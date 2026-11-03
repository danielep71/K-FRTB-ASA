"""Independent reference implementation of the prototype's FRTB-SA scope.

Written separately from the VBA engine (matrix formulation with numpy) so
that both can be reconciled on the same workbook. It reads parameters and
inputs from a workbook produced by the VBA setup (openpyxl, .xlsx) or takes
them as plain Python structures in unit tests.

Scope: SBM delta/vega/curvature for GIRR, CSR non-securitization, equity,
commodity and FX under three correlation scenarios; DRC non-securitization;
residual risk add-on. References are to the Basel Framework (MAR21-MAR23).
"""

from __future__ import annotations

import math
import re
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field

import numpy as np

RISK_CLASSES = ("GIRR", "CSR_NS", "EQ", "COM", "FX")
COMPONENTS = ("DELTA", "VEGA", "CURVATURE")
SCENARIOS = ("LOW", "MEDIUM", "HIGH")
EPS = 1e-6


# ---------------------------------------------------------------------------
# Parameters
# ---------------------------------------------------------------------------
@dataclass
class Params:
    reporting_ccy: str = "EUR"
    girr_specified_on: bool = False
    fx_specified_on: bool = False
    fx_curv_divisor_on: bool = False
    fx_curv_divisor_all_on: bool = False
    high_mult: float = 1.25
    low_mult: float = 0.75
    vega_rw_sigma: float = 0.55
    vega_alpha: float = 0.01
    vega_lh: list = field(default_factory=list)  # (class, bucket, lh)
    vertices: dict = field(default_factory=dict)
    girr_rw: dict = field(default_factory=dict)
    girr: dict = field(default_factory=dict)
    girr_specified: set = field(default_factory=set)
    csr_buckets: dict = field(default_factory=dict)  # b -> dict
    csr: dict = field(default_factory=dict)
    csr_sector: dict = field(default_factory=dict)  # (g1, g2) -> gamma
    eq_buckets: dict = field(default_factory=dict)
    eq: dict = field(default_factory=dict)
    eq_gamma: dict = field(default_factory=dict)
    com_buckets: dict = field(default_factory=dict)
    com: dict = field(default_factory=dict)
    com_gamma: dict = field(default_factory=dict)
    fx: dict = field(default_factory=dict)
    fx_pairs: set = field(default_factory=set)
    drc_rw: dict = field(default_factory=dict)
    drc_seniority: dict = field(default_factory=dict)  # name -> (rank, lgd)
    drc: dict = field(default_factory=dict)
    drc_buckets: list = field(default_factory=list)
    rrao_rw: dict = field(default_factory=dict)


# Numeric text as CORE_Util accepts it: optional sign, ASCII digits with an
# optional "." and an optional exponent, spaces around it at most. Commas,
# grouping, "_", "inf", "nan" and trailing text are not numbers, nor is a
# value beyond the Double range. Booleans and dates are not numbers.
_NUMERIC_TEXT = re.compile(r"[+-]?(\d*)(?:\.(\d*))?(?:[eE][+-]?\d+)?", re.ASCII)


def _number(v):
    """Value of a number or numeric text, or None when it is not a number."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    s = v.strip(" ")
    m = _NUMERIC_TEXT.fullmatch(s)
    if not m or not (m.group(1) or m.group(2)):
        return None
    x = float(s)
    return x if math.isfinite(x) else None


def _label(v) -> str:
    """Bucket label; a non-integral number matches no bucket instead of being rounded."""
    if v is None:
        return ""
    s = str(v).strip().upper()
    x = _number(v)
    if x is None:
        return s
    return str(int(x)) if x == int(x) else f"NON-INTEGER {s}"


def _text(v) -> str:
    return "" if v is None else str(v).strip().upper()


def _num(v, default=0.0) -> float:
    x = _number(v)
    return default if x is None else x


def _is_num(v) -> bool:
    return _number(v) is not None


def _blank_or_num(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "") or _is_num(v)


def _bool(v) -> bool:
    if isinstance(v, bool):
        return v
    return _text(v) in ("TRUE", "YES", "Y", "1", "-1")


def _flag(v):
    """An optional TRUE/FALSE input: False when blank, None for anything else.

    Only a Boolean or the text TRUE or FALSE is accepted; aliases such as YES,
    N or 1 are rejected (#144).
    """
    if v is None or (isinstance(v, str) and not v.strip()):
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, str):
        return {"TRUE": True, "FALSE": False}.get(_text(v))
    return None


FIRST_COL = 2  # tables start in column C (0-based index 2)
FIRST_DATA_ROW = 6  # input and output tables: header in row 5, data from row 6


def _tables(ws) -> dict:
    """Return {TAG: (header, rows)} for every block whose heading starts with "[TAG]"."""
    rows = [list(r) + [None] * 40 for r in ws.iter_rows(values_only=True)]
    out = {}
    for i, r in enumerate(rows):
        c0 = r[FIRST_COL]
        if isinstance(c0, str) and c0.startswith("[") and "]" in c0:
            header = []
            for v in rows[i + 1][FIRST_COL:]:
                if v is None or str(v).strip() == "":
                    break
                header.append(v)
            data = []
            j = i + 2
            while j < len(rows) and rows[j][FIRST_COL] not in (None, ""):
                data.append(rows[j][FIRST_COL : FIRST_COL + len(header)])
                j += 1
            out[c0[1 : c0.index("]")].upper()] = (header, data)
    return out


def _kv(tables, tag) -> dict:
    return {_text(r[0]): r[1] for r in tables[tag][1]}


def params_from_workbook(wb) -> Params:
    p = Params()
    s = _kv(_tables(wb["Control"]), "SETTINGS")
    p.reporting_ccy = _text(s.get("REPORTING_CCY", "EUR"))
    p.girr_specified_on = _bool(s.get("GIRR_SPECIFIED_CCY_DISCRETION"))
    p.fx_specified_on = _bool(s.get("FX_SPECIFIED_PAIR_DISCRETION"))
    p.fx_curv_divisor_on = _bool(s.get("FX_CURVATURE_DIVISOR"))
    p.fx_curv_divisor_all_on = _bool(s.get("FX_CURVATURE_DIVISOR_ALL"))

    t = _tables(wb["P_General"])
    kv = _kv(t, "CORR_SCENARIOS")
    p.high_mult, p.low_mult = float(kv["HIGH_MULTIPLIER"]), float(kv["LOW_MULTIPLIER"])
    kv = _kv(t, "VEGA")
    p.vega_rw_sigma, p.vega_alpha = float(kv["RW_SIGMA"]), float(kv["ALPHA"])
    p.vega_lh = [(_text(r[0]), _label(r[1]), float(r[2])) for r in t["VEGA_LH"][1]]
    p.vertices = {k: [float(x) for x in str(v).split(",") if x.strip()] for k, v in _kv(t, "VERTICES").items()}

    t = _tables(wb["P_GIRR"])
    p.girr_rw = {float(r[0]): float(r[1]) for r in t["GIRR_RW"][1]}
    p.girr = {k: float(v) for k, v in _kv(t, "GIRR_PARAMS").items()}
    p.girr_specified = {_text(r[0]) for r in t["GIRR_SPECIFIED_CCY"][1]}

    t = _tables(wb["P_CSR"])
    for r in t["CSR_BUCKETS"][1]:
        p.csr_buckets[_label(r[0])] = {"quality": _text(r[2]), "group": _label(r[3]), "type": _text(r[4]), "rw": float(r[5])}
    p.csr = {k: float(v) for k, v in _kv(t, "CSR_PARAMS").items()}
    header, rows = t["CSR_GAMMA_SECTOR"]
    cols = [_label(h) for h in header[1:]]
    for r in rows:
        for c, v in zip(cols, r[1:]):
            p.csr_sector[(_label(r[0]), c)] = float(v)

    t = _tables(wb["P_EQ"])
    for r in t["EQ_BUCKETS"][1]:
        p.eq_buckets[_label(r[0])] = {"type": _text(r[2]), "rw_spot": float(r[3]), "rw_repo": float(r[4]), "rho": float(r[5])}
    p.eq = {k: float(v) for k, v in _kv(t, "EQ_PARAMS").items()}
    header, rows = t["EQ_GAMMA"]
    for r in rows:
        for c, v in zip(header[1:], r[1:]):
            p.eq_gamma[(_label(r[0]), _label(c))] = float(v)

    t = _tables(wb["P_COM"])
    for r in t["COM_BUCKETS"][1]:
        p.com_buckets[_label(r[0])] = {"rw": float(r[2]), "rho": float(r[3])}
    p.com = {k: float(v) for k, v in _kv(t, "COM_PARAMS").items()}
    header, rows = t["COM_GAMMA"]
    for r in rows:
        for c, v in zip(header[1:], r[1:]):
            p.com_gamma[(_label(r[0]), _label(c))] = float(v)

    t = _tables(wb["P_FX"])
    p.fx = {k: float(v) for k, v in _kv(t, "FX_PARAMS").items()}
    p.fx_pairs = {_text(r[0]) for r in t["FX_SPECIFIED_PAIRS"][1]}

    t = _tables(wb["P_DRC"])
    p.drc_rw = {_text(r[0]): float(r[1]) for r in t["DRC_RW"][1]}
    p.drc_seniority = {_text(r[0]): (int(r[1]), float(r[2])) for r in t["DRC_SENIORITY"][1]}
    p.drc = {k: float(v) for k, v in _kv(t, "DRC_PARAMS").items()}
    p.drc_buckets = [_text(r[0]) for r in t["DRC_BUCKETS"][1]]

    t = _tables(wb["P_RRAO"])
    p.rrao_rw = {_text(r[0]): float(r[1]) for r in t["RRAO_RW"][1]}
    return p


def _cell_value(cell):
    """A cell's value, with a number formatted as TRUE/FALSE read as a Boolean.

    LibreOffice saves a Boolean cell as a number whose format shows TRUE or
    FALSE; the VBA reads such a cell as the flag it shows (CellBool, #144). A
    plain number keeps its value, so 1 or 0 is not a flag.
    """
    v = cell.value
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return v
    sections = [s.strip() for s in (cell.number_format or "").split(";")]
    if len(sections) < 2:
        return v
    shown = sections[0] if v > 0 else sections[1] if v < 0 else sections[2] if len(sections) > 2 else sections[0]
    shown = shown.strip('"').strip().upper()
    return {"TRUE": True, "FALSE": False}.get(shown, v)


def _input_rows(ws, ncols):
    out = []
    for cells in ws.iter_rows(min_row=FIRST_DATA_ROW, min_col=FIRST_COL + 1):
        r = [_cell_value(c) for c in cells] + [None] * ncols
        if r[0] in (None, "") and r[1] in (None, ""):
            break
        out.append(r[:ncols])
    return out


def inputs_from_workbook(wb) -> dict:
    return {
        "delta": _input_rows(wb["In_Delta"], 7),
        "vega": _input_rows(wb["In_Vega"], 8),
        "curvature": _input_rows(wb["In_Curvature"], 8),
        "drc": _input_rows(wb["In_DRC"], 9),
        "rrao": _input_rows(wb["In_RRAO"], 3),
    }


# ---------------------------------------------------------------------------
# Risk factors
# ---------------------------------------------------------------------------
@dataclass
class RF:
    rc: str
    bucket: str
    qual: str
    sub: str
    t1: float
    t2: float
    a1: float
    a2: float = 0.0
    rw: float = 0.0


def _in(x, values):
    return any(abs(x - v) < EPS for v in values)


def _bucket_valid(p: Params, rc, b) -> bool:
    if rc in ("GIRR", "FX"):
        return len(b) == 3 and not (rc == "FX" and b == p.reporting_ccy)
    table = {"CSR_NS": p.csr_buckets, "EQ": p.eq_buckets, "COM": p.com_buckets}[rc]
    return b in table


def _net(rfs):
    out = OrderedDict()
    for f in rfs:
        key = (f.rc, f.bucket, f.qual, f.sub, round(f.t1, 6), round(f.t2, 6))
        if key in out:
            out[key].a1 += f.a1
            out[key].a2 += f.a2
        else:
            out[key] = f
    return list(out.values())


def delta_risk_factors(p: Params, rows):
    rfs = []
    for r in rows:
        rc, b, q, sub = _text(r[1]), _label(r[2]), _text(r[3]), _text(r[4])
        t = _num(r[5]) if _is_num(r[5]) else 0.0
        if rc not in RISK_CLASSES or not _bucket_valid(p, rc, b) or not _is_num(r[6]):
            continue
        if not _blank_or_num(r[5]):
            continue
        if rc == "GIRR":
            sub = sub or "YIELD"
            if sub == "YIELD" and not _in(t, p.vertices["GIRR_TENORS"]):
                continue
            if sub in ("INFLATION", "XCCY"):
                t = 0.0
            if sub not in ("YIELD", "INFLATION", "XCCY"):
                continue
        elif rc == "CSR_NS":
            sub = sub or "BOND"
            if sub not in ("BOND", "CDS") or not _in(t, p.vertices["CSR_TENORS"]):
                continue
        elif rc == "EQ":
            sub = sub or "SPOT"
            if sub not in ("SPOT", "REPO"):
                continue
            t = 0.0
        elif rc == "COM":
            if not _in(t, p.vertices["COM_TENORS"]):
                continue
        elif rc == "FX":
            q, sub, t = b, "", 0.0
        if not q:
            continue
        rfs.append(RF(rc, b, q, sub, t, 0.0, _num(r[6])))
    rfs = _net(rfs)
    for f in rfs:
        f.rw = delta_rw(p, f)
    return rfs


def delta_rw(p: Params, f: RF) -> float:
    if f.rc == "GIRR":
        if f.sub == "INFLATION":
            rw = p.girr["RW_INFLATION"]
        elif f.sub == "XCCY":
            rw = p.girr["RW_XCCY"]
        else:
            rw = next(v for k, v in p.girr_rw.items() if abs(k - f.t1) < EPS)
        # MAR21.44: "the above risk weights" covers MAR21.42 and MAR21.43.
        if p.girr_specified_on and (f.bucket in p.girr_specified or f.bucket == p.reporting_ccy):
            rw /= p.girr["SPECIFIED_DIVISOR"]
        return rw
    if f.rc == "CSR_NS":
        return p.csr_buckets[f.bucket]["rw"]
    if f.rc == "EQ":
        b = p.eq_buckets[f.bucket]
        return b["rw_repo"] if f.sub == "REPO" else b["rw_spot"]
    if f.rc == "COM":
        return p.com_buckets[f.bucket]["rw"]
    rw = p.fx["RW"]
    if p.fx_specified_on and fx_pair_specified(p, f.bucket, p.reporting_ccy):
        rw /= p.fx["SPECIFIED_DIVISOR"]
    return rw


def fx_pair_specified(p: Params, a: str, b: str) -> bool:
    """Listed pair, or first-order cross of two currencies listed against USD (MAR21.88)."""
    if {f"{a}/{b}", f"{b}/{a}"} & p.fx_pairs:
        return True
    usd_paired = {"USD"} | {x.split("/")[1] for x in p.fx_pairs if x.startswith("USD/")} | {
        x.split("/")[0] for x in p.fx_pairs if x.endswith("/USD")
    }
    return a in usd_paired and b in usd_paired


def vega_rw(p: Params, rc, bucket) -> float:
    lh = next((v for c, b, v in p.vega_lh if c == rc and b == bucket), None)
    if lh is None:
        lh = next(v for c, b, v in p.vega_lh if c == rc and b == "*")
    return min(p.vega_rw_sigma * math.sqrt(lh / 10.0), 1.0)


def vega_risk_factors(p: Params, rows):
    rfs = []
    mats = p.vertices["VEGA_MATURITIES"]
    for r in rows:
        rc, b, q, sub = _text(r[1]), _label(r[2]), _text(r[3]), _text(r[4])
        if not (_blank_or_num(r[5]) and _blank_or_num(r[6])):
            continue
        topt, tund = _num(r[5]), _num(r[6])
        if rc not in RISK_CLASSES or not _bucket_valid(p, rc, b) or not _is_num(r[7]):
            continue
        if not _in(topt, mats):
            continue
        if rc == "GIRR":
            sub = sub or "YIELD"
            if sub == "YIELD":
                # one vega surface per currency: option x underlying maturity
                if not _in(tund, mats):
                    continue
                q = ""
            elif sub in ("INFLATION", "XCCY"):
                # option maturity only (MAR21.8 FAQ4): an underlying maturity is an input error
                if not (r[6] is None or (isinstance(r[6], str) and r[6].strip() == "")):
                    continue
                tund = 0.0
            else:
                continue
        else:
            tund = 0.0
            if rc in ("CSR_NS", "EQ", "COM"):
                sub = ""  # vega risk factors: name (or commodity) and option maturity only
            elif rc == "FX":
                q, sub = b, ""
            if not q:
                continue
        rfs.append(RF(rc, b, q, sub, topt, tund, _num(r[7])))
    rfs = _net(rfs)
    for f in rfs:
        f.rw = vega_rw(p, f.rc, f.bucket)
    return rfs


def curvature_risk_factors(p: Params, rows):
    rfs = []
    for r in rows:
        r = list(r) + [None] * (8 - len(r))
        rc, b, q, sub = _text(r[1]), _label(r[2]), _text(r[3]), _text(r[4])
        if rc not in RISK_CLASSES or not _bucket_valid(p, rc, b) or not (_is_num(r[5]) and _is_num(r[6])):
            continue
        ref_ccy = _flag(r[7])
        if ref_ccy is None:
            continue
        up, dn = _num(r[5]), _num(r[6])
        # MAR21.98: options not referencing the reporting currency may be divided
        # by 1.5; every FX instrument only with supervisory approval (_ALL)
        divided = p.fx_curv_divisor_all_on or (p.fx_curv_divisor_on and not ref_ccy)
        div = p.fx["CURVATURE_DIVISOR"] if divided else 1.0
        if rc == "GIRR":
            q, sub = "", ""
        elif rc in ("CSR_NS", "EQ", "COM"):
            sub = ""  # one curvature risk factor per issuer or commodity
        elif rc == "FX":
            q, sub = b, ""
            up, dn = up / div, dn / div
        if rc != "GIRR" and not q:
            continue
        rfs.append(RF(rc, b, q, sub, 0.0, 0.0, up, dn))
    return _net(rfs)


# ---------------------------------------------------------------------------
# Correlations
# ---------------------------------------------------------------------------
def scenario(rho, scen, p: Params):
    rho = np.asarray(rho, dtype=float)
    if scen == "HIGH":
        return np.minimum(p.high_mult * rho, 1.0)
    if scen == "LOW":
        return np.maximum(2.0 * rho - 1.0, p.low_mult * rho)
    return rho


def _issuer_rho(p: Params, rc, bucket, a: RF, b: RF) -> float:
    """Underlying correlation without the tenor component."""
    if rc == "CSR_NS":
        name = 1.0 if a.qual == b.qual else (p.csr["RHO_NAME_INDEX"] if p.csr_buckets[bucket]["type"] == "INDEX" else p.csr["RHO_NAME"])
        return name * (1.0 if a.sub == b.sub else p.csr["RHO_BASIS"])
    if rc == "EQ":
        return 1.0 if a.qual == b.qual else p.eq_buckets[bucket]["rho"]
    if rc == "COM":
        cty = 1.0 if a.qual == b.qual else p.com_buckets[bucket]["rho"]
        return cty * (1.0 if a.sub == b.sub else p.com["RHO_BASIS"])
    return 1.0


def delta_rho(p: Params, rc, bucket, a: RF, b: RF) -> float:
    if rc == "GIRR":
        g = p.girr
        if "XCCY" in (a.sub, b.sub):
            return g["RHO_XCCY"]
        if a.sub == b.sub == "INFLATION":
            return g["RHO_DIFF_CURVE"]
        if "INFLATION" in (a.sub, b.sub):
            return g["RHO_INFLATION"]
        if abs(a.t1 - b.t1) < EPS:
            rho = 1.0
        else:
            rho = max(math.exp(-g["THETA"] * abs(a.t1 - b.t1) / min(a.t1, b.t1)), g["RHO_FLOOR"])
        return rho * (1.0 if a.qual == b.qual else g["RHO_DIFF_CURVE"])
    if rc == "CSR_NS":
        return _issuer_rho(p, rc, bucket, a, b) * (1.0 if abs(a.t1 - b.t1) < EPS else p.csr["RHO_TENOR"])
    if rc == "EQ":
        return _issuer_rho(p, rc, bucket, a, b) * (1.0 if a.sub == b.sub else p.eq["RHO_SPOT_REPO"])
    if rc == "COM":
        return _issuer_rho(p, rc, bucket, a, b) * (1.0 if abs(a.t1 - b.t1) < EPS else p.com["RHO_TENOR"])
    return 1.0


def _mat_rho(p: Params, x, y):
    return 1.0 if abs(x - y) < EPS else math.exp(-p.vega_alpha * abs(x - y) / min(x, y))


def _name_rho(p: Params, rc, bucket, a: RF, b: RF) -> float:
    """Delta correlation on the dimension vega/curvature share with delta: name or commodity."""
    if a.qual == b.qual:
        return 1.0
    if rc == "CSR_NS":
        return p.csr["RHO_NAME_INDEX"] if p.csr_buckets[bucket]["type"] == "INDEX" else p.csr["RHO_NAME"]
    if rc == "EQ":
        return p.eq_buckets[bucket]["rho"]
    if rc == "COM":
        return p.com_buckets[bucket]["rho"]
    return 1.0


def _girr_vega_rho(p: Params, a: RF, b: RF) -> float:
    """Delta correlation of the GIRR delta risk factors two vega risk factors
    correspond to (MAR21.94, MAR21.46-MAR21.49): the underlying maturity for
    two yield surfaces, otherwise the sub-type and curve correlations."""
    g = p.girr
    if a.sub == b.sub == "YIELD":
        return _mat_rho(p, a.t2, b.t2)
    if a.sub == b.sub and a.qual == b.qual:
        return 1.0
    if "XCCY" in (a.sub, b.sub):
        return g["RHO_XCCY"]
    if a.sub == b.sub == "INFLATION":
        return g["RHO_DIFF_CURVE"]
    return g["RHO_INFLATION"]


def vega_rho(p: Params, rc, bucket, a: RF, b: RF) -> float:
    und = _girr_vega_rho(p, a, b) if rc == "GIRR" else _name_rho(p, rc, bucket, a, b)
    return min(_mat_rho(p, a.t1, b.t1) * und, 1.0)


def curvature_rho(p: Params, rc, bucket, a: RF, b: RF) -> float:
    return _name_rho(p, rc, bucket, a, b) ** 2


def gamma(p: Params, rc, b, c) -> float:
    if rc == "GIRR":
        return p.girr["GAMMA"]
    if rc == "FX":
        return p.fx["GAMMA"]
    if rc == "CSR_NS":
        bb, cc = p.csr_buckets[b], p.csr_buckets[c]
        both_named = bb["type"] == "NAME" and cc["type"] == "NAME"  # buckets 1-15
        rating = p.csr["GAMMA_RATING"] if both_named and bb["quality"] != cc["quality"] else 1.0
        return rating * p.csr_sector[(bb["group"], cc["group"])]
    if rc == "EQ":
        return p.eq_gamma[(b, c)]
    return p.com_gamma[(b, c)]


def is_other(p: Params, rc, b) -> bool:
    if rc == "CSR_NS":
        return p.csr_buckets[b]["type"] == "OTHER"
    if rc == "EQ":
        return p.eq_buckets[b]["type"] == "OTHER"
    return False


def _bucket_sort_key(b):
    return (0, int(b), "") if b.isdigit() else (1, 0, b)


# ---------------------------------------------------------------------------
# SBM aggregation
# ---------------------------------------------------------------------------
def _psi_matrix(x):
    neg = x < 0
    return np.where(np.outer(neg, neg), 0.0, 1.0)


def sbm_component(p: Params, comp, rfs):
    """Return (capital[rc][scen], alt[rc][scen], buckets[(rc, scen, bucket)] = (Kb, Sb, n, note))."""
    rho_fn = {"DELTA": delta_rho, "VEGA": vega_rho, "CURVATURE": curvature_rho}[comp]
    capital = {rc: {s: 0.0 for s in SCENARIOS} for rc in RISK_CLASSES}
    alt_used = {rc: {s: False for s in SCENARIOS} for rc in RISK_CLASSES}
    bucket_out = {}
    for rc in RISK_CLASSES:
        groups = defaultdict(list)
        for f in rfs:
            if f.rc == rc:
                groups[f.bucket].append(f)
        labels = sorted(groups, key=_bucket_sort_key)
        if not labels:
            continue
        base_rho = {}
        for b in labels:
            fs = groups[b]
            n = len(fs)
            m = np.eye(n)
            for i in range(n):
                for j in range(n):
                    if i != j:
                        m[i, j] = rho_fn(p, rc, b, fs[i], fs[j])
            base_rho[b] = m
        other = np.array([is_other(p, rc, b) for b in labels])
        # "Other" buckets are added outside the square root, so no gamma is needed for them
        base_gamma = np.array([
            [1.0 if (b == c or other[i] or other[j]) else gamma(p, rc, b, c) for j, c in enumerate(labels)]
            for i, b in enumerate(labels)
        ])
        if comp == "CURVATURE":
            base_gamma = base_gamma ** 2

        for s in SCENARIOS:
            K = np.zeros(len(labels))
            S = np.zeros(len(labels))
            notes = []
            for k, b in enumerate(labels):
                fs = groups[b]
                rho = scenario(base_rho[b], s, p)
                np.fill_diagonal(rho, 1.0)
                note = ""
                if comp == "CURVATURE":
                    up = np.array([f.a1 for f in fs])
                    dn = np.array([f.a2 for f in fs])
                    if other[k]:
                        kup, kdn = np.maximum(up, 0).sum(), np.maximum(dn, 0).sum()
                    else:
                        def kcurv(x):
                            off = rho * np.outer(x, x) * _psi_matrix(x)
                            np.fill_diagonal(off, 0.0)
                            return math.sqrt(max((np.maximum(x, 0) ** 2).sum() + off.sum(), 0.0))
                        kup, kdn = kcurv(up), kcurv(dn)
                    if math.isclose(kup, kdn, rel_tol=1e-7, abs_tol=1e-7):
                        pick_up = up.sum() > dn.sum()
                    else:
                        pick_up = kup > kdn
                    K[k], S[k] = (kup, up.sum()) if pick_up else (kdn, dn.sum())
                    note = "UP" if pick_up else "DOWN"
                else:
                    ws = np.array([f.rw * f.a1 for f in fs])
                    S[k] = ws.sum()
                    K[k] = np.abs(ws).sum() if other[k] else math.sqrt(max(ws @ rho @ ws, 0.0))
                if other[k]:
                    note = (note + ";OTHER") if note else "OTHER"
                notes.append(note)

            g = scenario(base_gamma, s, p)
            np.fill_diagonal(g, 0.0)
            reg = ~other

            def across(Sv):
                Kr, Sr, gr = K[reg], Sv[reg], g[np.ix_(reg, reg)]
                cross = gr * np.outer(Sr, Sr)
                if comp == "CURVATURE":
                    cross = cross * _psi_matrix(Sr)
                return (Kr ** 2).sum() + cross.sum()

            total_sq = across(S)
            alt = False
            S_report = S.copy()
            if total_sq < 0 and comp != "CURVATURE":
                alt = True
                S_alt = np.clip(S, -K, K)
                total_sq = across(S_alt)
                S_report = np.where(reg, S_alt, S)
            capital[rc][s] = math.sqrt(max(total_sq, 0.0)) + K[other].sum()
            alt_used[rc][s] = alt
            for k, b in enumerate(labels):
                note = notes[k]
                if alt and reg[k]:
                    note = (note + ";ALT_SB") if note else "ALT_SB"
                bucket_out[(rc, s, b)] = (K[k], S_report[k], len(groups[b]), note)
    return capital, alt_used, bucket_out


# ---------------------------------------------------------------------------
# DRC and RRAO
# ---------------------------------------------------------------------------
def drc(p: Params, rows):
    """Return (total, obligors{name: (bucket, quality, rw, long, short)}, buckets{b: (...)})."""
    positions = []
    for r in rows:
        obl, bucket, quality, sen = _text(r[1]), _text(r[2]), _text(r[3]), _text(r[4])
        if not obl or bucket not in p.drc_buckets or sen not in p.drc_seniority or quality not in p.drc_rw:
            continue
        if not (_is_num(r[5]) and _is_num(r[6])) or not (_blank_or_num(r[7]) and _blank_or_num(r[8])):
            continue
        notional, mv = _num(r[5]), _num(r[6])
        maturity = _num(r[7], p.drc["MAX_MATURITY"])
        # MAR22.16 and its FAQ: cash equity is three months or at least one year.
        if sen == "EQUITY" and not (abs(maturity - p.drc["MIN_MATURITY"]) < 1e-6 or maturity >= p.drc["MAX_MATURITY"]):
            continue
        lgd = _num(r[8]) if _is_num(r[8]) else p.drc_seniority[sen][1]
        jtd = lgd * notional + (mv - notional)
        jtd = max(jtd, 0.0) if notional >= 0 else min(jtd, 0.0)
        w = min(max(maturity, p.drc["MIN_MATURITY"]), p.drc["MAX_MATURITY"]) / p.drc["MAX_MATURITY"]
        positions.append((obl, bucket, quality, p.drc_seniority[sen][0], jtd * w))

    obligors = OrderedDict()
    for obl, bucket, quality, rank, jtd in positions:
        o = obligors.setdefault(obl, {"buckets": set(), "quality": quality, "pos": []})
        o["buckets"].add(bucket)
        # MAR22.24: the worst quality (highest risk weight); the first one on a tie
        if p.drc_rw[quality] > p.drc_rw[o["quality"]]:
            o["quality"] = quality
        o["pos"].append((rank, jtd))
    # MAR22.22: an obligor belongs to one bucket; a split obligor is an input error
    obligors = OrderedDict((k, v) for k, v in obligors.items() if len(v["buckets"]) == 1)
    for o in obligors.values():
        o["bucket"] = next(iter(o["buckets"]))

    max_rank = max(v[0] for v in p.drc_seniority.values())
    result = OrderedDict()
    for name, o in obligors.items():
        longs = [0.0] * (max_rank + 1)
        shorts = [0.0] * (max_rank + 1)
        for rank, jtd in o["pos"]:
            if jtd >= 0:
                longs[rank] += jtd
            else:
                shorts[rank] -= jtd
        # Shorts can only hedge longs of equal or higher seniority rank.
        remaining_long = 0.0
        pool = 0.0
        for rank in range(1, max_rank + 1):
            pool += shorts[rank]
            used = min(longs[rank], pool)
            remaining_long += longs[rank] - used
            pool -= used
        rw = p.drc_rw[o["quality"]]
        result[name] = (o["bucket"], o["quality"], rw, remaining_long, pool)

    buckets = OrderedDict()
    total = 0.0
    for b in p.drc_buckets:
        members = [v for v in result.values() if v[0] == b]
        lg = sum(v[3] for v in members)
        sh = sum(v[4] for v in members)
        hbr = lg / (lg + sh) if lg + sh > 0 else 0.0
        wl = sum(v[2] * v[3] for v in members)
        wsh = sum(v[2] * v[4] for v in members)
        charge = max(wl - hbr * wsh, 0.0)
        buckets[b] = (lg, sh, hbr, wl, wsh, charge)
        total += charge
    return total, result, buckets


def rrao(p: Params, rows) -> float:
    total = 0.0
    for r in rows:
        cat = _text(r[1])
        if cat in p.rrao_rw and _is_num(r[2]) and _num(r[2]) >= 0:
            total += p.rrao_rw[cat] * _num(r[2])
    return total


# ---------------------------------------------------------------------------
# Full run
# ---------------------------------------------------------------------------
def run(p: Params, inputs: dict) -> dict:
    loaders = {
        "DELTA": delta_risk_factors(p, inputs["delta"]),
        "VEGA": vega_risk_factors(p, inputs["vega"]),
        "CURVATURE": curvature_risk_factors(p, inputs["curvature"]),
    }
    capital, alt, buckets = {}, {}, {}
    for comp, rfs in loaders.items():
        c, a, b = sbm_component(p, comp, rfs)
        for rc in RISK_CLASSES:
            for s in SCENARIOS:
                capital[(rc, comp, s)] = c[rc][s]
                alt[(rc, comp, s)] = a[rc][s]
        for (rc, s, bk), v in b.items():
            buckets[(rc, comp, s, bk)] = v
    sbm_by_scen = {s: sum(capital[(rc, comp, s)] for rc in RISK_CLASSES for comp in COMPONENTS) for s in SCENARIOS}
    best = "LOW"
    for s in SCENARIOS:
        if sbm_by_scen[s] > sbm_by_scen[best]:
            best = s
    drc_total, drc_obl, drc_buckets = drc(p, inputs["drc"])
    rrao_total = rrao(p, inputs["rrao"])
    summary = {
        "SBM_LOW": sbm_by_scen["LOW"],
        "SBM_MEDIUM": sbm_by_scen["MEDIUM"],
        "SBM_HIGH": sbm_by_scen["HIGH"],
        "SBM_SCENARIO": best,
        "SBM": sbm_by_scen[best],
        "DRC_NONSEC": drc_total,
        "RRAO": rrao_total,
        "TOTAL": sbm_by_scen[best] + drc_total + rrao_total,
    }
    return {
        "capital": capital,
        "alt": alt,
        "buckets": buckets,
        "summary": summary,
        "drc_obligors": drc_obl,
        "drc_buckets": drc_buckets,
    }
