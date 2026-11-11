"""Reconciliation of the VBA engine (run in LibreOffice) with the reference.

The VBA modules are executed headless by tools/lo_runner.py; every output
cell is compared with tests/frtb_reference.py computed on the same
parameters and inputs. LibreOffice's VBA support is not Excel: these tests
show that the engine logic is right, not that Excel compiles or runs it.
"""

from __future__ import annotations

import math
from typing import Any

import openpyxl
import pytest

import acceptance_cases as cases
import frtb_reference as ref
import lo_runner
import regulatory_defaults as reg

REL = 1e-9
ABS = 1e-6


def close(a, b):
    return math.isclose(float(a), float(b), rel_tol=REL, abs_tol=ABS)


FIRST_COL = 3  # tables start in column C
FIRST_DATA_ROW = 6  # header in row 5


def sheet_rows(wb, name, min_row=FIRST_DATA_ROW):
    rows = wb[name].iter_rows(min_row=min_row, min_col=FIRST_COL, values_only=True)
    return [r for r in rows if r and r[0] not in (None, "")]


def reconcile(wb):  # noqa: C901 -- one pass over every output table of the engine
    """Return a list of mismatch descriptions (empty when everything agrees)."""
    p = ref.params_from_workbook(wb)
    res = ref.run(p, ref.inputs_from_workbook(wb))
    problems = []

    rows = sheet_rows(wb, "Out_RiskClass")
    if len(rows) != len(ref.RISK_CLASSES) * len(ref.COMPONENTS) * len(ref.SCENARIOS):
        problems.append(f"Out_RiskClass has {len(rows)} rows")
    for rc, comp, scen, cap, _n, alt in (r[:6] for r in rows):
        if not close(cap, res["capital"][(rc, comp, scen)]):
            problems.append(f"capital {rc}/{comp}/{scen}: vba={cap} ref={res['capital'][(rc, comp, scen)]}")
        if (alt == "Y") != res["alt"][(rc, comp, scen)]:
            problems.append(f"alt flag {rc}/{comp}/{scen}")

    vba_buckets = {(r[0], r[1], r[2], str(r[3])): r[4:8] for r in sheet_rows(wb, "Out_Buckets")}
    if set(vba_buckets) != set(res["buckets"]):
        problems.append(f"bucket keys differ: {set(vba_buckets) ^ set(res['buckets'])}")
    for key, (kb, sb, n, note) in res["buckets"].items():
        v = vba_buckets.get(key)
        if v is None:
            continue
        if not (close(v[0], kb) and close(v[1], sb) and v[2] == n and (v[3] or "") == note):
            problems.append(f"bucket {key}: vba={v} ref={(kb, sb, n, note)}")

    summary = {r[0]: r[1] for r in sheet_rows(wb, "Out_Summary")}
    for key, value in res["summary"].items():
        if isinstance(value, str):
            if summary.get(key) != value:
                problems.append(f"summary {key}: vba={summary.get(key)} ref={value}")
        elif not close(summary.get(key), value):
            problems.append(f"summary {key}: vba={summary.get(key)} ref={value}")

    drc_rows = sheet_rows(wb, "Out_DRC")
    vba_obl = {r[0]: r[1:6] for r in drc_rows}
    if set(vba_obl) != set(res["drc_obligors"]):
        problems.append("DRC obligor sets differ")
    for name, (bucket, quality, rw, lg, sh) in res["drc_obligors"].items():
        v = vba_obl.get(name)
        if v and not (v[0] == bucket and v[1] == quality and close(v[2], rw) and close(v[3], lg) and close(v[4], sh)):
            problems.append(f"DRC obligor {name}: vba={v} ref={(bucket, quality, rw, lg, sh)}")
    vba_drc_buckets = {
        r[7]: r[8:14]
        for r in wb["Out_DRC"].iter_rows(min_row=FIRST_DATA_ROW, min_col=FIRST_COL, values_only=True)
        if len(r) > 7 and r[7]
    }
    for b, vals in res["drc_buckets"].items():
        v = vba_drc_buckets.get(b)
        if v is None or not all(close(x, y) for x, y in zip(v, vals)):
            problems.append(f"DRC bucket {b}: vba={v} ref={vals}")
    return problems


def log_rows(wb, level=None):
    rows = sheet_rows(wb, "Log")
    return [r for r in rows if level is None or r[0] == level]


# --------------------------------------------------------------------------
# Default sample portfolio
# --------------------------------------------------------------------------
def test_default_run_reconciles_with_reference(lo_default_run):
    assert reconcile(lo_default_run) == []


def tie_notes(wb):
    return [v for row in wb["Results"].iter_rows(min_row=FIRST_DATA_ROW, values_only=True)
            for v in row if isinstance(v, str) and v.startswith("Scenario tie")]


def test_default_run_has_no_tie_note(lo_default_run):
    assert tie_notes(lo_default_run) == []


def test_default_run_has_clean_log(lo_default_run):
    assert log_rows(lo_default_run, "ERROR") == []
    assert log_rows(lo_default_run, "WARN") == []


def test_default_run_exercises_alternative_sb(lo_default_run):
    alt = [r for r in sheet_rows(lo_default_run, "Out_RiskClass") if r[5] == "Y"]
    assert alt, "the sample portfolio should trigger the alternative Sb specification"


def test_in_workbook_regression_test_passes(lo_default_run):
    ws = lo_default_run["Test_Expected"]
    assert str(ws["C3"].value).startswith("Result: PASS")
    statuses = [r[4] for r in ws.iter_rows(min_row=7, min_col=FIRST_COL, values_only=True) if r[0]]
    assert len(statuses) == 52 and set(statuses) == {"PASS"}


def test_workbook_layout(lo_default_run):
    expected = [
        "README", "Control", "Results", "In_Delta", "In_Vega", "In_Curvature", "In_DRC", "In_RRAO",
        "Out_Summary", "Out_RiskClass", "Out_Buckets", "Out_DRC", "Log", "P_General", "P_GIRR", "P_CSR",
        "P_EQ", "P_COM", "P_FX", "P_DRC", "P_RRAO", "Test_Expected",
    ]
    assert lo_default_run.sheetnames == expected


# --------------------------------------------------------------------------
# Parameters written by M_Setup vs the independently typed defaults
# --------------------------------------------------------------------------
def test_parameter_sheets_match_regulatory_defaults(lo_default_run):
    got = ref.params_from_workbook(lo_default_run)
    exp = reg.build_params()
    assert got.girr_rw == pytest.approx(exp.girr_rw)
    assert got.girr == pytest.approx(exp.girr)
    assert got.girr_specified == exp.girr_specified
    assert got.csr_buckets == exp.csr_buckets
    assert got.csr == pytest.approx(exp.csr)
    assert got.csr_sector == pytest.approx(exp.csr_sector)
    assert set(got.eq_buckets) == set(exp.eq_buckets)
    for b in exp.eq_buckets:
        assert got.eq_buckets[b]["type"] == exp.eq_buckets[b]["type"]
        assert got.eq_buckets[b]["rw_spot"] == pytest.approx(exp.eq_buckets[b]["rw_spot"])
        assert got.eq_buckets[b]["rw_repo"] == pytest.approx(exp.eq_buckets[b]["rw_repo"])
        assert got.eq_buckets[b]["rho"] == pytest.approx(exp.eq_buckets[b]["rho"])
    assert got.eq_gamma == pytest.approx(exp.eq_gamma)
    assert set(got.com_buckets) == set(exp.com_buckets)
    for b, v in exp.com_buckets.items():
        assert got.com_buckets[b]["rw"] == pytest.approx(v["rw"])
        assert got.com_buckets[b]["rho"] == pytest.approx(v["rho"])
    assert got.com_gamma == pytest.approx(exp.com_gamma)
    assert got.fx == pytest.approx(exp.fx)
    assert got.fx_pairs == exp.fx_pairs
    assert got.drc_rw == pytest.approx(exp.drc_rw)
    assert got.drc_seniority == exp.drc_seniority
    assert got.drc == pytest.approx(exp.drc)
    assert got.drc_buckets == exp.drc_buckets
    assert got.rrao_rw == pytest.approx(exp.rrao_rw)
    assert got.vertices == exp.vertices
    assert sorted(got.vega_lh) == sorted(exp.vega_lh)
    assert (got.high_mult, got.low_mult, got.vega_rw_sigma, got.vega_alpha) == (1.25, 0.75, 0.55, 0.01)


# --------------------------------------------------------------------------
# Discretions switched on and invalid input rows
# --------------------------------------------------------------------------
@pytest.fixture(scope="module")
def lo_variant_run(lo_default_path, tmp_path_factory):
    wb = openpyxl.load_workbook(lo_default_path)
    cases.apply_variant(wb)
    folder = tmp_path_factory.mktemp("lo-variant")
    source = folder / "variant-in.xlsx"
    wb.save(source)
    out = folder / "variant-out.xlsx"
    errors = lo_runner.run(out, ["M_Main.RunFRTB_Silent"], input_book=source)
    assert errors == [], errors
    return openpyxl.load_workbook(out, data_only=True)


def test_variant_run_reconciles_with_reference(lo_variant_run):
    assert reconcile(lo_variant_run) == []


def test_variant_run_logs_every_invalid_row(lo_variant_run):
    errors = log_rows(lo_variant_run, "ERROR")
    sources = sorted(r[1] for r in errors)
    assert len(errors) == cases.EXPECTED_ERRORS, errors
    assert sum(s.startswith("In_Delta") for s in sources) == 7
    assert sum(s.startswith("In_DRC") for s in sources) == 7
    assert any("Cash equity maturity" in r[2] for r in errors)
    flag_errors = sorted(r[1] for r in errors if "(RefReportingCcy) must be TRUE, FALSE or blank" in r[2])
    assert len(flag_errors) == 4 and all(s.startswith("In_Curvature") for s in flag_errors), flag_errors
    # BANK_B, split across CORP and SOV, is skipped; BANK_A takes its worst quality (#142)
    assert any("BANK_B has positions in several buckets" in r[2] for r in errors)
    warnings = log_rows(lo_variant_run, "WARN")
    assert len(warnings) == cases.EXPECTED_WARNINGS and "the worst, BBB, is used" in warnings[0][2]
    drc = {r[0]: r for r in sheet_rows(lo_variant_run, "Out_DRC")}
    assert "BANK_B" not in drc and drc["BANK_A"][2] == "BBB"


def test_variant_run_applies_discretions(lo_variant_run, lo_default_run):
    def cap(wb, rc, comp):
        return {(r[0], r[1], r[2]): r[3] for r in sheet_rows(wb, "Out_RiskClass")}[(rc, comp, "MEDIUM")]

    # Specified currency and pair discretions divide the GIRR and FX delta charges by
    # sqrt(2): the sample GIRR is EUR and USD only, inflation and cross-currency basis
    # included (#136), and all sample FX currencies are specified against EUR,
    # directly or as crosses
    assert cap(lo_variant_run, "GIRR", "DELTA") == pytest.approx(cap(lo_default_run, "GIRR", "DELTA") / math.sqrt(2))
    assert cap(lo_variant_run, "FX", "DELTA") == pytest.approx(cap(lo_default_run, "FX", "DELTA") / math.sqrt(2))
    # The FX curvature divisor skips the USD row, flagged as referencing EUR (#144)
    fx = fx_curvature_buckets(lo_variant_run)
    assert fx["GBP"][0] == pytest.approx(15_000 / 1.5) and fx["USD"][1] == pytest.approx(-35_000)
    # The Boolean-formatted 2 on JPY shows TRUE, so that row is not divided either
    assert fx["JPY"][0] == pytest.approx(9_000)


def fx_curvature_buckets(wb):
    """{bucket: (Kb, Sb)} for FX curvature in the medium scenario."""
    return {r[3]: r[4:6] for r in sheet_rows(wb, "Out_Buckets") if r[:3] == ("FX", "CURVATURE", "MEDIUM")}


def test_fx_curvature_divisor_all(lo_default_path, tmp_path):
    # The approved alternative of MAR21.98 divides every FX row, flagged or not (#144)
    wb = openpyxl.load_workbook(lo_default_path)
    for row in wb["Control"].iter_rows(min_row=1, min_col=FIRST_COL):
        if row[0].value == "FX_CURVATURE_DIVISOR_ALL":
            row[1].value = True
    source = tmp_path / "fx-all-in.xlsx"
    wb.save(source)
    out = tmp_path / "fx-all-out.xlsx"
    assert lo_runner.run(out, ["M_Main.RunFRTB_Silent"], input_book=source) == []
    result = openpyxl.load_workbook(out, data_only=True)
    assert reconcile(result) == []
    fx = fx_curvature_buckets(result)
    assert fx["GBP"][0] == pytest.approx(15_000 / 1.5) and fx["USD"][1] == pytest.approx(-35_000 / 1.5)



GIRR_VEGA_ROWS: list[list[Any]] = [
    # MAR21.8 FAQ4: inflation and cross-currency basis vega along the option maturity only (#209)
    ["V1", "GIRR", "EUR", "HICP", "INFLATION", 1, None, 300_000],
    ["V2", "GIRR", "EUR", "FR_CPI", "INFLATION", 5, None, -100_000],
    ["V3", "GIRR", "EUR", "EURUSD", "XCCY", 1, None, 200_000],
    ["V4", "GIRR", "USD", "CPI", "INFLATION", 3, None, 150_000],
    ["V5", "GIRR", "EUR", "HICP", "INFLATION", 1, 5, 1_000],       # an underlying maturity: rejected
    ["V6", "GIRR", "EUR", "", "BASIS", 1, 5, 1_000],              # unknown sub-type: rejected
]


def test_girr_vega_inflation_and_xccy_reconcile(lo_default_run, lo_default_path, tmp_path):
    wb = openpyxl.load_workbook(lo_default_path)
    ws = wb["In_Vega"]
    r = FIRST_DATA_ROW
    while ws.cell(r, FIRST_COL).value not in (None, ""):
        r += 1
    for values in GIRR_VEGA_ROWS:
        for c, v in enumerate(values, start=FIRST_COL):
            ws.cell(r, c).value = v
        r += 1
    source = tmp_path / "girr-vega-in.xlsx"
    wb.save(source)
    out = tmp_path / "girr-vega-out.xlsx"
    assert lo_runner.run(out, ["M_Main.RunFRTB_Silent"], input_book=source) == []
    result = openpyxl.load_workbook(out, data_only=True)
    assert reconcile(result) == []
    errors = sorted(r[2] for r in log_rows(result, "ERROR"))
    assert errors == ["GIRR INFLATION vega takes no underlying maturity; row skipped",
                      "GIRR sub-type must be YIELD, INFLATION or XCCY; row skipped"], errors
    # the valid rows are read as new risk factors: GIRR vega changes in every scenario
    def girr_vega(book):
        return {r[2]: r[3] for r in sheet_rows(book, "Out_RiskClass") if r[:2] == ("GIRR", "VEGA")}
    before, after = girr_vega(lo_default_run), girr_vega(result)
    assert set(after) == {"LOW", "MEDIUM", "HIGH"}
    assert all(after[s] != pytest.approx(before[s]) for s in after)

def test_undefined_vega_and_curvature_sub_types_are_rejected(lo_default_run, lo_default_path, tmp_path):
    wb = openpyxl.load_workbook(lo_default_path)
    for sheet, rows in cases.UNDEFINED_SUB_TYPE_ROWS.items():
        ws = wb[sheet]
        r = FIRST_DATA_ROW
        while ws.cell(r, FIRST_COL).value not in (None, ""):
            r += 1
        for values in rows:
            for c, v in enumerate(values, start=FIRST_COL):
                ws.cell(r, c).value = v
            r += 1
    source = tmp_path / "sub-types-in.xlsx"
    wb.save(source)
    out = tmp_path / "sub-types-out.xlsx"
    assert lo_runner.run(out, ["M_Main.RunFRTB_Silent"], input_book=source) == []
    result = openpyxl.load_workbook(out, data_only=True)
    assert reconcile(result) == []
    errors = sorted(r[2] for r in log_rows(result, "ERROR"))
    assert errors == cases.UNDEFINED_SUB_TYPE_ERRORS, errors
    # every row is skipped: the results are those of the sample, with nine errors
    def summary(book):
        return {r[0]: r[1] for r in sheet_rows(book, "Out_Summary")}
    after, before = summary(result), summary(lo_default_run)
    assert after.pop("ERRORS") == 9 and before.pop("ERRORS") == 0
    assert after == before


# --------------------------------------------------------------------------
# Empty inputs
# --------------------------------------------------------------------------
def test_empty_inputs_give_zero_capital(lo_default_path, tmp_path):
    wb = openpyxl.load_workbook(lo_default_path)
    for sheet in ("In_Delta", "In_Vega", "In_Curvature", "In_DRC", "In_RRAO"):
        ws = wb[sheet]
        ws.delete_rows(FIRST_DATA_ROW, ws.max_row)
    source = tmp_path / "empty-in.xlsx"
    wb.save(source)
    out = tmp_path / "empty-out.xlsx"
    assert lo_runner.run(out, ["M_Main.RunFRTB_Silent"], input_book=source) == []
    result = openpyxl.load_workbook(out, data_only=True)
    summary = {r[0]: r[1] for r in sheet_rows(result, "Out_Summary")}
    assert summary["TOTAL"] == 0 and summary["ERRORS"] == 0
    assert all(r[3] == 0 for r in sheet_rows(result, "Out_RiskClass"))
    assert sheet_rows(result, "Out_Buckets") == []
    # All three scenarios are zero: Results notes the tie and reports LOW (#138)
    assert tie_notes(result) == ["Scenario tie (LOW, MEDIUM, HIGH): same SBM total, LOW reported; "
                                 "ties resolve in the order low, medium, high"]
