"""Regulatory default parameters typed independently of src/modules/M_Setup.bas.

Used to spot-check the parameter sheets the VBA setup writes and to build
self-contained parameter sets for hand-calculated tests. Values follow the
Basel Framework MAR21-MAR23 (in force 1 January 2023), checked against the
BIS text on 2026-10-17; review them against the regulation that applies.
"""

from __future__ import annotations

import math

from frtb_reference import Params

GIRR_RW = {0.25: 0.017, 0.5: 0.017, 1: 0.016, 2: 0.013, 3: 0.012, 5: 0.011, 10: 0.011, 15: 0.011, 20: 0.011, 30: 0.011}

# bucket: (quality, sector group, type, rw)
CSR_BUCKETS = {
    1: ("IG", "1", "NAME", 0.005), 2: ("IG", "2", "NAME", 0.01), 3: ("IG", "3", "NAME", 0.05),
    4: ("IG", "4", "NAME", 0.03), 5: ("IG", "5", "NAME", 0.03), 6: ("IG", "6", "NAME", 0.02),
    7: ("IG", "7", "NAME", 0.015), 8: ("IG", "8", "NAME", 0.025), 9: ("HY", "1", "NAME", 0.02),
    10: ("HY", "2", "NAME", 0.04), 11: ("HY", "3", "NAME", 0.12), 12: ("HY", "4", "NAME", 0.07),
    13: ("HY", "5", "NAME", 0.085), 14: ("HY", "6", "NAME", 0.055), 15: ("HY", "7", "NAME", 0.05),
    16: ("NA", "16", "OTHER", 0.12), 17: ("IG", "17", "INDEX", 0.015), 18: ("HY", "18", "INDEX", 0.05),
}

# Upper triangle of MAR21.57 Table 5 (sector of buckets 1/9 .. 7/15, 8, 16, 17, 18)
CSR_SECTOR_GROUPS = ["1", "2", "3", "4", "5", "6", "7", "8", "16", "17", "18"]
CSR_SECTOR_UPPER: list[list[float]] = [
    [1, 0.75, 0.10, 0.20, 0.25, 0.20, 0.15, 0.10, 0, 0.45, 0.45],
    [1, 0.05, 0.15, 0.20, 0.15, 0.10, 0.10, 0, 0.45, 0.45],
    [1, 0.05, 0.15, 0.20, 0.05, 0.20, 0, 0.45, 0.45],
    [1, 0.20, 0.25, 0.05, 0.05, 0, 0.45, 0.45],
    [1, 0.25, 0.05, 0.15, 0, 0.45, 0.45],
    [1, 0.05, 0.20, 0, 0.45, 0.45],
    [1, 0.05, 0, 0.45, 0.45],
    [1, 0, 0.45, 0.45],
    [1, 0, 0],
    [1, 0.75],
    [1],
]

# bucket: (type, spot rw, rho between names)
EQ_BUCKETS = {
    1: ("NAME", 0.55, 0.15), 2: ("NAME", 0.60, 0.15), 3: ("NAME", 0.45, 0.15), 4: ("NAME", 0.55, 0.15),
    5: ("NAME", 0.30, 0.25), 6: ("NAME", 0.35, 0.25), 7: ("NAME", 0.40, 0.25), 8: ("NAME", 0.50, 0.25),
    9: ("NAME", 0.70, 0.075), 10: ("NAME", 0.50, 0.125), 11: ("OTHER", 0.70, 0.0),
    12: ("INDEX", 0.15, 0.80), 13: ("INDEX", 0.25, 0.80),
}

# bucket: (rw, rho between commodities)
COM_BUCKETS = {
    1: (0.30, 0.55), 2: (0.35, 0.95), 3: (0.60, 0.40), 4: (0.80, 0.80), 5: (0.40, 0.60), 6: (0.45, 0.65),
    7: (0.20, 0.55), 8: (0.35, 0.45), 9: (0.25, 0.15), 10: (0.35, 0.40), 11: (0.50, 0.15),
}

DRC_RW = {"AAA": 0.005, "AA": 0.02, "A": 0.03, "BBB": 0.06, "BB": 0.15, "B": 0.30, "CCC": 0.50,
          "UNRATED": 0.15, "DEFAULTED": 1.0}
DRC_SENIORITY = {"EQUITY": (1, 1.0), "NON_SENIOR": (2, 1.0), "SENIOR": (3, 0.75), "COVERED": (4, 0.25)}


def eq_gamma(b: int, c: int) -> float:
    if b == c:
        return 1.0
    if 11 in (b, c):
        return 0.0
    if b >= 12 and c >= 12:
        return 0.75
    if b >= 12 or c >= 12:
        return 0.45
    return 0.15


def com_gamma(b: int, c: int) -> float:
    if b == c:
        return 1.0
    return 0.0 if 11 in (b, c) else 0.20


def csr_sector() -> dict[tuple[str, str], float]:
    out: dict[tuple[str, str], float] = {}
    for i, row in enumerate(CSR_SECTOR_UPPER):
        for k, v in enumerate(row):
            j = i + k
            out[(CSR_SECTOR_GROUPS[i], CSR_SECTOR_GROUPS[j])] = float(v)
            out[(CSR_SECTOR_GROUPS[j], CSR_SECTOR_GROUPS[i])] = float(v)
    return out


def build_params() -> Params:
    p = Params()
    # Equity large cap and indices (1-8, 12, 13): 20 days; small cap and other (9-11): 60 days
    p.vega_lh = [("GIRR", "*", 60), ("CSR_NS", "*", 120), ("EQ", "*", 20), ("EQ", "9", 60), ("EQ", "10", 60),
                 ("EQ", "11", 60), ("COM", "*", 120), ("FX", "*", 40)]
    p.vertices = {
        "GIRR_TENORS": [0.25, 0.5, 1, 2, 3, 5, 10, 15, 20, 30],
        "CSR_TENORS": [0.5, 1, 3, 5, 10],
        "COM_TENORS": [0, 0.25, 0.5, 1, 2, 3, 5, 10, 15, 20, 30],
        "VEGA_MATURITIES": [0.5, 1, 3, 5, 10],
    }
    p.girr_rw = {float(k): v for k, v in GIRR_RW.items()}
    p.girr = {"THETA": 0.03, "RHO_FLOOR": 0.40, "RHO_DIFF_CURVE": 0.999, "RHO_INFLATION": 0.40, "RHO_XCCY": 0.0,
              "RW_INFLATION": 0.016, "RW_XCCY": 0.016, "GAMMA": 0.50, "SPECIFIED_DIVISOR": math.sqrt(2)}
    p.girr_specified = {"EUR", "USD", "GBP", "AUD", "JPY", "SEK", "CAD"}
    p.csr_buckets = {str(b): {"quality": q, "group": g, "type": t, "rw": rw} for b, (q, g, t, rw) in CSR_BUCKETS.items()}
    p.csr = {"RHO_NAME": 0.35, "RHO_NAME_INDEX": 0.80, "RHO_TENOR": 0.65, "RHO_BASIS": 0.999, "GAMMA_RATING": 0.50}
    p.csr_sector = csr_sector()
    p.eq_buckets = {str(b): {"type": t, "rw_spot": rw, "rw_repo": rw / 100, "rho": rho} for b, (t, rw, rho) in EQ_BUCKETS.items()}
    p.eq = {"RHO_SPOT_REPO": 0.999}
    p.eq_gamma = {(str(b), str(c)): eq_gamma(b, c) for b in range(1, 14) for c in range(1, 14)}
    p.com_buckets = {str(b): {"rw": rw, "rho": rho} for b, (rw, rho) in COM_BUCKETS.items()}
    p.com = {"RHO_TENOR": 0.99, "RHO_BASIS": 0.999}
    p.com_gamma = {(str(b), str(c)): com_gamma(b, c) for b in range(1, 12) for c in range(1, 12)}
    p.fx = {"RW": 0.15, "GAMMA": 0.60, "SPECIFIED_DIVISOR": math.sqrt(2), "CURVATURE_DIVISOR": 1.5}
    p.fx_pairs = {"USD/EUR", "USD/JPY", "USD/GBP", "USD/AUD", "USD/CAD", "USD/CHF", "USD/MXN", "USD/CNY",
                  "USD/NZD", "USD/RUB", "USD/HKD", "USD/SGD", "USD/TRY", "USD/KRW", "USD/SEK", "USD/ZAR",
                  "USD/INR", "USD/NOK", "USD/BRL", "EUR/JPY", "EUR/GBP", "EUR/CHF", "JPY/AUD"}
    p.drc_rw = dict(DRC_RW)
    p.drc_seniority = dict(DRC_SENIORITY)
    p.drc = {"MIN_MATURITY": 0.25, "MAX_MATURITY": 1.0}
    p.drc_buckets = ["CORP", "SOV", "LGM"]
    p.rrao_rw = {"EXOTIC": 0.01, "OTHER": 0.001}
    return p
