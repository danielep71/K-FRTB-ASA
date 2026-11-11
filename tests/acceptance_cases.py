"""The invalid-row variant shared by the LibreOffice tests and the Excel checks.

The variant switches the three discretions on and adds invalid rows to every
input sheet (#28). ``tests/test_vba_engine.py`` runs it in LibreOffice, and
``tools/build_acceptance_data.py`` writes the same rows and their reference
results into ``tests/modules/TEST_AcceptanceData.bas`` for the unattended
Excel checks (#274).
"""

from __future__ import annotations

from typing import Any

import openpyxl

FIRST_COL = 3  # tables start in column C
FIRST_DATA_ROW = 6  # header in row 5

DISCRETIONS = ("GIRR_SPECIFIED_CCY_DISCRETION", "FX_SPECIFIED_PAIR_DISCRETION", "FX_CURVATURE_DIVISOR")

INVALID_ROWS: dict[str, list[list[Any]]] = {
    "In_Delta": [
        ["X1", "CSR_SEC_CTP", 1, "A", "BOND", 5, 1000],
        ["X2", "GIRR", "EUR", "ESTR", "YIELD", 7, 1000],
        ["X3", "FX", "EUR", "EUR", None, None, 1000],
        ["X4", "EQ", 99, "A", "SPOT", None, 1000],
        ["X5", "EQ", 8, "A", "SPOT", None, "abc"],
        ["X6", "COM", 2, "BRENT", "ROTTERDAM", "abc", 1000],
        ["X7", "EQ", 8.6, "A", "SPOT", None, 1000],
    ],
    "In_Vega": [["Y1", "EQ", 8, "A", None, 2, None, 1000]],
    "In_Curvature": [
        ["Z1", "EQ", 8, "A", None, 1000, None],
        # RefReportingCcy accepts only TRUE, FALSE or blank; aliases are rejected (#144)
        ["Z2", "FX", "USD", "USD", None, 1000, 1000, "maybe"],
        ["Z3", "FX", "USD", "USD", None, 1000, 1000, "no"],
        ["Z4", "FX", "GBP", "GBP", None, 1000, 1000, "YES"],
        ["Z5", "FX", "GBP", "GBP", None, 1000, 1000, 1],
    ],
    "In_DRC": [
        ["K1", "A", "CORP", "ZZZ", "SENIOR", 1, 1, 1, None],
        ["K2", "A", "XXX", "A", "SENIOR", 1, 1, 1, None],
        ["K3", "A", "CORP", "A", "JUNK", 1, 1, 1, None],
        ["K4", "BANK_A", "CORP", "BBB", "SENIOR", 100, 100, 1, None],
        ["K5", "BANK_C", "CORP", "A", "SENIOR", 1000, 1000, "abc", None],
        ["K6", "BANK_C", "CORP", "A", "SENIOR", 1000, 1000, 1, "x"],
        ["K7", "BANK_D", "CORP", "A", "EQUITY", 1000, 1000, 0.5, None],
        ["K8", "BANK_B", "SOV", "BBB", "SENIOR", 100, 100, 1, None],
    ],
    "In_RRAO": [["Q1", "FOO", 1000], ["Q2", "OTHER", -5]],
}

# A Boolean-formatted 2 shows TRUE: the VBA and the reference read it as TRUE,
# so FX_CURVATURE_DIVISOR leaves the row undivided (#144). LibreOffice's VBA
# already returns a Boolean for this cell; the numeric path of CellBool is the
# one Excel takes (#28).
BOOLEAN_ROW = ["Z6", "FX", "JPY", "JPY", None, 9000, 6000, 2]
BOOLEAN_COLUMN = 8  # References reporting ccy, the eighth column of In_Curvature
BOOLEAN_FORMAT = '"TRUE";"TRUE";"FALSE"'

# Every invalid row is logged, BANK_B adds one error for its split buckets,
# and BANK_A's two credit qualities give the one warning.
EXPECTED_ERRORS = 22
EXPECTED_WARNINGS = 1


def first_free_row(ws: Any) -> int:
    """Return the first row below the sample whose first two columns are empty."""
    r = FIRST_DATA_ROW
    while ws.cell(r, FIRST_COL).value not in (None, "") or ws.cell(r, FIRST_COL + 1).value not in (None, ""):
        r += 1
    return r


def apply_variant(wb: openpyxl.Workbook) -> None:
    """Switch the discretions on and append the invalid rows, as the Excel check does."""
    for row in wb["Control"].iter_rows(min_row=1, min_col=FIRST_COL):
        if row[0].value in DISCRETIONS:
            row[1].value = True
    for sheet, rows in INVALID_ROWS.items():
        ws = wb[sheet]
        r = first_free_row(ws)
        for values in rows:
            for c, v in enumerate(values, start=FIRST_COL):
                ws.cell(r, c).value = v
            r += 1
        if sheet == "In_Curvature":
            for c, v in enumerate(BOOLEAN_ROW, start=FIRST_COL):
                ws.cell(r, c).value = v
            ws.cell(r, FIRST_COL + BOOLEAN_COLUMN - 1).number_format = BOOLEAN_FORMAT
