"""Worksheet data validation of the template (#348)."""

from __future__ import annotations

import re
import shutil
import zipfile
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

import openpyxl
import pytest
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, range_boundaries

import acceptance_cases as cases
import build_template
import input_validation as iv
import lo_runner
from conftest import require_libreoffice
from test_vba_engine import log_rows

TEMPLATE = build_template.TEMPLATE


def bounds(sqref: str) -> tuple[int, int, int, int]:
    """Left column, top row, right column and bottom row of a range."""
    left, top, right, bottom = range_boundaries(sqref)
    assert left is not None and top is not None and right is not None and bottom is not None
    return left, top, right, bottom


def stored() -> list[iv.Rule]:
    return iv.read_rules(TEMPLATE)


def test_template_carries_the_rules_it_must_carry():
    assert iv.problems(TEMPLATE) == []
    rules = stored()
    assert len(rules) == len(set(rules)) > 80
    assert {r.sheet for r in rules} >= {"Control", *iv.INPUT_SHEETS, "P_General", "P_DRC", "P_RRAO"}


def test_every_input_column_has_a_rule_or_a_reason():
    wb = openpyxl.load_workbook(TEMPLATE)
    rules = stored()
    for sheet in iv.INPUT_SHEETS:
        for cell in wb[sheet][5]:
            if cell.value in (None, "") or cell.column < 3:
                continue
            ruled = any(r.sheet == sheet and bounds(r.sqref)[0] == cell.column for r in rules)
            reason = iv.UNVALIDATED.get((sheet, iv._header_name(cell.value)))
            assert ruled != bool(reason), (sheet, cell.coordinate, cell.value, "rule and reason are exclusive")
    for key in iv.UNVALIDATED:
        assert key[0] in iv.INPUT_SHEETS


def test_input_rules_cover_every_data_row_and_no_other_row():
    first, last = iv._layout()
    assert (first, last) == (6, 205)
    for rule in stored():
        if rule.sheet in iv.INPUT_SHEETS:
            _, row1, _, row2 = bounds(rule.sqref)
            assert (row1, row2) == (first, last), rule


def test_every_numeric_parameter_cell_has_a_rule():
    wb = openpyxl.load_workbook(TEMPLATE)
    rules = stored()
    covered = set()
    for rule in rules:
        if rule.sheet.startswith("P_"):
            c1, r1, c2, r2 = bounds(rule.sqref)
            covered |= {(rule.sheet, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)}
    missing = []
    for ws in wb:
        if not ws.title.startswith("P_"):
            continue
        header = 0
        for row in range(1, ws.max_row + 1):
            key = ws.cell(row, 3).value
            if isinstance(key, str) and key.startswith("["):
                header = row + 1
                continue
            if header and row > header:
                for col in range(3, ws.max_column + 1):
                    value = ws.cell(row, col).value
                    if isinstance(value, (int, float)) and not isinstance(value, bool) and (ws.title, row, col) not in covered:
                        missing.append((ws.title, row, col))
    assert missing == []


def test_formulas_hold_no_array_constant_and_no_unguarded_error():
    for rule in stored():
        if rule.kind == "custom":
            assert "{" not in rule.formula, rule
            for function in ("ABS(", "INT(", "CODE("):
                for match in re.finditer(re.escape(function), rule.formula):
                    assert "IF(" in rule.formula[:match.start()], (rule.sheet, rule.sqref, function)


def test_rule_texts_fit_excel_limits():
    for rule in stored():
        assert len(rule.title) <= 32 and len(rule.message) <= 255, rule
        assert len(rule.formula) <= 255, rule


def test_rules_sit_between_the_autofilter_and_the_page_setup():
    with zipfile.ZipFile(TEMPLATE) as z:
        for name, part in build_template.sheet_parts(TEMPLATE).items():
            if name not in iv.INPUT_SHEETS:
                continue
            tags = re.findall(r"<(autoFilter|dataValidations|pageMargins|pageSetup|headerFooter)[ >/]", z.read(part).decode("utf-8"))
            assert tags == ["autoFilter", "dataValidations", "pageMargins", "pageSetup", "headerFooter"], (name, tags)


def test_applying_the_rules_twice_changes_nothing(tmp_path):
    copy = tmp_path / "copy.xlsx"
    shutil.copy(TEMPLATE, copy)
    before = iv.read_rules(copy)
    iv.apply(copy)
    assert iv.read_rules(copy) == before
    assert build_template.signature_differences(build_template.signature(TEMPLATE), build_template.signature(copy)) == []


def test_a_missing_or_changed_rule_is_reported(tmp_path):
    copy = tmp_path / "copy.xlsx"
    shutil.copy(TEMPLATE, copy)
    with zipfile.ZipFile(copy) as z:
        parts = {i.filename: z.read(i.filename) for i in z.infolist()}
    part = build_template.sheet_parts(copy)["In_Delta"]
    parts[part] = parts[part].replace(b"Invalid risk class", b"Invalid class")
    with zipfile.ZipFile(copy, "w") as z:
        for name, content in parts.items():
            z.writestr(name, content)
    found = iv.problems(copy)
    assert any("missing" in p and "Invalid risk class" in p for p in found), found
    assert any("unexpected" in p and "Invalid class" in p for p in found), found


# --------------------------------------------------------------------------
# The formulas, evaluated by LibreOffice
# --------------------------------------------------------------------------
def evaluate(tmp_path: Path, rule: iv.Rule, cases_: Sequence[Mapping[str, Any]]) -> list[bool]:
    """Evaluate a custom rule's formula for each case: a dict of column letter to value, on rows from the rule's first row."""
    require_libreoffice()
    first_cell = rule.sqref.split(":")[0]
    first_row = int(re.sub(r"\D", "", first_cell))
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    for offset, values in enumerate(cases_):
        row = first_row + offset
        for column, value in values.items():
            ws[f"{column}{row}"] = value
        ws[f"AA{row}"] = Translator("=" + rule.formula, origin=first_cell).translate_formula(f"{re.sub(r'[0-9]+', '', first_cell)}{row}")
    source = tmp_path / "formula-in.xlsx"
    wb.save(source)
    out = tmp_path / "formula-out.xlsx"
    assert lo_runner.run(out, [], input_book=source) == []
    result = openpyxl.load_workbook(out, data_only=True).active
    assert result is not None
    return [result[f"AA{first_row + i}"].value for i in range(len(cases_))]


def rule_of(sheet: str, column: str) -> iv.Rule:
    found = [r for r in stored() if r.sheet == sheet and r.sqref.startswith(f"{column}6")]
    assert len(found) == 1, (sheet, column, found)
    return found[0]


def test_bucket_rule(tmp_path):
    rule = rule_of("In_Delta", "E")
    rows: list[dict[str, Any]] = [
        {"D": "GIRR", "E": "EUR"}, {"D": "girr", "E": " usd "}, {"D": "GIRR", "E": "EURO"}, {"D": "GIRR", "E": 5},
        {"D": "FX", "E": "USD"}, {"D": "FX", "E": "US"},
        {"D": "CSR_NS", "E": 3}, {"D": "EQ", "E": 11}, {"D": "COM", "E": 1}, {"D": "COM", "E": 0},
        {"D": "EQ", "E": 2.5}, {"D": "CSR_NS", "E": "AB"},
        {"D": None, "E": "anything"}, {"D": "COM", "E": None},
    ]
    assert evaluate(tmp_path, rule, rows) == [
        True, True, False, False, True, False, True, True, True, False, False, False, True, True]


@pytest.mark.parametrize("sheet", ["In_Delta", "In_Vega", "In_Curvature"])
def test_sub_type_rule_follows_the_vocabulary_table(tmp_path, sheet):
    rule = rule_of(sheet, "G")
    rows: list[dict[str, Any]] = []
    expected: list[bool] = []
    for risk_class in iv.RISK_CLASSES:
        names = iv.SUBTYPES[sheet].get(risk_class, ())
        for value in (None, *names, "ZZZ"):
            rows.append({"D": risk_class, "G": value})
            expected.append(value is None or value in names or (risk_class == "COM"))
    rows.append({"D": "EQ", "G": " spot "})
    expected.append("SPOT" in iv.SUBTYPES[sheet]["EQ"])
    assert evaluate(tmp_path, rule, rows) == expected


def test_number_rules(tmp_path):
    tenor = rule_of("In_Delta", "H")
    assert evaluate(tmp_path, tenor, [{"H": None}, {"H": 0}, {"H": 2.5}, {"H": -1}, {"H": "5"}, {"H": "x"}]) == [
        True, True, True, False, False, False]
    amount = rule_of("In_Delta", "I")
    assert evaluate(tmp_path, amount, [{"I": None}, {"I": -5}, {"I": "12"}, {"I": "x"}]) == [True, True, False, False]
    gross = rule_of("In_RRAO", "E")
    assert evaluate(tmp_path, gross, [{"E": 100}, {"E": -1}, {"E": None}]) == [True, False, True]


def test_currency_rule(tmp_path):
    rule = [r for r in stored() if r.sheet == "Control" and r.sqref == "D7"][0]
    assert rule.alert == "warning"
    values = ["EUR", "usd", " eur ", "EU", "E", "EURO", "E1R", "E R", 123, None]
    assert evaluate(tmp_path, rule, [{"D": v} for v in values]) == [
        True, True, True, False, False, False, False, False, False, True]


def test_parameter_rules(tmp_path):
    weight = [r for r in stored() if r.sheet == "P_GIRR" and r.title == "Check the risk weight"][0]
    assert evaluate(tmp_path, weight, [{get_column_letter(bounds(weight.sqref)[0]): v} for v in (0.017, 0, -0.1, "x", None)]) == [
        True, True, False, False, True]
    correlation = [r for r in stored() if r.title == "Check the correlation"][0]
    column = get_column_letter(bounds(correlation.sqref)[0])
    assert evaluate(tmp_path, correlation, [{column: v} for v in (0.5, -1, 1, 1.2, "x")]) == [True, True, True, False, False]


# --------------------------------------------------------------------------
# The vocabularies against the run
# --------------------------------------------------------------------------
def _base_rows() -> dict[str, dict[str, list[Any]]]:
    """One valid row per risk class for each SBM sheet, with the sub-type left to fill in (column index 4)."""
    return {
        "In_Delta": {"GIRR": ["", "GIRR", "EUR", "ESTR", None, 1, 100.0], "CSR_NS": ["", "CSR_NS", 1, "ISSUER", None, 1, 100.0],
                     "EQ": ["", "EQ", 1, "NAME", None, None, 100.0], "COM": ["", "COM", 1, "CRUDE", None, 1, 100.0],
                     "FX": ["", "FX", "USD", "", None, None, 100.0]},
        "In_Vega": {"GIRR": ["", "GIRR", "EUR", "ESTR", None, 1, 1, 100.0], "CSR_NS": ["", "CSR_NS", 1, "ISSUER", None, 1, None, 100.0],
                    "EQ": ["", "EQ", 1, "NAME", None, 1, None, 100.0], "COM": ["", "COM", 1, "CRUDE", None, 1, None, 100.0],
                    "FX": ["", "FX", "USD", "", None, 1, None, 100.0]},
        "In_Curvature": {"GIRR": ["", "GIRR", "EUR", "ESTR", None, 10.0, 5.0, None], "CSR_NS": ["", "CSR_NS", 1, "ISSUER", None, 10.0, 5.0, None],
                         "EQ": ["", "EQ", 1, "NAME", None, 10.0, 5.0, None], "COM": ["", "COM", 1, "CRUDE", None, 10.0, 5.0, None],
                         "FX": ["", "FX", "USD", "", None, 10.0, 5.0, None]},
    }


def test_sub_type_vocabulary_matches_the_run(lo_default_path, tmp_path):
    require_libreoffice()
    wb = openpyxl.load_workbook(lo_default_path)
    expected: dict[tuple[str, int], tuple[str, str, bool]] = {}
    for sheet, bases in _base_rows().items():
        ws = wb[sheet]
        row = cases.first_free_row(ws)
        for risk_class, base in bases.items():
            names = iv.SUBTYPES[sheet].get(risk_class, ())
            for value in (*names, "ZZZ"):
                values = list(base)
                values[4] = value
                values[0] = f"V{row}"
                for column, item in enumerate(values, start=cases.FIRST_COL):
                    ws.cell(row, column).value = item
                allowed = value in names or risk_class == "COM"
                expected[(sheet, row)] = (risk_class, value, allowed)
                row += 1
    source = tmp_path / "vocabulary-in.xlsx"
    wb.save(source)
    out = tmp_path / "vocabulary-out.xlsx"
    assert lo_runner.run(out, ["M_Main.RunFRTB_Silent"], input_book=source) == []
    errors = [r for r in log_rows(openpyxl.load_workbook(out, data_only=True), "ERROR") if "sub-type" in str(r[2]).lower()]
    flagged = {(m.group(1), int(m.group(2))) for r in errors if (m := re.match(r"(\w+) row (\d+)", str(r[1])))}
    for (sheet, row), (risk_class, value, allowed) in expected.items():
        assert ((sheet, row) in flagged) == (not allowed), (sheet, risk_class, value, allowed)
    assert any(not allowed for _, _, allowed in expected.values()) and any(allowed for _, _, allowed in expected.values())


def test_risk_class_and_flag_vocabularies_match_the_run(lo_default_path, tmp_path):
    require_libreoffice()
    config = (Path(__file__).resolve().parents[1] / "src/modules/CORE_Config.bas").read_text(encoding="cp1252")
    codes = tuple(re.findall(r'Public Const RC_\w+\s+As String\s+=\s+"(\w+)"', config))
    assert codes == iv.RISK_CLASSES
    wb = openpyxl.load_workbook(lo_default_path)
    ws = wb["In_Curvature"]
    row = cases.first_free_row(ws)
    rows = {}
    for flag in (*iv.FLAGS, "YES"):
        for column, item in enumerate(["F", "GIRR", "EUR", "ESTR", None, 10.0, 5.0, flag], start=cases.FIRST_COL):
            ws.cell(row, column).value = item
        rows[row] = flag
        row += 1
    source = tmp_path / "flags-in.xlsx"
    wb.save(source)
    out = tmp_path / "flags-out.xlsx"
    assert lo_runner.run(out, ["M_Main.RunFRTB_Silent"], input_book=source) == []
    bad = {int(m.group(1)) for r in log_rows(openpyxl.load_workbook(out, data_only=True), "ERROR")
           if (m := re.match(r"In_Curvature row (\d+)", str(r[1]))) and "RefReportingCcy" in str(r[2])}
    assert {r for r, flag in rows.items() if flag not in iv.FLAGS} == bad
