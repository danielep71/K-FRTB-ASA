"""The unattended Excel checks of TEST_Acceptance, run in LibreOffice (#274).

LibreOffice runs every check that does not depend on Excel's window or fonts;
those report SKIPPED here and run on the Windows host through
tools/Check-FRTB-Excel.ps1. The data module the checks read is generated from
the independent reference and must be current.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

import lo_runner

REPO_ROOT = Path(__file__).resolve().parents[1]
HOST_INDEPENDENT = ("SAMPLE_RUN", "FAILED_RUN", "CLEAR_OUTPUTS", "NUMERIC_TEXT", "UNDEFINED_SUB_TYPES",
                    "NO_DATES", "INVALID_ROW_VARIANT")
EXCEL_ONLY = ("FROZEN_PANES", "AUTOFILTERS", "CONTROL_DESCRIPTIONS", "DISPLAYED_VALUES", "SHEET_VIEW", "DATA_VALIDATION",
              "NAVIGATION", "KEY_HIGHLIGHTS", "PROTECTION")


def test_the_acceptance_data_is_current():
    result = subprocess.run(
        [sys.executable, str(REPO_ROOT / "tools" / "build_acceptance_data.py"), "--check"],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.fixture(scope="module")
def acceptance_report(lo_default_path, tmp_path_factory):
    out = tmp_path_factory.mktemp("lo-acceptance") / "acceptance.xlsx"
    errors = lo_runner.run(out, ["TEST_Acceptance.Test_RunAcceptance_Silent"], input_book=lo_default_path)
    assert errors == [], errors
    ws = openpyxl.load_workbook(out, data_only=True)["Acceptance"]
    rows = {r[0]: (r[1], r[2]) for r in ws.iter_rows(min_row=5, min_col=3, max_col=5, values_only=True) if r[0]}
    return ws["C2"].value, rows


def test_every_host_independent_check_passes(acceptance_report):
    result, rows = acceptance_report
    assert result == "RESULT=PASS", rows
    for name in HOST_INDEPENDENT:
        assert rows[name][0] == "PASS", (name, rows[name])


def test_the_excel_only_checks_are_skipped_outside_excel(acceptance_report):
    _, rows = acceptance_report
    assert set(rows) == set(HOST_INDEPENDENT) | set(EXCEL_ONLY)
    for name in EXCEL_ONLY:
        assert rows[name][0] == "SKIPPED", (name, rows[name])


def test_the_log_names_the_row_and_the_reference_of_every_finding(lo_default_path, tmp_path):
    """After the invalid-row variant, the Log carries the sheet row and the ID of each rejected row,
    and the obligor of an obligor finding, with the messages unchanged (#350)."""
    out = tmp_path / "variant.xlsx"
    errors = lo_runner.run(out, ["TEST_Acceptance.Test_RunAcceptance_Silent"], input_book=lo_default_path)
    assert errors == [], errors
    wb = openpyxl.load_workbook(out, data_only=True)
    header = [c.value for c in wb["Log"][5]][2:7]
    assert header == ["Level", "Source", "Message", "Row", "Reference"]
    findings = [r for r in wb["Log"].iter_rows(min_row=6, min_col=3, max_col=7, values_only=True) if r[0] in ("ERROR", "WARN")]
    assert len(findings) == 23, findings  # 22 errors and 1 warning
    for level, source, message, row, reference in findings:
        if " row " in str(source):
            sheet, _, number = str(source).rpartition(" row ")
            number_of_row = int(number)
            assert row == number_of_row, (source, row)
            ids = [c.value for c in wb[sheet]["C"]]  # column C holds the ID of an input row
            expected_id = ids[number_of_row - 1]
            assert reference == (str(expected_id) if expected_id not in (None, "") else None), (source, reference)
        else:
            assert reference, (level, source, message)  # an obligor-level finding names the obligor
