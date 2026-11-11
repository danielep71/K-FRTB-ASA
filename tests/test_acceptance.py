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
HOST_INDEPENDENT = ("SAMPLE_RUN", "FAILED_RUN", "CLEAR_OUTPUTS", "NUMERIC_TEXT", "INVALID_ROW_VARIANT")
EXCEL_ONLY = ("FROZEN_PANES", "AUTOFILTERS", "CONTROL_DESCRIPTIONS")


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
