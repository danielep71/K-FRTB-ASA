"""The unattended Excel check, tools/Check-FRTB-Excel.ps1, and its helper (#274).

Excel itself runs only on the Windows host. Here the script must parse, stop
with exit code 2 before it changes anything when a tool is missing, and keep
the promises its help makes; tools/compare_templates.py is run on real
templates.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
from pathlib import Path

import openpyxl
import pytest

import compare_templates

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "tools" / "Check-FRTB-Excel.ps1"
LAUNCHER = ROOT / "tools" / "Check-FRTB-Excel.cmd"
TEMPLATE = ROOT / "src" / "workbook" / "FRTB_Template.xlsx"
PWSH = os.environ.get("PWSH") or shutil.which("pwsh")


def script_text() -> str:
    return SCRIPT.read_text(encoding="ascii")


@pytest.mark.skipif(not PWSH, reason="PowerShell 7 (pwsh) is not installed")
def test_the_script_parses():
    command = (
        "$e = $null; $t = $null; [void][System.Management.Automation.Language.Parser]::ParseFile("
        f"'{SCRIPT}', [ref]$t, [ref]$e); if ($e.Count) {{ $e | ForEach-Object {{ $_.ToString() }}; exit 1 }}"
    )
    completed = subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-Command", command],
                               capture_output=True, text=True, timeout=120, check=False)
    assert completed.returncode == 0, completed.stdout + completed.stderr


@pytest.mark.skipif(not PWSH, reason="PowerShell 7 (pwsh) is not installed")
def test_a_missing_tool_stops_with_exit_2_before_anything_is_written(tmp_path):
    tools = tmp_path / "tools"
    tools.mkdir()
    shutil.copy(SCRIPT, tools)
    empty = tmp_path / "empty"
    empty.mkdir()
    completed = subprocess.run(
        [str(PWSH), "-NoProfile", "-NonInteractive", "-File", str(tools / SCRIPT.name), "-NoPrompt"],
        env={**os.environ, "PATH": str(empty)},
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    assert completed.returncode == 2, completed.stdout + completed.stderr
    assert "NOT STARTED: git is not on PATH." in completed.stdout + completed.stderr
    assert not (tmp_path / "test-results").exists()


def test_the_trust_center_is_changed_only_on_request_and_restored():
    text = script_text()
    # The access is turned on only after a Y answer, and the error check
    # restores the value it found, even when the build fails
    turn_on = text.index("Set-VbaAccess $excelVersion 1")
    assert text.rindex("Read-YesNo", 0, turn_on) > text.rindex("if ($access.Value -ne 1)", 0, turn_on)
    check = text.split("if ($runTrustCheck) {", 1)[1].split("} else {", 1)[0]
    assert re.search(r"try \{\s+Set-VbaAccess \$excelVersion 0", check)
    assert re.search(r"\} finally \{\s+Set-VbaAccess \$excelVersion \$before\s+\}", check)
    assert text.count("Set-VbaAccess $excelVersion") == 3
    # Macros run only in the instances this script starts, never by policy
    assert "AutomationSecurity = $Security" in text
    assert "VBAWarnings" not in text


def test_the_script_never_pushes_and_keeps_its_excel_to_itself():
    text = script_text()
    assert not re.search(r"git\s+push", text)
    assert "New-Object -ComObject Excel.Application" in text
    assert "GetActiveObject" not in text  # never attaches to the operator's Excel
    assert "Stop-Process -Force" in text and "$before -notcontains $_.Id" in text


def test_the_acceptance_entry_point_is_the_excel_one():
    text = script_text()
    assert "TEST_Acceptance.Test_RunAcceptanceInExcel_Silent" in text
    module = (ROOT / "tests" / "modules" / "TEST_Acceptance.bas").read_text(encoding="cp1252")
    assert "Public Sub Test_RunAcceptanceInExcel_Silent()" in module


def test_the_launcher_passes_the_branch_and_reports_each_exit_code():
    launcher = LAUNCHER.read_text(encoding="ascii")
    assert '-File "%SCRIPT%" -Branch "%BRANCH%"' in launcher
    assert 'if not defined BRANCH set "BRANCH=release/v0.2.0"' in launcher
    for code, word in (("0", "PASS"), ("1", "FINDINGS")):
        assert f'"%RC%"=="{code}" (' in launcher and f"echo {word}:" in launcher
    assert "echo NOT STARTED:" in launcher


def test_identical_templates_have_no_differences():
    wb = openpyxl.load_workbook(TEMPLATE)
    assert compare_templates.compare(wb, openpyxl.load_workbook(TEMPLATE)) == []


def test_template_differences_are_listed_by_kind(tmp_path):
    other = openpyxl.load_workbook(TEMPLATE)
    ws = other["Control"]
    ws["E7"].value = "changed"
    ws.column_dimensions["E"].width = 20
    ws.freeze_panes = None
    path = tmp_path / "other.xlsx"
    other.save(path)
    lines = compare_templates.compare(openpyxl.load_workbook(TEMPLATE), openpyxl.load_workbook(path))
    assert "Control: 1 values differ" in lines
    assert any(line.startswith("  Control!E7: ") and line.endswith("-> 'changed'") for line in lines)
    assert any(line.startswith("  Control!E: width ") for line in lines)
    assert "Control: frozen panes 6 rows, 0 columns -> none" in lines


def test_compare_templates_prints_the_result(tmp_path, capsys):
    assert compare_templates.main([str(TEMPLATE), str(TEMPLATE)]) == 0
    assert "no differences" in capsys.readouterr().out
    assert compare_templates.main([str(TEMPLATE), str(tmp_path / "missing.xlsx")]) == 2
