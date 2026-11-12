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
    check = text.split("} elseif ($runTrustCheck) {", 1)[1].split("} else {", 1)[0]
    assert "$before = Get-UserVbaAccess $excelVersion" in check
    assert re.search(r"try \{\s+Set-VbaAccess \$excelVersion 0", check)
    assert re.search(r"\} finally \{\s+Restore-UserVbaAccess \$excelVersion \$before\s+\}", check)
    assert text.count("Set-VbaAccess $excelVersion") == 2
    # Macros run only in the instances this script starts, never by policy
    assert "AutomationSecurity = $Security" in text
    assert "VBAWarnings" not in text


def test_excel_that_cannot_start_stops_with_exit_2():
    text = script_text()
    probe = text.split("$probe = New-Object -ComObject Excel.Application", 1)
    assert len(probe) == 2
    # the first Excel is created inside a try block whose catch stops with exit 2
    assert probe[0].rstrip().endswith("try {")
    catch = probe[1].split("} catch {", 1)[1].split("\n}", 1)[0]
    assert "Stop-NotStarted" in catch and "Excel could not be started" in catch
    assert "exit 2" in text.split("function Stop-NotStarted {", 1)[1].split("\n}", 1)[0]


def test_a_policy_on_the_vba_access_skips_the_trust_center_check():
    text = script_text()
    # The user's setting is read and restored as found, absent included;
    # the effective value, which a policy can set, is never written back
    get = text.split("function Get-UserVbaAccess {", 1)[1].split("\n}", 1)[0]
    assert "Get-SecurityKey $Version" in get and "Policies" not in get
    assert "return $null" in get
    restore = text.split("function Restore-UserVbaAccess {", 1)[1].split("\n}", 1)[0]
    assert "Remove-ItemProperty" in restore and "Set-VbaAccess $Version" in restore
    assert "(Get-VbaAccess $excelVersion).Value" not in text
    # With a policy, nothing is asked or changed and the check reports why
    ask = text.split("$runTrustCheck = [bool]$TrustCenterCheck", 1)[1].split("Write-Host ''", 1)[0]
    assert ask.index("if ($access.Policy) {") < ask.index("Read-YesNo")
    stage = text.split("# ------------------------------------------------- 4. Trust Center path", 1)[1]
    policy = stage.split("if ($access.Policy) {", 1)[1].split("} elseif ($runTrustCheck) {", 1)[0]
    assert "'TRUST_CENTER_ERROR' 'NOT_RUN'" in policy and "group policy" in policy
    assert "VbaAccess" not in policy


def test_the_script_never_pushes_and_keeps_its_excel_to_itself():
    text = script_text()
    assert not re.search(r"git\s+push", text)
    assert "New-Object -ComObject Excel.Application" in text
    assert "GetActiveObject" not in text  # never attaches to the operator's Excel
    # The watchdog closes only Excel started by automation during the run,
    # never an Excel the operator opened meanwhile
    watchdog = text.split("$watchdog = Start-Job", 1)[1].split("\n}", 1)[0]
    assert "$before -notcontains $_.ProcessId" in watchdog
    assert "[/-]automation" in watchdog and "Stop-Process -Id $_.ProcessId" in watchdog
    assert "Get-Process -Name EXCEL" not in watchdog
    # Past the limit it keeps closing them until the run stops it, and no
    # stage starts after the limit
    assert "while ($true)" in watchdog
    body = text.split("$watchdog = Start-Job", 1)[1]
    stages = [line for line in body.splitlines() if line.strip().startswith(("Write-Host '==", 'Write-Host "=='))]
    assert stages and body.count("Assert-TimeLeft") == len(stages)


def test_a_timed_out_launcher_is_stopped_with_its_processes():
    text = script_text()
    stop = text.split("function Stop-TimedOutLauncher {", 1)[1].split("\n}", 1)[0]
    # the whole tree first, then only automation Excel started since the launcher began
    assert stop.index("taskkill.exe /PID $Process.Id /T /F") < stop.index("Get-AutomationExcelIds")
    assert "$ExcelBefore -notcontains $_" in stop and "Wait-Process" in stop
    automation = text.split("function Get-AutomationExcelIds {", 1)[1].split("\n}", 1)[0]
    assert "[/-]automation" in automation
    loop = text.split("Write-Host '== Synchronization launchers'", 1)[1].split("} catch {", 1)[0]
    assert loop.index("$excelBeforeLauncher = Get-AutomationExcelIds") < loop.index("Start-Process")
    assert "Stop-TimedOutLauncher $process $excelBeforeLauncher" in loop
    assert "$process.Kill()" not in loop


def test_the_launcher_exit_code_is_read():
    # Without reading the handle while the launcher runs, a Process from
    # Start-Process loses its exit code and the check reports FAIL (#297)
    loop = script_text().split("Write-Host '== Synchronization launchers'", 1)[1].split("} catch {", 1)[0]
    start = loop.index("Start-Process")
    assert start < loop.index("$null = $process.Handle") < loop.index("$process.WaitForExit(180000)")
    assert loop.index("$process.WaitForExit(180000)") < loop.index("$process.WaitForExit()") \
        < loop.index("Select-String -Path $logFile")
    assert "$process.ExitCode -eq 0 -and $succeeded" in loop


def test_each_launcher_is_judged_by_its_own_success_line():
    text = script_text()
    for launcher in ("Sync-FRTB-VBA.cmd", "Sync-FRTB-VBA-Downloads.cmd"):
        source = (ROOT / "tools" / launcher).read_text(encoding="ascii")
        success = next(line.strip()[len("echo "):] for line in source.splitlines() if "echo SUCCESS:" in line)
        assert f"Name = '{launcher}'" in text and f"Success = '{success}'" in text, (launcher, success)


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


def test_the_build_turns_alerts_off_again_before_it_saves():
    # BuildWorkbook and FinishWorkbook end by turning Excel alerts back on;
    # a hidden Excel then waits on the macro-free question of the .xlsx
    # template save, which nobody can answer (#302)
    build = (ROOT / "tools" / "Build-Workbook.ps1").read_text(encoding="ascii")
    last_macro = max(build.index("'BuildWorkbook')"), build.index("'FinishWorkbook')"),
                     build.index("'Test_RunRegression_Silent')"))
    again = build.index("$excel.DisplayAlerts = $false", last_macro)
    assert again < build.index("$workbook.SaveAs($OutputPath, 51)")
    assert again < build.index("$workbook.SaveAs($OutputPath, 52)")


def test_the_default_printer_is_switched_for_the_run_and_restored():
    # A slow default printer made Excel wait in every build; the run uses
    # Microsoft Print to PDF and puts the operator's printer back (#304)
    text = script_text()
    body = text.split("$watchdog = Start-Job", 1)[1]
    run = body.split("\ntry {", 1)[1]
    switch = run.index("Set-DefaultPrinter $pdfPrinter")
    assert switch < run.index("Write-Host '== Recorded run")
    final = run.split("\n} finally {", 1)[1].split("\n}", 1)[0]
    assert "if ($printerChanged)" in final and "Set-DefaultPrinter $printerBefore" in final
    assert "$pdfPrinter = 'Microsoft Print to PDF'" in text
    # without the PDF printer nothing changes; the log names no printer of the operator
    assert "default printer unchanged: $pdfPrinter is not installed" in run
    assert not [line for line in text.splitlines() if "Write-Log" in line and "$printerBefore" in line]
    escape = text.split("function Set-DefaultPrinter {", 1)[1].split("\n}", 1)[0]
    assert r".Replace('\', '\\')" in escape and r""".Replace("'", "\'")""" in escape


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
