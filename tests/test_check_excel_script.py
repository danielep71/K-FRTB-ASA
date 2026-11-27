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
    assert 'if not defined BRANCH set "BRANCH=release/v0.3.0"' in launcher
    for code, word in (("0", "PASS"), ("1", "FINDINGS")):
        assert f'"%RC%"=="{code}" (' in launcher and f"echo {word}:" in launcher
    assert "echo NOT STARTED:" in launcher


def test_the_build_turns_alerts_off_again_before_it_saves():
    # BuildWorkbook and FinishWorkbook end by turning Excel alerts back on;
    # a hidden Excel then waits on the macro-free question of the .xlsx
    # template save, which nobody can answer (#302)
    build = (ROOT / "tools" / "Build-Workbook.ps1").read_text(encoding="ascii")
    last_macro = max(build.index("'BuildWorkbook')"), build.index("'FinishWorkbook')"),
                     build.index("'AddNavigationBars')"), build.index("'Test_RunRegression_Silent')"))
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
    assert "$wql = ConvertTo-WqlLiteral $Name" in text


@pytest.mark.skipif(not PWSH, reason="PowerShell 7 (pwsh) is not installed")
def test_printer_names_are_escaped_for_wql():
    # The script's own function, run on a shared printer and a quote (#304)
    command = (
        "$t = $null; $e = $null; $ast = [System.Management.Automation.Language.Parser]::ParseFile("
        f"'{SCRIPT}', [ref]$t, [ref]$e); "
        "$f = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] "
        "-and $n.Name -eq 'ConvertTo-WqlLiteral' }, $true) | Select-Object -First 1; "
        "Invoke-Expression $f.Extent.Text; ConvertTo-WqlLiteral \"\\\\server\\queue O'Brien\""
    )
    completed = subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-Command", command],
                               capture_output=True, text=True, timeout=120, check=False)
    assert completed.returncode == 0, completed.stdout + completed.stderr
    assert completed.stdout.strip() == r"\\\\server\\queue O\'Brien"


def test_every_sheet_is_pictured_before_and_after_a_sample_run():
    # The review of #206 needs each sheet as Excel shows it (#316)
    text = script_text()
    stage = text[text.index("== Sheet pictures"):text.index("# ------------------------------------------------------------ 3. builds")]
    assert "foreach ($stage in @('before', 'after'))" in stage
    assert "M_Main.RunFRTB_Silent" in stage
    # The development workbook is saved with results: the before pictures
    # need Clear outputs first
    assert "if ($stage -eq 'before') { [void](Invoke-Macro $excel $workbook 'M_Main.ClearOutputs') }" in stage
    assert "foreach ($sheet in @($workbook.Worksheets))" in stage
    assert "Get-PictureAddress $sheet" in stage
    assert "Add-Result 'SHEET_PICTURES'" in stage
    assert "$sheetPictureRows = 60" in text
    assert "$sheetPictureColumns = 26" in text


def test_an_empty_paste_is_retried_and_never_counted():
    # Chart.Paste sometimes pastes nothing after CopyPicture; a blank picture
    # must be retried, and reported as not taken if it stays blank (#324)
    text = script_text()
    export = text.split("function Export-SheetPicture {", 1)[1].split("function Test-LogMentions {", 1)[0]
    assert "for ($attempt = 1; $attempt -le $pictureAttempts; $attempt++)" in export
    assert "if ([int]$chartObject.Chart.Shapes.Count -gt 0) { break }" in export
    empty = export.index("if ([int]$chartObject.Chart.Shapes.Count -eq 0) {")
    assert export.index("return $false", empty) < export.index("[void]$chartObject.Chart.Export($Path)")
    assert re.search(r"\} finally \{\s+if \(\$null -ne \$chartObject\)", export)
    assert "$pictureAttempts = 4" in text
    # The chart is created right of the range, never over the cells pictured
    assert "ChartObjects().Add([double]$range.Left + [double]$range.Width + 20, [double]$range.Top," in export
    assert "ChartObjects().Add(0, 0" not in text


def test_the_trust_center_message_is_found_across_line_breaks():
    text = script_text()
    check = text.split("} elseif ($runTrustCheck) {", 1)[1].split("'not requested'", 1)[0]
    assert "Test-LogMentions (Join-Path $script:bundle 'build-trust-center.log')" in check
    assert "Select-String" not in check


def run_script_function(name, body):
    # Defines one function of the script in pwsh and runs the body after it
    command = (
        "$t = $null; $e = $null; $ast = [System.Management.Automation.Language.Parser]::ParseFile("
        f"'{SCRIPT}', [ref]$t, [ref]$e); "
        "$f = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] "
        f"-and $n.Name -eq '{name}' }}, $true) | Select-Object -First 1; "
        f"Invoke-Expression $f.Extent.Text; {body}"
    )
    return subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-Command", command],
                          capture_output=True, text=True, timeout=120, check=False)


@pytest.mark.skipif(not PWSH, reason="PowerShell 7 (pwsh) is not installed")
def test_a_wrapped_message_is_found_and_a_missing_one_is_not(tmp_path):
    # The log of November 15, 2026 wrapped the message at the console width
    wrapped = tmp_path / "wrapped.log"
    wrapped.write_text("Excel denied access to the VBA project. Enable 'Trust access to the VBA project \r\n"
                       "object model' and retry.\r\nEXIT=1\r\n", encoding="utf-8-sig")
    other = tmp_path / "other.log"
    other.write_text("Excel could not start.\r\nEXIT=1\r\n", encoding="utf-8-sig")
    phrase = "'Trust access to the VBA project object model'"
    body = (f"(Test-LogMentions '{wrapped}' {phrase}), (Test-LogMentions '{other}' {phrase}), "
            f"(Test-LogMentions '{tmp_path / 'absent.log'}' {phrase}) -join ','")
    completed = run_script_function("Test-LogMentions", body)
    assert completed.returncode == 0, completed.stdout + completed.stderr
    assert completed.stdout.strip() == "True,False,False"


@pytest.mark.skipif(not PWSH, reason="PowerShell 7 (pwsh) is not installed")
def test_column_numbers_become_column_names():
    command = (
        "$t = $null; $e = $null; $ast = [System.Management.Automation.Language.Parser]::ParseFile("
        f"'{SCRIPT}', [ref]$t, [ref]$e); "
        "$f = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] "
        "-and $n.Name -eq 'ConvertTo-ColumnName' }, $true) | Select-Object -First 1; "
        "Invoke-Expression $f.Extent.Text; (1, 26, 27, 52, 703 | ForEach-Object { ConvertTo-ColumnName $_ }) -join ','"
    )
    completed = subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-Command", command],
                               capture_output=True, text=True, timeout=120, check=False)
    assert completed.returncode == 0, completed.stdout + completed.stderr
    assert completed.stdout.strip() == "A,Z,AA,AZ,AAA"


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
    assert "Control: frozen panes 6 rows, 2 columns -> none" in lines


def test_compare_templates_prints_the_result(tmp_path, capsys):
    assert compare_templates.main([str(TEMPLATE), str(TEMPLATE)]) == 0
    assert "no differences" in capsys.readouterr().out
    assert compare_templates.main([str(TEMPLATE), str(tmp_path / "missing.xlsx")]) == 2


QUICK_LAUNCHER = ROOT / "tools" / "Check-FRTB-Excel-Quick.cmd"


def test_the_quick_profile_builds_the_development_workbook_and_runs_the_acceptance_checks_only():
    text = script_text()
    assert "[switch]$Quick" in text and "quick profile, not evidence" in text
    quick = text.replace("\r\n", "\n").split("    if ($Quick) {\n        Assert-TimeLeft\n        Write-Host '== Development build", 1)[1].split("    } else {", 1)[0]
    assert "Build-Workbook.ps1" in quick and "'-RunTests'" in quick and "Record-FRTB-ExcelRun" not in quick
    flat = text.replace("\r\n", "\n")
    # The acceptance checks run in both profiles; everything after them but the summary is the full profile's
    assert flat.index("TEST_Acceptance.Test_RunAcceptanceInExcel_Silent") < flat.index("    if (-not $Quick) {")
    block = flat.split("    if (-not $Quick) {", 1)[1].split("    } # end of the full profile", 1)[0]
    for stage in ("== Sheet pictures", "== Build from source with the regression test", "== Template built by Excel",
                  "== Trust Center error path", "== Synchronization: $mode", "LAUNCHER_"):
        assert stage in block, stage
    assert "NOT_RUN_IN_QUICK" in flat


def test_the_quick_profile_asks_nothing_about_the_trust_center_and_is_not_recorded_as_evidence():
    flat = script_text().replace("\r\n", "\n")
    # With the VBA project access off it stops instead of asking to turn it on
    access = flat.split("if ($access.Value -ne 1) {", 1)[1].split("Read-YesNo", 1)[0]
    assert "if ($Quick) {" in access and "Stop-NotStarted" in access
    assert "elseif (-not $runTrustCheck -and -not $Quick) {" in flat
    assert "if (-not $Quick) { Write-Host \"Recorded evidence:" in flat
    assert '"$short-quick"' in flat


def test_the_quick_launcher_passes_the_branch_and_the_switch_and_reports_each_exit_code():
    text = QUICK_LAUNCHER.read_text(encoding="ascii")
    assert '-File "%SCRIPT%" -Branch "%BRANCH%" -Quick' in text
    for line in ('if not defined BRANCH set "BRANCH=release/v0.3.0"', 'if "%RC%"=="0"', 'if "%RC%"=="1"', "exit /b %RC%"):
        assert line in text


WORKFLOW = ROOT / ".github" / "workflows" / "excel-check.yml"


def test_the_self_hosted_workflow_runs_only_for_my_pushes_and_publishes_nothing():
    """The runner executes the commit's code on my computer: no pull request trigger, only my
    pushes, no secret, read-only permission, no artifact (the bundle names the computer) and a
    checkout that keeps the bundles."""
    import yaml

    document = yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))
    events = document[True] if True in document else document["on"]  # YAML reads "on" as a boolean key
    assert set(events) == {"push", "workflow_dispatch"}
    assert events["push"] == {"branches": ["main", "release/**"]}
    assert document["permissions"] == {"contents": "read"}
    job = document["jobs"]["excel"]
    assert job["runs-on"] == ["self-hosted", "windows", "excel"]
    assert "github.actor == github.repository_owner" in job["if"] and "github.repository ==" in job["if"]
    text = WORKFLOW.read_text(encoding="utf-8")
    assert "secrets." not in text and "upload-artifact" not in text
    assert "pull_request" not in text.replace("never runs for a pull request", "")
    # The step runs in cmd and starts a script with -ExecutionPolicy Bypass: the host's policy refuses
    # the temporary script the runner writes for a PowerShell step
    step = job["steps"][1]
    assert step["shell"] == "cmd" and "-ExecutionPolicy Bypass -File tools\\Run-Excel-Check-CI.ps1" in step["run"]
    wrapper = (ROOT / "tools" / "Run-Excel-Check-CI.ps1").read_text(encoding="ascii")
    assert "-NoPrompt" in wrapper and "-Commit $env:GITHUB_SHA" in wrapper and "-Branch $env:BRANCH" in wrapper
    assert "exit $code" in wrapper and "GITHUB_STEP_SUMMARY" in wrapper
    assert job["steps"][0]["with"]["clean"] is False  # the default git clean would delete test-results
    # The untrusted branch name reaches the script through an environment variable, not the command line text
    assert "${{ github.ref_name }}" not in job["steps"][1]["run"]
    assert (ROOT / ".github" / "actionlint.yaml").read_text(encoding="utf-8").count("excel") >= 1


VERIFY_LAUNCHER = ROOT / "tools" / "Verify-FRTB-Branch.cmd"


def test_the_verify_launcher_asks_for_the_branch_and_the_mode_and_passes_both_to_the_check():
    text = VERIFY_LAUNCHER.read_text(encoding="ascii")
    assert '-File "%SCRIPT%" -Branch "%BRANCH%" %QUICK%' in text
    assert 'set "DEFAULT_BRANCH=release/v0.3.0"' in text
    # The branch and the mode are asked for only when the file is double-clicked
    assert 'if not defined BRANCH (' in text and 'set /p "BRANCH=' in text and 'set /p "MODE=' in text
    # Quotes are removed from what was typed before it reaches the command line
    assert 'set "BRANCH=%BRANCH:"=%"' in text and 'set "MODE=%MODE:"=%"' in text
    assert text.count('set "QUICK=-Quick"') == 4
    for line in ('if "%RC%"=="0"', 'if "%RC%"=="1"', "exit /b %RC%"):
        assert line in text


def test_the_ribbon_stage_is_in_the_full_profile_only_and_judges_each_step_by_itself():
    flat = script_text().replace("\r\n", "\n")
    full = flat.split("    if (-not $Quick) {", 1)[1].split("    } # end of the full profile", 1)[0]
    stage = full.split("# ------------------------------------------------------ 3b. ribbon tab", 1)[1].split("# ------------------------------------------------- 4. Trust Center path", 1)[0]
    # The build adds the part, and the file is judged by the control ids of the source part
    assert "'-Ribbon', '-RunTests'" in stage and "src\\ribbon\\customUI14.xml" in stage
    for name in ("BUILD_RIBBON", "RIBBON_CALLBACKS", "RIBBON_KEPT_BY_EXCEL", "RIBBON_ROUND_TRIP", "RIBBON_SYNC"):
        assert f"'{name}' $(if" in stage or f"'{name}' 'FAIL'" in stage, name
    # The comparison takes its mode in lower case, as compare_vba_export.py declares it
    assert "compare_vba_export.py --export $export --mode development --candidate-sha" in stage
    # What needs a person is never reported as passed
    assert "'RIBBON_TAB_SHOWN' 'NOT_RUN'" in stage
    # Excel saves the workbook as macro-enabled (52) and the check finds the part in the saved file
    assert "SaveAs($savedCopy, 52)" in stage and "Test-RibbonPart $savedCopy" in stage
    # The relationship is read as XML and must lead to the part whose ids are compared
    assert "/ui/extensibility$" in flat and "$archive.GetEntry($target)" in flat and "-match 'ui/extensibility'" not in flat
    # The quick profile has none of it
    assert "RIBBON" not in flat.split("    if (-not $Quick) {", 1)[0].split("== Acceptance checks in Excel", 1)[1]


def test_the_macros_the_ribbon_stage_runs_exist_with_the_signatures_it_passes():
    flat = script_text().replace("\r\n", "\n")
    main = (ROOT / "src" / "modules" / "M_Main.bas").read_text(encoding="utf-8")
    setup = (ROOT / "src" / "modules" / "M_Setup.bas").read_text(encoding="utf-8")
    assert "Public Sub Ribbon_ToggleNav(ByVal control As Object, ByVal pressed As Boolean)" in main
    assert "Public Sub Ribbon_Clear(ByVal control As Object)" in main
    assert "Public Function NavigationVisible() As Boolean" in setup
    for call in ("'M_Main.Ribbon_ToggleNav' @($workbook, $false)", "'M_Main.Ribbon_ToggleNav' @($workbook, $true)",
                 "'M_Main.Ribbon_Clear' @($workbook)", "'M_Setup.NavigationVisible'"):
        assert call in flat, call
