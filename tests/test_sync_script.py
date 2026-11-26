"""Offline tests of tools/Sync-FRTB-VBA.ps1: the functions that need no Excel.

The functions are loaded from the script's syntax tree and run in PowerShell
7 (pwsh, or the executable named by the PWSH environment variable); the
tests are skipped where it is not installed. Nothing here starts Excel or
touches a workbook.
"""

import json
import os
import shutil
import subprocess
import time
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = REPO_ROOT / "tools" / "Sync-FRTB-VBA.ps1"
PWSH = os.environ.get("PWSH") or shutil.which("pwsh")

# Defines every function of the script, and the $NewModulePattern it uses,
# without running the script body.
LOADER = r"""
$ErrorActionPreference = 'Stop'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($env:SYNC_SCRIPT, [ref]$null, [ref]$null)
foreach ($f in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    . ([scriptblock]::Create($f.Extent.Text))
}
foreach ($a in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $false)) {
    if ($a.Left.Extent.Text -eq '$NewModulePattern') { . ([scriptblock]::Create($a.Extent.Text)) }
}
"""

needs_pwsh = pytest.mark.skipif(not PWSH, reason="PowerShell 7 (pwsh) is not installed")


def run_pwsh(tmp_path, body):
    """Run body after the loader; body writes a JSON result to $env:RESULT."""
    result = tmp_path / "result.json"
    completed = subprocess.run(
        [str(PWSH), "-NoProfile", "-NonInteractive", "-Command", LOADER + body],
        env={**os.environ, "SYNC_SCRIPT": str(SCRIPT), "RESULT": str(result)},
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr + completed.stdout
    return json.loads(result.read_text(encoding="utf-8-sig"))


def module(name, body="Option Explicit\r\n"):
    return f'Attribute VB_Name = "{name}"\r\n{body}'.encode("cp1252")


@needs_pwsh
def test_script_parses_without_errors(tmp_path):
    errors = run_pwsh(
        tmp_path,
        "$e = $null; $t = $null\n"
        "[void][System.Management.Automation.Language.Parser]::ParseFile($env:SYNC_SCRIPT, [ref]$t, [ref]$e)\n"
        "ConvertTo-Json -InputObject @($e | ForEach-Object { $_.ToString() }) | Set-Content -LiteralPath $env:RESULT",
    )
    assert errors == []


@needs_pwsh
@pytest.mark.parametrize(
    ("content", "name", "problem"),
    [
        (module("CORE_SBM"), "CORE_SBM", ""),
        (b"Option Explicit\r\n", "", "no Attribute VB_Name line"),
        (module("CORE_SBM") + b'Attribute VB_Name = "CORE_SBM"\r\n', "", "several Attribute VB_Name lines"),
        (b"Attribute VB_Name = CORE_SBM\r\n", "", "malformed Attribute VB_Name line"),
        (module("1CORE"), "", "is not a valid VBA name"),
    ],
)
def test_declared_name_is_the_only_identity(tmp_path, content, name, problem):
    source = tmp_path / "CORE_SBM.bas"
    source.write_bytes(content)
    declared = run_pwsh(
        tmp_path,
        f"Get-DeclaredVbaName -SourcePath '{source}' | ConvertTo-Json | Set-Content -LiteralPath $env:RESULT",
    )
    assert declared["Name"] == name
    assert problem in declared["Problem"]
    assert bool(declared["Problem"]) == bool(problem)


@needs_pwsh
def test_folder_selection_takes_repository_and_new_modules_only(tmp_path):
    folder = tmp_path / "downloads"
    folder.mkdir()
    (folder / "CORE_SBM.bas").write_bytes(module("CORE_SBM", "' older\r\n"))
    newer = folder / "CORE_SBM (1).bas"
    newer.write_bytes(module("CORE_SBM", "' newer\r\n"))
    stamp = time.time()
    os.utime(folder / "CORE_SBM.bas", (stamp - 60, stamp - 60))
    os.utime(newer, (stamp, stamp))
    (folder / "renamed.bas").write_bytes(module("CORE_Util"))
    (folder / "TEST_New.bas").write_bytes(module("TEST_New"))
    (folder / "Helper.bas").write_bytes(module("Helper"))
    (folder / "M_Main.bas").write_bytes(b"Option Explicit\r\n")
    (folder / "notes.txt").write_text("not a module")

    found = run_pwsh(
        tmp_path,
        "$names = @{ 'core_sbm' = 'x'; 'core_util' = 'x'; 'm_main' = 'x' }\n"
        f"$r = Select-FolderModules -FolderPath '{folder}' -RepositoryNames $names\n"
        "@{ Selection = @($r.Selection | ForEach-Object { @{ Name = $_.Name; File = $_.File.Name; IsNew = $_.IsNew } });"
        " Ignored = @($r.Ignored) } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $env:RESULT",
    )
    selection = {item["Name"]: item for item in found["Selection"]}
    assert set(selection) == {"CORE_SBM", "CORE_Util", "TEST_New"}
    assert selection["CORE_SBM"]["File"] == "CORE_SBM (1).bas"
    assert selection["CORE_Util"]["File"] == "renamed.bas"
    assert selection["TEST_New"]["IsNew"] is True
    assert selection["CORE_SBM"]["IsNew"] is False
    ignored = " | ".join(found["Ignored"])
    assert "CORE_SBM.bas (older copy of 'CORE_SBM')" in ignored
    assert "Helper.bas (VB_Name 'Helper' is not a repository module)" in ignored
    assert "M_Main.bas (no Attribute VB_Name line" in ignored
    assert "notes.txt" not in ignored


@needs_pwsh
def test_copy_converts_line_endings_and_keeps_cp1252_bytes(tmp_path):
    source = tmp_path / "in.bas"
    target = tmp_path / "out.bas"
    source.write_bytes(b'Attribute VB_Name = "M_X"\n\' caf\xe9\rlast\r\n')
    run_pwsh(
        tmp_path,
        f"Copy-WithCrlf -SourcePath '{source}' -TargetPath '{target}'\n"
        "'{}' | Set-Content -LiteralPath $env:RESULT",
    )
    assert target.read_bytes() == b'Attribute VB_Name = "M_X"\r\n\' caf\xe9\r\nlast\r\n'


@needs_pwsh
def test_mode_answers(tmp_path):
    answers = ["p", "PROD", " production ", "d", "Dev", "development", "", "x", "pd"]
    modes = run_pwsh(
        tmp_path,
        "$a = " + json.dumps(answers).replace("[", "@(").replace("]", ")") + "\n"
        "ConvertTo-Json -InputObject @($a | ForEach-Object { [string](ConvertTo-SyncMode -Answer $_) })"
        " | Set-Content -LiteralPath $env:RESULT",
    )
    assert modes == ["Production"] * 3 + ["Development"] * 3 + [""] * 3


@needs_pwsh
def test_repository_modules_match_the_build_sets(tmp_path):
    sets = run_pwsh(
        tmp_path,
        f"$p = @(Get-RepositoryModules -RepoPath '{REPO_ROOT}' -SyncMode Production | ForEach-Object Name)\n"
        f"$d = @(Get-RepositoryModules -RepoPath '{REPO_ROOT}' -SyncMode Development | ForEach-Object Name)\n"
        "@{ Production = $p; Development = $d } | ConvertTo-Json | Set-Content -LiteralPath $env:RESULT",
    )
    production = sorted(path.stem for path in (REPO_ROOT / "src" / "modules").glob("*.bas"))
    tests = sorted(path.stem for path in (REPO_ROOT / "tests" / "modules").glob("*.bas"))
    assert sets["Production"] == production
    assert sets["Development"] == production + tests


@needs_pwsh
def test_repository_modules_reject_a_duplicate_name(tmp_path):
    for folder in ("src/modules", "tests/modules"):
        (tmp_path / folder).mkdir(parents=True)
    (tmp_path / "src/modules/CORE_A.bas").write_bytes(module("CORE_A"))
    (tmp_path / "tests/modules/TEST_A.bas").write_bytes(module("CORE_A"))
    outcome = run_pwsh(
        tmp_path,
        f"try {{ Get-RepositoryModules -RepoPath '{tmp_path}' -SyncMode Development | Out-Null; $m = '' }}"
        " catch { $m = $_.Exception.Message }\n"
        "ConvertTo-Json -InputObject $m | Set-Content -LiteralPath $env:RESULT",
    )
    assert "Duplicate module name 'CORE_A'" in outcome


def test_launchers_call_existing_scripts():
    for launcher, script in [
        ("Sync-FRTB-VBA.cmd", "Sync-FRTB-VBA.ps1"),
        ("Sync-FRTB-VBA-Downloads.cmd", "Sync-FRTB-VBA.ps1"),
        ("Build-Workbook.cmd", "Build-Workbook.ps1"),
        ("Verify-FRTB-Branch.cmd", "Check-FRTB-Excel.ps1"),
    ]:
        text = (REPO_ROOT / "tools" / launcher).read_text(encoding="ascii")
        assert f"%~dp0{script}" in text
        assert (REPO_ROOT / "tools" / script).is_file()


def test_launchers_never_echo_a_path_directly():
    # A path may contain & | < > or parentheses; expanded into an echo line
    # with %VAR%, cmd.exe would run the rest as a command. Paths go through
    # the :echovar helper, which inserts the value with delayed expansion.
    path_variables = ("SCRIPT", "WORKBOOK", "SOURCE", "REPO_ROOT", "OUTPUT", "DEFAULT_SOURCE", "BUILD_SCRIPT", "FRESH",
                      "BRANCH", "MODE")
    launchers = ("Sync-FRTB-VBA.cmd", "Sync-FRTB-VBA-Downloads.cmd", "Build-Workbook.cmd", "Verify-FRTB-Branch.cmd")
    for launcher in launchers:
        lines = (REPO_ROOT / "tools" / launcher).read_text(encoding="ascii").splitlines()
        for line in lines:
            if line.strip().lower().startswith("echo"):
                assert not any(f"%{name}%" in line for name in path_variables), (launcher, line)
        assert "setlocal EnableDelayedExpansion" in lines and "echo(  !%~1!" in lines, launcher


def test_the_sync_launcher_builds_a_development_workbook_without_one_and_asks_the_mode_with_one():
    text = (REPO_ROOT / "tools" / "Sync-FRTB-VBA.cmd").read_text(encoding="ascii")
    # Without a workbook: a development workbook from the template, in the ignored test-results folder
    assert "if not defined WORKBOOK goto :fresh" in text
    assert 'set "FRESH=%CD%\\test-results\\K-FRTB-ASA-Development.xlsm"' in text
    assert '-File "%BUILD_SCRIPT%" -OutputPath "%FRESH%" -Force' in text
    assert "-Production" not in text.split(":fresh", 1)[1]
    # With one: the mode comes from the second argument or from a question, and goes to the script
    assert 'set "MODE=%~2"' in text and 'set /p "MODE=' in text
    assert 'set "MODE=%MODE:"=%"' in text
    for mode in ("Production", "Development"):
        assert f'set "MODE_ARG=-Mode {mode}"' in text
    assert '-RepoRoot "%REPO_ROOT%" %MODE_ARG%' in text
    # A mode that is neither stops before anything is changed
    assert "CANCELLED: the mode must be P or D" in text
