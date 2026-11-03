"""Offline tests of tools/compare_vba_export.py; nothing here runs Excel.

Each test writes an "export" from the repository source at HEAD, with the
CRLF line endings the VBE uses, changes it the way a faulty round trip would,
and runs the comparison.
"""

import hashlib
import io
import json
import os
import shutil
import subprocess
import zipfile
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import pytest

import compare_vba_export as tool

ROOT = Path(__file__).resolve().parents[1]
# An empty ThisWorkbook as the VBE exports it.
THIS_WORKBOOK = (
    b"VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
    b'Attribute VB_Name = "ThisWorkbook"\r\n'
    b'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
    b"Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
    b"Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n"
    b"Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n"
)


# A sheet module as the VBE's Export writes it: no VB_Base line, so only
# the manifest's component type marks it as a document module.
SHEET1 = (
    b"VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
    b'Attribute VB_Name = "Sheet1"\r\n'
    b"Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
    b"Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n"
    b"Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n"
)


def committed_modules(mode):
    folders = tool.MODE_FOLDERS[mode]
    listing = subprocess.run(["git", "-C", str(ROOT), "ls-tree", "--name-only", "HEAD", *[f"{f}/" for f in folders]],
                             capture_output=True, check=True, text=True).stdout.split()
    return [path for path in listing if path.endswith(".bas")]


def source(path):
    return subprocess.run(["git", "-C", str(ROOT), "show", f"HEAD:{path}"], capture_output=True, check=True).stdout


@pytest.fixture
def export(tmp_path):
    folder = tmp_path / "export"
    folder.mkdir()

    def write(mode, document=THIS_WORKBOOK):
        for path in committed_modules(mode):
            (folder / Path(path).name).write_bytes(source(path).replace(b"\r\n", b"\n").replace(b"\n", b"\r\n"))
        if document is not None:
            (folder / "ThisWorkbook.cls").write_bytes(document)
        return folder

    return write


def run_tool(folder, mode, *extra):
    out, err = io.StringIO(), io.StringIO()
    with redirect_stdout(out), redirect_stderr(err):
        code = tool.main(["--root", str(ROOT), "--export", str(folder), "--mode", mode, *extra])
    return code, out.getvalue() + err.getvalue()


def replace_in(path, old, new):
    data = path.read_bytes()
    assert old in data, (path, old)
    path.write_bytes(data.replace(old, new, 1))


def workbook(folder, name, project=b"vba project"):
    """A stand-in for a built .xlsm: a zip holding a VBA project part."""
    path = folder / name
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/vbaProject.bin", project)
    return path


def write_manifest(folder, book, types=None):
    components = [{"name": tool.parse(path.read_bytes(), path.name).name, "file": path.name,
                   "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
                  for path in sorted(folder.iterdir()) if path.suffix in tool.SUFFIXES]
    for item in components:
        if types and item["file"] in types:
            item["type"] = types[item["file"]]
    record = {"format": tool.EXPORT_FORMAT,
              "workbook": {"file": book.name, "sha256": hashlib.sha256(book.read_bytes()).hexdigest()},
              "components": components}
    (folder / tool.EXPORT_MANIFEST).write_text(json.dumps(record, indent=2), encoding="utf-8")


@pytest.mark.parametrize("mode", ["production", "development"])
def test_an_exact_export_passes(export, mode):
    code, output = run_tool(export(mode), mode)
    assert code == 0, output
    assert "RESULT=PASS" in output
    assert "binding=UNVERIFIED" in output
    assert f"modules={len(committed_modules(mode))}" in output
    assert tool.TEMPLATE in output


def test_recased_identifiers_are_counted_not_failed(export):
    folder = export("production")
    replace_in(folder / "CORE_SBM.bas", b"Option Explicit", b"option explicit")
    code, output = run_tool(folder, "production")
    assert code == 0, output
    assert "Re-cased:  1 line(s)" in output


@pytest.mark.parametrize(
    ("old", "new"),
    [
        (b"Option Explicit", b"Option Explicit\r\nDim extra As Long"),
        (b"' MODULE: CORE_SBM", b"' module: CORE_SBM"),
        (b'Attribute VB_Name = "CORE_SBM"', b'Attribute VB_Name = "CORE_SBM"\r\nAttribute VB_Description = "x"'),
    ],
    ids=["changed code", "changed comment case", "changed attribute"],
)
def test_a_changed_module_fails(export, old, new):
    folder = export("production")
    replace_in(folder / "CORE_SBM.bas", old, new)
    code, output = run_tool(folder, "production")
    assert code == 1, output
    assert "FAIL src/modules/CORE_SBM.bas:" in output


def test_a_case_change_inside_a_string_fails(export):
    folder = export("production")
    data = (folder / "CORE_Config.bas").read_bytes()
    literal = data.split(b'"')[3]  # the first string literal after the VB_Name
    assert literal and literal.lower() != literal
    replace_in(folder / "CORE_Config.bas", b'"' + literal + b'"', b'"' + literal.lower() + b'"')
    code, output = run_tool(folder, "production")
    assert code == 1, output
    assert "CORE_Config.bas: code differs" in output


def test_missing_and_unexpected_modules_fail(export):
    folder = export("production")
    (folder / "CORE_DRC.bas").unlink()
    (folder / "M_Extra.bas").write_bytes(b'Attribute VB_Name = "M_Extra"\r\nOption Explicit\r\n')
    code, output = run_tool(folder, "production")
    assert code == 1, output
    assert "missing from the export: CORE_DRC (src/modules/CORE_DRC.bas)" in output
    assert "not in the source for this mode: M_Extra (M_Extra.bas)" in output


def test_a_test_module_in_a_production_export_fails(export):
    folder = export("development")
    code, output = run_tool(folder, "production")
    assert code == 1, output
    assert "not in the source for this mode: TEST_Regression" in output


def test_a_document_module_with_code_fails(export):
    folder = export("production", THIS_WORKBOOK + b"Option Explicit\r\nPrivate Sub Workbook_Open()\r\nEnd Sub\r\n")
    code, output = run_tool(folder, "production")
    assert code == 1, output
    assert "document module ThisWorkbook (ThisWorkbook.cls) holds code that is not in the source" in output


def test_a_document_module_with_only_option_lines_passes(export):
    folder = export("production", THIS_WORKBOOK + b"Option Explicit\r\n\r\n")
    code, output = run_tool(folder, "production")
    assert code == 0, output


def test_a_sheet_exported_without_vb_base_is_a_document_by_its_manifest_type(export, tmp_path):
    folder = export("development")
    (folder / "Sheet1.cls").write_bytes(SHEET1 + b"Option Explicit\r\n")
    book = workbook(tmp_path, "K-FRTB-SA.xlsm")
    write_manifest(folder, book, {"ThisWorkbook.cls": 100, "Sheet1.cls": 100, "CORE_SBM.bas": 1})
    code, output = run_tool(folder, "development", "--workbook", str(book))
    assert code == 0, output
    assert "RESULT=PASS" in output and "binding=VERIFIED" in output


def test_a_sheet_without_vb_base_or_manifest_type_is_not_in_the_source(export, tmp_path):
    folder = export("development")
    (folder / "Sheet1.cls").write_bytes(SHEET1)
    book = workbook(tmp_path, "K-FRTB-SA.xlsm")
    write_manifest(folder, book)
    code, output = run_tool(folder, "development", "--workbook", str(book))
    assert code == 1, output
    assert "not in the source for this mode: Sheet1 (Sheet1.cls)" in output


def test_a_manifest_document_type_cannot_hide_code(export, tmp_path):
    folder = export("development")
    (folder / "Sheet1.cls").write_bytes(SHEET1 + b"Sub Hidden()\r\nEnd Sub\r\n")
    book = workbook(tmp_path, "K-FRTB-SA.xlsm")
    write_manifest(folder, book, {"Sheet1.cls": 100})
    code, output = run_tool(folder, "development", "--workbook", str(book))
    assert code == 1, output
    assert "document module Sheet1 (Sheet1.cls) holds code that is not in the source" in output


def test_a_component_type_that_is_not_an_integer_fails_the_binding(export, tmp_path):
    folder = export("development")
    book = workbook(tmp_path, "K-FRTB-SA.xlsm")
    write_manifest(folder, book, {"ThisWorkbook.cls": "100"})
    code, output = run_tool(folder, "development", "--workbook", str(book))
    assert code == 1, output
    assert "a component type is not an integer" in output


def test_a_matching_manifest_verifies_the_binding(export, tmp_path):
    folder = export("development")
    book = workbook(tmp_path, "K-FRTB-SA.xlsm")
    write_manifest(folder, book)
    code, output = run_tool(folder, "development", "--workbook", str(book))
    assert code == 0, output
    assert "binding=VERIFIED" in output


def test_a_manifest_without_workbook_stays_unverified(export, tmp_path):
    folder = export("development")
    write_manifest(folder, workbook(tmp_path, "K-FRTB-SA.xlsm"))
    code, output = run_tool(folder, "development")
    assert code == 0, output
    assert "binding=UNVERIFIED" in output
    assert "no --workbook given" in output


def test_a_file_changed_after_the_export_fails_the_binding(export, tmp_path):
    folder = export("development")
    book = workbook(tmp_path, "K-FRTB-SA.xlsm")
    write_manifest(folder, book)
    replace_in(folder / "CORE_SBM.bas", b"Option Explicit", b"option explicit")
    code, output = run_tool(folder, "development", "--workbook", str(book))
    assert code == 1, output
    assert "CORE_SBM.bas changed after the export" in output


def test_another_workbook_fails_the_binding(export, tmp_path):
    folder = export("development")
    write_manifest(folder, workbook(tmp_path, "K-FRTB-SA.xlsm"))
    other = workbook(tmp_path, "other.xlsm", b"another project")
    code, output = run_tool(folder, "development", "--workbook", str(other))
    assert code == 1, output
    assert "other.xlsm is not the workbook the export was made from" in output


def test_a_renamed_workbook_is_a_note(export, tmp_path):
    folder = export("development")
    book = workbook(tmp_path, "K-FRTB-SA.xlsm")
    write_manifest(folder, book)
    renamed = book.rename(tmp_path / "K-FRTB-SA (1).xlsm")
    code, output = run_tool(folder, "development", "--workbook", str(renamed))
    assert code == 0, output
    assert "NOTE export-manifest.json names the workbook K-FRTB-SA.xlsm" in output


def test_a_workbook_without_vba_fails_the_binding(export, tmp_path):
    folder = export("development")
    book = tmp_path / "K-FRTB-SA.xlsm"
    with zipfile.ZipFile(book, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
    write_manifest(folder, book)
    code, output = run_tool(folder, "development", "--workbook", str(book))
    assert code == 1, output
    assert "is not a macro-enabled workbook" in output


def test_unreadable_inputs_exit_2(export, tmp_path):
    folder = export("production")
    assert run_tool(folder, "production", "--candidate-sha", "not-a-commit")[0] == 2
    assert run_tool(tmp_path / "missing", "production")[0] == 2
    (folder / "CORE_SBM (1).bas").write_bytes(source("src/modules/CORE_SBM.bas"))
    code, output = run_tool(folder, "production")
    assert code == 2
    assert "exported twice" in output


def test_an_unreadable_component_exits_2(export):
    folder = export("production")
    (folder / "broken.bas").mkdir()
    code, output = run_tool(folder, "production")
    assert code == 2, output
    assert "broken.bas: not a file" in output
    (folder / "broken.bas").rmdir()
    (folder / "CORE_SBM.bas").chmod(0)
    try:
        if not os.access(folder / "CORE_SBM.bas", os.R_OK):  # root can still read it
            code, output = run_tool(folder, "production")
            assert code == 2, output
            assert "Traceback" not in output
    finally:
        (folder / "CORE_SBM.bas").chmod(0o644)


def test_a_component_that_is_not_cp1252_exits_2(export):
    folder = export("production")
    (folder / "CORE_SBM.bas").write_bytes(b'Attribute VB_Name = "CORE_SBM"\r\n\x81\r\n')
    code, output = run_tool(folder, "production")
    assert code == 2, output


def test_exporter_writes_the_format_and_manifest_the_tool_reads():
    script = (ROOT / "tools" / "Export-FRTB-VBA.ps1").read_text(encoding="ascii")
    assert f"$exportFormat = '{tool.EXPORT_FORMAT}'" in script
    assert f"$exportManifestName = '{tool.EXPORT_MANIFEST}'" in script
    # Read-only open with macros disabled, and a hash check after the close.
    assert "$excel.Workbooks.Open($workbookFullPath, 0, $true)" in script
    assert "$excel.AutomationSecurity = $msoAutomationSecurityForceDisable" in script
    assert "The workbook changed while it was exported" in script
    # The manifest records each component's type, which marks a document module.
    assert "type   = [int]$component.Type" in script


@pytest.mark.skipif(not (os.environ.get("PWSH") or shutil.which("pwsh")), reason="PowerShell 7 (pwsh) is not installed")
def test_exporter_parses_without_errors():
    pwsh = os.environ.get("PWSH") or shutil.which("pwsh")
    command = ("$e = $null; $t = $null; [void][System.Management.Automation.Language.Parser]::ParseFile("
               f"'{ROOT / 'tools' / 'Export-FRTB-VBA.ps1'}', [ref]$t, [ref]$e); "
               "if ($e) { $e | ForEach-Object { $_.ToString() }; exit 1 }")
    completed = subprocess.run([pwsh, "-NoProfile", "-NonInteractive", "-Command", command],
                               capture_output=True, text=True, timeout=120, check=False)
    assert completed.returncode == 0, completed.stdout + completed.stderr
