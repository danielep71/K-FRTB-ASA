"""The FRTB-SA ribbon tab: its custom UI part, its contract and its callbacks (#353)."""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
PART = REPO_ROOT / "src" / "ribbon" / "customUI14.xml"
CONTRACT = REPO_ROOT / "docs" / "ribbon.md"
MAIN = REPO_ROOT / "src" / "modules" / "M_Main.bas"
SETUP = REPO_ROOT / "src" / "modules" / "M_Setup.bas"
SCRIPT = REPO_ROOT / "tools" / "Add-RibbonTab.ps1"
NS = "{http://schemas.microsoft.com/office/2009/07/customui}"
PWSH = os.environ.get("PWSH") or shutil.which("pwsh")
CALLBACK_ATTRIBUTES = ("onLoad", "onAction", "getEnabled", "getPressed")


def controls() -> list[ET.Element]:
    root = ET.parse(PART).getroot()
    return [c for c in root.iter() if c.tag in (f"{NS}button", f"{NS}toggleButton")]


def contract_rows() -> dict[str, list[str]]:
    rows = {}
    for line in CONTRACT.read_text(encoding="utf-8").splitlines():
        cells = [c.strip().strip("`") for c in line.strip().strip("|").split("|")]
        if len(cells) == 7 and cells[1].startswith(("btn", "tgl")):
            rows[cells[1]] = cells
    return rows


def procedure_bodies(path: Path) -> dict[str, list[str]]:
    text = path.read_text(encoding="cp1252")
    bodies = {}
    for match in re.finditer(r"^Public (?:Sub|Function) (\w+)\(.*?\)[^\r\n]*\r?\n(.*?)^End (?:Sub|Function)", text, re.S | re.M):
        lines = [ln.strip() for ln in match.group(2).splitlines()]
        bodies[match.group(1)] = [ln for ln in lines if ln and not ln.startswith("'")]
    return bodies


def test_the_part_is_valid_and_every_control_has_a_screen_tip():
    root = ET.parse(PART).getroot()
    assert root.tag == f"{NS}customUI"
    assert root.get("onLoad") == "Ribbon_OnLoad"
    found = controls()
    assert [c.get("id") for c in found] == ["btnRun", "btnClear", "btnTest", "tglNav"]
    for control in found:
        assert control.get("screentip") and control.get("supertip"), control.get("id")
    ids = [e.get("id") for e in root.iter() if e.get("id")]
    assert len(ids) == len(set(ids))


def test_the_contract_and_the_part_agree_on_every_control_and_callback():
    rows = contract_rows()
    found = controls()
    assert set(rows) == {c.get("id") for c in found}
    for control in found:
        _, identifier, kind, callback, _, enabled, tip = rows[str(control.get("id"))]
        assert (kind == "toggle button") == (control.tag == f"{NS}toggleButton"), identifier
        assert control.get("onAction") == callback, identifier
        assert control.get("screentip") == tip, identifier
        assert ("Ribbon_IsIdle" in enabled) == (control.get("getEnabled") == "Ribbon_IsIdle"), identifier


def test_every_callback_exists_in_m_main_and_calls_only_what_the_contract_names():
    bodies = procedure_bodies(MAIN)
    rows = contract_rows()
    used = {c.get(a) for c in [ET.parse(PART).getroot(), *ET.parse(PART).getroot().iter()] for a in CALLBACK_ATTRIBUTES} - {None}
    assert used <= set(bodies), used - set(bodies)
    for control in controls():
        calls = rows[str(control.get("id"))][4]
        body = bodies[str(control.get("onAction"))]
        assert len(body) == 1 and calls in body[0], (control.get("id"), body)
    # The state callbacks return a state and nothing else; the load callback keeps the object
    assert bodies["Ribbon_IsIdle"] == ["returnedVal = Not mBusy"]
    assert bodies["Ribbon_NavShown"] == ["returnedVal = NavigationVisible()"]
    assert bodies["Ribbon_OnLoad"] == ["Set mRibbon = ribbon"]
    assert "SetNavigationVisible" in procedure_bodies(SETUP) and "NavigationVisible" in procedure_bodies(SETUP)


def test_a_run_marks_the_workbook_busy_and_idle_again():
    text = MAIN.read_text(encoding="cp1252")
    run = text.split("Private Function ExecuteRun()", 1)[1].split("End Function", 1)[0]
    assert run.index("SetBusy True") < run.index("ClearOutputs") < run.index("SetBusy False")
    assert "mRibbon.Invalidate" in text


@pytest.mark.skipif(not PWSH, reason="PowerShell 7 (pwsh) is not installed")
def test_the_script_adds_the_part_and_one_relationship_and_is_idempotent(tmp_path):
    book = tmp_path / "book.xlsm"
    with zipfile.ZipFile(book, "w") as archive:
        archive.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                         '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>')
        archive.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr("xl/workbook.xml", "<workbook/>")
        archive.writestr("xl/vbaProject.bin", b"vba")
    for _ in range(2):
        done = subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-File", str(SCRIPT), "-Workbook", str(book)],
                              capture_output=True, text=True, check=False)
        assert done.returncode == 0, done.stdout + done.stderr
    with zipfile.ZipFile(book) as archive:
        names = archive.namelist()
        assert names.count("customUI/customUI14.xml") == 1
        assert archive.read("customUI/customUI14.xml") == PART.read_bytes()
        assert archive.read("xl/vbaProject.bin") == b"vba" and archive.read("xl/workbook.xml") == b"<workbook/>"
        rels = archive.read("_rels/.rels").decode("utf-8")
        assert rels.count("ui/extensibility") == 1 and "officeDocument" in rels
        assert 'Extension="xml"' in archive.read("[Content_Types].xml").decode("utf-8")


def test_the_build_script_adds_the_tab_only_on_request_after_the_save():
    build = (REPO_ROOT / "tools" / "Build-Workbook.ps1").read_text(encoding="cp1252")
    assert "[switch]$Ribbon" in build and "-Template saves a macro-free template" in build
    assert build.index("$workbook.SaveAs($OutputPath, 52)") < build.index("Join-Path $PSScriptRoot 'Add-RibbonTab.ps1'")
    assert "if ($Ribbon -and $exitCode -eq 0)" in build
