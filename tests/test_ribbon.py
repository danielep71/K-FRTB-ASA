"""The FRTB-ASA ribbon tab: its custom UI part, its contract and its callbacks (#353)."""

from __future__ import annotations

import os
import re
import shutil
import struct
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
ICON_DIR = REPO_ROOT / "src" / "ribbon" / "icons"
ICONS = {"btnRun": "run", "btnClear": "clear", "btnTest": "test", "tglNav": "nav"}


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
        types = archive.read("[Content_Types].xml").decode("utf-8")
        assert 'Extension="xml"' in types and types.count('Extension="png"') == 1
        # The four icons, one relationship each, in the rels of the part
        for name in ICONS.values():
            assert names.count(f"customUI/images/{name}.png") == 1
            assert archive.read(f"customUI/images/{name}.png") == (ICON_DIR / f"{name}.png").read_bytes()
        part_rels = archive.read("customUI/_rels/customUI14.xml.rels").decode("utf-8")
        assert part_rels.count("/relationships/image") == len(ICONS)
        for name in ICONS.values():
            assert f'Id="{name}"' in part_rels and f'Target="images/{name}.png"' in part_rels


def test_the_build_script_adds_the_tab_only_on_request_after_the_save():
    build = (REPO_ROOT / "tools" / "Build-Workbook.ps1").read_text(encoding="cp1252")
    assert "[switch]$Ribbon" in build and "-Template saves a macro-free template" in build
    assert build.index("$workbook.SaveAs($OutputPath, 52)") < build.index("Join-Path $PSScriptRoot 'Add-RibbonTab.ps1'")
    assert "if ($Ribbon -and $exitCode -eq 0)" in build


def test_the_tab_is_called_frtb_asa_in_the_part_and_in_the_documents():
    tabs = ET.parse(PART).getroot().findall(f".//{NS}tab")
    assert [tab.get("label") for tab in tabs] == ["FRTB-ASA"]
    for document in (CONTRACT, REPO_ROOT / "docs" / "workbook-guide.md"):
        text = document.read_text(encoding="utf-8")
        assert "**FRTB-ASA** tab" in text or "**FRTB-ASA** ribbon tab" in text, document.name
        assert "**FRTB-SA** tab" not in text and "**FRTB-SA** ribbon tab" not in text, document.name


def png_header(path: Path) -> tuple[int, int, int, int]:
    data = path.read_bytes()
    assert data[:8] == b"\x89PNG\r\n\x1a\n", path.name
    return struct.unpack(">IIBB", data[16:26])


def test_every_command_has_an_icon_of_its_own_and_the_files_and_the_contract_agree():
    found = {str(c.get("id")): c for c in controls()}
    assert {i: c.get("image") for i, c in found.items()} == ICONS
    assert all(c.get("imageMso") is None for c in found.values())
    assert len(set(ICONS.values())) == len(ICONS)
    contract = CONTRACT.read_text(encoding="utf-8")
    for control_id, name in ICONS.items():
        # The icon table of the contract names the control, its image and both files
        row = next(ln for ln in contract.splitlines() if ln.startswith(f"| `{control_id}` | `{name}` |"))
        assert f"`{name}.svg`" in row and f"`{name}.png`" in row, control_id
        # 64 by 64 pixels, 8 bits per channel, color type 6 (RGBA), so the corners are transparent
        assert png_header(ICON_DIR / f"{name}.png") == (64, 64, 8, 6), name
        assert (ICON_DIR / f"{name}.svg").read_text(encoding="utf-8").lstrip().startswith("<svg"), name
    assert sorted(p.name for p in ICON_DIR.iterdir()) == sorted(f"{n}.{e}" for n in ICONS.values() for e in ("png", "svg"))


def test_the_icons_share_one_style():
    # A navy tile (#1F3864) with white glyphs, as the contract says: every source draws the tile first
    for name in ICONS.values():
        svg = (ICON_DIR / f"{name}.svg").read_text(encoding="utf-8")
        assert re.search(r'<rect x="1.5" y="1.5" width="29" height="29" rx="6" fill="#1F3864"/>', svg), name
        assert "#FFFFFF" in svg and 'viewBox="0 0 32 32"' in svg, name


def test_the_toggle_hides_column_a_with_the_bar_and_the_bar_is_rebuilt_with_it_shown():
    setup = SETUP.read_text(encoding="cp1252")
    toggle = setup.split("Public Sub SetNavigationVisible", 1)[1].split("End Sub", 1)[0]
    assert "CurrentSheet.Shapes(ShapeIndex).Visible = shown" in toggle
    assert "CurrentSheet.Columns(1).Hidden = Not shown" in toggle
    # The buttons move with their cells, so the toggle moves none; it keeps the Saved flag
    assert "Buttons" not in toggle and "Offset" not in toggle
    assert "WasSaved = ThisWorkbook.Saved" in toggle and "ThisWorkbook.Saved = WasSaved" in toggle
    # A hidden column has no width to place the bar by
    bar = setup.split("Private Sub AddNavigationBar", 1)[1].split("End Sub", 1)[0]
    assert bar.index("ws.Columns(1).Hidden = False") < bar.index("BarWidth = ws.Columns(1).Width - 12")
    # The acceptance check toggles the bar and reads column A and the shapes of every sheet
    acceptance = (REPO_ROOT / "tests" / "modules" / "TEST_Acceptance.bas").read_text(encoding="cp1252")
    check = acceptance.split("Private Function CheckNavigation", 1)[1].split("End Function", 1)[0]
    assert "SetNavigationVisible (ToggleStep = 1)" in check and "Columns(1).Hidden" in check
    # The baseline is the shown bar, whatever state the workbook was saved in
    assert check.index("SetNavigationVisible True") < check.index("NavWidth(SheetIndex) = ")


def test_the_sheet_buttons_are_rounded_rectangles_that_move_with_their_cells():
    setup = SETUP.read_text(encoding="cp1252")
    row = setup.split("Private Sub AddButtonRow", 1)[1].split("End Sub", 1)[0]
    # A shape, not a form control, named btn_ plus its macro, with the macro assigned
    assert "ws.Shapes.AddShape(5," in row and "ws.Buttons.Add" not in row
    assert 'Button.Name = "btn_" & ' in row and "Button.OnAction = " in row
    # It follows its cells when column A is hidden, and keeps its size
    assert "Button.Placement = 2" in row and "Placement = 3" not in row
    # The style: navy run command, light blue others, a navy line, Aptos 11 pt bold
    assert 'IsPrimary = (CStr(ButtonDefs(LBound(ButtonDefs) + Index)(1)) = "RunFRTB")' in row
    assert "IIf(IsPrimary, COLOR_INPUT_HEADER, COLOR_TOTAL_FILL)" in row
    assert "IIf(IsPrimary, COLOR_WHITE, COLOR_INPUT_HEADER)" in row
    assert "Button.Line.Weight = 0.75" in row and "Font.Size = 11" in row and "Font.Bold = -1" in row
    # The shapes are removed by name before they are added again
    removal = setup.split("Private Sub RemoveButtons", 1)[1].split("End Sub", 1)[0]
    assert 'Left$(ws.Shapes(ShapeIndex).Name, 4) = "btn_"' in removal


def test_the_write_window_holds_screen_updating_off_and_gives_it_back_once():
    util = (REPO_ROOT / "src" / "modules" / "CORE_Util.bas").read_text(encoding="cp1252")
    begin = util.split("Public Sub BeginWrite", 1)[1].split("End Sub", 1)[0]
    end = util.split("Public Sub EndWrite", 1)[1].split("End Sub", 1)[0]
    # Held before the protection is lifted, so that the unprotecting does not flicker either
    assert begin.index("Application.ScreenUpdating = False") < begin.index("SetProtection False")
    assert "mScreenBeforeWrite = Application.ScreenUpdating" in begin
    # Given back after the protection returns, only when this window held it
    assert end.index("SetProtection True") < end.index("Application.ScreenUpdating = mScreenBeforeWrite")
    assert "If mScreenHeld Then" in end


def test_the_bar_has_start_at_the_top_and_two_triangles_that_toggle_it():
    setup = SETUP.read_text(encoding="cp1252")
    bar = setup.split("Private Sub AddNavigationBar", 1)[1].split("End Sub", 1)[0]
    # START and Home come first, in the rows that stay in view; the rest starts below every frozen row
    assert "BarTop = ws.Cells(1, 1).Top + 5" in bar
    assert "If Position = 2 Then BarTop = ws.Cells(LAYOUT_FIRST_DATA_ROW + 1, 1).Top" in bar
    assert "For Position = 1 To NAV_COUNT" in bar
    # Two triangles in B1, turned left and right, that follow their cells and run the macro
    assert "ws.Shapes.AddShape(7," in bar and "ws.Cells(1, 2)" in bar
    assert "NavShape.Rotation = IIf(ToggleStep = 0, 270, 90)" in bar and "NavShape.Placement = 2" in bar
    assert 'NavShape.OnAction = "ToggleNavigationBar"' in bar
    assert "NavShape.Name = IIf(ToggleStep = 0, NAV_TOGGLE_HIDE, NAV_TOGGLE_SHOW)" in bar
    assert "NavShape.Visible = IIf(ToggleStep = 0, -1, 0)" in bar
    # The toggle gives the triangle that shows the bar the opposite state
    toggle = setup.split("Public Sub SetNavigationVisible", 1)[1].split("End Sub", 1)[0]
    assert "Name = NAV_TOGGLE_SHOW Then" in toggle and ".Visible = Not shown" in toggle
    # The macro asks for the other state and refreshes the ribbon
    main = procedure_bodies(MAIN)
    assert main["ToggleNavigationBar"][:2] == ["If mBusy Then Exit Sub", "SetNavigationVisible Not NavigationVisible()"]
    assert "mRibbon.Invalidate" in " ".join(main["ToggleNavigationBar"])
    config = (REPO_ROOT / "src" / "modules" / "CORE_Config.bas").read_text(encoding="cp1252")
    assert 'NAV_TOGGLE_HIDE As String = "NavToggleHide"' in config and 'NAV_TOGGLE_SHOW As String = "NavToggleShow"' in config
    # Both names start with Nav, so RemoveNavigation clears them with the bar
    assert 'Left$(ws.Shapes(Index).Name, 3) = "Nav"' in setup


def test_the_title_bar_starts_in_the_gutter_and_column_a_stays_black():
    setup = SETUP.read_text(encoding="cp1252")
    frame = setup.split("Private Sub ApplyFrame", 1)[1].split("End Sub", 1)[0]
    assert "ws.Cells(2, 1).Interior.Color = COLOR_BRAND_BAR" in frame
    assert "With ws.Range(ws.Cells(2, LAYOUT_FIRST_COL - 1), ws.Cells(2, lastCol))" in frame
