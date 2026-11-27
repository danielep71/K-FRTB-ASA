"""The committed workbook template and the build that starts from it."""

from __future__ import annotations

import re
import zipfile

import openpyxl
from openpyxl.styles import Protection

import build_template
import lo_runner
from conftest import require_libreoffice
from test_vba_engine import FIRST_COL, log_rows, reconcile


def test_template_is_current_and_macro_free():
    require_libreoffice()
    assert build_template.check() == []


def test_source_build_print_areas_match_template(tmp_path):
    """VBA source builds set print extents without the OOXML normalizer."""
    require_libreoffice()
    out = tmp_path / "source-print-areas.xlsx"
    assert lo_runner.run(out, ["M_Setup.BuildWorkbook"]) == []
    source = openpyxl.load_workbook(out)
    template = openpyxl.load_workbook(build_template.TEMPLATE)
    for ws in source:
        assert str(ws.print_area) == str(template[ws.title].print_area), ws.title


def test_template_saved_presentation():
    """The XLSX must look complete before any macro can run."""
    wb = openpyxl.load_workbook(build_template.TEMPLATE)
    for ws in wb:
        assert ws.sheet_view.showGridLines is False, ws.title
        expected_tab = "FF1F3864" if ws.title in ("README", "Home") else "FF548235" if ws.title == "Test_Expected" else "FF2E75B6" if ws.title in ("Results", "Log") or ws.title.startswith("Out_") else "FFBF9000"
        assert ws.sheet_properties.tabColor is not None, ws.title
        assert ws.sheet_properties.tabColor.rgb == expected_tab, ws.title
        wide = ws.title.startswith(("In_", "Out_")) or ws.title == "Log"
        # Columns A and B stay in view on every sheet; the wide tables keep their key columns too
        expected_freeze = ("C4" if ws.title in ("README", "Home") else "D6" if wide else "C6" if ws.title == "Results"
                           else "C7")
        assert ws.freeze_panes == expected_freeze, ws.title
        assert ws.page_setup.orientation == "landscape", ws.title
        assert ws.page_setup.fitToWidth == 1 and ws.page_setup.fitToHeight == 0
        assert ws.sheet_properties.pageSetUpPr is not None
        assert ws.sheet_properties.pageSetUpPr.fitToPage
        assert "Aptos" in str(ws.oddFooter)
    for name in ("Control", "In_Delta", "In_Vega", "In_Curvature", "In_DRC", "In_RRAO"):
        ws = wb[name]
        row = 7 if name == "Control" else 6
        assert ws.row_dimensions[row].height == 20.1, name
        assert ws.cell(row, 4).border.bottom.style == "thin", name
    assert wb["README"]["D7"].border.bottom.style == "dashed"
    assert wb["README"]["C6"].border.top.style == "thin"
    assert wb["README"]["C6"].border.right.style == "dashed"
    assert wb["README"]["D12"].border.bottom.style == "thin"
    assert wb["In_Delta"]["C2"].value == "Delta sensitivities"
    assert wb["In_Delta"]["D5"].value == "Risk class"
    for name in ("Results", "Out_Summary", "Out_RiskClass", "Out_Buckets", "Out_DRC", "Log"):
        assert wb[name]["C5"].fill.fgColor.rgb == "FF1F3864", name  # the input navy (#206)
        assert wb[name]["C6"].value == "No results yet: press Run FRTB-SA.", name


def test_template_type_sizes():
    """Table headers are 11 pt bold, data rows 10 pt, subtitles 30 pt high, zoom unset (#351)."""
    wb = openpyxl.load_workbook(build_template.TEMPLATE)
    headers = 0
    for ws in wb:
        assert ws.sheet_view.zoomScale in (None, 100), ws.title
        assert ws.row_dimensions[3].height == 30, ws.title
        for row in ws.iter_rows():
            for cell in row:
                if cell.value in (None, ""):
                    continue
                is_header = (cell.fill.fill_type and cell.fill.fgColor.rgb == "FF1F3864" and cell.font.b
                             and cell.font.color is not None and cell.font.color.rgb == "FFFFFFFF")
                if is_header:
                    headers += 1
                    assert cell.font.sz == 11, (ws.title, cell.coordinate, cell.font.sz)
                elif cell.row > 5 and cell.font.sz is not None and cell.font.sz > 10:
                    # the only larger text below the banner is a section heading: 14 pt navy
                    assert cell.font.sz == 14 and cell.font.b and cell.font.color.rgb == "FF1F3864", \
                        (ws.title, cell.coordinate, cell.font.sz)
    assert headers > 100
    # the cells below a header stay at 10 pt
    assert wb["In_Delta"]["C6"].font.sz == 10
    assert wb["P_GIRR"]["D7"].font.sz == 10


def test_home_lists_every_sheet_once_and_links_to_it():
    """The start sheet is the first sheet and lists every sheet once, as a link to cell A1 (#347)."""
    wb = openpyxl.load_workbook(build_template.TEMPLATE)
    values = openpyxl.load_workbook(build_template.TEMPLATE, data_only=True)
    assert wb.sheetnames[0] == "Home" and wb.sheetnames[-1] == "README"
    ws = wb["Home"]
    listed = []
    for row in range(1, ws.max_row + 1):
        formula = ws.cell(row, 3).value
        if isinstance(formula, str) and formula.startswith("=HYPERLINK("):
            name = str(values["Home"].cell(row, 3).value)
            assert formula == f'=HYPERLINK("#\'{name}\'!A1","{name}")', (row, formula)
            listed.append(name)
    assert sorted(listed) == sorted(wb.sheetnames) and len(listed) == len(set(listed))


def test_every_sheet_leaves_room_for_the_navigation_bar():
    """Column A is 18 wide, as the shared toolkit's, column B is a gutter 2 wide, and the brand and title bars
    start in column A (#347). The product name and the title sit in column C with the subtitle and the tables."""
    wb = openpyxl.load_workbook(build_template.TEMPLATE)
    for ws in wb:
        assert round(ws.column_dimensions["A"].width) == 19, ws.title      # 18 plus Excel's padding
        assert round(ws.column_dimensions["B"].width) == 3, ws.title       # 2 plus Excel's padding
        assert ws["C1"].value == "K-FRTB-ASA" and ws["B1"].value is None, ws.title
        assert ws["C2"].value and ws["B2"].value is None, ws.title
        assert ws["A1"].fill.fgColor.rgb == "FF000000", ws.title
        assert ws["A2"].fill.fgColor.rgb == "FF265A99", ws.title


def header_pixels(text):
    """Width of a bold 11 pt header: 9.4 px for a capital, 7 for an underscore, 7.5 for the rest.
    Calibrated on the Excel pictures of 27eb55a and 9390788, where capitals such as RHO_NAME
    and SectorGroup broke mid-word although a flat 7.5 px a character said they fit."""
    return sum(9.4 if ch.isupper() else 7.0 if ch == "_" else 7.5 for ch in text)


def header_lines(text, width_px):
    """Greedy word wrap of a header, with 8 px of padding."""
    lines, current = [], ""
    for word in text.split():
        trial = f"{current} {word}" if current else word
        if not current or header_pixels(trial) <= width_px - 8:
            current = trial
        else:
            lines.append(current)
            current = word
    return [*lines, current]


def test_headers_fit_their_cells():
    """An 11 pt header wraps in at most two lines, none wider than its cell, and a second
    line leaves room for the filter arrow. The model is an estimate calibrated on the
    Excel pictures of 27eb55a and 9390788, where it flags exactly the headers that came out cut (#351)."""
    wb = openpyxl.load_workbook(build_template.TEMPLATE)
    problems = []
    for ws in wb:
        widths = {}
        for dimension in ws.column_dimensions.values():
            if dimension.width and dimension.min is not None and dimension.max is not None:
                for column in range(dimension.min, dimension.max + 1):
                    widths[column] = dimension.width
        filtered = bool(ws.auto_filter.ref)
        for row in ws.iter_rows():
            for cell in row:
                header = (cell.fill.fill_type and cell.fill.fgColor.rgb == "FF1F3864" and cell.font.b
                          and cell.font.sz == 11 and cell.value not in (None, ""))
                if not header:
                    continue
                pixels = widths.get(cell.column or 0, 8.43) * 7 + 5
                lines = header_lines(str(cell.value), pixels)
                if len(lines) > 2 or any(header_pixels(line) > pixels - 8 for line in lines):
                    problems.append((ws.title, cell.coordinate, cell.value, lines))
                elif filtered and len(lines) == 2 and header_pixels(lines[1]) > pixels - 36:
                    problems.append((ws.title, cell.coordinate, cell.value, "second line under the filter arrow"))
    assert problems == []


def wrapped_lines(text: str, per_line: int) -> int:
    """M_Setup.WrappedLines: words, broken after a hyphen too."""
    lines, used, after_hyphen = 1, 0, False
    for word in text.replace("-", "- ").split(" "):
        if not word:
            continue
        gap = 1 if used and not after_hyphen else 0
        if used and used + gap + len(word) > per_line:
            lines, used, gap = lines + 1, 0, 0
        used += gap + len(word)
        while used > per_line:
            lines, used = lines + 1, used - per_line
        after_hyphen = word.endswith("-")
    return lines


def test_template_formats_and_wrapped_rows():
    # Excel ends #,##0.#### with the separator on a whole number (#206)
    wb = openpyxl.load_workbook(build_template.TEMPLATE)
    expected = {
        "In_Delta": {"H": "0.0##", "I": "#,##0.00"},
        "In_Vega": {"H": "0.0##", "I": "0.0##", "J": "#,##0.00"},
        "In_Curvature": {"H": "#,##0.00", "I": "#,##0.00"},
        "In_DRC": {"H": "#,##0.00", "I": "#,##0.00", "J": "0.0##", "K": "0.0##"},
        "In_RRAO": {"E": "#,##0.00"},
    }
    for name, columns in expected.items():
        ws = wb[name]
        for row in (6, 200):
            for column, number_format in columns.items():
                assert ws[f"{column}{row}"].number_format == number_format, (name, column, row)
            assert all("#.####" not in c.number_format and ",##0.####" not in c.number_format for c in ws[row]), name
    # A wrapped description gets one line of height per line it needs
    for name, tag_row in (("P_CSR", 7), ("P_EQ", 7)):
        ws = wb[name]
        per_line = int(round(ws.column_dimensions["D"].width) * 1.2)
        for row in range(tag_row, tag_row + 18):
            text = ws.cell(row, 4).value
            if isinstance(text, str) and ws.cell(row, 4).alignment.wrap_text:
                lines = wrapped_lines(text, per_line)
                height = ws.row_dimensions[row].height or 20.1
                assert height >= 13.5 * lines, (name, row, lines, height)
    for name in ("P_GIRR", "P_FX"):
        ws = wb[name]
        divisor = [c for row in ws.iter_rows(min_col=FIRST_COL, max_col=FIRST_COL) for c in row
                   if c.value == "SPECIFIED_DIVISOR"][0]
        assert ws.cell(divisor.row, FIRST_COL + 1).number_format == "0.0000", name


def test_presentation_normalization_preserves_content_and_is_idempotent(tmp_path):
    import shutil

    import template_presentation

    path = tmp_path / "template.xlsx"
    shutil.copyfile(build_template.TEMPLATE, path)
    before = build_template.signature(path)
    template_presentation.normalize(path)
    after = build_template.signature(path)
    assert build_template.signature_differences(before, after) == []
    template_presentation.normalize(path)
    assert build_template.signature_differences(after, build_template.signature(path)) == []


def _repack(source, target, part, old, new):
    """Copy a package, replacing one occurrence of old by new in one part."""
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part and old:
                text = data.decode("utf-8")
                assert old in text, (part, old)
                data = text.replace(old, new, 1).encode("utf-8")
            zout.writestr(item, data)


def test_signature_detects_layout_changes(tmp_path):
    """A width-, height-, tab-, fill- or format-only change must make the template stale."""
    template = build_template.TEMPLATE
    base = build_template.signature(template)
    control = build_template.sheet_parts(template)["Control"]
    control_xml = zipfile.ZipFile(template).read(control).decode("utf-8")
    tab_match = re.search(r"<tabColor[^>]*/>", control_xml)
    assert tab_match is not None
    tab_color = tab_match.group(0)

    unchanged = tmp_path / "unchanged.xlsx"
    _repack(template, unchanged, control, "", "")
    assert build_template.signature_differences(base, build_template.signature(unchanged)) == []

    cases = {
        "width": (control, 'width="34.75"', 'width="60"'),
        "height": (control, 'ht="26.1"', 'ht="40"'),
        "tab": (control, tab_color, '<tabColor rgb="FF00FF00"/>'),
        "gridlines": (control, 'showGridLines="0"', 'showGridLines="1"'),
        "freeze": (control, 'topLeftCell="C7"', 'topLeftCell="C8"'),
        "print": (control, 'orientation="landscape"', 'orientation="portrait"'),
        "fill": ("xl/styles.xml", "FF265A99", "FF265A98"),
        "number_format": ("xl/styles.xml", 'formatCode="0.0##%"', 'formatCode="0.00%"'),
    }
    for name, (part, old, new) in cases.items():
        path = tmp_path / f"{name}.xlsx"
        _repack(template, path, part, old, new)
        assert build_template.signature_differences(base, build_template.signature(path)), name


def test_build_from_template_runs_and_reconciles(tmp_path):
    require_libreoffice()
    out = tmp_path / "from-template.xlsx"
    errors = lo_runner.run(
        out,
        ["M_Setup.FinishWorkbook", "M_Main.RunFRTB_Silent", "TEST_Regression.Test_RunRegression_Silent"],
        input_book=build_template.TEMPLATE,
    )
    assert errors == []
    wb = openpyxl.load_workbook(out, data_only=True)
    assert str(wb["Test_Expected"].cell(3, FIRST_COL).value).startswith("Result: PASS")
    assert log_rows(wb, "ERROR") == []
    assert reconcile(wb) == []


def test_template_protects_sheets_and_structure_without_a_password():
    """Every sheet and the workbook structure are protected without a password; the cells in
    the input fill are the unlocked ones, and the input sheets allow inserted rows, sorting
    and the autofilter (#349)."""
    wb = openpyxl.load_workbook(build_template.TEMPLATE)
    security = wb.security
    assert security is not None and security.lockStructure
    assert not security.workbookPassword and not security.workbookHashValue
    problems = []
    for ws in wb:
        protection = ws.protection
        assert protection.sheet, ws.title
        assert not protection.password and not protection.hashValue, ws.title
        is_input = ws.title.startswith("In_")
        actions = (protection.insertRows, protection.sort, protection.autoFilter)
        assert actions == ((False,) * 3 if is_input else (True,) * 3), (ws.title, actions)
        assert (protection.formatCells, protection.deleteRows, protection.selectLockedCells) == (True, True, False), ws.title
        for row in ws.iter_rows():
            for cell in row:
                filled = bool(cell.fill.fill_type and cell.fill.fgColor.rgb == "FFFFF2CC")
                if cell.protection.locked == filled:
                    problems.append((ws.title, cell.coordinate, "locked" if filled else "unlocked"))
    assert problems == []
    # The cells to edit are where the guide says they are
    unlocked = {ws.title: sum(not c.protection.locked for r in ws.iter_rows() for c in r) for ws in wb}
    assert unlocked["Control"] == 5 and unlocked["In_Delta"] > 0
    assert all(unlocked[name] == 0 for name in ("Home", "Results", "Log", "Test_Expected", "README")
               + tuple(n for n in unlocked if n.startswith("Out_")))


def test_a_workbook_with_every_cell_locked_gets_its_input_cells_unlocked(tmp_path):
    """A workbook made before the protection, or updated by the sync, has every cell locked and
    no protection; the first command unlocks the cells in the input fill and protects it (#349)."""
    require_libreoffice()
    old = openpyxl.load_workbook(build_template.TEMPLATE)
    for ws in old:
        ws.protection.sheet = False
        for row in ws.iter_rows():
            for cell in row:
                cell.protection = Protection(locked=True)
    old.security = None
    source = tmp_path / "old.xlsx"
    old.save(source)
    out = tmp_path / "migrated.xlsx"
    assert lo_runner.run(out, ["M_Main.ClearOutputs"], input_book=source) == []
    migrated = openpyxl.load_workbook(out)
    wrong = []
    for ws in migrated:
        assert ws.protection.sheet, ws.title
        for row in ws.iter_rows():
            for cell in row:
                filled = bool(cell.fill.fill_type and cell.fill.fgColor.rgb == "FFFFF2CC")
                if cell.protection.locked == filled:
                    wrong.append((ws.title, cell.coordinate))
    assert wrong == []


def test_every_sheet_has_a_black_navigation_column():
    """Column A holds the navigation bar on a black column (#347)."""
    wb = openpyxl.load_workbook(build_template.TEMPLATE)
    for ws in wb:
        fill = ws.column_dimensions["A"].fill
        assert fill.fill_type == "solid" and fill.fgColor.rgb == "FF000000", ws.title
