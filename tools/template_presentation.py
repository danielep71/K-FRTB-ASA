"""Preserve workbook presentation lost by headless VBA compatibility.

The VBA builder owns cell content and the visual design. This OOXML step
restores view/print properties and README guide borders that the headless
host does not reliably expose through Excel's compatibility APIs.
"""

from __future__ import annotations

import re
from copy import deepcopy
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS)
ET.register_namespace("r", REL)


def _tag(name: str) -> str:
    return f"{{{NS}}}{name}"


def _child(parent, name, **attributes):
    element = parent.find(_tag(name))
    if element is None:
        element = ET.SubElement(parent, _tag(name))
    element.attrib.update({key: str(value) for key, value in attributes.items()})
    return element


def _palette() -> dict[str, str]:
    """Use the VBA palette, converting Excel's BGR Long to OOXML ARGB."""
    source = (Path(__file__).resolve().parents[1] / "src/modules/CORE_Config.bas").read_text()
    colors = {}
    for name, value in re.findall(r"Public Const (TAB_\w+) As Long = (\d+)", source):
        number = int(value)
        colors[name] = f"FF{number & 255:02X}{(number >> 8) & 255:02X}{(number >> 16) & 255:02X}"
    return colors


def _guide_borders(root, styles):
    """Restore WriteGuideRows' dashed interiors and solid block outlines.

    LibreOffice applies the block's outer edge style to its interior too.
    Guide blocks are consecutive README rows containing both label and text.
    Clone/deduplicate styles so no other cell or worksheet is affected.
    """
    borders = styles.find(_tag("borders"))
    formats = styles.find(_tag("cellXfs"))
    cells = {c.attrib["r"]: c for c in root.findall(f"{_tag('sheetData')}/{_tag('row')}/{_tag('c')}")}
    def populated(address):
        cell = cells.get(address)
        return cell is not None and any(cell.find(_tag(t)) is not None for t in ("v", "is"))
    rows = {int(address[1:]) for address in cells if address.startswith("C") and address[1:].isdigit() and int(address[1:]) > 3 and populated(address) and populated("D" + address[1:])}
    def intern(collection, element):
        serialized = ET.tostring(element)
        for index, existing in enumerate(collection):
            if ET.tostring(existing) == serialized:
                return index
        collection.append(element)
        collection.set("count", str(len(collection)))
        return len(collection) - 1
    for row in sorted(rows):
        for column in ("C", "D"):
            cell = cells[f"{column}{row}"]
            fmt = deepcopy(formats[int(cell.get("s", "0"))])
            border = deepcopy(borders[int(fmt.get("borderId", "0"))])
            for side, interior in (("left", column == "D"), ("right", column == "C"), ("top", row - 1 in rows), ("bottom", row + 1 in rows)):
                _child(border, side, style="dashed" if interior else "thin")
            fmt.set("borderId", str(intern(borders, border)))
            fmt.set("applyBorder", "1")
            cell.set("s", str(intern(formats, fmt)))


def normalize(path: Path) -> None:
    """Restore presentation only; never change cell values or formulas."""
    palette = _palette()
    with zipfile.ZipFile(path) as archive:
        parts = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    book = ET.fromstring(parts["xl/workbook.xml"])
    styles = ET.fromstring(parts["xl/styles.xml"])
    relations = {r.attrib["Id"]: r.attrib["Target"] for r in ET.fromstring(parts["xl/_rels/workbook.xml.rels"])}
    names = _child(book, "definedNames")
    for defined in list(names):
        if defined.get("name") in ("_xlnm.Print_Area", "_xlnm.Print_Titles"):
            names.remove(defined)
    for index, sheet in enumerate(book.find(_tag("sheets"))):
        name = sheet.attrib["name"]
        target = relations[sheet.attrib[f"{{{REL}}}id"]]
        target = "xl/" + target.lstrip("/").removeprefix("xl/")
        root = ET.fromstring(parts[target])
        if name == "README":
            _guide_borders(root, styles)
        props = root.find(_tag("sheetPr"))
        if props is None:
            props = ET.Element(_tag("sheetPr"))
            root.insert(0, props)
        role = "TAB_REFERENCE" if name == "README" else "TAB_TEST" if name == "Test_Expected" else "TAB_OUTPUT" if name in ("Results", "Log") or name.startswith("Out_") else "TAB_INPUT"
        color = props.find(_tag("tabColor"))
        if color is None:
            color = ET.Element(_tag("tabColor"))
            props.insert(0, color)
        color.attrib.clear()
        color.set("rgb", palette[role])
        _child(props, "pageSetUpPr", fitToPage="1")

        frozen_rows = 3 if name == "README" else 6 if name in ("Control", "Test_Expected") or name.startswith("P_") else 5
        # The wide input, output and Log tables keep their key columns in
        # view; the narrow tables of the other sheets freeze rows only.
        frozen_cols = 3 if name.startswith(("In_", "Out_")) or name == "Log" else 0
        for view in root.findall(f"{_tag('sheetViews')}/{_tag('sheetView')}"):
            view.set("showGridLines", "0")
            view.set("topLeftCell", "A1")
            for old in list(view):
                if old.tag in (_tag("pane"), _tag("selection")):
                    view.remove(old)
            pane = ET.Element(_tag("pane"), ySplit=str(frozen_rows), topLeftCell=f"{'D' if frozen_cols else 'A'}{frozen_rows + 1}", activePane="bottomRight" if frozen_cols else "bottomLeft", state="frozen")
            if frozen_cols:
                pane.set("xSplit", str(frozen_cols))
            view.insert(0, pane)
            ET.SubElement(view, _tag("selection"), pane=pane.attrib["activePane"], activeCell=pane.attrib["topLeftCell"], sqref=pane.attrib["topLeftCell"])

        _child(root, "pageMargins", left="0.25", right="0.25", top="0.5", bottom="0.5", header="0.2", footer="0.2")
        setup = _child(root, "pageSetup", paperSize="9", orientation="landscape", fitToWidth="1", fitToHeight="0")
        setup.attrib.pop("scale", None)
        footer = _child(root, "headerFooter")
        footer.clear()
        ET.SubElement(footer, _tag("oddFooter")).text = '&C&"Aptos,Regular"&9Page &P of &N'
        # Bound printing to populated or intentionally styled cells, excluding
        # default column styles that may extend to the worksheet's last column.
        coordinates = [c.attrib["r"] for c in root.findall(f"{_tag('sheetData')}/{_tag('row')}/{_tag('c')}")]
        def column_number(address):
            value = 0
            for letter in re.match(r"[A-Z]+", address).group():
                value = value * 26 + ord(letter) - 64
            return value
        last_col = re.match(r"[A-Z]+", max(coordinates, key=column_number)).group()
        last_row = max(int(re.search(r"\d+", address).group()) for address in coordinates)
        quoted = "'" + name.replace("'", "''") + "'"
        ET.SubElement(names, _tag("definedName"), name="_xlnm.Print_Titles", localSheetId=str(index)).text = f"{quoted}!$1:${frozen_rows}"
        # Outputs grow at run time, so do not freeze their print extent here.
        if role != "TAB_OUTPUT":
            ET.SubElement(names, _tag("definedName"), name="_xlnm.Print_Area", localSheetId=str(index)).text = f"{quoted}!$B$1:${last_col}${last_row}"
        parts[target] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    # definedNames precedes calcPr/external-link metadata in workbook order.
    book.remove(names)
    sheets_index = list(book).index(book.find(_tag("sheets")))
    book.insert(sheets_index + 1, names)
    parts["xl/workbook.xml"] = ET.tostring(book, encoding="utf-8", xml_declaration=True)
    parts["xl/styles.xml"] = ET.tostring(styles, encoding="utf-8", xml_declaration=True)
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as temp:
        replacement = Path(temp.name)
    try:
        with zipfile.ZipFile(replacement, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, content in parts.items():
                archive.writestr(name, content)
        replacement.replace(path)
    finally:
        replacement.unlink(missing_ok=True)
