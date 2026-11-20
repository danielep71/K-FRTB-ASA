"""Highlight a repeated key in the parameter tables the run reads as Key | Value (#350).

The run ignores the second of two equal keys in a table it reads with
ReadKeyValues and logs "Duplicate key ignored in [TAG]". A conditional format
makes the repeat visible where it is typed or pasted. The tables are the ones
CORE_Params reads that way, so the list follows the engine. build_template.build
writes the rules into the template's sheet XML after the headless build, because
LibreOffice's VBA compatibility cannot store a conditional format; M_Setup
applies the same rule when Excel builds or resets a sheet.

The rule is an expression, COUNTIF($C$7:$C$9,C7)>1, with a red fill and text.
COUNTIF ignores case, as the run does; it does not trim blanks, which the run
does, so a key with a trailing blank is not highlighted.

Usage:
    python3 tools/key_highlights.py FILE.xlsx   # list the rules of a workbook
"""

from __future__ import annotations

import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl

import input_validation

REPO_ROOT = Path(__file__).resolve().parents[1]
NS = input_validation.NS
FILL = "FFFFC7CE"  # COLOR_FAIL_FILL
TEXT = "FF9C0006"  # COLOR_FAIL_TEXT
KEY_COLUMN = "C"   # a tagged table starts in column C (input_validation.table_rows)
AFTER_FORMATS = ("dataValidations", *input_validation.AFTER_VALIDATIONS)


@dataclass(frozen=True)
class Highlight:
    sheet: str
    sqref: str
    formula: str


def key_tables() -> list[tuple[str, str]]:
    """(sheet, tag) of every table the run reads as Key | Value, from CORE_Params."""
    config = (REPO_ROOT / "src/modules/CORE_Config.bas").read_text(encoding="cp1252")
    params = (REPO_ROOT / "src/modules/CORE_Params.bas").read_text(encoding="cp1252")
    names = dict(re.findall(r'Const (SH_\w+) As String = "(\w+)"', config))
    return [(names[sheet], tag) for sheet, tag in re.findall(r'ReadKeyValues\((SH_\w+), "(\w+)"\)', params)]


def expected_highlights(path: Path) -> list[Highlight]:
    wb = openpyxl.load_workbook(path)
    out = []
    for sheet, tag in key_tables():
        first, last = input_validation.table_rows(wb[sheet], tag)
        out.append(Highlight(sheet, f"{KEY_COLUMN}{first}:{KEY_COLUMN}{last}",
                             f"COUNTIF(${KEY_COLUMN}${first}:${KEY_COLUMN}${last},{KEY_COLUMN}{first})>1"))
    return out


def _tag(name: str) -> str:
    return f"{{{NS}}}{name}"


def _dxf(styles: ET.Element) -> int:
    """Append the red style and return its index."""
    dxfs = styles.find(_tag("dxfs"))
    if dxfs is None:
        dxfs = ET.Element(_tag("dxfs"))
        anchor = styles.find(_tag("cellStyles"))
        assert anchor is not None  # required part of styles.xml
        styles.insert(list(styles).index(anchor) + 1, dxfs)
    dxf = ET.SubElement(dxfs, _tag("dxf"))
    ET.SubElement(ET.SubElement(dxf, _tag("font")), _tag("color"), rgb=TEXT)
    ET.SubElement(ET.SubElement(ET.SubElement(dxf, _tag("fill")), _tag("patternFill")), _tag("bgColor"), rgb=FILL)
    dxfs.set("count", str(len(dxfs)))
    return len(dxfs) - 1


def _write(root: ET.Element, highlights: list[Highlight], dxf_id: int) -> None:
    for old in root.findall(_tag("conditionalFormatting")):
        root.remove(old)
    position = len(root)
    for index, child in enumerate(root):
        if child.tag.removeprefix(f"{{{NS}}}") in AFTER_FORMATS:
            position = index
            break
    for priority, highlight in enumerate(highlights, start=1):
        block = ET.Element(_tag("conditionalFormatting"), sqref=highlight.sqref)
        rule = ET.SubElement(block, _tag("cfRule"), type="expression", dxfId=str(dxf_id), priority=str(priority))
        ET.SubElement(rule, _tag("formula")).text = highlight.formula
        root.insert(position + priority - 1, block)


def apply(path: Path) -> None:
    """Write the highlights of the file's own cells into its sheets."""
    highlights = expected_highlights(path)
    with zipfile.ZipFile(path) as archive:
        parts = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    styles = ET.fromstring(parts["xl/styles.xml"])
    dxf_id = _dxf(styles)
    parts["xl/styles.xml"] = ET.tostring(styles, encoding="utf-8", xml_declaration=True)
    by_sheet: dict[str, list[Highlight]] = {}
    for highlight in highlights:
        by_sheet.setdefault(highlight.sheet, []).append(highlight)
    for name, part in input_validation._sheet_parts(parts).items():
        root = ET.fromstring(parts[part])
        _write(root, by_sheet.get(name, []), dxf_id)
        parts[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as temp:
        replacement = Path(temp.name)
    try:
        with zipfile.ZipFile(replacement, "w", zipfile.ZIP_DEFLATED) as archive:
            for filename, content in parts.items():
                archive.writestr(filename, content)
        replacement.replace(path)
    finally:
        replacement.unlink(missing_ok=True)


def read_highlights(path: Path) -> list[Highlight]:
    """The expression rules with the red style that a workbook stores."""
    with zipfile.ZipFile(path) as archive:
        parts = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    styles = ET.fromstring(parts["xl/styles.xml"])
    dxfs = styles.find(_tag("dxfs"))
    red = set()
    if dxfs is not None:
        for index, dxf in enumerate(dxfs):
            color = dxf.find(f"{_tag('font')}/{_tag('color')}")
            fill = dxf.find(f"{_tag('fill')}/{_tag('patternFill')}/{_tag('bgColor')}")
            if color is not None and fill is not None and color.get("rgb") == TEXT and fill.get("rgb") == FILL:
                red.add(str(index))
    out = []
    for name, part in input_validation._sheet_parts(parts).items():
        for block in ET.fromstring(parts[part]).findall(_tag("conditionalFormatting")):
            for rule in block.findall(_tag("cfRule")):
                formula = rule.find(_tag("formula"))
                if rule.get("type") == "expression" and rule.get("dxfId") in red and formula is not None:
                    out.append(Highlight(name, block.attrib["sqref"], formula.text or ""))
    return out


def problems(path: Path) -> list[str]:
    """How the highlights stored in a template differ from the ones it must carry."""
    stored, expected = read_highlights(path), expected_highlights(path)
    out = [f"key highlight missing: {h.sheet}!{h.sqref}" for h in expected if h not in stored]
    out += [f"unexpected key highlight: {h.sheet}!{h.sqref}" for h in stored if h not in expected]
    return out


def main(argv: list[str]) -> int:
    if len(argv) != 1:
        print(__doc__, file=sys.stderr)
        return 2
    for highlight in read_highlights(Path(argv[0])):
        print(f"{highlight.sheet}!{highlight.sqref}: {highlight.formula}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
