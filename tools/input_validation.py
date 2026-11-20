"""Worksheet data validation of the input and parameter sheets (#348).

The rules are defined here once. build_template.build writes them into the
template's sheet XML after the headless build, because LibreOffice's VBA
compatibility cannot store custom rules or literal lists. A workbook built
from the template carries them; a workbook built from source with
tools/Build-Workbook.ps1 -FromSource does not, and the guide says so.

A rule guards typing and pasting into a cell. Formulas never return an error:
Excel evaluates every argument of AND and OR, so a text cell inside ABS or INT
sits behind an IF. The run's own validation stays
the authority: a pasted value bypasses a rule, and the run rejects the same
rows with the same messages whatever the rules say. Rules never accept what
the run rejects, with these exceptions that only narrow what a person can
type: text lists are case-sensitive in Excel while the run ignores case, and
bucket numbers must be numbers, not numeric text.

Usage:
    python3 tools/input_validation.py FILE.xlsx   # list the rules of a workbook
"""

from __future__ import annotations

import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils import get_column_letter

REPO_ROOT = Path(__file__).resolve().parents[1]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS)
ET.register_namespace("r", REL)

#: Child elements of a worksheet that follow dataValidations, in schema order.
AFTER_VALIDATIONS = (
    "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
    "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
)

INPUT_SHEETS = ("In_Delta", "In_Vega", "In_Curvature", "In_DRC", "In_RRAO")
RISK_CLASSES = ("GIRR", "CSR_NS", "EQ", "COM", "FX")
FLAGS = ("TRUE", "FALSE")

#: Sub-types the run accepts per input sheet and risk class (CORE_SBM). A blank
#: sub-type is always accepted; a commodity sub-type is free text; FX takes none.
SUBTYPES: dict[str, dict[str, tuple[str, ...]]] = {
    "In_Delta": {"GIRR": ("YIELD", "INFLATION", "XCCY"), "CSR_NS": ("BOND", "CDS"), "EQ": ("SPOT", "REPO")},
    "In_Vega": {"GIRR": ("YIELD", "INFLATION", "XCCY"), "CSR_NS": ("BOND", "CDS"), "EQ": ("SPOT",)},
    "In_Curvature": {"GIRR": ("YIELD",), "CSR_NS": ("BOND", "CDS"), "EQ": ("SPOT",)},
}
SUBTYPE_MESSAGE = {
    "In_Delta": "GIRR: YIELD, INFLATION or XCCY. CSR_NS: BOND or CDS. EQ: SPOT or REPO. COM: free text. FX: none. Blank takes the default.",
    "In_Vega": "GIRR: YIELD, INFLATION or XCCY. CSR_NS: BOND or CDS. EQ: SPOT. COM: free text. FX: none. Blank takes the default.",
    "In_Curvature": "GIRR: YIELD. CSR_NS: BOND or CDS. EQ: SPOT. COM: free text. FX: none. Blank takes the default.",
}


#: Input columns without a rule, and why. Every other input column has one.
UNVALIDATED: dict[tuple[str, str], str] = {
    **{(sheet, "ID"): "A free label; the run does not read it" for sheet in INPUT_SHEETS},
    **{(sheet, "Comment"): "Free text; the run does not read it" for sheet in INPUT_SHEETS},
    **{(sheet, "Qualifier"): "Free text; the run requires it for most risk classes and reports a blank one"
       for sheet in ("In_Delta", "In_Vega", "In_Curvature")},
    ("In_DRC", "Obligor"): "Free text; the run reports a blank obligor",
}


@dataclass(frozen=True)
class Rule:
    sheet: str
    sqref: str
    kind: str          # "list" (a literal list), "range" (a list from cells) or "custom" (a formula)
    formula: str       # the list text, the range, or the formula without "="
    alert: str         # "stop" or "warning"
    title: str
    message: str

    def __post_init__(self) -> None:
        assert self.kind in ("list", "range", "custom"), self
        assert self.alert in ("stop", "warning"), self
        assert 0 < len(self.title) <= 32, f"title over 32 characters: {self.title!r}"
        assert 0 < len(self.message) <= 255, f"message over 255 characters: {self.message!r}"
        if self.kind == "list":
            assert len(self.formula) <= 255, self


def _layout() -> tuple[int, int]:
    """First and last data row of the input sheets, from the VBA constants."""
    config = (REPO_ROOT / "src/modules/CORE_Config.bas").read_text(encoding="cp1252")
    setup = (REPO_ROOT / "src/modules/M_Setup.bas").read_text(encoding="cp1252")
    first = re.search(r"Const LAYOUT_FIRST_DATA_ROW As Long = (\d+)", config)
    count = re.search(r"Const INPUT_ROWS As Long = (\d+)", setup)
    assert first is not None and count is not None
    return int(first.group(1)), int(first.group(1)) + int(count.group(1)) - 1


def _header_name(value: Any) -> str:
    """The column name an input sheet header shows, as UNVALIDATED keys it."""
    return str(value).strip()


def _text(value: Any) -> str:
    return value.strip() if isinstance(value, str) else ""


def table_rows(ws: Any, tag: str) -> tuple[int, int]:
    """Data rows of the tagged table [TAG] in column C: from two rows below its
    heading to the last row before a blank key."""
    for row in range(1, ws.max_row + 1):
        if _text(ws.cell(row, 3).value).startswith(f"[{tag}]"):
            first = row + 2
            last = first
            while ws.cell(last + 1, 3).value not in (None, "") and not _text(ws.cell(last + 1, 3).value).startswith("["):
                last += 1
            return first, last
    raise KeyError(f"{ws.title}: no table [{tag}]")


def _custom(sheet: str, sqref: str, formula: str, title: str, message: str, alert: str = "stop") -> Rule:
    return Rule(sheet, sqref, "custom", formula, alert, title, message)


def _numeric(sheet: str, column: str, first: int, last: int, title: str, message: str, minimum: bool = False) -> Rule:
    cell = f"{column}{first}"
    test = f"AND(ISNUMBER({cell}),{cell}>=0)" if minimum else f"ISNUMBER({cell})"
    return _custom(sheet, f"{column}{first}:{column}{last}", f'OR({cell}="",{test})', title, message)


def _risk_class_rules(sheet: str, first: int, last: int) -> list[Rule]:
    r = first
    rc = f"TRIM($D{r})"
    bucket = (
        f'OR(E{r}="",{rc}="",'
        f'AND(OR({rc}="GIRR",{rc}="FX"),LEN(TRIM(E{r}))=3,NOT(ISNUMBER(E{r}))),'
        f'AND(OR({rc}="CSR_NS",{rc}="EQ",{rc}="COM"),IF(ISNUMBER(E{r}),AND(E{r}>=1,E{r}=INT(E{r})),FALSE)))'
    )
    parts = [f'TRIM(G{r})=""', f'{rc}=""', f'{rc}="COM"']
    for risk_class, names in SUBTYPES[sheet].items():
        options = ",".join(f'TRIM(G{r})="{name}"' for name in names)
        parts.append(f'AND({rc}="{risk_class}",OR({options}))')
    return [
        Rule(sheet, f"D{first}:D{last}", "list", ",".join(RISK_CLASSES), "stop", "Invalid risk class",
             "Use GIRR, CSR_NS, EQ, COM or FX."),
        _custom(sheet, f"E{first}:E{last}", bucket, "Invalid bucket",
                "GIRR and FX: a three-letter currency. CSR_NS, EQ and COM: a bucket number of 1 or more, typed as a number."),
        _custom(sheet, f"G{first}:G{last}", f"OR({','.join(parts)})", "Invalid sub-type", SUBTYPE_MESSAGE[sheet]),
    ]


def _flag(sheet: str, sqref: str) -> Rule:
    return Rule(sheet, sqref, "list", ",".join(FLAGS), "stop", "Invalid flag", "Enter TRUE or FALSE, or leave the cell blank.")


def _list_from(sheet: str, column: str, first: int, last: int, source: str, title: str, parameter_sheet: str) -> Rule:
    return Rule(sheet, f"{column}{first}:{column}{last}", "range", source, "warning", title,
                f"Not in the table on {parameter_sheet}. Add it there first, or continue if you are about to.")


def input_rules(wb: Any) -> list[Rule]:
    """Rules of Control and the five input sheets."""
    first, last = _layout()
    rules = [
        _custom("Control", "D7", 'OR(D7="",AND(ISTEXT(D7),LEN(TRIM(D7))=3,'
                'SUMPRODUCT(--ISNUMBER(FIND(MID(UPPER(TRIM(D7)),{1,2,3},1),"ABCDEFGHIJKLMNOPQRSTUVWXYZ")))=3))',
                "Check the currency", "A currency code has three letters, such as EUR. The run reads the text as typed.",
                alert="warning"),
        _flag("Control", "D8:D11"),
    ]
    # In_Delta: ID C, risk class D, bucket E, qualifier F, sub-type G, tenor H, sensitivity I
    rules += _risk_class_rules("In_Delta", first, last)
    rules += [
        _numeric("In_Delta", "H", first, last, "Invalid tenor", "A number of years of zero or more, or blank.", minimum=True),
        _numeric("In_Delta", "I", first, last, "Invalid sensitivity", "A number, or blank."),
    ]
    # In_Vega: option maturity H, underlying maturity I, sensitivity J
    rules += _risk_class_rules("In_Vega", first, last)
    rules += [
        _numeric("In_Vega", "H", first, last, "Invalid maturity", "A number of years of zero or more, or blank.", minimum=True),
        _numeric("In_Vega", "I", first, last, "Invalid maturity", "A number of years of zero or more, or blank.", minimum=True),
        _numeric("In_Vega", "J", first, last, "Invalid sensitivity", "A number, or blank."),
    ]
    # In_Curvature: CVR up H, CVR down I, references reporting currency J
    rules += _risk_class_rules("In_Curvature", first, last)
    rules += [
        _numeric("In_Curvature", "H", first, last, "Invalid CVR", "A number, or blank."),
        _numeric("In_Curvature", "I", first, last, "Invalid CVR", "A number, or blank."),
        _flag("In_Curvature", f"J{first}:J{last}"),
    ]
    # In_DRC: obligor D, bucket E, credit quality F, seniority G, notional H, market value I, maturity J, LGD K
    drc = wb["P_DRC"]
    for column, tag, title in (("E", "DRC_BUCKETS", "Unknown bucket"), ("F", "DRC_RW", "Unknown credit quality"),
                               ("G", "DRC_SENIORITY", "Unknown seniority")):
        a, b = table_rows(drc, tag)
        rules.append(_list_from("In_DRC", column, first, last, f"P_DRC!$C${a}:$C${b}", title, "P_DRC"))
    rules += [
        _numeric("In_DRC", "H", first, last, "Invalid amount", "A number, or blank."),
        _numeric("In_DRC", "I", first, last, "Invalid amount", "A number, or blank."),
        _numeric("In_DRC", "J", first, last, "Invalid maturity", "A number of years, or blank for one year."),
        _numeric("In_DRC", "K", first, last, "Invalid LGD", "A number, or blank to use the seniority's LGD."),
    ]
    # In_RRAO: category D, gross notional E
    a, b = table_rows(wb["P_RRAO"], "RRAO_RW")
    rules.append(_list_from("In_RRAO", "D", first, last, f"P_RRAO!$C${a}:$C${b}", "Unknown category", "P_RRAO"))
    rules.append(_numeric("In_RRAO", "E", first, last, "Invalid gross notional",
                          "A number of zero or more, or blank.", minimum=True))
    return rules


def _range_of(cells: list[int], column: str) -> list[str]:
    """Contiguous runs of rows as A1 ranges in one column."""
    runs, start, previous = [], cells[0], cells[0]
    for row in cells[1:]:
        if row != previous + 1:
            runs.append((start, previous))
            start = row
        previous = row
    runs.append((start, previous))
    return [f"{column}{a}:{column}{b}" if a != b else f"{column}{a}" for a, b in runs]


def parameter_rules(wb: Any) -> list[Rule]:
    """A number rule for every numeric cell of the P_ sheets, grouped per table
    and column. Risk weights must be zero or more, and correlations lie between
    -1 and 1; those two only warn, because the run accepts any number."""
    rules: list[Rule] = []
    for ws in wb.worksheets:
        if not ws.title.startswith("P_"):
            continue
        tag, header = None, 0
        blocks: dict[tuple[str, int], list[int]] = {}
        headers: dict[tuple[str, int], str] = {}
        tags: dict[str, str] = {}
        for row in range(1, ws.max_row + 1):
            key = _text(ws.cell(row, 3).value)
            if key.startswith("["):
                tag, header = key[1:].split("]")[0], row + 1
                continue
            if tag is None or row <= header:
                continue
            for column in range(3, ws.max_column + 1):
                value = ws.cell(row, column).value
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    letter = get_column_letter(column)
                    blocks.setdefault((tag, column), []).append(row)
                    headers[(tag, column)] = _text(ws.cell(header, column).value) or str(ws.cell(header, column).value)
                    tags[tag] = tag
        for (tag, column), rows in sorted(blocks.items(), key=lambda item: (item[0][1], item[1][0])):
            letter = get_column_letter(column)
            title_of = headers[(tag, column)]
            for sqref in _range_of(rows, letter):
                first_cell = sqref.split(":")[0]
                if title_of.upper().startswith("RW"):
                    rules.append(_custom(ws.title, sqref, f'OR({first_cell}="",AND(ISNUMBER({first_cell}),{first_cell}>=0))',
                                         "Check the risk weight", "A risk weight is a number of zero or more.", alert="warning"))
                elif "GAMMA" in tag or title_of.upper().startswith("RHO"):
                    rules.append(_custom(ws.title, sqref, f'OR({first_cell}="",IF(ISNUMBER({first_cell}),ABS({first_cell})<=1,FALSE))',
                                         "Check the correlation", "A correlation lies between -1 and 1.", alert="warning"))
                else:
                    rules.append(_custom(ws.title, sqref, f'OR({first_cell}="",ISNUMBER({first_cell}))',
                                         "Invalid parameter", "Enter a number."))
    return rules


def expected_rules(path: Path) -> list[Rule]:
    """The rules a template must carry, from the cells of the file itself."""
    wb = openpyxl.load_workbook(path)
    return input_rules(wb) + parameter_rules(wb)


# --------------------------------------------------------------------------
# Writing and reading the sheet XML
# --------------------------------------------------------------------------
def _tag(name: str) -> str:
    return f"{{{NS}}}{name}"


def _element(rule: Rule) -> ET.Element:
    attributes = {"type": "list" if rule.kind in ("list", "range") else "custom"}
    if rule.alert == "warning":
        attributes["errorStyle"] = "warning"
    attributes.update(allowBlank="1", showErrorMessage="1", errorTitle=rule.title, error=rule.message, sqref=rule.sqref)
    element = ET.Element(_tag("dataValidation"), attributes)
    ET.SubElement(element, _tag("formula1")).text = f'"{rule.formula}"' if rule.kind == "list" else rule.formula
    return element


def _sheet_parts(parts: dict[str, bytes]) -> dict[str, str]:
    book = ET.fromstring(parts["xl/workbook.xml"])
    relations = {r.attrib["Id"]: r.attrib["Target"] for r in ET.fromstring(parts["xl/_rels/workbook.xml.rels"])}
    sheets = book.find(_tag("sheets"))
    assert sheets is not None
    return {s.attrib["name"]: "xl/" + relations[s.attrib[f"{{{REL}}}id"]].lstrip("/").removeprefix("xl/") for s in sheets}


def _write_rules(root: ET.Element, rules: list[Rule]) -> None:
    for old in root.findall(_tag("dataValidations")):
        root.remove(old)
    if not rules:
        return
    block = ET.Element(_tag("dataValidations"), count=str(len(rules)))
    block.extend(_element(rule) for rule in rules)
    position = len(root)
    for index, child in enumerate(root):
        if child.tag.removeprefix(f"{{{NS}}}") in AFTER_VALIDATIONS:
            position = index
            break
    root.insert(position, block)


def apply(path: Path) -> None:
    """Write the rules of the file's own cells into its sheets."""
    rules = expected_rules(path)
    with zipfile.ZipFile(path) as archive:
        parts = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    by_sheet: dict[str, list[Rule]] = {}
    for rule in rules:
        by_sheet.setdefault(rule.sheet, []).append(rule)
    for name, part in _sheet_parts(parts).items():
        root = ET.fromstring(parts[part])
        _write_rules(root, by_sheet.get(name, []))
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


def read_rules(path: Path) -> list[Rule]:
    """The rules stored in a workbook, in sheet and file order."""
    with zipfile.ZipFile(path) as archive:
        parts = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    rules: list[Rule] = []
    for name, part in _sheet_parts(parts).items():
        for node in ET.fromstring(parts[part]).findall(f"{_tag('dataValidations')}/{_tag('dataValidation')}"):
            formula = (node.findtext(_tag("formula1")) or "")
            if node.get("type") == "custom":
                kind = "custom"
            elif formula.startswith('"'):
                kind, formula = "list", formula.strip('"')
            else:
                kind = "range"
            rules.append(Rule(name, node.attrib["sqref"], kind, formula, "warning" if node.get("errorStyle") == "warning" else "stop",
                              node.get("errorTitle", ""), node.get("error", "")))
    return rules


def problems(path: Path) -> list[str]:
    """How the rules stored in a template differ from the rules it must carry."""
    stored, expected = read_rules(path), expected_rules(path)
    out = [f"validation rule missing: {r.sheet}!{r.sqref} ({r.title})" for r in expected if r not in stored]
    out += [f"unexpected validation rule: {r.sheet}!{r.sqref} ({r.title})" for r in stored if r not in expected]
    return out


def main(argv: list[str]) -> int:
    if len(argv) != 1:
        print(__doc__, file=sys.stderr)
        return 2
    for rule in read_rules(Path(argv[0])):
        print(f"{rule.sheet}!{rule.sqref}: {rule.kind} {rule.alert} {rule.title} | {rule.formula}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
