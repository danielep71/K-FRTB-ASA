"""The highlight of a repeated key in the Key | Value parameter tables (#350)."""

from __future__ import annotations

import re
import shutil
import zipfile

import build_template
import key_highlights

SETUP = key_highlights.REPO_ROOT / "src" / "modules" / "M_Setup.bas"


def test_the_template_highlights_a_repeated_key_in_every_key_value_table():
    assert key_highlights.problems(build_template.TEMPLATE) == []
    stored = key_highlights.read_highlights(build_template.TEMPLATE)
    tables = key_highlights.key_tables()
    assert len(tables) == 10 and len(stored) == 10
    assert {(h.sheet) for h in stored} == {sheet for sheet, _ in tables}
    assert all(h.formula.startswith("AND(C") and "COUNTIF($C$" in h.formula and h.formula.endswith(")>1)") for h in stored)
    # a single-row table such as [EQ_PARAMS] covers its row and the one below it
    eq = next(h for h in stored if h.sheet == "P_EQ")
    first, last = (int(part[1:]) for part in eq.sqref.split(":"))
    assert last == first + 1


def test_the_vba_marks_the_same_tables_as_the_engine_reads():
    """M_Setup applies the rule where Excel builds or resets a sheet, to the tables CORE_Params reads."""
    lines = SETUP.read_text(encoding="cp1252").splitlines()
    marked, last_tag = [], ""
    for line in lines:
        tag = re.search(r'PutTable\(\w+, [^"]*"(\w+)"', line)
        if tag:
            last_tag = tag.group(1)
        if re.search(r"FormatParamRows \w+,", line) and line.rstrip().endswith(", True"):
            marked.append(last_tag)
    assert sorted(marked) == sorted(tag for _, tag in key_highlights.key_tables())


def test_applying_the_highlights_twice_changes_nothing_and_a_missing_one_is_reported(tmp_path):
    copy = tmp_path / "copy.xlsx"
    shutil.copy(build_template.TEMPLATE, copy)
    before = key_highlights.read_highlights(copy)
    key_highlights.apply(copy)
    assert key_highlights.read_highlights(copy) == before
    # Remove the rules of one sheet from the XML
    stripped = tmp_path / "stripped.xlsx"
    with zipfile.ZipFile(copy) as src, zipfile.ZipFile(stripped, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and b"conditionalFormatting" in data:
                data = re.sub(rb"<conditionalFormatting.*?</conditionalFormatting>", b"", data, count=1, flags=re.S)
            out.writestr(item.filename, data)
    problems = key_highlights.problems(stripped)
    assert len(problems) >= 1 and all(p.startswith("key highlight missing") for p in problems)
