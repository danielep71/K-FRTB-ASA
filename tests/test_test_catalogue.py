"""Every regression item has an area, a description, a source and a source status (#352)."""

from __future__ import annotations

import openpyxl

import build_template
import input_validation
import test_catalogue
import update_expected
from frtb_reference import COMPONENTS, RISK_CLASSES, SCENARIOS

FIRST_COL = 3


def expected_names() -> list[str]:
    summary = list(update_expected.SUMMARY_ITEMS)
    return summary + [f"{rc}|{comp}|{scen}" for rc in RISK_CLASSES for comp in COMPONENTS for scen in SCENARIOS]


def test_the_catalogue_describes_every_item_and_a_missing_one_raises():
    names = expected_names()
    assert len(names) == 52
    for name in names:
        area, description, source, status = test_catalogue.describe(name)
        assert all(text.strip() for text in (area, description, source, status)), name
        assert "independent reference" in source, name
    for unknown in ("SBM_EXTRA", "GIRR|DELTA", "XX|DELTA|LOW", "GIRR|THETA|LOW", "GIRR|DELTA|VERY_HIGH"):
        try:
            test_catalogue.describe(unknown)
        except (KeyError, ValueError):
            continue
        raise AssertionError(f"{unknown} was described")


def test_the_template_lists_every_item_with_its_catalogue_columns():
    ws = openpyxl.load_workbook(build_template.TEMPLATE)["Test_Expected"]
    first, last = input_validation.table_rows(ws, "EXPECTED")
    header = [str(ws.cell(first - 1, FIRST_COL + k).value).replace("\n", " ") for k in range(9)]
    assert header[5:] == ["Area", "Description", "Source", "Source status"]
    items = []
    for row in range(first, last + 1):
        item, *_, area, description, source, status = (ws.cell(row, FIRST_COL + k).value for k in range(9))
        items.append(item)
        assert (area, description, source, status) == test_catalogue.describe(str(item)), item
    assert items == expected_names()
