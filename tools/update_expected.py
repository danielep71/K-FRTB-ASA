"""Regenerate the expected results embedded in src/modules/M_Setup.bas.

The expected values for the sample portfolio come from the independent
Python reference (tests/frtb_reference.py), never from the VBA engine
itself. The script builds the workbook headless in LibreOffice (setup only,
so parameters and sample inputs are exactly what M_Setup writes), computes
the reference results and rewrites the generated block of WriteExpectedRows;
when the block changes, it also rebuilds src/workbook/FRTB_Template.xlsx.

Usage:
    python3 tools/update_expected.py          # rewrite M_Setup.bas
    python3 tools/update_expected.py --check  # exit 1 if the block is stale
"""

from __future__ import annotations

import argparse
import re
import sys
import tempfile
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

import openpyxl

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "tests"))
sys.path.insert(0, str(REPO_ROOT / "tools"))

import frtb_reference as ref  # noqa: E402
import lo_runner  # noqa: E402

SETUP = REPO_ROOT / "src" / "modules" / "M_Setup.bas"
BEGIN = "    ' BEGIN GENERATED EXPECTED VALUES"
END = "    ' END GENERATED EXPECTED VALUES"
SUMMARY_ITEMS = ("SBM_LOW", "SBM_MEDIUM", "SBM_HIGH", "SBM", "DRC_NONSEC", "RRAO", "TOTAL")


def expected_items(result: Mapping[str, Any]) -> list[tuple[str, float]]:
    items = [(k, float(result["summary"][k])) for k in SUMMARY_ITEMS]
    for rc in ref.RISK_CLASSES:
        for comp in ref.COMPONENTS:
            for scen in ref.SCENARIOS:
                items.append((f"{rc}|{comp}|{scen}", float(result["capital"][(rc, comp, scen)])))
    return items


def vba_literal(x: float) -> str:
    s = repr(round(x, 6))
    return s[:-2] if s.endswith(".0") else s


def generated_block(items: Sequence[tuple[str, float]], newline: str) -> str:
    lines = [BEGIN]
    for i, (name, value) in enumerate(items):
        lines.append(f'    PutRow ws, r + {i}, Array("{name}", {vba_literal(value)})')
    lines.append(f"    WriteExpectedRows = {len(items)}")
    lines.append(END)
    return newline.join(lines)


def compute() -> list[tuple[str, float]]:
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "setup.xlsx"
        errors = lo_runner.run(out, ["M_Setup.BuildWorkbook"])
        if errors:
            raise RuntimeError("; ".join(errors))
        wb = openpyxl.load_workbook(out, data_only=True)
        result = ref.run(ref.params_from_workbook(wb), ref.inputs_from_workbook(wb))
    return expected_items(result)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args(argv)
    with SETUP.open(encoding="cp1252", newline="") as fh:
        text = fh.read()
    newline = "\r\n" if "\r\n" in text else "\n"
    pattern = re.compile(re.escape(BEGIN) + r".*?" + re.escape(END), re.S)
    if not pattern.search(text):
        print("Generated block markers not found in M_Setup.bas", file=sys.stderr)
        return 2
    items = compute()
    new_text = pattern.sub(lambda _m: generated_block(items, newline), text)
    if args.check:
        if new_text != text:
            print("M_Setup.bas expected values are stale; run tools/update_expected.py", file=sys.stderr)
            return 1
        print("M_Setup.bas expected values are up to date")
        return 0
    if new_text == text:
        print("M_Setup.bas expected values already current")
        return 0
    SETUP.write_text(new_text, encoding="cp1252", newline="")
    print("M_Setup.bas updated")
    # The template carries the expected values too: rebuild it from the new source
    import build_template  # noqa: PLC0415

    build_template.build(build_template.TEMPLATE)
    print("src/workbook/FRTB_Template.xlsx rebuilt")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
