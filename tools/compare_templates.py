"""List the differences between two workbook templates (#274).

tools/Check-FRTB-Excel.ps1 compares the template that desktop Excel builds
from source with the committed template src/workbook/FRTB_Template.xlsx,
which LibreOffice builds. It reports the differences for a person to explain
(fonts, widths, number formats); it never fails on them, because the two
hosts are expected to differ in presentation (#28).

Usage:
    python tools/compare_templates.py COMMITTED.xlsx OTHER.xlsx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

SHOWN = 20  # differences listed per kind; the rest are counted


def _load(path: Path) -> Any:
    import openpyxl

    return openpyxl.load_workbook(path)


def _cells(ws: Any) -> dict[str, tuple[Any, str]]:
    return {
        cell.coordinate: (cell.value, cell.number_format)
        for row in ws.iter_rows()
        for cell in row
        if cell.value is not None
    }


def _pane(ws: Any) -> str:
    pane = ws.sheet_view.pane
    if pane is None or pane.state != "frozen":
        return "none"
    return f"{int(pane.ySplit or 0)} rows, {int(pane.xSplit or 0)} columns"


def compare(committed: Any, other: Any) -> list[str]:
    """Return one line per difference, grouped by kind and sheet."""
    lines: list[str] = []
    a_names, b_names = committed.sheetnames, other.sheetnames
    if a_names != b_names:
        lines.append(f"sheets: committed {a_names}; other {b_names}")
    for name in (n for n in a_names if n in b_names):
        a, b = committed[name], other[name]
        a_cells, b_cells = _cells(a), _cells(b)
        values = [
            f"{name}!{ref}: {a_cells.get(ref, (None, ''))[0]!r} -> {b_cells.get(ref, (None, ''))[0]!r}"
            for ref in sorted(set(a_cells) | set(b_cells))
            if a_cells.get(ref, (None, ""))[0] != b_cells.get(ref, (None, ""))[0]
        ]
        formats = [
            f"{name}!{ref}: {a_cells[ref][1]!r} -> {b_cells[ref][1]!r}"
            for ref in sorted(set(a_cells) & set(b_cells))
            if a_cells[ref][1] != b_cells[ref][1]
        ]
        widths = []
        for letter in sorted(set(a.column_dimensions) | set(b.column_dimensions)):
            wa, wb = a.column_dimensions[letter].width, b.column_dimensions[letter].width
            if wa and wb and abs(wa - wb) > 0.5:
                widths.append(f"{name}!{letter}: width {wa:.2f} -> {wb:.2f}")
        for kind, items in (("values", values), ("number formats", formats), ("column widths", widths)):
            if items:
                lines.append(f"{name}: {len(items)} {kind} differ")
                lines += ["  " + item for item in items[:SHOWN]]
                if len(items) > SHOWN:
                    lines.append(f"  ... and {len(items) - SHOWN} more")
        if _pane(a) != _pane(b):
            lines.append(f"{name}: frozen panes {_pane(a)} -> {_pane(b)}")
    return lines


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("committed", type=Path)
    parser.add_argument("other", type=Path)
    args = parser.parse_args(argv)
    try:
        lines = compare(_load(args.committed), _load(args.other))
    except ImportError:
        print("not compared: openpyxl is not installed (pip install -r tests/requirements.txt)")
        return 0
    except (OSError, ValueError, KeyError) as exc:
        print(f"not compared: {exc}")
        return 2
    print(f"committed: {args.committed}")
    print(f"other:     {args.other}")
    if not lines:
        print("no differences in sheets, values, number formats, column widths or frozen panes")
    for line in lines:
        print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
