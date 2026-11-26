"""Repository tooling checks: VBA static lint and generated expected values."""

import subprocess
import sys
from pathlib import Path

import vba_lint
from conftest import require_libreoffice

REPO_ROOT = Path(__file__).resolve().parents[1]


def test_vba_lint_is_clean():
    assert vba_lint.main() == 0


def test_vba_lint_rejects_a_declaration_after_a_procedure(tmp_path):
    # Excel accepts only comments after End Sub; LibreOffice compiled it (#327)
    module = tmp_path / "M_Sample.bas"
    module.write_bytes(
        b'Attribute VB_Name = "M_Sample"\r\n'
        b"' MODULE: M_Sample\r\n"
        b"Option Explicit\r\n"
        b"Private Const FIRST As Long = 1\r\n"
        b"Public Sub Run()\r\n"
        b"End Sub\r\n"
        b"' a comment may follow\r\n"
        b'Private Const LATE As String = "x"\r\n'
        b"Public Sub Again()\r\n"
        b"End Sub\r\n"
    )
    _, errors = vba_lint.parse(module)
    assert errors == ["M_Sample.bas:8: only comments may appear after End Sub, End Function or End Property"]


def test_expected_values_in_setup_are_current():
    require_libreoffice()
    proc = subprocess.run(
        [sys.executable, str(REPO_ROOT / "tools" / "update_expected.py"), "--check"],
        capture_output=True,
        text=True,
        timeout=300,
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr


def test_the_build_turns_printer_communication_off_while_it_sets_page_setup():
    # Each PageSetup change queries the default printer; an unreachable one
    # made the Excel build hang (#295). The build helper turns it off with
    # alerts and screen updating, ignored by hosts without the property.
    setup = (REPO_ROOT / "src" / "modules" / "M_Setup.bas").read_text(encoding="cp1252")
    helper = setup.split("Private Sub SetAlerts(ByVal enabled As Boolean)", 1)[1].split("End Sub", 1)[0]
    code = "\n".join(line for line in helper.splitlines() if line.strip() and not line.strip().startswith("'"))
    assert code.lstrip().startswith("On Error Resume Next")
    assert "Application.PrintCommunication = enabled" in helper
    build = setup.split("Private Sub BuildAllSheets()", 1)[1].split("End Sub", 1)[0]
    assert build.index("SetAlerts False") < build.index("BuildReadme") < build.index("SetAlerts True")


def test_the_scratch_cell_of_the_duplicate_key_rule_is_not_the_last_cell_of_the_sheet():
    # Writing to the last cell stretches the used range to the last row and column in Excel, and the
    # acceptance checks that walk UsedRange cell by cell then do not end (#368).
    setup = (REPO_ROOT / "src" / "modules" / "M_Setup.bas").read_text(encoding="cp1252")
    function = setup.split("Private Function LocalFormula(", 1)[1].split("End Function", 1)[0]
    code = "\n".join(line for line in function.splitlines() if not line.strip().startswith("'"))
    assert "Rows.Count" not in code and "Columns.Count" not in code
    assert "ws.UsedRange" in code


def test_no_trailing_comment_touches_the_code_before_it():
    # The VBE puts a space before a trailing comment when it exports a module, so a comment that touches its code
    # fails the round trip of the development workbook (found in Excel on d8ff9b5).
    touching = []
    for path in sorted([*(REPO_ROOT / "src" / "modules").glob("*.bas"), *(REPO_ROOT / "tests" / "modules").glob("*.bas")]):
        for number, line in enumerate(path.read_text(encoding="cp1252").splitlines(), 1):
            in_string = False
            for index, char in enumerate(line):
                if char == '"':
                    in_string = not in_string
                elif char == "'" and not in_string:
                    if index and line[:index].strip() and not line[index - 1].isspace():
                        touching.append(f"{path.name}:{number}")
                    break
    assert not touching, touching
