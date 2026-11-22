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
