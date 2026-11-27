"""Protection for the interface only: the sheets stay protected for the user, and a macro writes without opening them (#349, #353).

Excel alone shows it; here the text of the modules is checked, and the LibreOffice runs of the other tests show that a host
which does not keep the option still works through the unprotect and protect of each write window.
"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
UTIL = (ROOT / "src" / "modules" / "CORE_Util.bas").read_text(encoding="cp1252")
ACCEPTANCE = (ROOT / "tests" / "modules" / "TEST_Acceptance.bas").read_text(encoding="cp1252")


def body(text: str, start: str, end: str) -> str:
    return text.split(start, 1)[1].split(end, 1)[0]


def test_a_sheet_is_protected_for_the_interface_only_in_excel_and_the_old_way_elsewhere() -> None:
    sheet = body(UTIL, "Private Sub ProtectSheet", "End Sub")
    excel, other = sheet.split("ElseIf Left$(ws.Name, 3)", 1)
    assert excel.count("UserInterfaceOnly:=True") == 2 and "UserInterfaceOnly" not in other
    assert "If KeepsInterfaceOnly() Then" in excel


def test_the_window_protects_each_sheet_once_and_leaves_it_alone_afterwards() -> None:
    protection = body(UTIL, "Private Sub SetProtection", "End Sub")
    # A sheet already protected for the interface only is skipped; the others are opened, unlocked and protected that way
    assert "If Not (TargetSheet.ProtectContents And TargetSheet.ProtectionMode) Then" in protection
    assert protection.index("ProtectSheet TargetSheet") < protection.index("TargetSheet.ProtectionMode Then")
    # Closing also protects managed sheets created after the window opened.
    excel_branch = protection.split("If KeepsInterfaceOnly() Then", 1)[1].split("Else\n", 1)[0]
    assert "If Not turnOn Then" in excel_branch and "TargetSheet.Protect " not in excel_branch
    closing = excel_branch.split("ElseIf Not TargetSheet.ProtectContents Then", 1)[1].split("End If", 1)[0]
    assert closing.index("UnlockInputCells TargetSheet") < closing.index("ProtectSheet TargetSheet")
    assert "ThisWorkbook.Protect Structure:=True, Windows:=False" in protection


def test_a_host_that_does_not_keep_the_option_falls_back_to_the_old_way() -> None:
    protection = body(UTIL, "Private Sub SetProtection", "End Sub")
    assert "TargetSheet.Unprotect \"\"\n                                    mUiOnlyRefused = True" in protection
    host = body(UTIL, "Private Function KeepsInterfaceOnly", "End Function")
    assert "ThisWorkbook.Worksheets(1).ProtectionMode" in host and "KeepsInterfaceOnly = mHostKnowsMode And Not mUiOnlyRefused" in host
    # LibreOffice calls itself Excel, so the host is not told by its name
    assert "Application.Name" not in host.split("'------------------------------------------------------------------------------\n' DECLARE", 1)[1]


def test_the_acceptance_check_requires_the_interface_only_protection_and_a_script_that_writes() -> None:
    check = body(ACCEPTANCE, "Private Function CheckProtection", "End Function")
    assert "ManagedSheet.ProtectContents And ManagedSheet.ProtectionMode" in check
    assert "If WriteRefused(GetSheet(SH_OUT_SUMMARY).Cells(LAYOUT_HEADER_ROW, LAYOUT_FIRST_COL)) Then" in check
    assert "BeginWrite" in check and "EndWrite" in check
