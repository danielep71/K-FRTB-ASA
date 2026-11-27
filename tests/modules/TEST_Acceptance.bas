Attribute VB_Name = "TEST_Acceptance"

Option Explicit

'
'==============================================================================
' MODULE: TEST_Acceptance
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the checks of the Excel acceptance without a person at the keyboard
'   (#274): the sample run, the state after a failed run and after Clear
'   outputs (#85), numeric text with a comma and with a point, the rejection
'   of sub-types MAR21 does not define (#245, #276), the frozen panes, the
'   autofilters, the width of the Control descriptions (#203), the values
'   Excel displays, the absence of dates and the view of every sheet (#206,
'   #316), the data validation of the input and parameter sheets (#348), the
'   navigation bar and the Home sheet (#347), the key highlights (#350), the
'   protection of the sheets and the structure (#349), and the invalid-row
'   variant (#28).
'
' WHY THIS EXISTS
'   The Excel acceptance has to run without a person at the keyboard:
'   tools/Check-FRTB-Excel.ps1 runs it in desktop Excel, and the LibreOffice
'   tests run it through Test_RunAcceptance_Silent with the checks only Excel
'   can make reported as SKIPPED.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Runs the sixteen checks in a fixed order and collects one line per check
'   (check name, PASS, FAIL or SKIPPED, and a detail). Writes the lines to the
'   Acceptance sheet, which it adds when it is missing, with RESULT=PASS or
'   RESULT=FAIL in cell C2.
'
' ERROR POLICY
'   An error inside a check makes that check FAIL with the error's text; the
'   other checks still run. No dialog is shown.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, M_Main, M_Setup (ManagedSheets),
'   TEST_AcceptanceData.
'
' NOTES
'   Excel VBA; also runs in LibreOffice Calc with VBA support for the
'   automated tests through Test_RunAcceptance_Silent. LibreOffice reports
'   "Microsoft Excel" as its Application.Name, so the caller says which host
'   runs the checks. No references beyond the defaults.
'
'   Development build only: tools/Build-Workbook.ps1 -Production leaves this
'   module out.
'
' PUBLIC SURFACE
'   Test_RunAcceptanceInExcel_Silent: run every check and write one line per
'   check (PASS, FAIL or SKIPPED, with a detail) to the Acceptance sheet,
'   which it adds; cell C2 holds RESULT=PASS or RESULT=FAIL.
'   tools/Check-FRTB-Excel.ps1 runs it in desktop Excel.
'   Test_RunAcceptance_Silent: the same, with the checks only Excel can make
'   (the window panes, the autofilters, the column widths, the displayed
'   values, the sheet view, the data validation, the navigation bar, the key
'   highlights and the protection) SKIPPED; the LibreOffice tests run it.
'   Test_TimeNavigationToggle: times the navigation bar toggle and shows the times.
'
' STATE OWNERSHIP
'   Changes inputs, parameters, outputs and the window of the workbook it
'   runs in, and adds the Acceptance sheet: run it on a copy and close the
'   copy without saving. The variant runs last and leaves its rows behind.
'
' UPDATED
'   2026-11-20
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const SH_ACCEPTANCE         As String = "Acceptance"        'Name of the report sheet

    Private Const CLEARED_NOTE          As String = "No results yet: press Run FRTB-SA."    'Note an output sheet shows after Clear outputs
    Private Const FAILED_NOTE           As String = "Run failed (see the Log sheet): "      'Start of the alert Results shows after a failed run

    Private Const AMOUNT_TOLERANCE      As Double = 0.01    'Out_Summary amounts, in currency units
    Private Const TEXT_COMPARE          As Long = 1         'vbTextCompare
    Private Const VAR_DATE              As Long = 7         'vbDate, as VarType returns it
    Private Const MAX_LISTED            As Long = 10        'Cells named in a FAIL detail

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mReport                     As Collection   'Array(check, result, detail), one per check
    Private mInExcel                    As Boolean      'True: run the Excel-only checks


'
'------------------------------------------------------------------------------
'
'                                PUBLIC MACROS
'
'------------------------------------------------------------------------------
'

Public Sub Test_RunAcceptanceInExcel_Silent()

'
'------------------------------------------------------------------------------
'                         RUN THE ACCEPTANCE IN EXCEL
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every acceptance check, including the checks only Excel can make, and
'   writes the report to the Acceptance sheet.
'
' WHY THIS EXISTS
'   Desktop Excel needs a macro that runs the checks without a person at the
'   keyboard; tools/Check-FRTB-Excel.ps1 runs this one.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The report is on the Acceptance sheet: one line per check and
'   RESULT=PASS or RESULT=FAIL in cell C2.
'
' BEHAVIOR
'   Calls RunAll with True, so the checks that depend on Excel (panes,
'   autofilters, column widths, displayed values, sheet view, data validation,
'   navigation bar, key highlights, protection) run instead of reporting
'   SKIPPED.
'
' ERROR POLICY
'   None of its own. An error inside a check becomes a FAIL line (see
'   RunCheck); an error elsewhere in the run is not handled here.
'
' DEPENDENCIES
'   RunAll.
'
' NOTES
'   Changes the workbook it runs in: run it on a copy and close the copy
'   without saving.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN CHECKS
'------------------------------------------------------------------------------
    'Run the checks, with the Excel-only checks included
        RunAll True

End Sub

Public Sub Test_RunAcceptance_Silent()

'
'------------------------------------------------------------------------------
'                 RUN THE ACCEPTANCE WITHOUT EXCEL-ONLY CHECKS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the acceptance checks that do not depend on Excel and writes the report
'   to the Acceptance sheet.
'
' WHY THIS EXISTS
'   The LibreOffice tests run the acceptance through this macro; the checks that
'   need Excel's window, fonts or column widths report SKIPPED there.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The report is on the Acceptance sheet: one line per check and
'   RESULT=PASS or RESULT=FAIL in cell C2.
'
' BEHAVIOR
'   Calls RunAll with False, so the checks that read Excel's window, fonts,
'   column widths, validation rules, shapes, conditional formats or protection
'   report SKIPPED.
'
' ERROR POLICY
'   None of its own. An error inside a check becomes a FAIL line (see
'   RunCheck); an error elsewhere in the run is not handled here.
'
' DEPENDENCIES
'   RunAll.
'
' NOTES
'   Changes the workbook it runs in: run it on a copy and close the copy
'   without saving.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN CHECKS
'------------------------------------------------------------------------------
    'Run the checks, with the Excel-only checks skipped
        RunAll False

End Sub

Public Sub Test_TimeNavigationToggle()

'
'------------------------------------------------------------------------------
'                        TIME THE NAVIGATION BAR TOGGLE
'------------------------------------------------------------------------------
' PURPOSE
'   Times an empty write window twice and the two toggles of the navigation bar twice, and
'   reports the six times in a message.
'
' WHY THIS EXISTS
'   The first write window of a session protects the sheets for the interface only, which
'   takes a moment; later ones do not. Timing the same steps twice shows that first cost
'   apart from the cost of the shapes and the columns (#353).
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. A message gives the six times in seconds.
'
' BEHAVIOR
'   Opens and closes a write window twice and times each, then calls SetNavigationVisible with
'   False and with True and times each call, and does it once more. The bar is left shown.
'
' ERROR POLICY
'   None. An error stops the macro.
'
' DEPENDENCIES
'   BeginWrite, EndWrite (CORE_Util); SetNavigationVisible (M_Setup).
'
' NOTES
'   Excel only; the shapes and the columns are the ones of Excel. Run it on a workbook just
'   opened, so that the first window is the first of the session, and again to see the
'   steady state. It changes only the state of the bar, which it leaves shown.
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.Test_TimeNavigationToggle"

    Dim Started                 As Double       'Timer when a step started
    Dim FirstWindow             As Double       'Seconds of the first empty write window
    Dim SecondWindow            As Double       'Seconds of the second empty write window
    Dim HideFirst               As Double       'Seconds of the first toggle that hides the bar
    Dim ShowFirst               As Double       'Seconds of the first toggle that shows it
    Dim HideSecond              As Double       'Seconds of the second toggle that hides the bar
    Dim ShowSecond              As Double       'Seconds of the second toggle that shows it

'------------------------------------------------------------------------------
' TIME THE STEPS
'------------------------------------------------------------------------------
    'Time two empty write windows: the first of a session, then one with the sheets already protected
        Started = Timer
        BeginWrite
        EndWrite
        FirstWindow = Timer - Started
        Started = Timer
        BeginWrite
        EndWrite
        SecondWindow = Timer - Started
    'Time the toggle that hides the bar and the one that shows it, twice
        Started = Timer
        SetNavigationVisible False
        HideFirst = Timer - Started
        Started = Timer
        SetNavigationVisible True
        ShowFirst = Timer - Started
        Started = Timer
        SetNavigationVisible False
        HideSecond = Timer - Started
        Started = Timer
        SetNavigationVisible True
        ShowSecond = Timer - Started

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Report the six times
        MsgBox "Empty write window, first and second: " & Round(FirstWindow, 2) & " s, " & Round(SecondWindow, 2) & " s" & _
            vbCrLf & "Hide the bar, first and second: " & Round(HideFirst, 2) & " s, " & Round(HideSecond, 2) & " s" & _
            vbCrLf & "Show the bar, first and second: " & Round(ShowFirst, 2) & " s, " & Round(ShowSecond, 2) & " s", vbInformation

End Sub

'
'------------------------------------------------------------------------------
'
'                                    RUNNER
'
'------------------------------------------------------------------------------
'

Private Sub RunAll(ByVal inExcel As Boolean)

'
'------------------------------------------------------------------------------
'                                RUN ALL CHECKS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the sixteen checks in order and writes the report.
'
' WHY THIS EXISTS
'   The two public macros differ only in whether the Excel-only checks run; this
'   routine holds the shared list and order of the checks.
'
' INPUTS
'   inExcel
'     True to run the Excel-only checks; False to have them report
'     SKIPPED
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the host flag, starts an empty report, runs each check by name and
'   writes the report to the Acceptance sheet.
'
' ERROR POLICY
'   None of its own. RunCheck turns an error inside a check into a FAIL line.
'
' DEPENDENCIES
'   RunCheck, WriteReport, mInExcel, mReport.
'
' NOTES
'   INVALID_ROW_VARIANT runs last because it adds input rows.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' STORE HOST
'------------------------------------------------------------------------------
    'Remember which host runs the checks
        mInExcel = inExcel
    'Start an empty report
        Set mReport = New Collection

'------------------------------------------------------------------------------
' RUN CHECKS
'------------------------------------------------------------------------------
    'Run the checks in order; the variant is last because it adds input rows
        RunCheck "SAMPLE_RUN"
        RunCheck "FAILED_RUN"
        RunCheck "CLEAR_OUTPUTS"
        RunCheck "NUMERIC_TEXT"
        RunCheck "UNDEFINED_SUB_TYPES"
        RunCheck "FROZEN_PANES"
        RunCheck "AUTOFILTERS"
        RunCheck "CONTROL_DESCRIPTIONS"
        RunCheck "DISPLAYED_VALUES"
        RunCheck "NO_DATES"
        RunCheck "SHEET_VIEW"
        RunCheck "DATA_VALIDATION"
        RunCheck "NAVIGATION"
        RunCheck "KEY_HIGHLIGHTS"
        RunCheck "PROTECTION"
        RunCheck "INVALID_ROW_VARIANT"

'------------------------------------------------------------------------------
' WRITE REPORT
'------------------------------------------------------------------------------
    'Write the collected lines to the Acceptance sheet
        WriteReport

End Sub

Private Sub RunCheck(ByVal checkName As String)

'
'------------------------------------------------------------------------------
'                                RUN ONE CHECK
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the check that carries the given name and adds its line to the report.
'
' WHY THIS EXISTS
'   One place turns the outcome of a check, or an error inside it, into a report
'   line, so a failing check does not stop the others.
'
' INPUTS
'   checkName
'     Name of the check: SAMPLE_RUN, FAILED_RUN, CLEAR_OUTPUTS,
'     NUMERIC_TEXT, UNDEFINED_SUB_TYPES, FROZEN_PANES, AUTOFILTERS,
'     CONTROL_DESCRIPTIONS, DISPLAYED_VALUES, NO_DATES, SHEET_VIEW,
'     DATA_VALIDATION, NAVIGATION, KEY_HIGHLIGHTS, PROTECTION or
'     INVALID_ROW_VARIANT
'
' RETURNS
'   Nothing. The line Array(check, result, detail) is added to mReport.
'
' BEHAVIOR
'   Calls the Check function that matches the name, which returns the result
'   and fills the detail, and adds the line to the report. A name that matches
'   no check adds a line with an empty result and detail.
'
' ERROR POLICY
'   An error inside the check is caught: CloseWriteWindows runs, and the line is
'   FAIL with the text "Error <number>: <description>". Nothing is raised.
'
' DEPENDENCIES
'   The sixteen Check functions, CloseWriteWindows (CORE_Util), mReport.
'
' NOTES
'   CloseWriteWindows is called because a check may fail inside a write window.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.RunCheck"

    Dim CheckDetail             As String       'Detail text the check returns
    Dim CheckResult             As String       'PASS, FAIL or SKIPPED the check returns

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Report an error inside a check as a FAIL line
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN CHECK
'------------------------------------------------------------------------------
    'Run the check that carries this name
        Select Case checkName
            Case "SAMPLE_RUN"
                CheckResult = CheckSampleRun(CheckDetail)
            Case "FAILED_RUN"
                CheckResult = CheckFailedRun(CheckDetail)
            Case "CLEAR_OUTPUTS"
                CheckResult = CheckClearOutputs(CheckDetail)
            Case "NUMERIC_TEXT"
                CheckResult = CheckNumericText(CheckDetail)
            Case "UNDEFINED_SUB_TYPES"
                CheckResult = CheckSubTypes(CheckDetail)
            Case "FROZEN_PANES"
                CheckResult = CheckFrozenPanes(CheckDetail)
            Case "AUTOFILTERS"
                CheckResult = CheckAutofilters(CheckDetail)
            Case "CONTROL_DESCRIPTIONS"
                CheckResult = CheckControlDescriptions(CheckDetail)
            Case "DISPLAYED_VALUES"
                CheckResult = CheckDisplayedValues(CheckDetail)
            Case "NO_DATES"
                CheckResult = CheckNoDates(CheckDetail)
            Case "SHEET_VIEW"
                CheckResult = CheckSheetView(CheckDetail)
            Case "DATA_VALIDATION"
                CheckResult = CheckDataValidation(CheckDetail)
            Case "NAVIGATION"
                CheckResult = CheckNavigation(CheckDetail)
            Case "KEY_HIGHLIGHTS"
                CheckResult = CheckKeyHighlights(CheckDetail)
            Case "PROTECTION"
                CheckResult = CheckProtection(CheckDetail)
            Case "INVALID_ROW_VARIANT"
                CheckResult = CheckVariant(CheckDetail)
            Case Else
                'No other name is run; nothing is done
        End Select

'------------------------------------------------------------------------------
' RECORD RESULT
'------------------------------------------------------------------------------
    'Add the check, its result and its detail to the report
        mReport.Add Array(checkName, CheckResult, CheckDetail)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Close a write window the failed check left open
        CloseWriteWindows
    'Add a FAIL line with the error number and description
        mReport.Add Array(checkName, "FAIL", "Error " & Err.Number & ": " & Err.Description)

End Sub

Private Sub WriteReport()

'
'------------------------------------------------------------------------------
'                               WRITE THE REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the collected check lines to the Acceptance sheet and the overall
'   verdict to cell C2.
'
' WHY THIS EXISTS
'   The checks write their lines to a collection; the Check-FRTB-Excel tooling
'   and the LibreOffice tests read the lines and the verdict from the sheet.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens a write window (BeginWrite), because a new sheet needs the structure
'   unprotected (#349). Clears the Acceptance sheet when it exists, otherwise
'   adds it after the last sheet. Writes the headers Check, Result and Detail in
'   row 4 (columns C to E) and one row per check below them. Writes RESULT=FAIL
'   in C2 when any line is FAIL, otherwise RESULT=PASS. Closes the write window
'   (EndWrite).
'
' ERROR POLICY
'   None. Errors are not handled here.
'
' DEPENDENCIES
'   BeginWrite, EndWrite, SheetExists, GetSheet (CORE_Util); mReport,
'   SH_ACCEPTANCE, ThisWorkbook.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.WriteReport"

    Dim ReportSheet             As Object       'Acceptance sheet
    Dim ReportEntry             As Variant      'Line of the report: check, result, detail
    Dim ReportRow               As Long         'Row being written
    Dim AnyFailed               As Boolean      'True when a line is FAIL

'------------------------------------------------------------------------------
' OPEN WRITE WINDOW
'------------------------------------------------------------------------------
    'Unprotect the structure, which a new sheet needs (#349)
        BeginWrite

'------------------------------------------------------------------------------
' PREPARE SHEET
'------------------------------------------------------------------------------
    'Clear the report sheet, or add it after the last sheet
        If SheetExists(SH_ACCEPTANCE) Then
            Set ReportSheet = GetSheet(SH_ACCEPTANCE)
            ReportSheet.Cells.Clear
        Else
            Set ReportSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ReportSheet.Name = SH_ACCEPTANCE
        End If

'------------------------------------------------------------------------------
' WRITE LINES
'------------------------------------------------------------------------------
    'Write the headers in row 4
        ReportRow = 4
        ReportSheet.Cells(ReportRow, 3).Value = "Check"
        ReportSheet.Cells(ReportRow, 4).Value = "Result"
        ReportSheet.Cells(ReportRow, 5).Value = "Detail"
    'Write one row per check and note whether any is FAIL
        For Each ReportEntry In mReport
            ReportRow = ReportRow + 1
            ReportSheet.Cells(ReportRow, 3).Value = ReportEntry(0)
            ReportSheet.Cells(ReportRow, 4).Value = ReportEntry(1)
            ReportSheet.Cells(ReportRow, 5).Value = ReportEntry(2)
            If ReportEntry(1) = "FAIL" Then AnyFailed = True
        Next ReportEntry

'------------------------------------------------------------------------------
' WRITE VERDICT
'------------------------------------------------------------------------------
    'Write the overall result in C2
        ReportSheet.Cells(2, 3).Value = IIf(AnyFailed, "RESULT=FAIL", "RESULT=PASS")

'------------------------------------------------------------------------------
' CLOSE WRITE WINDOW
'------------------------------------------------------------------------------
    'Protect the structure again
        EndWrite

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CHECKS
'
'------------------------------------------------------------------------------
'

Private Function CheckSampleRun(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                             CHECK THE SAMPLE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the sample portfolio and checks that the run ends without error and that
'   the results are the reference results.
'
' INPUTS
'   detail
'     Receives the differences found, or on a pass the total and the
'     counts of errors and warnings
'
' RETURNS
'   PASS when CompareSummary finds no difference, otherwise FAIL.
'
' ERROR POLICY
'   None. An error raised by the run is caught by RunCheck.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN AND COMPARE
'------------------------------------------------------------------------------
    'Run the calculation
        RunFRTB_Silent
    'Compare the results with the reference; no difference means PASS
        detail = CompareSummary(AcceptanceSampleExpected(), 0, 0)
        CheckSampleRun = IIf(detail = vbNullString, "PASS", "FAIL")

'------------------------------------------------------------------------------
' DESCRIBE PASS
'------------------------------------------------------------------------------
    'On a pass, describe the total and the counts
        If detail = vbNullString Then
            detail = "total " & Format(OutSummaryValue("TOTAL"), "0.00") & ", 0 errors, 0 warnings"
        End If

End Function

Private Function CheckFailedRun(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                      CHECK THE STATE AFTER A FAILED RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Blanks the GIRR risk weight, runs, and checks the state a failed run leaves:
'   Results shows the alert, the Out_* sheets their placeholder, and the Log the
'   error (#85).
'
' WHY THIS EXISTS
'   A blank GIRR risk weight stops the run; the sheets must then show that no
'   results exist and why.
'
' INPUTS
'   detail
'     Receives the problems found, or on a pass the first line of
'     Results (C6)
'
' RETURNS
'   PASS when there is no problem, otherwise FAIL.
'
' BEHAVIOR
'   Clears the cell two rows below the GIRR_RW tag of the GIRR parameter
'   sheet, runs, and records a problem when:
'     - the run did not fail
'     - the error does not name GIRR_RW
'     - Results, Out_Summary, Out_Class, Out_Buckets or Out_DRC do not show
'       their placeholder
'     - the Log has no ERROR row naming GIRR_RW
'   Then writes the saved value back.
'
' ERROR POLICY
'   An error in the check restores the saved value and is raised again with the
'   same number and description.
'
' DEPENDENCIES
'   RunFRTB_Silent, LastRunStatus, PlaceholderProblems, LogHas, FindTagRow,
'   GetSheet, TableCell.
'
' NOTES
'   The saved value is restored on success and on error.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckFailedRun"

    Dim GirrSheet               As Object       'GIRR parameter sheet
    Dim TargetCell              As Object       'Cell two rows below the GIRR_RW tag
    Dim SavedValue              As Variant      'Value of the cell before it is cleared
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Find the GIRR parameter sheet
        Set GirrSheet = GetSheet(SH_P_GIRR)
    'Point at the cell two rows below the GIRR_RW tag
        Set TargetCell = GirrSheet.Cells(FindTagRow(GirrSheet, "GIRR_RW") + 2, LAYOUT_FIRST_COL + 1)
    'Keep its value to write it back
        SavedValue = TargetCell.Value
    'Blank the risk weight
        TargetCell.ClearContents
    'Write the value back and raise again when an error occurs from here on
        On Error GoTo Restore

'------------------------------------------------------------------------------
' RUN WITH A BLANK WEIGHT
'------------------------------------------------------------------------------
    'Run the calculation
        RunFRTB_Silent
    'Require that the run failed
        If LastRunStatus() = "OK" Then Problems = Problems & "the run did not fail; "
    'Require that the error names GIRR_RW
        If InStr(1, LastRunStatus(), "GIRR_RW", TEXT_COMPARE) = 0 Then
            Problems = Problems & "the error does not name GIRR_RW; "
        End If

'------------------------------------------------------------------------------
' CHECK PLACEHOLDERS
'------------------------------------------------------------------------------
    'Require the alert on Results and the cleared note on the Out_* sheets
        Problems = Problems & PlaceholderProblems(SH_RESULTS, FAILED_NOTE, True)
        Problems = Problems & PlaceholderProblems(SH_OUT_SUMMARY, CLEARED_NOTE, False)
        Problems = Problems & PlaceholderProblems(SH_OUT_CLASS, CLEARED_NOTE, False)
        Problems = Problems & PlaceholderProblems(SH_OUT_BUCKETS, CLEARED_NOTE, False)
        Problems = Problems & PlaceholderProblems(SH_OUT_DRC, CLEARED_NOTE, False)
    'Require an ERROR row in the Log that names GIRR_RW
        If Not LogHas("ERROR", "", "GIRR_RW") Then
            Problems = Problems & "the Log has no GIRR_RW error; "
        End If

'------------------------------------------------------------------------------
' RESTORE VALUE
'------------------------------------------------------------------------------
    'Write the saved value back and restore normal error handling
        TargetCell.Value = SavedValue
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the first line of Results on a pass
        detail = IIf(Problems = vbNullString, _
            "Results C6: " & CStr(TableCell(GetSheet(SH_RESULTS), LAYOUT_FIRST_DATA_ROW, 1).Value), _
            Problems)
        CheckFailedRun = IIf(Problems = vbNullString, "PASS", "FAIL")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Restore:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Write the saved value back
        TargetCell.Value = SavedValue
    'Raise the error again
        Err.Raise ErrorNumber, "TEST_Acceptance", ErrorDescription

End Function

Private Function CheckClearOutputs(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                     CHECK THE STATE AFTER CLEAR OUTPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the sample, calls Clear outputs and checks that the sheets are back to
'   the template's empty state (#85).
'
' WHY THIS EXISTS
'   Clear outputs after a successful run must leave the template's empty state.
'
' INPUTS
'   detail
'     Receives the problems found, or on a pass the description of the
'     empty state
'
' RETURNS
'   PASS when there is no problem, otherwise FAIL.
'
' BEHAVIOR
'   Runs the calculation and records a problem when it did not end with OK. Calls
'   ClearOutputs and checks the placeholder of Results, Out_Summary, Out_Class,
'   Out_Buckets, Out_DRC and Log, and that column A of each is still the black of the
'   brand bar below the banner, where the navigation bar sits.
'
' ERROR POLICY
'   None. An error raised by the run or by Clear outputs is caught by RunCheck.
'
' DEPENDENCIES
'   RunFRTB_Silent, LastRunStatus, ClearOutputs, PlaceholderProblems, GetSheet,
'   LAYOUT_FIRST_DATA_ROW, COLOR_BRAND_BAR.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckClearOutputs"

    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim SheetName               As Variant      'Output sheet whose column A is read

'------------------------------------------------------------------------------
' RUN AND CLEAR
'------------------------------------------------------------------------------
    'Run the calculation and record a problem when it did not end with OK
        RunFRTB_Silent
        If LastRunStatus() <> "OK" Then
            Problems = "the run before Clear outputs failed: " & LastRunStatus() & "; "
        End If
    'Clear the outputs
        ClearOutputs

'------------------------------------------------------------------------------
' CHECK PLACEHOLDERS
'------------------------------------------------------------------------------
    'Require the cleared note on each output sheet and on the Log
        Problems = Problems & PlaceholderProblems(SH_RESULTS, CLEARED_NOTE, False)
        Problems = Problems & PlaceholderProblems(SH_OUT_SUMMARY, CLEARED_NOTE, False)
        Problems = Problems & PlaceholderProblems(SH_OUT_CLASS, CLEARED_NOTE, False)
        Problems = Problems & PlaceholderProblems(SH_OUT_BUCKETS, CLEARED_NOTE, False)
        Problems = Problems & PlaceholderProblems(SH_OUT_DRC, CLEARED_NOTE, False)
        Problems = Problems & PlaceholderProblems(SH_LOG, CLEARED_NOTE, False)

'------------------------------------------------------------------------------
' CHECK COLUMN A
'------------------------------------------------------------------------------
    'Require column A to keep the black of the brand bar below the banner on each of them
        For Each SheetName In Array(SH_RESULTS, SH_OUT_SUMMARY, SH_OUT_CLASS, SH_OUT_BUCKETS, SH_OUT_DRC, SH_LOG)
            If GetSheet(CStr(SheetName)).Cells(LAYOUT_FIRST_DATA_ROW + 1, 1).Interior.Color <> COLOR_BRAND_BAR Then
                Problems = Problems & CStr(SheetName) & " lost the black of column A; "
            End If
        Next SheetName

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the description of the empty state on a pass
        detail = IIf(Problems = vbNullString, _
            "headers in row 5 and the note in row 6 on Results, the Out_* sheets and Log; column A stays black", Problems)
        CheckClearOutputs = IIf(Problems = vbNullString, "PASS", "FAIL")

End Function

Private Function CheckNumericText(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                              CHECK NUMERIC TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Runs CheckNumericTextInWindow inside a write window and raises again any
'   error it produced.
'
' WHY THIS EXISTS
'   Numeric text has one, locale-independent form: a point for the decimals. A
'   comma is rejected whatever the regional settings; a typed number is not text
'   and is accepted (docs/workbook-guide.md). A cell format is locked on a
'   protected sheet, so the check needs the write window (#349).
'
' INPUTS
'   detail
'     Receives the problems found, or on a pass the forms rejected and
'     accepted
'
' RETURNS
'   PASS when there is no problem, otherwise FAIL.
'
' BEHAVIOR
'   Opens the write window (BeginWrite), calls CheckNumericTextInWindow with
'   errors suppressed (On Error Resume Next) and keeps its result and the number,
'   source and description of any error. Restores normal error handling, closes
'   the write window (EndWrite) and raises the error again when there was one.
'
' ERROR POLICY
'   An error in CheckNumericTextInWindow is held back until the write window is
'   closed, then raised again with the same number, source and description.
'
' DEPENDENCIES
'   CheckNumericTextInWindow, BeginWrite, EndWrite.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckNumericText"

    Dim CheckResult             As String       'Result the inner check returns
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorSource             As String       'Captured error source
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' OPEN WRITE WINDOW
'------------------------------------------------------------------------------
    'Unprotect what the check needs: a cell format is locked on a protected sheet (#349)
        BeginWrite

'------------------------------------------------------------------------------
' RUN THE CHECK
'------------------------------------------------------------------------------
    'Run the check and keep any error, so the write window is closed first
        On Error Resume Next
        CheckResult = CheckNumericTextInWindow(detail)
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' CLOSE WRITE WINDOW
'------------------------------------------------------------------------------
    'Protect the sheets again
        EndWrite

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Raise the error again when the check failed
        If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription
    'Return the result of the check
        CheckNumericText = CheckResult

End Function

Private Function CheckNumericTextInWindow(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                    CHECK NUMERIC TEXT IN THE WRITE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Writes numeric text in several forms into the first row of In_Delta and
'   checks which forms the run rejects.
'
' WHY THIS EXISTS
'   CheckNumericText needs the write window open; this routine holds the check
'   itself, so that the window is closed whatever happens inside it.
'
' INPUTS
'   detail
'     Receives the problems found, or on a pass the forms rejected and
'     accepted
'
' RETURNS
'   PASS when there is no problem, otherwise FAIL.
'
' BEHAVIOR
'   For each text, formats the Sensitivity cell of the first data row as text,
'   writes the text, runs, and checks whether the Log holds the error
'   "Sensitivity is not numeric" for that row:
'     - '0,5', '1.000,5' and '1,5' must be rejected
'     - '0.5' and '-1.25e2' must be accepted
'   Then writes the number 0.5 and checks that the Log has no ERROR for that
'   row. Writes the original format and value back.
'
' ERROR POLICY
'   An error in the check writes the original format and value back and is
'   raised again with the same number and description.
'
' DEPENDENCIES
'   RunFRTB_Silent, LogHas, HeaderColumn, TableCell, GetSheet.
'
' NOTES
'   The original format and value are restored on success and on error.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckNumericTextInWindow"

    Dim InputSheet              As Object       'In_Delta sheet
    Dim TargetCell              As Object       'Sensitivity cell of the first data row
    Dim SavedValue              As Variant      'Value of the cell before the check
    Dim SavedFormat             As String       'Number format of the cell before the check
    Dim TextCases               As Variant      'Pairs: text, then True when the row must be rejected
    Dim CaseIndex               As Long         'Position in the pairs
    Dim WasRejected             As Boolean      'True when the Log holds the error for the row
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim RowSource               As String       'Log source of the row
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Find the In_Delta sheet and the Sensitivity cell of the first data row
        Set InputSheet = GetSheet(SH_IN_DELTA)
        Set TargetCell = TableCell(InputSheet, LAYOUT_FIRST_DATA_ROW, _
            HeaderColumn(InputSheet, "Sensitivity", LAYOUT_HEADER_ROW))
    'Keep its value and format to write them back
        SavedValue = TargetCell.Value
        SavedFormat = TargetCell.NumberFormat
    'Build the Log source of the row
        RowSource = SH_IN_DELTA & " row " & LAYOUT_FIRST_DATA_ROW
    'List the texts, each followed by True when the row must be rejected
        TextCases = Array("0,5", True, "1.000,5", True, "1,5", True, "0.5", False, "-1.25e2", False)
    'Write the original format and value back and raise again when an error occurs from here on
        On Error GoTo Restore

'------------------------------------------------------------------------------
' CHECK TEXT FORMS
'------------------------------------------------------------------------------
    'Write each text, run, and compare the rejection with the expectation
        For CaseIndex = LBound(TextCases) To UBound(TextCases) Step 2
            TargetCell.NumberFormat = "@"
            TargetCell.Value = CStr(TextCases(CaseIndex))
            RunFRTB_Silent
            WasRejected = LogHas("ERROR", RowSource, "Sensitivity is not numeric")
            If WasRejected <> CBool(TextCases(CaseIndex + 1)) Then
                Problems = Problems & "text '" & TextCases(CaseIndex) & "' was " & _
                    IIf(WasRejected, "rejected", "accepted") & "; "
            End If
        Next CaseIndex

'------------------------------------------------------------------------------
' CHECK A TYPED NUMBER
'------------------------------------------------------------------------------
    'Write the number 0.5 in the original format and run
        TargetCell.NumberFormat = SavedFormat
        TargetCell.Value = 0.5
        RunFRTB_Silent
    'Require that the Log holds no ERROR for the row
        If LogHas("ERROR", RowSource, "") Then Problems = Problems & "the typed number 0.5 was rejected; "

'------------------------------------------------------------------------------
' RESTORE VALUE
'------------------------------------------------------------------------------
    'Write the original format and value back and restore normal error handling
        TargetCell.NumberFormat = SavedFormat
        TargetCell.Value = SavedValue
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the forms rejected and accepted on a pass
        detail = IIf(Problems = vbNullString, _
            "rejected as text: '0,5', '1.000,5', '1,5'; accepted: '0.5', '-1.25e2' and the number 0.5", _
            Problems)
        CheckNumericTextInWindow = IIf(Problems = vbNullString, "PASS", "FAIL")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Restore:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Write the original format and value back
        TargetCell.NumberFormat = SavedFormat
        TargetCell.Value = SavedValue
    'Raise the error again
        Err.Raise ErrorNumber, "TEST_Acceptance", ErrorDescription

End Function

Private Function CheckSubTypes(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                        CHECK THE UNDEFINED SUB-TYPES
'------------------------------------------------------------------------------
' PURPOSE
'   Adds rows whose sub-type MAR21 does not define, runs, and checks that each is
'   rejected with its message while the results stay those of the sample
'   (#245).
'
' WHY THIS EXISTS
'   Undefined sub-types must be rejected row by row without changing the results.
'   The rows are removed afterwards, so the invalid-row variant is unaffected.
'
' INPUTS
'   detail
'     Receives the problems found, or on a pass the number of rows
'     rejected
'
' RETURNS
'   PASS when there is no problem, otherwise FAIL.
'
' BEHAVIOR
'   Appends each row of AcceptanceSubTypeRows below the last data row of its
'   sheet, runs, and compares Out_Summary with the sample's results, with the
'   expected number of errors and no warnings. Then checks that each expected
'   message is logged as many times as it is expected. Clears the rows it
'   wrote.
'
' ERROR POLICY
'   An error in the check clears the rows it wrote and is raised again with the
'   same number and description.
'
' DEPENDENCIES
'   RunFRTB_Silent, CompareSummary, MessageProblems, RemoveRows, LastDataRow,
'   TableCell, GetSheet, AcceptanceSubTypeRows, AcceptanceSubTypeErrors,
'   AcceptanceSampleExpected.
'
' NOTES
'   Element 0 of each row definition is the sheet name; the elements from index
'   2 on are written to table columns 1 onward, empty ones skipped.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckSubTypes"

    Dim SubTypeRow              As Variant          'Row definition: sheet name, then the values
    Dim TargetSheet             As Object           'Sheet that receives the row
    Dim TargetRow               As Long             'Row being written
    Dim ColumnIndex             As Long             'Element of the row definition being written
    Dim WrittenRows             As New Collection   'Array(sheet name, row, columns) of each row written
    Dim Problems                As String           'Problems found, each ending with a semicolon
    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Clear the rows written and raise again when an error occurs from here on
        On Error GoTo Restore

'------------------------------------------------------------------------------
' WRITE ROWS
'------------------------------------------------------------------------------
    'Append each row below the last data row of its sheet and note where it went
        For Each SubTypeRow In AcceptanceSubTypeRows()
            Set TargetSheet = GetSheet(CStr(SubTypeRow(0)))
            TargetRow = LastDataRow(TargetSheet) + 1
            For ColumnIndex = 2 To UBound(SubTypeRow)
                If Not IsEmpty(SubTypeRow(ColumnIndex)) Then
                    TableCell(TargetSheet, TargetRow, ColumnIndex - 1).Value = SubTypeRow(ColumnIndex)
                End If
            Next ColumnIndex
            WrittenRows.Add Array(TargetSheet.Name, TargetRow, UBound(SubTypeRow) - 1)
        Next SubTypeRow

'------------------------------------------------------------------------------
' RUN AND COMPARE
'------------------------------------------------------------------------------
    'Run the calculation
        RunFRTB_Silent
    'Compare the results with the sample's and the error count with the expected messages
        Problems = CompareSummary(AcceptanceSampleExpected(), AcceptanceSubTypeErrors().Count, 0)
    'Require each expected message in the Log as many times as expected
        Problems = Problems & MessageProblems(AcceptanceSubTypeErrors())

'------------------------------------------------------------------------------
' REMOVE ROWS
'------------------------------------------------------------------------------
    'Clear the rows written and restore normal error handling
        RemoveRows WrittenRows
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the number of rows rejected on a pass
        detail = IIf(Problems = vbNullString, _
            WrittenRows.Count & " rows rejected with their messages; the results are the sample's", _
            Problems)
        CheckSubTypes = IIf(Problems = vbNullString, "PASS", "FAIL")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Restore:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Clear the rows written
        RemoveRows WrittenRows
    'Raise the error again
        Err.Raise ErrorNumber, "TEST_Acceptance", ErrorDescription

End Function

Private Function CheckFrozenPanes(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                            CHECK THE FROZEN PANES
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that every sheet listed by AcceptancePanes keeps the frozen panes of
'   the committed template (#203).
'
' WHY THIS EXISTS
'   The panes depend on the Excel window, so the check runs only in Excel.
'
' INPUTS
'   detail
'     Receives the problems found, the number of sheets frozen as in the
'     template, or the reason the check is skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when there is no problem,
'   otherwise FAIL.
'
' BEHAVIOR
'   For each sheet of AcceptancePanes, activates it and compares the freeze
'   state, the split row and the split column of the active window with the
'   expected values.
'
' ERROR POLICY
'   None. An error raised by the window calls is caught by RunCheck.
'
' DEPENDENCIES
'   AcceptancePanes, GetSheet, Application.ActiveWindow, mInExcel.
'
' NOTES
'   Activates each sheet it checks.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckFrozenPanes"

    Dim PaneSpec                As Variant      'Array(sheet name, split row, split column)
    Dim PaneSheet               As Object       'Sheet being checked
    Dim PaneWindow              As Object       'Active window
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim SheetCount              As Long         'Number of sheets checked

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: the window panes as Excel shows them"
            CheckFrozenPanes = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' CHECK PANES
'------------------------------------------------------------------------------
    'Activate each sheet and compare its panes with the expected ones
        For Each PaneSpec In AcceptancePanes()
            Set PaneSheet = GetSheet(CStr(PaneSpec(0)))
            PaneSheet.Activate
            Set PaneWindow = Application.ActiveWindow
            If Not PaneWindow.FreezePanes Or PaneWindow.SplitRow <> PaneSpec(1) Or _
                PaneWindow.SplitColumn <> PaneSpec(2) Then
                Problems = Problems & PaneSpec(0) & " frozen at " & PaneWindow.SplitRow & " rows and " & _
                    PaneWindow.SplitColumn & " columns, not " & PaneSpec(1) & " and " & PaneSpec(2) & "; "
            End If
            SheetCount = SheetCount + 1
        Next PaneSpec

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the number of sheets on a pass
        detail = IIf(Problems = vbNullString, SheetCount & " sheets frozen as in the template", Problems)
        CheckFrozenPanes = IIf(Problems = vbNullString, "PASS", "FAIL")

End Function

Private Function CheckAutofilters(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                            CHECK THE AUTOFILTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that every input sheet has its autofilter and that filtering In_Delta
'   on GIRR shows exactly the GIRR rows (#28).
'
' WHY THIS EXISTS
'   The autofilters depend on Excel, so the check runs only in Excel.
'
' INPUTS
'   detail
'     Receives the problems found, the description of the filtered view,
'     or the reason the check is skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when there is no problem,
'   otherwise FAIL.
'
' BEHAVIOR
'   For each sheet of AcceptanceFilters, requires an autofilter whose first row
'   has the expected address. Counts the GIRR rows of In_Delta, then opens a
'   write window (BeginWrite), because a script filters the protected sheet,
'   filters column 2 on GIRR, counts the visible GIRR rows and the visible other
'   rows, shows all data again and closes the write window (EndWrite). A problem
'   is recorded when there are no GIRR rows, when a GIRR row is hidden or when
'   another row is visible.
'
' ERROR POLICY
'   None. An error raised by the filter calls is caught by RunCheck.
'
' DEPENDENCIES
'   AcceptanceFilters, GetSheet, LastDataRow, HeaderCount, TableCell,
'   BeginWrite, EndWrite, mInExcel.
'
' NOTES
'   Leaves In_Delta unfiltered when it was filtered by this check. A person's
'   filter is not run here.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckAutofilters"

    Dim FilterSpec              As Variant      'Array(sheet name, expected address of the first row)
    Dim FilterSheet             As Object       'Sheet being checked
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim RowIndex                As Long         'Row being read
    Dim LastRow                 As Long         'Last data row of In_Delta
    Dim GirrRowCount            As Long         'GIRR rows of In_Delta
    Dim VisibleGirrCount        As Long         'GIRR rows shown after filtering
    Dim VisibleOtherCount       As Long         'Other rows shown after filtering

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: the autofilters as Excel applies them"
            CheckAutofilters = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' CHECK SHEET FILTERS
'------------------------------------------------------------------------------
    'Require an autofilter on the expected range of each input sheet
        For Each FilterSpec In AcceptanceFilters()
            Set FilterSheet = GetSheet(CStr(FilterSpec(0)))
            If Not FilterSheet.AutoFilterMode Then
                Problems = Problems & FilterSpec(0) & " has no autofilter; "
            ElseIf FilterSheet.AutoFilter.Range.Rows(1).Address(False, False) <> FilterSpec(1) Then
                Problems = Problems & FilterSpec(0) & " filters " & _
                    FilterSheet.AutoFilter.Range.Rows(1).Address(False, False) & _
                    ", not " & FilterSpec(1) & "; "
            End If
        Next FilterSpec

'------------------------------------------------------------------------------
' COUNT GIRR ROWS
'------------------------------------------------------------------------------
    'Find the last data row of In_Delta
        Set FilterSheet = GetSheet(SH_IN_DELTA)
        LastRow = LastDataRow(FilterSheet)
    'Count the rows whose second column is GIRR
        For RowIndex = LAYOUT_FIRST_DATA_ROW To LastRow
            If CStr(TableCell(FilterSheet, RowIndex, 2).Value) = "GIRR" Then
                GirrRowCount = GirrRowCount + 1
            End If
        Next RowIndex

'------------------------------------------------------------------------------
' FILTER ON GIRR
'------------------------------------------------------------------------------
    'Open the write window: a script filters the protected sheet, a person's filter is not run here
        BeginWrite
    'Filter the second column on GIRR
        FilterSheet.Range(TableCell(FilterSheet, LAYOUT_HEADER_ROW, 1), _
            TableCell(FilterSheet, LastRow, HeaderCount(FilterSheet))).AutoFilter 2, "GIRR"
    'Count the GIRR rows and the other rows that stay visible
        For RowIndex = LAYOUT_FIRST_DATA_ROW To LastRow
            If Not FilterSheet.Rows(RowIndex).Hidden Then
                If CStr(TableCell(FilterSheet, RowIndex, 2).Value) = "GIRR" Then
                    VisibleGirrCount = VisibleGirrCount + 1
                Else
                    VisibleOtherCount = VisibleOtherCount + 1
                End If
            End If
        Next RowIndex
    'Show all data again
        If FilterSheet.FilterMode Then FilterSheet.ShowAllData
    'Close the write window
        EndWrite
    'Require every GIRR row visible and no other row
        If GirrRowCount = 0 Or VisibleGirrCount <> GirrRowCount Or VisibleOtherCount <> 0 Then
            Problems = Problems & "the GIRR filter on In_Delta showed " & VisibleGirrCount & " of " & _
                GirrRowCount & " GIRR rows and " & VisibleOtherCount & " other rows; "
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the description of the filtered view on a pass
        detail = IIf(Problems = vbNullString, _
            "5 input sheets filtered as in the template; In_Delta filtered on GIRR shows its " & _
            GirrRowCount & " GIRR rows only", Problems)
        CheckAutofilters = IIf(Problems = vbNullString, "PASS", "FAIL")

End Function

Private Function CheckControlDescriptions(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                        CHECK THE CONTROL DESCRIPTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that every description on Control fits its column: Excel's own
'   AutoFit of those cells does not widen the column (#203).
'
' WHY THIS EXISTS
'   The column width depends on Excel's fonts, so the check runs only in Excel.
'
' INPUTS
'   detail
'     Receives the saved and fitted widths, or the reason the check is
'     skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when the fitted width is at most
'   0.5 above the saved width, otherwise FAIL.
'
' BEHAVIOR
'   Finds the Description column of Control and saves its width. Opens a write
'   window (BeginWrite), because a column width is locked on a protected sheet,
'   applies AutoFit to the description cells, reads the fitted width, writes the
'   saved width back and closes the write window (EndWrite).
'
' ERROR POLICY
'   None. An error raised by the sheet calls is caught by RunCheck.
'
' DEPENDENCIES
'   GetSheet, HeaderColumn, LastDataRow, BeginWrite, EndWrite, mInExcel.
'
' NOTES
'   The saved width is not restored when an error interrupts the check.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckControlDescriptions"

    Dim ControlSheet            As Object       'Control sheet
    Dim DescriptionColumn       As Long         'Sheet column of the Description header
    Dim FirstRow                As Long         'First description row
    Dim LastRow                 As Long         'Last description row
    Dim SavedWidth              As Double       'Column width before AutoFit
    Dim FittedWidth             As Double       'Column width after AutoFit

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: the column width depends on Excel's fonts"
            CheckControlDescriptions = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' LOCATE DESCRIPTIONS
'------------------------------------------------------------------------------
    'Find the Description column and the rows of the descriptions
        Set ControlSheet = GetSheet(SH_CONTROL)
        DescriptionColumn = LAYOUT_FIRST_COL + _
            HeaderColumn(ControlSheet, "Description", LAYOUT_HEADER_ROW + 1) - 1
        FirstRow = LAYOUT_HEADER_ROW + 2
        LastRow = LastDataRow(ControlSheet)
    'Keep the column width to write it back
        SavedWidth = ControlSheet.Columns(DescriptionColumn).ColumnWidth

'------------------------------------------------------------------------------
' FIT COLUMN
'------------------------------------------------------------------------------
    'Open the write window: a column width is locked on a protected sheet
        BeginWrite
    'Apply AutoFit to the description cells and read the fitted width
        ControlSheet.Range(ControlSheet.Cells(FirstRow, DescriptionColumn), _
            ControlSheet.Cells(LastRow, DescriptionColumn)).Columns.AutoFit
        FittedWidth = ControlSheet.Columns(DescriptionColumn).ColumnWidth
    'Write the saved width back
        ControlSheet.Columns(DescriptionColumn).ColumnWidth = SavedWidth
    'Close the write window
        EndWrite

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Describe the widths and pass when the texts fit within half a unit
        detail = "Description column " & Format(SavedWidth, "0.00") & " wide; its texts need " & _
            Format(FittedWidth, "0.00")
        CheckControlDescriptions = IIf(FittedWidth <= SavedWidth + 0.5, "PASS", "FAIL")

End Function

Private Function CheckDisplayedValues(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                          CHECK THE DISPLAYED VALUES
'------------------------------------------------------------------------------
' PURPOSE
'   After a sample run, checks that no cell of the managed sheets shows "####":
'   every column is wide enough for the values Excel displays (#206).
'
' WHY THIS EXISTS
'   The displayed text depends on Excel's fonts and column widths, so the check
'   runs only in Excel.
'
' INPUTS
'   detail
'     Receives the number of cells and sheets read, the cells that show
'     ####, or the reason the check is skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when no cell shows only "#"
'   characters, otherwise FAIL.
'
' BEHAVIOR
'   Runs the calculation, then reads the displayed text of every cell of the
'   used range of each sheet of ManagedSheets and counts those that IsHashes
'   accepts. Names the first cells found, up to MAX_LISTED.
'
' ERROR POLICY
'   None. An error raised by the run or the sheet calls is caught by RunCheck.
'
' DEPENDENCIES
'   RunFRTB_Silent, ManagedSheets, GetSheet, IsHashes, mInExcel, MAX_LISTED.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckDisplayedValues"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim SheetIndex              As Long         'Position in the sheet names
    Dim ManagedSheet            As Object       'Sheet being read
    Dim CurrentCell             As Object       'Cell being read
    Dim HashCount               As Long         'Cells that show only #
    Dim CellCount               As Long         'Cells read
    Dim ListedCells             As String       'Addresses of the first cells found

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: the displayed text depends on Excel's fonts and column widths"
            CheckDisplayedValues = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RUN AND READ CELLS
'------------------------------------------------------------------------------
    'Run the calculation
        RunFRTB_Silent
    'Read the displayed text of every used cell of every managed sheet
        SheetNames = ManagedSheets()
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
            For Each CurrentCell In ManagedSheet.UsedRange.Cells
                CellCount = CellCount + 1
                If IsHashes(CStr(CurrentCell.Text)) Then
                    HashCount = HashCount + 1
                    If HashCount <= MAX_LISTED Then
                        ListedCells = ListedCells & ManagedSheet.Name & "!" & _
                            CurrentCell.Address(False, False) & " "
                    End If
                End If
            Next CurrentCell
        Next SheetIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Describe the cells read on a pass, or the cells found on a fail
        If HashCount = 0 Then
            detail = CellCount & " cells on " & (UBound(SheetNames) - LBound(SheetNames) + 1) & _
                " sheets after a sample run; none shows ####"
            CheckDisplayedValues = "PASS"
        Else
            detail = HashCount & " cells show ####: " & ListedCells & _
                IIf(HashCount > MAX_LISTED, "and more", "")
            CheckDisplayedValues = "FAIL"
        End If

End Function

Private Function CheckNoDates(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                          CHECK THE ABSENCE OF DATES
'------------------------------------------------------------------------------
' PURPOSE
'   After a sample run, checks that no cell of the managed sheets holds a date:
'   MAR references, codes and tenors stay text or numbers (#206).
'
' WHY THIS EXISTS
'   MAR references, codes and tenors must stay text or numbers; a date value in
'   one of the sheets is a defect (#206).
'
' INPUTS
'   detail
'     Receives the number of cells and sheets read, or the cells that
'     hold a date
'
' RETURNS
'   PASS when no cell holds a date, otherwise FAIL.
'
' BEHAVIOR
'   Runs the calculation, then reads the value type of every cell of the used
'   range of each sheet of ManagedSheets and counts those of type Date. Names
'   the first cells found, up to MAX_LISTED.
'
' ERROR POLICY
'   None. An error raised by the run or the sheet calls is caught by RunCheck.
'
' DEPENDENCIES
'   RunFRTB_Silent, ManagedSheets, GetSheet, VAR_DATE, MAX_LISTED.
'
' NOTES
'   Runs in Excel and in LibreOffice.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckNoDates"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim SheetIndex              As Long         'Position in the sheet names
    Dim ManagedSheet            As Object       'Sheet being read
    Dim CurrentCell             As Object       'Cell being read
    Dim DateCount               As Long         'Cells that hold a date
    Dim CellCount               As Long         'Cells read
    Dim ListedCells             As String       'Addresses of the first cells found

'------------------------------------------------------------------------------
' RUN AND READ CELLS
'------------------------------------------------------------------------------
    'Run the calculation
        RunFRTB_Silent
    'Read the value type of every used cell of every managed sheet
        SheetNames = ManagedSheets()
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
            For Each CurrentCell In ManagedSheet.UsedRange.Cells
                CellCount = CellCount + 1
                If VarType(CurrentCell.Value) = VAR_DATE Then
                    DateCount = DateCount + 1
                    If DateCount <= MAX_LISTED Then
                        ListedCells = ListedCells & ManagedSheet.Name & "!" & _
                            CurrentCell.Address(False, False) & " "
                    End If
                End If
            Next CurrentCell
        Next SheetIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Describe the cells read on a pass, or the cells found on a fail
        If DateCount = 0 Then
            detail = CellCount & " cells on " & (UBound(SheetNames) - LBound(SheetNames) + 1) & _
                " sheets after a sample run; none holds a date"
            CheckNoDates = "PASS"
        Else
            detail = DateCount & " cells hold a date: " & ListedCells & _
                IIf(DateCount > MAX_LISTED, "and more", "")
            CheckNoDates = "FAIL"
        End If

End Function

Private Function CheckSheetView(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                             CHECK THE SHEET VIEW
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that every managed sheet hides its gridlines and has the tab color of
'   its role: README navy, inputs and parameters gold, outputs blue,
'   Test_Expected green (#206).
'
' WHY THIS EXISTS
'   Gridlines and tab colors are read from the Excel window and the sheet tab, so
'   the check runs only in Excel.
'
' INPUTS
'   detail
'     Receives the problems found, the number of sheets checked, or the
'     reason the check is skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when there is no problem,
'   otherwise FAIL.
'
' BEHAVIOR
'   Activates each sheet of ManagedSheets, requires that the window shows no
'   gridlines and that the tab color equals ExpectedTabColor for the sheet.
'   Activates the Home sheet at the end.
'
' ERROR POLICY
'   None. An error raised by the window or sheet calls is caught by RunCheck.
'
' DEPENDENCIES
'   ManagedSheets, GetSheet, ExpectedTabColor, Application.ActiveWindow, mInExcel.
'
' NOTES
'   Activates each sheet it checks.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckSheetView"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim SheetIndex              As Long         'Position in the sheet names
    Dim ManagedSheet            As Object       'Sheet being checked
    Dim Problems                As String       'Problems found, each ending with a semicolon

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: gridlines and tab colors as Excel shows them"
            CheckSheetView = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' CHECK SHEETS
'------------------------------------------------------------------------------
    'Activate each managed sheet; require no gridlines and the tab color of its role
        SheetNames = ManagedSheets()
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
            ManagedSheet.Activate
            If Application.ActiveWindow.DisplayGridlines Then
                Problems = Problems & ManagedSheet.Name & " shows gridlines; "
            End If
            If ManagedSheet.Tab.Color <> ExpectedTabColor(ManagedSheet.Name) Then
                Problems = Problems & ManagedSheet.Name & " tab color " & ManagedSheet.Tab.Color & _
                    ", not " & ExpectedTabColor(ManagedSheet.Name) & "; "
            End If
        Next SheetIndex
    'Leave Home active
        GetSheet(SH_HOME).Activate

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the number of sheets on a pass
        detail = IIf(Problems = vbNullString, (UBound(SheetNames) - LBound(SheetNames) + 1) & _
            " sheets without gridlines, each with the tab color of its role", Problems)
        CheckSheetView = IIf(Problems = vbNullString, "PASS", "FAIL")

End Function

Private Function CheckDataValidation(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                          CHECK THE DATA VALIDATION
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that every data validation rule of the template is on its sheet, with
'   the same type, alert style and error title (#348), also after Load sample
'   data and Reset parameters rewrite the input and parameter sheets.
'
' WHY THIS EXISTS
'   The rules are read from the cells as Excel holds them, so the check runs only
'   in Excel. The formulas are not compared: Excel returns them in the language
'   of the host.
'
' INPUTS
'   detail
'     Receives the problems found, the number of rules checked, or the
'     reason the check is skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when there is no problem,
'   otherwise FAIL.
'
' BEHAVIOR
'   Runs LoadSampleData and ResetParameters, then for each rule of
'   AcceptanceValidation (sheet, address, type, alert style, error title) reads
'   the rule of that cell with ReadValidation and records a problem when the cell
'   has none or when its type, alert style or error title differs.
'
' ERROR POLICY
'   None. An error raised by the setup commands is caught by RunCheck.
'
' DEPENDENCIES
'   AcceptanceValidation, ReadValidation, LoadSampleData, ResetParameters,
'   mInExcel.
'
' NOTES
'   Rewrites the input and parameter sheets with the sample data and the default
'   parameters.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckDataValidation"

    Dim RuleSpec                As Variant      'Array(sheet name, address, type, alert style, error title)
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim RuleCount               As Long         'Number of rules checked
    Dim RuleType                As Long         'Type of the rule read from the cell
    Dim RuleAlert               As Long         'Alert style of the rule read from the cell
    Dim RuleTitle               As String       'Error title of the rule read from the cell

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: the validation rules as Excel reads them"
            CheckDataValidation = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' REWRITE SHEETS
'------------------------------------------------------------------------------
    'Run both commands that rewrite their sheets: the rules must survive
        LoadSampleData
        ResetParameters

'------------------------------------------------------------------------------
' CHECK RULES
'------------------------------------------------------------------------------
    'Read the rule of each cell and compare its type, alert style and error title
        For Each RuleSpec In AcceptanceValidation()
            RuleCount = RuleCount + 1
            If Not ReadValidation(CStr(RuleSpec(0)), CStr(RuleSpec(1)), RuleType, RuleAlert, RuleTitle) Then
                Problems = Problems & RuleSpec(0) & "!" & RuleSpec(1) & " has no rule; "
            ElseIf RuleType <> RuleSpec(2) Or RuleAlert <> RuleSpec(3) Or RuleTitle <> RuleSpec(4) Then
                Problems = Problems & RuleSpec(0) & "!" & RuleSpec(1) & " has type " & RuleType & _
                    ", alert " & RuleAlert & " and title '" & RuleTitle & "'; "
            End If
        Next RuleSpec

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the number of rules on a pass
        detail = IIf(Problems = vbNullString, _
            RuleCount & " rules on the input and parameter sheets, as in the template, also after " & _
            "Load sample data and Reset parameters", Problems)
        CheckDataValidation = IIf(Problems = vbNullString, "PASS", "FAIL")

End Function

Private Function ReadValidation( _
    ByVal sheetName As String, _
    ByVal address As String, _
    ByRef kind As Long, _
    ByRef alert As Long, _
    ByRef title As String) As Boolean
'
'------------------------------------------------------------------------------
'                           READ THE RULE OF A CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the data validation rule of a cell: its type, alert style and error
'   title.
'
' INPUTS
'   sheetName
'     Name of the sheet that holds the cell
'
'   address
'     Address of the cell
'
'   kind
'     Receives the type of the rule
'
'   alert
'     Receives the alert style of the rule
'
'   title
'     Receives the error title of the rule
'
' RETURNS
'   True when the cell has a rule; False when it has none.
'
' ERROR POLICY
'   Errors are suppressed (On Error Resume Next) and the error state is cleared:
'   a cell without a rule makes the read fail, and that is the False result.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.ReadValidation"

    Dim CellRule                As Object       'Validation object of the cell

'------------------------------------------------------------------------------
' READ RULE
'------------------------------------------------------------------------------
    'Read the rule, ignoring the error a cell without a rule raises
        On Error Resume Next
        Set CellRule = GetSheet(sheetName).Range(address).Validation
        kind = CellRule.Type
    'Read the alert style and the error title when the rule exists
        If Err.Number = 0 Then
            alert = CellRule.AlertStyle
            title = CellRule.ErrorTitle
        End If
    'Return True when no error occurred, then restore normal error handling
        ReadValidation = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0

End Function

Private Function CheckNavigation(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                           CHECK THE NAVIGATION BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that every sheet carries the navigation bar of #347: a marker for
'   itself and a link to each of the other sheets, which Load sample data and
'   Reset parameters leave in place; and that Home lists every sheet once.
'
' WHY THIS EXISTS
'   The bar is made of shapes and hyperlinks that Excel builds, so the check runs
'   only in Excel.
'
' INPUTS
'   detail
'     Receives the problems found, the description of the bar, or the
'     reason the check is skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when there is no problem,
'   otherwise FAIL.
'
' BEHAVIOR
'   For each sheet of ManagedSheets records a problem when:
'     - its number of hyperlinks is not the number of sheets minus one
'     - it has no shape named Nav_Current
'     - it has no shape named Nav_ plus the name of another sheet
'     - column C of Home does not hold its name exactly once
'   Then, for each sheet, compares the number of button shapes (named btn_) with the number
'   of buttons SheetButtons defines for it (none when it returns Empty), and requires a
'   shape named btn_ plus the macro for each of them.
'   Requires, on each sheet, the two triangles NAV_TOGGLE_HIDE and NAV_TOGGLE_SHOW, and the
'   Home entry in the rows above the first scrolling row: START stays in view.
'   Last, calls SetNavigationVisible with True to start from the shown bar, then with
'   False and with True again, as the ribbon's toggle does, and after each call records a
'   problem for every sheet whose column A or whose Nav shapes disagree (the triangle that
'   shows the bar takes the opposite state of the others), or whose first
'   first button is not where it was, moved left by the width of column A while the bar is
'   hidden, or whose column A is not LAYOUT_NAV_WIDTH wide when the bar is shown again.
'   The bar is left shown. A pass reports how long the two toggles took, with the checks.
'
' ERROR POLICY
'   None. An error raised by the sheet calls is caught by RunCheck.
'
' DEPENDENCIES
'   ManagedSheets, GetSheet, ShapeExists, SheetButtons, SetNavigationVisible, SH_HOME,
'   LAYOUT_NAV_WIDTH, LAYOUT_FIRST_DATA_ROW, NAV_TOGGLE_HIDE, NAV_TOGGLE_SHOW, mInExcel.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckNavigation"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim SheetIndex              As Long         'Position of the sheet being checked in the sheet names
    Dim InnerIndex              As Long         'Other sheet position, then the buttons expected
    Dim ManagedSheet            As Object       'Sheet being checked
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim SheetCount              As Long         'Number of managed sheets
    Dim SheetButtonList         As Variant      'Buttons the sheet is defined to have, or Empty
    Dim ToggleStep              As Long         'Step of the toggle check: 0 hides the bar, 1 shows it
    Dim ShapeIndex              As Long         'Shape being read
    Dim NavWidth()              As Double       'Width of column A of each sheet before the toggle
    Dim ButtonLeft()            As Double       'Left edge of the first button of each sheet before the toggle
    Dim ButtonShapes            As Long         'Shapes named btn_ on the sheet
    Dim ToggleStart             As Double       'Timer when a toggle started
    Dim ToggleSeconds           As Double       'Seconds the two toggles took, without the checks

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: the bar of shapes and links that Excel builds"
            CheckNavigation = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' COUNT SHEETS
'------------------------------------------------------------------------------
    'Read the managed sheets and count them
        SheetNames = ManagedSheets()
        SheetCount = UBound(SheetNames) - LBound(SheetNames) + 1

'------------------------------------------------------------------------------
' CHECK BARS
'------------------------------------------------------------------------------
    'Check the links, the marker, the link shapes and the Home listing of each sheet
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
            If ManagedSheet.Hyperlinks.Count <> SheetCount - 1 Then
                Problems = Problems & ManagedSheet.Name & " has " & ManagedSheet.Hyperlinks.Count & _
                    " links, not " & (SheetCount - 1) & "; "
            End If
            If Not ShapeExists(ManagedSheet, "Nav_Current") Then
                Problems = Problems & ManagedSheet.Name & " has no marker; "
            End If
            If Not ShapeExists(ManagedSheet, NAV_TOGGLE_HIDE) Or Not ShapeExists(ManagedSheet, NAV_TOGGLE_SHOW) Then
                Problems = Problems & ManagedSheet.Name & " lacks a triangle to hide or show the bar; "
            End If
            If ManagedSheet.Name = SH_HOME Then
                If ManagedSheet.Shapes("Nav_Current").Top >= ManagedSheet.Cells(LAYOUT_FIRST_DATA_ROW + 1, 1).Top Then
                    Problems = Problems & "START is not in the rows that stay in view; "
                End If
            ElseIf ShapeExists(ManagedSheet, "Nav_" & SH_HOME) Then
                If ManagedSheet.Shapes("Nav_" & SH_HOME).Top >= ManagedSheet.Cells(LAYOUT_FIRST_DATA_ROW + 1, 1).Top Then
                    Problems = Problems & ManagedSheet.Name & " has START below the rows that stay in view; "
                End If
            End If
            For InnerIndex = LBound(SheetNames) To UBound(SheetNames)
                If CStr(SheetNames(InnerIndex)) <> ManagedSheet.Name Then
                    If Not ShapeExists(ManagedSheet, "Nav_" & CStr(SheetNames(InnerIndex))) Then
                        Problems = Problems & ManagedSheet.Name & " has no link to " & _
                            SheetNames(InnerIndex) & "; "
                    End If
                End If
            Next InnerIndex
            If Application.WorksheetFunction.CountIf(GetSheet(SH_HOME).Range("C:C"), _
                CStr(SheetNames(SheetIndex))) <> 1 Then
                Problems = Problems & "Home does not list " & SheetNames(SheetIndex) & " once; "
            End If
        Next SheetIndex

'------------------------------------------------------------------------------
' CHECK BUTTONS
'------------------------------------------------------------------------------
    'Compare the buttons of each sheet with the buttons it is defined to have;
    'Empty counts as none, because IIf would evaluate UBound of Empty too
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
            SheetButtonList = SheetButtons(ManagedSheet.Name)
            If IsEmpty(SheetButtonList) Then
                InnerIndex = 0
            Else
                InnerIndex = UBound(SheetButtonList) - LBound(SheetButtonList) + 1
            End If
            ButtonShapes = 0
            For ShapeIndex = 1 To ManagedSheet.Shapes.Count
                If Left$(ManagedSheet.Shapes(ShapeIndex).Name, 4) = "btn_" Then ButtonShapes = ButtonShapes + 1
            Next ShapeIndex
            If ButtonShapes <> InnerIndex Then
                Problems = Problems & ManagedSheet.Name & " has " & ButtonShapes & _
                    " buttons, not " & InnerIndex & "; "
            End If
            For ShapeIndex = 0 To InnerIndex - 1
                If Not ShapeExists(ManagedSheet, "btn_" & CStr(SheetButtonList(LBound(SheetButtonList) + ShapeIndex)(1))) Then
                    Problems = Problems & ManagedSheet.Name & " has no button for " & _
                        CStr(SheetButtonList(LBound(SheetButtonList) + ShapeIndex)(1)) & "; "
                End If
            Next ShapeIndex
        Next SheetIndex

'------------------------------------------------------------------------------
' CHECK THE TOGGLE
'------------------------------------------------------------------------------
    'Show the bar first, so the baseline is its shown state whatever state the workbook was in,
    'then hide it and show it again, as the ribbon's toggle does; after each call column A
    'is hidden exactly while the bar is, every Nav shape follows, and shown again column A
    'has the width it was built with
        SetNavigationVisible True
        ReDim NavWidth(LBound(SheetNames) To UBound(SheetNames))
        ReDim ButtonLeft(LBound(SheetNames) To UBound(SheetNames))
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
            NavWidth(SheetIndex) = ManagedSheet.Columns(1).Width
            SheetButtonList = SheetButtons(ManagedSheet.Name)
            If Not IsEmpty(SheetButtonList) Then
                ButtonLeft(SheetIndex) = ManagedSheet.Shapes("btn_" & CStr(SheetButtonList(LBound(SheetButtonList))(1))).Left
            End If
        Next SheetIndex
        For ToggleStep = 0 To 1
            ToggleStart = Timer
            SetNavigationVisible (ToggleStep = 1)
            ToggleSeconds = ToggleSeconds + Timer - ToggleStart
            For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
                Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
                If ManagedSheet.Columns(1).Hidden = (ToggleStep = 1) Then
                    Problems = Problems & ManagedSheet.Name & " has column A " & _
                        IIf(ToggleStep = 1, "hidden", "shown") & " while the bar is " & _
                        IIf(ToggleStep = 1, "shown", "hidden") & "; "
                End If
                SheetButtonList = SheetButtons(ManagedSheet.Name)
                If Not IsEmpty(SheetButtonList) Then
                    If Abs(ManagedSheet.Shapes("btn_" & CStr(SheetButtonList(LBound(SheetButtonList))(1))).Left - _
                        (ButtonLeft(SheetIndex) - IIf(ToggleStep = 0, NavWidth(SheetIndex), 0))) > 0.5 Then
                        Problems = Problems & ManagedSheet.Name & " has its buttons off their place; "
                    End If
                End If
                For ShapeIndex = 1 To ManagedSheet.Shapes.Count
                    If Left$(ManagedSheet.Shapes(ShapeIndex).Name, 3) = "Nav" Then
                        If (ManagedSheet.Shapes(ShapeIndex).Visible <> (ToggleStep = 1)) <> _
                            (ManagedSheet.Shapes(ShapeIndex).Name = NAV_TOGGLE_SHOW) Then
                            Problems = Problems & ManagedSheet.Name & " keeps " & _
                                ManagedSheet.Shapes(ShapeIndex).Name & IIf(ToggleStep = 1, " hidden", " shown") & "; "
                            Exit For
                        End If
                    End If
                Next ShapeIndex
                If ToggleStep = 1 Then
                    If Abs(ManagedSheet.Columns(1).ColumnWidth - LAYOUT_NAV_WIDTH) > 0.5 Then
                        Problems = Problems & ManagedSheet.Name & " has column A " & _
                            ManagedSheet.Columns(1).ColumnWidth & " wide, not " & LAYOUT_NAV_WIDTH & "; "
                    End If
                End If
            Next SheetIndex
        Next ToggleStep

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the description of the bar on a pass
        detail = IIf(Problems = vbNullString, SheetCount & " sheets, each with a marker and " & _
            (SheetCount - 1) & " links, all listed on Home; the buttons as set; " & _
            "the toggle hides and shows column A and the bar in " & Round(ToggleSeconds, 1) & " s", Problems)
        CheckNavigation = IIf(Problems = vbNullString, "PASS", "FAIL")

End Function

Private Function CheckKeyHighlights(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                           CHECK THE KEY HIGHLIGHTS
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that a key that repeats in a Key | Value table is highlighted (#350).
'   After Reset parameters, which clears the formats of its sheets and applies
'   the rule again, every table shows what the rule is for: its first key is not
'   red, and the same key typed into the row below makes that row red.
'
' WHY THIS EXISTS
'   The run on 0d7fbf9 showed the P_* tables without the rule after Reset
'   parameters, while Control, which it does not reset, had the template's. The
'   conditional format is read as Excel displays it, so the check runs only in
'   Excel.
'
' INPUTS
'   detail
'     Receives the problems found, the number of tables checked, or the
'     reason the check is skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when there is no problem,
'   otherwise FAIL.
'
' BEHAVIOR
'   Runs ResetParameters. For each table of AcceptanceKeyTables, reads the first
'   key cell and the cell below it, and notes whether the first key shows the
'   fail fill. Opens a write window (BeginWrite), because a key is locked on a
'   protected sheet, copies the first key into the cell below, notes whether that
'   cell shows the fail fill, writes the cell's value back and closes the write
'   window (EndWrite). Records a problem when the first key is red or when the
'   repeated key is not red.
'
' ERROR POLICY
'   None. An error raised by the sheet calls is caught by RunCheck.
'
' DEPENDENCIES
'   AcceptanceKeyTables, ResetParameters, GetSheet, TableCell, BeginWrite,
'   EndWrite, COLOR_FAIL_FILL, mInExcel.
'
' NOTES
'   Each probe restores the cell.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckKeyHighlights"

    Dim KeySpec                 As Variant      'Array(sheet name, row of the first key, ...)
    Dim KeySheet                As Object       'Sheet being checked
    Dim FirstKeyCell            As Object       'First key of the table
    Dim SecondKeyCell           As Object       'Cell below the first key
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim TableCount              As Long         'Number of tables checked
    Dim SavedValue              As Variant      'Value of the second cell before the probe
    Dim IsRed                   As Boolean      'True when the repeated key shows the fail fill
    Dim IsPlain                 As Boolean      'True when the first key does not show the fail fill

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: the conditional formats as Excel applies them"
            CheckKeyHighlights = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESET PARAMETERS
'------------------------------------------------------------------------------
    'Clear the formats of the parameter sheets and apply the rule again
        ResetParameters

'------------------------------------------------------------------------------
' CHECK TABLES
'------------------------------------------------------------------------------
    'Probe each table: the first key is plain and a repeat of it in the row below is red;
    'the repeat is typed inside a write window, because a key is locked on a protected sheet,
    'and the cell is restored
        For Each KeySpec In AcceptanceKeyTables()
            Set KeySheet = GetSheet(CStr(KeySpec(0)))
            Set FirstKeyCell = TableCell(KeySheet, CLng(KeySpec(1)), 1)
            Set SecondKeyCell = TableCell(KeySheet, CLng(KeySpec(1)) + 1, 1)
            TableCount = TableCount + 1
            IsPlain = (FirstKeyCell.DisplayFormat.Interior.Color <> COLOR_FAIL_FILL)
            BeginWrite
            SavedValue = SecondKeyCell.Value
            SecondKeyCell.Value = FirstKeyCell.Value
            IsRed = (SecondKeyCell.DisplayFormat.Interior.Color = COLOR_FAIL_FILL)
            SecondKeyCell.Value = SavedValue
            EndWrite
            If Not IsPlain Then
                Problems = Problems & KeySheet.Name & "!" & FirstKeyCell.Address(False, False) & _
                    " is red without a repeat; "
            End If
            If Not IsRed Then
                Problems = Problems & "a repeat at " & KeySheet.Name & "!" & _
                    SecondKeyCell.Address(False, False) & " is not red; "
            End If
        Next KeySpec

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the number of tables on a pass
        detail = IIf(Problems = vbNullString, TableCount & _
            " Key | Value tables after Reset parameters: a key is plain, and the same key in the row below is red", _
            Problems)
        CheckKeyHighlights = IIf(Problems = vbNullString, "PASS", "FAIL")

End Function

Private Function CheckProtection(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                             CHECK THE PROTECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that the sheets and the workbook structure are protected after every
'   command that ran above, without a password: the input fill marks exactly the
'   unlocked cells, the input sheets allow inserted rows, sorting and the
'   autofilter, a script cannot write to a locked cell and can write to an input
'   cell (#349).
'
' WHY THIS EXISTS
'   The protection is applied and enforced by Excel, so the check runs only in
'   Excel. Typing and filtering by a person are not run here.
'
' INPUTS
'   detail
'     Receives the problems found, the description of the protection, or
'     the reason the check is skipped
'
' RETURNS
'   SKIPPED when the host is not Excel, PASS when there is no problem,
'   otherwise FAIL.
'
' BEHAVIOR
'   For each sheet of ManagedSheets records a problem when:
'     - the sheet contents are not protected
'     - inserting rows, sorting and filtering are not allowed exactly on the
'       sheets whose name starts with In_
'   For each cell of the used range, counts the cells whose locked state equals
'   whether they have the input fill (COLOR_INPUT_CELL), names the first ones up
'   to MAX_LISTED, and counts the unlocked cells. Then records a problem when the
'   workbook structure is not protected, when a sheet is not protected for the
'   interface only after a write window (ProtectionMode), when a write to the
'   header of Out_Summary is refused, or when a write to the first data cell of
'   In_Delta is refused: a script writes, the user interface cannot.
'
' ERROR POLICY
'   None. An error raised by the sheet calls is caught by RunCheck.
'
' DEPENDENCIES
'   ManagedSheets, GetSheet, TableCell, WriteRefused, BeginWrite, EndWrite,
'   COLOR_INPUT_CELL, MAX_LISTED, mInExcel.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckProtection"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim SheetIndex              As Long         'Position in the sheet names
    Dim ManagedSheet            As Object       'Sheet being checked
    Dim CurrentCell             As Object       'Cell being read
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim WrongCount              As Long         'Cells whose locked state does not match their fill
    Dim UnlockedCount           As Long         'Cells that are not locked
    Dim ListedCells             As String       'Addresses of the first wrong cells found
    Dim IsInputSheet            As Boolean      'True when the sheet name starts with In_

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Report SKIPPED outside Excel
        If Not mInExcel Then
            detail = "Excel only: the protection as Excel applies and enforces it"
            CheckProtection = "SKIPPED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' CHECK SHEETS
'------------------------------------------------------------------------------
    'Check the protection, the allowed actions and the cell fills of each managed sheet
        SheetNames = ManagedSheets()
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
            IsInputSheet = (Left$(ManagedSheet.Name, 3) = "In_")
            If Not ManagedSheet.ProtectContents Then
                Problems = Problems & ManagedSheet.Name & " is not protected; "
            End If
            If ManagedSheet.Protection.AllowInsertingRows <> IsInputSheet Or _
                ManagedSheet.Protection.AllowSorting <> IsInputSheet Or _
                ManagedSheet.Protection.AllowFiltering <> IsInputSheet Then
                Problems = Problems & ManagedSheet.Name & " allows the wrong actions; "
            End If
            For Each CurrentCell In ManagedSheet.UsedRange.Cells
                If CurrentCell.Locked = (CurrentCell.Interior.Color = COLOR_INPUT_CELL) Then
                    WrongCount = WrongCount + 1
                    If WrongCount <= MAX_LISTED Then
                        ListedCells = ListedCells & ManagedSheet.Name & "!" & _
                            CurrentCell.Address(False, False) & " "
                    End If
                End If
                If Not CurrentCell.Locked Then UnlockedCount = UnlockedCount + 1
            Next CurrentCell
        Next SheetIndex

'------------------------------------------------------------------------------
' CHECK THE WORKBOOK
'------------------------------------------------------------------------------
    'Require the cells locked against their fill to be none
        If WrongCount > 0 Then
            Problems = Problems & WrongCount & " cells are locked against their fill: " & ListedCells & _
                IIf(WrongCount > MAX_LISTED, "and more", "") & "; "
        End If
    'Require the workbook structure to be protected
        If Not ThisWorkbook.ProtectStructure Then
            Problems = Problems & "the workbook structure is not protected; "
        End If
    'Require every sheet to be protected for the interface only after a write window, so that a
    'script writes to a header while the user interface cannot
        BeginWrite
        EndWrite
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            Set ManagedSheet = GetSheet(CStr(SheetNames(SheetIndex)))
            If Not (ManagedSheet.ProtectContents And ManagedSheet.ProtectionMode) Then
                Problems = Problems & ManagedSheet.Name & " is not protected for the interface only; "
            End If
        Next SheetIndex
        If WriteRefused(GetSheet(SH_OUT_SUMMARY).Cells(LAYOUT_HEADER_ROW, LAYOUT_FIRST_COL)) Then
            Problems = Problems & "a header on Out_Summary refused a script's write; "
        End If
    'Require a script's write to an input cell to be accepted
        If WriteRefused(TableCell(GetSheet(SH_IN_DELTA), LAYOUT_FIRST_DATA_ROW, 1)) Then
            Problems = Problems & "an input cell on In_Delta refused a script's write; "
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems, or the description of the protection on a pass
        detail = IIf(Problems = vbNullString, _
            (UBound(SheetNames) - LBound(SheetNames) + 1) & _
            " sheets and the structure protected without a password; " & _
            UnlockedCount & " unlocked cells, all in the input fill; protected for the interface only: a script writes to " & _
            "a header, the user interface cannot", _
            Problems)
        CheckProtection = IIf(Problems = vbNullString, "PASS", "FAIL")

End Function

Private Function WriteRefused(ByVal cell As Object) As Boolean

'
'------------------------------------------------------------------------------
'                       TEST WHETHER A WRITE IS REFUSED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when writing the cell's own value back raises; the value is
'   unchanged.
'
' INPUTS
'   cell
'     Cell to write to
'
' RETURNS
'   True when the write raised an error; False otherwise.
'
' ERROR POLICY
'   The error of the write is suppressed (On Error Resume Next) and cleared.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
    'Write the cell's value back to itself, ignoring the error a locked cell raises
        On Error Resume Next
        cell.Value = cell.Value
    'Return True when the write raised an error, then restore normal error handling
        WriteRefused = (Err.Number <> 0)
        Err.Clear
        On Error GoTo 0

End Function

Private Function ShapeExists(ByVal ws As Object, ByVal shapeName As String) As Boolean

'
'------------------------------------------------------------------------------
'                         TEST WHETHER A SHAPE EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the sheet has a shape of the given name.
'
' INPUTS
'   ws
'     Sheet to search
'
'   shapeName
'     Name of the shape
'
' RETURNS
'   True when the sheet has the shape; False otherwise.
'
' ERROR POLICY
'   The error of looking up a missing shape is suppressed (On Error Resume Next)
'   and cleared.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.ShapeExists"

    Dim FoundShape              As Object       'Shape of the given name, Nothing when the sheet has none

'------------------------------------------------------------------------------
' LOOK UP SHAPE
'------------------------------------------------------------------------------
    'Look the shape up by name, ignoring the error a missing shape raises
        On Error Resume Next
        Set FoundShape = ws.Shapes(shapeName)
    'Return True when no error occurred, then restore normal error handling
        ShapeExists = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0

End Function

Private Function CheckVariant(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'                        CHECK THE INVALID-ROW VARIANT
'------------------------------------------------------------------------------
' PURPOSE
'   Runs CheckVariantInWindow inside a write window and raises again any error it
'   produced.
'
' WHY THIS EXISTS
'   The variant writes the number format of cells, which is locked on a protected
'   sheet, so the check needs the write window (#349).
'
' INPUTS
'   detail
'     Receives the differences found, or on a pass the total and the
'     counts of errors and warnings
'
' RETURNS
'   PASS when CompareSummary finds no difference, otherwise FAIL.
'
' BEHAVIOR
'   Opens the write window (BeginWrite), calls CheckVariantInWindow with errors
'   suppressed (On Error Resume Next) and keeps its result and the number, source
'   and description of any error. Restores normal error handling, closes the
'   write window (EndWrite) and raises the error again when there was one.
'
' ERROR POLICY
'   An error in CheckVariantInWindow is held back until the write window is
'   closed, then raised again with the same number, source and description.
'
' DEPENDENCIES
'   CheckVariantInWindow, BeginWrite, EndWrite.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckVariant"

    Dim CheckResult             As String       'Result the inner check returns
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorSource             As String       'Captured error source
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' OPEN WRITE WINDOW
'------------------------------------------------------------------------------
    'Unprotect what the check needs: a cell format is locked on a protected sheet (#349)
        BeginWrite

'------------------------------------------------------------------------------
' RUN THE CHECK
'------------------------------------------------------------------------------
    'Run the check and keep any error, so the write window is closed first
        On Error Resume Next
        CheckResult = CheckVariantInWindow(detail)
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' CLOSE WRITE WINDOW
'------------------------------------------------------------------------------
    'Protect the sheets again
        EndWrite

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Raise the error again when the check failed
        If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription
    'Return the result of the check
        CheckVariant = CheckResult

End Function

Private Function CheckVariantInWindow(ByRef detail As String) As String

'
'------------------------------------------------------------------------------
'              CHECK THE INVALID-ROW VARIANT IN THE WRITE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Switches the discretions on, adds the invalid rows to the input sheets, runs,
'   and checks that every invalid row is logged and that the results equal the
'   reference for the same inputs (#28).
'
' WHY THIS EXISTS
'   The invalid rows must each be logged, and the results must stay equal to the
'   reference for the same inputs. CheckVariant needs the write window open; this
'   routine holds the check itself.
'
' INPUTS
'   detail
'     Receives the differences found, or on a pass the total and the
'     counts of errors and warnings
'
' RETURNS
'   PASS when CompareSummary finds no difference, otherwise FAIL.
'
' BEHAVIOR
'   Sets each setting of AcceptanceDiscretions to True on Control. Appends each
'   row of AcceptanceRows below the last data row of its sheet and applies the
'   boolean format to the column its definition names, when that is above
'   zero. Runs and compares the results with AcceptanceVariantExpected and the
'   Log counts with ACCEPTANCE_VARIANT_ERRORS and ACCEPTANCE_VARIANT_WARNINGS.
'
' ERROR POLICY
'   None. An error raised by the run or the sheet calls is raised to
'   CheckVariant, which raises it again to RunCheck.
'
' DEPENDENCIES
'   TEST_AcceptanceData (AcceptanceDiscretions, AcceptanceRows,
'   AcceptanceBooleanFormat, AcceptanceVariantExpected, ACCEPTANCE_VARIANT_ERRORS,
'   ACCEPTANCE_VARIANT_WARNINGS), SetControlValue, LastDataRow, TableCell,
'   RunFRTB_Silent, CompareSummary, OutSummaryValue, LogErrorCount,
'   LogWarningCount.
'
' NOTES
'   Runs last and leaves its rows and settings behind: the rows stay on the input
'   sheets of the workbook it runs in.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CheckVariantInWindow"

    Dim DiscretionCode          As Variant      'Code of a setting on Control
    Dim TargetSheet             As Object       'Sheet that receives the row
    Dim VariantRow              As Variant      'Row definition: sheet name, boolean column, values
    Dim TargetRow               As Long         'Row being written
    Dim ColumnIndex             As Long         'Element of the row definition being written

'------------------------------------------------------------------------------
' SET DISCRETIONS
'------------------------------------------------------------------------------
    'Switch each discretion on
        For Each DiscretionCode In AcceptanceDiscretions()
            SetControlValue CStr(DiscretionCode), True
        Next DiscretionCode

'------------------------------------------------------------------------------
' WRITE ROWS
'------------------------------------------------------------------------------
    'Append each invalid row below the last data row of its sheet
        For Each VariantRow In AcceptanceRows()
            Set TargetSheet = GetSheet(CStr(VariantRow(0)))
            TargetRow = LastDataRow(TargetSheet) + 1
            For ColumnIndex = 2 To UBound(VariantRow)
                If Not IsEmpty(VariantRow(ColumnIndex)) Then
                    TableCell(TargetSheet, TargetRow, ColumnIndex - 1).Value = VariantRow(ColumnIndex)
                End If
            Next ColumnIndex
            If VariantRow(1) > 0 Then
                TableCell(TargetSheet, TargetRow, CLng(VariantRow(1))).NumberFormat = _
                    AcceptanceBooleanFormat()
            End If
        Next VariantRow

'------------------------------------------------------------------------------
' RUN AND COMPARE
'------------------------------------------------------------------------------
    'Run the calculation
        RunFRTB_Silent
    'Compare the results and the Log counts with the reference; no difference means PASS
        detail = CompareSummary(AcceptanceVariantExpected(), ACCEPTANCE_VARIANT_ERRORS, _
            ACCEPTANCE_VARIANT_WARNINGS)
        CheckVariantInWindow = IIf(detail = vbNullString, "PASS", "FAIL")

'------------------------------------------------------------------------------
' DESCRIBE PASS
'------------------------------------------------------------------------------
    'On a pass, describe the total and the counts
        If detail = vbNullString Then
            detail = "total " & Format(OutSummaryValue("TOTAL"), "0.00") & ", " & LogErrorCount() & _
                " errors and " & LogWarningCount() & " warning, as the reference"
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function CompareSummary( _
    ByVal expected As Collection, _
    ByVal errors As Long, _
    ByVal warnings As Long) As String
'
'------------------------------------------------------------------------------
'                    COMPARE THE SUMMARY WITH THE REFERENCE
'------------------------------------------------------------------------------
' PURPOSE
'   Compares Out_Summary with the expected items and the Log counts with the
'   expected counts, and returns the differences.
'
' WHY THIS EXISTS
'   The sample run and the variant both compare their results with a reference;
'   this is the one place that does it.
'
' INPUTS
'   expected
'     Collection of Array(item, value) to compare with Out_Summary;
'     a string value is compared as text, any other as a number
'
'   errors
'     Expected number of ERROR rows in the Log
'
'   warnings
'     Expected number of warning rows in the Log
'
' RETURNS
'   The differences, each ending with a semicolon, or an empty string when
'   everything agrees.
'
' BEHAVIOR
'   Records a problem when the last run did not end with OK. For each expected
'   item reads its value in Out_Summary and records a problem when a text
'   differs, when a number is missing or when it differs by more than
'   AMOUNT_TOLERANCE. Then compares the Log error and warning counts.
'
' ERROR POLICY
'   None. Differences are returned, not raised.
'
' DEPENDENCIES
'   LastRunStatus, OutSummaryValue, LogErrorCount, LogWarningCount,
'   AMOUNT_TOLERANCE.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.CompareSummary"

    Dim ExpectedItem            As Variant      'Array(item, expected value)
    Dim SummaryValue            As Variant      'Value found in Out_Summary
    Dim Problems                As String       'Differences found, each ending with a semicolon

'------------------------------------------------------------------------------
' CHECK RUN STATUS
'------------------------------------------------------------------------------
    'Record a problem when the run did not end with OK
        If LastRunStatus() <> "OK" Then Problems = "the run failed: " & LastRunStatus() & "; "

'------------------------------------------------------------------------------
' COMPARE ITEMS
'------------------------------------------------------------------------------
    'Compare each expected item with its value in Out_Summary
        For Each ExpectedItem In expected
            SummaryValue = OutSummaryValue(CStr(ExpectedItem(0)))
            If VarType(ExpectedItem(1)) = vbString Then
                If CStr(SummaryValue) <> ExpectedItem(1) Then
                    Problems = Problems & ExpectedItem(0) & " " & SummaryValue & ", not " & _
                        ExpectedItem(1) & "; "
                End If
            ElseIf Not IsNumeric(SummaryValue) Or IsEmpty(SummaryValue) Then
                Problems = Problems & ExpectedItem(0) & " is missing; "
            ElseIf Abs(CDbl(SummaryValue) - CDbl(ExpectedItem(1))) > AMOUNT_TOLERANCE Then
                Problems = Problems & ExpectedItem(0) & " " & SummaryValue & ", not " & _
                    ExpectedItem(1) & "; "
            End If
        Next ExpectedItem

'------------------------------------------------------------------------------
' COMPARE LOG COUNTS
'------------------------------------------------------------------------------
    'Compare the error and warning counts of the Log with the expected ones
        If LogErrorCount() <> errors Then
            Problems = Problems & LogErrorCount() & " errors, not " & errors & "; "
        End If
        If LogWarningCount() <> warnings Then
            Problems = Problems & LogWarningCount() & " warnings, not " & warnings & "; "
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the differences
        CompareSummary = Problems

End Function

Private Function PlaceholderProblems( _
    ByVal sheetName As String, _
    ByVal note As String, _
    ByVal noteIsPrefix As Boolean) As String
'
'------------------------------------------------------------------------------
'                       CHECK THE EMPTY STATE OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Checks the empty state of an output sheet: its headers in row 5, the note in
'   C6 (on Results after a failure, the start of the alert) and nothing below.
'
' WHY THIS EXISTS
'   After a failed run and after Clear outputs the sheets must show the same
'   placeholder as the template (#85).
'
' INPUTS
'   sheetName
'     Sheet to check: Results, Out_Summary, Out_Class, Out_Buckets,
'     Out_DRC or Log
'
'   note
'     Text expected in the first data cell
'
'   noteIsPrefix
'     True when the cell may continue after the note (the alert on
'     Results); False when it must equal the note
'
' RETURNS
'   The problems found, each ending with a semicolon, or an empty string.
'
' BEHAVIOR
'   Compares the header captions of row 5 with the captions of the sheet.
'   For Out_DRC also requires the bucket table headers in J5 and P5. Compares
'   the first data cell with the note, and requires the cell below it to be
'   blank.
'
' ERROR POLICY
'   None. Problems are returned, not raised.
'
' DEPENDENCIES
'   GetSheet, TableCell, HeaderCaption, IsBlank, LAYOUT_HEADER_ROW,
'   LAYOUT_FIRST_DATA_ROW.
'
' NOTES
'   For a sheet name not listed, no headers are compared.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.PlaceholderProblems"

    Dim OutputSheet             As Object       'Sheet being checked
    Dim HeaderCodes             As Variant      'Header codes expected in row 5
    Dim HeaderIndex             As Long         'Header being compared
    Dim Problems                As String       'Problems found, each ending with a semicolon
    Dim ShownNote               As String       'Text in the first data cell

'------------------------------------------------------------------------------
' LOCATE SHEET
'------------------------------------------------------------------------------
    'Find the sheet to check
        Set OutputSheet = GetSheet(sheetName)
    'Select the header codes this sheet shows
        Select Case sheetName
            Case SH_RESULTS
                HeaderCodes = Array("Risk class", "Component", "Low", "Medium", "High")
            Case SH_OUT_SUMMARY
                HeaderCodes = Array("Item", "Value")
            Case SH_OUT_CLASS
                HeaderCodes = Array("RiskClass", "Component", "Scenario", "Capital", "Buckets", "AltSbUsed")
            Case SH_OUT_BUCKETS
                HeaderCodes = Array("RiskClass", "Component", "Scenario", "Bucket", "Kb", "Sb", "RiskFactors", "Note")
            Case SH_OUT_DRC
                HeaderCodes = Array("Obligor", "Bucket", "CreditQuality", "RW", "NetLong", "NetShort")
            Case SH_LOG
                HeaderCodes = Array("Level", "Source", "Message", "Row", "Reference")
            Case Else
                'Any other sheet name: nothing is done
        End Select

'------------------------------------------------------------------------------
' CHECK HEADERS
'------------------------------------------------------------------------------
    'Compare each header caption in row 5 with the expected caption
        For HeaderIndex = LBound(HeaderCodes) To UBound(HeaderCodes)
            If CStr(TableCell(OutputSheet, LAYOUT_HEADER_ROW, HeaderIndex + 1).Value) <> _
                CStr(HeaderCaption(HeaderCodes(HeaderIndex))) Then
                Problems = Problems & sheetName & " header " & (HeaderIndex + 1) & " is '" & _
                    TableCell(OutputSheet, LAYOUT_HEADER_ROW, HeaderIndex + 1).Value & "'; "
            End If
        Next HeaderIndex
    'Out_DRC also keeps the bucket table headers in J5 and P5
        If sheetName = SH_OUT_DRC Then
            If CStr(OutputSheet.Range("J5").Value) <> CStr(HeaderCaption("Bucket")) Or _
                CStr(OutputSheet.Range("P5").Value) <> CStr(HeaderCaption("DRC")) Then
                Problems = Problems & "Out_DRC bucket table headers missing in J5:P5; "
            End If
        End If

'------------------------------------------------------------------------------
' CHECK NOTE
'------------------------------------------------------------------------------
    'Read the first data cell
        ShownNote = CStr(TableCell(OutputSheet, LAYOUT_FIRST_DATA_ROW, 1).Value)
    'Compare it with the note, as a prefix or in full
        If noteIsPrefix Then
            If Left$(ShownNote, Len(note)) <> note Then
                Problems = Problems & sheetName & " C6 is '" & ShownNote & "'; "
            End If
        ElseIf ShownNote <> note Then
            Problems = Problems & sheetName & " C6 is '" & ShownNote & "'; "
        End If
    'Require that nothing sits below the note
        If Not IsBlank(TableCell(OutputSheet, LAYOUT_FIRST_DATA_ROW + 1, 1).Value) Then
            Problems = Problems & sheetName & " keeps rows below the note; "
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the problems
        PlaceholderProblems = Problems

End Function

Private Function MessageProblems(ByVal expected As Collection) As String

'
'------------------------------------------------------------------------------
'                      CHECK THE EXPECTED ERROR MESSAGES
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that each expected ERROR message is logged as many times as it is
'   expected.
'
' WHY THIS EXISTS
'   A message that must appear more than once is listed once for each time, so
'   the check compares counts and not only presence.
'
' INPUTS
'   expected
'     Collection of the expected ERROR messages, a message repeated as
'     many times as it must be logged
'
' RETURNS
'   The problems found, each ending with a semicolon, or an empty string.
'
' BEHAVIOR
'   For each expected message, counts how often it appears in the collection and
'   compares that count with the number of Log rows of level ERROR that carry
'   exactly that message.
'
' ERROR POLICY
'   None. Problems are returned, not raised.
'
' DEPENDENCIES
'   LogMessageCount.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.MessageProblems"

    Dim ExpectedMessage         As Variant      'Message being checked
    Dim OtherMessage            As Variant      'Message compared with it
    Dim WantedCount             As Long         'Times the message is expected
    Dim Problems                As String       'Problems found, each ending with a semicolon

'------------------------------------------------------------------------------
' CHECK MESSAGES
'------------------------------------------------------------------------------
    'Count how often each message is expected and compare with the Log
        For Each ExpectedMessage In expected
            WantedCount = 0
            For Each OtherMessage In expected
                If OtherMessage = ExpectedMessage Then WantedCount = WantedCount + 1
            Next OtherMessage
            If LogMessageCount("ERROR", CStr(ExpectedMessage)) <> WantedCount Then
                Problems = Problems & "'" & ExpectedMessage & "' logged " & _
                    LogMessageCount("ERROR", CStr(ExpectedMessage)) & " times, not " & WantedCount & "; "
            End If
        Next ExpectedMessage
    'Return the problems
        MessageProblems = Problems

End Function

Private Function LogMessageCount(ByVal level As String, ByVal text As String) As Long

'
'------------------------------------------------------------------------------
'                       COUNT THE LOG ROWS OF A MESSAGE
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the Log rows with the given level and exactly the given message.
'
' INPUTS
'   level
'     Level in the first column of the Log
'
'   text
'     Message in the third column of the Log
'
' RETURNS
'   Number of rows that match both; 0 when none.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.LogMessageCount"

    Dim LogSheet                As Object       'Log sheet
    Dim RowIndex                As Long         'Row being read

'------------------------------------------------------------------------------
' COUNT ROWS
'------------------------------------------------------------------------------
    'Read the rows from the first data row to the first blank first cell
        Set LogSheet = GetSheet(SH_LOG)
        RowIndex = LAYOUT_FIRST_DATA_ROW
    'Count the rows with this level and message
        Do While Not IsBlank(TableCell(LogSheet, RowIndex, 1).Value)
            If CStr(TableCell(LogSheet, RowIndex, 1).Value) = level And _
                CStr(TableCell(LogSheet, RowIndex, 3).Value) = text Then
                LogMessageCount = LogMessageCount + 1
            End If
            RowIndex = RowIndex + 1
        Loop

End Function

Private Sub RemoveRows(ByVal written As Collection)

'
'------------------------------------------------------------------------------
'                         CLEAR THE ROWS A CHECK WROTE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the rows a check wrote.
'
' INPUTS
'   written
'     Collection of Array(sheet name, row, number of columns), one for
'     each row to clear
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Errors while clearing a row are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.RemoveRows"

    Dim WrittenRow              As Variant      'Array(sheet name, row, number of columns)
    Dim TargetSheet             As Object       'Sheet that holds the row

'------------------------------------------------------------------------------
' CLEAR ROWS
'------------------------------------------------------------------------------
    'Clear the cells of each row, ignoring errors, then restore normal error handling
        On Error Resume Next
        For Each WrittenRow In written
            Set TargetSheet = GetSheet(CStr(WrittenRow(0)))
            TargetSheet.Range(TableCell(TargetSheet, CLng(WrittenRow(1)), 1), _
                TableCell(TargetSheet, CLng(WrittenRow(1)), CLng(WrittenRow(2)))).ClearContents
        Next WrittenRow
        On Error GoTo 0

End Sub

Private Function LogHas( _
    ByVal level As String, _
    ByVal source As String, _
    ByVal text As String) As Boolean
'
'------------------------------------------------------------------------------
'                          TEST THE LOG FOR A MESSAGE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the Log holds a row of the given level whose source starts
'   with source and whose message contains text.
'
' INPUTS
'   level
'     Level in the first column of the Log
'
'   source
'     Start of the source in the second column; an empty string matches
'     any source
'
'   text
'     Text the message in the third column must contain, compared without
'     regard to case; an empty string matches any message
'
' RETURNS
'   True when a row matches; False otherwise.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.LogHas"

    Dim LogSheet                As Object       'Log sheet
    Dim RowIndex                As Long         'Row being read

'------------------------------------------------------------------------------
' SEARCH LOG
'------------------------------------------------------------------------------
    'Read the rows from the first data row to the first blank first cell
        Set LogSheet = GetSheet(SH_LOG)
        RowIndex = LAYOUT_FIRST_DATA_ROW
    'Stop at the first row that matches level, source and message
        Do While Not IsBlank(TableCell(LogSheet, RowIndex, 1).Value)
            If CStr(TableCell(LogSheet, RowIndex, 1).Value) = level And _
                Left$(CStr(TableCell(LogSheet, RowIndex, 2).Value), Len(source)) = source And _
                InStr(1, CStr(TableCell(LogSheet, RowIndex, 3).Value), text, TEXT_COMPARE) > 0 Then
                LogHas = True
                Exit Function
            End If
            RowIndex = RowIndex + 1
        Loop

End Function

Private Function HeaderColumn( _
    ByVal ws As Object, _
    ByVal caption As String, _
    ByVal headerRow As Long) As Long
'
'------------------------------------------------------------------------------
'                             FIND A HEADER COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the column, counted from column C as 1, of a header in the header
'   row of a table.
'
' INPUTS
'   ws
'     Sheet that holds the table
'
'   caption
'     Caption to find
'
'   headerRow
'     Row that holds the headers
'
' RETURNS
'   The position of the first column, among the first 30, whose caption equals the
'   given one.
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP when no column carries the caption.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.HeaderColumn"

    Dim ColumnIndex             As Long         'Column being read

'------------------------------------------------------------------------------
' FIND COLUMN
'------------------------------------------------------------------------------
    'Read up to 30 columns and return the first whose caption matches
        For ColumnIndex = 1 To 30
            If CStr(TableCell(ws, headerRow, ColumnIndex).Value) = caption Then
                HeaderColumn = ColumnIndex
                Exit Function
            End If
        Next ColumnIndex

'------------------------------------------------------------------------------
' REPORT MISSING COLUMN
'------------------------------------------------------------------------------
    'Raise an error naming the sheet and the caption
        Err.Raise vbObjectError + ERR_APP, "TEST_Acceptance", _
            ws.Name & " has no column '" & caption & "'"

End Function

Private Function HeaderCount(ByVal ws As Object) As Long

'
'------------------------------------------------------------------------------
'                              COUNT THE HEADERS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of non-blank headers in the header row, from the first
'   column to the first blank one.
'
' INPUTS
'   ws
'     Sheet that holds the table
'
' RETURNS
'   The number of headers; at least 1.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.HeaderCount"

    Dim ColumnCount             As Long         'Headers counted

'------------------------------------------------------------------------------
' COUNT HEADERS
'------------------------------------------------------------------------------
    'Count the headers up to the first blank cell
        ColumnCount = 1
        Do While Not IsBlank(TableCell(ws, LAYOUT_HEADER_ROW, ColumnCount + 1).Value)
            ColumnCount = ColumnCount + 1
        Loop
    'Return the count
        HeaderCount = ColumnCount

End Function

Private Function IsHashes(ByVal shown As String) As Boolean

'
'------------------------------------------------------------------------------
'                          TEST FOR A HASH-ONLY TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the displayed text of a cell is only "#" characters, as
'   Excel shows a value too wide for its column.
'
' INPUTS
'   shown
'     Displayed text of the cell
'
' RETURNS
'   True when the text is not empty and holds nothing but "#"; False otherwise.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
    'An empty text is not a hash-only text
        If Len(shown) = 0 Then Exit Function
    'Remove the # characters; nothing left means a hash-only text
        IsHashes = (Replace(shown, "#", "") = vbNullString)

End Function

Private Function ExpectedTabColor(ByVal sheetName As String) As Long

'
'------------------------------------------------------------------------------
'                        RETURN THE EXPECTED TAB COLOR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the tab color M_Setup gives a sheet of this name.
'
' INPUTS
'   sheetName
'     Name of the sheet
'
' RETURNS
'   TAB_REFERENCE for SH_README and SH_HOME, TAB_TEST for SH_TEST, TAB_OUTPUT for
'   SH_RESULTS, SH_LOG and the sheets whose name starts with Out_, and TAB_INPUT
'   for every other sheet.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' SELECT COLOR
'------------------------------------------------------------------------------
    'Pick the color of the role of the sheet
        Select Case True
            Case sheetName = SH_README, sheetName = SH_HOME
                ExpectedTabColor = TAB_REFERENCE
            Case sheetName = SH_TEST
                ExpectedTabColor = TAB_TEST
            Case sheetName = SH_RESULTS, sheetName = SH_LOG, Left$(sheetName, 4) = "Out_"
                ExpectedTabColor = TAB_OUTPUT
            Case Else
                ExpectedTabColor = TAB_INPUT
        End Select

End Function

Private Function LastDataRow(ByVal ws As Object) As Long

'
'------------------------------------------------------------------------------
'                            FIND THE LAST DATA ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last row below the header with a value in column C or D.
'
' INPUTS
'   ws
'     Sheet that holds the table
'
' RETURNS
'   The last row of the block of rows that start at the first data row and have
'   a value in the first or the second table column; one row above the first
'   data row when there is none.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.LastDataRow"

    Dim RowIndex                As Long         'Row being read

'------------------------------------------------------------------------------
' FIND ROW
'------------------------------------------------------------------------------
    'Walk down from the first data row while column C or D holds a value
        RowIndex = LAYOUT_FIRST_DATA_ROW
        Do While Not IsBlank(TableCell(ws, RowIndex, 1).Value) Or _
            Not IsBlank(TableCell(ws, RowIndex, 2).Value)
            RowIndex = RowIndex + 1
        Loop
    'Return the row above the first row with both cells blank
        LastDataRow = RowIndex - 1

End Function

Private Sub SetControlValue(ByVal code As String, ByVal settingValue As Variant)

'
'------------------------------------------------------------------------------
'                            SET A CONTROL SETTING
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the value of one setting on Control.
'
' INPUTS
'   code
'     Code of the setting, in the first table column
'
'   settingValue
'     Value written in the second table column
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP when no setting carries the code.
'
' NOTES
'   Searches the header row and the 60 rows below it.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Acceptance.SetControlValue"

    Dim ControlSheet            As Object       'Control sheet
    Dim RowIndex                As Long         'Row being read

'------------------------------------------------------------------------------
' SET VALUE
'------------------------------------------------------------------------------
    'Find the Control sheet
        Set ControlSheet = GetSheet(SH_CONTROL)
    'Write the value on the row whose first cell holds the code
        For RowIndex = LAYOUT_HEADER_ROW To LAYOUT_HEADER_ROW + 60
            If CStr(TableCell(ControlSheet, RowIndex, 1).Value) = code Then
                TableCell(ControlSheet, RowIndex, 2).Value = settingValue
                Exit Sub
            End If
        Next RowIndex

'------------------------------------------------------------------------------
' REPORT MISSING SETTING
'------------------------------------------------------------------------------
    'Raise an error naming the code
        Err.Raise vbObjectError + ERR_APP, "TEST_Acceptance", "Control has no setting " & code

End Sub
