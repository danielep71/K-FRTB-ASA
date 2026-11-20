Attribute VB_Name = "TEST_Acceptance"
'==============================================================================
' MODULE: TEST_Acceptance
'------------------------------------------------------------------------------
' PURPOSE
'   Run the checks of the Excel acceptance without a person at the keyboard
'   (#274): the sample run, the state after a failed run and after Clear
'   outputs (#85), numeric text with a comma and with a point, the rejection
'   of sub-types MAR21 does not define (#245, #276), the frozen panes, the
'   autofilters, the width of the Control descriptions (#203), the values
'   Excel displays, the absence of dates and the view of every sheet (#206,
'   #316), the data validation of the input and parameter sheets (#348), the
'   navigation bar and the Home sheet (#347), and the invalid-row variant (#28).
'
' PUBLIC SURFACE
'   Test_RunAcceptanceInExcel_Silent: run every check and write one line per
'   check (PASS, FAIL or SKIPPED, with a detail) to the Acceptance sheet,
'   which it adds; cell C2 holds RESULT=PASS or RESULT=FAIL.
'   tools/Check-FRTB-Excel.ps1 runs it in desktop Excel.
'   Test_RunAcceptance_Silent: the same, with the checks only Excel can make
'   (the window panes, the autofilters, the column widths, the displayed
'   values, the sheet view, the data validation and the navigation bar)
'   SKIPPED; the LibreOffice tests run it.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, M_Main, M_Setup (ManagedSheets), TEST_AcceptanceData.
'
' STATE OWNERSHIP
'   Changes inputs, parameters, outputs and the window of the workbook it
'   runs in, and adds the Acceptance sheet: run it on a copy and close the
'   copy without saving. The variant runs last and leaves its rows behind.
'
' ERROR POLICY
'   An error inside a check makes that check FAIL with the error's text; the
'   other checks still run. No dialog is shown.
'
' COMPATIBILITY
'   Excel VBA; also runs in LibreOffice Calc with VBA support for the
'   automated tests through Test_RunAcceptance_Silent. LibreOffice reports
'   "Microsoft Excel" as its Application.Name, so the caller says which host
'   runs the checks. No references beyond the defaults.
'   Development build only: tools/Build-Workbook.ps1 -Production leaves
'   this module out.
'==============================================================================

'------------------------------------------------------------------------------
'                              MODULE SETTINGS
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
'                                CONSTANTS
'------------------------------------------------------------------------------
Private Const SH_ACCEPTANCE As String = "Acceptance"
Private Const AMOUNT_TOLERANCE As Double = 0.01             'Out_Summary amounts, in currency units
Private Const CLEARED_NOTE As String = "No results yet: press Run FRTB-SA."
Private Const FAILED_NOTE As String = "Run failed (see the Log sheet): "
Private Const TEXT_COMPARE As Long = 1                      'vbTextCompare
Private Const VAR_DATE As Long = 7                          'vbDate, as VarType returns it
Private Const MAX_LISTED As Long = 10                       'cells named in a FAIL detail

'------------------------------------------------------------------------------
'                               MODULE STATE
'------------------------------------------------------------------------------
Private mReport As Collection                                'Array(check, result, detail)
Private mInExcel As Boolean                                  'True: run the Excel-only checks

'------------------------------------------------------------------------------
'                              PUBLIC MACROS
'------------------------------------------------------------------------------
Public Sub Test_RunAcceptanceInExcel_Silent()
    RunAll True
End Sub


Public Sub Test_RunAcceptance_Silent()
    RunAll False
End Sub

'------------------------------------------------------------------------------
'                                  RUNNER
'------------------------------------------------------------------------------
Private Sub RunAll(ByVal inExcel As Boolean)
    mInExcel = inExcel
    Set mReport = New Collection
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
    RunCheck "PROTECTION"
    RunCheck "INVALID_ROW_VARIANT"                           'last: it adds input rows
    WriteReport
End Sub


Private Sub RunCheck(ByVal checkName As String)
    Dim detail As String, result As String
    On Error GoTo Failed
    Select Case checkName
        Case "SAMPLE_RUN": result = CheckSampleRun(detail)
        Case "FAILED_RUN": result = CheckFailedRun(detail)
        Case "CLEAR_OUTPUTS": result = CheckClearOutputs(detail)
        Case "NUMERIC_TEXT": result = CheckNumericText(detail)
        Case "UNDEFINED_SUB_TYPES": result = CheckSubTypes(detail)
        Case "FROZEN_PANES": result = CheckFrozenPanes(detail)
        Case "AUTOFILTERS": result = CheckAutofilters(detail)
        Case "CONTROL_DESCRIPTIONS": result = CheckControlDescriptions(detail)
        Case "DISPLAYED_VALUES": result = CheckDisplayedValues(detail)
        Case "NO_DATES": result = CheckNoDates(detail)
        Case "SHEET_VIEW": result = CheckSheetView(detail)
        Case "DATA_VALIDATION": result = CheckDataValidation(detail)
        Case "NAVIGATION": result = CheckNavigation(detail)
        Case "PROTECTION": result = CheckProtection(detail)
        Case "INVALID_ROW_VARIANT": result = CheckVariant(detail)
    End Select
    mReport.Add Array(checkName, result, detail)
    Exit Sub

Failed:
    CloseWriteWindows                                        'a check that failed inside a write window
    mReport.Add Array(checkName, "FAIL", "Error " & Err.Number & ": " & Err.Description)
End Sub


Private Sub WriteReport()
    Dim ws As Object, item As Variant, r As Long, anyFailed As Boolean
    BeginWrite                                               'a new sheet needs the structure unprotected (#349)
    If SheetExists(SH_ACCEPTANCE) Then
        Set ws = GetSheet(SH_ACCEPTANCE)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_ACCEPTANCE
    End If
    r = 4
    ws.Cells(r, 3).Value = "Check"
    ws.Cells(r, 4).Value = "Result"
    ws.Cells(r, 5).Value = "Detail"
    For Each item In mReport
        r = r + 1
        ws.Cells(r, 3).Value = item(0)
        ws.Cells(r, 4).Value = item(1)
        ws.Cells(r, 5).Value = item(2)
        If item(1) = "FAIL" Then anyFailed = True
    Next item
    ws.Cells(2, 3).Value = IIf(anyFailed, "RESULT=FAIL", "RESULT=PASS")
    EndWrite
End Sub

'------------------------------------------------------------------------------
'                                  CHECKS
'------------------------------------------------------------------------------
' The sample portfolio: no error and the reference results.
Private Function CheckSampleRun(ByRef detail As String) As String
    RunFRTB_Silent
    detail = CompareSummary(AcceptanceSampleExpected(), 0, 0)
    CheckSampleRun = IIf(detail = "", "PASS", "FAIL")
    If detail = "" Then detail = "total " & Format(OutSummaryValue("TOTAL"), "0.00") & ", 0 errors, 0 warnings"
End Function


' A blank GIRR risk weight stops the run: Results shows the alert, the Out_*
' sheets their placeholder, and the Log the error (#85).
Private Function CheckFailedRun(ByRef detail As String) As String
    Dim ws As Object, target As Object, saved As Variant, problems As String
    Dim errNumber As Long, errDescription As String
    Set ws = GetSheet(SH_P_GIRR)
    Set target = ws.Cells(FindTagRow(ws, "GIRR_RW") + 2, LAYOUT_FIRST_COL + 1)
    saved = target.Value
    target.ClearContents
    On Error GoTo Restore
    RunFRTB_Silent
    If LastRunStatus() = "OK" Then problems = problems & "the run did not fail; "
    If InStr(1, LastRunStatus(), "GIRR_RW", TEXT_COMPARE) = 0 Then problems = problems & "the error does not name GIRR_RW; "
    problems = problems & PlaceholderProblems(SH_RESULTS, FAILED_NOTE, True)
    problems = problems & PlaceholderProblems(SH_OUT_SUMMARY, CLEARED_NOTE, False)
    problems = problems & PlaceholderProblems(SH_OUT_CLASS, CLEARED_NOTE, False)
    problems = problems & PlaceholderProblems(SH_OUT_BUCKETS, CLEARED_NOTE, False)
    problems = problems & PlaceholderProblems(SH_OUT_DRC, CLEARED_NOTE, False)
    If Not LogHas("ERROR", "", "GIRR_RW") Then problems = problems & "the Log has no GIRR_RW error; "
    target.Value = saved
    On Error GoTo 0
    detail = IIf(problems = "", "Results C6: " & CStr(TableCell(GetSheet(SH_RESULTS), LAYOUT_FIRST_DATA_ROW, 1).Value), problems)
    CheckFailedRun = IIf(problems = "", "PASS", "FAIL")
    Exit Function

Restore:
    errNumber = Err.Number
    errDescription = Err.Description
    target.Value = saved
    Err.Raise errNumber, "TEST_Acceptance", errDescription
End Function


' Clear outputs after a successful run leaves the template's empty state (#85).
Private Function CheckClearOutputs(ByRef detail As String) As String
    Dim problems As String
    RunFRTB_Silent
    If LastRunStatus() <> "OK" Then problems = "the run before Clear outputs failed: " & LastRunStatus() & "; "
    ClearOutputs
    problems = problems & PlaceholderProblems(SH_RESULTS, CLEARED_NOTE, False)
    problems = problems & PlaceholderProblems(SH_OUT_SUMMARY, CLEARED_NOTE, False)
    problems = problems & PlaceholderProblems(SH_OUT_CLASS, CLEARED_NOTE, False)
    problems = problems & PlaceholderProblems(SH_OUT_BUCKETS, CLEARED_NOTE, False)
    problems = problems & PlaceholderProblems(SH_OUT_DRC, CLEARED_NOTE, False)
    problems = problems & PlaceholderProblems(SH_LOG, CLEARED_NOTE, False)
    detail = IIf(problems = "", "headers in row 5 and the note in row 6 on Results, the Out_* sheets and Log", problems)
    CheckClearOutputs = IIf(problems = "", "PASS", "FAIL")
End Function


' Numeric text has one, locale-independent form: a point for the decimals.
' A comma is rejected whatever the regional settings; a typed number is not
' text and is accepted (docs/workbook-guide.md).
Private Function CheckNumericText(ByRef detail As String) As String
    Dim result As String, errNumber As Long, errSource As String, errDescription As String
    BeginWrite                                               'a cell format is locked on a protected sheet (#349)
    On Error Resume Next
    result = CheckNumericTextInWindow(detail)
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error GoTo 0
    EndWrite
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    CheckNumericText = result
End Function


Private Function CheckNumericTextInWindow(ByRef detail As String) As String
    Dim ws As Object, target As Object, savedValue As Variant, savedFormat As String
    Dim cases As Variant, i As Long, rejected As Boolean, problems As String, source As String
    Dim errNumber As Long, errDescription As String
    Set ws = GetSheet(SH_IN_DELTA)
    Set target = TableCell(ws, LAYOUT_FIRST_DATA_ROW, HeaderColumn(ws, "Sensitivity", LAYOUT_HEADER_ROW))
    savedValue = target.Value
    savedFormat = target.NumberFormat
    source = SH_IN_DELTA & " row " & LAYOUT_FIRST_DATA_ROW
    ' text, then True when the row must be rejected
    cases = Array("0,5", True, "1.000,5", True, "1,5", True, "0.5", False, "-1.25e2", False)
    On Error GoTo Restore
    For i = LBound(cases) To UBound(cases) Step 2
        target.NumberFormat = "@"
        target.Value = CStr(cases(i))
        RunFRTB_Silent
        rejected = LogHas("ERROR", source, "Sensitivity is not numeric")
        If rejected <> CBool(cases(i + 1)) Then
            problems = problems & "text '" & cases(i) & "' was " & IIf(rejected, "rejected", "accepted") & "; "
        End If
    Next i
    target.NumberFormat = savedFormat
    target.Value = 0.5
    RunFRTB_Silent
    If LogHas("ERROR", source, "") Then problems = problems & "the typed number 0.5 was rejected; "
    target.NumberFormat = savedFormat
    target.Value = savedValue
    On Error GoTo 0
    detail = IIf(problems = "", "rejected as text: '0,5', '1.000,5', '1,5'; accepted: '0.5', '-1.25e2' and the number 0.5", problems)
    CheckNumericTextInWindow = IIf(problems = "", "PASS", "FAIL")
    Exit Function

Restore:
    errNumber = Err.Number
    errDescription = Err.Description
    target.NumberFormat = savedFormat
    target.Value = savedValue
    Err.Raise errNumber, "TEST_Acceptance", errDescription
End Function


' Rows whose sub-type MAR21 does not define are each rejected with their
' message, and the results stay those of the sample; the rows are removed
' afterwards, so the variant below is unaffected (#245).
Private Function CheckSubTypes(ByRef detail As String) As String
    Dim rowData As Variant, ws As Object, r As Long, k As Long, written As New Collection
    Dim problems As String, errNumber As Long, errDescription As String
    On Error GoTo Restore
    For Each rowData In AcceptanceSubTypeRows()
        Set ws = GetSheet(CStr(rowData(0)))
        r = LastDataRow(ws) + 1
        For k = 2 To UBound(rowData)
            If Not IsEmpty(rowData(k)) Then TableCell(ws, r, k - 1).Value = rowData(k)
        Next k
        written.Add Array(ws.Name, r, UBound(rowData) - 1)
    Next rowData
    RunFRTB_Silent
    problems = CompareSummary(AcceptanceSampleExpected(), AcceptanceSubTypeErrors().Count, 0)
    problems = problems & MessageProblems(AcceptanceSubTypeErrors())
    RemoveRows written
    On Error GoTo 0
    detail = IIf(problems = "", written.Count & " rows rejected with their messages; the results are the sample's", problems)
    CheckSubTypes = IIf(problems = "", "PASS", "FAIL")
    Exit Function

Restore:
    errNumber = Err.Number
    errDescription = Err.Description
    RemoveRows written
    Err.Raise errNumber, "TEST_Acceptance", errDescription
End Function


' Every sheet keeps the frozen panes of the committed template (#203).
Private Function CheckFrozenPanes(ByRef detail As String) As String
    Dim item As Variant, ws As Object, win As Object, problems As String, n As Long
    If Not mInExcel Then
        detail = "Excel only: the window panes as Excel shows them"
        CheckFrozenPanes = "SKIPPED"
        Exit Function
    End If
    For Each item In AcceptancePanes()
        Set ws = GetSheet(CStr(item(0)))
        ws.Activate
        Set win = Application.ActiveWindow
        If Not win.FreezePanes Or win.SplitRow <> item(1) Or win.SplitColumn <> item(2) Then
            problems = problems & item(0) & " frozen at " & win.SplitRow & " rows and " & win.SplitColumn & _
                       " columns, not " & item(1) & " and " & item(2) & "; "
        End If
        n = n + 1
    Next item
    detail = IIf(problems = "", n & " sheets frozen as in the template", problems)
    CheckFrozenPanes = IIf(problems = "", "PASS", "FAIL")
End Function


' Every input sheet has its autofilter, and filtering In_Delta on GIRR shows
' exactly the GIRR rows (#28).
Private Function CheckAutofilters(ByRef detail As String) As String
    Dim item As Variant, ws As Object, problems As String, r As Long, lastRow As Long
    Dim girrRows As Long, visibleGirr As Long, visibleOther As Long
    If Not mInExcel Then
        detail = "Excel only: the autofilters as Excel applies them"
        CheckAutofilters = "SKIPPED"
        Exit Function
    End If
    For Each item In AcceptanceFilters()
        Set ws = GetSheet(CStr(item(0)))
        If Not ws.AutoFilterMode Then
            problems = problems & item(0) & " has no autofilter; "
        ElseIf ws.AutoFilter.Range.Rows(1).Address(False, False) <> item(1) Then
            problems = problems & item(0) & " filters " & ws.AutoFilter.Range.Rows(1).Address(False, False) & _
                       ", not " & item(1) & "; "
        End If
    Next item
    Set ws = GetSheet(SH_IN_DELTA)
    lastRow = LastDataRow(ws)
    For r = LAYOUT_FIRST_DATA_ROW To lastRow
        If CStr(TableCell(ws, r, 2).Value) = "GIRR" Then girrRows = girrRows + 1
    Next r
    BeginWrite                                               'a script filters the protected sheet; a person's filter is not run here
    ws.Range(TableCell(ws, LAYOUT_HEADER_ROW, 1), TableCell(ws, lastRow, HeaderCount(ws))).AutoFilter 2, "GIRR"
    For r = LAYOUT_FIRST_DATA_ROW To lastRow
        If Not ws.Rows(r).Hidden Then
            If CStr(TableCell(ws, r, 2).Value) = "GIRR" Then visibleGirr = visibleGirr + 1 Else visibleOther = visibleOther + 1
        End If
    Next r
    If ws.FilterMode Then ws.ShowAllData
    EndWrite
    If girrRows = 0 Or visibleGirr <> girrRows Or visibleOther <> 0 Then
        problems = problems & "the GIRR filter on In_Delta showed " & visibleGirr & " of " & girrRows & _
                   " GIRR rows and " & visibleOther & " other rows; "
    End If
    detail = IIf(problems = "", "5 input sheets filtered as in the template; In_Delta filtered on GIRR shows its " & _
                 girrRows & " GIRR rows only", problems)
    CheckAutofilters = IIf(problems = "", "PASS", "FAIL")
End Function


' Every description on Control fits its column: Excel's own AutoFit of those
' cells does not widen the column (#203).
Private Function CheckControlDescriptions(ByRef detail As String) As String
    Dim ws As Object, col As Long, first As Long, last As Long, saved As Double, fitted As Double
    If Not mInExcel Then
        detail = "Excel only: the column width depends on Excel's fonts"
        CheckControlDescriptions = "SKIPPED"
        Exit Function
    End If
    Set ws = GetSheet(SH_CONTROL)
    col = LAYOUT_FIRST_COL + HeaderColumn(ws, "Description", LAYOUT_HEADER_ROW + 1) - 1
    first = LAYOUT_HEADER_ROW + 2
    last = LastDataRow(ws)
    saved = ws.Columns(col).ColumnWidth
    BeginWrite                                               'a column width is locked on a protected sheet
    ws.Range(ws.Cells(first, col), ws.Cells(last, col)).Columns.AutoFit
    fitted = ws.Columns(col).ColumnWidth
    ws.Columns(col).ColumnWidth = saved
    EndWrite
    detail = "Description column " & Format(saved, "0.00") & " wide; its texts need " & Format(fitted, "0.00")
    CheckControlDescriptions = IIf(fitted <= saved + 0.5, "PASS", "FAIL")
End Function


' After a sample run, no cell of any sheet shows "####": every column is
' wide enough for the values Excel displays (#206).
Private Function CheckDisplayedValues(ByRef detail As String) As String
    Dim names As Variant, i As Long, ws As Object, cell As Object, found As Long, cellCount As Long
    Dim listed As String
    If Not mInExcel Then
        detail = "Excel only: the displayed text depends on Excel's fonts and column widths"
        CheckDisplayedValues = "SKIPPED"
        Exit Function
    End If
    RunFRTB_Silent
    names = ManagedSheets()
    For i = LBound(names) To UBound(names)
        Set ws = GetSheet(CStr(names(i)))
        For Each cell In ws.UsedRange.Cells
            cellCount = cellCount + 1
            If IsHashes(CStr(cell.Text)) Then
                found = found + 1
                If found <= MAX_LISTED Then listed = listed & ws.Name & "!" & cell.Address(False, False) & " "
            End If
        Next cell
    Next i
    If found = 0 Then
        detail = cellCount & " cells on " & (UBound(names) - LBound(names) + 1) & _
                 " sheets after a sample run; none shows ####"
        CheckDisplayedValues = "PASS"
    Else
        detail = found & " cells show ####: " & listed & IIf(found > MAX_LISTED, "and more", "")
        CheckDisplayedValues = "FAIL"
    End If
End Function


' After a sample run, no cell of any sheet holds a date: MAR references,
' codes and tenors stay text or numbers (#206).
Private Function CheckNoDates(ByRef detail As String) As String
    Dim names As Variant, i As Long, ws As Object, cell As Object, found As Long, cellCount As Long
    Dim listed As String
    RunFRTB_Silent
    names = ManagedSheets()
    For i = LBound(names) To UBound(names)
        Set ws = GetSheet(CStr(names(i)))
        For Each cell In ws.UsedRange.Cells
            cellCount = cellCount + 1
            If VarType(cell.Value) = VAR_DATE Then
                found = found + 1
                If found <= MAX_LISTED Then listed = listed & ws.Name & "!" & cell.Address(False, False) & " "
            End If
        Next cell
    Next i
    If found = 0 Then
        detail = cellCount & " cells on " & (UBound(names) - LBound(names) + 1) & _
                 " sheets after a sample run; none holds a date"
        CheckNoDates = "PASS"
    Else
        detail = found & " cells hold a date: " & listed & IIf(found > MAX_LISTED, "and more", "")
        CheckNoDates = "FAIL"
    End If
End Function


' Every sheet hides its gridlines and has the tab color of its role: README
' navy, inputs and parameters gold, outputs blue, Test_Expected green (#206).
Private Function CheckSheetView(ByRef detail As String) As String
    Dim names As Variant, i As Long, ws As Object, problems As String
    If Not mInExcel Then
        detail = "Excel only: gridlines and tab colors as Excel shows them"
        CheckSheetView = "SKIPPED"
        Exit Function
    End If
    names = ManagedSheets()
    For i = LBound(names) To UBound(names)
        Set ws = GetSheet(CStr(names(i)))
        ws.Activate
        If Application.ActiveWindow.DisplayGridlines Then problems = problems & ws.Name & " shows gridlines; "
        If ws.Tab.Color <> ExpectedTabColor(ws.Name) Then
            problems = problems & ws.Name & " tab color " & ws.Tab.Color & ", not " & ExpectedTabColor(ws.Name) & "; "
        End If
    Next i
    GetSheet(SH_HOME).Activate
    detail = IIf(problems = "", (UBound(names) - LBound(names) + 1) & _
                 " sheets without gridlines, each with the tab color of its role", problems)
    CheckSheetView = IIf(problems = "", "PASS", "FAIL")
End Function


' Every data validation rule of the template is on its sheet, with the same
' type, alert style and error title (#348), also after Load sample data and
' Reset parameters rewrite the input and parameter sheets. The formulas are
' not compared: Excel returns them in the language of the host.
Private Function CheckDataValidation(ByRef detail As String) As String
    Dim item As Variant, problems As String, n As Long, kind As Long, alert As Long, title As String
    If Not mInExcel Then
        detail = "Excel only: the validation rules as Excel reads them"
        CheckDataValidation = "SKIPPED"
        Exit Function
    End If
    LoadSampleData                                           'both rewrite their sheets: the rules must survive
    ResetParameters
    For Each item In AcceptanceValidation()
        n = n + 1
        If Not ReadValidation(CStr(item(0)), CStr(item(1)), kind, alert, title) Then
            problems = problems & item(0) & "!" & item(1) & " has no rule; "
        ElseIf kind <> item(2) Or alert <> item(3) Or title <> item(4) Then
            problems = problems & item(0) & "!" & item(1) & " has type " & kind & ", alert " & alert & _
                       " and title '" & title & "'; "
        End If
    Next item
    detail = IIf(problems = "", n & " rules on the input and parameter sheets, as in the template, also after " & _
                 "Load sample data and Reset parameters", problems)
    CheckDataValidation = IIf(problems = "", "PASS", "FAIL")
End Function


' The rule of a cell: False when the cell has none.
Private Function ReadValidation(ByVal sheetName As String, ByVal address As String, ByRef kind As Long, _
                                ByRef alert As Long, ByRef title As String) As Boolean
    Dim v As Object
    On Error Resume Next
    Set v = GetSheet(sheetName).Range(address).Validation
    kind = v.Type
    If Err.Number = 0 Then
        alert = v.AlertStyle
        title = v.ErrorTitle
    End If
    ReadValidation = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' Every sheet carries the navigation bar of #347: a marker for itself and a
' link to each of the other sheets, which Load sample data and Reset
' parameters leave in place; and Home lists every sheet once.
Private Function CheckNavigation(ByRef detail As String) As String
    Dim names As Variant, i As Long, j As Long, ws As Object, problems As String, n As Long
    If Not mInExcel Then
        detail = "Excel only: the bar of shapes and links that Excel builds"
        CheckNavigation = "SKIPPED"
        Exit Function
    End If
    names = ManagedSheets()
    n = UBound(names) - LBound(names) + 1
    For i = LBound(names) To UBound(names)
        Set ws = GetSheet(CStr(names(i)))
        If ws.Hyperlinks.Count <> n - 1 Then
            problems = problems & ws.Name & " has " & ws.Hyperlinks.Count & " links, not " & (n - 1) & "; "
        End If
        If Not ShapeExists(ws, "Nav_Current") Then problems = problems & ws.Name & " has no marker; "
        For j = LBound(names) To UBound(names)
            If CStr(names(j)) <> ws.Name Then
                If Not ShapeExists(ws, "Nav_" & CStr(names(j))) Then
                    problems = problems & ws.Name & " has no link to " & names(j) & "; "
                End If
            End If
        Next j
        If Application.WorksheetFunction.CountIf(GetSheet(SH_HOME).Range("C:C"), CStr(names(i))) <> 1 Then
            problems = problems & "Home does not list " & names(i) & " once; "
        End If
    Next i
    detail = IIf(problems = "", n & " sheets, each with a marker and " & (n - 1) & " links, all listed on Home", problems)
    CheckNavigation = IIf(problems = "", "PASS", "FAIL")
End Function


' The sheets and the workbook structure are protected after every command
' that ran above, without a password: the input fill marks exactly the
' unlocked cells, the input sheets allow inserted rows, sorting and the
' autofilter, a script cannot write to a locked cell and can write to an
' input cell (#349). Typing and filtering by a person are not run here.
Private Function CheckProtection(ByRef detail As String) As String
    Dim names As Variant, i As Long, ws As Object, cell As Object, problems As String
    Dim wrong As Long, unlocked As Long, listed As String, isInput As Boolean
    If Not mInExcel Then
        detail = "Excel only: the protection as Excel applies and enforces it"
        CheckProtection = "SKIPPED"
        Exit Function
    End If
    names = ManagedSheets()
    For i = LBound(names) To UBound(names)
        Set ws = GetSheet(CStr(names(i)))
        isInput = (Left$(ws.Name, 3) = "In_")
        If Not ws.ProtectContents Then problems = problems & ws.Name & " is not protected; "
        If ws.Protection.AllowInsertingRows <> isInput Or ws.Protection.AllowSorting <> isInput Or _
           ws.Protection.AllowFiltering <> isInput Then
            problems = problems & ws.Name & " allows the wrong actions; "
        End If
        For Each cell In ws.UsedRange.Cells
            If cell.Locked = (cell.Interior.Color = COLOR_INPUT_CELL) Then
                wrong = wrong + 1
                If wrong <= MAX_LISTED Then listed = listed & ws.Name & "!" & cell.Address(False, False) & " "
            End If
            If Not cell.Locked Then unlocked = unlocked + 1
        Next cell
    Next i
    If wrong > 0 Then problems = problems & wrong & " cells are locked against their fill: " & listed & IIf(wrong > MAX_LISTED, "and more", "") & "; "
    If Not ThisWorkbook.ProtectStructure Then problems = problems & "the workbook structure is not protected; "
    If Not WriteRefused(GetSheet(SH_OUT_SUMMARY).Cells(LAYOUT_HEADER_ROW, LAYOUT_FIRST_COL)) Then
        problems = problems & "a header on Out_Summary accepted a script's write; "
    End If
    If WriteRefused(TableCell(GetSheet(SH_IN_DELTA), LAYOUT_FIRST_DATA_ROW, 1)) Then
        problems = problems & "an input cell on In_Delta refused a script's write; "
    End If
    detail = IIf(problems = "", (UBound(names) - LBound(names) + 1) & " sheets and the structure protected without a password; " & _
                 unlocked & " unlocked cells, all in the input fill; a script's write to a header is refused", problems)
    CheckProtection = IIf(problems = "", "PASS", "FAIL")
End Function


' True when writing the cell's own value back raises; the value is unchanged.
Private Function WriteRefused(ByVal cell As Object) As Boolean
    On Error Resume Next
    cell.Value = cell.Value
    WriteRefused = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
End Function


Private Function ShapeExists(ByVal ws As Object, ByVal shapeName As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = ws.Shapes(shapeName)
    ShapeExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' The discretions on and invalid rows on every input sheet: every invalid row
' is logged and the results equal the reference for the same inputs (#28).
Private Function CheckVariant(ByRef detail As String) As String
    Dim result As String, errNumber As Long, errSource As String, errDescription As String
    BeginWrite                                               'a cell format is locked on a protected sheet (#349)
    On Error Resume Next
    result = CheckVariantInWindow(detail)
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error GoTo 0
    EndWrite
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    CheckVariant = result
End Function


Private Function CheckVariantInWindow(ByRef detail As String) As String
    Dim item As Variant, ws As Object, rowData As Variant, r As Long, k As Long
    For Each item In AcceptanceDiscretions()
        SetControlValue CStr(item), True
    Next item
    For Each rowData In AcceptanceRows()
        Set ws = GetSheet(CStr(rowData(0)))
        r = LastDataRow(ws) + 1
        For k = 2 To UBound(rowData)
            If Not IsEmpty(rowData(k)) Then TableCell(ws, r, k - 1).Value = rowData(k)
        Next k
        If rowData(1) > 0 Then TableCell(ws, r, CLng(rowData(1))).NumberFormat = AcceptanceBooleanFormat()
    Next rowData
    RunFRTB_Silent
    detail = CompareSummary(AcceptanceVariantExpected(), ACCEPTANCE_VARIANT_ERRORS, ACCEPTANCE_VARIANT_WARNINGS)
    CheckVariantInWindow = IIf(detail = "", "PASS", "FAIL")
    If detail = "" Then
        detail = "total " & Format(OutSummaryValue("TOTAL"), "0.00") & ", " & LogErrorCount() & " errors and " & _
                 LogWarningCount() & " warning, as the reference"
    End If
End Function

'------------------------------------------------------------------------------
'                                 HELPERS
'------------------------------------------------------------------------------
' Compares Out_Summary with the expected items and the Log counts; returns
' the differences, or "" when everything agrees.
Private Function CompareSummary(ByVal expected As Collection, ByVal errors As Long, ByVal warnings As Long) As String
    Dim item As Variant, actual As Variant, problems As String
    If LastRunStatus() <> "OK" Then problems = "the run failed: " & LastRunStatus() & "; "
    For Each item In expected
        actual = OutSummaryValue(CStr(item(0)))
        If VarType(item(1)) = vbString Then
            If CStr(actual) <> item(1) Then problems = problems & item(0) & " " & actual & ", not " & item(1) & "; "
        ElseIf Not IsNumeric(actual) Or IsEmpty(actual) Then
            problems = problems & item(0) & " is missing; "
        ElseIf Abs(CDbl(actual) - CDbl(item(1))) > AMOUNT_TOLERANCE Then
            problems = problems & item(0) & " " & actual & ", not " & item(1) & "; "
        End If
    Next item
    If LogErrorCount() <> errors Then problems = problems & LogErrorCount() & " errors, not " & errors & "; "
    If LogWarningCount() <> warnings Then problems = problems & LogWarningCount() & " warnings, not " & warnings & "; "
    CompareSummary = problems
End Function


' The empty state of an output sheet: its headers in row 5, the note in C6
' (on Results after a failure, the start of the alert) and nothing below.
Private Function PlaceholderProblems(ByVal sheetName As String, ByVal note As String, ByVal noteIsPrefix As Boolean) As String
    Dim ws As Object, codes As Variant, i As Long, problems As String, shown As String
    Set ws = GetSheet(sheetName)
    Select Case sheetName
        Case SH_RESULTS: codes = Array("Risk class", "Component", "Low", "Medium", "High")
        Case SH_OUT_SUMMARY: codes = Array("Item", "Value")
        Case SH_OUT_CLASS: codes = Array("RiskClass", "Component", "Scenario", "Capital", "Buckets", "AltSbUsed")
        Case SH_OUT_BUCKETS: codes = Array("RiskClass", "Component", "Scenario", "Bucket", "Kb", "Sb", "RiskFactors", "Note")
        Case SH_OUT_DRC: codes = Array("Obligor", "Bucket", "CreditQuality", "RW", "NetLong", "NetShort")
        Case SH_LOG: codes = Array("Level", "Source", "Message")
    End Select
    For i = LBound(codes) To UBound(codes)
        If CStr(TableCell(ws, LAYOUT_HEADER_ROW, i + 1).Value) <> CStr(HeaderCaption(codes(i))) Then
            problems = problems & sheetName & " header " & (i + 1) & " is '" & TableCell(ws, LAYOUT_HEADER_ROW, i + 1).Value & "'; "
        End If
    Next i
    If sheetName = SH_OUT_DRC Then
        If CStr(ws.Range("J5").Value) <> CStr(HeaderCaption("Bucket")) Or CStr(ws.Range("P5").Value) <> CStr(HeaderCaption("DRC")) Then
            problems = problems & "Out_DRC bucket table headers missing in J5:P5; "
        End If
    End If
    shown = CStr(TableCell(ws, LAYOUT_FIRST_DATA_ROW, 1).Value)
    If noteIsPrefix Then
        If Left$(shown, Len(note)) <> note Then problems = problems & sheetName & " C6 is '" & shown & "'; "
    ElseIf shown <> note Then
        problems = problems & sheetName & " C6 is '" & shown & "'; "
    End If
    If Not IsBlank(TableCell(ws, LAYOUT_FIRST_DATA_ROW + 1, 1).Value) Then
        problems = problems & sheetName & " keeps rows below the note; "
    End If
    PlaceholderProblems = problems
End Function


' The expected ERROR messages, each logged as many times as it is expected.
Private Function MessageProblems(ByVal expected As Collection) As String
    Dim message As Variant, other As Variant, wanted As Long, problems As String
    For Each message In expected
        wanted = 0
        For Each other In expected
            If other = message Then wanted = wanted + 1
        Next other
        If LogMessageCount("ERROR", CStr(message)) <> wanted Then
            problems = problems & "'" & message & "' logged " & LogMessageCount("ERROR", CStr(message)) & _
                       " times, not " & wanted & "; "
        End If
    Next message
    MessageProblems = problems
End Function


Private Function LogMessageCount(ByVal level As String, ByVal text As String) As Long
    Dim ws As Object, r As Long
    Set ws = GetSheet(SH_LOG)
    r = LAYOUT_FIRST_DATA_ROW
    Do While Not IsBlank(TableCell(ws, r, 1).Value)
        If CStr(TableCell(ws, r, 1).Value) = level And CStr(TableCell(ws, r, 3).Value) = text Then
            LogMessageCount = LogMessageCount + 1
        End If
        r = r + 1
    Loop
End Function


' Clears the rows a check wrote: Array(sheet, row, number of columns) each.
Private Sub RemoveRows(ByVal written As Collection)
    Dim item As Variant, ws As Object
    On Error Resume Next
    For Each item In written
        Set ws = GetSheet(CStr(item(0)))
        ws.Range(TableCell(ws, CLng(item(1)), 1), TableCell(ws, CLng(item(1)), CLng(item(2)))).ClearContents
    Next item
    On Error GoTo 0
End Sub


' True when the Log holds a row of that level whose source starts with
' source and whose message contains text ("" matches any).
Private Function LogHas(ByVal level As String, ByVal source As String, ByVal text As String) As Boolean
    Dim ws As Object, r As Long
    Set ws = GetSheet(SH_LOG)
    r = LAYOUT_FIRST_DATA_ROW
    Do While Not IsBlank(TableCell(ws, r, 1).Value)
        If CStr(TableCell(ws, r, 1).Value) = level _
           And Left$(CStr(TableCell(ws, r, 2).Value), Len(source)) = source _
           And InStr(1, CStr(TableCell(ws, r, 3).Value), text, TEXT_COMPARE) > 0 Then
            LogHas = True
            Exit Function
        End If
        r = r + 1
    Loop
End Function


' Column (1 = column C) of a header in the table header row.
Private Function HeaderColumn(ByVal ws As Object, ByVal caption As String, ByVal headerRow As Long) As Long
    Dim k As Long
    For k = 1 To 30
        If CStr(TableCell(ws, headerRow, k).Value) = caption Then
            HeaderColumn = k
            Exit Function
        End If
    Next k
    Err.Raise vbObjectError + ERR_APP, "TEST_Acceptance", ws.Name & " has no column '" & caption & "'"
End Function


Private Function HeaderCount(ByVal ws As Object) As Long
    Dim k As Long
    k = 1
    Do While Not IsBlank(TableCell(ws, LAYOUT_HEADER_ROW, k + 1).Value)
        k = k + 1
    Loop
    HeaderCount = k
End Function


' True when a cell's displayed text is only "#", as Excel shows a value too
' wide for its column.
Private Function IsHashes(ByVal shown As String) As Boolean
    If Len(shown) = 0 Then Exit Function
    IsHashes = (Replace(shown, "#", "") = "")
End Function


' The tab color M_Setup gives a sheet of this name.
Private Function ExpectedTabColor(ByVal sheetName As String) As Long
    Select Case True
        Case sheetName = SH_README, sheetName = SH_HOME: ExpectedTabColor = TAB_REFERENCE
        Case sheetName = SH_TEST: ExpectedTabColor = TAB_TEST
        Case sheetName = SH_RESULTS, sheetName = SH_LOG, Left$(sheetName, 4) = "Out_": ExpectedTabColor = TAB_OUTPUT
        Case Else: ExpectedTabColor = TAB_INPUT
    End Select
End Function


' Last row below the header with a value in column C or D.
Private Function LastDataRow(ByVal ws As Object) As Long
    Dim r As Long
    r = LAYOUT_FIRST_DATA_ROW
    Do While Not IsBlank(TableCell(ws, r, 1).Value) Or Not IsBlank(TableCell(ws, r, 2).Value)
        r = r + 1
    Loop
    LastDataRow = r - 1
End Function


Private Sub SetControlValue(ByVal code As String, ByVal settingValue As Variant)
    Dim ws As Object, r As Long
    Set ws = GetSheet(SH_CONTROL)
    For r = LAYOUT_HEADER_ROW To LAYOUT_HEADER_ROW + 60
        If CStr(TableCell(ws, r, 1).Value) = code Then
            TableCell(ws, r, 2).Value = settingValue
            Exit Sub
        End If
    Next r
    Err.Raise vbObjectError + ERR_APP, "TEST_Acceptance", "Control has no setting " & code
End Sub
