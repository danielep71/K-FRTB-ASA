Attribute VB_Name = "M_Main"
'==============================================================================
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the macros behind the sheet buttons, run the calculation with
'   Excel prepared for speed, put Excel back as it was, write the Results
'   and Out_Summary sheets and report the outcome to the user.
'
' PUBLIC SURFACE
'   RunFRTB: calculate SBM, DRC and RRAO and rewrite every output sheet.
'   RunFRTB_Silent: RunFRTB without dialogs, for automation and tests.
'   ClearOutputs: empty every output sheet and the log, keeping their
'   headers and a note, as in a new workbook.
'   LastRunStatus: "OK" or the error of the last run.
'   OutSummaryValue: a value of Out_Summary by item name.
'
' DEPENDENCIES
'   CORE_Params, CORE_SBM, CORE_DRC, CORE_RRAO for the calculation;
'   CORE_Util and CORE_Config for sheets, layout, palette and the log.
'
' STATE OWNERSHIP
'   Each run captures screen updating and calculation mode, changes them
'   for the run and restores the captured values on success and failure.
'   Owns the status of the last run (mLastStatus).
'
' ERROR POLICY
'   A run error is logged, stored in LastRunStatus, noted on Results and
'   shown in a message box by RunFRTB; RunFRTB_Silent never shows a dialog.
'   Every run first clears the output sheets, and a failed run clears them
'   again, so it leaves neither earlier nor partial capital behind. A
'   parameter that is blank or not a number is a run error. Errors in the
'   input rows are not VBA errors: they are listed on the Log sheet.
'
' KNOWN DEVIATION
'   The sheet buttons call this module directly, and it writes the Results
'   and Out_Summary sheets itself; there is no separate UI layer yet.
'
' COMPATIBILITY
'   Excel VBA; also runs in LibreOffice Calc with VBA support for the
'   automated tests. No references beyond the defaults.
'==============================================================================

'------------------------------------------------------------------------------
'                              MODULE SETTINGS
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
'                              MODULE CONSTANTS
'------------------------------------------------------------------------------
Private Const XL_CALCULATION_MANUAL As Long = -4135           'xlCalculationManual

'------------------------------------------------------------------------------
'                               MODULE STATE
'------------------------------------------------------------------------------
Private mLastStatus As String                                'OK or the error of the last run

'------------------------------------------------------------------------------
'                              BUTTON MACROS
'------------------------------------------------------------------------------
Public Sub RunFRTB()
    Dim ok As Boolean
    ok = ExecuteRun()
    If ok Then
        MsgBox "Calculation complete." & vbCrLf & _
               "Total capital: " & Format(OutSummaryValue("TOTAL"), "#,##0") & vbCrLf & _
               "Errors: " & LogErrorCount() & "   Warnings: " & LogWarningCount() & vbCrLf & _
               "See the Results and Log sheets.", vbInformation, APP_NAME
    Else
        MsgBox "Calculation failed: " & mLastStatus, vbCritical, APP_NAME
    End If
End Sub


Public Sub RunFRTB_Silent()
    ExecuteRun
End Sub


Public Function LastRunStatus() As String
    LastRunStatus = mLastStatus
End Function


Public Sub ClearOutputs()
    ClearResultSheets "No results yet: press Run FRTB-SA.", False
    If SheetExists(SH_LOG) Then
        ClearTableArea GetSheet(SH_LOG)
        WriteOutputPlaceholder GetSheet(SH_LOG)
    End If
End Sub

' Returns Results and the Out_* sheets, but not the Log, to their empty
' state: headers and a note, resultsNote on Results.
Private Sub ClearResultSheets(ByVal resultsNote As String, ByVal isAlert As Boolean)
    Dim names As Variant, i As Long, ws As Object
    names = Array(SH_OUT_CLASS, SH_OUT_BUCKETS, SH_OUT_DRC, SH_OUT_SUMMARY, SH_RESULTS)
    For i = LBound(names) To UBound(names)
        If SheetExists(CStr(names(i))) Then
            Set ws = GetSheet(CStr(names(i)))
            ClearTableArea ws
            If ws.Name = SH_RESULTS Then
                WriteOutputPlaceholder ws, resultsNote, isAlert
            Else
                WriteOutputPlaceholder ws
            End If
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
'                                  RUN
'------------------------------------------------------------------------------
Private Function ExecuteRun() As Boolean
    Dim savedScreen As Boolean                               'ScreenUpdating before the run
    Dim savedCalculation As Long                             'Calculation mode before the run
    Dim captured As Boolean                                  'True once both were read
    Dim errNumber As Long                                    'Primary error; 0 for none
    Dim errDescription As String

    On Error GoTo Failed
    CaptureSettings savedScreen, savedCalculation, captured
    ClearOutputs                                             'no earlier results after a failure
    ResetLog
    LogMessage "INFO", "RunFRTB", APP_NAME & " " & APP_VERSION & " started"
    LoadAllParameters
    RunSBM
    RunDRC
    RunRRAO
    WriteSummary
    LogMessage "INFO", "RunFRTB", "Finished with " & LogErrorCount() & " error(s) and " & LogWarningCount() & " warning(s)"
    mLastStatus = "OK"

CleanUp:
    On Error Resume Next
    If errNumber <> 0 Then
        mLastStatus = "Error " & errNumber & ": " & errDescription
        LogMessage "ERROR", "RunFRTB", mLastStatus
        ' Drop the rows written before the error.
        ClearResultSheets "Run failed (see the Log sheet): " & errDescription, True
    End If
    If captured Then RestoreSettings savedScreen, savedCalculation
    On Error GoTo 0
    ExecuteRun = (errNumber = 0)
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume CleanUp
End Function


' Best effort: a host without these properties leaves captured False.
Private Sub CaptureSettings(ByRef savedScreen As Boolean, ByRef savedCalculation As Long, ByRef captured As Boolean)
    On Error Resume Next
    savedScreen = Application.ScreenUpdating
    savedCalculation = Application.Calculation
    captured = (Err.Number = 0)
    Err.Clear
    Application.ScreenUpdating = False
    Application.Calculation = XL_CALCULATION_MANUAL
    On Error GoTo 0
End Sub


Private Sub RestoreSettings(ByVal savedScreen As Boolean, ByVal savedCalculation As Long)
    On Error Resume Next
    Application.Calculation = savedCalculation
    Application.ScreenUpdating = savedScreen
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
'                                 REPORTS
'------------------------------------------------------------------------------
Private Sub WriteSummary()
    Dim ws As Object, s As Long, best As Long
    Dim sbm As Double, drc As Double, rrao As Double

    best = SCEN_LOW
    For s = SCEN_LOW To SCEN_HIGH
        If SbmScenarioTotal(s) > SbmScenarioTotal(best) Then best = s
    Next s
    sbm = SbmScenarioTotal(best)
    drc = DrcTotal()
    rrao = RraoTotal()

    Set ws = GetSheet(SH_OUT_SUMMARY)
    ClearTableArea ws
    WriteTableHeader ws, LAYOUT_HEADER_ROW, Array("Item", "Value"), COLOR_OUTPUT_HEADER
    PutSummary ws, 1, "SBM_LOW", SbmScenarioTotal(SCEN_LOW)
    PutSummary ws, 2, "SBM_MEDIUM", SbmScenarioTotal(SCEN_MEDIUM)
    PutSummary ws, 3, "SBM_HIGH", SbmScenarioTotal(SCEN_HIGH)
    PutSummary ws, 4, "SBM_SCENARIO", ScenarioName(best)
    PutSummary ws, 5, "SBM", sbm
    PutSummary ws, 6, "DRC_NONSEC", drc
    PutSummary ws, 7, "RRAO", rrao
    PutSummary ws, 8, "TOTAL", sbm + drc + rrao
    PutSummary ws, 9, "ERRORS", LogErrorCount()
    PutSummary ws, 10, "WARNINGS", LogWarningCount()
    FormatDataRows ws, LAYOUT_FIRST_DATA_ROW, LAYOUT_HEADER_ROW + 10, 2, 2, "#,##0", False
    On Error Resume Next
    TableCell(ws, LAYOUT_HEADER_ROW + 4, 2).HorizontalAlignment = -4152    'the scenario name, as the numbers (#206)
    On Error GoTo 0

    WriteResultsSheet best, sbm, drc, rrao
End Sub


Private Sub PutSummary(ByVal ws As Object, ByVal i As Long, ByVal item As String, ByVal value As Variant)
    TableCell(ws, LAYOUT_HEADER_ROW + i, 1).Value = item
    TableCell(ws, LAYOUT_HEADER_ROW + i, 2).Value = value
End Sub


' Results: capital by risk class, component and scenario, then the capital
' requirement block, the run note and, on a scenario tie, the tie note.
Private Sub WriteResultsSheet(ByVal best As Long, ByVal sbm As Double, ByVal drc As Double, ByVal rrao As Double)
    Dim ws As Object, r As Long, i As Long, c As Long, s As Long, blockRow As Long
    Dim tied As String, nTied As Long
    Set ws = GetSheet(SH_RESULTS)
    ClearTableArea ws
    WriteTableHeader ws, LAYOUT_HEADER_ROW, Array("Risk class", "Component", "Low", "Medium", "High"), COLOR_OUTPUT_HEADER
    r = LAYOUT_HEADER_ROW
    For i = 1 To RC_COUNT
        For c = 1 To COMP_COUNT
            r = r + 1
            TableCell(ws, r, 1).Value = RiskClassCode(i)
            TableCell(ws, r, 2).Value = ComponentName(c)
            For s = SCEN_LOW To SCEN_HIGH
                TableCell(ws, r, 2 + s).Value = SbmCapital(i, c, s)
            Next s
        Next c
    Next i
    r = r + 1
    TableCell(ws, r, 1).Value = "SBM total"
    For s = SCEN_LOW To SCEN_HIGH
        TableCell(ws, r, 2 + s).Value = SbmScenarioTotal(s)
    Next s
    FormatDataRows ws, LAYOUT_FIRST_DATA_ROW, r, 5, 3, "#,##0", False
    StyleTotalRow ws, r, 5
    ' The binding scenario's total in navy: red would read as an error (#206)
    On Error Resume Next
    TableCell(ws, r, 2 + best).Font.Color = COLOR_INPUT_HEADER
    On Error GoTo 0

    blockRow = r + 2
    WriteKeyRow ws, blockRow + 1, "Binding correlation scenario", ScenarioName(best)
    WriteKeyRow ws, blockRow + 2, "SBM capital (max over scenarios)", sbm
    WriteKeyRow ws, blockRow + 3, "DRC (non-securitization)", drc
    WriteKeyRow ws, blockRow + 4, "Residual risk add-on", rrao
    WriteKeyRow ws, blockRow + 5, "Total capital requirement", sbm + drc + rrao
    FormatDataRows ws, blockRow + 1, blockRow + 5, 3, 3, "#,##0", False
    StyleTotalRow ws, blockRow + 5, 3
    ' After the rows, so their borders do not replace the heading rule, and
    ' as wide as the block below it (#206)
    WriteSectionHeading ws, blockRow, "Capital requirement", 3

    TableCell(ws, blockRow + 7, 1).Value = "Last run " & Format(Now, "yyyy-mm-dd hh:mm:ss") & _
        " - " & LogErrorCount() & " error(s), " & LogWarningCount() & " warning(s) on Log"
    StyleNote TableCell(ws, blockRow + 7, 1), (LogErrorCount() > 0)

    ' Ties resolve in the order low, medium, high: a reporting convention, not
    ' a reading of MAR21.7, since the capital is the same (#29, item 5)
    For s = SCEN_LOW To SCEN_HIGH
        If SbmScenarioTotal(s) = SbmScenarioTotal(best) Then
            nTied = nTied + 1
            If nTied > 1 Then tied = tied & ", "
            tied = tied & ScenarioName(s)
        End If
    Next s
    If nTied > 1 Then
        TableCell(ws, blockRow + 8, 1).Value = "Scenario tie (" & tied & "): same SBM total, " & _
            ScenarioName(best) & " reported; ties resolve in the order low, medium, high"
        StyleNote TableCell(ws, blockRow + 8, 1), False
    End If
End Sub


Private Sub WriteKeyRow(ByVal ws As Object, ByVal r As Long, ByVal label As String, ByVal value As Variant)
    TableCell(ws, r, 1).Value = label
    TableCell(ws, r, 3).Value = value
    On Error Resume Next
    TableCell(ws, r, 3).HorizontalAlignment = -4152          'a scenario name aligns with the amounts (#206)
    ws.Range(TableCell(ws, r, 1), TableCell(ws, r, 2)).Merge
    TableCell(ws, r, 1).Font.Bold = True
    On Error GoTo 0
End Sub


Private Sub WriteSectionHeading(ByVal ws As Object, ByVal r As Long, ByVal text As String, ByVal nCols As Long)
    TableCell(ws, r, 1).Value = text
    On Error Resume Next
    With TableCell(ws, r, 1).Font
        .Bold = True
        .Size = 14
        .Color = COLOR_INPUT_HEADER
    End With
    With ws.Range(TableCell(ws, r, 1), TableCell(ws, r, nCols)).Borders(9)    'xlEdgeBottom
        .LineStyle = 1
        .Weight = -4138                                                      'xlMedium
        .Color = COLOR_INPUT_HEADER
    End With
    ws.Rows(r).RowHeight = 30
    On Error GoTo 0
End Sub


Private Sub StyleTotalRow(ByVal ws As Object, ByVal r As Long, ByVal nCols As Long)
    On Error Resume Next
    With ws.Range(TableCell(ws, r, 1), TableCell(ws, r, nCols))
        .Font.Bold = True
        .Interior.Color = COLOR_TOTAL_FILL
    End With
    On Error GoTo 0
End Sub


Private Sub StyleNote(ByVal cell As Object, ByVal isAlert As Boolean)
    On Error Resume Next
    cell.Font.Italic = True
    cell.Font.Size = 10
    If isAlert Then
        cell.Font.Color = COLOR_FAIL_TEXT
        cell.Interior.Color = COLOR_FAIL_FILL
    Else
        cell.Font.Color = COLOR_NOTE
    End If
    On Error GoTo 0
End Sub


Public Function OutSummaryValue(ByVal item As String) As Variant
    Dim ws As Object, r As Long
    Set ws = GetSheet(SH_OUT_SUMMARY)
    OutSummaryValue = Empty
    r = LAYOUT_FIRST_DATA_ROW
    Do While Not IsBlank(TableCell(ws, r, 1).Value)
        If UCase$(CStr(TableCell(ws, r, 1).Value)) = UCase$(item) Then
            OutSummaryValue = TableCell(ws, r, 2).Value
            Exit Function
        End If
        r = r + 1
    Loop
End Function
