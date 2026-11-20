Attribute VB_Name = "TEST_Regression"
'==============================================================================
' MODULE: TEST_Regression
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the outputs of the last run with the [EXPECTED] table on
'   Test_Expected and mark every item PASS or FAIL. Items are Out_Summary
'   keys (for example "TOTAL") or "RiskClass|Component|Scenario" keys of
'   Out_RiskClass (for example "GIRR|DELTA|MEDIUM").
'
' PUBLIC SURFACE
'   Test_RunRegression: compare and report in a message box.
'   Test_RunRegression_Silent: compare without dialogs; the result line is
'   written to the subtitle cell of Test_Expected.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util; M_Main.OutSummaryValue.
'
' STATE OWNERSHIP
'   Writes the Actual, AbsDiff and Status columns of Test_Expected and its
'   result line; changes no other sheet.
'
' ERROR POLICY
'   A missing output counts as FAIL; a missing sheet raises (CORE_Util).
'
' COMPATIBILITY
'   Excel VBA; also runs in LibreOffice Calc with VBA support for the
'   automated tests. No references beyond the defaults.
'   Development build only: tools/Build-Workbook.ps1 -Production leaves
'   this module out.
'==============================================================================

'------------------------------------------------------------------------------
'                              MODULE SETTINGS
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
'                              MODULE CONSTANTS
'------------------------------------------------------------------------------
Private Const ABS_TOLERANCE As Double = 0.01                 'Absolute tolerance, reporting currency
Private Const REL_TOLERANCE As Double = 0.000000001          'Relative tolerance

'------------------------------------------------------------------------------
'                              PUBLIC MACROS
'------------------------------------------------------------------------------
Public Sub Test_RunRegression()
    Dim failures As Long, total As Long
    failures = ExecuteRegressionTest(total)
    If total = 0 Then
        MsgBox "No expected values found on " & SH_TEST & ".", vbExclamation, APP_NAME
    ElseIf failures = 0 Then
        MsgBox "Regression test PASSED (" & total & " items).", vbInformation, APP_NAME
    Else
        MsgBox "Regression test FAILED: " & failures & " of " & total & " items differ." & vbCrLf & _
               "See the " & SH_TEST & " sheet.", vbCritical, APP_NAME
    End If
End Sub


Public Sub Test_RunRegression_Silent()
    Dim total As Long
    ExecuteRegressionTest total
End Sub

'------------------------------------------------------------------------------
'                               COMPARISON
'------------------------------------------------------------------------------
Private Function ExecuteRegressionTest(ByRef total As Long) As Long
    Dim failures As Long
    Dim errNumber As Long, errSource As String, errDescription As String
    BeginWrite                                               'Test_Expected is protected (#349)
    On Error Resume Next
    failures = CompareExpected(total)
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error GoTo 0
    EndWrite
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    ExecuteRegressionTest = failures
End Function


Private Function CompareExpected(ByRef total As Long) As Long
    Dim ws As Object, r As Long, firstRow As Long, failures As Long
    Dim item As String, expected As Double, actual As Variant, diff As Double

    Set ws = GetSheet(SH_TEST)
    firstRow = FindTagRow(ws, "EXPECTED") + 2
    failures = 0
    total = 0
    r = firstRow
    Do While Not IsBlank(TableCell(ws, r, 1).Value) And firstRow > 2
        item = CStr(TableCell(ws, r, 1).Value)
        expected = ToDouble(TableCell(ws, r, 2).Value, 0#)
        actual = ActualValue(item)
        total = total + 1
        If IsEmpty(actual) Then
            TableCell(ws, r, 3).Value = "missing"
            TableCell(ws, r, 4).Value = ""
            MarkStatus TableCell(ws, r, 5), False
            failures = failures + 1
        Else
            diff = Abs(CDbl(actual) - expected)
            TableCell(ws, r, 3).Value = CDbl(actual)
            TableCell(ws, r, 4).Value = diff
            If diff <= ABS_TOLERANCE + REL_TOLERANCE * Abs(expected) Then
                MarkStatus TableCell(ws, r, 5), True
            Else
                MarkStatus TableCell(ws, r, 5), False
                failures = failures + 1
            End If
        End If
        r = r + 1
    Loop
    If failures = 0 And total > 0 Then
        TableCell(ws, LAYOUT_SUBTITLE_ROW, 1).Value = "Result: PASS (" & total & " items)"
        MarkStatus TableCell(ws, LAYOUT_SUBTITLE_ROW, 1), True
    Else
        TableCell(ws, LAYOUT_SUBTITLE_ROW, 1).Value = "Result: FAIL (" & failures & " of " & total & " items)"
        MarkStatus TableCell(ws, LAYOUT_SUBTITLE_ROW, 1), False
    End If
    CompareExpected = failures
End Function


Private Sub MarkStatus(ByVal cell As Object, ByVal passed As Boolean)
    If cell.Column <> LAYOUT_FIRST_COL Then cell.Value = IIf(passed, "PASS", "FAIL")
    On Error Resume Next
    cell.Font.Bold = True
    If passed Then
        cell.Interior.Color = COLOR_PASS_FILL
        cell.Font.Color = COLOR_PASS_TEXT
    Else
        cell.Interior.Color = COLOR_FAIL_FILL
        cell.Font.Color = COLOR_FAIL_TEXT
    End If
    On Error GoTo 0
End Sub


Private Function ActualValue(ByVal item As String) As Variant
    Dim parts As Variant, ws As Object, r As Long, v As Variant
    ActualValue = Empty
    parts = Split(item, "|")
    If UBound(parts) = 2 Then
        Set ws = GetSheet(SH_OUT_CLASS)
        r = LAYOUT_FIRST_DATA_ROW
        Do While Not IsBlank(TableCell(ws, r, 1).Value)
            If CStr(TableCell(ws, r, 1).Value) = parts(0) And CStr(TableCell(ws, r, 2).Value) = parts(1) _
               And CStr(TableCell(ws, r, 3).Value) = parts(2) Then
                ActualValue = TableCell(ws, r, 4).Value
                Exit Function
            End If
            r = r + 1
        Loop
    Else
        v = OutSummaryValue(item)
        If IsNumber(v) Then ActualValue = ToDouble(v, 0#)
    End If
End Function
