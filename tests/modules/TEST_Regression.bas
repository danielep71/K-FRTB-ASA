Attribute VB_Name = "TEST_Regression"

Option Explicit

'
'==============================================================================
' MODULE: TEST_Regression
'------------------------------------------------------------------------------
' PURPOSE
'   Compares the outputs of the last run with the [EXPECTED] table on
'   Test_Expected and marks
'   every item PASS or FAIL. Items are Out_Summary keys (for example TOTAL) or
'   RiskClass|Component|Scenario keys of Out_RiskClass (for example
'   GIRR|DELTA|MEDIUM).
'
' WHY THIS EXISTS
'   The expected values are the regression oracle of the prototype: a run that
'   changes a
'   number shows up here.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Test_RunRegression compares and reports in a message box;
'   Test_RunRegression_Silent compares
'   without dialogs and writes the result line to the subtitle cell of
'   Test_Expected.
'
' ERROR POLICY
'   A missing output counts as FAIL; a missing sheet raises (CORE_Util).
'
' DEPENDENCIES
'   CORE_Config, CORE_Util; M_Main.OutSummaryValue.
'
' NOTES
'   Development build only: tools/Build-Workbook.ps1 -Production leaves this
'   module out.
'
'   The module also runs in LibreOffice Calc with VBA support for the automated
'   tests. It needs no
'   references beyond the defaults.
'
' PUBLIC SURFACE
'   Test_RunRegression: compare and report in a message box.
'   Test_RunRegression_Silent: compare without dialogs; the result line is
'   written to the
'   subtitle cell of Test_Expected.
'
' STATE OWNERSHIP
'   Writes the Actual, AbsDiff and Status columns of Test_Expected and its
'   result line;
'   changes no other sheet.
'
' UPDATED
'   2026-11-20
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const ABS_TOLERANCE As Double = 0.01        'Absolute tolerance, reporting currency
    Private Const REL_TOLERANCE As Double = 0.000000001 'Relative tolerance


'
'------------------------------------------------------------------------------
'
'                                PUBLIC MACROS
'
'------------------------------------------------------------------------------
'

Public Sub Test_RunRegression()
'
'------------------------------------------------------------------------------
'                           RUN THE REGRESSION TEST
'------------------------------------------------------------------------------
' PURPOSE
'   Compares the outputs of the last run with the expected values and reports
'   in a message box.
'
' WHY THIS EXISTS
'   The Test button needs one entry point that compares and tells the person
'   the outcome.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs ExecuteRegressionTest, then shows one of three messages: no expected
'   values found, PASSED with the
'   number of items, or FAILED with the number of items that differ.
'
' ERROR POLICY
'   Errors from the comparison are raised by ExecuteRegressionTest.
'
' DEPENDENCIES
'   ExecuteRegressionTest, APP_NAME, SH_TEST.
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
    Const PROC_NAME             As String = "TEST_Regression.Test_RunRegression"

    Dim FailureCount            As Long          'Items that differ
    Dim TotalCount              As Long          'Items compared

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Compare and report the outcome
        FailureCount = ExecuteRegressionTest(TotalCount)
        If TotalCount = 0 Then
            MsgBox "No expected values found on " & SH_TEST & ".", vbExclamation, APP_NAME
        ElseIf FailureCount = 0 Then
            MsgBox "Regression test PASSED (" & TotalCount & " items).", vbInformation, APP_NAME
        Else
            MsgBox "Regression test FAILED: " & FailureCount & " of " & TotalCount & " items differ." & vbCrLf & _
                   "See the " & SH_TEST & " sheet.", vbCritical, APP_NAME
        End If

End Sub

Public Sub Test_RunRegression_Silent()
'
'------------------------------------------------------------------------------
'                       RUN THE REGRESSION TEST SILENTLY
'------------------------------------------------------------------------------
' PURPOSE
'   Compares the outputs of the last run with the expected values without
'   dialogs.
'
' WHY THIS EXISTS
'   Automation and the tests need the comparison without a message box; the
'   result line is written to the
'   subtitle cell of Test_Expected.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs ExecuteRegressionTest and ignores the count it returns.
'
' ERROR POLICY
'   Errors from the comparison are raised by ExecuteRegressionTest.
'
' DEPENDENCIES
'   ExecuteRegressionTest
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
    Const PROC_NAME             As String = "TEST_Regression.Test_RunRegression_Silent"

    Dim TotalCount              As Long          'Items compared

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Compare without a dialog
        ExecuteRegressionTest TotalCount

End Sub


'
'------------------------------------------------------------------------------
'
'                                  COMPARISON
'
'------------------------------------------------------------------------------
'

Private Function ExecuteRegressionTest(ByRef total As Long) As Long
'
'------------------------------------------------------------------------------
'                         EXECUTE THE REGRESSION TEST
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the comparison inside a write window of the protected Test_Expected
'   sheet.
'
' WHY THIS EXISTS
'   Test_Expected is protected (#349), so the comparison needs a write window,
'   and the window must be closed
'   whatever the comparison does.
'
' INPUTS
'   total
'     Receives the number of items compared
'
' RETURNS
'   The number of items that differ.
'
' BEHAVIOR
'   Opens a write window, runs CompareExpected holding any error back, closes
'   the window and raises the held
'   error again.
'
' ERROR POLICY
'   An error from the comparison is raised again with its number, source and
'   description once the write window
'   is closed.
'
' DEPENDENCIES
'   BeginWrite, EndWrite (CORE_Util), CompareExpected.
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
    Const PROC_NAME             As String = "TEST_Regression.ExecuteRegressionTest"

    Dim FailureCount            As Long          'Items that differ
    Dim ErrorNumber             As Long          'Error number held back
    Dim ErrorSource             As String        'Error source held back
    Dim ErrorDescription        As String        'Error description held back

'------------------------------------------------------------------------------
' OPEN THE WRITE WINDOW
'------------------------------------------------------------------------------
    'Allow writes to the protected sheet (#349)
        BeginWrite

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Compare, holding any error back
        On Error Resume Next
        FailureCount = CompareExpected(total)
    'Keep the error that was met
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description

'------------------------------------------------------------------------------
' CLOSE THE WRITE WINDOW
'------------------------------------------------------------------------------
    'Close the window and raise a held-back error
        On Error GoTo 0
        EndWrite
        If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of items that differ
        ExecuteRegressionTest = FailureCount

End Function

Private Function CompareExpected(ByRef total As Long) As Long
'
'------------------------------------------------------------------------------
'                       COMPARE WITH THE EXPECTED VALUES
'------------------------------------------------------------------------------
' PURPOSE
'   Compares each expected item with the output of the last run and marks it
'   PASS or FAIL.
'
' WHY THIS EXISTS
'   The expected values on Test_Expected are the regression oracle of the
'   prototype.
'
' INPUTS
'   total
'     Receives the number of items compared
'
' RETURNS
'   The number of items that differ.
'
' BEHAVIOR
'   Reads the [EXPECTED] table of Test_Expected row by row. A missing output
'   counts as a failure; a value passes
'   when it is within the absolute tolerance plus the relative tolerance of the
'   expected amount. Writes the
'   actual value, the absolute difference and the status of each row, and the
'   result line in the subtitle cell.
'
' ERROR POLICY
'   A missing output counts as FAIL; a missing sheet raises (CORE_Util).
'
' DEPENDENCIES
'   CORE_Util table helpers, ActualValue, MarkStatus, ABS_TOLERANCE,
'   REL_TOLERANCE.
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
    Const PROC_NAME             As String = "TEST_Regression.CompareExpected"

    Dim ExpectedSheet           As Object        'Test_Expected sheet
    Dim RowIndex                As Long          'Row being compared
    Dim FirstRow                As Long          'First data row of the table
    Dim FailureCount            As Long          'Items that differ
    Dim ExpectedItem            As String        'Item name of the row
    Dim ExpectedValue           As Double        'Expected amount
    Dim ActualResult            As Variant       'Output of the last run, or Empty
    Dim Difference              As Double        'Absolute difference

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Locate the table and start the counts
        Set ExpectedSheet = GetSheet(SH_TEST)
        FirstRow = FindTagRow(ExpectedSheet, "EXPECTED") + 2
        FailureCount = 0
        total = 0
        RowIndex = FirstRow

'------------------------------------------------------------------------------
' COMPARE ROWS
'------------------------------------------------------------------------------
    'Compare each row until the first blank item
        Do While Not IsBlank(TableCell(ExpectedSheet, RowIndex, 1).Value) And FirstRow > 2
            ExpectedItem = CStr(TableCell(ExpectedSheet, RowIndex, 1).Value)
            ExpectedValue = ToDouble(TableCell(ExpectedSheet, RowIndex, 2).Value, 0#)
            ActualResult = ActualValue(ExpectedItem)
            total = total + 1
            If IsEmpty(ActualResult) Then
                TableCell(ExpectedSheet, RowIndex, 3).Value = "missing"
                TableCell(ExpectedSheet, RowIndex, 4).Value = ""
                MarkStatus TableCell(ExpectedSheet, RowIndex, 5), False
                FailureCount = FailureCount + 1
            Else
                Difference = Abs(CDbl(ActualResult) - ExpectedValue)
                TableCell(ExpectedSheet, RowIndex, 3).Value = CDbl(ActualResult)
                TableCell(ExpectedSheet, RowIndex, 4).Value = Difference
                If Difference <= ABS_TOLERANCE + REL_TOLERANCE * Abs(ExpectedValue) Then
                    MarkStatus TableCell(ExpectedSheet, RowIndex, 5), True
                Else
                    MarkStatus TableCell(ExpectedSheet, RowIndex, 5), False
                    FailureCount = FailureCount + 1
                End If
            End If
            RowIndex = RowIndex + 1
        Loop

'------------------------------------------------------------------------------
' WRITE THE RESULT
'------------------------------------------------------------------------------
    'Write the result line in the subtitle cell
        If FailureCount = 0 And total > 0 Then
            TableCell(ExpectedSheet, LAYOUT_SUBTITLE_ROW, 1).Value = "Result: PASS (" & total & " items)"
            MarkStatus TableCell(ExpectedSheet, LAYOUT_SUBTITLE_ROW, 1), True
        Else
            TableCell(ExpectedSheet, LAYOUT_SUBTITLE_ROW, 1).Value = "Result: FAIL (" & FailureCount & " of " & _
                total & " items)"
            MarkStatus TableCell(ExpectedSheet, LAYOUT_SUBTITLE_ROW, 1), False
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of items that differ
        CompareExpected = FailureCount

End Function

Private Sub MarkStatus(ByVal cell As Object, ByVal passed As Boolean)
'
'------------------------------------------------------------------------------
'                              MARK A STATUS CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Writes PASS or FAIL in a cell and colors it.
'
' INPUTS
'   cell
'     The cell to mark
'
'   passed
'     True for PASS
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' NOTES
'   A cell in the first table column gets the color only, without the text (the
'   result line).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' MARK
'------------------------------------------------------------------------------
    'Write the status, then color the cell, ignoring errors
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
'
'------------------------------------------------------------------------------
'                            READ THE ACTUAL VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the actual value of an expected item from the output sheets.
'
' WHY THIS EXISTS
'   Items are Out_Summary keys (for example TOTAL) or
'   RiskClass|Component|Scenario keys of Out_RiskClass (for
'   example GIRR|DELTA|MEDIUM).
'
' INPUTS
'   item
'     The item name
'
' RETURNS
'   The value, or Empty when the output does not have it.
'
' BEHAVIOR
'   An item with three parts separated by a vertical bar is looked up in
'   Out_RiskClass; any other item is
'   read from Out_Summary.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   GetSheet, TableCell, IsBlank, IsNumber, ToDouble (CORE_Util);
'   OutSummaryValue (M_Main).
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
    Const PROC_NAME             As String = "TEST_Regression.ActualValue"

    Dim KeyParts                As Variant       'Parts of the item name
    Dim OutputSheet             As Object        'Out_RiskClass sheet
    Dim RowIndex                As Long          'Row being read
    Dim SummaryResult           As Variant       'Out_Summary value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return Empty unless the item is found
        ActualValue = Empty

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'Look up a three-part item in Out_RiskClass, any other item in Out_Summary
        KeyParts = Split(item, "|")
        If UBound(KeyParts) = 2 Then
            Set OutputSheet = GetSheet(SH_OUT_CLASS)
            RowIndex = LAYOUT_FIRST_DATA_ROW
            Do While Not IsBlank(TableCell(OutputSheet, RowIndex, 1).Value)
                If CStr(TableCell(OutputSheet, RowIndex, 1).Value) = KeyParts(0) And _
                    CStr(TableCell(OutputSheet, RowIndex, 2).Value) = KeyParts(1) _
                   And CStr(TableCell(OutputSheet, RowIndex, 3).Value) = KeyParts(2) Then
                    ActualValue = TableCell(OutputSheet, RowIndex, 4).Value
                    Exit Function
                End If
                RowIndex = RowIndex + 1
            Loop
        Else
            SummaryResult = OutSummaryValue(item)
            If IsNumber(SummaryResult) Then ActualValue = ToDouble(SummaryResult, 0#)
        End If

End Function
