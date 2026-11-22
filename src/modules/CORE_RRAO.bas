Attribute VB_Name = "CORE_RRAO"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_RRAO
'------------------------------------------------------------------------------
' PURPOSE
'   Calculates the residual risk add-on (RRAO): the gross notional of each row of
'   In_RRAO times the risk weight of its residual risk category (exotic or other),
'   summed.
'
' WHY THIS EXISTS
'   The RRAO is a capital component of its own, next to the sensitivities-based
'   method and the default risk charge, so it has its own engine module.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Reads the rows of In_RRAO, logs the rows it rejects and keeps the total of the
'   last run.
'
' ERROR POLICY
'   Invalid rows are logged as ERROR and skipped.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, CORE_Params (the parameters must be loaded).
'
' NOTES
'   Reference: Basel Framework MAR23.8.
'
'   The module also runs in LibreOffice Calc with VBA support for the automated
'   tests. It needs no references beyond the defaults.
'
' PUBLIC SURFACE
'   RunRRAO; RraoTotal.
'
' STATE OWNERSHIP
'   Owns the total of the last run. Reads In_RRAO and logs rejected rows.
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mTotal  As Double   'RRAO of the last run

Public Function RraoTotal() As Double

'
'------------------------------------------------------------------------------
'                          RETURN THE LAST RRAO TOTAL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the RRAO total of the last RunRRAO call.
'
' WHY THIS EXISTS
'   The total stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The total, or 0 when RunRRAO has not run.
'
' BEHAVIOR
'   Returns the module's total.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the total of the last run
        RraoTotal = mTotal

End Function

Public Sub RunRRAO()

'
'------------------------------------------------------------------------------
'                              CALCULATE THE RRAO
'------------------------------------------------------------------------------
' PURPOSE
'   Calculates the RRAO of the rows of In_RRAO and stores the total.
'
' WHY THIS EXISTS
'   The run needs the add-on for the capital total, and a row that cannot be used
'   must show in the Log and not stop the run.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The total is read with RraoTotal.
'
' BEHAVIOR
'   Walks the rows of In_RRAO from the first data row to the first row whose first
'   two columns are blank. For each row:
'     - an unknown category is logged as ERROR and skipped
'     - a gross notional that is not a number is logged as ERROR and skipped
'     - a negative gross notional is logged as ERROR and skipped
'     - otherwise the risk weight of the category times the gross notional is
'       added to the total
'
' ERROR POLICY
'   Does not raise for a bad row. Errors from the callees (for example a missing
'   sheet) are not handled here.
'
' DEPENDENCIES
'   GetSheet, TableCell, IsBlank, IsNumber, ToText, ToDouble, LogMessage (CORE_Util);
'   RraoRiskWeight (CORE_Params).
'
' NOTES
'   The parameters must be loaded before the call.
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_RRAO.RunRRAO"

    Dim InputSheet              As Object       'In_RRAO sheet
    Dim RowIndex                As Long         'Row being read
    Dim Source                  As String       'Row reference for the Log
    Dim Category                As String       'Residual risk category of the row
    Dim RiskWeight              As Double       'Risk weight of the category
    Dim Found                   As Boolean      'True when the category is defined

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the total at zero
        mTotal = 0#
    'Locate the input sheet
        Set InputSheet = GetSheet(SH_IN_RRAO)
    'Start at the first data row
        RowIndex = LAYOUT_FIRST_DATA_ROW

'------------------------------------------------------------------------------
' PROCESS ROWS
'------------------------------------------------------------------------------
    'Read rows until the first two columns are both blank
        Do While Not (IsBlank(TableCell(InputSheet, RowIndex, 1).Value) And _
                IsBlank(TableCell(InputSheet, RowIndex, 2).Value))
            'Name the row for the Log
                Source = SH_IN_RRAO & " row " & RowIndex
            'Look up the risk weight of the category
                Category = ToText(TableCell(InputSheet, RowIndex, 2).Value)
                RiskWeight = RraoRiskWeight(Category, Found)
            'Log and skip a row that cannot be used, add the others
                If Not Found Then
                    LogMessage "ERROR", Source, "Unknown RRAO category '" & Category & "'; row skipped"
                ElseIf Not IsNumber(TableCell(InputSheet, RowIndex, 3).Value) Then
                    LogMessage "ERROR", Source, "Gross notional is not numeric; row skipped"
                ElseIf ToDouble(TableCell(InputSheet, RowIndex, 3).Value, 0#) < 0# Then
                    LogMessage "ERROR", Source, "Gross notional must be non-negative; row skipped"
                Else
                    mTotal = mTotal + RiskWeight * ToDouble(TableCell(InputSheet, RowIndex, 3).Value, 0#)
                End If
            'Move to the next row
                RowIndex = RowIndex + 1
        Loop

End Sub
