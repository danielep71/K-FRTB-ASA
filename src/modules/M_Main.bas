Attribute VB_Name = "M_Main"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the macros behind the sheet buttons, runs the calculation with
'   Excel prepared for
'   speed, puts Excel back as it was, writes the Results and Out_Summary sheets
'   and reports the
'   outcome to the user.
'
' WHY THIS EXISTS
'   The sheet buttons and the ribbon need entry points that run the engines in
'   the right
'   order, with the Excel settings and the sheet protection handled in one
'   place.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunFRTB calculates SBM, DRC and RRAO and rewrites every output sheet;
'   RunFRTB_Silent does
'   the same without dialogs, for automation and tests; ClearOutputs empties
'   every output sheet
'   and the log, keeping their headers and a note, as in a new workbook;
'   LastRunStatus returns OK
'   or the error of the last run; OutSummaryValue returns a value of
'   Out_Summary by item name.
'
' ERROR POLICY
'   A run error is logged, stored in LastRunStatus, noted on Results and shown
'   in a message
'   box by RunFRTB; RunFRTB_Silent never shows a dialog. Every run first clears
'   the output
'   sheets, and a failed run clears them again, so it leaves neither earlier
'   nor partial
'   capital behind. A parameter that is blank or not a number is a run error.
'   Errors in the
'   input rows are not VBA errors: they are listed on the Log sheet.
'
' DEPENDENCIES
'   CORE_Params, CORE_SBM, CORE_DRC and CORE_RRAO for the calculation;
'   CORE_Util and
'   CORE_Config for sheets, layout, palette and the log; M_Setup for the
'   navigation toggle of
'   the ribbon (SetNavigationVisible).
'
' NOTES
'   The module also runs in LibreOffice Calc with VBA support for the automated
'   tests. It needs no
'   references beyond the defaults.
'
' PUBLIC SURFACE
'   RunFRTB, RunFRTB_Silent, ClearOutputs, LastRunStatus, OutSummaryValue.
'   Ribbon_OnLoad, Ribbon_Run, Ribbon_Clear, Ribbon_Test, Ribbon_IsIdle,
'   Ribbon_ToggleNav,
'   Ribbon_NavShown: the callbacks of the FRTB-SA ribbon tab
'   (src/ribbon/customUI14.xml,
'   docs/ribbon.md). Each calls one procedure and holds no logic.
'
' STATE OWNERSHIP
'   Each run captures the screen updating and calculation mode, changes them
'   for the run
'   and restores the captured values on success and failure. The module owns
'   the status of
'   the last run (mLastStatus), whether a run is busy (mBusy) and the ribbon
'   object that the
'   busy state invalidates (mRibbon).
'
' KNOWN DEVIATION
'   The sheet buttons call this module directly, and it writes the Results and
'   Out_Summary
'   sheets itself; there is no separate UI layer yet.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const XL_CALCULATION_MANUAL As Long = -4135 'xlCalculationManual

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mLastStatus As String   'OK or the error of the last run
    Private mBusy       As Boolean  'True while a run is in progress; the ribbon disables its commands
    Private mRibbon     As Object   'IRibbonUI of the FRTB-SA tab, set by Ribbon_OnLoad


'
'------------------------------------------------------------------------------
'
'                                BUTTON MACROS
'
'------------------------------------------------------------------------------
'

Public Sub RunFRTB()
'
'------------------------------------------------------------------------------
'                             RUN THE CALCULATION
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the whole calculation from the Run button and reports the outcome in a
'   message box.
'
' WHY THIS EXISTS
'   The button needs one entry point that runs, then tells the person what
'   happened.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs ExecuteRun. On success shows the total capital, the error and warning
'   counts and a pointer to
'   the Results and Log sheets; on failure shows the error of the run.
'
' ERROR POLICY
'   Does not raise: a run error is stored in LastRunStatus and shown in a
'   message box.
'
' DEPENDENCIES
'   ExecuteRun, OutSummaryValue, LogErrorCount, LogWarningCount, APP_NAME.
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
    Const PROC_NAME             As String = "M_Main.RunFRTB"

    Dim Succeeded               As Boolean       'True when the run finished without error

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run the calculation
        Succeeded = ExecuteRun()
        If Succeeded Then
            MsgBox "Calculation complete." & vbCrLf & _
                   "Total capital: " & Format(OutSummaryValue("TOTAL"), "#,##0") & vbCrLf & _
                   "Errors: " & LogErrorCount() & "   Warnings: " & LogWarningCount() & vbCrLf & _
                   "See the Results and Log sheets.", vbInformation, APP_NAME
        Else
            MsgBox "Calculation failed: " & mLastStatus, vbCritical, APP_NAME
        End If

End Sub

Public Sub RunFRTB_Silent()
'
'------------------------------------------------------------------------------
'                             RUN WITHOUT DIALOGS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the calculation without any dialog.
'
' WHY THIS EXISTS
'   Automation and the tests run the calculation and read the outcome with
'   LastRunStatus.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The outcome is read with LastRunStatus.
'
' BEHAVIOR
'   Runs ExecuteRun and ignores its result.
'
' ERROR POLICY
'   Does not raise and shows no dialog.
'
' DEPENDENCIES
'   ExecuteRun
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run the calculation and keep the status for LastRunStatus
        ExecuteRun

End Sub

Private Sub SetBusy(ByVal busy As Boolean)
'
'------------------------------------------------------------------------------
'                          MARK THE RUN BUSY OR IDLE
'------------------------------------------------------------------------------
' PURPOSE
'   Marks a run as busy or idle and has the ribbon read its enabled states
'   again.
'
' INPUTS
'   busy
'     True while a run is in progress
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   The ribbon refresh ignores errors (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' UPDATE STATE
'------------------------------------------------------------------------------
    'Store the state and refresh the ribbon, ignoring errors
        mBusy = busy
        On Error Resume Next
        If Not mRibbon Is Nothing Then mRibbon.Invalidate
        On Error GoTo 0

End Sub

Public Function LastRunStatus() As String
'
'------------------------------------------------------------------------------
'                          RETURN THE LAST RUN STATUS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns OK or the error of the last run.
'
' WHY THIS EXISTS
'   Callers and tests read the outcome of a run without a dialog.
'
' INPUTS
'   None
'
' RETURNS
'   The status text: OK, or Error followed by the number and description.
'
' BEHAVIOR
'   Returns the module's status.
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
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the status of the last run
        LastRunStatus = mLastStatus

End Function

Public Sub ClearOutputs()
'
'------------------------------------------------------------------------------
'                           CLEAR THE OUTPUT SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Empties every output sheet and the log, keeping their headers and a note,
'   as in a new workbook.
'
' WHY THIS EXISTS
'   The Clear outputs command and every run start from empty output sheets, so
'   a failed run leaves no
'   earlier or partial capital behind.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens a write window, clears the Results and Out_* sheets, then the Log,
'   restores the protection and
'   raises the error it met, if any.
'
' ERROR POLICY
'   Errors while clearing are held back until the write window is closed, then
'   raised again with their
'   number, source and description.
'
' DEPENDENCIES
'   BeginWrite, EndWrite, ClearTableArea, GetSheet, SheetExists,
'   WriteOutputPlaceholder (CORE_Util).
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
    Const PROC_NAME             As String = "M_Main.ClearOutputs"

    Dim ErrorNumber             As Long          'Error number held back
    Dim ErrorSource             As String        'Error source held back
    Dim ErrorDescription        As String        'Error description held back

'------------------------------------------------------------------------------
' OPEN THE WRITE WINDOW
'------------------------------------------------------------------------------
    'Allow writes to the protected sheets
        BeginWrite

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Clear the result sheets and the Log, holding any error back
        On Error Resume Next
        ClearResultSheets "No results yet: press Run FRTB-SA.", False
        If Err.Number = 0 And SheetExists(SH_LOG) Then
            ClearTableArea GetSheet(SH_LOG)
            WriteOutputPlaceholder GetSheet(SH_LOG)
        End If
    'Keep the error that was met
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description

'------------------------------------------------------------------------------
' CLOSE THE WRITE WINDOW
'------------------------------------------------------------------------------
    'Close the write window and raise a held-back error
        On Error GoTo 0
        EndWrite
        If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Private Sub ClearResultSheets(ByVal resultsNote As String, ByVal isAlert As Boolean)
'
'------------------------------------------------------------------------------
'                           CLEAR THE RESULT SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Results and the Out_* sheets, but not the Log, to their empty
'   state: headers and a note.
'
' WHY THIS EXISTS
'   A failed run drops the rows it wrote, and the Clear outputs command starts
'   from empty sheets.
'
' INPUTS
'   resultsNote
'     Note shown on Results
'
'   isAlert
'     True to show the note as an alert
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   For each of Out_RiskClass, Out_Buckets, Out_DRC, Out_Summary and Results
'   that exists, clears the table
'   area and writes the placeholder; Results gets resultsNote.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   SheetExists, GetSheet, ClearTableArea, WriteOutputPlaceholder (CORE_Util).
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
    Const PROC_NAME             As String = "M_Main.ClearResultSheets"

    Dim SheetNames              As Variant       'Names of the sheets to clear
    Dim SheetIndex              As Long          'Name being read
    Dim TargetSheet             As Object        'Sheet being cleared

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Clear each output sheet that exists and write its placeholder
        SheetNames = Array(SH_OUT_CLASS, SH_OUT_BUCKETS, SH_OUT_DRC, SH_OUT_SUMMARY, SH_RESULTS)
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            If SheetExists(CStr(SheetNames(SheetIndex))) Then
                Set TargetSheet = GetSheet(CStr(SheetNames(SheetIndex)))
                ClearTableArea TargetSheet
                If TargetSheet.Name = SH_RESULTS Then
                    WriteOutputPlaceholder TargetSheet, resultsNote, isAlert
                Else
                    WriteOutputPlaceholder TargetSheet
                End If
            End If
        Next SheetIndex

End Sub


'
'------------------------------------------------------------------------------
'
'                           RIBBON CALLBACKS (#353)
'
'------------------------------------------------------------------------------
'

Public Sub Ribbon_OnLoad(ByVal ribbon As Object)
'
'------------------------------------------------------------------------------
'                            KEEP THE RIBBON OBJECT
'------------------------------------------------------------------------------
' PURPOSE
'   Keeps the ribbon object so that the module can refresh it.
'
' WHY THIS EXISTS
'   The FRTB-SA ribbon tab (src/ribbon/customUI14.xml, docs/ribbon.md) calls
'   this procedure by name; it holds no logic.
'
' INPUTS
'   control, ribbon, returnedVal and pressed are the arguments the ribbon
'   passes.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the ribbon object.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   See the module header.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CALLBACK
'------------------------------------------------------------------------------
    'Do what the command asks
        Set mRibbon = ribbon

End Sub

Public Sub Ribbon_Run(ByVal control As Object)
'
'------------------------------------------------------------------------------
'                                 RIBBON: RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Ribbon callback of the Run command.
'
' WHY THIS EXISTS
'   The FRTB-SA ribbon tab (src/ribbon/customUI14.xml, docs/ribbon.md) calls
'   this procedure by name; it holds no logic.
'
' INPUTS
'   control
'     The ribbon control
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls RunFRTB.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   See the module header.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CALLBACK
'------------------------------------------------------------------------------
    'Do what the command asks
        RunFRTB

End Sub

Public Sub Ribbon_Clear(ByVal control As Object)
'
'------------------------------------------------------------------------------
'                            RIBBON: CLEAR OUTPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Ribbon callback of the Clear outputs command.
'
' WHY THIS EXISTS
'   The FRTB-SA ribbon tab (src/ribbon/customUI14.xml, docs/ribbon.md) calls
'   this procedure by name; it holds no logic.
'
' INPUTS
'   control, ribbon, returnedVal and pressed are the arguments the ribbon
'   passes.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls ClearOutputs.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   See the module header.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CALLBACK
'------------------------------------------------------------------------------
    'Do what the command asks
        ClearOutputs

End Sub

Public Sub Ribbon_Test(ByVal control As Object)
'
'------------------------------------------------------------------------------
'                             RIBBON: RUN THE TEST
'------------------------------------------------------------------------------
' PURPOSE
'   Ribbon callback of the Test command.
'
' WHY THIS EXISTS
'   The FRTB-SA ribbon tab (src/ribbon/customUI14.xml, docs/ribbon.md) calls
'   this procedure by name; it holds no logic.
'
' INPUTS
'   control, ribbon, returnedVal and pressed are the arguments the ribbon
'   passes.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the regression test through Application.Run.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   See the module header.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CALLBACK
'------------------------------------------------------------------------------
    'Do what the command asks
        Application.Run "Test_RunRegression"

End Sub

Public Sub Ribbon_IsIdle(ByVal control As Object, ByRef returnedVal As Variant)
'
'------------------------------------------------------------------------------
'                            RIBBON: ENABLED STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Ribbon callback that enables a command only when no run is in progress.
'
' WHY THIS EXISTS
'   The FRTB-SA ribbon tab (src/ribbon/customUI14.xml, docs/ribbon.md) calls
'   this procedure by name; it holds no logic.
'
' INPUTS
'   control, ribbon, returnedVal and pressed are the arguments the ribbon
'   passes.
'
' RETURNS
'   Nothing. The value goes back through returnedVal.
'
' BEHAVIOR
'   Returns True when no run is busy.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   See the module header.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CALLBACK
'------------------------------------------------------------------------------
    'Do what the command asks
        returnedVal = Not mBusy

End Sub

Public Sub Ribbon_ToggleNav(ByVal control As Object, ByVal pressed As Boolean)
'
'------------------------------------------------------------------------------
'                      RIBBON: TOGGLE THE NAVIGATION BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Ribbon callback of the navigation bar toggle.
'
' WHY THIS EXISTS
'   The FRTB-SA ribbon tab (src/ribbon/customUI14.xml, docs/ribbon.md) calls
'   this procedure by name; it holds no logic.
'
' INPUTS
'   control, ribbon, returnedVal and pressed are the arguments the ribbon
'   passes.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Shows or hides the navigation bar.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   See the module header.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CALLBACK
'------------------------------------------------------------------------------
    'Do what the command asks
        SetNavigationVisible pressed

End Sub

Public Sub Ribbon_NavShown(ByVal control As Object, ByRef returnedVal As Variant)
'
'------------------------------------------------------------------------------
'                         RIBBON: NAVIGATION BAR STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Ribbon callback that reports whether the navigation bar is shown.
'
' WHY THIS EXISTS
'   The FRTB-SA ribbon tab (src/ribbon/customUI14.xml, docs/ribbon.md) calls
'   this procedure by name; it holds no logic.
'
' INPUTS
'   control, ribbon, returnedVal and pressed are the arguments the ribbon
'   passes.
'
' RETURNS
'   Nothing. The value goes back through returnedVal.
'
' BEHAVIOR
'   Returns the navigation bar state.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   See the module header.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CALLBACK
'------------------------------------------------------------------------------
    'Do what the command asks
        returnedVal = NavigationVisible()

End Sub


'
'------------------------------------------------------------------------------
'
'                                     RUN
'
'------------------------------------------------------------------------------
'

Private Function ExecuteRun() As Boolean
'
'------------------------------------------------------------------------------
'                                EXECUTE A RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Runs SBM, DRC and RRAO with Excel prepared for speed, and puts Excel back
'   as it was.
'
' WHY THIS EXISTS
'   RunFRTB, RunFRTB_Silent and the tests share one run, with one place that
'   restores the Excel settings
'   and the protection after a success and after a failure.
'
' INPUTS
'   None
'
' RETURNS
'   True when the run finished without error.
'
' BEHAVIOR
'   In order:
'     - marks the run busy and captures the screen updating and calculation
'     settings
'     - opens a write window and clears the outputs and the log
'     - loads the parameters, runs SBM, DRC and RRAO and writes the summary
'     - on an error, stores it in the status, logs it and clears the result
'     sheets again
'     - closes the write window, restores the settings and marks the run idle
'
' ERROR POLICY
'   Does not raise. The error is stored in the status and logged; the cleanup
'   ignores its own errors.
'
' DEPENDENCIES
'   CORE_Params, CORE_SBM, CORE_DRC, CORE_RRAO and CORE_Util, and the helpers
'   of this module for the Excel
'   settings, the busy state, the clearing of the outputs and the summary.
'
' NOTES
'   The error handler resumes at CleanUp, the label that the success path falls
'   into.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Main.ExecuteRun"

    Dim SavedScreen             As Boolean       'ScreenUpdating before the run
    Dim SavedCalculation        As Long          'Calculation mode before the run
    Dim Captured                As Boolean       'True once both were read
    Dim Opened                  As Boolean       'True once the write window is open
    Dim ErrorNumber             As Long          'Primary error; 0 for none
    Dim ErrorDescription        As String        'Description of the primary error

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route errors to the handler
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Mark the run busy and capture the Excel settings
        SetBusy True
        CaptureSettings SavedScreen, SavedCalculation, Captured
    'Open the write window; the sheets are protected (#349)
        BeginWrite
        Opened = True
    'Clear the outputs first, so that no earlier results survive a failure
        ClearOutputs
    'Log the start, load the parameters and run the three engines
        ResetLog
        LogMessage "INFO", "RunFRTB", APP_NAME & " " & APP_VERSION & " started"
        LoadAllParameters
        RunSBM
        RunDRC
        RunRRAO
    'Summarize the capital and record the status
        WriteSummary
        LogMessage "INFO", "RunFRTB", "Finished with " & LogErrorCount() & " error(s) and " & LogWarningCount() & _
            " warning(s)"
        mLastStatus = "OK"

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
CleanUp:
    'Store and log an error, then drop the rows written before it
        On Error Resume Next
        If ErrorNumber <> 0 Then
            mLastStatus = "Error " & ErrorNumber & ": " & ErrorDescription
            LogMessage "ERROR", "RunFRTB", mLastStatus
            ClearResultSheets "Run failed (see the Log sheet): " & ErrorDescription, True
        End If
        If Opened Then EndWrite
        If Captured Then RestoreSettings SavedScreen, SavedCalculation
        SetBusy False
        On Error GoTo 0

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Return the outcome
        ExecuteRun = (ErrorNumber = 0)
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Capture the error and resume at the cleanup
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
        Resume CleanUp

End Function

Private Sub CaptureSettings(ByRef savedScreen As Boolean, ByRef savedCalculation As Long, ByRef captured As Boolean)
'
'------------------------------------------------------------------------------
'                          CAPTURE THE EXCEL SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the screen updating and calculation settings, then turns screen
'   updating off and calculation to manual.
'
' INPUTS
'   savedScreen, savedCalculation
'     Receive the settings before the change
'
'   captured
'     Receives True when both settings were read
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Best effort: a host without these properties leaves captured False.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
    'Read the settings, then prepare Excel for the run; ignore errors
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
'
'------------------------------------------------------------------------------
'                          RESTORE THE EXCEL SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Puts the calculation mode and the screen updating back to the captured
'   values.
'
' INPUTS
'   savedScreen, savedCalculation
'     The values captured before the run
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore the settings, ignoring errors
        On Error Resume Next
        Application.Calculation = savedCalculation
        Application.ScreenUpdating = savedScreen
        On Error GoTo 0

End Sub


'
'------------------------------------------------------------------------------
'
'                                   REPORTS
'
'------------------------------------------------------------------------------
'

Private Sub WriteSummary()
'
'------------------------------------------------------------------------------
'                              WRITE THE SUMMARY
'------------------------------------------------------------------------------
' PURPOSE
'   Writes Out_Summary: the SBM total of each scenario, the binding scenario,
'   the SBM, DRC and RRAO capital, the total and the error and warning counts.
'
' WHY THIS EXISTS
'   The summary is the table that OutSummaryValue and the tests read, and it
'   feeds the Results sheet.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Finds the scenario with the highest SBM total, writes the ten summary rows,
'   formats them and calls
'   WriteResultsSheet.
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   SbmScenarioTotal (CORE_SBM), DrcTotal (CORE_DRC), RraoTotal (CORE_RRAO),
'   CORE_Util table helpers, PutSummary,
'   WriteResultsSheet.
'
' NOTES
'   On a tie between scenarios the first one in the order low, medium, high is
'   the binding one.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Main.WriteSummary"

    Dim SummarySheet            As Object        'Out_Summary sheet
    Dim ScenarioIndex           As Long          'Scenario being read
    Dim BestScenario            As Long          'Scenario with the highest SBM total
    Dim SbmValue                As Double        'SBM capital
    Dim DrcValue                As Double        'DRC capital
    Dim RraoValue               As Double        'RRAO capital

'------------------------------------------------------------------------------
' FIND THE BINDING SCENARIO
'------------------------------------------------------------------------------
    'Keep the scenario with the highest SBM total; a tie keeps the first
        BestScenario = SCEN_LOW
        For ScenarioIndex = SCEN_LOW To SCEN_HIGH
            If SbmScenarioTotal(ScenarioIndex) > SbmScenarioTotal(BestScenario) Then BestScenario = ScenarioIndex
        Next ScenarioIndex

'------------------------------------------------------------------------------
' READ THE CAPITAL
'------------------------------------------------------------------------------
    'Take the SBM capital of that scenario, the DRC and the RRAO
        SbmValue = SbmScenarioTotal(BestScenario)
        DrcValue = DrcTotal()
        RraoValue = RraoTotal()

'------------------------------------------------------------------------------
' WRITE THE TABLE
'------------------------------------------------------------------------------
    'Clear the sheet and write the header
        Set SummarySheet = GetSheet(SH_OUT_SUMMARY)
        ClearTableArea SummarySheet
        WriteTableHeader SummarySheet, LAYOUT_HEADER_ROW, Array("Item", "Value"), COLOR_OUTPUT_HEADER
    'Write the ten summary rows
        PutSummary SummarySheet, 1, "SBM_LOW", SbmScenarioTotal(SCEN_LOW)
        PutSummary SummarySheet, 2, "SBM_MEDIUM", SbmScenarioTotal(SCEN_MEDIUM)
        PutSummary SummarySheet, 3, "SBM_HIGH", SbmScenarioTotal(SCEN_HIGH)
        PutSummary SummarySheet, 4, "SBM_SCENARIO", ScenarioName(BestScenario)
        PutSummary SummarySheet, 5, "SBM", SbmValue
        PutSummary SummarySheet, 6, "DRC_NONSEC", DrcValue
        PutSummary SummarySheet, 7, "RRAO", RraoValue
        PutSummary SummarySheet, 8, "TOTAL", SbmValue + DrcValue + RraoValue
        PutSummary SummarySheet, 9, "ERRORS", LogErrorCount()
        PutSummary SummarySheet, 10, "WARNINGS", LogWarningCount()
    'Format the rows
        FormatDataRows SummarySheet, LAYOUT_FIRST_DATA_ROW, LAYOUT_HEADER_ROW + 10, 2, 2, "#,##0", False
    'Align the scenario name with the numbers (#206)
        On Error Resume Next
        TableCell(SummarySheet, LAYOUT_HEADER_ROW + 4, 2).HorizontalAlignment = -4152
        On Error GoTo 0

'------------------------------------------------------------------------------
' WRITE THE RESULTS SHEET
'------------------------------------------------------------------------------
    'Write the Results sheet from the same figures
        WriteResultsSheet BestScenario, SbmValue, DrcValue, RraoValue

End Sub

Private Sub PutSummary(ByVal ws As Object, ByVal i As Long, ByVal item As String, ByVal value As Variant)
'
'------------------------------------------------------------------------------
'                             WRITE A SUMMARY ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one item name and its value to a row of Out_Summary.
'
' INPUTS
'   ws
'     Out_Summary sheet
'
'   i
'     Row number below the header
'
'   item, value
'     The item name and its value
'
' RETURNS
'   Nothing
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Write the item and its value
        TableCell(ws, LAYOUT_HEADER_ROW + i, 1).Value = item
        TableCell(ws, LAYOUT_HEADER_ROW + i, 2).Value = value

End Sub

Private Sub WriteResultsSheet(ByVal best As Long, ByVal sbm As Double, ByVal drc As Double, ByVal rrao As Double)
'
'------------------------------------------------------------------------------
'                           WRITE THE RESULTS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Writes Results: the capital by risk class, component and scenario, the
'   capital requirement block, the run note and, on a scenario tie, the tie
'   note.
'
' WHY THIS EXISTS
'   The Results sheet is the page that a person reads after a run.
'
' INPUTS
'   best
'     Binding scenario
'
'   sbm, drc, rrao
'     The SBM, DRC and RRAO capital
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes one row per risk class and component, the SBM total of each
'   scenario, the capital requirement
'   block with its heading, the time of the run with the error and warning
'   counts and, when scenarios tie,
'   the tie note.
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   SbmCapital, SbmScenarioTotal (CORE_SBM), RiskClassCode, ComponentName,
'   ScenarioName (CORE_Config),
'   CORE_Util table helpers, WriteKeyRow, WriteSectionHeading, StyleTotalRow,
'   StyleNote.
'
' NOTES
'   The binding scenario's total is in navy, because red would read as an error
'   (#206).
'
'   Ties resolve in the order low, medium, high: a reporting convention, not a
'   reading of MAR21.7, since
'   the capital is the same (#29, item 5).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Main.WriteResultsSheet"

    Dim ResultsSheet            As Object        'Results sheet
    Dim RowIndex                As Long          'Row being written
    Dim ClassIndex              As Long          'Risk class being written
    Dim ComponentIndex          As Long          'Component being written
    Dim ScenarioIndex           As Long          'Scenario being written
    Dim BlockRow                As Long          'Heading row of the requirement block
    Dim TiedNames               As String        'Names of the tied scenarios
    Dim TiedCount               As Long          'Number of tied scenarios

'------------------------------------------------------------------------------
' PREPARE THE SHEET
'------------------------------------------------------------------------------
    'Clear the sheet and write the header
        Set ResultsSheet = GetSheet(SH_RESULTS)
        ClearTableArea ResultsSheet
        WriteTableHeader ResultsSheet, LAYOUT_HEADER_ROW, Array("Risk class", "Component", "Low", "Medium", "High"), _
            COLOR_OUTPUT_HEADER

'------------------------------------------------------------------------------
' WRITE THE CAPITAL TABLE
'------------------------------------------------------------------------------
    'Write one row per risk class and component
        RowIndex = LAYOUT_HEADER_ROW
        For ClassIndex = 1 To RC_COUNT
            For ComponentIndex = 1 To COMP_COUNT
                RowIndex = RowIndex + 1
                TableCell(ResultsSheet, RowIndex, 1).Value = RiskClassCode(ClassIndex)
                TableCell(ResultsSheet, RowIndex, 2).Value = ComponentName(ComponentIndex)
                For ScenarioIndex = SCEN_LOW To SCEN_HIGH
                    TableCell(ResultsSheet, RowIndex, 2 + ScenarioIndex).Value = SbmCapital(ClassIndex, _
                        ComponentIndex, ScenarioIndex)
                Next ScenarioIndex
            Next ComponentIndex
        Next ClassIndex
    'Write the SBM total of each scenario
        RowIndex = RowIndex + 1
        TableCell(ResultsSheet, RowIndex, 1).Value = "SBM total"
        For ScenarioIndex = SCEN_LOW To SCEN_HIGH
            TableCell(ResultsSheet, RowIndex, 2 + ScenarioIndex).Value = SbmScenarioTotal(ScenarioIndex)
        Next ScenarioIndex
    'Format the table and its total row
        FormatDataRows ResultsSheet, LAYOUT_FIRST_DATA_ROW, RowIndex, 5, 3, "#,##0", False
        StyleTotalRow ResultsSheet, RowIndex, 5
    'Show the binding scenario's total in navy (#206)
        On Error Resume Next
        TableCell(ResultsSheet, RowIndex, 2 + best).Font.Color = COLOR_INPUT_HEADER
        On Error GoTo 0

'------------------------------------------------------------------------------
' WRITE THE REQUIREMENT BLOCK
'------------------------------------------------------------------------------
    'Write the capital requirement rows
        BlockRow = RowIndex + 2
        WriteKeyRow ResultsSheet, BlockRow + 1, "Binding correlation scenario", ScenarioName(best)
        WriteKeyRow ResultsSheet, BlockRow + 2, "SBM capital (max over scenarios)", sbm
        WriteKeyRow ResultsSheet, BlockRow + 3, "DRC (non-securitization)", drc
        WriteKeyRow ResultsSheet, BlockRow + 4, "Residual risk add-on", rrao
        WriteKeyRow ResultsSheet, BlockRow + 5, "Total capital requirement", sbm + drc + rrao
    'Format the rows
        FormatDataRows ResultsSheet, BlockRow + 1, BlockRow + 5, 3, 3, "#,##0", False
        StyleTotalRow ResultsSheet, BlockRow + 5, 3
    'Add the heading after the rows, so that their borders do not replace its rule (#206)
        WriteSectionHeading ResultsSheet, BlockRow, "Capital requirement", 3

'------------------------------------------------------------------------------
' WRITE THE NOTES
'------------------------------------------------------------------------------
    'Write the run note
        TableCell(ResultsSheet, BlockRow + 7, 1).Value = "Last run " & Format(Now, "yyyy-mm-dd hh:mm:ss") & _
            " - " & LogErrorCount() & " error(s), " & LogWarningCount() & " warning(s) on Log"
        StyleNote TableCell(ResultsSheet, BlockRow + 7, 1), (LogErrorCount() > 0)
    'List the scenarios that share the binding total
        For ScenarioIndex = SCEN_LOW To SCEN_HIGH
            If SbmScenarioTotal(ScenarioIndex) = SbmScenarioTotal(best) Then
                TiedCount = TiedCount + 1
                If TiedCount > 1 Then TiedNames = TiedNames & ", "
                TiedNames = TiedNames & ScenarioName(ScenarioIndex)
            End If
        Next ScenarioIndex
    'Write the tie note when several scenarios tie
        If TiedCount > 1 Then
            TableCell(ResultsSheet, BlockRow + 8, 1).Value = "Scenario tie (" & TiedNames & "): same SBM total, " & _
                ScenarioName(best) & " reported; ties resolve in the order low, medium, high"
            StyleNote TableCell(ResultsSheet, BlockRow + 8, 1), False
        End If

End Sub

Private Sub WriteKeyRow(ByVal ws As Object, ByVal r As Long, ByVal label As String, ByVal value As Variant)
'
'------------------------------------------------------------------------------
'                               WRITE A KEY ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a label and a value of the capital requirement block and merges the
'   label cells.
'
' INPUTS
'   ws, r
'     Sheet and row
'
'   label, value
'     The label and its value
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' NOTES
'   A scenario name aligns with the amounts (#206).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Write the label and the value
        TableCell(ws, r, 1).Value = label
        TableCell(ws, r, 3).Value = value
    'Align, merge and embolden, ignoring errors
        On Error Resume Next
        TableCell(ws, r, 3).HorizontalAlignment = -4152
        ws.Range(TableCell(ws, r, 1), TableCell(ws, r, 2)).Merge
        TableCell(ws, r, 1).Font.Bold = True
        On Error GoTo 0

End Sub

Private Sub WriteSectionHeading(ByVal ws As Object, ByVal r As Long, ByVal text As String, ByVal nCols As Long)
'
'------------------------------------------------------------------------------
'                           WRITE A SECTION HEADING
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a heading with a rule under it, across a number of columns.
'
' INPUTS
'   ws, r
'     Sheet and row
'
'   text
'     The heading
'
'   nCols
'     Number of columns the rule spans
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Write the heading text
        TableCell(ws, r, 1).Value = text
    'Format the text and draw the rule, ignoring errors
        On Error Resume Next
        With TableCell(ws, r, 1).Font
            .Bold = True
            .Size = 14
            .Color = COLOR_INPUT_HEADER
        End With
        With ws.Range(TableCell(ws, r, 1), TableCell(ws, r, nCols)).Borders(9)
            .LineStyle = 1
            .Weight = -4138
            .Color = COLOR_INPUT_HEADER
        End With
        ws.Rows(r).RowHeight = 30
        On Error GoTo 0

End Sub

Private Sub StyleTotalRow(ByVal ws As Object, ByVal r As Long, ByVal nCols As Long)
'
'------------------------------------------------------------------------------
'                              STYLE A TOTAL ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Makes a row bold on the total fill.
'
' INPUTS
'   ws, r
'     Sheet and row
'
'   nCols
'     Number of columns to style
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Style the row, ignoring errors
        On Error Resume Next
        With ws.Range(TableCell(ws, r, 1), TableCell(ws, r, nCols))
            .Font.Bold = True
            .Interior.Color = COLOR_TOTAL_FILL
        End With
        On Error GoTo 0

End Sub

Private Sub StyleNote(ByVal cell As Object, ByVal isAlert As Boolean)
'
'------------------------------------------------------------------------------
'                                 STYLE A NOTE
'------------------------------------------------------------------------------
' PURPOSE
'   Formats a note cell in italics, as an alert or not.
'
' INPUTS
'   cell
'     The note cell
'
'   isAlert
'     True for the alert colors
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Format the note, ignoring errors
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
'
'------------------------------------------------------------------------------
'                             READ A SUMMARY VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the value of an Out_Summary item by name.
'
' WHY THIS EXISTS
'   Callers and tests read the capital figures by item name and not by row.
'
' INPUTS
'   item
'     Item name, compared without regard to case
'
' RETURNS
'   The value, or Empty when the item is not found.
'
' BEHAVIOR
'   Walks the rows of Out_Summary until the first blank item and returns the
'   value of the matching one.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   GetSheet, TableCell, IsBlank (CORE_Util).
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
    Const PROC_NAME             As String = "M_Main.OutSummaryValue"

    Dim SummarySheet            As Object        'Out_Summary sheet
    Dim RowIndex                As Long          'Row being read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Locate the sheet and return Empty unless the item is found
        Set SummarySheet = GetSheet(SH_OUT_SUMMARY)
        OutSummaryValue = Empty
        RowIndex = LAYOUT_FIRST_DATA_ROW

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Search the items for the name
        Do While Not IsBlank(TableCell(SummarySheet, RowIndex, 1).Value)
            If UCase$(CStr(TableCell(SummarySheet, RowIndex, 1).Value)) = UCase$(item) Then
                OutSummaryValue = TableCell(SummarySheet, RowIndex, 2).Value
                Exit Function
            End If
            RowIndex = RowIndex + 1
        Loop

End Function
