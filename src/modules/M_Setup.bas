Attribute VB_Name = "M_Setup"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_Setup
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the workbook from an empty one: every sheet with the K-FRTB-ASA frame (brand bar,
'   title bar, subtitle, tab color), the default regulatory parameters (Basel MAR21-MAR23),
'   a synthetic sample portfolio, the buttons and the expected results of the regression
'   test.
'
' WHY THIS EXISTS
'   The macro-free template src/workbook/FRTB_Template.xlsx is this build saved without VBA;
'   FinishWorkbook adds the buttons and reapplies the view settings to a workbook opened
'   from it.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Rebuilds the sheets, parameters, sample data and expected values, adds the sheet buttons
'   and the navigation bars, and sets the tab colors and the window settings.
'
' ERROR POLICY
'   Formatting is best effort: a host that rejects a format property keeps the content and
'   skips the format. Content writes raise normally.
'
' DEPENDENCIES
'   CORE_Config for names, layout and palette; CORE_Util for sheet access and table
'   formatting.
'
' NOTES
'   Excel VBA; also runs in LibreOffice Calc with VBA support for the automated tests. No
'   references beyond the defaults.
'
' PUBLIC SURFACE
'   BuildWorkbook: rebuild every sheet (overwrites the inputs), then FinishWorkbook.
'   FinishWorkbook: add the sheet buttons and window settings to a workbook opened from the
'   template.
'   AddNavigationBars: add the bar of links in column A of every sheet; Excel only, because
'   the bar is made of shapes (#347).
'   ResetParameters: rewrite the P_* sheets with the defaults.
'   LoadSampleData: rewrite the In_* sheets with the sample portfolio.
'   SetNavigationVisible, NavigationVisible: show, hide and read the navigation bar, for the
'   ribbon's toggle.
'   OpenWriteWindow, CloseWriteWindow: lift and restore the protection for a caller outside
'   the workbook, such as the build script.
'
'   Every public macro opens a write window (BeginWrite, EndWrite of CORE_Util) around its
'   work, because the sheets and the workbook structure are protected without a password
'   (#349); FinishWorkbook protects them last.
'
' STATE OWNERSHIP
'   Turns alerts, screen updating and printer communication off while building and back on
'   at the end, so Excel applies the page setup of every sheet in one batch instead of
'   querying the printer for each property (#295). Activates each sheet once to set
'   gridlines and frozen panes.
'
' REFERENCE
'   Parameter values: Basel Framework MAR21-MAR23 (in force 1 January 2023), checked against
'   the BIS text on 2026-10-17. Verify them against the regulation that applies to you
'   before relying on any result.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const NAV_COUNT     As Long = 23            'Entries of NavEntry: every managed sheet once (#347)

    'Input number formats. A format such as #,##0.#### ends a whole number with the decimal
    'separator in Excel ("5.000.000,"), so amounts show two decimals and years one to three (#206).
    'Declarations stay above the first procedure: Excel accepts only comments after End Sub (#327)
    Private Const FMT_AMOUNT    As String = "#,##0.00"  'Number format of amounts: two decimals
    Private Const FMT_YEARS     As String = "0.0##"     'Number format of years: one to three decimals


'
'------------------------------------------------------------------------------
'
'                       BUILD, FINISH AND RESET COMMANDS
'
'------------------------------------------------------------------------------
'

Public Sub BuildWorkbook()

'
'------------------------------------------------------------------------------
'                              BUILD THE WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Rebuilds every managed sheet from empty content and finishes the workbook.
'
' WHY THIS EXISTS
'   The workbook is built from code: the macro-free template is this build saved without
'   VBA.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens a write window, runs BuildAllSheets, closes the write window and raises again the
'   error that the build raised, if any.
'
' ERROR POLICY
'   The error of the build is captured, the write window is closed, then the same number,
'   source and description are raised again.
'
' DEPENDENCIES
'   BeginWrite, EndWrite (CORE_Util); BuildAllSheets.
'
' NOTES
'   Overwrites the input sheets (In_*) and the parameter sheets (P_*) with the sample
'   portfolio and the default parameters.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.BuildWorkbook"

    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorSource             As String       'Captured error source
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' RUN THE BUILD
'------------------------------------------------------------------------------
    'Open the write window
        BeginWrite
    'Run the build and keep the error it raises, if any
        On Error Resume Next
        BuildAllSheets
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' CLOSE AND RAISE
'------------------------------------------------------------------------------
    'Close the write window
        EndWrite
    'Raise the error of the build again
        If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Private Sub BuildAllSheets()

'
'------------------------------------------------------------------------------
'                              BUILD EVERY SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or clears every managed sheet, builds each of them and finishes the workbook.
'
' WHY THIS EXISTS
'   The build steps run in one fixed order here; BuildWorkbook adds the write window around
'   them.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - turns alerts and screen updating off
'     - adds each managed sheet that is missing and clears the others
'     - deletes the blank sheets that are not managed
'     - builds Home, the ReadMe sheet, Control, the output sheets and the input sheets
'     - writes the parameters, the sample data and the expected results
'     - turns alerts back on and runs FinishWorkbook
'
' ERROR POLICY
'   No handler of its own: an error leaves the procedure at once and is taken by the caller.
'
' DEPENDENCIES
'   ManagedSheets (CORE_Config); EnsureSheet, RemoveDefaultSheets, FinishWorkbook and the
'   Build* and Write* procedures of this module.
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
    Const PROC_NAME             As String = "M_Setup.BuildAllSheets"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim Index                   As Long         'Loop counter

'------------------------------------------------------------------------------
' PREPARE THE SHEETS
'------------------------------------------------------------------------------
    'Turn alerts and screen updating off
        SetAlerts False
    'Add each managed sheet that is missing, clear the others
        SheetNames = ManagedSheets()
        For Index = LBound(SheetNames) To UBound(SheetNames)
            EnsureSheet CStr(SheetNames(Index))
        Next Index
    'Delete the blank default sheets
        RemoveDefaultSheets

'------------------------------------------------------------------------------
' BUILD THE SHEETS
'------------------------------------------------------------------------------
    'Build the start, ReadMe, Control, output and input sheets
        BuildHome
        BuildReadme
        BuildControl
        BuildOutputSheets
        BuildInputSheets

'------------------------------------------------------------------------------
' WRITE THE CONTENT
'------------------------------------------------------------------------------
    'Write the default parameters, the sample portfolio and the expected results
        WriteParameters
        WriteSampleData
        WriteExpectedResults

'------------------------------------------------------------------------------
' FINISH
'------------------------------------------------------------------------------
    'Turn alerts and screen updating back on
        SetAlerts True
    'Add the buttons and apply the window settings
        FinishWorkbook

End Sub

Public Sub FinishWorkbook()

'
'------------------------------------------------------------------------------
'                             FINISH THE WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Adds the sheet buttons and reapplies the tab colors and the window settings to a
'   workbook opened from the template.
'
' WHY THIS EXISTS
'   The XLSX template retains the presentation settings, but the macro buttons are
'   deliberately excluded from it.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens a write window, runs FinishAllSheets, closes the write window and raises again the
'   error that the run raised, if any.
'
' ERROR POLICY
'   The error is captured, the write window is closed, then the same number, source and
'   description are raised again.
'
' DEPENDENCIES
'   BeginWrite, EndWrite (CORE_Util); FinishAllSheets.
'
' NOTES
'   BuildAllSheets calls it last, and the build script calls it for a workbook opened from
'   the template (tools/Build-Workbook.ps1).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.FinishWorkbook"

    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorSource             As String       'Captured error source
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' RUN THE FINISH
'------------------------------------------------------------------------------
    'Open the write window
        BeginWrite
    'Run the finish and keep the error it raises, if any
        On Error Resume Next
        FinishAllSheets
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' CLOSE AND RAISE
'------------------------------------------------------------------------------
    'Close the write window
        EndWrite
    'Raise the error of the run again
        If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Private Sub FinishAllSheets()

'
'------------------------------------------------------------------------------
'                          APPLY THE FINISHING STEPS
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the tab color to every managed sheet that exists, adds the buttons, applies the
'   window settings and activates Home.
'
' BEHAVIOR
'   Turns alerts and screen updating off while it works and back on at the end.
'
' ERROR POLICY
'   Activating Home is best effort (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.FinishAllSheets"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim Index                   As Long         'Loop counter

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Turn alerts and screen updating off
        SetAlerts False
    'Set the tab color of each managed sheet that exists
        SheetNames = ManagedSheets()
        For Index = LBound(SheetNames) To UBound(SheetNames)
            If SheetExists(CStr(SheetNames(Index))) Then ApplyTabColor GetSheet(CStr(SheetNames(Index)))
        Next Index
    'Add the buttons and apply the window settings
        AddWorkbookButtons
        ApplyWindowSettings
    'Activate Home; ignore a host that refuses
        On Error Resume Next
        ThisWorkbook.Worksheets(SH_HOME).Activate
        On Error GoTo 0
    'Turn alerts and screen updating back on
        SetAlerts True

End Sub

Public Sub ResetParameters()

'
'------------------------------------------------------------------------------
'                             RESET THE PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Rewrites the P_* sheets with the default regulatory parameters.
'
' WHY THIS EXISTS
'   The parameter sheets hold values that can be edited; this command writes the defaults
'   again.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens a write window, turns alerts off, runs WriteParameters, turns alerts on, closes
'   the write window and raises again the error of the run, if any.
'
' ERROR POLICY
'   The error is captured, alerts are turned on and the write window is closed, then the
'   same number, source and description are raised again.
'
' DEPENDENCIES
'   BeginWrite, EndWrite (CORE_Util); SetAlerts, WriteParameters.
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
    Const PROC_NAME             As String = "M_Setup.ResetParameters"

    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorSource             As String       'Captured error source
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' RUN THE RESET
'------------------------------------------------------------------------------
    'Open the write window
        BeginWrite
    'Rewrite the parameters with alerts off and keep the error, if any
        On Error Resume Next
        SetAlerts False
        WriteParameters
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description

'------------------------------------------------------------------------------
' RESTORE AND RAISE
'------------------------------------------------------------------------------
    'Turn alerts back on
        SetAlerts True
    'Restore normal error handling
        On Error GoTo 0
    'Close the write window
        EndWrite
    'Raise the error of the run again
        If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Public Sub LoadSampleData()

'
'------------------------------------------------------------------------------
'                          LOAD THE SAMPLE PORTFOLIO
'------------------------------------------------------------------------------
' PURPOSE
'   Rewrites the In_* sheets with the synthetic sample portfolio.
'
' WHY THIS EXISTS
'   The regression test compares the sample portfolio with its expected values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens a write window and turns alerts off. Rebuilds the input sheets and, when that
'   raised no error, writes the sample records. Then turns alerts on, closes the write
'   window and raises again the error of the run, if any.
'
' ERROR POLICY
'   The error is captured, alerts are turned on and the write window is closed, then the
'   same number, source and description are raised again.
'
' DEPENDENCIES
'   BeginWrite, EndWrite (CORE_Util); SetAlerts, BuildInputSheets, WriteSampleData.
'
' NOTES
'   Clearing an input sheet keeps the data validation of its input cells (#348).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.LoadSampleData"

    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorSource             As String       'Captured error source
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' RUN THE LOAD
'------------------------------------------------------------------------------
    'Open the write window
        BeginWrite
    'Rebuild the input sheets with alerts off, then write the samples when no error was raised
        On Error Resume Next
        SetAlerts False
        BuildInputSheets
        If Err.Number = 0 Then WriteSampleData
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description

'------------------------------------------------------------------------------
' RESTORE AND RAISE
'------------------------------------------------------------------------------
    'Turn alerts back on
        SetAlerts True
    'Restore normal error handling
        On Error GoTo 0
    'Close the write window
        EndWrite
    'Raise the error of the run again
        If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

'
'------------------------------------------------------------------------------
'
'                                SHEET HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub SetAlerts(ByVal enabled As Boolean)

'
'------------------------------------------------------------------------------
'                        SET ALERTS AND SCREEN UPDATING
'------------------------------------------------------------------------------
' PURPOSE
'   Turns Excel alerts, screen updating and printer communication on or off together.
'
' INPUTS
'   enabled
'     True turns the three settings on, False turns them off.
'
' BEHAVIOR
'   Sets Application.DisplayAlerts, ScreenUpdating and PrintCommunication to the value. Each
'   PageSetup change queries the default printer, which can wait for a timeout per property
'   when that printer is unreachable (#295).
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next): a host without a property skips it.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' SET THE STATE
'------------------------------------------------------------------------------
    'Ignore a host that lacks a property
        On Error Resume Next
    'Set alerts and screen updating
        Application.DisplayAlerts = enabled
        Application.ScreenUpdating = enabled
    'Each PageSetup change queries the default printer, which can wait for a timeout per property
    'when that printer is unreachable (#295); hosts without the property ignore it
        Application.PrintCommunication = enabled
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub ClearCells(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                          CLEAR THE CELLS OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the contents and the formats of every cell of a sheet.
'
' WHY THIS EXISTS
'   Clear would also remove the data validation of the input cells (#348), so contents and
'   formats are cleared and the validation stays.
'
' INPUTS
'   ws
'     Sheet to empty
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Clear the contents and then the formats
        ws.Cells.ClearContents
        ws.Cells.ClearFormats

End Sub

Private Sub EnsureSheet(ByVal sheetName As String)

'
'------------------------------------------------------------------------------
'                            ENSURE A SHEET EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Makes sure a worksheet with the given name exists: an existing one is cleared, a missing
'   one is added.
'
' INPUTS
'   sheetName
'     Name of the worksheet
'
' BEHAVIOR
'   A new worksheet is added after the last worksheet of this workbook and named.
'
' ERROR POLICY
'   None handled here: an error from the host is raised.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.EnsureSheet"

    Dim NewSheet                As Object       'Worksheet that is added
    Dim Book                    As Object       'This workbook

'------------------------------------------------------------------------------
' ENSURE THE SHEET
'------------------------------------------------------------------------------
    'Refer to this workbook
        Set Book = ThisWorkbook
    'Clear the sheet when it exists, otherwise add it after the last sheet and name it
        If SheetExists(sheetName) Then
            Book.Worksheets(sheetName).Cells.Clear
        Else
            Set NewSheet = Book.Worksheets.Add(After:=Book.Worksheets(Book.Worksheets.Count))
            NewSheet.Name = sheetName
        End If

End Sub

Private Function IsManaged(ByVal sheetName As String) As Boolean

'
'------------------------------------------------------------------------------
'                           TEST FOR A MANAGED SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a sheet name is in the list of managed sheets.
'
' INPUTS
'   sheetName
'     Name to look up. The comparison ignores case
'
' RETURNS
'   True when the name is in the list returned by ManagedSheets, otherwise False.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.IsManaged"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim Index                   As Long         'Loop counter

'------------------------------------------------------------------------------
' SEARCH THE LIST
'------------------------------------------------------------------------------
    'Read the managed sheet names
        SheetNames = ManagedSheets()
    'Assume the sheet is not managed
        IsManaged = False
    'Stop at the first name that matches, ignoring case
        For Index = LBound(SheetNames) To UBound(SheetNames)
            If UCase$(CStr(SheetNames(Index))) = UCase$(sheetName) Then
                IsManaged = True
                Exit Function
            End If
        Next Index

End Function

Private Sub RemoveDefaultSheets()

'
'------------------------------------------------------------------------------
'                       DELETE THE BLANK DEFAULT SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes the blank worksheets that are not part of the layout, such as the default Sheet1
'   of a new workbook.
'
' BEHAVIOR
'   Walks the worksheets from the last to the first. A sheet that is not managed and has A1,
'   A2 and B1 blank is deleted. Sheets with content are never deleted.
'
' ERROR POLICY
'   A sheet that cannot be deleted is skipped (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.RemoveDefaultSheets"

    Dim Index                   As Long         'Loop counter
    Dim CurrentSheet            As Object       'Worksheet being tested

'------------------------------------------------------------------------------
' DELETE BLANK SHEETS
'------------------------------------------------------------------------------
    'Walk the worksheets from the last to the first and delete the blank ones that are not managed
        For Index = ThisWorkbook.Worksheets.Count To 1 Step -1
            Set CurrentSheet = ThisWorkbook.Worksheets(Index)
            If Not IsManaged(CurrentSheet.Name) Then
                If IsBlank(CurrentSheet.Cells(1, 1).Value) And IsBlank(CurrentSheet.Cells(2, 1).Value) And _
                    IsBlank(CurrentSheet.Cells(1, 2).Value) Then
                    On Error Resume Next
                    CurrentSheet.Delete
                    On Error GoTo 0
                End If
            End If
        Next Index

End Sub

Private Sub ApplyFrame( _
    ByVal ws As Object, _
    ByVal title As String, _
    ByVal subtitle As String, _
    ByVal lastCol As Long)
'
'------------------------------------------------------------------------------
'                            APPLY THE SHEET FRAME
'------------------------------------------------------------------------------
' PURPOSE
'   Writes and formats the frame shared by every sheet: brand bar, title bar, subtitle and
'   spacer rows, with the page setup and the tab color.
'
' WHY THIS EXISTS
'   Every sheet of the workbook shows the same frame (the workbook frame).
'
' INPUTS
'   ws
'     Sheet to format
'
'   title
'     Sheet name that SheetTitle turns into the text of the title bar
'
'   subtitle
'     Text of the subtitle row
'
'   lastCol
'     Last column of the brand bar and the title bar
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - writes the product name, the title and the subtitle
'     - sets the font, the width of columns A and B and the black fill of column A, where
'       the navigation bar sits (#347)
'     - colors and sizes the brand bar (row 1), the title bar (row 2) and the subtitle (row
'       3)
'     - sets the heights of rows 1 to 4
'     - sets the page setup: A4 landscape, one page wide, margins, footer and the title rows
'       that repeat
'     - extends the print area to the frame and sets the tab color
'
' ERROR POLICY
'   Formatting is best effort (On Error Resume Next); the text writes raise normally.
'
' DEPENDENCIES
'   APP_NAME, layout and color constants (CORE_Config); SheetTitle, ApplyTabColor;
'   ExtendPrintArea (CORE_Util).
'
' NOTES
'   The title rows that repeat depend on the sheet: $1:$3 on README and Home, $1:$6 on
'   Control, Test_Expected and the P_* sheets, $1:$5 on the others.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE THE TEXTS
'------------------------------------------------------------------------------
    'Write the product name, the sheet title and the subtitle
        ws.Cells(1, 2).Value = APP_NAME
        ws.Cells(2, 2).Value = SheetTitle(title)
        ws.Cells(LAYOUT_SUBTITLE_ROW, LAYOUT_FIRST_COL).Value = subtitle

'------------------------------------------------------------------------------
' FORMAT THE FRAME
'------------------------------------------------------------------------------
    'Ignore formatting errors from here to the end of the formatting
        On Error Resume Next
    'Set the sheet font and center the cells vertically
        ws.Cells.Font.Name = FONT_NAME
        ws.Cells.Font.Size = 10
        ws.Cells.VerticalAlignment = -4108  'xlCenter
    'Set the width of column A and fill it black: the navigation bar sits there (#347)
        ws.Columns(1).ColumnWidth = LAYOUT_NAV_WIDTH
        ws.Columns(1).Interior.Color = COLOR_BRAND_BAR
    'Set the width of column B
        ws.Columns(2).ColumnWidth = 4
    'Color the brand bar and size the product name
        With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
            .Interior.Color = COLOR_BRAND_BAR
            .Font.Color = COLOR_BRAND_TEXT
        End With
        ws.Cells(1, 2).Font.Size = 14
    'Color the title bar and size the title
        With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
            .Interior.Color = COLOR_TITLE_BAR
            .Font.Color = COLOR_WHITE
        End With
        ws.Cells(2, 2).Font.Size = 18
        ws.Cells(2, 2).Font.Bold = True
    'Style the subtitle
        With ws.Cells(LAYOUT_SUBTITLE_ROW, LAYOUT_FIRST_COL).Font
            .Italic = True
            .Color = COLOR_SUBTITLE
        End With
    'Set the heights of the frame rows
        ws.Rows(1).RowHeight = 24
        ws.Rows(2).RowHeight = 26.1
        ws.Rows(3).RowHeight = 30
        ws.Rows(4).RowHeight = 9.95
    'Set the page setup, with the repeated title rows chosen by sheet
        With ws.PageSetup
            .PrintArea = ""
            .Orientation = 2  'xlLandscape
            .PaperSize = 9  'xlPaperA4
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .LeftMargin = 18
            .RightMargin = 18
            .TopMargin = 36
            .BottomMargin = 36
            .CenterHeader = ""
            .CenterFooter = "&""Aptos,Regular""&9Page &P of &N"
            If ws.Name = SH_README Or ws.Name = SH_HOME Then
                .PrintTitleRows = "$1:$3"
            ElseIf ws.Name = SH_CONTROL Or ws.Name = SH_TEST Or Left$(ws.Name, 2) = "P_" Then
                .PrintTitleRows = "$1:$6"
            Else
                .PrintTitleRows = "$1:$5"
            End If
        End With
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' SET PRINT AREA AND TAB COLOR
'------------------------------------------------------------------------------
    'Extend the print area to the frame
        ExtendPrintArea ws, 4, lastCol
    'Set the tab color of the sheet
        ApplyTabColor ws

End Sub

Private Function SheetTitle(ByVal sheetName As String) As String

'
'------------------------------------------------------------------------------
'                         RETURN THE TITLE OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the text of the title bar for a sheet name; a sheet without a title in the list
'   keeps its name.
'
' WHY THIS EXISTS
'   The sheet names are short identifiers; the title bar shows a descriptive text.
'
' INPUTS
'   sheetName
'     Name of the sheet
'
' RETURNS
'   The title of the 18 sheets in the list; the sheet name itself for any other name.
'
' BEHAVIOR
'   Looks the name up with a Select Case.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Sheet name constants (CORE_Config).
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
    'Look up the title; keep the name when there is none
        Select Case sheetName
            Case SH_IN_DELTA
                SheetTitle = "Delta sensitivities"
            Case SH_IN_VEGA
                SheetTitle = "Vega sensitivities"
            Case SH_IN_CURV
                SheetTitle = "Curvature sensitivities"
            Case SH_IN_DRC
                SheetTitle = "Default risk positions"
            Case SH_IN_RRAO
                SheetTitle = "Residual risk positions"
            Case SH_OUT_SUMMARY
                SheetTitle = "Calculation summary"
            Case SH_OUT_CLASS
                SheetTitle = "Capital by risk class"
            Case SH_OUT_BUCKETS
                SheetTitle = "Bucket aggregation"
            Case SH_OUT_DRC
                SheetTitle = "Default risk calculation"
            Case SH_TEST
                SheetTitle = "Regression tests"
            Case SH_P_GENERAL
                SheetTitle = "General parameters"
            Case SH_P_GIRR
                SheetTitle = "Interest rate parameters"
            Case SH_P_CSR
                SheetTitle = "Credit spread parameters"
            Case SH_P_EQ
                SheetTitle = "Equity parameters"
            Case SH_P_COM
                SheetTitle = "Commodity parameters"
            Case SH_P_FX
                SheetTitle = "Foreign exchange parameters"
            Case SH_P_DRC
                SheetTitle = "Default risk parameters"
            Case SH_P_RRAO
                SheetTitle = "Residual risk parameters"
            Case Else
                SheetTitle = sheetName
        End Select

End Function

Private Sub ApplyTabColor(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                              SET THE TAB COLOR
'------------------------------------------------------------------------------
' PURPOSE
'   Sets the tab color of a sheet by its role (RoleColor).
'
' WHY THIS EXISTS
'   Set again by FinishWorkbook because not every host keeps the tab color in the saved
'   template.
'
' INPUTS
'   ws
'     Sheet to color
'
' BEHAVIOR
'   Roles: navy for README and Home, green for tests, blue for outputs, gold for inputs and
'   parameters.
'
' ERROR POLICY
'   A host that rejects the property is ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' SET THE TAB COLOR
'------------------------------------------------------------------------------
    'Set the color of the role of the sheet; ignore a host that rejects it
        On Error Resume Next
        ws.Tab.Color = RoleColor(ws.Name)
        On Error GoTo 0

End Sub

Private Function RoleColor(ByVal sheetName As String) As Long

'
'------------------------------------------------------------------------------
'                        RETURN THE TAB COLOR OF A ROLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the tab color of a sheet's role, chosen from the sheet name.
'
' INPUTS
'   sheetName
'     Name of the sheet
'
' RETURNS
'   TAB_REFERENCE for README and Home, TAB_TEST for the test sheet, TAB_OUTPUT for Results,
'   Log and the Out_* sheets, TAB_INPUT for every other sheet.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Choose the color by the name of the sheet
        Select Case True
            Case sheetName = SH_README, sheetName = SH_HOME
                RoleColor = TAB_REFERENCE
            Case sheetName = SH_TEST
                RoleColor = TAB_TEST
            Case sheetName = SH_RESULTS, sheetName = SH_LOG, Left$(sheetName, 4) = "Out_"
                RoleColor = TAB_OUTPUT
            Case Else
                RoleColor = TAB_INPUT
        End Select

End Function

'
'------------------------------------------------------------------------------
'
'                                TABLE HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub PutRow(ByVal ws As Object, ByVal r As Long, ByVal values As Variant)

'
'------------------------------------------------------------------------------
'                            WRITE A ROW OF VALUES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a list of values into one row of the table area, from the first table column (C).
'
' WHY THIS EXISTS
'   Text is stored as text, so that a reference such as "MAR21.54" is never converted to a
'   date or a number.
'
' INPUTS
'   ws
'     Sheet to write to
'
'   r
'     Row to write
'
'   values
'     Array of the values, in column order
'
' BEHAVIOR
'   A non-empty text value gets the Text number format ("@") before it is written.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.PutRow"

    Dim Index                   As Long         'Position in the array
    Dim ColumnIndex             As Long         'Table column being written

'------------------------------------------------------------------------------
' WRITE THE VALUES
'------------------------------------------------------------------------------
    'Start before the first table column
        ColumnIndex = 0
    'Write each value; give a non-empty text the Text format first
        For Index = LBound(values) To UBound(values)
            ColumnIndex = ColumnIndex + 1
            If VarType(values(Index)) = vbString Then
                If values(Index) <> vbNullString Then TableCell(ws, r, ColumnIndex).NumberFormat = "@"
            End If
            TableCell(ws, r, ColumnIndex).Value = values(Index)
        Next Index

End Sub

Private Sub PutSample(ByVal ws As Object, ByVal n As Long, ByVal values As Variant)

'
'------------------------------------------------------------------------------
'                            WRITE A SAMPLE RECORD
'------------------------------------------------------------------------------
' PURPOSE
'   Writes input record n (1-based) of an input sheet.
'
' INPUTS
'   ws
'     Input sheet
'
'   n
'     Number of the record: it goes in the n-th row below the header row
'
'   values
'     Array of the values, in column order
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE THE RECORD
'------------------------------------------------------------------------------
    'Write the values in the row of the record
        PutRow ws, LAYOUT_HEADER_ROW + n, values

End Sub

Private Sub SetWidths(ByVal ws As Object, ByVal widths As Variant)

'
'------------------------------------------------------------------------------
'                            SET THE COLUMN WIDTHS
'------------------------------------------------------------------------------
' PURPOSE
'   Sets the widths of the table columns of a sheet from a list, starting at the first table
'   column (C).
'
' INPUTS
'   ws
'     Sheet to format
'
'   widths
'     Array of the column widths, in column order
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.SetWidths"

    Dim Index                   As Long         'Position in the array
    Dim ColumnIndex             As Long         'Table column being sized

'------------------------------------------------------------------------------
' SET THE WIDTHS
'------------------------------------------------------------------------------
    'Ignore formatting errors
        On Error Resume Next
    'Start before the first table column
        ColumnIndex = 0
    'Set the width of each column from the list
        For Index = LBound(widths) To UBound(widths)
            ColumnIndex = ColumnIndex + 1
            ws.Columns(LAYOUT_FIRST_COL + ColumnIndex - 1).ColumnWidth = widths(Index)
        Next Index
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub WriteSectionHeading( _
    ByVal ws As Object, _
    ByVal r As Long, _
    ByVal text As String, _
    ByVal nCols As Long)
'
'------------------------------------------------------------------------------
'                           WRITE A SECTION HEADING
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a section heading in the first table column of a row: navy bold text over a
'   medium navy rule.
'
' INPUTS
'   ws
'     Sheet to write to
'
'   r
'     Row of the heading
'
'   text
'     Text of the heading
'
'   nCols
'     Number of columns the rule spans
'
' BEHAVIOR
'   The text is 14 pt bold, the rule is the bottom border of the first nCols columns, the
'   row is 30 high.
'
' ERROR POLICY
'   The text write raises normally; the formatting ignores errors (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE THE HEADING
'------------------------------------------------------------------------------
    'Write the text
        TableCell(ws, r, 1).Value = text
    'Ignore formatting errors
        On Error Resume Next
    'Make the text bold, 14 pt, navy
        With TableCell(ws, r, 1).Font
            .Bold = True
            .Size = 14
            .Color = COLOR_INPUT_HEADER
        End With
    'Draw the medium navy rule under the columns
        With ws.Range(TableCell(ws, r, 1), TableCell(ws, r, nCols)).Borders(9)  'xlEdgeBottom
            .LineStyle = 1
            .Weight = -4138  'xlMedium
            .Color = COLOR_INPUT_HEADER
        End With
    'Set the row height
        ws.Rows(r).RowHeight = 30
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function PutTable( _
    ByVal ws As Object, _
    ByVal r As Long, _
    ByVal tag As String, _
    ByVal title As String, _
    ByVal headers As Variant) As Long
'
'------------------------------------------------------------------------------
'                         WRITE A TAGGED TABLE HEADING
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the heading of a tagged parameter table: a section heading "[TAG] title" that the
'   engine reads, then a navy header row.
'
' INPUTS
'   ws
'     Sheet to write to
'
'   r
'     Row of the section heading
'
'   tag
'     Tag of the table, written in brackets
'
'   title
'     Title that follows the tag
'
'   headers
'     Array of the column captions
'
' RETURNS
'   The first data row, which is two rows below the heading.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.PutTable"

    Dim ColumnCount             As Long         'Number of columns of the table

'------------------------------------------------------------------------------
' WRITE THE HEADING
'------------------------------------------------------------------------------
    'Count the columns from the captions
        ColumnCount = UBound(headers) - LBound(headers) + 1
    'Write the section heading and the header row
        WriteSectionHeading ws, r, "[" & tag & "]  " & title, ColumnCount
        WriteTableHeader ws, r + 1, headers, COLOR_INPUT_HEADER
    'Return the first data row
        PutTable = r + 2

End Function

Private Sub MarkDuplicateKeys(ByVal ws As Object, ByVal r1 As Long, ByVal r2 As Long)

'
'------------------------------------------------------------------------------
'                           HIGHLIGHT REPEATED KEYS
'------------------------------------------------------------------------------
' PURPOSE
'   Highlights a key that repeats in the first column of a Key | Value table, which the run
'   ignores with a warning (#350).
'
' WHY THIS EXISTS
'   The rule is conditional formatting, so it needs no event and follows the cells.
'
' INPUTS
'   ws
'     Sheet of the table
'
'   r1
'     First data row of the table
'
'   r2
'     Last data row of the table
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds a conditional format to the key cells from r1 to the row below r2, which the run
'   reads too. The format applies to a key that is not blank and appears more than once,
'   with the fail fill and the fail text colors. The formula is first written in the
'   language of the running Excel (LocalFormula); when the host does not take it, the
'   English form is used (#350).
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   LocalFormula; COLOR_FAIL_FILL, COLOR_FAIL_TEXT (CORE_Config); TableCell (CORE_Util).
'
' NOTES
'   Excel reads a relative reference in a formula added by a macro from the active cell, so
'   the formula finds its own cell with ROW.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.MarkDuplicateKeys"

    Dim KeyRange                As Object       'Key cells and the row below
    Dim FormatRule              As Object       'Conditional format being added
    Dim FirstCellAddress        As String       'Address of the first key cell
    Dim KeyExpression           As String       'Formula text that reads the key of the current row
    Dim EnglishFormula          As String       'Rule formula in English

'------------------------------------------------------------------------------
' ADD THE RULE
'------------------------------------------------------------------------------
    'Ignore errors from the host
        On Error Resume Next
    'Take the key cells and the row below, which the run reads too
        Set KeyRange = ws.Range(TableCell(ws, r1, 1), TableCell(ws, r2 + 1, 1))
    'Build the formula: a key that is not blank and appears more than once
        FirstCellAddress = TableCell(ws, r1, 1).Address
        KeyExpression = "INDEX(" & KeyRange.Address & ",ROW()-ROW(" & FirstCellAddress & ")+1)"
        EnglishFormula = "=AND(" & KeyExpression & "<>"""",COUNTIF(" & KeyRange.Address & "," & KeyExpression & ")>1)"
    'FormatConditions.Add reads its formula in the language of the Excel that runs it, so the local
    'form comes first; the English form is the fallback for a host that takes it (#350)
        Set FormatRule = KeyRange.FormatConditions.Add(2, , LocalFormula(ws, EnglishFormula))
        If FormatRule Is Nothing Then
            Err.Clear
            Set FormatRule = KeyRange.FormatConditions.Add(2, , EnglishFormula)
        End If
    'Color the repeated key with the fail fill and text
        FormatRule.Interior.Color = COLOR_FAIL_FILL
        FormatRule.Font.Color = COLOR_FAIL_TEXT
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function LocalFormula(ByVal ws As Object, ByVal english As String) As String

'
'------------------------------------------------------------------------------
'                    TRANSLATE A FORMULA TO THE LOCAL EXCEL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a formula in the language and separators of the running Excel.
'
' INPUTS
'   ws
'     Sheet whose last cell is used as the scratch cell
'
'   english
'     Formula in English
'
' RETURNS
'   The local form, or the English form when the host has no FormulaLocal.
'
' BEHAVIOR
'   Writes the English form into a scratch cell, reads FormulaLocal and clears the cell.
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.LocalFormula"

    Dim ScratchCell             As Object       'Last cell of the sheet, used to translate the formula

'------------------------------------------------------------------------------
' TRANSLATE
'------------------------------------------------------------------------------
    'Ignore errors from the host
        On Error Resume Next
    'Write the English form, read the local form and clear the cell
        Set ScratchCell = ws.Cells(ws.Rows.Count, ws.Columns.Count)
        ScratchCell.Formula = english
        LocalFormula = ScratchCell.FormulaLocal
        ScratchCell.Clear
    'Restore normal error handling
        On Error GoTo 0
    'Fall back to the English form when there is no local form
        If LocalFormula = vbNullString Then LocalFormula = english

End Function

Private Sub FormatParamRows( _
    ByVal ws As Object, _
    ByVal r1 As Long, _
    ByVal r2 As Long, _
    ByVal nCols As Long, _
    ByVal firstValueCol As Long, _
    ByVal lastValueCol As Long, _
    ByVal pctCols As String, _
    Optional ByVal keyTable As Boolean = False)
'
'------------------------------------------------------------------------------
'                     FORMAT THE ROWS OF A PARAMETER TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Formats the data rows of a parameter table: the editable columns in the input fill,
'   percentages where the column (or, in Key | Value tables, the key) denotes a rate.
'
' WHY THIS EXISTS
'   The parameter tables of the P_* sheets and of Control share one set of formats.
'
' INPUTS
'   ws
'     Sheet of the table
'
'   r1
'     First data row
'
'   r2
'     Last data row
'
'   nCols
'     Number of columns of the table
'
'   firstValueCol
'     First editable column
'
'   lastValueCol
'     Last editable column
'
'   pctCols
'     Comma-separated list of the columns shown as percentages, or KEY to choose by the key
'     in the first column (RHO*, RW*, GAMMA*, THETA, ALPHA)
'
'   keyTable
'     True for a Key | Value table: repeated keys are highlighted
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Formats the borders and fonts of the rows, optionally marks repeated keys, fills the
'   editable cells with the input fill and unlocks them, makes the first column bold and
'   sets the percentage format on the numeric cells that qualify.
'
' ERROR POLICY
'   The formatting is best effort (On Error Resume Next).
'
' DEPENDENCIES
'   FormatDataRows, TableCell, IsNumber (CORE_Util); MarkDuplicateKeys; COLOR_INPUT_CELL
'   (CORE_Config).
'
' NOTES
'   The editable cells stay editable on a protected sheet (#349).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.FormatParamRows"

    Dim RowIndex                As Long         'Row being formatted
    Dim ColumnIndex             As Long         'Column being formatted
    Dim KeyName                 As String       'Upper-case key of the row

'------------------------------------------------------------------------------
' FORMAT THE ROWS
'------------------------------------------------------------------------------
    'Format the borders and the fonts of the data rows
        FormatDataRows ws, r1, r2, nCols, 0, "", False
    'Highlight repeated keys in a Key | Value table
        If keyTable Then MarkDuplicateKeys ws, r1, r2
    'Ignore formatting errors
        On Error Resume Next
    'Fill the editable columns with the input fill and unlock them: editable on a protected sheet (#349)
        ws.Range(TableCell(ws, r1, firstValueCol), TableCell(ws, r2, lastValueCol)).Interior.Color = COLOR_INPUT_CELL
        ws.Range(TableCell(ws, r1, firstValueCol), TableCell(ws, r2, lastValueCol)).Locked = False
    'Make the first column bold
        ws.Range(TableCell(ws, r1, 1), TableCell(ws, r2, 1)).Font.Bold = True
    'Show the cells of the percentage columns, or the rate keys of a Key | Value table, as percentages;
    'the compare mode 0 of InStr is vbBinaryCompare
        For RowIndex = r1 To r2
            For ColumnIndex = firstValueCol To lastValueCol
                If InStr(1, "," & pctCols & ",", "," & CStr(ColumnIndex) & ",", 0) > 0 Then
                    If IsNumber(TableCell(ws, RowIndex, ColumnIndex).Value) Then
                        TableCell(ws, RowIndex, ColumnIndex).NumberFormat = "0.0##%"
                    End If
                ElseIf pctCols = "KEY" Then
                    KeyName = UCase$(CStr(TableCell(ws, RowIndex, 1).Value))
                    If (Left$(KeyName, 3) = "RHO" Or Left$(KeyName, 2) = "RW" Or Left$(KeyName, 5) = "GAMMA" Or _
                        KeyName = "THETA" Or KeyName = "ALPHA") And IsNumber(TableCell(ws, RowIndex, ColumnIndex).Value) Then
                        TableCell(ws, RowIndex, ColumnIndex).NumberFormat = "0.0##%"
                    End If
                End If
            Next ColumnIndex
        Next RowIndex
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub WrapColumn(ByVal ws As Object, ByVal r1 As Long, ByVal r2 As Long, ByVal k As Long)

'
'------------------------------------------------------------------------------
'                    WRAP A COLUMN AND SET THE ROW HEIGHTS
'------------------------------------------------------------------------------
' PURPOSE
'   Wraps column k of rows r1 to r2 and sets each row high enough for its lines.
'
' WHY THIS EXISTS
'   AutoFit measured the text in the fallback font of the headless build, which is narrower
'   than Aptos, so Excel cut the second line off (#206).
'
' INPUTS
'   ws
'     Sheet of the table
'
'   r1
'     First row
'
'   r2
'     Last row
'
'   k
'     Table column to wrap
'
' BEHAVIOR
'   The line count is the same on every host: words, broken after a hyphen too, filled into
'   1.2 characters per unit of column width (Excel fit 50 characters of these texts into a
'   column 40 wide). A row that needs more than one line gets 4 + 13.5 points per line.
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WrapColumn"

    Dim RowIndex                As Long         'Row being sized
    Dim LineCount               As Long         'Lines the text needs
    Dim CharsPerLine            As Long         'Characters that fit in one line

'------------------------------------------------------------------------------
' WRAP THE COLUMN
'------------------------------------------------------------------------------
    'Ignore formatting errors
        On Error Resume Next
    'Turn the wrap on for the cells of the column
        ws.Range(TableCell(ws, r1, k), TableCell(ws, r2, k)).WrapText = True
    'Compute how many characters fit in one line of the column
        CharsPerLine = Int(ws.Columns(LAYOUT_FIRST_COL + k - 1).ColumnWidth * 1.2)
    'Set the height of each row from the lines its text needs
        For RowIndex = r1 To r2
            LineCount = WrappedLines(CStr(TableCell(ws, RowIndex, k).Value), CharsPerLine)
            If LineCount > 1 Then ws.Rows(RowIndex).RowHeight = 4 + 13.5 * LineCount
        Next RowIndex
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function WrappedLines(ByVal text As String, ByVal perLine As Long) As Long

'
'------------------------------------------------------------------------------
'                          COUNT THE LINES OF A TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the lines that a text needs at perLine characters a line, breaking at spaces and
'   after hyphens.
'
' WHY THIS EXISTS
'   Every host gets the same line count, unlike AutoFit.
'
' INPUTS
'   text
'     Text to measure
'
'   perLine
'     Characters that fit in one line
'
' RETURNS
'   The number of lines, at least 1. 1 when perLine is less than 1 or the text is empty.
'
' BEHAVIOR
'   Splits the text into words, breaking after a hyphen too. Fills the words into lines,
'   with one space between words except after a hyphen. A word longer than a line takes as
'   many lines as its length needs.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   tests/test_template.py repeats this algorithm in Python to check the row heights.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WrappedLines"

    Dim WordList                As Variant      'Words of the text
    Dim Index                   As Long         'Loop counter
    Dim CurrentWord             As String       'Word being placed
    Dim UsedChars               As Long         'Characters used in the current line
    Dim LineCount               As Long         'Lines counted so far
    Dim GapChars                As Long         'Space before the word, 0 or 1
    Dim AfterHyphen             As Boolean      'True when the previous word ended with a hyphen

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Start with one line
        WrappedLines = 1
    'Return one line when there is no width or no text
        If perLine < 1 Or Len(text) = 0 Then Exit Function
    'Split the text into words, breaking after each hyphen
        LineCount = 1
        WordList = Split(Replace(text, "-", "- "), " ")

'------------------------------------------------------------------------------
' COUNT THE LINES
'------------------------------------------------------------------------------
    'Place each word on the current line or on a new one
        For Index = LBound(WordList) To UBound(WordList)
            CurrentWord = WordList(Index)
            If Len(CurrentWord) > 0 Then
                If UsedChars > 0 And Not AfterHyphen Then
                    GapChars = 1
                Else
                    GapChars = 0
                End If
                If UsedChars > 0 And UsedChars + GapChars + Len(CurrentWord) > perLine Then
                    LineCount = LineCount + 1
                    UsedChars = 0
                    GapChars = 0
                End If
                UsedChars = UsedChars + GapChars + Len(CurrentWord)
                Do While UsedChars > perLine
                    LineCount = LineCount + 1
                    UsedChars = UsedChars - perLine
                Loop
                AfterHyphen = (Right$(CurrentWord, 1) = "-")
            End If
        Next Index

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the line count
        WrappedLines = LineCount

End Function

'
'------------------------------------------------------------------------------
'
'                         BUTTONS AND WINDOW SETTINGS
'
'------------------------------------------------------------------------------
'

Private Sub AddButtonRow(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                        ADD THE BUTTON ROW OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Adds the command buttons of a sheet as one row in its title bar, right-aligned to the
'   bar and never over the title.
'
' WHY THIS EXISTS
'   The buttons of a sheet come from SheetButtons and sit in one row of its title bar
'   (#347).
'
' INPUTS
'   ws
'     Sheet that receives the buttons
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the buttons of the sheet and removes the ones already there. A sheet without
'   buttons ends there. Otherwise sizes each button from its caption (24 + 6.4 per
'   character, 6 between them), finds the right end of the title bar by walking back from
'   column 60, and places the row so that it ends 8 points before that edge but never starts
'   before the end of the title. Each button gets its caption, its macro (OnAction),
'   free-floating placement and the workbook font.
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   SheetButtons, RemoveButtons; FONT_NAME, COLOR_TITLE_BAR (CORE_Config).
'
' NOTES
'   Each button is Aptos 11 pt and 19.5 pt high.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.AddButtonRow"

    Dim ButtonDefs              As Variant      'Buttons of the sheet, or Empty
    Dim Index                   As Long         'Loop counter
    Dim ButtonCount             As Long         'Number of buttons
    Dim TotalWidth              As Double       'Width of the whole row
    Dim ButtonLeft              As Double       'Left edge of the next button
    Dim BarRight                As Double       'Right end of the title bar
    Dim TitleEnd                As Double       'Estimated right end of the title text
    Dim ColumnIndex             As Long         'Column being tested
    Dim Button                  As Object       'Button being added
    Dim ButtonWidths()          As Double       'Width of each button

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Read the buttons of the sheet
        ButtonDefs = SheetButtons(ws.Name)
    'Remove the buttons that are already on the sheet
        RemoveButtons ws
    'Leave when the sheet has no buttons
        If IsEmpty(ButtonDefs) Then Exit Sub
    'Ignore formatting errors
        On Error Resume Next

'------------------------------------------------------------------------------
' MEASURE THE ROW
'------------------------------------------------------------------------------
    'Size each button from its caption and add the widths and the gaps
        ButtonCount = UBound(ButtonDefs) - LBound(ButtonDefs) + 1
        ReDim ButtonWidths(0 To ButtonCount - 1)
        For Index = 0 To ButtonCount - 1
            ButtonWidths(Index) = 24 + 6.4 * Len(CStr(ButtonDefs(LBound(ButtonDefs) + Index)(0)))
            TotalWidth = TotalWidth + ButtonWidths(Index) + IIf(Index < ButtonCount - 1, 6, 0)
        Next Index
    'Find the right end of the title bar, walking back from column 60
        For ColumnIndex = 60 To 1 Step -1
            If ws.Cells(2, ColumnIndex).Interior.Color = COLOR_TITLE_BAR Then
                BarRight = ws.Cells(2, ColumnIndex).Left + ws.Cells(2, ColumnIndex).Width
                Exit For
            End If
        Next ColumnIndex
    'Estimate where the title text ends
        TitleEnd = ws.Cells(2, 2).Left + 11 * Len(CStr(ws.Cells(2, 2).Value)) + 18
    'Right-align the row to the bar, but never over the title
        ButtonLeft = BarRight - TotalWidth - 8
        If ButtonLeft < TitleEnd Then ButtonLeft = TitleEnd

'------------------------------------------------------------------------------
' ADD THE BUTTONS
'------------------------------------------------------------------------------
    'Add each button with its caption, macro and font
        For Index = 0 To ButtonCount - 1
            Set Button = Nothing
            Set Button = ws.Buttons.Add(ButtonLeft, ws.Cells(2, 1).Top + 3.5, ButtonWidths(Index), 19.5)
            If Not Button Is Nothing Then
                Button.Caption = CStr(ButtonDefs(LBound(ButtonDefs) + Index)(0))
                Button.OnAction = CStr(ButtonDefs(LBound(ButtonDefs) + Index)(1))
                Button.Placement = 3  'xlFreeFloating
                Button.Font.Name = FONT_NAME
                Button.Font.Size = 11
            End If
            ButtonLeft = ButtonLeft + ButtonWidths(Index) + 6
        Next Index
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Function SheetButtons(ByVal sheetName As String) As Variant

'
'------------------------------------------------------------------------------
'                        RETURN THE BUTTONS OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the buttons of a sheet as one Array(caption, macro) each, or Empty.
'
' WHY THIS EXISTS
'   The buttons depend on the role of the sheet: the run and its companions, Clear outputs
'   on every output sheet, and the regression test on its own sheet.
'
' INPUTS
'   sheetName
'     Name of the sheet
'
' RETURNS
'   An array of Array(caption, macro name) items, or Empty for a sheet without buttons:
'     - Control: Run FRTB-SA, Clear outputs, Run regression test
'     - Results: Run FRTB-SA, Clear outputs
'     - Log and the Out_* sheets: Clear outputs
'     - the test sheet: Run regression test
'
' BEHAVIOR
'   Builds the three items, then picks the list by the sheet name.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Sheet name constants (CORE_Config).
'
' NOTES
'   The macros are RunFRTB, ClearOutputs and Test_RunRegression.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.SheetButtons"

    Dim RunItem                 As Variant      'Caption and macro of the run button
    Dim ClearItem               As Variant      'Caption and macro of the clear button
    Dim TestItem                As Variant      'Caption and macro of the test button

'------------------------------------------------------------------------------
' DEFINE THE BUTTONS
'------------------------------------------------------------------------------
    'Define the run, clear and test buttons
        RunItem = Array("Run FRTB-SA", "RunFRTB")
        ClearItem = Array("Clear outputs", "ClearOutputs")
        TestItem = Array("Run regression test", "Test_RunRegression")

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Choose the list of buttons by the name of the sheet
        Select Case True
            Case sheetName = SH_CONTROL
                SheetButtons = Array(RunItem, ClearItem, TestItem)
            Case sheetName = SH_RESULTS
                SheetButtons = Array(RunItem, ClearItem)
            Case sheetName = SH_LOG, Left$(sheetName, 4) = "Out_"
                SheetButtons = Array(ClearItem)
            Case sheetName = SH_TEST
                SheetButtons = Array(TestItem)
            Case Else
                SheetButtons = Empty
        End Select

End Function

Private Sub RemoveButtons(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                        REMOVE THE BUTTONS OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes the buttons of a sheet.
'
' INPUTS
'   ws
'     Sheet to clear
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REMOVE
'------------------------------------------------------------------------------
    'Delete the buttons; ignore a host that refuses
        On Error Resume Next
        ws.Buttons.Delete
        On Error GoTo 0

End Sub

Private Sub ApplyWindowSettings()

'
'------------------------------------------------------------------------------
'                          APPLY THE WINDOW SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the gridlines and freezes the panes below the table header on every managed sheet.
'
' WHY THIS EXISTS
'   Needs each sheet to be active once. The wide input, output and Log sheets also keep
'   their key columns in view; the narrow tables of Results, Control, Test_Expected and the
'   parameter sheets freeze rows only.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   For each managed sheet: activates it, hides the gridlines, removes the panes and freezes
'   new ones:
'     - In_*, Out_* and Log: the rows above the first data row and the columns up to the
'       first table column
'     - Results: the rows above the first data row
'     - Control, the test sheet and the P_* sheets: rows 1 to 6, the header row and the row
'       below it
'     - any other sheet: rows 1 to 3
'   Then selects A1.
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   ActiveWindow of the Excel application; ManagedSheets (CORE_Config).
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
    Const PROC_NAME             As String = "M_Setup.ApplyWindowSettings"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim Index                   As Long         'Loop counter
    Dim CurrentSheet            As Object       'Sheet being set up

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Read the managed sheet names
        SheetNames = ManagedSheets()
    'Ignore errors from the host
        On Error Resume Next

'------------------------------------------------------------------------------
' SET UP THE SHEETS
'------------------------------------------------------------------------------
    'Activate each sheet, hide the gridlines and freeze the panes chosen by the kind of sheet
        For Index = LBound(SheetNames) To UBound(SheetNames)
            Set CurrentSheet = ThisWorkbook.Worksheets(CStr(SheetNames(Index)))
            CurrentSheet.Activate
            Application.ActiveWindow.DisplayGridlines = False
            Application.ActiveWindow.FreezePanes = False
            If Left$(CurrentSheet.Name, 3) = "In_" Or Left$(CurrentSheet.Name, 4) = "Out_" Or CurrentSheet.Name = SH_LOG Then
                CurrentSheet.Cells(LAYOUT_FIRST_DATA_ROW, LAYOUT_FIRST_COL + 1).Select
                Application.ActiveWindow.FreezePanes = True
            ElseIf CurrentSheet.Name = SH_RESULTS Then
                CurrentSheet.Cells(LAYOUT_FIRST_DATA_ROW, 1).Select
                Application.ActiveWindow.FreezePanes = True
            ElseIf CurrentSheet.Name = SH_CONTROL Or CurrentSheet.Name = SH_TEST Or _
                Left$(CurrentSheet.Name, 2) = "P_" Then
                CurrentSheet.Cells(LAYOUT_HEADER_ROW + 2, 1).Select
                Application.ActiveWindow.FreezePanes = True
            Else
                CurrentSheet.Cells(4, 1).Select
                Application.ActiveWindow.FreezePanes = True
            End If
            CurrentSheet.Cells(1, 1).Select
        Next Index
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                                 README SHEET
'
'------------------------------------------------------------------------------
'

Private Sub BuildReadme()

'
'------------------------------------------------------------------------------
'                            BUILD THE README SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the README sheet: the frame and three guide blocks, "How to use", "Input
'   conventions" and "Sheet map".
'
' WHY THIS EXISTS
'   The sheet tells the user how to run the workbook and how to fill the input sheets.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the frame and sets the column widths. Then writes each block as a section heading
'   followed by guide rows (label | text): seven usage steps, seven input conventions and
'   two sheet-map rows.
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the text writes raise normally.
'
' DEPENDENCIES
'   ApplyFrame, SetWidths, WriteSectionHeading, WriteGuideRows; APP_TITLE, APP_VERSION
'   (CORE_Config).
'
' NOTES
'   The subtitle names the version and states that the results are not validated for
'   regulatory use.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.BuildReadme"

    Dim ReadmeSheet             As Object       'README sheet
    Dim RowIndex                As Long         'Row being written

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Locate the README sheet
        Set ReadmeSheet = GetSheet(SH_README)
    'Write the frame and set the column widths
        ApplyFrame ReadmeSheet, "ReadMe", APP_TITLE & " (Basel MAR21-MAR23) in Excel VBA. Version " & APP_VERSION & _
            ": results are not validated for regulatory use.", 4
        SetWidths ReadmeSheet, Array(55, 150)

'------------------------------------------------------------------------------
' WRITE THE GUIDE
'------------------------------------------------------------------------------
    'Write the "How to use" block
        RowIndex = 5
        WriteSectionHeading ReadmeSheet, RowIndex, "How to use", 2
        RowIndex = WriteGuideRows(ReadmeSheet, RowIndex + 1, Array( _
            "1. Enable macros", "Open the .xlsm and click Enable Content. If Excel shows that macros are blocked, close the file, open its Properties, tick Unblock and reopen it.", _
            "2. Settings", "Control: reporting currency and the national discretions (sqrt(2) for specified GIRR currencies and FX pairs, FX curvature divisor). Cream cells are inputs.", _
            "3. Sensitivities", "In_Delta, In_Vega, In_Curvature: one row per sensitivity, in the reporting currency. Rows with the same risk factor are netted.", _
            "4. Positions", "In_DRC: one row per position for the default risk charge. In_RRAO: gross notional of instruments with residual risks.", _
            "5. Run", "Click Run FRTB-SA (macro RunFRTB). Results, Out_*, Out_Summary and Log are rewritten on every run.", _
            "6. Review", "Log lists every rejected row with the reason; a rejected row is never counted as zero. Results shows the capital by risk class, component and correlation scenario.", _
            "7. Test", "Test_Expected compares the sample portfolio with values from an independent reference implementation (Run regression test, development build)."))
    'Write the "Input conventions" block
        RowIndex = RowIndex + 1
        WriteSectionHeading ReadmeSheet, RowIndex, "Input conventions", 2
        RowIndex = WriteGuideRows(ReadmeSheet, RowIndex + 1, Array( _
            "Risk classes", "GIRR, CSR_NS (credit spread, non-securitization), EQ, COM, FX. CSR securitization and the correlation trading portfolio are not implemented and are rejected.", _
            "Buckets", "GIRR and FX: ISO currency code, one for all variants of a currency (KRW, not KRO). CSR_NS 1-18, EQ 1-13, COM 1-11 as in MAR21.51, MAR21.72 and MAR21.82 (see the P_* sheets).", _
            "Qualifier", "GIRR: curve name. CSR_NS and EQ: issuer or index. COM: commodity. FX: the currency.", _
            "Sub-type", "GIRR: YIELD, INFLATION or XCCY; curvature YIELD only. CSR_NS: BOND or CDS. EQ: SPOT or REPO; vega and curvature SPOT only. COM: delivery location. FX: none.", _
            "Tenors", "GIRR 0.25-30y, CSR 0.5-10y, COM 0-30y; vega option and GIRR underlying maturities 0.5, 1, 3, 5, 10y. GIRR INFLATION and XCCY vega: option maturity only.", _
            "Amounts", "Sensitivities as defined in MAR21.19-26 (for example GIRR: PV change for 1bp / 0.0001). Vega: vega x implied volatility. Curvature: CVR up and CVR down per risk factor; References reporting ccy TRUE marks FX options that reference the reporting currency (MAR21.98).", _
            "DRC", "Signed notional and market value (negative for shorts), maturity in years (cash equity: 0.25 or at least 1, MAR22.16), bucket CORP, SOV or LGM, credit quality AAA ... DEFAULTED, seniority EQUITY, NON_SENIOR, SENIOR or COVERED."))
    'Write the "Sheet map" block
        RowIndex = RowIndex + 1
        WriteSectionHeading ReadmeSheet, RowIndex, "Sheet map", 2
        RowIndex = WriteGuideRows(ReadmeSheet, RowIndex + 1, Array( _
            "Where is a sheet", "Home lists every sheet with its role and links to it. On every other sheet the bar in column A does the same.", _
            "Tab colors", "Navy: Home and README. Gold: inputs and parameters (Control, In_*, P_*); parameter codes in column C are read by the engine, and only cream cells are meant to be changed. Blue: outputs (Results, Out_*, Log). Green: Test_Expected."))

End Sub

Private Function WriteGuideRows(ByVal ws As Object, ByVal r As Long, ByVal pairs As Variant) As Long

'
'------------------------------------------------------------------------------
'                       WRITE THE GUIDE ROWS OF A BLOCK
'------------------------------------------------------------------------------
' PURPOSE
'   Writes two-column guide rows (label | text) with dashed inner borders and a thin outer
'   border.
'
' WHY THIS EXISTS
'   The README sheet shows its text as blocks of label and text rows.
'
' INPUTS
'   ws
'     Sheet to write to
'
'   r
'     First row of the block
'
'   pairs
'     Array of label and text values: label, text, label, text, ...
'
' RETURNS
'   The row after the block.
'
' BEHAVIOR
'   Writes each pair in the first two table columns, wraps the text, draws the borders,
'   makes the labels bold, sets the text to 9 pt and the rows to 33.95 points, and extends
'   the print area.
'
' ERROR POLICY
'   The writes raise normally; the formatting ignores errors (On Error Resume Next).
'
' DEPENDENCIES
'   TableCell, FormatBorders, ExtendPrintArea (CORE_Util); COLOR_GRID (CORE_Config).
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
    Const PROC_NAME             As String = "M_Setup.WriteGuideRows"

    Dim Index                   As Long         'Position in the array
    Dim FirstRow                As Long         'First row of the block

'------------------------------------------------------------------------------
' WRITE THE ROWS
'------------------------------------------------------------------------------
    'Remember the first row of the block
        FirstRow = r
    'Write each label and text pair in its row
        For Index = LBound(pairs) To UBound(pairs) Step 2
            TableCell(ws, r, 1).Value = pairs(Index)
            TableCell(ws, r, 2).Value = pairs(Index + 1)
            r = r + 1
        Next Index

'------------------------------------------------------------------------------
' FORMAT THE BLOCK
'------------------------------------------------------------------------------
    'Ignore formatting errors
        On Error Resume Next
    'Wrap the text of the block
        With ws.Range(TableCell(ws, FirstRow, 1), TableCell(ws, r - 1, 2))
            .WrapText = True
        End With
    'Draw the dashed inner borders
        FormatBorders ws.Range(TableCell(ws, FirstRow, 1), TableCell(ws, r - 1, 2)), COLOR_GRID, -4115
    'Draw the thin outer border
        With ws.Range(TableCell(ws, FirstRow, 1), TableCell(ws, r - 1, 2))
            .Borders(7).LineStyle = 1  'xlEdgeLeft
            .Borders(8).LineStyle = 1  'xlEdgeTop
            .Borders(9).LineStyle = 1  'xlEdgeBottom
            .Borders(10).LineStyle = 1  'xlEdgeRight
        End With
    'Make the labels bold and the texts 9 pt
        ws.Range(TableCell(ws, FirstRow, 1), TableCell(ws, r - 1, 1)).Font.Bold = True
        ws.Range(TableCell(ws, FirstRow, 2), TableCell(ws, r - 1, 2)).Font.Size = 9
    'Set the row height of the block
        ws.Range(ws.Rows(FirstRow), ws.Rows(r - 1)).RowHeight = 33.95
    'Restore normal error handling
        On Error GoTo 0
    'Extend the print area to the block
        ExtendPrintArea ws, r - 1, LAYOUT_FIRST_COL + 1

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the row after the block
        WriteGuideRows = r

End Function

'
'------------------------------------------------------------------------------
'
'                      HOME AND THE NAVIGATION BAR (#347)
'
'------------------------------------------------------------------------------
'

Private Sub BuildHome()

'
'------------------------------------------------------------------------------
'                             BUILD THE HOME SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the start sheet: what the workbook does, the steps of a run and every sheet with
'   its role and a link to it.
'
' WHY THIS EXISTS
'   Home is the first sheet of the workbook and lists every sheet in navigation order.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the frame and the column widths, then three tables:
'     - "What <product> does": what it calculates, what it does not cover and its status
'     - "How to run": five steps
'     - "Sheets": one row per entry of NavEntry, with the name as a HYPERLINK formula to
'       cell A1 of the sheet
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the writes raise normally. The links are
'   best effort.
'
' DEPENDENCIES
'   ApplyFrame, SetWidths, WriteSectionHeading, PutRow, NavEntry; WriteTableHeader,
'   FormatDataRows (CORE_Util).
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
    Const PROC_NAME             As String = "M_Setup.BuildHome"

    Dim HomeSheet               As Object       'Home sheet
    Dim RowIndex                As Long         'Row being written
    Dim Index                   As Long         'Loop counter
    Dim FirstRow                As Long         'First data row of the Sheets table
    Dim Entry                   As Variant      'Navigation entry

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Locate the Home sheet
        Set HomeSheet = GetSheet(SH_HOME)
    'Write the frame and set the column widths
        ApplyFrame HomeSheet, "Home", "Start here: what " & APP_NAME & " does, the steps of a run and where every sheet is.", 5
        SetWidths HomeSheet, Array(26, 22, 118)

'------------------------------------------------------------------------------
' WRITE THE TABLES
'------------------------------------------------------------------------------
    'Write the table of what the workbook does
        RowIndex = LAYOUT_HEADER_ROW
        WriteSectionHeading HomeSheet, RowIndex, "What " & APP_NAME & " does", 3
        WriteTableHeader HomeSheet, RowIndex + 1, Array("Topic", "Kind", "Description"), COLOR_INPUT_HEADER
        RowIndex = RowIndex + 2
        PutRow HomeSheet, RowIndex, Array("Calculates", "Capital", "The standardized approach capital requirement of the Basel Framework (MAR21-MAR23): SBM, the default risk charge and the RRAO.")
        PutRow HomeSheet, RowIndex + 1, Array("Does not cover", "Out of scope", "CSR securitization, the correlation trading portfolio, securitization DRC, jurisdiction profiles and the pricing of sensitivities.")
        PutRow HomeSheet, RowIndex + 2, Array("Status", "Prototype", "Results are not validated for regulatory use. The README sheet has the input conventions.")
        FormatDataRows HomeSheet, RowIndex, RowIndex + 2, 3, 0, "", False
    'Write the table of the steps of a run
        RowIndex = RowIndex + 4
        WriteSectionHeading HomeSheet, RowIndex, "How to run", 3
        WriteTableHeader HomeSheet, RowIndex + 1, Array("Step", "Where", "What to do"), COLOR_INPUT_HEADER
        RowIndex = RowIndex + 2
        PutRow HomeSheet, RowIndex, Array("1. Check the settings", "Control", "Reporting currency and the national discretions. Cream cells are inputs; every other cell is locked, with no password.")
        PutRow HomeSheet, RowIndex + 1, Array("2. Enter the inputs", "In_* sheets", "Delta, vega and curvature sensitivities, DRC positions and RRAO notional, one row each, in the reporting currency.")
        PutRow HomeSheet, RowIndex + 2, Array("3. Run", "Control", "Click Run FRTB-SA. The output sheets and the Log are rewritten on every run.")
        PutRow HomeSheet, RowIndex + 3, Array("4. Read the results", "Results, Log", "Results shows the capital; the Log lists every rejected row with its reason.")
        PutRow HomeSheet, RowIndex + 4, Array("Macros", "Excel", "The commands need macros enabled: open the .xlsm and click Enable Content.")
        FormatDataRows HomeSheet, RowIndex, RowIndex + 4, 3, 0, "", False
    'Write the table of the sheets
        RowIndex = RowIndex + 6
        WriteSectionHeading HomeSheet, RowIndex, "Sheets", 3
        WriteTableHeader HomeSheet, RowIndex + 1, Array("Sheet", "Role", "Contents"), COLOR_INPUT_HEADER
        FirstRow = RowIndex + 2
        For Index = 1 To NAV_COUNT
            Entry = NavEntry(Index)
            PutRow HomeSheet, FirstRow + Index - 1, Array(Entry(1), Entry(2), Entry(3))
        Next Index
        FormatDataRows HomeSheet, FirstRow, FirstRow + NAV_COUNT - 1, 3, 0, "", False

'------------------------------------------------------------------------------
' LINK THE SHEETS
'------------------------------------------------------------------------------
    'Turn each sheet name into a link to the sheet; ignore errors from the host
        On Error Resume Next
        For Index = 1 To NAV_COUNT
            Entry = NavEntry(Index)
            With TableCell(HomeSheet, FirstRow + Index - 1, 1)
                .Formula = "=HYPERLINK(""#'" & Entry(1) & "'!A1"",""" & Entry(1) & """)"
                .Font.Color = COLOR_INPUT_HEADER
                .Font.Underline = 2  'xlUnderlineStyleSingle
            End With
        Next Index
        On Error GoTo 0

End Sub

Private Function NavEntry(ByVal i As Long) As Variant

'
'------------------------------------------------------------------------------
'                          RETURN A NAVIGATION ENTRY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the entry of the i-th sheet in navigation order: Array(section heading, sheet,
'   role, contents).
'
' WHY THIS EXISTS
'   Home lists the entries and the navigation bars link to them.
'
' INPUTS
'   i
'     Position of the entry, from 1 to NAV_COUNT
'
' RETURNS
'   Array(section heading, sheet name, role, contents). Any value other than 1 to 22 returns
'   the entry of the README sheet.
'
' BEHAVIOR
'   Looks the position up with a Select Case.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Sheet name constants (CORE_Config).
'
' NOTES
'   NAV_COUNT must equal the number of entries.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the entry of the position
        Select Case i
            Case 1
                NavEntry = Array("START", SH_HOME, "Start", "This sheet: what the workbook does, how to run it and where every sheet is.")
            Case 2
                NavEntry = Array("INPUTS", SH_CONTROL, "Inputs", "Run settings and the buttons Run FRTB-SA, Clear outputs and Run regression test.")
            Case 3
                NavEntry = Array("INPUTS", SH_IN_DELTA, "Inputs", "Delta sensitivities in the reporting currency, one row per risk factor and tenor.")
            Case 4
                NavEntry = Array("INPUTS", SH_IN_VEGA, "Inputs", "Vega sensitivities per option maturity, and GIRR underlying maturity.")
            Case 5
                NavEntry = Array("INPUTS", SH_IN_CURV, "Inputs", "Curvature CVR up and CVR down per risk factor.")
            Case 6
                NavEntry = Array("INPUTS", SH_IN_DRC, "Inputs", "Positions for the default risk charge: notional, market value, maturity and seniority.")
            Case 7
                NavEntry = Array("INPUTS", SH_IN_RRAO, "Inputs", "Gross notional of instruments with residual risks, by category.")
            Case 8
                NavEntry = Array("PARAMETERS", SH_P_GENERAL, "Parameters", "Correlation scenarios, vega parameters and regulatory vertices.")
            Case 9
                NavEntry = Array("PARAMETERS", SH_P_GIRR, "Parameters", "General interest rate risk: delta risk weights and correlations.")
            Case 10
                NavEntry = Array("PARAMETERS", SH_P_CSR, "Parameters", "Credit spread risk, non-securitization: buckets, risk weights and correlations.")
            Case 11
                NavEntry = Array("PARAMETERS", SH_P_EQ, "Parameters", "Equity risk: buckets, risk weights and correlations.")
            Case 12
                NavEntry = Array("PARAMETERS", SH_P_COM, "Parameters", "Commodity risk: buckets, risk weights and correlations.")
            Case 13
                NavEntry = Array("PARAMETERS", SH_P_FX, "Parameters", "Foreign exchange risk: risk weights and specified pairs.")
            Case 14
                NavEntry = Array("PARAMETERS", SH_P_DRC, "Parameters", "Default risk charge: risk weights, LGD, maturity scaling and buckets.")
            Case 15
                NavEntry = Array("PARAMETERS", SH_P_RRAO, "Parameters", "Residual risk add-on risk weights.")
            Case 16
                NavEntry = Array("OUTPUTS", SH_RESULTS, "Outputs", "Capital by risk class, component and correlation scenario, with the total.")
            Case 17
                NavEntry = Array("OUTPUTS", SH_OUT_SUMMARY, "Outputs", "Key results of the last run, read by the regression test and automation.")
            Case 18
                NavEntry = Array("OUTPUTS", SH_OUT_CLASS, "Outputs", "Capital per risk class, component and scenario.")
            Case 19
                NavEntry = Array("OUTPUTS", SH_OUT_BUCKETS, "Outputs", "Kb and Sb per bucket.")
            Case 20
                NavEntry = Array("OUTPUTS", SH_OUT_DRC, "Outputs", "Net jump-to-default per obligor, and the charge per bucket.")
            Case 21
                NavEntry = Array("OUTPUTS", SH_LOG, "Outputs", "Every rejected input row with the reason, warnings and the run status.")
            Case 22
                NavEntry = Array("TESTS & REFERENCE", SH_TEST, "Tests", "Expected values of the sample portfolio, actual values and PASS or FAIL per item, with the area and source of each.")
            Case Else
                NavEntry = Array("TESTS & REFERENCE", SH_README, "Reference", "How to use the workbook, and the input conventions.")
        End Select

End Function

Public Sub SetNavigationVisible(ByVal shown As Boolean)

'
'------------------------------------------------------------------------------
'                       SHOW OR HIDE THE NAVIGATION BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Shows or hides the navigation bar of every sheet.
'
' WHY THIS EXISTS
'   The ribbon's toggle calls it (#353).
'
' INPUTS
'   shown
'     True shows the bar, False hides it
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens a write window, then for each managed sheet that exists sets the visibility of
'   every shape whose name starts with "Nav", then closes the write window.
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next) while the shapes are set.
'
' DEPENDENCIES
'   BeginWrite, EndWrite, GetSheet, SheetExists (CORE_Util); ManagedSheets (CORE_Config).
'
' NOTES
'   A shape is locked on a protected sheet, hence the write window.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.SetNavigationVisible"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim Index                   As Long         'Loop counter
    Dim ShapeIndex              As Long         'Shape being set
    Dim CurrentSheet            As Object       'Sheet being set

'------------------------------------------------------------------------------
' OPEN THE WRITE WINDOW
'------------------------------------------------------------------------------
    'Open the write window: a shape is locked on a protected sheet
        BeginWrite

'------------------------------------------------------------------------------
' SET THE VISIBILITY
'------------------------------------------------------------------------------
    'Ignore errors from the host
        On Error Resume Next
    'Set the visibility of the navigation shapes of each managed sheet
        SheetNames = ManagedSheets()
        For Index = LBound(SheetNames) To UBound(SheetNames)
            If SheetExists(CStr(SheetNames(Index))) Then
                Set CurrentSheet = GetSheet(CStr(SheetNames(Index)))
                For ShapeIndex = 1 To CurrentSheet.Shapes.Count
                    If Left$(CurrentSheet.Shapes(ShapeIndex).Name, 3) = "Nav" Then
                        CurrentSheet.Shapes(ShapeIndex).Visible = shown
                    End If
                Next ShapeIndex
            End If
        Next Index
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' CLOSE THE WRITE WINDOW
'------------------------------------------------------------------------------
    'Close the write window
        EndWrite

End Sub

Public Function NavigationVisible() As Boolean

'
'------------------------------------------------------------------------------
'                          READ THE NAVIGATION STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True while the navigation bar is shown.
'
' WHY THIS EXISTS
'   Read from the marker of Home, so a saved workbook with the bar hidden opens with the
'   toggle released.
'
' INPUTS
'   None
'
' RETURNS
'   The visibility of the shape Nav_Current of Home; True when it cannot be read.
'
' BEHAVIOR
'   Starts at True, then reads the Visible property of the shape.
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   GetSheet (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ THE MARKER
'------------------------------------------------------------------------------
    'Assume the bar is shown
        NavigationVisible = True
    'Read the visibility of the marker; keep True when it cannot be read
        On Error Resume Next
        NavigationVisible = GetSheet(SH_HOME).Shapes("Nav_Current").Visible
        On Error GoTo 0

End Function

Public Sub OpenWriteWindow()

'
'------------------------------------------------------------------------------
'                             OPEN A WRITE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Lifts the protection of the sheets and the workbook structure until CloseWriteWindow
'   restores it (#349).
'
' WHY THIS EXISTS
'   For the build script, which deletes the shapes of a template.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls BeginWrite.
'
' ERROR POLICY
'   Raises what BeginWrite raises.
'
' DEPENDENCIES
'   BeginWrite (CORE_Util).
'
' NOTES
'   Every call must be matched by one call of CloseWriteWindow.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' OPEN THE WINDOW
'------------------------------------------------------------------------------
    'Open the write window
        BeginWrite

End Sub

Public Sub CloseWriteWindow()

'
'------------------------------------------------------------------------------
'                             CLOSE A WRITE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Closes a write window opened with OpenWriteWindow and restores the protection.
'
' WHY THIS EXISTS
'   For a caller outside the workbook, such as the build script.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls EndWrite.
'
' ERROR POLICY
'   None: EndWrite does not raise.
'
' DEPENDENCIES
'   EndWrite (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLOSE THE WINDOW
'------------------------------------------------------------------------------
    'Close the write window
        EndWrite

End Sub

Public Sub AddNavigationBars()

'
'------------------------------------------------------------------------------
'                           ADD THE NAVIGATION BARS
'------------------------------------------------------------------------------
' PURPOSE
'   Adds the bar of links in column A of every sheet.
'
' WHY THIS EXISTS
'   The bar is made of shapes, which the template does not hold: Excel adds them at build
'   time, and a sheet shows its own entry as the marker.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens a write window and turns alerts off. Adds the bar to each managed sheet that
'   exists, activates Home, turns alerts on and closes the write window.
'
' ERROR POLICY
'   No handler: an error leaves the procedure before alerts are turned on and the write
'   window is closed.
'
' DEPENDENCIES
'   BeginWrite, EndWrite, GetSheet, SheetExists (CORE_Util); SetAlerts, AddNavigationBar.
'
' NOTES
'   Excel only. tools/Build-Workbook.ps1 runs it after FinishWorkbook.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.AddNavigationBars"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim Index                   As Long         'Loop counter

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Open the write window
        BeginWrite
    'Turn alerts and screen updating off
        SetAlerts False

'------------------------------------------------------------------------------
' ADD THE BARS
'------------------------------------------------------------------------------
    'Add the bar to each managed sheet that exists
        SheetNames = ManagedSheets()
        For Index = LBound(SheetNames) To UBound(SheetNames)
            If SheetExists(CStr(SheetNames(Index))) Then AddNavigationBar GetSheet(CStr(SheetNames(Index)))
        Next Index

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Activate Home; ignore a host that refuses
        On Error Resume Next
        ThisWorkbook.Worksheets(SH_HOME).Activate
        On Error GoTo 0
    'Turn alerts and screen updating back on
        SetAlerts True
    'Close the write window
        EndWrite

End Sub

Private Sub AddNavigationBar(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                      ADD THE NAVIGATION BAR OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Adds to a sheet the shapes of the navigation bar: a heading for each section, an entry
'   for each sheet and a strip of the role color at the left edge of each entry.
'
' WHY THIS EXISTS
'   Every sheet shows the same bar; the entry of the sheet itself is the marker
'   (Nav_Current).
'
' INPUTS
'   ws
'     Sheet that receives the bar
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Removes the bar the sheet already has. Starts in column A, below every frozen row, and
'   walks the entries of NavEntry from the second to the last, then Home, so that the bar
'   ends on START. For each entry: adds a section heading when the section changes, adds the
'   entry as a rectangle with its caption and, unless it is the current sheet, a hyperlink
'   to cell A1 of that sheet, then adds the strip. The current sheet is drawn in the title
'   bar color with white bold text and an amber strip.
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   RemoveNavigation, NavEntry, RoleColor; ManagedSheets and the palette (CORE_Config).
'
' NOTES
'   Shape names: Nav_Current, Nav_<sheet>, Nav_<sheet>_Mark and NavSection_<section>.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.AddNavigationBar"

    Dim Position                As Long         'Loop position, 2 to NAV_COUNT + 1
    Dim EntryNumber             As Long         'Entry of NavEntry
    Dim Entry                   As Variant      'Navigation entry
    Dim SectionName             As String       'Section heading written last
    Dim NavShape                As Object       'Heading or entry shape
    Dim StripShape              As Object       'Strip at the edge of the entry
    Dim BarLeft                 As Double       'Left edge of the shapes
    Dim BarTop                  As Double       'Top of the next shape
    Dim BarWidth                As Double       'Width of the shapes
    Dim IsCurrent               As Boolean      'True for the entry of this sheet
    Dim StripColor              As Long         'Color of the strip

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Ignore errors from the host
        On Error Resume Next
    'Remove the bar the sheet already has
        RemoveNavigation ws
    'Place the bar in column A, below every frozen row
        BarLeft = ws.Cells(1, 1).Left + 6
        BarWidth = ws.Columns(1).Width - 12
        BarTop = ws.Cells(LAYOUT_FIRST_DATA_ROW + 1, 1).Top  'below every frozen row

'------------------------------------------------------------------------------
' ADD THE ENTRIES
'------------------------------------------------------------------------------
    'Add one entry per sheet, with a heading where the section changes
        For Position = 2 To NAV_COUNT + 1  'Home, the first entry, comes last: the bar ends on START
            EntryNumber = IIf(Position > NAV_COUNT, 1, Position)
            Entry = NavEntry(EntryNumber)
            If CStr(Entry(0)) <> SectionName Then
                SectionName = CStr(Entry(0))
                If Position > 2 Then BarTop = BarTop + 6
                Set NavShape = ws.Shapes.AddTextbox(1, BarLeft, BarTop, BarWidth, 12.75)
                NavShape.Name = "NavSection_" & Replace(SectionName, " ", "_")
                NavShape.Placement = 3  'xlFreeFloating
                NavShape.Fill.Visible = 0
                NavShape.Line.Visible = 0
                NavShape.TextFrame.MarginLeft = 3
                NavShape.TextFrame.MarginTop = 0
                NavShape.TextFrame.MarginBottom = 1.5
                NavShape.TextFrame.VerticalAlignment = -4107  'xlBottom
                NavShape.TextFrame.Characters.Text = SectionName
                NavShape.TextFrame.Characters.Font.Name = FONT_NAME
                NavShape.TextFrame.Characters.Font.Size = 7
                NavShape.TextFrame.Characters.Font.Bold = True
                NavShape.TextFrame.Characters.Font.Color = COLOR_BRAND_TEXT
                BarTop = BarTop + 13
            End If
            IsCurrent = (CStr(Entry(1)) = ws.Name)
            Set NavShape = ws.Shapes.AddShape(1, BarLeft, BarTop, BarWidth, 15.75)  'msoShapeRectangle
            NavShape.Name = IIf(IsCurrent, "Nav_Current", "Nav_" & CStr(Entry(1)))
            NavShape.Placement = 3
            NavShape.Fill.ForeColor.RGB = IIf(IsCurrent, COLOR_TITLE_BAR, COLOR_NAV_ENTRY)
            NavShape.Line.Visible = 0
            NavShape.TextFrame.MarginLeft = 9
            NavShape.TextFrame.MarginTop = 0
            NavShape.TextFrame.MarginBottom = 0
            NavShape.TextFrame.Characters.Text = CStr(Entry(1))
            NavShape.TextFrame.Characters.Font.Name = FONT_NAME
            NavShape.TextFrame.Characters.Font.Size = 9
            NavShape.TextFrame.Characters.Font.Bold = IsCurrent
            NavShape.TextFrame.Characters.Font.Color = IIf(IsCurrent, COLOR_WHITE, COLOR_NAV_TEXT)
            NavShape.TextFrame.HorizontalAlignment = -4131  'xlLeft
            NavShape.TextFrame.VerticalAlignment = -4108  'xlCenter
            If Not IsCurrent Then
                ws.Hyperlinks.Add NavShape, "", "'" & CStr(Entry(1)) & "'!A1", "Go to " & CStr(Entry(1))
            End If
            'A strip of the sheet's role color at the entry's left edge; amber on the current sheet
            StripColor = IIf(CStr(Entry(1)) = SH_HOME Or CStr(Entry(1)) = SH_README, RGB(143, 170, 220), _
                RoleColor(CStr(Entry(1))))
            Set StripShape = ws.Shapes.AddShape(1, BarLeft, BarTop, 2.25, 15.75)
            StripShape.Name = "Nav_" & CStr(Entry(1)) & "_Mark"
            StripShape.Placement = 3
            StripShape.Fill.ForeColor.RGB = IIf(IsCurrent, COLOR_BRAND_TEXT, StripColor)
            StripShape.Line.Visible = 0
            BarTop = BarTop + 17
        Next Position
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub RemoveNavigation(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                     REMOVE THE NAVIGATION BAR OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes the shapes of a sheet whose name starts with "Nav".
'
' INPUTS
'   ws
'     Sheet to clear
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.RemoveNavigation"

    Dim Index                   As Long         'Loop counter

'------------------------------------------------------------------------------
' REMOVE
'------------------------------------------------------------------------------
    'Ignore errors from the host
        On Error Resume Next
    'Delete the navigation shapes, from the last to the first
        For Index = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(Index).Name, 3) = "Nav" Then ws.Shapes(Index).Delete
        Next Index
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                          CONTROL AND OUTPUT SHEETS
'
'------------------------------------------------------------------------------
'

Private Sub BuildControl()

'
'------------------------------------------------------------------------------
'                           BUILD THE CONTROL SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the Control sheet: the frame and the table of run settings.
'
' WHY THIS EXISTS
'   The run reads its settings from this table.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the frame and the column widths, then the [SETTINGS] table with the reporting
'   currency and four national-discretion flags, all with their defaults and references. The
'   Value column is the editable one and is centered; repeated keys are highlighted.
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the text writes raise normally.
'
' DEPENDENCIES
'   ApplyFrame, SetWidths, PutTable, PutRow, FormatParamRows.
'
' NOTES
'   The flags default to FALSE and the reporting currency to EUR.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.BuildControl"

    Dim ControlSheet            As Object       'Control sheet
    Dim RowIndex                As Long         'Row being written

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Locate the Control sheet
        Set ControlSheet = GetSheet(SH_CONTROL)
    'Write the frame and set the column widths
        ApplyFrame ControlSheet, "Control", "Run settings and actions.", 6
        SetWidths ControlSheet, Array(34, 14, 125, 16)

'------------------------------------------------------------------------------
' WRITE THE SETTINGS
'------------------------------------------------------------------------------
    'Write the settings heading and the header row
        RowIndex = PutTable(ControlSheet, LAYOUT_HEADER_ROW, "SETTINGS", "Run settings", Array("Code", "Value", "Description", "Reference"))
    'Write the five settings with their defaults and references
        PutRow ControlSheet, RowIndex, Array("REPORTING_CCY", "EUR", "Reporting currency; FX sensitivities are against this currency", "MAR21.14")
        PutRow ControlSheet, RowIndex + 1, Array("GIRR_SPECIFIED_CCY_DISCRETION", False, "TRUE: GIRR delta risk weights (tenor, inflation and cross-currency basis) divided by sqrt(2) for the specified currencies and the reporting currency", "MAR21.44")
        PutRow ControlSheet, RowIndex + 2, Array("FX_SPECIFIED_PAIR_DISCRETION", False, "TRUE: FX risk weight divided by sqrt(2) for specified pairs and their first-order crosses", "MAR21.88")
        PutRow ControlSheet, RowIndex + 3, Array("FX_CURVATURE_DIVISOR", False, "TRUE: FX curvature CVRs divided by 1.5 for options that do not reference the reporting currency (In_Curvature References reporting ccy not TRUE)", "MAR21.98")
        PutRow ControlSheet, RowIndex + 4, Array("FX_CURVATURE_DIVISOR_ALL", False, "TRUE: all FX curvature CVRs divided by 1.5; needs supervisory approval and curvature sensitivities for all currencies, including the reporting currency", "MAR21.98")
    'Format the rows: the Value column is editable and repeated keys are highlighted
        FormatParamRows ControlSheet, RowIndex, RowIndex + 4, 4, 2, 2, "", True
    'Center the Value column; ignore formatting errors
        On Error Resume Next
        ControlSheet.Range(TableCell(ControlSheet, RowIndex, 2), _
            TableCell(ControlSheet, RowIndex + 4, 2)).HorizontalAlignment = -4108
        On Error GoTo 0

End Sub

Private Sub AddWorkbookButtons()

'
'------------------------------------------------------------------------------
'                       ADD THE BUTTONS OF THE WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Removes and recreates the buttons of every managed sheet that exists.
'
' BEHAVIOR
'   Calls AddButtonRow for each managed sheet that exists; AddButtonRow removes the sheet's
'   buttons first.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.AddWorkbookButtons"

    Dim SheetNames              As Variant      'Names of the managed sheets
    Dim Index                   As Long         'Loop counter

'------------------------------------------------------------------------------
' ADD THE BUTTONS
'------------------------------------------------------------------------------
    'Read the managed sheet names
        SheetNames = ManagedSheets()
    'Add the button row to each managed sheet that exists
        For Index = LBound(SheetNames) To UBound(SheetNames)
            If SheetExists(CStr(SheetNames(Index))) Then AddButtonRow GetSheet(CStr(SheetNames(Index)))
        Next Index

End Sub

Private Sub BuildOutputSheets()

'
'------------------------------------------------------------------------------
'                           BUILD THE OUTPUT SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the frame and the placeholder of the six output sheets: Results, Out_Summary,
'   Out_RiskClass, Out_Buckets, Out_DRC and Log.
'
' WHY THIS EXISTS
'   The run rewrites these sheets; until then they show a placeholder.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   For each sheet: writes the frame with its title and subtitle, sets the column widths and
'   writes the output placeholder.
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the text writes raise normally.
'
' DEPENDENCIES
'   ApplyFrame, SetWidths; WriteOutputPlaceholder (CORE_Util).
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
    Const PROC_NAME             As String = "M_Setup.BuildOutputSheets"

    Dim OutputSheet             As Object       'Output sheet being built

'------------------------------------------------------------------------------
' BUILD RESULTS
'------------------------------------------------------------------------------
    'Build the Results sheet
        Set OutputSheet = GetSheet(SH_RESULTS)
        ApplyFrame OutputSheet, "Results", "Capital by risk class, component and correlation scenario (MAR21.7), reporting currency.", 9
        SetWidths OutputSheet, Array(26, 14, 16, 16, 16)
        WriteOutputPlaceholder OutputSheet

'------------------------------------------------------------------------------
' BUILD OUT_SUMMARY
'------------------------------------------------------------------------------
    'Build the Out_Summary sheet
        Set OutputSheet = GetSheet(SH_OUT_SUMMARY)
        ApplyFrame OutputSheet, "Out_Summary", "Key results of the last run, read by the regression test and automation.", 5
        SetWidths OutputSheet, Array(18, 20)
        WriteOutputPlaceholder OutputSheet

'------------------------------------------------------------------------------
' BUILD OUT_RISKCLASS
'------------------------------------------------------------------------------
    'Build the Out_RiskClass sheet
        Set OutputSheet = GetSheet(SH_OUT_CLASS)
        ApplyFrame OutputSheet, "Out_RiskClass", "Capital per risk class, component and scenario; Alternative Sb used = Y when the alternative Sb of MAR21.4(5)(b) applied.", 9
        SetWidths OutputSheet, Array(12, 13, 11, 16, 9, 13)
        WriteOutputPlaceholder OutputSheet

'------------------------------------------------------------------------------
' BUILD OUT_BUCKETS
'------------------------------------------------------------------------------
    'Build the Out_Buckets sheet
        Set OutputSheet = GetSheet(SH_OUT_BUCKETS)
        ApplyFrame OutputSheet, "Out_Buckets", "Kb and Sb per bucket. Notes: ALT_SB alternative Sb, UP/DOWN selected curvature scenario, OTHER other-sector bucket.", 11
        SetWidths OutputSheet, Array(12, 13, 11, 9, 16, 16, 11, 16)
        WriteOutputPlaceholder OutputSheet

'------------------------------------------------------------------------------
' BUILD OUT_DRC
'------------------------------------------------------------------------------
    'Build the Out_DRC sheet
        Set OutputSheet = GetSheet(SH_OUT_DRC)
        ApplyFrame OutputSheet, "Out_DRC", "Net jump-to-default per obligor after seniority netting, and the hedge benefit ratio and charge per bucket (MAR22).", 18
        SetWidths OutputSheet, Array(16, 9, 13, 9, 14, 14, 3, 9, 14, 14, 10, 14, 14, 14)
        WriteOutputPlaceholder OutputSheet

'------------------------------------------------------------------------------
' BUILD LOG
'------------------------------------------------------------------------------
    'Build the Log sheet
        Set OutputSheet = GetSheet(SH_LOG)
        ApplyFrame OutputSheet, "Log", "Every rejected input row with the reason, warnings and the run status; Row and Reference name the row and its ID.", 8
        SetWidths OutputSheet, Array(9, 24, 110, 7, 24)
        WriteOutputPlaceholder OutputSheet

End Sub

'
'------------------------------------------------------------------------------
'
'                                 INPUT SHEETS
'
'------------------------------------------------------------------------------
'

Private Sub BuildInputSheets()

'
'------------------------------------------------------------------------------
'                            BUILD THE INPUT SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the five input sheets: In_Delta, In_Vega, In_Curvature, In_DRC and In_RRAO.
'
' WHY THIS EXISTS
'   Each sheet differs in its subtitle, column captions, column widths and number formats.
'
' BEHAVIOR
'   Calls BuildInputSheet once per sheet with its subtitle, captions, widths and formats.
'
' DEPENDENCIES
'   BuildInputSheet; FMT_AMOUNT, FMT_YEARS.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BUILD THE SHEETS
'------------------------------------------------------------------------------
    'Build the delta input sheet
        BuildInputSheet SH_IN_DELTA, "Delta sensitivities in the reporting currency; rows with the same risk factor are netted (MAR21.4(2)).", _
            Array("ID", "RiskClass", "Bucket", "Qualifier", "SubType", "Tenor", "Sensitivity", "Comment"), _
            Array(10, 11, 9, 18, 13, 9, 16, 34), Array("", "", "", "", "", FMT_YEARS, FMT_AMOUNT, "")
    'Build the vega input sheet
        BuildInputSheet SH_IN_VEGA, "Vega x implied volatility per option maturity (and GIRR YIELD underlying maturity; GIRR INFLATION and XCCY: none), MAR21.25 and MAR21.8 FAQ4.", _
            Array("ID", "RiskClass", "Bucket", "Qualifier", "SubType", "OptionMaturity", "UnderlyingMaturity", "Sensitivity", "Comment"), _
            Array(10, 11, 9, 18, 13, 18, 22, 16, 34), Array("", "", "", "", "", FMT_YEARS, FMT_YEARS, FMT_AMOUNT, "")
    'Build the curvature input sheet
        BuildInputSheet SH_IN_CURV, "Curvature risk CVR up and CVR down per risk factor, net across instruments, positive for a loss and before the FX divisor; shocks per MAR21.98-99 (MAR21.5).", _
            Array("ID", "RiskClass", "Bucket", "Qualifier", "SubType", "CVR_Up", "CVR_Down", "RefReportingCcy", "Comment"), _
            Array(10, 11, 9, 18, 13, 14, 14, 20, 34), Array("", "", "", "", "", FMT_AMOUNT, FMT_AMOUNT, "", "")
    'Build the DRC input sheet
        BuildInputSheet SH_IN_DRC, "Positions for the default risk charge: signed notional and market value, maturity in years, optional LGD override (MAR22).", _
            Array("ID", "Obligor", "Bucket", "CreditQuality", "Seniority", "Notional", "MarketValue", "Maturity", "LGD", "Comment"), _
            Array(10, 16, 9, 14, 13, 14, 14, 10, 8, 30), _
            Array("", "", "", "", "", FMT_AMOUNT, FMT_AMOUNT, FMT_YEARS, FMT_YEARS, "")
    'Build the RRAO input sheet
        BuildInputSheet SH_IN_RRAO, "Gross notional of instruments with residual risks: EXOTIC or OTHER (MAR23).", _
            Array("ID", "Category", "GrossNotional", "Comment"), Array(10, 12, 16, 50), Array("", "", FMT_AMOUNT, "")

End Sub

Private Sub BuildInputSheet( _
    ByVal sheetName As String, _
    ByVal subtitle As String, _
    ByVal headers As Variant, _
    ByVal widths As Variant, _
    ByVal formats As Variant)
'
'------------------------------------------------------------------------------
'                             BUILD AN INPUT SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Builds one input sheet: frame, column widths, header row, INPUT_ROWS formatted input
'   rows, number formats and an autofilter.
'
' WHY THIS EXISTS
'   The five input sheets share one layout and differ only in their subtitle, captions,
'   widths and formats.
'
' INPUTS
'   sheetName
'     Name of the input sheet
'
'   subtitle
'     Text of the subtitle row
'
'   headers
'     Array of the column captions
'
'   widths
'     Array of the column widths
'
'   formats
'     Array with the number format of each column, or an empty text for none
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the contents and the formats of the sheet (the data validation stays). Writes the
'   frame, widths and header row, formats the input rows as input cells, applies the number
'   formats and adds the autofilter. Centers the risk class and bucket columns (the second
'   and third).
'
' ERROR POLICY
'   Formatting is best effort (On Error Resume Next); the writes raise normally.
'
' DEPENDENCIES
'   ClearCells, ApplyFrame, SetWidths; WriteTableHeader, FormatDataRows, TableCell
'   (CORE_Util).
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
    Const PROC_NAME             As String = "M_Setup.BuildInputSheet"

    Dim InputSheet              As Object       'Input sheet being built
    Dim ColumnCount             As Long         'Number of columns
    Dim ColumnIndex             As Long         'Column being formatted

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Locate the sheet
        Set InputSheet = GetSheet(sheetName)
    'Clear the contents and formats; the data validation stays
        ClearCells InputSheet
    'Count the columns from the captions
        ColumnCount = UBound(headers) - LBound(headers) + 1
    'Write the frame and set the column widths
        ApplyFrame InputSheet, sheetName, subtitle, LAYOUT_FIRST_COL + ColumnCount - 1
        SetWidths InputSheet, widths
    'Write the header row
        WriteTableHeader InputSheet, LAYOUT_HEADER_ROW, headers, COLOR_INPUT_HEADER
    'Format the input rows as input cells
        FormatDataRows InputSheet, LAYOUT_FIRST_DATA_ROW, LAYOUT_FIRST_DATA_ROW + INPUT_ROWS - 1, ColumnCount, 0, _
            "", True

'------------------------------------------------------------------------------
' FORMAT THE COLUMNS
'------------------------------------------------------------------------------
    'Ignore formatting errors
        On Error Resume Next
    'Apply the number format of each column that has one
        For ColumnIndex = 1 To ColumnCount
            If formats(LBound(formats) + ColumnIndex - 1) <> vbNullString Then
                InputSheet.Range(TableCell(InputSheet, LAYOUT_FIRST_DATA_ROW, ColumnIndex), _
                    TableCell(InputSheet, LAYOUT_FIRST_DATA_ROW + INPUT_ROWS - 1, _
                    ColumnIndex)).NumberFormat = formats(LBound(formats) + ColumnIndex - 1)
            End If
        Next ColumnIndex
    'Add the autofilter to the header row
        InputSheet.Range(TableCell(InputSheet, LAYOUT_HEADER_ROW, 1), _
            TableCell(InputSheet, LAYOUT_HEADER_ROW, ColumnCount)).AutoFilter
    'Center the short codes, risk class and bucket
        InputSheet.Range(TableCell(InputSheet, LAYOUT_FIRST_DATA_ROW, 2), _
            TableCell(InputSheet, LAYOUT_FIRST_DATA_ROW + INPUT_ROWS - 1, 3)).HorizontalAlignment = -4108
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                               PARAMETER SHEETS
'
'------------------------------------------------------------------------------
'

Private Sub WriteParameters()

'
'------------------------------------------------------------------------------
'                          WRITE THE PARAMETER SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the default regulatory parameters to the eight P_* sheets.
'
' BEHAVIOR
'   Calls the Write procedure of each sheet in turn: General, GIRR, CSR, equity, commodity,
'   FX, DRC, RRAO.
'
' DEPENDENCIES
'   WriteParamGeneral, WriteParamGirr, WriteParamCsr, WriteParamEquity, WriteParamCommodity,
'   WriteParamFx, WriteParamDrc, WriteParamRrao.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE THE SHEETS
'------------------------------------------------------------------------------
    'Write the parameter sheets one after the other
        WriteParamGeneral
        WriteParamGirr
        WriteParamCsr
        WriteParamEquity
        WriteParamCommodity
        WriteParamFx
        WriteParamDrc
        WriteParamRrao

End Sub

Private Function ParamSheet( _
    ByVal sheetName As String, _
    ByVal subtitle As String, _
    ByVal lastCol As Long) As Object
'
'------------------------------------------------------------------------------
'                          PREPARE A PARAMETER SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Clears a parameter sheet, writes its frame and returns it.
'
' INPUTS
'   sheetName
'     Name of the parameter sheet
'
'   subtitle
'     Text of the subtitle row
'
'   lastCol
'     Last column of the brand bar and the title bar
'
' RETURNS
'   The worksheet.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.ParamSheet"

    Dim TargetSheet             As Object       'Parameter sheet

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Locate the sheet
        Set TargetSheet = GetSheet(sheetName)
    'Clear the contents and formats
        ClearCells TargetSheet
    'Write the frame
        ApplyFrame TargetSheet, sheetName, subtitle, lastCol

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the sheet
        Set ParamSheet = TargetSheet

End Function

Private Sub WriteParamGeneral()

'
'------------------------------------------------------------------------------
'                         WRITE THE GENERAL PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the P_General sheet: correlation scenarios, vega parameters, vega liquidity
'   horizons and regulatory vertices.
'
' WHY THIS EXISTS
'   The engine reads these tagged tables from the sheet; the defaults are written here.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes four tagged tables:
'     - CORR_SCENARIOS: the high and low correlation multipliers (MAR21.6)
'     - VEGA: the vega risk weight sigma and the option maturity correlation parameter
'       (MAR21.92, MAR21.93)
'     - VEGA_LH: the vega liquidity horizons by risk class and bucket (MAR21.92)
'     - VERTICES: the tenors of GIRR, CSR, COM and the vega maturities (MAR21.8, MAR21.9,
'       MAR21.13)
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the cell writes raise normally.
'
' DEPENDENCIES
'   ParamSheet, SetWidths, PutTable, PutRow, FormatParamRows.
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
    Const PROC_NAME             As String = "M_Setup.WriteParamGeneral"

    Dim GeneralSheet            As Object       'P_General sheet
    Dim RowIndex                As Long         'Row being written

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Write the frame of the sheet and set the column widths
        Set GeneralSheet = ParamSheet(SH_P_GENERAL, "Correlation scenarios, vega parameters and regulatory vertices. Codes in column C are read by the engine.", 7)
        SetWidths GeneralSheet, Array(22, 30, 10, 60)

'------------------------------------------------------------------------------
' CORR_SCENARIOS
'------------------------------------------------------------------------------
    'Write the correlation scenario multipliers
        RowIndex = PutTable(GeneralSheet, LAYOUT_HEADER_ROW, "CORR_SCENARIOS", "Correlation scenarios", Array("Code", "Value", "Reference"))
        PutRow GeneralSheet, RowIndex, Array("HIGH_MULTIPLIER", 1.25, "MAR21.6(2): rho_high = min(1.25 x rho, 100%)")
        PutRow GeneralSheet, RowIndex + 1, Array("LOW_MULTIPLIER", 0.75, "MAR21.6(3): rho_low = max(2 x rho - 100%, 75% x rho)")
        FormatParamRows GeneralSheet, RowIndex, RowIndex + 1, 3, 2, 2, "", True

'------------------------------------------------------------------------------
' VEGA
'------------------------------------------------------------------------------
    'Write the vega risk weight and maturity correlation parameters
        RowIndex = PutTable(GeneralSheet, RowIndex + 3, "VEGA", "Vega risk weight and maturity correlation", Array("Code", "Value", "Reference"))
        PutRow GeneralSheet, RowIndex, Array("RW_SIGMA", 0.55, "MAR21.92: RW = min(RW_sigma x sqrt(LH / 10), 100%)")
        PutRow GeneralSheet, RowIndex + 1, Array("ALPHA", 0.01, "MAR21.93: option maturity correlation parameter")
        FormatParamRows GeneralSheet, RowIndex, RowIndex + 1, 3, 2, 2, "2", True

'------------------------------------------------------------------------------
' VEGA_LH
'------------------------------------------------------------------------------
    'Write the vega liquidity horizons
        RowIndex = PutTable(GeneralSheet, RowIndex + 3, "VEGA_LH", "Vega liquidity horizons (days)", Array("RiskClass", "Bucket", "LH", "Reference"))
        PutRow GeneralSheet, RowIndex, Array("GIRR", "*", 60, "MAR21.92 Table 13")
        PutRow GeneralSheet, RowIndex + 1, Array("CSR_NS", "*", 120, "MAR21.92 Table 13")
        PutRow GeneralSheet, RowIndex + 2, Array("EQ", "*", 20, "MAR21.92: large cap and indices (buckets 1-8, 12, 13)")
        PutRow GeneralSheet, RowIndex + 3, Array("EQ", 9, 60, "MAR21.92: small cap and other")
        PutRow GeneralSheet, RowIndex + 4, Array("EQ", 10, 60, "MAR21.92: small cap and other")
        PutRow GeneralSheet, RowIndex + 5, Array("EQ", 11, 60, "MAR21.92: small cap and other")
        PutRow GeneralSheet, RowIndex + 6, Array("COM", "*", 120, "MAR21.92 Table 13")
        PutRow GeneralSheet, RowIndex + 7, Array("FX", "*", 40, "MAR21.92 Table 13")
        FormatParamRows GeneralSheet, RowIndex, RowIndex + 7, 4, 3, 3, ""

'------------------------------------------------------------------------------
' VERTICES
'------------------------------------------------------------------------------
    'Write the regulatory vertices
        RowIndex = PutTable(GeneralSheet, RowIndex + 9, "VERTICES", "Regulatory vertices (years)", Array("Code", "Value", "Reference"))
        PutRow GeneralSheet, RowIndex, Array("GIRR_TENORS", "0.25,0.5,1,2,3,5,10,15,20,30", "MAR21.8(1)")
        PutRow GeneralSheet, RowIndex + 1, Array("CSR_TENORS", "0.5,1,3,5,10", "MAR21.9(1)")
        PutRow GeneralSheet, RowIndex + 2, Array("COM_TENORS", "0,0.25,0.5,1,2,3,5,10,15,20,30", "MAR21.13(1)")
        PutRow GeneralSheet, RowIndex + 3, Array("VEGA_MATURITIES", "0.5,1,3,5,10", "MAR21.8(4)")
        FormatParamRows GeneralSheet, RowIndex, RowIndex + 3, 3, 2, 2, "", True

End Sub

Private Sub WriteParamGirr()

'
'------------------------------------------------------------------------------
'                          WRITE THE GIRR PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the P_GIRR sheet: delta risk weights by tenor, correlations and other risk
'   weights, and the specified currencies.
'
' WHY THIS EXISTS
'   The engine reads these tagged tables from the sheet; the defaults are written here.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes three tagged tables:
'     - GIRR_RW: the delta risk weights for the ten tenors (MAR21.42 Table 1)
'     - GIRR_PARAMS: THETA, the correlation floor, the correlations, the inflation and
'       cross-currency risk weights, GAMMA and the divisor sqrt(2) shown to four decimals
'       (#206)
'     - GIRR_SPECIFIED_CCY: the seven specified currencies (MAR21.44)
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the cell writes raise normally.
'
' DEPENDENCIES
'   ParamSheet, SetWidths, PutTable, PutRow, FormatParamRows; TableCell (CORE_Util).
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
    Const PROC_NAME             As String = "M_Setup.WriteParamGirr"

    Dim GirrSheet               As Object       'P_GIRR sheet
    Dim RowIndex                As Long         'Row being written
    Dim Index                   As Long         'Loop counter
    Dim Tenors                  As Variant      'Tenors in years
    Dim RiskWeights             As Variant      'Delta risk weight of each tenor
    Dim Currencies              As Variant      'Specified currencies

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Write the frame of the sheet and set the column widths
        Set GirrSheet = ParamSheet(SH_P_GIRR, "General interest rate risk: delta risk weights, correlations and specified currencies.", 7)
        SetWidths GirrSheet, Array(22, 12, 26)

'------------------------------------------------------------------------------
' GIRR_RW
'------------------------------------------------------------------------------
    'Define the tenors and their risk weights
        Tenors = Array(0.25, 0.5, 1, 2, 3, 5, 10, 15, 20, 30)
        RiskWeights = Array(0.017, 0.017, 0.016, 0.013, 0.012, 0.011, 0.011, 0.011, 0.011, 0.011)
    'Write the delta risk weights by tenor
        RowIndex = PutTable(GirrSheet, LAYOUT_HEADER_ROW, "GIRR_RW", "Delta risk weights by tenor (MAR21.42 Table 1)", Array("Tenor", "RW"))
        For Index = 0 To 9
            PutRow GirrSheet, RowIndex + Index, Array(Tenors(Index), RiskWeights(Index))
        Next Index
        FormatParamRows GirrSheet, RowIndex, RowIndex + 9, 2, 2, 2, "2"

'------------------------------------------------------------------------------
' GIRR_PARAMS
'------------------------------------------------------------------------------
    'Write the correlations and the other risk weights
        RowIndex = PutTable(GirrSheet, RowIndex + 11, "GIRR_PARAMS", "Correlations and other risk weights", Array("Code", "Value", "Reference"))
        PutRow GirrSheet, RowIndex, Array("THETA", 0.03, "MAR21.46")
        PutRow GirrSheet, RowIndex + 1, Array("RHO_FLOOR", 0.4, "MAR21.46")
        PutRow GirrSheet, RowIndex + 2, Array("RHO_DIFF_CURVE", 0.999, "MAR21.45, MAR21.47 (also two inflation curves, FAQ)")
        PutRow GirrSheet, RowIndex + 3, Array("RHO_INFLATION", 0.4, "MAR21.48")
        PutRow GirrSheet, RowIndex + 4, Array("RHO_XCCY", 0, "MAR21.49")
        PutRow GirrSheet, RowIndex + 5, Array("RW_INFLATION", 0.016, "MAR21.43")
        PutRow GirrSheet, RowIndex + 6, Array("RW_XCCY", 0.016, "MAR21.43")
        PutRow GirrSheet, RowIndex + 7, Array("GAMMA", 0.5, "MAR21.50")
        PutRow GirrSheet, RowIndex + 8, Array("SPECIFIED_DIVISOR", Sqr(2), "MAR21.44")
        FormatParamRows GirrSheet, RowIndex, RowIndex + 8, 3, 2, 2, "KEY", True
        TableCell(GirrSheet, RowIndex + 8, 2).NumberFormat = "0.0000"  'sqrt(2), shown to four decimals (#206)

'------------------------------------------------------------------------------
' GIRR_SPECIFIED_CCY
'------------------------------------------------------------------------------
    'Write the specified currencies
        Currencies = Array("EUR", "USD", "GBP", "AUD", "JPY", "SEK", "CAD")
        RowIndex = PutTable(GirrSheet, RowIndex + 10, "GIRR_SPECIFIED_CCY", "Specified currencies (MAR21.44)", Array("Currency"))
        For Index = 0 To UBound(Currencies)
            TableCell(GirrSheet, RowIndex + Index, 1).Value = Currencies(Index)
        Next Index
        FormatParamRows GirrSheet, RowIndex, RowIndex + UBound(Currencies), 1, 1, 1, ""

End Sub

Private Sub WriteParamCsr()

'
'------------------------------------------------------------------------------
'                           WRITE THE CSR PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the P_CSR sheet: the buckets with their delta risk weights, the correlations and
'   the sector correlation matrix.
'
' WHY THIS EXISTS
'   The engine reads these tagged tables from the sheet; the defaults are written here.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes three tagged tables:
'     - CSR_BUCKETS: the 18 buckets with description, quality, sector group, type and risk
'       weight (MAR21.51, MAR21.53 Table 4)
'     - CSR_PARAMS: the name, tenor and basis correlations and the rating gamma (MAR21.54 to
'       MAR21.57)
'     - CSR_GAMMA_SECTOR: the cross-bucket sector correlation of the sector groups (MAR21.57
'       Table 5), symmetric
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the cell writes raise normally.
'
' DEPENDENCIES
'   ParamSheet, SetWidths, PutTable, PutRow, FormatParamRows, WrapColumn; TableCell
'   (CORE_Util).
'
' NOTES
'   The bucket description column is wrapped, with the row height set for its lines.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WriteParamCsr"

    Dim CsrSheet                As Object       'P_CSR sheet
    Dim RowIndex                As Long         'Row being written
    Dim RowOffset               As Long         'Matrix row being written
    Dim ColumnOffset            As Long         'Matrix column being written
    Dim SectorGroups            As Variant      'Sector group of each matrix row and column
    Dim SectorMatrix            As Variant      'Sector correlations
    Dim Headers                 As Variant      'Captions of the matrix table

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Write the frame of the sheet and set the column widths
        Set CsrSheet = ParamSheet(SH_P_CSR, "Credit spread risk, non-securitization: buckets, risk weights and correlations.", 15)
        SetWidths CsrSheet, Array(22, 40, 12, 15, 9, 9, 9, 9, 9, 9, 9, 9)

'------------------------------------------------------------------------------
' CSR_BUCKETS
'------------------------------------------------------------------------------
    'Write the buckets and their delta risk weights
        RowIndex = PutTable(CsrSheet, LAYOUT_HEADER_ROW, "CSR_BUCKETS", "Buckets and delta risk weights (MAR21.51, MAR21.53 Table 4)", _
            Array("Bucket", "Description", "Quality", "SectorGroup", "Type", "RW"))
        PutRow CsrSheet, RowIndex, Array(1, "Sovereigns incl. central banks, MDBs", "IG", 1, "NAME", 0.005)
        PutRow CsrSheet, RowIndex + 1, Array(2, "Local government, government-backed non-financials, education, public administration", "IG", 2, "NAME", 0.01)
        PutRow CsrSheet, RowIndex + 2, Array(3, "Financials incl. government-backed financials", "IG", 3, "NAME", 0.05)
        PutRow CsrSheet, RowIndex + 3, Array(4, "Basic materials, energy, industrials, agriculture, manufacturing, mining", "IG", 4, "NAME", 0.03)
        PutRow CsrSheet, RowIndex + 4, Array(5, "Consumer goods and services, transportation and storage, administrative services", "IG", 5, "NAME", 0.03)
        PutRow CsrSheet, RowIndex + 5, Array(6, "Technology, telecommunications", "IG", 6, "NAME", 0.02)
        PutRow CsrSheet, RowIndex + 6, Array(7, "Health care, utilities, professional and technical activities", "IG", 7, "NAME", 0.015)
        PutRow CsrSheet, RowIndex + 7, Array(8, "Covered bonds (1.5% for AA- or better at the bank's discretion, fn 17: not modeled; an override applies to the whole bucket)", "IG", 8, "NAME", 0.025)
        PutRow CsrSheet, RowIndex + 8, Array(9, "Sovereigns incl. central banks, MDBs", "HY", 1, "NAME", 0.02)
        PutRow CsrSheet, RowIndex + 9, Array(10, "Local government, government-backed non-financials, education, public administration", "HY", 2, "NAME", 0.04)
        PutRow CsrSheet, RowIndex + 10, Array(11, "Financials incl. government-backed financials", "HY", 3, "NAME", 0.12)
        PutRow CsrSheet, RowIndex + 11, Array(12, "Basic materials, energy, industrials, agriculture, manufacturing, mining", "HY", 4, "NAME", 0.07)
        PutRow CsrSheet, RowIndex + 12, Array(13, "Consumer goods and services, transportation and storage, administrative services", "HY", 5, "NAME", 0.085)
        PutRow CsrSheet, RowIndex + 13, Array(14, "Technology, telecommunications", "HY", 6, "NAME", 0.055)
        PutRow CsrSheet, RowIndex + 14, Array(15, "Health care, utilities, professional and technical activities", "HY", 7, "NAME", 0.05)
        PutRow CsrSheet, RowIndex + 15, Array(16, "Other sector", "NA", 16, "OTHER", 0.12)
        PutRow CsrSheet, RowIndex + 16, Array(17, "IG indices", "IG", 17, "INDEX", 0.015)
        PutRow CsrSheet, RowIndex + 17, Array(18, "HY indices", "HY", 18, "INDEX", 0.05)
    'Format the rows as percentages where they are rates and wrap the descriptions
        FormatParamRows CsrSheet, RowIndex, RowIndex + 17, 6, 6, 6, "6"
        WrapColumn CsrSheet, RowIndex, RowIndex + 17, 2

'------------------------------------------------------------------------------
' CSR_PARAMS
'------------------------------------------------------------------------------
    'Write the correlations
        RowIndex = PutTable(CsrSheet, RowIndex + 19, "CSR_PARAMS", "Correlations", Array("Code", "Value", "Reference"))
        PutRow CsrSheet, RowIndex, Array("RHO_NAME", 0.35, "MAR21.54")
        PutRow CsrSheet, RowIndex + 1, Array("RHO_NAME_INDEX", 0.8, "MAR21.55 (buckets 17-18)")
        PutRow CsrSheet, RowIndex + 2, Array("RHO_TENOR", 0.65, "MAR21.54")
        PutRow CsrSheet, RowIndex + 3, Array("RHO_BASIS", 0.999, "MAR21.54 (bond vs CDS curve)")
        PutRow CsrSheet, RowIndex + 4, Array("GAMMA_RATING", 0.5, "MAR21.57: IG vs HY/NR, both buckets in 1-15")
        FormatParamRows CsrSheet, RowIndex, RowIndex + 4, 3, 2, 2, "KEY", True

'------------------------------------------------------------------------------
' CSR_GAMMA_SECTOR
'------------------------------------------------------------------------------
    'Define the sector correlation of the cross-bucket correlation (MAR21.57 Table 5), symmetric
        SectorGroups = Array(1, 2, 3, 4, 5, 6, 7, 8, 16, 17, 18)
        SectorMatrix = Array( _
            Array(1, 0.75, 0.1, 0.2, 0.25, 0.2, 0.15, 0.1, 0, 0.45, 0.45), _
            Array(0.75, 1, 0.05, 0.15, 0.2, 0.15, 0.1, 0.1, 0, 0.45, 0.45), _
            Array(0.1, 0.05, 1, 0.05, 0.15, 0.2, 0.05, 0.2, 0, 0.45, 0.45), _
            Array(0.2, 0.15, 0.05, 1, 0.2, 0.25, 0.05, 0.05, 0, 0.45, 0.45), _
            Array(0.25, 0.2, 0.15, 0.2, 1, 0.25, 0.05, 0.15, 0, 0.45, 0.45), _
            Array(0.2, 0.15, 0.2, 0.25, 0.25, 1, 0.05, 0.2, 0, 0.45, 0.45), _
            Array(0.15, 0.1, 0.05, 0.05, 0.05, 0.05, 1, 0.05, 0, 0.45, 0.45), _
            Array(0.1, 0.1, 0.2, 0.05, 0.15, 0.2, 0.05, 1, 0, 0.45, 0.45), _
            Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0), _
            Array(0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0, 1, 0.75), _
            Array(0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0, 0.75, 1))
        Headers = Array("Group", 1, 2, 3, 4, 5, 6, 7, 8, 16, 17, 18)
    'Write the matrix table with its header
        RowIndex = PutTable(CsrSheet, RowIndex + 6, "CSR_GAMMA_SECTOR", "Cross-bucket sector correlation (MAR21.57 Table 5)", Headers)
        For RowOffset = 0 To 10
            TableCell(CsrSheet, RowIndex + RowOffset, 1).Value = SectorGroups(RowOffset)
            For ColumnOffset = 0 To 10
                TableCell(CsrSheet, RowIndex + RowOffset, _
                    ColumnOffset + 2).Value = SectorMatrix(RowOffset)(ColumnOffset)
            Next ColumnOffset
        Next RowOffset
    'Format the rows as percentages
        FormatParamRows CsrSheet, RowIndex, RowIndex + 10, 12, 2, 12, "2,3,4,5,6,7,8,9,10,11,12"

End Sub

Private Sub WriteParamEquity()

'
'------------------------------------------------------------------------------
'                         WRITE THE EQUITY PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the P_EQ sheet: the buckets with spot and repo risk weights, the correlations and
'   the cross-bucket correlation matrix.
'
' WHY THIS EXISTS
'   The engine reads these tagged tables from the sheet; the defaults are written here.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes three tagged tables:
'     - EQ_BUCKETS: the 13 buckets with description, type, spot and repo risk weights and
'       name correlation (MAR21.72, MAR21.77 Table 10)
'     - EQ_PARAMS: the spot-repo correlation (MAR21.78)
'     - EQ_GAMMA: the cross-bucket correlation (MAR21.80): 1 on the diagonal, 0 with bucket
'       11, 0.75 between two index buckets, 0.45 between an index bucket and another, 0.15
'       otherwise
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the cell writes raise normally.
'
' DEPENDENCIES
'   ParamSheet, SetWidths, PutTable, PutRow, FormatParamRows, WrapColumn; TableCell
'   (CORE_Util).
'
' NOTES
'   The bucket description column is wrapped, with the row height set for its lines.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WriteParamEquity"

    Dim EquitySheet             As Object       'P_EQ sheet
    Dim RowIndex                As Long         'Row being written
    Dim BucketRow               As Long         'Bucket of the matrix row
    Dim BucketColumn            As Long         'Bucket of the matrix column
    Dim Correlation             As Double       'Cross-bucket correlation of the pair

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Write the frame of the sheet and set the column widths
        Set EquitySheet = ParamSheet(SH_P_EQ, "Equity risk: buckets, spot and repo risk weights, correlations.", 16)
        SetWidths EquitySheet, Array(18, 40, 12, 10, 10, 13, 9, 9, 9, 9, 9, 9, 9, 9)

'------------------------------------------------------------------------------
' EQ_BUCKETS
'------------------------------------------------------------------------------
    'Write the buckets with their risk weights and correlations
        RowIndex = PutTable(EquitySheet, LAYOUT_HEADER_ROW, "EQ_BUCKETS", "Buckets and delta risk weights (MAR21.72, MAR21.77 Table 10)", _
            Array("Bucket", "Description", "Type", "RW_SPOT", "RW_REPO", "RHO_NAME"))
        PutRow EquitySheet, RowIndex, Array(1, "Large cap, EM: consumer goods/services, transportation, admin, health care, utilities", "NAME", 0.55, 0.0055, 0.15)
        PutRow EquitySheet, RowIndex + 1, Array(2, "Large cap, EM: telecommunications, industrials", "NAME", 0.6, 0.006, 0.15)
        PutRow EquitySheet, RowIndex + 2, Array(3, "Large cap, EM: basic materials, energy, agriculture, manufacturing, mining", "NAME", 0.45, 0.0045, 0.15)
        PutRow EquitySheet, RowIndex + 3, Array(4, "Large cap, EM: financials, real estate, technology", "NAME", 0.55, 0.0055, 0.15)
        PutRow EquitySheet, RowIndex + 4, Array(5, "Large cap, AE: consumer goods/services, transportation, admin, health care, utilities", "NAME", 0.3, 0.003, 0.25)
        PutRow EquitySheet, RowIndex + 5, Array(6, "Large cap, AE: telecommunications, industrials", "NAME", 0.35, 0.0035, 0.25)
        PutRow EquitySheet, RowIndex + 6, Array(7, "Large cap, AE: basic materials, energy, agriculture, manufacturing, mining", "NAME", 0.4, 0.004, 0.25)
        PutRow EquitySheet, RowIndex + 7, Array(8, "Large cap, AE: financials, real estate, technology", "NAME", 0.5, 0.005, 0.25)
        PutRow EquitySheet, RowIndex + 8, Array(9, "Small cap, EM: all sectors", "NAME", 0.7, 0.007, 0.075)
        PutRow EquitySheet, RowIndex + 9, Array(10, "Small cap, AE: all sectors", "NAME", 0.5, 0.005, 0.125)
        PutRow EquitySheet, RowIndex + 10, Array(11, "Other sector", "OTHER", 0.7, 0.007, 0)
        PutRow EquitySheet, RowIndex + 11, Array(12, "Large cap, AE indices", "INDEX", 0.15, 0.0015, 0.8)
        PutRow EquitySheet, RowIndex + 12, Array(13, "Other indices", "INDEX", 0.25, 0.0025, 0.8)
    'Format the rows as percentages and wrap the descriptions
        FormatParamRows EquitySheet, RowIndex, RowIndex + 12, 6, 4, 6, "4,5,6"
        WrapColumn EquitySheet, RowIndex, RowIndex + 12, 2

'------------------------------------------------------------------------------
' EQ_PARAMS
'------------------------------------------------------------------------------
    'Write the spot-repo correlation
        RowIndex = PutTable(EquitySheet, RowIndex + 14, "EQ_PARAMS", "Correlations", Array("Code", "Value", "Reference"))
        PutRow EquitySheet, RowIndex, Array("RHO_SPOT_REPO", 0.999, "MAR21.78")
        FormatParamRows EquitySheet, RowIndex, RowIndex, 3, 2, 2, "KEY", True

'------------------------------------------------------------------------------
' EQ_GAMMA
'------------------------------------------------------------------------------
    'Write the cross-bucket correlation matrix
        RowIndex = PutTable(EquitySheet, RowIndex + 2, "EQ_GAMMA", "Cross-bucket correlation (MAR21.80)", _
            Array("Bucket", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13))
        For BucketRow = 1 To 13
            TableCell(EquitySheet, RowIndex + BucketRow - 1, 1).Value = BucketRow
            For BucketColumn = 1 To 13
                If BucketRow = BucketColumn Then
                    Correlation = 1
                ElseIf BucketRow = 11 Or BucketColumn = 11 Then
                    Correlation = 0
                ElseIf BucketRow >= 12 And BucketColumn >= 12 Then
                    Correlation = 0.75
                ElseIf BucketRow >= 12 Or BucketColumn >= 12 Then
                    Correlation = 0.45
                Else
                    Correlation = 0.15
                End If
                TableCell(EquitySheet, RowIndex + BucketRow - 1, BucketColumn + 1).Value = Correlation
            Next BucketColumn
        Next BucketRow
    'Format the rows as percentages
        FormatParamRows EquitySheet, RowIndex, RowIndex + 12, 14, 2, 14, "2,3,4,5,6,7,8,9,10,11,12,13,14"

End Sub

Private Sub WriteParamCommodity()

'
'------------------------------------------------------------------------------
'                        WRITE THE COMMODITY PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the P_COM sheet: the buckets with risk weights and commodity correlation, the
'   correlations and the cross-bucket correlation matrix.
'
' WHY THIS EXISTS
'   The engine reads these tagged tables from the sheet; the defaults are written here.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes three tagged tables:
'     - COM_BUCKETS: the 11 buckets with description, risk weight and commodity correlation
'       (MAR21.82, MAR21.83)
'     - COM_PARAMS: the tenor and basis correlations (MAR21.83)
'     - COM_GAMMA: the cross-bucket correlation (MAR21.85): 1 on the diagonal, 0 with bucket
'       11, 0.2 otherwise
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the cell writes raise normally.
'
' DEPENDENCIES
'   ParamSheet, SetWidths, PutTable, PutRow, FormatParamRows; TableCell (CORE_Util).
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
    Const PROC_NAME             As String = "M_Setup.WriteParamCommodity"

    Dim CommoditySheet          As Object       'P_COM sheet
    Dim RowIndex                As Long         'Row being written
    Dim BucketRow               As Long         'Bucket of the matrix row
    Dim BucketColumn            As Long         'Bucket of the matrix column
    Dim Correlation             As Double       'Cross-bucket correlation of the pair

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Write the frame of the sheet and set the column widths
        Set CommoditySheet = ParamSheet(SH_P_COM, "Commodity risk: buckets, risk weights and correlations.", 14)
        SetWidths CommoditySheet, Array(16, 45, 12, 10, 9, 9, 9, 9, 9, 9, 9, 9)

'------------------------------------------------------------------------------
' COM_BUCKETS
'------------------------------------------------------------------------------
    'Write the buckets with their risk weights and correlations
        RowIndex = PutTable(CommoditySheet, LAYOUT_HEADER_ROW, "COM_BUCKETS", "Buckets, risk weights and commodity correlation (MAR21.82-83)", _
            Array("Bucket", "Description", "RW", "RHO_CTY"))
        PutRow CommoditySheet, RowIndex, Array(1, "Energy - solid combustibles", 0.3, 0.55)
        PutRow CommoditySheet, RowIndex + 1, Array(2, "Energy - liquid combustibles", 0.35, 0.95)
        PutRow CommoditySheet, RowIndex + 2, Array(3, "Energy - electricity and carbon trading", 0.6, 0.4)
        PutRow CommoditySheet, RowIndex + 3, Array(4, "Freight", 0.8, 0.8)
        PutRow CommoditySheet, RowIndex + 4, Array(5, "Metals - non-precious", 0.4, 0.6)
        PutRow CommoditySheet, RowIndex + 5, Array(6, "Gaseous combustibles", 0.45, 0.65)
        PutRow CommoditySheet, RowIndex + 6, Array(7, "Precious metals (including gold)", 0.2, 0.55)
        PutRow CommoditySheet, RowIndex + 7, Array(8, "Grains and oilseed", 0.35, 0.45)
        PutRow CommoditySheet, RowIndex + 8, Array(9, "Livestock and dairy", 0.25, 0.15)
        PutRow CommoditySheet, RowIndex + 9, Array(10, "Softs and other agriculturals", 0.35, 0.4)
        PutRow CommoditySheet, RowIndex + 10, Array(11, "Other commodity", 0.5, 0.15)
        FormatParamRows CommoditySheet, RowIndex, RowIndex + 10, 4, 3, 4, "3,4"

'------------------------------------------------------------------------------
' COM_PARAMS
'------------------------------------------------------------------------------
    'Write the tenor and basis correlations
        RowIndex = PutTable(CommoditySheet, RowIndex + 12, "COM_PARAMS", "Correlations", Array("Code", "Value", "Reference"))
        PutRow CommoditySheet, RowIndex, Array("RHO_TENOR", 0.99, "MAR21.83")
        PutRow CommoditySheet, RowIndex + 1, Array("RHO_BASIS", 0.999, "MAR21.83 (delivery location)")
        FormatParamRows CommoditySheet, RowIndex, RowIndex + 1, 3, 2, 2, "KEY", True

'------------------------------------------------------------------------------
' COM_GAMMA
'------------------------------------------------------------------------------
    'Write the cross-bucket correlation matrix
        RowIndex = PutTable(CommoditySheet, RowIndex + 3, "COM_GAMMA", "Cross-bucket correlation (MAR21.85)", _
            Array("Bucket", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11))
        For BucketRow = 1 To 11
            TableCell(CommoditySheet, RowIndex + BucketRow - 1, 1).Value = BucketRow
            For BucketColumn = 1 To 11
                If BucketRow = BucketColumn Then
                    Correlation = 1
                ElseIf BucketRow = 11 Or BucketColumn = 11 Then
                    Correlation = 0
                Else
                    Correlation = 0.2
                End If
                TableCell(CommoditySheet, RowIndex + BucketRow - 1, BucketColumn + 1).Value = Correlation
            Next BucketColumn
        Next BucketRow
    'Format the rows as percentages
        FormatParamRows CommoditySheet, RowIndex, RowIndex + 10, 12, 2, 12, "2,3,4,5,6,7,8,9,10,11,12"

End Sub

Private Sub WriteParamFx()

'
'------------------------------------------------------------------------------
'                           WRITE THE FX PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the P_FX sheet: the risk weight and correlation parameters and the specified
'   currency pairs.
'
' WHY THIS EXISTS
'   The engine reads these tagged tables from the sheet; the defaults are written here.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes two tagged tables:
'     - FX_PARAMS: the risk weight, GAMMA, the specified-pair divisor sqrt(2) shown to four
'       decimals (#206) and the curvature divisor 1.5
'     - FX_SPECIFIED_PAIRS: the 23 specified pairs (MAR21.88)
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the cell writes raise normally.
'
' DEPENDENCIES
'   ParamSheet, SetWidths, PutTable, PutRow, FormatParamRows; TableCell (CORE_Util).
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
    Const PROC_NAME             As String = "M_Setup.WriteParamFx"

    Dim FxSheet                 As Object       'P_FX sheet
    Dim RowIndex                As Long         'Row being written
    Dim Index                   As Long         'Loop counter
    Dim CurrencyPairs           As Variant      'Specified currency pairs

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Write the frame of the sheet and set the column widths
        Set FxSheet = ParamSheet(SH_P_FX, "Foreign exchange risk: risk weight, correlation and specified currency pairs.", 7)
        SetWidths FxSheet, Array(22, 12, 40)

'------------------------------------------------------------------------------
' FX_PARAMS
'------------------------------------------------------------------------------
    'Write the risk weight, correlation and divisors
        RowIndex = PutTable(FxSheet, LAYOUT_HEADER_ROW, "FX_PARAMS", "Risk weight and correlation", Array("Code", "Value", "Reference"))
        PutRow FxSheet, RowIndex, Array("RW", 0.15, "MAR21.87")
        PutRow FxSheet, RowIndex + 1, Array("GAMMA", 0.6, "MAR21.89")
        PutRow FxSheet, RowIndex + 2, Array("SPECIFIED_DIVISOR", Sqr(2), "MAR21.88")
        PutRow FxSheet, RowIndex + 3, Array("CURVATURE_DIVISOR", 1.5, "MAR21.98")
        FormatParamRows FxSheet, RowIndex, RowIndex + 3, 3, 2, 2, "KEY", True
        TableCell(FxSheet, RowIndex + 2, 2).NumberFormat = "0.0000"  'sqrt(2), shown to four decimals (#206)

'------------------------------------------------------------------------------
' FX_SPECIFIED_PAIRS
'------------------------------------------------------------------------------
    'Define the specified currency pairs
        CurrencyPairs = Array("USD/EUR", "USD/JPY", "USD/GBP", "USD/AUD", "USD/CAD", "USD/CHF", "USD/MXN", "USD/CNY", _
            "USD/NZD", "USD/RUB", "USD/HKD", "USD/SGD", "USD/TRY", "USD/KRW", "USD/SEK", "USD/ZAR", _
            "USD/INR", "USD/NOK", "USD/BRL", "EUR/JPY", "EUR/GBP", "EUR/CHF", "JPY/AUD")
    'Write the specified pairs
        RowIndex = PutTable(FxSheet, RowIndex + 5, "FX_SPECIFIED_PAIRS", "Specified pairs; first-order crosses also qualify (MAR21.88)", Array("Pair"))
        For Index = 0 To UBound(CurrencyPairs)
            TableCell(FxSheet, RowIndex + Index, 1).Value = CurrencyPairs(Index)
        Next Index
        FormatParamRows FxSheet, RowIndex, RowIndex + UBound(CurrencyPairs), 1, 1, 1, ""

End Sub

Private Sub WriteParamDrc()

'
'------------------------------------------------------------------------------
'                           WRITE THE DRC PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the P_DRC sheet: risk weights by credit quality, seniority rank and LGD, maturity
'   scaling and buckets.
'
' WHY THIS EXISTS
'   The engine reads these tagged tables from the sheet; the defaults are written here.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes four tagged tables:
'     - DRC_RW: the risk weights by credit quality (MAR22.24 Table 2)
'     - DRC_SENIORITY: the rank and LGD of each seniority (MAR22.12, MAR22.19)
'     - DRC_PARAMS: the minimum and maximum maturity in years (MAR22.18, MAR22.15)
'     - DRC_BUCKETS: the three buckets, CORP, SOV and LGM (MAR22.22)
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the cell writes raise normally.
'
' DEPENDENCIES
'   ParamSheet, SetWidths, PutTable, PutRow, FormatParamRows.
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
    Const PROC_NAME             As String = "M_Setup.WriteParamDrc"

    Dim DrcSheet                As Object       'P_DRC sheet
    Dim RowIndex                As Long         'Row being written

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Write the frame of the sheet and set the column widths
        Set DrcSheet = ParamSheet(SH_P_DRC, "Default risk charge, non-securitization: risk weights, LGD, maturity scaling and buckets.", 7)
        SetWidths DrcSheet, Array(16, 12, 12, 40)  'room for the 11 pt headers (#351)

'------------------------------------------------------------------------------
' DRC_RW
'------------------------------------------------------------------------------
    'Write the risk weights by credit quality
        RowIndex = PutTable(DrcSheet, LAYOUT_HEADER_ROW, "DRC_RW", "Risk weights by credit quality (MAR22.24 Table 2)", Array("CreditQuality", "RW"))
        PutRow DrcSheet, RowIndex, Array("AAA", 0.005)
        PutRow DrcSheet, RowIndex + 1, Array("AA", 0.02)
        PutRow DrcSheet, RowIndex + 2, Array("A", 0.03)
        PutRow DrcSheet, RowIndex + 3, Array("BBB", 0.06)
        PutRow DrcSheet, RowIndex + 4, Array("BB", 0.15)
        PutRow DrcSheet, RowIndex + 5, Array("B", 0.3)
        PutRow DrcSheet, RowIndex + 6, Array("CCC", 0.5)
        PutRow DrcSheet, RowIndex + 7, Array("UNRATED", 0.15)
        PutRow DrcSheet, RowIndex + 8, Array("DEFAULTED", 1)
        FormatParamRows DrcSheet, RowIndex, RowIndex + 8, 2, 2, 2, "2"

'------------------------------------------------------------------------------
' DRC_SENIORITY
'------------------------------------------------------------------------------
    'Rank: 1 = most junior. A short may offset a long of the same or a more senior rank
        RowIndex = PutTable(DrcSheet, RowIndex + 10, "DRC_SENIORITY", "Seniority rank and LGD (MAR22.12, MAR22.19)", Array("Seniority", "Rank", "LGD", "Reference"))
        PutRow DrcSheet, RowIndex, Array("EQUITY", 1, 1, "MAR22.12")
        PutRow DrcSheet, RowIndex + 1, Array("NON_SENIOR", 2, 1, "MAR22.12")
        PutRow DrcSheet, RowIndex + 2, Array("SENIOR", 3, 0.75, "MAR22.12")
        PutRow DrcSheet, RowIndex + 3, Array("COVERED", 4, 0.25, "MAR22.12")
        FormatParamRows DrcSheet, RowIndex, RowIndex + 3, 4, 2, 3, "3"

'------------------------------------------------------------------------------
' DRC_PARAMS
'------------------------------------------------------------------------------
    'Write the maturity scaling parameters
        RowIndex = PutTable(DrcSheet, RowIndex + 5, "DRC_PARAMS", "Maturity scaling (years)", Array("Code", "Value", "Reference"))
        PutRow DrcSheet, RowIndex, Array("MIN_MATURITY", 0.25, "MAR22.18: floor of three months")
        PutRow DrcSheet, RowIndex + 1, Array("MAX_MATURITY", 1, "MAR22.15: one-year capital horizon")
        FormatParamRows DrcSheet, RowIndex, RowIndex + 1, 3, 2, 2, "", True

'------------------------------------------------------------------------------
' DRC_BUCKETS
'------------------------------------------------------------------------------
    'Write the buckets
        RowIndex = PutTable(DrcSheet, RowIndex + 3, "DRC_BUCKETS", "Buckets (MAR22.22)", Array("Bucket", "Description"))
        PutRow DrcSheet, RowIndex, Array("CORP", "Corporates")
        PutRow DrcSheet, RowIndex + 1, Array("SOV", "Sovereigns")
        PutRow DrcSheet, RowIndex + 2, Array("LGM", "Local governments and municipalities")
        FormatParamRows DrcSheet, RowIndex, RowIndex + 2, 2, 2, 2, ""

End Sub

Private Sub WriteParamRrao()

'
'------------------------------------------------------------------------------
'                          WRITE THE RRAO PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the P_RRAO sheet: the residual risk add-on risk weights.
'
' WHY THIS EXISTS
'   The engine reads these tagged tables from the sheet; the defaults are written here.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes one tagged table, RRAO_RW: the risk weights of the categories EXOTIC and OTHER
'   (MAR23.8).
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the cell writes raise normally.
'
' DEPENDENCIES
'   ParamSheet, SetWidths, PutTable, PutRow, FormatParamRows.
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
    Const PROC_NAME             As String = "M_Setup.WriteParamRrao"

    Dim RraoSheet               As Object       'P_RRAO sheet
    Dim RowIndex                As Long         'Row being written

'------------------------------------------------------------------------------
' WRITE THE TABLE
'------------------------------------------------------------------------------
    'Write the frame of the sheet and set the column widths
        Set RraoSheet = ParamSheet(SH_P_RRAO, "Residual risk add-on risk weights.", 6)
        SetWidths RraoSheet, Array(12, 10, 40)
    'Write the risk weights
        RowIndex = PutTable(RraoSheet, LAYOUT_HEADER_ROW, "RRAO_RW", "Risk weights (MAR23.8)", Array("Category", "RW", "Reference"))
        PutRow RraoSheet, RowIndex, Array("EXOTIC", 0.01, "MAR23.8(2)(a): exotic underlying")
        PutRow RraoSheet, RowIndex + 1, Array("OTHER", 0.001, "MAR23.8(2)(b): other residual risks")
        FormatParamRows RraoSheet, RowIndex, RowIndex + 1, 3, 2, 2, "2"

End Sub

'
'------------------------------------------------------------------------------
'
'          SYNTHETIC SAMPLE PORTFOLIO (AMOUNTS IN REPORTING CURRENCY)
'
'------------------------------------------------------------------------------
'

Private Sub WriteSampleData()

'
'------------------------------------------------------------------------------
'                          WRITE THE SAMPLE PORTFOLIO
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the synthetic sample portfolio to the five input sheets.
'
' BEHAVIOR
'   Calls the Write procedure of each input sheet in turn: delta, vega, curvature, DRC,
'   RRAO.
'
' DEPENDENCIES
'   WriteSampleDelta, WriteSampleVega, WriteSampleCurvature, WriteSampleDrc,
'   WriteSampleRrao.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE THE SAMPLES
'------------------------------------------------------------------------------
    'Write the samples of each input sheet in turn
        WriteSampleDelta
        WriteSampleVega
        WriteSampleCurvature
        WriteSampleDrc
        WriteSampleRrao

End Sub

Private Sub WriteSampleDelta()

'
'------------------------------------------------------------------------------
'                           WRITE THE DELTA SAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the 36 delta sensitivity records D001 to D036 to In_Delta.
'
' WHY THIS EXISTS
'   The regression test compares the engine results on the sample portfolio with its
'   expected values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes one record per call of PutSample: GIRR, CSR_NS, EQ, COM and FX sensitivities in
'   the reporting currency.
'
' ERROR POLICY
'   The writes raise normally.
'
' DEPENDENCIES
'   PutSample, GetSheet (CORE_Util).
'
' NOTES
'   The comment column of each record says what it represents, for example the record that
'   nets with D008.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WriteSampleDelta"

    Dim DeltaSheet              As Object       'In_Delta sheet

'------------------------------------------------------------------------------
' WRITE THE RECORDS
'------------------------------------------------------------------------------
    'Locate the sheet
        Set DeltaSheet = GetSheet(SH_IN_DELTA)
    'Write the delta records
        PutSample DeltaSheet, 1, Array("D001", "GIRR", "EUR", "ESTR", "YIELD", 1, 5000000, "Swap book")
        PutSample DeltaSheet, 2, Array("D002", "GIRR", "EUR", "ESTR", "YIELD", 5, -12000000, "Swap book")
        PutSample DeltaSheet, 3, Array("D003", "GIRR", "EUR", "ESTR", "YIELD", 10, 8000000, "Swap book")
        PutSample DeltaSheet, 4, Array("D004", "GIRR", "EUR", "EURIBOR6M", "YIELD", 5, 3000000, "Basis swaps")
        PutSample DeltaSheet, 5, Array("D005", "GIRR", "EUR", "EURIBOR6M", "YIELD", 10, -1500000, "Basis swaps")
        PutSample DeltaSheet, 6, Array("D006", "GIRR", "EUR", "HICP", "INFLATION", "", 400000, "Inflation swaps")
        PutSample DeltaSheet, 7, Array("D007", "GIRR", "USD", "SOFR", "YIELD", 2, -4000000, "UST hedge")
        PutSample DeltaSheet, 8, Array("D008", "GIRR", "USD", "SOFR", "YIELD", 10, 2000000, "UST hedge")
        PutSample DeltaSheet, 9, Array("D009", "GIRR", "USD", "SOFR", "YIELD", 10, 500000, "Netted with D008")
        PutSample DeltaSheet, 10, Array("D010", "GIRR", "USD", "USD-EUR", "XCCY", "", 1000000, "Cross-currency swaps")
        PutSample DeltaSheet, 11, Array("D011", "CSR_NS", 1, "ITALY", "BOND", 5, 3000000, "BTP")
        PutSample DeltaSheet, 12, Array("D012", "CSR_NS", 1, "SPAIN", "BOND", 5, 3000000, "Bono")
        PutSample DeltaSheet, 13, Array("D013", "CSR_NS", 1, "ITALY", "CDS", 5, 500000, "Sovereign CDS protection sold")
        PutSample DeltaSheet, 14, Array("D014", "CSR_NS", 2, "REGION_X", "BOND", 3, -2000000, "Regional bond short")
        PutSample DeltaSheet, 15, Array("D015", "CSR_NS", 2, "REGION_Y", "BOND", 5, -1800000, "Regional bond short")
        PutSample DeltaSheet, 16, Array("D016", "CSR_NS", 3, "BANK_A", "BOND", 3, 1500000, "Senior bond")
        PutSample DeltaSheet, 17, Array("D017", "CSR_NS", 3, "BANK_A", "CDS", 3, -1400000, "CDS hedge")
        PutSample DeltaSheet, 18, Array("D018", "CSR_NS", 11, "BANK_Z", "BOND", 1, 100000, "HY financial")
        PutSample DeltaSheet, 19, Array("D019", "CSR_NS", 16, "MISC_CO", "BOND", 5, 200000, "Other sector")
        PutSample DeltaSheet, 20, Array("D020", "CSR_NS", 16, "MISC_CO2", "BOND", 5, -150000, "Other sector")
        PutSample DeltaSheet, 21, Array("D021", "CSR_NS", 17, "ITRAXX_MAIN", "CDS", 5, -800000, "Index protection")
        PutSample DeltaSheet, 22, Array("D022", "EQ", 8, "BANK_A", "SPOT", "", 2000000, "Shares")
        PutSample DeltaSheet, 23, Array("D023", "EQ", 8, "BANK_B", "SPOT", "", -1200000, "Short shares")
        PutSample DeltaSheet, 24, Array("D024", "EQ", 8, "BANK_A", "REPO", "", 50000, "Securities lending")
        PutSample DeltaSheet, 25, Array("D025", "EQ", 5, "CONSUMER_X", "SPOT", "", 900000, "Shares")
        PutSample DeltaSheet, 26, Array("D026", "EQ", 12, "EUROSTOXX50", "SPOT", "", -1500000, "Index future short")
        PutSample DeltaSheet, 27, Array("D027", "EQ", 11, "OTHER_EQ", "SPOT", "", 100000, "Other sector")
        PutSample DeltaSheet, 28, Array("D028", "COM", 2, "BRENT", "ROTTERDAM", 0.5, 800000, "Futures")
        PutSample DeltaSheet, 29, Array("D029", "COM", 2, "BRENT", "ROTTERDAM", 1, -500000, "Futures")
        PutSample DeltaSheet, 30, Array("D030", "COM", 2, "WTI", "CUSHING", 1, 300000, "Futures")
        PutSample DeltaSheet, 31, Array("D031", "COM", 7, "GOLD", "LONDON", 0, 1000000, "Spot gold")
        PutSample DeltaSheet, 32, Array("D032", "COM", 3, "POWER_DE", "EEX", 0.25, 200000, "Power forward")
        PutSample DeltaSheet, 33, Array("D033", "FX", "USD", "USD", "", "", 3000000, "USD exposure")
        PutSample DeltaSheet, 34, Array("D034", "FX", "GBP", "GBP", "", "", -1500000, "GBP exposure")
        PutSample DeltaSheet, 35, Array("D035", "FX", "JPY", "JPY", "", "", 800000, "JPY exposure")
        PutSample DeltaSheet, 36, Array("D036", "FX", "CHF", "CHF", "", "", -400000, "CHF exposure")

End Sub

Private Sub WriteSampleVega()

'
'------------------------------------------------------------------------------
'                            WRITE THE VEGA SAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the ten vega sensitivity records V001 to V010 to In_Vega.
'
' WHY THIS EXISTS
'   The regression test compares the engine results on the sample portfolio with its
'   expected values.
'
' ERROR POLICY
'   The writes raise normally.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WriteSampleVega"

    Dim VegaSheet               As Object       'In_Vega sheet

'------------------------------------------------------------------------------
' WRITE THE RECORDS
'------------------------------------------------------------------------------
    'Locate the sheet
        Set VegaSheet = GetSheet(SH_IN_VEGA)
    'Write the vega records
        PutSample VegaSheet, 1, Array("V001", "GIRR", "EUR", "", "", 1, 5, 400000, "Swaptions")
        PutSample VegaSheet, 2, Array("V002", "GIRR", "EUR", "", "", 5, 10, -250000, "Swaptions")
        PutSample VegaSheet, 3, Array("V003", "GIRR", "USD", "", "", 3, 10, 200000, "Swaptions")
        PutSample VegaSheet, 4, Array("V004", "CSR_NS", 1, "ITALY", "BOND", 1, "", 60000, "Bond options")
        PutSample VegaSheet, 5, Array("V005", "EQ", 8, "BANK_A", "", 0.5, "", 300000, "Single-stock options")
        PutSample VegaSheet, 6, Array("V006", "EQ", 8, "BANK_A", "", 1, "", -100000, "Single-stock options")
        PutSample VegaSheet, 7, Array("V007", "EQ", 12, "EUROSTOXX50", "", 1, "", -500000, "Index options")
        PutSample VegaSheet, 8, Array("V008", "COM", 2, "BRENT", "ROTTERDAM", 1, "", 120000, "Oil options")
        PutSample VegaSheet, 9, Array("V009", "FX", "USD", "USD", "", 1, "", 150000, "FX options")
        PutSample VegaSheet, 10, Array("V010", "FX", "GBP", "GBP", "", 0.5, "", -80000, "FX options")

End Sub

Private Sub WriteSampleCurvature()

'
'------------------------------------------------------------------------------
'                         WRITE THE CURVATURE SAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the nine curvature records C001 to C009 to In_Curvature.
'
' WHY THIS EXISTS
'   The regression test compares the engine results on the sample portfolio with its
'   expected values.
'
' ERROR POLICY
'   The writes raise normally.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WriteSampleCurvature"

    Dim CurvatureSheet          As Object       'In_Curvature sheet

'------------------------------------------------------------------------------
' WRITE THE RECORDS
'------------------------------------------------------------------------------
    'Locate the sheet
        Set CurvatureSheet = GetSheet(SH_IN_CURV)
    'Write the curvature records
        PutSample CurvatureSheet, 1, Array("C001", "GIRR", "EUR", "", "", -50000, 30000, "", "Swaptions")
        PutSample CurvatureSheet, 2, Array("C002", "GIRR", "USD", "", "", 10000, -20000, "", "Swaptions")
        PutSample CurvatureSheet, 3, Array("C003", "CSR_NS", 1, "ITALY", "BOND", -15000, 5000, "", "Bond options")
        PutSample CurvatureSheet, 4, Array("C004", "EQ", 8, "BANK_A", "", -80000, 40000, "", "Single-stock options")
        PutSample CurvatureSheet, 5, Array("C005", "EQ", 8, "BANK_B", "", 20000, -60000, "", "Single-stock options")
        PutSample CurvatureSheet, 6, Array("C006", "EQ", 12, "EUROSTOXX50", "", 30000, -10000, "", "Index options")
        PutSample CurvatureSheet, 7, Array("C007", "COM", 2, "BRENT", "ROTTERDAM", -25000, 10000, "", "Oil options")
        PutSample CurvatureSheet, 8, Array("C008", "FX", "USD", "USD", "", -40000, -35000, True, "USD/EUR options")
        PutSample CurvatureSheet, 9, Array("C009", "FX", "GBP", "GBP", "", 15000, 12000, "", "GBP/USD options")

End Sub

Private Sub WriteSampleDrc()

'
'------------------------------------------------------------------------------
'                            WRITE THE DRC SAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the ten default risk positions J001 to J010 to In_DRC.
'
' WHY THIS EXISTS
'   The regression test compares the engine results on the sample portfolio with its
'   expected values.
'
' ERROR POLICY
'   The writes raise normally.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WriteSampleDrc"

    Dim DrcSheet                As Object       'In_DRC sheet

'------------------------------------------------------------------------------
' WRITE THE RECORDS
'------------------------------------------------------------------------------
    'Locate the sheet
        Set DrcSheet = GetSheet(SH_IN_DRC)
    'Write the DRC positions
        PutSample DrcSheet, 1, Array("J001", "ITALY", "SOV", "BBB", "SENIOR", 10000000, 9800000, 5, "", "BTP long")
        PutSample DrcSheet, 2, Array("J002", "GERMANY", "SOV", "AAA", "SENIOR", -5000000, -5100000, 10, "", "Bund short")
        PutSample DrcSheet, 3, Array("J003", "BANK_A", "CORP", "A", "SENIOR", 4000000, 4050000, 3, "", "Senior bond")
        PutSample DrcSheet, 4, Array("J004", "BANK_A", "CORP", "A", "EQUITY", 2000000, 2000000, 1, "", "Shares")
        PutSample DrcSheet, 5, Array("J005", "BANK_A", "CORP", "A", "SENIOR", -3000000, -3020000, 0.5, "", "Short senior bond")
        PutSample DrcSheet, 6, Array("J006", "BANK_B", "CORP", "BBB", "NON_SENIOR", -1500000, -1450000, 2, "", "Short subordinated")
        PutSample DrcSheet, 7, Array("J007", "CONSUMER_X", "CORP", "BB", "SENIOR", 1000000, 950000, 0.1, "", "Short-dated bond")
        PutSample DrcSheet, 8, Array("J008", "REGION_X", "LGM", "AA", "SENIOR", 2000000, 2010000, 4, "", "Regional bond")
        PutSample DrcSheet, 9, Array("J009", "MISC_CO", "CORP", "UNRATED", "SENIOR", 500000, 480000, 2, "", "Unrated bond")
        PutSample DrcSheet, 10, Array("J010", "BANK_Z", "CORP", "CCC", "COVERED", 800000, 700000, 1.5, "", "Covered bond")

End Sub

Private Sub WriteSampleRrao()

'
'------------------------------------------------------------------------------
'                            WRITE THE RRAO SAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the three residual risk records R001 to R003 to In_RRAO.
'
' WHY THIS EXISTS
'   The regression test compares the engine results on the sample portfolio with its
'   expected values.
'
' ERROR POLICY
'   The writes raise normally.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WriteSampleRrao"

    Dim RraoSheet               As Object       'In_RRAO sheet

'------------------------------------------------------------------------------
' WRITE THE RECORDS
'------------------------------------------------------------------------------
    'Locate the sheet
        Set RraoSheet = GetSheet(SH_IN_RRAO)
    'Write the RRAO records
        PutSample RraoSheet, 1, Array("R001", "EXOTIC", 5000000, "Weather derivative")
        PutSample RraoSheet, 2, Array("R002", "OTHER", 20000000, "Correlation basket option")
        PutSample RraoSheet, 3, Array("R003", "OTHER", 7500000, "Behavioral prepayment option")

End Sub

'
'------------------------------------------------------------------------------
'
'                       REGRESSION TEST EXPECTED RESULTS
'
'------------------------------------------------------------------------------
'

Private Sub WriteExpectedResults()

'
'------------------------------------------------------------------------------
'                       WRITE THE EXPECTED RESULTS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the Test_Expected sheet: the frame and the table of expected values of the sample
'   portfolio.
'
' WHY THIS EXISTS
'   The regression test compares the actual results of a run with these values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the sheet, writes the frame, the column widths and the [EXPECTED] table heading,
'   writes the rows with WriteExpectedRows, formats them and shows the Expected, Actual and
'   AbsDiff columns with two decimals.
'
' ERROR POLICY
'   Formatting is best effort inside the helpers; the writes raise normally.
'
' DEPENDENCIES
'   ApplyFrame, SetWidths, PutTable, WriteExpectedRows; FormatDataRows, TableCell
'   (CORE_Util).
'
' NOTES
'   The expected values are for the default parameters and settings.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Setup.WriteExpectedResults"

    Dim TestSheet               As Object       'Test_Expected sheet
    Dim RowIndex                As Long         'First data row
    Dim ItemCount               As Long         'Number of expected values written

'------------------------------------------------------------------------------
' WRITE THE FRAME
'------------------------------------------------------------------------------
    'Locate the sheet and clear it
        Set TestSheet = GetSheet(SH_TEST)
        TestSheet.Cells.Clear
    'Write the frame and set the column widths
        ApplyFrame TestSheet, "Test_Expected", "Expected results of the sample portfolio with default parameters and settings. " & _
            "Run Run FRTB-SA, then Run regression test. Agreement with the reference does not prove the reading of the regulation.", 12
        SetWidths TestSheet, Array(28, 18, 18, 12, 10, 24, 62, 54, 70)

'------------------------------------------------------------------------------
' WRITE THE TABLE
'------------------------------------------------------------------------------
    'Write the table heading and header row
        RowIndex = PutTable(TestSheet, LAYOUT_HEADER_ROW, "EXPECTED", "Expected values", _
            Array("Item", "Expected", "Actual", "AbsDiff", "Status", "Area", "Description", "Source", "Source status"))
    'Write the expected values and keep their count
        ItemCount = WriteExpectedRows(TestSheet, RowIndex)
    'Format the rows and show the numbers with two decimals
        FormatDataRows TestSheet, RowIndex, RowIndex + ItemCount - 1, 9, 0, "", False
        TestSheet.Range(TableCell(TestSheet, RowIndex, 2), _
            TableCell(TestSheet, RowIndex + ItemCount - 1, 4)).NumberFormat = "#,##0.00"

End Sub

Private Function WriteExpectedRows(ByVal ws As Object, ByVal r As Long) As Long

'
'------------------------------------------------------------------------------
'                          WRITE THE EXPECTED VALUES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one row per expected value of the sample portfolio into the table of
'   Test_Expected.
'
' WHY THIS EXISTS
'   The regression test needs the values of an independent reference implementation for each
'   result item.
'
' INPUTS
'   ws
'     Test_Expected sheet
'
'   r
'     First data row
'
' RETURNS
'   The number of rows written.
'
' BEHAVIOR
'   Writes each row with PutRow: item name, expected value, three empty cells (actual value,
'   difference and status), area, description, source and source status.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   PutRow.
'
' NOTES
'   The rows between the generated markers are rewritten by tools/update_expected.py from
'   the reference implementation under tests/; do not edit them by hand.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE THE ROWS
'------------------------------------------------------------------------------
    'Write one row per expected value and return their count; the lines between the markers are
    'generated by tools/update_expected.py
    ' BEGIN GENERATED EXPECTED VALUES
    PutRow ws, r + 0, Array("SBM_LOW", 3163703.416013, Empty, Empty, Empty, "SBM", "Sum over risk classes of delta, vega and curvature, low scenario", "MAR21.7; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 1, Array("SBM_MEDIUM", 3039394.382502, Empty, Empty, Empty, "SBM", "Sum over risk classes of delta, vega and curvature, medium scenario", "MAR21.7; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 2, Array("SBM_HIGH", 2932555.615174, Empty, Empty, Empty, "SBM", "Sum over risk classes of delta, vega and curvature, high scenario", "MAR21.7; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 3, Array("SBM", 3163703.416013, Empty, Empty, Empty, "SBM", "The largest of the three scenario totals", "MAR21.7; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 4, Array("DRC_NONSEC", 636589.396433, Empty, Empty, Empty, "DRC", "Default risk capital of non-securitizations, sum of the buckets", "MAR22.22-MAR22.26; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 5, Array("RRAO", 77500, Empty, Empty, Empty, "RRAO", "Gross notional times 1.0% (exotic) and 0.1% (other)", "MAR23.8; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 6, Array("TOTAL", 3877792.812446, Empty, Empty, Empty, "Total", "SBM plus DRC plus RRAO", "docs/methodology.md; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 7, Array("GIRR|DELTA|LOW", 60423.50359, Empty, Empty, Empty, "SBM | GIRR | Delta", "Delta capital of GIRR, low scenario", "MAR21.42-MAR21.50; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 8, Array("GIRR|DELTA|MEDIUM", 56980.229937, Empty, Empty, Empty, "SBM | GIRR | Delta", "Delta capital of GIRR, medium scenario", "MAR21.42-MAR21.50; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 9, Array("GIRR|DELTA|HIGH", 53315.039354, Empty, Empty, Empty, "SBM | GIRR | Delta", "Delta capital of GIRR, high scenario", "MAR21.42-MAR21.50; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 10, Array("GIRR|VEGA|LOW", 323277.327073, Empty, Empty, Empty, "SBM | GIRR | Vega", "Vega capital of GIRR, low scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 11, Array("GIRR|VEGA|MEDIUM", 319771.9736, Empty, Empty, Empty, "SBM | GIRR | Vega", "Vega capital of GIRR, medium scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 12, Array("GIRR|VEGA|HIGH", 316227.766017, Empty, Empty, Empty, "SBM | GIRR | Vega", "Vega capital of GIRR, high scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 13, Array("GIRR|CURVATURE|LOW", 33354.16016, Empty, Empty, Empty, "SBM | GIRR | Curvature", "Curvature capital of GIRR, low scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 14, Array("GIRR|CURVATURE|MEDIUM", 33911.649916, Empty, Empty, Empty, "SBM | GIRR | Curvature", "Curvature capital of GIRR, medium scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 15, Array("GIRR|CURVATURE|HIGH", 34460.12188, Empty, Empty, Empty, "SBM | GIRR | Curvature", "Curvature capital of GIRR, high scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 16, Array("CSR_NS|DELTA|LOW", 62844.491179, Empty, Empty, Empty, "SBM | CSR_NS | Delta", "Delta capital of CSR non-securitization, low scenario", "MAR21.51-MAR21.57; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 17, Array("CSR_NS|DELTA|MEDIUM", 48260.890512, Empty, Empty, Empty, "SBM | CSR_NS | Delta", "Delta capital of CSR non-securitization, medium scenario", "MAR21.51-MAR21.57; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 18, Array("CSR_NS|DELTA|HIGH", 60572.001077, Empty, Empty, Empty, "SBM | CSR_NS | Delta", "Delta capital of CSR non-securitization, high scenario", "MAR21.51-MAR21.57; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 19, Array("CSR_NS|VEGA|LOW", 60000, Empty, Empty, Empty, "SBM | CSR_NS | Vega", "Vega capital of CSR non-securitization, low scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 20, Array("CSR_NS|VEGA|MEDIUM", 60000, Empty, Empty, Empty, "SBM | CSR_NS | Vega", "Vega capital of CSR non-securitization, medium scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 21, Array("CSR_NS|VEGA|HIGH", 60000, Empty, Empty, Empty, "SBM | CSR_NS | Vega", "Vega capital of CSR non-securitization, high scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 22, Array("CSR_NS|CURVATURE|LOW", 5000, Empty, Empty, Empty, "SBM | CSR_NS | Curvature", "Curvature capital of CSR non-securitization, low scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 23, Array("CSR_NS|CURVATURE|MEDIUM", 5000, Empty, Empty, Empty, "SBM | CSR_NS | Curvature", "Curvature capital of CSR non-securitization, medium scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 24, Array("CSR_NS|CURVATURE|HIGH", 5000, Empty, Empty, Empty, "SBM | CSR_NS | Curvature", "Curvature capital of CSR non-securitization, high scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 25, Array("EQ|DELTA|LOW", 1156963.1261, Empty, Empty, Empty, "SBM | EQ | Delta", "Delta capital of equity, low scenario", "MAR21.72-MAR21.80; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 26, Array("EQ|DELTA|MEDIUM", 1109540.409267, Empty, Empty, Empty, "SBM | EQ | Delta", "Delta capital of equity, medium scenario", "MAR21.72-MAR21.80; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 27, Array("EQ|DELTA|HIGH", 1059848.3154, Empty, Empty, Empty, "SBM | EQ | Delta", "Delta capital of equity, high scenario", "MAR21.72-MAR21.80; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 28, Array("EQ|VEGA|LOW", 367878.895929, Empty, Empty, Empty, "SBM | EQ | Vega", "Vega capital of equity, low scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 29, Array("EQ|VEGA|MEDIUM", 348369.331364, Empty, Empty, Empty, "SBM | EQ | Vega", "Vega capital of equity, medium scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 30, Array("EQ|VEGA|HIGH", 327700.320415, Empty, Empty, Empty, "SBM | EQ | Vega", "Vega capital of equity, high scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 31, Array("EQ|CURVATURE|LOW", 45746.584572, Empty, Empty, Empty, "SBM | EQ | Curvature", "Curvature capital of equity, low scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 32, Array("EQ|CURVATURE|MEDIUM", 44237.992721, Empty, Empty, Empty, "SBM | EQ | Curvature", "Curvature capital of equity, medium scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 33, Array("EQ|CURVATURE|HIGH", 42676.105727, Empty, Empty, Empty, "SBM | EQ | Curvature", "Curvature capital of equity, high scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 34, Array("COM|DELTA|LOW", 352784.668318, Empty, Empty, Empty, "SBM | COM | Delta", "Delta capital of commodity, low scenario", "MAR21.82-MAR21.85; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 35, Array("COM|DELTA|MEDIUM", 366440.324064, Empty, Empty, Empty, "SBM | COM | Delta", "Delta capital of commodity, medium scenario", "MAR21.82-MAR21.85; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 36, Array("COM|DELTA|HIGH", 379605.057922, Empty, Empty, Empty, "SBM | COM | Delta", "Delta capital of commodity, high scenario", "MAR21.82-MAR21.85; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 37, Array("COM|VEGA|LOW", 120000, Empty, Empty, Empty, "SBM | COM | Vega", "Vega capital of commodity, low scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 38, Array("COM|VEGA|MEDIUM", 120000, Empty, Empty, Empty, "SBM | COM | Vega", "Vega capital of commodity, medium scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 39, Array("COM|VEGA|HIGH", 120000, Empty, Empty, Empty, "SBM | COM | Vega", "Vega capital of commodity, high scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 40, Array("COM|CURVATURE|LOW", 10000, Empty, Empty, Empty, "SBM | COM | Curvature", "Curvature capital of commodity, low scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 41, Array("COM|CURVATURE|MEDIUM", 10000, Empty, Empty, Empty, "SBM | COM | Curvature", "Curvature capital of commodity, medium scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 42, Array("COM|CURVATURE|HIGH", 10000, Empty, Empty, Empty, "SBM | COM | Curvature", "Curvature capital of commodity, high scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 43, Array("FX|DELTA|LOW", 430894.418622, Empty, Empty, Empty, "SBM | FX | Delta", "Delta capital of FX, low scenario", "MAR21.87-MAR21.89; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 44, Array("FX|DELTA|MEDIUM", 396465.635333, Empty, Empty, Empty, "SBM | FX | Delta", "Delta capital of FX, medium scenario", "MAR21.87-MAR21.89; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 45, Array("FX|DELTA|HIGH", 358747.822293, Empty, Empty, Empty, "SBM | FX | Delta", "Delta capital of FX, high scenario", "MAR21.87-MAR21.89; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 46, Array("FX|VEGA|LOW", 134536.240471, Empty, Empty, Empty, "SBM | FX | Vega", "Vega capital of FX, low scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 47, Array("FX|VEGA|MEDIUM", 120415.945788, Empty, Empty, Empty, "SBM | FX | Vega", "Vega capital of FX, medium scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 48, Array("FX|VEGA|HIGH", 104403.065089, Empty, Empty, Empty, "SBM | FX | Vega", "Vega capital of FX, high scenario", "MAR21.25, MAR21.92, MAR21.94; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 49, Array("FX|CURVATURE|LOW", 0, Empty, Empty, Empty, "SBM | FX | Curvature", "Curvature capital of FX, low scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 50, Array("FX|CURVATURE|MEDIUM", 0, Empty, Empty, Empty, "SBM | FX | Curvature", "Curvature capital of FX, medium scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    PutRow ws, r + 51, Array("FX|CURVATURE|HIGH", 0, Empty, Empty, Empty, "SBM | FX | Curvature", "Curvature capital of FX, high scenario", "MAR21.5, MAR21.98-MAR21.101; independent reference", "Parameters checked against the BIS text; reading not independently reviewed (#29)")
    WriteExpectedRows = 52
    ' END GENERATED EXPECTED VALUES

End Function
