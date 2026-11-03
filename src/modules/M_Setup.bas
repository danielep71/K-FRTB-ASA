Attribute VB_Name = "M_Setup"
'==============================================================================
' MODULE: M_Setup
'------------------------------------------------------------------------------
' PURPOSE
'   Build the workbook from an empty one: every sheet with the K-FRTB-SA
'   frame (brand bar, title bar, subtitle, tab color), the default
'   regulatory parameters (Basel MAR21-MAR23), a synthetic sample portfolio,
'   the buttons and the expected results of the regression test. The
'   macro-free template src/workbook/FRTB_Template.xlsx is this build saved
'   without VBA; FinishWorkbook adds buttons and reapplies view settings.
'
' PUBLIC SURFACE
'   BuildWorkbook: rebuild every sheet (overwrites the inputs), then
'   FinishWorkbook.
'   FinishWorkbook: add the sheet buttons and window settings to a workbook
'   opened from the template.
'   ResetParameters: rewrite the P_* sheets with the defaults.
'   LoadSampleData: rewrite the In_* sheets with the sample portfolio.
'   ManagedSheets: the sheet names in workbook order.
'
' DEPENDENCIES
'   CORE_Config for names, layout and palette; CORE_Util for sheet access
'   and table formatting.
'
' STATE OWNERSHIP
'   Turns alerts and screen updating off while building and back on at the
'   end. Activates each sheet once to set gridlines and frozen panes.
'
' ERROR POLICY
'   Formatting is best effort: a host that rejects a format property keeps
'   the content and skips the format. Content writes raise normally.
'
' REFERENCE
'   Parameter values: Basel Framework MAR21-MAR23 (in force 1 January
'   2023), checked against the BIS text on 2026-10-17. Verify them against
'   the regulation that applies to you before relying on any result.
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
Private Const INPUT_ROWS As Long = 200                       'Pre-formatted input rows per input sheet

'------------------------------------------------------------------------------
'                              PUBLIC MACROS
'------------------------------------------------------------------------------
Public Sub BuildWorkbook()
    Dim names As Variant, i As Long
    SetAlerts False
    names = ManagedSheets()
    For i = LBound(names) To UBound(names)
        EnsureSheet CStr(names(i))
    Next i
    RemoveDefaultSheets

    BuildReadme
    BuildControl
    BuildOutputSheets
    BuildInputSheets
    WriteParameters
    WriteSampleData
    WriteExpectedResults
    SetAlerts True
    FinishWorkbook
End Sub


' The XLSX retains presentation settings. Reapply them when adding the
' macro buttons, which are deliberately excluded from the template.
Public Sub FinishWorkbook()
    Dim names As Variant, i As Long
    SetAlerts False
    names = ManagedSheets()
    For i = LBound(names) To UBound(names)
        If SheetExists(CStr(names(i))) Then ApplyTabColor GetSheet(CStr(names(i)))
    Next i
    AddWorkbookButtons
    ApplyWindowSettings
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_README).Activate
    On Error GoTo 0
    SetAlerts True
End Sub


Public Sub ResetParameters()
    SetAlerts False
    WriteParameters
    SetAlerts True
End Sub


Public Sub LoadSampleData()
    SetAlerts False
    BuildInputSheets
    WriteSampleData
    SetAlerts True
End Sub


Public Function ManagedSheets() As Variant
    ManagedSheets = Array(SH_README, SH_CONTROL, SH_RESULTS, _
                          SH_IN_DELTA, SH_IN_VEGA, SH_IN_CURV, SH_IN_DRC, SH_IN_RRAO, _
                          SH_OUT_SUMMARY, SH_OUT_CLASS, SH_OUT_BUCKETS, SH_OUT_DRC, SH_LOG, _
                          SH_P_GENERAL, SH_P_GIRR, SH_P_CSR, SH_P_EQ, SH_P_COM, SH_P_FX, _
                          SH_P_DRC, SH_P_RRAO, SH_TEST)
End Function

'------------------------------------------------------------------------------
'                              SHEET HELPERS
'------------------------------------------------------------------------------
Private Sub SetAlerts(ByVal enabled As Boolean)
    On Error Resume Next
    Application.DisplayAlerts = enabled
    Application.ScreenUpdating = enabled
    On Error GoTo 0
End Sub


Private Sub EnsureSheet(ByVal sheetName As String)
    Dim ws As Object, wb As Object
    Set wb = ThisWorkbook
    If SheetExists(sheetName) Then
        wb.Worksheets(sheetName).Cells.Clear
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
End Sub


Private Function IsManaged(ByVal sheetName As String) As Boolean
    Dim names As Variant, i As Long
    names = ManagedSheets()
    IsManaged = False
    For i = LBound(names) To UBound(names)
        If UCase$(CStr(names(i))) = UCase$(sheetName) Then
            IsManaged = True
            Exit Function
        End If
    Next i
End Function


' Deletes blank sheets that are not part of the layout (the default "Sheet1"
' of a new workbook). Sheets with content are never deleted.
Private Sub RemoveDefaultSheets()
    Dim i As Long, ws As Object
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Not IsManaged(ws.Name) Then
            If IsBlank(ws.Cells(1, 1).Value) And IsBlank(ws.Cells(2, 1).Value) And IsBlank(ws.Cells(1, 2).Value) Then
                On Error Resume Next
                ws.Delete
                On Error GoTo 0
            End If
        End If
    Next i
End Sub


' Brand bar, title bar, subtitle and spacer rows shared by every sheet
' (the workbook frame); lastCol is the last column of the bars.
Private Sub ApplyFrame(ByVal ws As Object, ByVal title As String, ByVal subtitle As String, _
                       ByVal lastCol As Long)
    ws.Cells(1, 2).Value = APP_NAME
    ws.Cells(2, 2).Value = SheetTitle(title)
    ws.Cells(LAYOUT_SUBTITLE_ROW, LAYOUT_FIRST_COL).Value = subtitle
    On Error Resume Next
    ws.Cells.Font.Name = FONT_NAME
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = -4108                       'xlCenter
    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 4
    With ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol))
        .Interior.Color = COLOR_BRAND_BAR
        .Font.Color = COLOR_BRAND_TEXT
    End With
    ws.Cells(1, 2).Font.Size = 14
    With ws.Range(ws.Cells(2, 2), ws.Cells(2, lastCol))
        .Interior.Color = COLOR_TITLE_BAR
        .Font.Color = COLOR_WHITE
    End With
    ws.Cells(2, 2).Font.Size = 18
    ws.Cells(2, 2).Font.Bold = True
    With ws.Cells(LAYOUT_SUBTITLE_ROW, LAYOUT_FIRST_COL).Font
        .Italic = True
        .Color = COLOR_SUBTITLE
    End With
    ws.Rows(1).RowHeight = 24
    ws.Rows(2).RowHeight = 26.1
    ws.Rows(3).RowHeight = 30
    ws.Rows(4).RowHeight = 9.95
    With ws.PageSetup
        .PrintArea = ""
        .Orientation = 2                                    'xlLandscape
        .PaperSize = 9                                      'xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = 18
        .RightMargin = 18
        .TopMargin = 36
        .BottomMargin = 36
        .CenterHeader = ""
        .CenterFooter = "&""Aptos,Regular""&9Page &P of &N"
        If ws.Name = SH_README Then
            .PrintTitleRows = "$1:$3"
        ElseIf ws.Name = SH_CONTROL Or ws.Name = SH_TEST Or Left$(ws.Name, 2) = "P_" Then
            .PrintTitleRows = "$1:$6"
        Else
            .PrintTitleRows = "$1:$5"
        End If
    End With
    On Error GoTo 0
    ExtendPrintArea ws, 4, lastCol
    ApplyTabColor ws
End Sub


Private Function SheetTitle(ByVal sheetName As String) As String
    Select Case sheetName
        Case SH_IN_DELTA: SheetTitle = "Delta sensitivities"
        Case SH_IN_VEGA: SheetTitle = "Vega sensitivities"
        Case SH_IN_CURV: SheetTitle = "Curvature sensitivities"
        Case SH_IN_DRC: SheetTitle = "Default risk positions"
        Case SH_IN_RRAO: SheetTitle = "Residual risk positions"
        Case SH_OUT_SUMMARY: SheetTitle = "Calculation summary"
        Case SH_OUT_CLASS: SheetTitle = "Capital by risk class"
        Case SH_OUT_BUCKETS: SheetTitle = "Bucket aggregation"
        Case SH_OUT_DRC: SheetTitle = "Default risk calculation"
        Case SH_TEST: SheetTitle = "Regression tests"
        Case SH_P_GENERAL: SheetTitle = "General parameters"
        Case SH_P_GIRR: SheetTitle = "Interest rate parameters"
        Case SH_P_CSR: SheetTitle = "Credit spread parameters"
        Case SH_P_EQ: SheetTitle = "Equity parameters"
        Case SH_P_COM: SheetTitle = "Commodity parameters"
        Case SH_P_FX: SheetTitle = "Foreign exchange parameters"
        Case SH_P_DRC: SheetTitle = "Default risk parameters"
        Case SH_P_RRAO: SheetTitle = "Residual risk parameters"
        Case Else: SheetTitle = sheetName
    End Select
End Function


' Tab color by role, from the sheet name: navy README, green tests, blue
' outputs, gold inputs and parameters. Set again by FinishWorkbook because
' not every host keeps it in the saved template.
Private Sub ApplyTabColor(ByVal ws As Object)
    Dim color As Long
    Select Case True
        Case ws.Name = SH_README: color = TAB_REFERENCE
        Case ws.Name = SH_TEST: color = TAB_TEST
        Case ws.Name = SH_RESULTS, ws.Name = SH_LOG, Left$(ws.Name, 4) = "Out_": color = TAB_OUTPUT
        Case Else: color = TAB_INPUT
    End Select
    On Error Resume Next
    ws.Tab.Color = color
    On Error GoTo 0
End Sub


' Writes values from column C. Text is stored as text, so that a reference
' such as "MAR21.54" is never converted to a date or a number.
Private Sub PutRow(ByVal ws As Object, ByVal r As Long, ByVal values As Variant)
    Dim i As Long, k As Long
    k = 0
    For i = LBound(values) To UBound(values)
        k = k + 1
        If VarType(values(i)) = vbString Then
            If values(i) <> "" Then TableCell(ws, r, k).NumberFormat = "@"
        End If
        TableCell(ws, r, k).Value = values(i)
    Next i
End Sub


' Input record n (1-based) of an input sheet.
Private Sub PutSample(ByVal ws As Object, ByVal n As Long, ByVal values As Variant)
    PutRow ws, LAYOUT_HEADER_ROW + n, values
End Sub


Private Sub SetWidths(ByVal ws As Object, ByVal widths As Variant)
    Dim i As Long, k As Long
    On Error Resume Next
    k = 0
    For i = LBound(widths) To UBound(widths)
        k = k + 1
        ws.Columns(LAYOUT_FIRST_COL + k - 1).ColumnWidth = widths(i)
    Next i
    On Error GoTo 0
End Sub


' Section heading: navy bold text over a medium navy rule.
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


' Tagged parameter table: a section heading "[TAG] title" read by the engine,
' then a navy header row. Returns the first data row.
Private Function PutTable(ByVal ws As Object, ByVal r As Long, ByVal tag As String, ByVal title As String, _
                          ByVal headers As Variant) As Long
    Dim nCols As Long
    nCols = UBound(headers) - LBound(headers) + 1
    WriteSectionHeading ws, r, "[" & tag & "]  " & title, nCols
    WriteTableHeader ws, r + 1, headers, COLOR_INPUT_HEADER
    PutTable = r + 2
End Function


' Formats the data rows of a parameter table: editable columns from
' firstValueCol to lastValueCol in the input fill, percentages where the
' column (or, in Key | Value tables, the key) denotes a rate.
Private Sub FormatParamRows(ByVal ws As Object, ByVal r1 As Long, ByVal r2 As Long, ByVal nCols As Long, _
                            ByVal firstValueCol As Long, ByVal lastValueCol As Long, ByVal pctCols As String)
    Dim r As Long, k As Long, key As String
    FormatDataRows ws, r1, r2, nCols, 0, "", False
    On Error Resume Next
    ws.Range(TableCell(ws, r1, firstValueCol), TableCell(ws, r2, lastValueCol)).Interior.Color = COLOR_INPUT_CELL
    ws.Range(TableCell(ws, r1, 1), TableCell(ws, r2, 1)).Font.Bold = True
    For r = r1 To r2
        For k = firstValueCol To lastValueCol
            If InStr("," & pctCols & ",", "," & CStr(k) & ",") > 0 Then
                If IsNumber(TableCell(ws, r, k).Value) Then TableCell(ws, r, k).NumberFormat = "0.0##%"
            ElseIf pctCols = "KEY" Then
                key = UCase$(CStr(TableCell(ws, r, 1).Value))
                If (Left$(key, 3) = "RHO" Or Left$(key, 2) = "RW" Or Left$(key, 5) = "GAMMA" Or _
                    key = "THETA" Or key = "ALPHA") And IsNumber(TableCell(ws, r, k).Value) Then
                    TableCell(ws, r, k).NumberFormat = "0.0##%"
                End If
            End If
        Next k
    Next r
    On Error GoTo 0
End Sub


Private Sub WrapColumn(ByVal ws As Object, ByVal r1 As Long, ByVal r2 As Long, ByVal k As Long)
    On Error Resume Next
    ws.Range(TableCell(ws, r1, k), TableCell(ws, r2, k)).WrapText = True
    ws.Range(ws.Rows(r1), ws.Rows(r2)).AutoFit
    On Error GoTo 0
End Sub


' Button in the subtitle row, slot-th (0-based) from the left edge of anchorCol.
Private Sub AddButton(ByVal ws As Object, ByVal anchorCol As Long, ByVal slot As Long, _
                      ByVal caption As String, ByVal macro As String)
    Dim btn As Object
    Set btn = Nothing
    On Error Resume Next
    With ws.Cells(LAYOUT_SUBTITLE_ROW, anchorCol)
        Set btn = ws.Buttons.Add(.Left + 2 + slot * 130, .Top + 3, 124, 24)
    End With
    If Not btn Is Nothing Then
        btn.Caption = caption
        btn.OnAction = macro
    End If
    On Error GoTo 0
End Sub


Private Sub RemoveButtons(ByVal ws As Object)
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
End Sub


' Hides gridlines and freezes the panes below the table header on every
' sheet; needs each sheet to be active once. The wide input, output and Log
' sheets also keep their key columns in view; the narrow tables of Results,
' Control, Test_Expected and the parameter sheets freeze rows only.
Private Sub ApplyWindowSettings()
    Dim names As Variant, i As Long, ws As Object
    names = ManagedSheets()
    On Error Resume Next
    For i = LBound(names) To UBound(names)
        Set ws = ThisWorkbook.Worksheets(CStr(names(i)))
        ws.Activate
        Application.ActiveWindow.DisplayGridlines = False
        Application.ActiveWindow.FreezePanes = False
        If Left$(ws.Name, 3) = "In_" Or Left$(ws.Name, 4) = "Out_" Or ws.Name = SH_LOG Then
            ws.Cells(LAYOUT_FIRST_DATA_ROW, LAYOUT_FIRST_COL + 1).Select
            Application.ActiveWindow.FreezePanes = True
        ElseIf ws.Name = SH_RESULTS Then
            ws.Cells(LAYOUT_FIRST_DATA_ROW, 1).Select
            Application.ActiveWindow.FreezePanes = True
        ElseIf ws.Name = SH_CONTROL Or ws.Name = SH_TEST Or Left$(ws.Name, 2) = "P_" Then
            ws.Cells(LAYOUT_HEADER_ROW + 2, 1).Select
            Application.ActiveWindow.FreezePanes = True
        Else
            ws.Cells(4, 1).Select
            Application.ActiveWindow.FreezePanes = True
        End If
        ws.Cells(1, 1).Select
    Next i
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
'                         README, CONTROL, OUTPUTS
'------------------------------------------------------------------------------
Private Sub BuildReadme()
    Dim ws As Object, r As Long
    Set ws = GetSheet(SH_README)
    ApplyFrame ws, "ReadMe", APP_TITLE & " (Basel MAR21-MAR23) in Excel VBA. Prototype " & APP_VERSION & _
        ": results are not validated for regulatory use.", 4
    SetWidths ws, Array(55, 150)

    r = 5
    WriteSectionHeading ws, r, "How to use", 2
    r = WriteGuideRows(ws, r + 1, Array( _
        "1. Enable macros", "Open the .xlsm and click Enable Content. If Excel shows that macros are blocked, close the file, open its Properties, tick Unblock and reopen it.", _
        "2. Settings", "Control: reporting currency and the national discretions (sqrt(2) for specified GIRR currencies and FX pairs, FX curvature divisor). Cream cells are inputs.", _
        "3. Sensitivities", "In_Delta, In_Vega, In_Curvature: one row per sensitivity, in the reporting currency. Rows with the same risk factor are netted.", _
        "4. Positions", "In_DRC: one row per position for the default risk charge. In_RRAO: gross notional of instruments with residual risks.", _
        "5. Run", "Click Run FRTB-SA (macro RunFRTB). Results, Out_*, Out_Summary and Log are rewritten on every run.", _
        "6. Review", "Log lists every rejected row with the reason; a rejected row is never counted as zero. Results shows the capital by risk class, component and correlation scenario.", _
        "7. Test", "Test_Expected compares the sample portfolio with values from an independent reference implementation (Run regression test, development build)."))

    r = r + 1
    WriteSectionHeading ws, r, "Input conventions", 2
    r = WriteGuideRows(ws, r + 1, Array( _
        "Risk classes", "GIRR, CSR_NS (credit spread, non-securitization), EQ, COM, FX. CSR securitization and the correlation trading portfolio are not implemented and are rejected.", _
        "Buckets", "GIRR and FX: currency code. CSR_NS 1-18, EQ 1-13, COM 1-11 as in MAR21.51, MAR21.72 and MAR21.82 (see the P_* sheets).", _
        "Qualifier", "GIRR: curve name. CSR_NS and EQ: issuer or index. COM: commodity. FX: the currency.", _
        "Sub-type", "GIRR: YIELD, INFLATION or XCCY. CSR_NS delta: BOND or CDS. EQ delta: SPOT or REPO. COM delta: delivery location.", _
        "Tenors", "GIRR 0.25-30y, CSR 0.5-10y, COM 0-30y; vega option and GIRR underlying maturities 0.5, 1, 3, 5, 10y. GIRR INFLATION and XCCY vega: option maturity only.", _
        "Amounts", "Sensitivities as defined in MAR21.19-26 (for example GIRR: PV change for 1bp / 0.0001). Vega: vega x implied volatility. Curvature: CVR up and CVR down per risk factor; RefReportingCcy TRUE marks FX options that reference the reporting currency (MAR21.98).", _
        "DRC", "Signed notional and market value (negative for shorts), maturity in years (cash equity: 0.25 or at least 1, MAR22.16), bucket CORP, SOV or LGM, credit quality AAA ... DEFAULTED, seniority EQUITY, NON_SENIOR, SENIOR or COVERED."))

    r = r + 1
    WriteSectionHeading ws, r, "Sheet map", 2
    r = WriteGuideRows(ws, r + 1, Array( _
        "Gold tabs", "Inputs and parameters: Control, In_*, P_*. Parameter codes in column C are read by the engine; only cream cells are meant to be changed.", _
        "Blue tabs", "Outputs: Results, Out_Summary, Out_RiskClass, Out_Buckets (Kb, Sb and notes ALT_SB, UP/DOWN, OTHER), Out_DRC, Log.", _
        "Green tab", "Test_Expected: in-workbook regression test of the sample portfolio."))
End Sub


' Two-column guide rows (label | text) with dashed inner borders and a thin
' outer border; returns the row after the block.
Private Function WriteGuideRows(ByVal ws As Object, ByVal r As Long, ByVal pairs As Variant) As Long
    Dim i As Long, first As Long
    first = r
    For i = LBound(pairs) To UBound(pairs) Step 2
        TableCell(ws, r, 1).Value = pairs(i)
        TableCell(ws, r, 2).Value = pairs(i + 1)
        r = r + 1
    Next i
    On Error Resume Next
    With ws.Range(TableCell(ws, first, 1), TableCell(ws, r - 1, 2))
        .WrapText = True
    End With
    FormatBorders ws.Range(TableCell(ws, first, 1), TableCell(ws, r - 1, 2)), COLOR_GRID, -4115
    With ws.Range(TableCell(ws, first, 1), TableCell(ws, r - 1, 2))
        .Borders(7).LineStyle = 1                            'xlEdgeLeft
        .Borders(8).LineStyle = 1                            'xlEdgeTop
        .Borders(9).LineStyle = 1                            'xlEdgeBottom
        .Borders(10).LineStyle = 1                           'xlEdgeRight
    End With
    ws.Range(TableCell(ws, first, 1), TableCell(ws, r - 1, 1)).Font.Bold = True
    ws.Range(TableCell(ws, first, 2), TableCell(ws, r - 1, 2)).Font.Size = 9
    ws.Range(ws.Rows(first), ws.Rows(r - 1)).RowHeight = 33.95
    On Error GoTo 0
    ExtendPrintArea ws, r - 1, LAYOUT_FIRST_COL + 1
    WriteGuideRows = r
End Function


Private Sub BuildControl()
    Dim ws As Object, r As Long
    Set ws = GetSheet(SH_CONTROL)
    ApplyFrame ws, "Control", "Run settings and actions.", 6
    SetWidths ws, Array(34, 14, 125, 16)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "SETTINGS", "Run settings", Array("Code", "Value", "Description", "Reference"))
    PutRow ws, r, Array("REPORTING_CCY", "EUR", "Reporting currency; FX sensitivities are against this currency", "MAR21.14")
    PutRow ws, r + 1, Array("GIRR_SPECIFIED_CCY_DISCRETION", False, "TRUE: GIRR delta risk weights (tenor, inflation and cross-currency basis) divided by sqrt(2) for the specified currencies and the reporting currency", "MAR21.44")
    PutRow ws, r + 2, Array("FX_SPECIFIED_PAIR_DISCRETION", False, "TRUE: FX risk weight divided by sqrt(2) for specified pairs and their first-order crosses", "MAR21.88")
    PutRow ws, r + 3, Array("FX_CURVATURE_DIVISOR", False, "TRUE: FX curvature CVRs divided by 1.5 for options that do not reference the reporting currency (In_Curvature RefReportingCcy not TRUE)", "MAR21.98")
    PutRow ws, r + 4, Array("FX_CURVATURE_DIVISOR_ALL", False, "TRUE: all FX curvature CVRs divided by 1.5; needs supervisory approval and curvature sensitivities for all currencies, including the reporting currency", "MAR21.98")
    FormatParamRows ws, r, r + 4, 4, 2, 2, ""
    On Error Resume Next
    ws.Range(TableCell(ws, r, 2), TableCell(ws, r + 4, 2)).HorizontalAlignment = -4108
    On Error GoTo 0
End Sub


' Removes and recreates the buttons of Control and Results.
Private Sub AddWorkbookButtons()
    Dim ws As Object
    Set ws = GetSheet(SH_CONTROL)
    RemoveButtons ws
    AddButton ws, LAYOUT_FIRST_COL + 2, 0, "Run FRTB-SA", "RunFRTB"
    AddButton ws, LAYOUT_FIRST_COL + 2, 1, "Clear outputs", "ClearOutputs"
    AddButton ws, LAYOUT_FIRST_COL + 2, 2, "Run regression test", "Test_RunRegression"
    Set ws = GetSheet(SH_RESULTS)
    RemoveButtons ws
    AddButton ws, LAYOUT_FIRST_COL + 5, 0, "Run FRTB-SA", "RunFRTB"
End Sub


Private Sub BuildOutputSheets()
    Dim ws As Object
    Set ws = GetSheet(SH_RESULTS)
    ApplyFrame ws, "Results", "Capital by risk class, component and correlation scenario (MAR21.7), reporting currency.", 9
    SetWidths ws, Array(26, 14, 16, 16, 16)
    WriteOutputPlaceholder ws

    Set ws = GetSheet(SH_OUT_SUMMARY)
    ApplyFrame ws, "Out_Summary", "Key results of the last run, read by the regression test and automation.", 5
    SetWidths ws, Array(18, 20)
    WriteOutputPlaceholder ws

    Set ws = GetSheet(SH_OUT_CLASS)
    ApplyFrame ws, "Out_RiskClass", "Capital per risk class, component and scenario; AltSbUsed = Y when the alternative Sb of MAR21.4(5)(b) applied.", 9
    SetWidths ws, Array(12, 13, 11, 16, 9, 11)
    WriteOutputPlaceholder ws

    Set ws = GetSheet(SH_OUT_BUCKETS)
    ApplyFrame ws, "Out_Buckets", "Kb and Sb per bucket. Notes: ALT_SB alternative Sb, UP/DOWN selected curvature scenario, OTHER other-sector bucket.", 11
    SetWidths ws, Array(12, 13, 11, 9, 16, 16, 11, 16)
    WriteOutputPlaceholder ws

    Set ws = GetSheet(SH_OUT_DRC)
    ApplyFrame ws, "Out_DRC", "Net jump-to-default per obligor after seniority netting, and the hedge benefit ratio and charge per bucket (MAR22).", 18
    SetWidths ws, Array(16, 9, 13, 9, 14, 14, 3, 9, 14, 14, 10, 14, 14, 14)
    WriteOutputPlaceholder ws

    Set ws = GetSheet(SH_LOG)
    ApplyFrame ws, "Log", "Every rejected input row with the reason, warnings and the run status.", 6
    SetWidths ws, Array(9, 24, 110)
    WriteOutputPlaceholder ws
End Sub

'------------------------------------------------------------------------------
'                              INPUT SHEETS
'------------------------------------------------------------------------------
Private Sub BuildInputSheets()
    BuildInputSheet SH_IN_DELTA, "Delta sensitivities in the reporting currency; rows with the same risk factor are netted (MAR21.4(2)).", _
        Array("ID", "RiskClass", "Bucket", "Qualifier", "SubType", "Tenor", "Sensitivity", "Comment"), _
        Array(10, 11, 9, 18, 13, 9, 16, 34), 7
    BuildInputSheet SH_IN_VEGA, "Vega x implied volatility per option maturity (and GIRR YIELD underlying maturity; GIRR INFLATION and XCCY: none), MAR21.25 and MAR21.8 FAQ4.", _
        Array("ID", "RiskClass", "Bucket", "Qualifier", "SubType", "OptionMaturity", "UnderlyingMaturity", "Sensitivity", "Comment"), _
        Array(10, 11, 9, 18, 13, 15, 18, 16, 34), 8
    BuildInputSheet SH_IN_CURV, "Curvature risk CVR up and CVR down per risk factor, net across instruments, positive for a loss and before the FX divisor; shocks per MAR21.98-99 (MAR21.5).", _
        Array("ID", "RiskClass", "Bucket", "Qualifier", "SubType", "CVR_Up", "CVR_Down", "RefReportingCcy", "Comment"), _
        Array(10, 11, 9, 18, 13, 14, 14, 16, 34), 6
    BuildInputSheet SH_IN_DRC, "Positions for the default risk charge: signed notional and market value, maturity in years, optional LGD override (MAR22).", _
        Array("ID", "Obligor", "Bucket", "CreditQuality", "Seniority", "Notional", "MarketValue", "Maturity", "LGD", "Comment"), _
        Array(10, 16, 9, 14, 13, 14, 14, 10, 8, 30), 6
    BuildInputSheet SH_IN_RRAO, "Gross notional of instruments with residual risks: EXOTIC or OTHER (MAR23).", _
        Array("ID", "Category", "GrossNotional", "Comment"), Array(10, 12, 16, 50), 3
End Sub


Private Sub BuildInputSheet(ByVal sheetName As String, ByVal subtitle As String, ByVal headers As Variant, _
                            ByVal widths As Variant, ByVal firstNumCol As Long)
    Dim ws As Object, nCols As Long
    Set ws = GetSheet(sheetName)
    ws.Cells.Clear
    nCols = UBound(headers) - LBound(headers) + 1
    ApplyFrame ws, sheetName, subtitle, LAYOUT_FIRST_COL + nCols - 1
    SetWidths ws, widths
    WriteTableHeader ws, LAYOUT_HEADER_ROW, headers, COLOR_INPUT_HEADER
    FormatDataRows ws, LAYOUT_FIRST_DATA_ROW, LAYOUT_FIRST_DATA_ROW + INPUT_ROWS - 1, nCols, firstNumCol, "#,##0.####", True
    On Error Resume Next
    ws.Range(TableCell(ws, LAYOUT_HEADER_ROW, 1), TableCell(ws, LAYOUT_HEADER_ROW, nCols)).AutoFilter
    ' Short codes (risk class and bucket) centered
    ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, 2), TableCell(ws, LAYOUT_FIRST_DATA_ROW + INPUT_ROWS - 1, 3)).HorizontalAlignment = -4108
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
'                               PARAMETERS
'------------------------------------------------------------------------------
Private Sub WriteParameters()
    WriteParamGeneral
    WriteParamGirr
    WriteParamCsr
    WriteParamEquity
    WriteParamCommodity
    WriteParamFx
    WriteParamDrc
    WriteParamRrao
End Sub


Private Function ParamSheet(ByVal sheetName As String, ByVal subtitle As String, ByVal lastCol As Long) As Object
    Dim ws As Object
    Set ws = GetSheet(sheetName)
    ws.Cells.Clear
    ApplyFrame ws, sheetName, subtitle, lastCol
    Set ParamSheet = ws
End Function


Private Sub WriteParamGeneral()
    Dim ws As Object, r As Long
    Set ws = ParamSheet(SH_P_GENERAL, "Correlation scenarios, vega parameters and regulatory vertices. Codes in column C are read by the engine.", 7)
    SetWidths ws, Array(22, 30, 10, 60)

    r = PutTable(ws, LAYOUT_HEADER_ROW, "CORR_SCENARIOS", "Correlation scenarios", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("HIGH_MULTIPLIER", 1.25, "MAR21.6(2): rho_high = min(1.25 x rho, 100%)")
    PutRow ws, r + 1, Array("LOW_MULTIPLIER", 0.75, "MAR21.6(3): rho_low = max(2 x rho - 100%, 75% x rho)")
    FormatParamRows ws, r, r + 1, 3, 2, 2, ""

    r = PutTable(ws, r + 3, "VEGA", "Vega risk weight and maturity correlation", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("RW_SIGMA", 0.55, "MAR21.92: RW = min(RW_sigma x sqrt(LH / 10), 100%)")
    PutRow ws, r + 1, Array("ALPHA", 0.01, "MAR21.93: option maturity correlation parameter")
    FormatParamRows ws, r, r + 1, 3, 2, 2, "2"

    r = PutTable(ws, r + 3, "VEGA_LH", "Vega liquidity horizons (days)", Array("RiskClass", "Bucket", "LH", "Reference"))
    PutRow ws, r, Array("GIRR", "*", 60, "MAR21.92 Table 13")
    PutRow ws, r + 1, Array("CSR_NS", "*", 120, "MAR21.92 Table 13")
    PutRow ws, r + 2, Array("EQ", "*", 20, "MAR21.92: large cap and indices (buckets 1-8, 12, 13)")
    PutRow ws, r + 3, Array("EQ", 9, 60, "MAR21.92: small cap and other")
    PutRow ws, r + 4, Array("EQ", 10, 60, "MAR21.92: small cap and other")
    PutRow ws, r + 5, Array("EQ", 11, 60, "MAR21.92: small cap and other")
    PutRow ws, r + 6, Array("COM", "*", 120, "MAR21.92 Table 13")
    PutRow ws, r + 7, Array("FX", "*", 40, "MAR21.92 Table 13")
    FormatParamRows ws, r, r + 7, 4, 3, 3, ""

    r = PutTable(ws, r + 9, "VERTICES", "Regulatory vertices (years)", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("GIRR_TENORS", "0.25,0.5,1,2,3,5,10,15,20,30", "MAR21.8(1)")
    PutRow ws, r + 1, Array("CSR_TENORS", "0.5,1,3,5,10", "MAR21.9(1)")
    PutRow ws, r + 2, Array("COM_TENORS", "0,0.25,0.5,1,2,3,5,10,15,20,30", "MAR21.13(1)")
    PutRow ws, r + 3, Array("VEGA_MATURITIES", "0.5,1,3,5,10", "MAR21.8(4)")
    FormatParamRows ws, r, r + 3, 3, 2, 2, ""
End Sub


Private Sub WriteParamGirr()
    Dim ws As Object, r As Long, i As Long, tenors As Variant, rws As Variant, ccys As Variant
    Set ws = ParamSheet(SH_P_GIRR, "General interest rate risk: delta risk weights, correlations and specified currencies.", 7)
    SetWidths ws, Array(22, 12, 26)
    tenors = Array(0.25, 0.5, 1, 2, 3, 5, 10, 15, 20, 30)
    rws = Array(0.017, 0.017, 0.016, 0.013, 0.012, 0.011, 0.011, 0.011, 0.011, 0.011)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "GIRR_RW", "Delta risk weights by tenor (MAR21.42 Table 1)", Array("Tenor", "RW"))
    For i = 0 To 9
        PutRow ws, r + i, Array(tenors(i), rws(i))
    Next i
    FormatParamRows ws, r, r + 9, 2, 2, 2, "2"

    r = PutTable(ws, r + 11, "GIRR_PARAMS", "Correlations and other risk weights", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("THETA", 0.03, "MAR21.46")
    PutRow ws, r + 1, Array("RHO_FLOOR", 0.4, "MAR21.46")
    PutRow ws, r + 2, Array("RHO_DIFF_CURVE", 0.999, "MAR21.45, MAR21.47 (also two inflation curves, FAQ)")
    PutRow ws, r + 3, Array("RHO_INFLATION", 0.4, "MAR21.48")
    PutRow ws, r + 4, Array("RHO_XCCY", 0, "MAR21.49")
    PutRow ws, r + 5, Array("RW_INFLATION", 0.016, "MAR21.43")
    PutRow ws, r + 6, Array("RW_XCCY", 0.016, "MAR21.43")
    PutRow ws, r + 7, Array("GAMMA", 0.5, "MAR21.50")
    PutRow ws, r + 8, Array("SPECIFIED_DIVISOR", Sqr(2), "MAR21.44")
    FormatParamRows ws, r, r + 8, 3, 2, 2, "KEY"

    ccys = Array("EUR", "USD", "GBP", "AUD", "JPY", "SEK", "CAD")
    r = PutTable(ws, r + 10, "GIRR_SPECIFIED_CCY", "Specified currencies (MAR21.44)", Array("Currency"))
    For i = 0 To UBound(ccys)
        TableCell(ws, r + i, 1).Value = ccys(i)
    Next i
    FormatParamRows ws, r, r + UBound(ccys), 1, 1, 1, ""
End Sub


Private Sub WriteParamCsr()
    Dim ws As Object, r As Long, i As Long, j As Long, groups As Variant, m As Variant, headers As Variant
    Set ws = ParamSheet(SH_P_CSR, "Credit spread risk, non-securitization: buckets, risk weights and correlations.", 15)
    SetWidths ws, Array(22, 40, 9, 12, 9, 9, 9, 9, 9, 9, 9, 9)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "CSR_BUCKETS", "Buckets and delta risk weights (MAR21.51, MAR21.53 Table 4)", _
        Array("Bucket", "Description", "Quality", "SectorGroup", "Type", "RW"))
    PutRow ws, r, Array(1, "Sovereigns incl. central banks, MDBs", "IG", 1, "NAME", 0.005)
    PutRow ws, r + 1, Array(2, "Local government, government-backed non-financials, education, public administration", "IG", 2, "NAME", 0.01)
    PutRow ws, r + 2, Array(3, "Financials incl. government-backed financials", "IG", 3, "NAME", 0.05)
    PutRow ws, r + 3, Array(4, "Basic materials, energy, industrials, agriculture, manufacturing, mining", "IG", 4, "NAME", 0.03)
    PutRow ws, r + 4, Array(5, "Consumer goods and services, transportation and storage, administrative services", "IG", 5, "NAME", 0.03)
    PutRow ws, r + 5, Array(6, "Technology, telecommunications", "IG", 6, "NAME", 0.02)
    PutRow ws, r + 6, Array(7, "Health care, utilities, professional and technical activities", "IG", 7, "NAME", 0.015)
    PutRow ws, r + 7, Array(8, "Covered bonds (1.5% for AA- or better at the bank's discretion, fn 17: not modeled; an override applies to the whole bucket)", "IG", 8, "NAME", 0.025)
    PutRow ws, r + 8, Array(9, "Sovereigns incl. central banks, MDBs", "HY", 1, "NAME", 0.02)
    PutRow ws, r + 9, Array(10, "Local government, government-backed non-financials, education, public administration", "HY", 2, "NAME", 0.04)
    PutRow ws, r + 10, Array(11, "Financials incl. government-backed financials", "HY", 3, "NAME", 0.12)
    PutRow ws, r + 11, Array(12, "Basic materials, energy, industrials, agriculture, manufacturing, mining", "HY", 4, "NAME", 0.07)
    PutRow ws, r + 12, Array(13, "Consumer goods and services, transportation and storage, administrative services", "HY", 5, "NAME", 0.085)
    PutRow ws, r + 13, Array(14, "Technology, telecommunications", "HY", 6, "NAME", 0.055)
    PutRow ws, r + 14, Array(15, "Health care, utilities, professional and technical activities", "HY", 7, "NAME", 0.05)
    PutRow ws, r + 15, Array(16, "Other sector", "NA", 16, "OTHER", 0.12)
    PutRow ws, r + 16, Array(17, "IG indices", "IG", 17, "INDEX", 0.015)
    PutRow ws, r + 17, Array(18, "HY indices", "HY", 18, "INDEX", 0.05)
    FormatParamRows ws, r, r + 17, 6, 6, 6, "6"
    WrapColumn ws, r, r + 17, 2

    r = PutTable(ws, r + 19, "CSR_PARAMS", "Correlations", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("RHO_NAME", 0.35, "MAR21.54")
    PutRow ws, r + 1, Array("RHO_NAME_INDEX", 0.8, "MAR21.55 (buckets 17-18)")
    PutRow ws, r + 2, Array("RHO_TENOR", 0.65, "MAR21.54")
    PutRow ws, r + 3, Array("RHO_BASIS", 0.999, "MAR21.54 (bond vs CDS curve)")
    PutRow ws, r + 4, Array("GAMMA_RATING", 0.5, "MAR21.57: IG vs HY/NR, both buckets in 1-15")
    FormatParamRows ws, r, r + 4, 3, 2, 2, "KEY"

    ' Sector component of the cross-bucket correlation (MAR21.57 Table 5), symmetric
    groups = Array(1, 2, 3, 4, 5, 6, 7, 8, 16, 17, 18)
    m = Array( _
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
    headers = Array("Group", 1, 2, 3, 4, 5, 6, 7, 8, 16, 17, 18)
    r = PutTable(ws, r + 6, "CSR_GAMMA_SECTOR", "Cross-bucket sector correlation (MAR21.57 Table 5)", headers)
    For i = 0 To 10
        TableCell(ws, r + i, 1).Value = groups(i)
        For j = 0 To 10
            TableCell(ws, r + i, j + 2).Value = m(i)(j)
        Next j
    Next i
    FormatParamRows ws, r, r + 10, 12, 2, 12, "2,3,4,5,6,7,8,9,10,11,12"
End Sub


Private Sub WriteParamEquity()
    Dim ws As Object, r As Long, b As Long, c As Long, g As Double
    Set ws = ParamSheet(SH_P_EQ, "Equity risk: buckets, spot and repo risk weights, correlations.", 16)
    SetWidths ws, Array(18, 40, 9, 10, 10, 10, 9, 9, 9, 9, 9, 9, 9, 9)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "EQ_BUCKETS", "Buckets and delta risk weights (MAR21.72, MAR21.77 Table 10)", _
        Array("Bucket", "Description", "Type", "RW_SPOT", "RW_REPO", "RHO_NAME"))
    PutRow ws, r, Array(1, "Large cap, EM: consumer goods/services, transportation, admin, health care, utilities", "NAME", 0.55, 0.0055, 0.15)
    PutRow ws, r + 1, Array(2, "Large cap, EM: telecommunications, industrials", "NAME", 0.6, 0.006, 0.15)
    PutRow ws, r + 2, Array(3, "Large cap, EM: basic materials, energy, agriculture, manufacturing, mining", "NAME", 0.45, 0.0045, 0.15)
    PutRow ws, r + 3, Array(4, "Large cap, EM: financials, real estate, technology", "NAME", 0.55, 0.0055, 0.15)
    PutRow ws, r + 4, Array(5, "Large cap, AE: consumer goods/services, transportation, admin, health care, utilities", "NAME", 0.3, 0.003, 0.25)
    PutRow ws, r + 5, Array(6, "Large cap, AE: telecommunications, industrials", "NAME", 0.35, 0.0035, 0.25)
    PutRow ws, r + 6, Array(7, "Large cap, AE: basic materials, energy, agriculture, manufacturing, mining", "NAME", 0.4, 0.004, 0.25)
    PutRow ws, r + 7, Array(8, "Large cap, AE: financials, real estate, technology", "NAME", 0.5, 0.005, 0.25)
    PutRow ws, r + 8, Array(9, "Small cap, EM: all sectors", "NAME", 0.7, 0.007, 0.075)
    PutRow ws, r + 9, Array(10, "Small cap, AE: all sectors", "NAME", 0.5, 0.005, 0.125)
    PutRow ws, r + 10, Array(11, "Other sector", "OTHER", 0.7, 0.007, 0)
    PutRow ws, r + 11, Array(12, "Large cap, AE indices", "INDEX", 0.15, 0.0015, 0.8)
    PutRow ws, r + 12, Array(13, "Other indices", "INDEX", 0.25, 0.0025, 0.8)
    FormatParamRows ws, r, r + 12, 6, 4, 6, "4,5,6"
    WrapColumn ws, r, r + 12, 2

    r = PutTable(ws, r + 14, "EQ_PARAMS", "Correlations", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("RHO_SPOT_REPO", 0.999, "MAR21.78")
    FormatParamRows ws, r, r, 3, 2, 2, "KEY"

    r = PutTable(ws, r + 2, "EQ_GAMMA", "Cross-bucket correlation (MAR21.80)", _
        Array("Bucket", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13))
    For b = 1 To 13
        TableCell(ws, r + b - 1, 1).Value = b
        For c = 1 To 13
            If b = c Then
                g = 1
            ElseIf b = 11 Or c = 11 Then
                g = 0
            ElseIf b >= 12 And c >= 12 Then
                g = 0.75
            ElseIf b >= 12 Or c >= 12 Then
                g = 0.45
            Else
                g = 0.15
            End If
            TableCell(ws, r + b - 1, c + 1).Value = g
        Next c
    Next b
    FormatParamRows ws, r, r + 12, 14, 2, 14, "2,3,4,5,6,7,8,9,10,11,12,13,14"
End Sub


Private Sub WriteParamCommodity()
    Dim ws As Object, r As Long, b As Long, c As Long, g As Double
    Set ws = ParamSheet(SH_P_COM, "Commodity risk: buckets, risk weights and correlations.", 14)
    SetWidths ws, Array(16, 45, 9, 10, 9, 9, 9, 9, 9, 9, 9, 9)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "COM_BUCKETS", "Buckets, risk weights and commodity correlation (MAR21.82-83)", _
        Array("Bucket", "Description", "RW", "RHO_CTY"))
    PutRow ws, r, Array(1, "Energy - solid combustibles", 0.3, 0.55)
    PutRow ws, r + 1, Array(2, "Energy - liquid combustibles", 0.35, 0.95)
    PutRow ws, r + 2, Array(3, "Energy - electricity and carbon trading", 0.6, 0.4)
    PutRow ws, r + 3, Array(4, "Freight", 0.8, 0.8)
    PutRow ws, r + 4, Array(5, "Metals - non-precious", 0.4, 0.6)
    PutRow ws, r + 5, Array(6, "Gaseous combustibles", 0.45, 0.65)
    PutRow ws, r + 6, Array(7, "Precious metals (including gold)", 0.2, 0.55)
    PutRow ws, r + 7, Array(8, "Grains and oilseed", 0.35, 0.45)
    PutRow ws, r + 8, Array(9, "Livestock and dairy", 0.25, 0.15)
    PutRow ws, r + 9, Array(10, "Softs and other agriculturals", 0.35, 0.4)
    PutRow ws, r + 10, Array(11, "Other commodity", 0.5, 0.15)
    FormatParamRows ws, r, r + 10, 4, 3, 4, "3,4"

    r = PutTable(ws, r + 12, "COM_PARAMS", "Correlations", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("RHO_TENOR", 0.99, "MAR21.83")
    PutRow ws, r + 1, Array("RHO_BASIS", 0.999, "MAR21.83 (delivery location)")
    FormatParamRows ws, r, r + 1, 3, 2, 2, "KEY"

    r = PutTable(ws, r + 3, "COM_GAMMA", "Cross-bucket correlation (MAR21.85)", _
        Array("Bucket", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11))
    For b = 1 To 11
        TableCell(ws, r + b - 1, 1).Value = b
        For c = 1 To 11
            If b = c Then
                g = 1
            ElseIf b = 11 Or c = 11 Then
                g = 0
            Else
                g = 0.2
            End If
            TableCell(ws, r + b - 1, c + 1).Value = g
        Next c
    Next b
    FormatParamRows ws, r, r + 10, 12, 2, 12, "2,3,4,5,6,7,8,9,10,11,12"
End Sub


Private Sub WriteParamFx()
    Dim ws As Object, r As Long, i As Long, pairs As Variant
    Set ws = ParamSheet(SH_P_FX, "Foreign exchange risk: risk weight, correlation and specified currency pairs.", 7)
    SetWidths ws, Array(22, 12, 40)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "FX_PARAMS", "Risk weight and correlation", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("RW", 0.15, "MAR21.87")
    PutRow ws, r + 1, Array("GAMMA", 0.6, "MAR21.89")
    PutRow ws, r + 2, Array("SPECIFIED_DIVISOR", Sqr(2), "MAR21.88")
    PutRow ws, r + 3, Array("CURVATURE_DIVISOR", 1.5, "MAR21.98")
    FormatParamRows ws, r, r + 3, 3, 2, 2, "KEY"

    pairs = Array("USD/EUR", "USD/JPY", "USD/GBP", "USD/AUD", "USD/CAD", "USD/CHF", "USD/MXN", "USD/CNY", _
                  "USD/NZD", "USD/RUB", "USD/HKD", "USD/SGD", "USD/TRY", "USD/KRW", "USD/SEK", "USD/ZAR", _
                  "USD/INR", "USD/NOK", "USD/BRL", "EUR/JPY", "EUR/GBP", "EUR/CHF", "JPY/AUD")
    r = PutTable(ws, r + 5, "FX_SPECIFIED_PAIRS", "Specified pairs; first-order crosses also qualify (MAR21.88)", Array("Pair"))
    For i = 0 To UBound(pairs)
        TableCell(ws, r + i, 1).Value = pairs(i)
    Next i
    FormatParamRows ws, r, r + UBound(pairs), 1, 1, 1, ""
End Sub


Private Sub WriteParamDrc()
    Dim ws As Object, r As Long
    Set ws = ParamSheet(SH_P_DRC, "Default risk charge, non-securitization: risk weights, LGD, maturity scaling and buckets.", 7)
    SetWidths ws, Array(16, 10, 10, 40)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "DRC_RW", "Risk weights by credit quality (MAR22.24 Table 2)", Array("CreditQuality", "RW"))
    PutRow ws, r, Array("AAA", 0.005)
    PutRow ws, r + 1, Array("AA", 0.02)
    PutRow ws, r + 2, Array("A", 0.03)
    PutRow ws, r + 3, Array("BBB", 0.06)
    PutRow ws, r + 4, Array("BB", 0.15)
    PutRow ws, r + 5, Array("B", 0.3)
    PutRow ws, r + 6, Array("CCC", 0.5)
    PutRow ws, r + 7, Array("UNRATED", 0.15)
    PutRow ws, r + 8, Array("DEFAULTED", 1)
    FormatParamRows ws, r, r + 8, 2, 2, 2, "2"

    ' Rank: 1 = most junior. A short may offset a long of the same or a more senior rank.
    r = PutTable(ws, r + 10, "DRC_SENIORITY", "Seniority rank and LGD (MAR22.12, MAR22.19)", Array("Seniority", "Rank", "LGD", "Reference"))
    PutRow ws, r, Array("EQUITY", 1, 1, "MAR22.12")
    PutRow ws, r + 1, Array("NON_SENIOR", 2, 1, "MAR22.12")
    PutRow ws, r + 2, Array("SENIOR", 3, 0.75, "MAR22.12")
    PutRow ws, r + 3, Array("COVERED", 4, 0.25, "MAR22.12")
    FormatParamRows ws, r, r + 3, 4, 2, 3, "3"

    r = PutTable(ws, r + 5, "DRC_PARAMS", "Maturity scaling (years)", Array("Code", "Value", "Reference"))
    PutRow ws, r, Array("MIN_MATURITY", 0.25, "MAR22.18: floor of three months")
    PutRow ws, r + 1, Array("MAX_MATURITY", 1, "MAR22.15: one-year capital horizon")
    FormatParamRows ws, r, r + 1, 3, 2, 2, ""

    r = PutTable(ws, r + 3, "DRC_BUCKETS", "Buckets (MAR22.22)", Array("Bucket", "Description"))
    PutRow ws, r, Array("CORP", "Corporates")
    PutRow ws, r + 1, Array("SOV", "Sovereigns")
    PutRow ws, r + 2, Array("LGM", "Local governments and municipalities")
    FormatParamRows ws, r, r + 2, 2, 2, 2, ""
End Sub


Private Sub WriteParamRrao()
    Dim ws As Object, r As Long
    Set ws = ParamSheet(SH_P_RRAO, "Residual risk add-on risk weights.", 6)
    SetWidths ws, Array(12, 10, 40)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "RRAO_RW", "Risk weights (MAR23.8)", Array("Category", "RW", "Reference"))
    PutRow ws, r, Array("EXOTIC", 0.01, "MAR23.8(2)(a): exotic underlying")
    PutRow ws, r + 1, Array("OTHER", 0.001, "MAR23.8(2)(b): other residual risks")
    FormatParamRows ws, r, r + 1, 3, 2, 2, "2"
End Sub

'------------------------------------------------------------------------------
'          SYNTHETIC SAMPLE PORTFOLIO (AMOUNTS IN REPORTING CURRENCY)
'------------------------------------------------------------------------------
Private Sub WriteSampleData()
    WriteSampleDelta
    WriteSampleVega
    WriteSampleCurvature
    WriteSampleDrc
    WriteSampleRrao
End Sub

Private Sub WriteSampleDelta()
    Dim ws As Object
    Set ws = GetSheet(SH_IN_DELTA)
    PutSample ws, 1, Array("D001", "GIRR", "EUR", "ESTR", "YIELD", 1, 5000000, "Swap book")
    PutSample ws, 2, Array("D002", "GIRR", "EUR", "ESTR", "YIELD", 5, -12000000, "Swap book")
    PutSample ws, 3, Array("D003", "GIRR", "EUR", "ESTR", "YIELD", 10, 8000000, "Swap book")
    PutSample ws, 4, Array("D004", "GIRR", "EUR", "EURIBOR6M", "YIELD", 5, 3000000, "Basis swaps")
    PutSample ws, 5, Array("D005", "GIRR", "EUR", "EURIBOR6M", "YIELD", 10, -1500000, "Basis swaps")
    PutSample ws, 6, Array("D006", "GIRR", "EUR", "HICP", "INFLATION", "", 400000, "Inflation swaps")
    PutSample ws, 7, Array("D007", "GIRR", "USD", "SOFR", "YIELD", 2, -4000000, "UST hedge")
    PutSample ws, 8, Array("D008", "GIRR", "USD", "SOFR", "YIELD", 10, 2000000, "UST hedge")
    PutSample ws, 9, Array("D009", "GIRR", "USD", "SOFR", "YIELD", 10, 500000, "Netted with D008")
    PutSample ws, 10, Array("D010", "GIRR", "USD", "USD-EUR", "XCCY", "", 1000000, "Cross-currency swaps")
    PutSample ws, 11, Array("D011", "CSR_NS", 1, "ITALY", "BOND", 5, 3000000, "BTP")
    PutSample ws, 12, Array("D012", "CSR_NS", 1, "SPAIN", "BOND", 5, 3000000, "Bono")
    PutSample ws, 13, Array("D013", "CSR_NS", 1, "ITALY", "CDS", 5, 500000, "Sovereign CDS protection sold")
    PutSample ws, 14, Array("D014", "CSR_NS", 2, "REGION_X", "BOND", 3, -2000000, "Regional bond short")
    PutSample ws, 15, Array("D015", "CSR_NS", 2, "REGION_Y", "BOND", 5, -1800000, "Regional bond short")
    PutSample ws, 16, Array("D016", "CSR_NS", 3, "BANK_A", "BOND", 3, 1500000, "Senior bond")
    PutSample ws, 17, Array("D017", "CSR_NS", 3, "BANK_A", "CDS", 3, -1400000, "CDS hedge")
    PutSample ws, 18, Array("D018", "CSR_NS", 11, "BANK_Z", "BOND", 1, 100000, "HY financial")
    PutSample ws, 19, Array("D019", "CSR_NS", 16, "MISC_CO", "BOND", 5, 200000, "Other sector")
    PutSample ws, 20, Array("D020", "CSR_NS", 16, "MISC_CO2", "BOND", 5, -150000, "Other sector")
    PutSample ws, 21, Array("D021", "CSR_NS", 17, "ITRAXX_MAIN", "CDS", 5, -800000, "Index protection")
    PutSample ws, 22, Array("D022", "EQ", 8, "BANK_A", "SPOT", "", 2000000, "Shares")
    PutSample ws, 23, Array("D023", "EQ", 8, "BANK_B", "SPOT", "", -1200000, "Short shares")
    PutSample ws, 24, Array("D024", "EQ", 8, "BANK_A", "REPO", "", 50000, "Securities lending")
    PutSample ws, 25, Array("D025", "EQ", 5, "CONSUMER_X", "SPOT", "", 900000, "Shares")
    PutSample ws, 26, Array("D026", "EQ", 12, "EUROSTOXX50", "SPOT", "", -1500000, "Index future short")
    PutSample ws, 27, Array("D027", "EQ", 11, "OTHER_EQ", "SPOT", "", 100000, "Other sector")
    PutSample ws, 28, Array("D028", "COM", 2, "BRENT", "ROTTERDAM", 0.5, 800000, "Futures")
    PutSample ws, 29, Array("D029", "COM", 2, "BRENT", "ROTTERDAM", 1, -500000, "Futures")
    PutSample ws, 30, Array("D030", "COM", 2, "WTI", "CUSHING", 1, 300000, "Futures")
    PutSample ws, 31, Array("D031", "COM", 7, "GOLD", "LONDON", 0, 1000000, "Spot gold")
    PutSample ws, 32, Array("D032", "COM", 3, "POWER_DE", "EEX", 0.25, 200000, "Power forward")
    PutSample ws, 33, Array("D033", "FX", "USD", "USD", "", "", 3000000, "USD exposure")
    PutSample ws, 34, Array("D034", "FX", "GBP", "GBP", "", "", -1500000, "GBP exposure")
    PutSample ws, 35, Array("D035", "FX", "JPY", "JPY", "", "", 800000, "JPY exposure")
    PutSample ws, 36, Array("D036", "FX", "CHF", "CHF", "", "", -400000, "CHF exposure")
End Sub


Private Sub WriteSampleVega()
    Dim ws As Object
    Set ws = GetSheet(SH_IN_VEGA)
    PutSample ws, 1, Array("V001", "GIRR", "EUR", "", "", 1, 5, 400000, "Swaptions")
    PutSample ws, 2, Array("V002", "GIRR", "EUR", "", "", 5, 10, -250000, "Swaptions")
    PutSample ws, 3, Array("V003", "GIRR", "USD", "", "", 3, 10, 200000, "Swaptions")
    PutSample ws, 4, Array("V004", "CSR_NS", 1, "ITALY", "BOND", 1, "", 60000, "Bond options")
    PutSample ws, 5, Array("V005", "EQ", 8, "BANK_A", "", 0.5, "", 300000, "Single-stock options")
    PutSample ws, 6, Array("V006", "EQ", 8, "BANK_A", "", 1, "", -100000, "Single-stock options")
    PutSample ws, 7, Array("V007", "EQ", 12, "EUROSTOXX50", "", 1, "", -500000, "Index options")
    PutSample ws, 8, Array("V008", "COM", 2, "BRENT", "ROTTERDAM", 1, "", 120000, "Oil options")
    PutSample ws, 9, Array("V009", "FX", "USD", "USD", "", 1, "", 150000, "FX options")
    PutSample ws, 10, Array("V010", "FX", "GBP", "GBP", "", 0.5, "", -80000, "FX options")
End Sub


Private Sub WriteSampleCurvature()
    Dim ws As Object
    Set ws = GetSheet(SH_IN_CURV)
    PutSample ws, 1, Array("C001", "GIRR", "EUR", "", "", -50000, 30000, "", "Swaptions")
    PutSample ws, 2, Array("C002", "GIRR", "USD", "", "", 10000, -20000, "", "Swaptions")
    PutSample ws, 3, Array("C003", "CSR_NS", 1, "ITALY", "BOND", -15000, 5000, "", "Bond options")
    PutSample ws, 4, Array("C004", "EQ", 8, "BANK_A", "", -80000, 40000, "", "Single-stock options")
    PutSample ws, 5, Array("C005", "EQ", 8, "BANK_B", "", 20000, -60000, "", "Single-stock options")
    PutSample ws, 6, Array("C006", "EQ", 12, "EUROSTOXX50", "", 30000, -10000, "", "Index options")
    PutSample ws, 7, Array("C007", "COM", 2, "BRENT", "ROTTERDAM", -25000, 10000, "", "Oil options")
    PutSample ws, 8, Array("C008", "FX", "USD", "USD", "", -40000, -35000, True, "USD/EUR options")
    PutSample ws, 9, Array("C009", "FX", "GBP", "GBP", "", 15000, 12000, "", "GBP/USD options")
End Sub


Private Sub WriteSampleDrc()
    Dim ws As Object
    Set ws = GetSheet(SH_IN_DRC)
    PutSample ws, 1, Array("J001", "ITALY", "SOV", "BBB", "SENIOR", 10000000, 9800000, 5, "", "BTP long")
    PutSample ws, 2, Array("J002", "GERMANY", "SOV", "AAA", "SENIOR", -5000000, -5100000, 10, "", "Bund short")
    PutSample ws, 3, Array("J003", "BANK_A", "CORP", "A", "SENIOR", 4000000, 4050000, 3, "", "Senior bond")
    PutSample ws, 4, Array("J004", "BANK_A", "CORP", "A", "EQUITY", 2000000, 2000000, 1, "", "Shares")
    PutSample ws, 5, Array("J005", "BANK_A", "CORP", "A", "SENIOR", -3000000, -3020000, 0.5, "", "Short senior bond")
    PutSample ws, 6, Array("J006", "BANK_B", "CORP", "BBB", "NON_SENIOR", -1500000, -1450000, 2, "", "Short subordinated")
    PutSample ws, 7, Array("J007", "CONSUMER_X", "CORP", "BB", "SENIOR", 1000000, 950000, 0.1, "", "Short-dated bond")
    PutSample ws, 8, Array("J008", "REGION_X", "LGM", "AA", "SENIOR", 2000000, 2010000, 4, "", "Regional bond")
    PutSample ws, 9, Array("J009", "MISC_CO", "CORP", "UNRATED", "SENIOR", 500000, 480000, 2, "", "Unrated bond")
    PutSample ws, 10, Array("J010", "BANK_Z", "CORP", "CCC", "COVERED", 800000, 700000, 1.5, "", "Covered bond")
End Sub


Private Sub WriteSampleRrao()
    Dim ws As Object
    Set ws = GetSheet(SH_IN_RRAO)
    PutSample ws, 1, Array("R001", "EXOTIC", 5000000, "Weather derivative")
    PutSample ws, 2, Array("R002", "OTHER", 20000000, "Correlation basket option")
    PutSample ws, 3, Array("R003", "OTHER", 7500000, "Behavioral prepayment option")
End Sub

'------------------------------------------------------------------------------
'   EXPECTED RESULTS OF THE SAMPLE PORTFOLIO FOR THE REGRESSION TEST
'   (values from the independent reference implementation under tests/)
'------------------------------------------------------------------------------
Private Sub WriteExpectedResults()
    Dim ws As Object, r As Long, n As Long
    Set ws = GetSheet(SH_TEST)
    ws.Cells.Clear
    ApplyFrame ws, "Test_Expected", "Expected results of the sample portfolio with default parameters and settings. " & _
        "Run Run FRTB-SA, then Run regression test.", 8
    SetWidths ws, Array(28, 18, 18, 12, 10)
    r = PutTable(ws, LAYOUT_HEADER_ROW, "EXPECTED", "Expected values", Array("Item", "Expected", "Actual", "AbsDiff", "Status"))
    n = WriteExpectedRows(ws, r)
    FormatDataRows ws, r, r + n - 1, 5, 2, "#,##0.00", False
End Sub


Private Function WriteExpectedRows(ByVal ws As Object, ByVal r As Long) As Long
    ' BEGIN GENERATED EXPECTED VALUES
    PutRow ws, r + 0, Array("SBM_LOW", 3163703.416013)
    PutRow ws, r + 1, Array("SBM_MEDIUM", 3039394.382502)
    PutRow ws, r + 2, Array("SBM_HIGH", 2932555.615174)
    PutRow ws, r + 3, Array("SBM", 3163703.416013)
    PutRow ws, r + 4, Array("DRC_NONSEC", 636589.396433)
    PutRow ws, r + 5, Array("RRAO", 77500)
    PutRow ws, r + 6, Array("TOTAL", 3877792.812446)
    PutRow ws, r + 7, Array("GIRR|DELTA|LOW", 60423.50359)
    PutRow ws, r + 8, Array("GIRR|DELTA|MEDIUM", 56980.229937)
    PutRow ws, r + 9, Array("GIRR|DELTA|HIGH", 53315.039354)
    PutRow ws, r + 10, Array("GIRR|VEGA|LOW", 323277.327073)
    PutRow ws, r + 11, Array("GIRR|VEGA|MEDIUM", 319771.9736)
    PutRow ws, r + 12, Array("GIRR|VEGA|HIGH", 316227.766017)
    PutRow ws, r + 13, Array("GIRR|CURVATURE|LOW", 33354.16016)
    PutRow ws, r + 14, Array("GIRR|CURVATURE|MEDIUM", 33911.649916)
    PutRow ws, r + 15, Array("GIRR|CURVATURE|HIGH", 34460.12188)
    PutRow ws, r + 16, Array("CSR_NS|DELTA|LOW", 62844.491179)
    PutRow ws, r + 17, Array("CSR_NS|DELTA|MEDIUM", 48260.890512)
    PutRow ws, r + 18, Array("CSR_NS|DELTA|HIGH", 60572.001077)
    PutRow ws, r + 19, Array("CSR_NS|VEGA|LOW", 60000)
    PutRow ws, r + 20, Array("CSR_NS|VEGA|MEDIUM", 60000)
    PutRow ws, r + 21, Array("CSR_NS|VEGA|HIGH", 60000)
    PutRow ws, r + 22, Array("CSR_NS|CURVATURE|LOW", 5000)
    PutRow ws, r + 23, Array("CSR_NS|CURVATURE|MEDIUM", 5000)
    PutRow ws, r + 24, Array("CSR_NS|CURVATURE|HIGH", 5000)
    PutRow ws, r + 25, Array("EQ|DELTA|LOW", 1156963.1261)
    PutRow ws, r + 26, Array("EQ|DELTA|MEDIUM", 1109540.409267)
    PutRow ws, r + 27, Array("EQ|DELTA|HIGH", 1059848.3154)
    PutRow ws, r + 28, Array("EQ|VEGA|LOW", 367878.895929)
    PutRow ws, r + 29, Array("EQ|VEGA|MEDIUM", 348369.331364)
    PutRow ws, r + 30, Array("EQ|VEGA|HIGH", 327700.320415)
    PutRow ws, r + 31, Array("EQ|CURVATURE|LOW", 45746.584572)
    PutRow ws, r + 32, Array("EQ|CURVATURE|MEDIUM", 44237.992721)
    PutRow ws, r + 33, Array("EQ|CURVATURE|HIGH", 42676.105727)
    PutRow ws, r + 34, Array("COM|DELTA|LOW", 352784.668318)
    PutRow ws, r + 35, Array("COM|DELTA|MEDIUM", 366440.324064)
    PutRow ws, r + 36, Array("COM|DELTA|HIGH", 379605.057922)
    PutRow ws, r + 37, Array("COM|VEGA|LOW", 120000)
    PutRow ws, r + 38, Array("COM|VEGA|MEDIUM", 120000)
    PutRow ws, r + 39, Array("COM|VEGA|HIGH", 120000)
    PutRow ws, r + 40, Array("COM|CURVATURE|LOW", 10000)
    PutRow ws, r + 41, Array("COM|CURVATURE|MEDIUM", 10000)
    PutRow ws, r + 42, Array("COM|CURVATURE|HIGH", 10000)
    PutRow ws, r + 43, Array("FX|DELTA|LOW", 430894.418622)
    PutRow ws, r + 44, Array("FX|DELTA|MEDIUM", 396465.635333)
    PutRow ws, r + 45, Array("FX|DELTA|HIGH", 358747.822293)
    PutRow ws, r + 46, Array("FX|VEGA|LOW", 134536.240471)
    PutRow ws, r + 47, Array("FX|VEGA|MEDIUM", 120415.945788)
    PutRow ws, r + 48, Array("FX|VEGA|HIGH", 104403.065089)
    PutRow ws, r + 49, Array("FX|CURVATURE|LOW", 0)
    PutRow ws, r + 50, Array("FX|CURVATURE|MEDIUM", 0)
    PutRow ws, r + 51, Array("FX|CURVATURE|HIGH", 0)
    WriteExpectedRows = 52
    ' END GENERATED EXPECTED VALUES
End Function
