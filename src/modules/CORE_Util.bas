Attribute VB_Name = "CORE_Util"
'==============================================================================
' MODULE: CORE_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Provide sheet access, tagged parameter tables, keyed lookups, value
'   conversions, table formatting in the workbook palette and the run log.
'   No regulatory logic lives here.
'
' PUBLIC SURFACE
'   Sheets: SheetExists, GetSheet, TableCell, ClearTableArea.
'   Tagged tables: FindTagRow, ReadTable, ReadKeyValues, KVNumber, KVText,
'   KVBool, TableNumber. A tagged table starts with a heading cell in
'   column C whose text begins with "[TAG]"; the next row is the header and
'   data rows continue to the next tag or used-range end. Blank keys and
'   internal blank rows raise rather than silently truncating the table.
'   Keyed lookups on Collections: KeyExists, KeyIndex, AddKey.
'   Conversions: IsBlank, IsBlankOrNumber, IsBlankOrBool, IsBlankOrBoolCell,
'   CellBool, ToDouble, IsNumber, ToBool, ToText, BucketLabel, MaxD, MinD, NearlyEqual, TenorKey, ParseNumberList,
'   InNumberList.
'   Sorting: BucketLess, SortLabels.
'   Formatting: WriteTableHeader, FormatDataRows, HeaderCaption, FormatBorders,
'   ExtendPrintArea.
'   Output placeholders: WriteOutputPlaceholder.
'   Log: ResetLog, LogMessage, LogErrorCount, LogWarningCount.
'
' DEPENDENCIES
'   CORE_Config for sheet names, layout and palette.
'
' STATE OWNERSHIP
'   Owns the log cursor and the error and warning counters of the current
'   run (mLogRow, mErrorCount, mWarningCount); ResetLog resets them.
'
' ERROR POLICY
'   GetSheet and ReadTable raise vbObjectError + ERR_APP for a missing sheet
'   or table, WriteOutputPlaceholder for a sheet that is not an output;
'   KVNumber, TableNumber and ParseNumberList raise it for a required
'   parameter that is blank or not a number. Formatting is cosmetic: a host
'   that rejects a format property leaves the cell unformatted instead of
'   stopping the run.
'
' COMPATIBILITY
'   Excel VBA; also runs in LibreOffice Calc with VBA support for the
'   automated tests. No references beyond the defaults.
'==============================================================================

'------------------------------------------------------------------------------
'                              MODULE SETTINGS
'------------------------------------------------------------------------------
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
'                              MODULE CONSTANTS
'------------------------------------------------------------------------------
Public Const ERR_APP As Long = 2048                          'Offset from vbObjectError
Private Const TAG_SCAN_ROWS As Long = 1000                   'Rows searched for a table tag

'------------------------------------------------------------------------------
'                              MODULE STATE
'------------------------------------------------------------------------------
Private mLogRow As Long                                      'Last row written on the Log sheet
Private mErrorCount As Long                                  'ERROR lines of the current run
Private mWarningCount As Long                                'WARN lines of the current run

'------------------------------------------------------------------------------
'                                  SHEETS
'------------------------------------------------------------------------------
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Object
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Err.Clear
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function


Public Function GetSheet(ByVal sheetName As String) As Object
    If Not SheetExists(sheetName) Then
        Err.Raise vbObjectError + ERR_APP, "CORE_Util.GetSheet", "Sheet not found: " & sheetName
    End If
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function


' Cell in column k of a table that starts in LAYOUT_FIRST_COL (k = 1 is column C).
Public Function TableCell(ByVal ws As Object, ByVal r As Long, ByVal k As Long) As Object
    Set TableCell = ws.Cells(r, LAYOUT_FIRST_COL + k - 1)
End Function


' Clears everything from the table header row down, keeping the banner rows.
Public Sub ClearTableArea(ByVal ws As Object)
    ws.Range(ws.Cells(LAYOUT_HEADER_ROW, 1), ws.Cells(ws.Rows.Count, 60)).Clear
End Sub

'------------------------------------------------------------------------------
'                              TAGGED TABLES
'------------------------------------------------------------------------------
Public Function FindTagRow(ByVal ws As Object, ByVal tag As String) As Long
    Dim r As Long
    Dim marker As String
    marker = "[" & UCase$(tag) & "]"
    FindTagRow = 0
    For r = 1 To TAG_SCAN_ROWS
        If Left$(UCase$(Trim$(CStr(ws.Cells(r, LAYOUT_FIRST_COL).Value))), Len(marker)) = marker Then
            FindTagRow = r
            Exit Function
        End If
    Next r
End Function


' Returns the data rows of a tagged table as a 1-based 2D array
' (rows x columns) and sets nRows/nCols. Reads to the next tag or used-range
' end; rejects blank keys and internal gaps instead of truncating the table.
Public Function ReadTable(ByVal sheetName As String, ByVal tag As String, _
                          ByRef nRows As Long, ByRef nCols As Long) As Variant
    Dim ws As Object, tagRow As Long, r As Long, c As Long
    Dim lastRow As Long, firstBlank As Long, rowHasData As Boolean
    Dim data() As Variant

    Set ws = GetSheet(sheetName)
    tagRow = FindTagRow(ws, tag)
    If tagRow = 0 Then
        Err.Raise vbObjectError + ERR_APP, "CORE_Util.ReadTable", _
                  "Table [" & tag & "] not found on sheet " & sheetName
    End If

    nCols = 0
    Do While Trim$(CStr(TableCell(ws, tagRow + 1, nCols + 1).Value)) <> ""
        nCols = nCols + 1
    Loop
    nRows = 0
    firstBlank = 0
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = tagRow + 2 To lastRow
        If Left$(ToText(TableCell(ws, r, 1).Value), 1) = "[" Then Exit For
        rowHasData = False
        For c = 1 To nCols
            If Not IsBlank(TableCell(ws, r, c).Value) Then rowHasData = True
        Next c
        If rowHasData Then
            If IsBlank(TableCell(ws, r, 1).Value) Then
                Err.Raise vbObjectError + ERR_APP, "CORE_Util.ReadTable", _
                          "Blank key in [" & tag & "] data row " & r - tagRow - 1 & " (" & sheetName & ")"
            End If
            If firstBlank > 0 Then
                Err.Raise vbObjectError + ERR_APP, "CORE_Util.ReadTable", _
                          "Blank data row in [" & tag & "] data row " & firstBlank & " (" & sheetName & ")"
            End If
            nRows = nRows + 1
        ElseIf firstBlank = 0 Then
            firstBlank = r - tagRow - 1
        End If
    Next r

    If nRows = 0 Or nCols = 0 Then
        ReadTable = Empty
        Exit Function
    End If

    ReDim data(1 To nRows, 1 To nCols)
    For r = 1 To nRows
        For c = 1 To nCols
            data(r, c) = TableCell(ws, tagRow + 1 + r, c).Value
        Next c
    Next r
    ReadTable = data
End Function


' Reads a [TAG] Key | Value table into a Collection keyed by upper-case key.
Public Function ReadKeyValues(ByVal sheetName As String, ByVal tag As String) As Collection
    Dim t As Variant, nRows As Long, nCols As Long, r As Long
    Dim col As New Collection
    t = ReadTable(sheetName, tag, nRows, nCols)
    For r = 1 To nRows
        On Error Resume Next
        col.Add t(r, 2), UCase$(Trim$(CStr(t(r, 1))))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            LogMessage "WARN", sheetName, "Duplicate key ignored in [" & tag & "]: " & CStr(t(r, 1))
        End If
        On Error GoTo 0
    Next r
    Set ReadKeyValues = col
End Function


' Value of a required key; raises an error when missing or nonnumeric.
Public Function KVNumber(ByVal col As Collection, ByVal key As String, ByVal source As String) As Double
    Dim v As Variant
    If Not KeyExists(col, key) Then
        Err.Raise vbObjectError + ERR_APP, "CORE_Util.KVNumber", "Missing parameter " & key & " (" & source & ")"
    End If
    v = col.Item(UCase$(key))
    If Not IsNumber(v) Then
        Err.Raise vbObjectError + ERR_APP, "CORE_Util.KVNumber", _
                  "Invalid numeric parameter " & key & " (" & source & ")"
    End If
    KVNumber = ToDouble(v, 0#)
End Function


' Required number in data row r, column c of a table returned by ReadTable;
' raises an error naming the sheet, table, row and column when the value is
' blank or not a number, so that a parameter is never read as zero.
Public Function TableNumber(ByRef t As Variant, ByVal r As Long, ByVal c As Long, _
                            ByVal sheetName As String, ByVal tag As String) As Double
    If Not IsNumber(t(r, c)) Then
        Err.Raise vbObjectError + ERR_APP, "CORE_Util.TableNumber", _
                  "Invalid numeric parameter in [" & tag & "] data row " & r & ", column " & c & _
                  " (" & sheetName & "): " & DescribeValue(t(r, c))
    End If
    TableNumber = ToDouble(t(r, c), 0#)
End Function


Private Function DescribeValue(ByVal v As Variant) As String
    If IsError(v) Then
        DescribeValue = "error value"
    ElseIf IsBlank(v) Then
        DescribeValue = "blank"
    Else
        DescribeValue = "'" & CStr(v) & "'"
    End If
End Function


Public Function KVText(ByVal col As Collection, ByVal key As String, ByVal defaultValue As String) As String
    If KeyExists(col, key) Then
        KVText = Trim$(CStr(col.Item(UCase$(key))))
    Else
        KVText = defaultValue
    End If
End Function


Public Function KVBool(ByVal col As Collection, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    If KeyExists(col, key) Then
        KVBool = ToBool(col.Item(UCase$(key)), defaultValue)
    Else
        KVBool = defaultValue
    End If
End Function

'------------------------------------------------------------------------------
'                    COLLECTIONS AS KEY -> VALUE MAPS
'------------------------------------------------------------------------------
Public Function KeyExists(ByVal col As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error GoTo Missing
    v = col.Item(UCase$(key))
    KeyExists = True
    Exit Function
Missing:
    KeyExists = False
End Function


' Index stored under key, or 0 when the key is absent.
Public Function KeyIndex(ByVal col As Collection, ByVal key As String) As Long
    Dim v As Variant
    On Error GoTo Missing
    v = col.Item(UCase$(key))
    KeyIndex = CLng(v)
    Exit Function
Missing:
    KeyIndex = 0
End Function


Public Sub AddKey(ByVal col As Collection, ByVal key As String, ByVal value As Variant)
    col.Add value, UCase$(key)
End Sub

'------------------------------------------------------------------------------
'                          CONVERSIONS AND NUMERICS
'------------------------------------------------------------------------------
Public Function IsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        IsBlank = True
    ElseIf IsError(v) Then
        IsBlank = False
    Else
        IsBlank = (Trim$(CStr(v)) = "")
    End If
End Function


' Numeric cells retain their value. Text must match the complete invariant
' decimal grammar (optional sign, dot and E exponent); no grouping or comma.
' Blank/error optional values retain the caller's explicit default.
Public Function ToDouble(ByVal v As Variant, ByVal defaultValue As Double) As Double
    Dim parsed As Double
    If IsBlank(v) Or IsError(v) Then
        ToDouble = defaultValue
    ElseIf TryNumber(v, parsed) Then
        ToDouble = parsed
    Else
        Err.Raise vbObjectError + ERR_APP, "CORE_Util.ToDouble", _
                  "Invalid numeric value; use a numeric cell or dot-decimal text"
    End If
End Function


Public Function IsNumber(ByVal v As Variant) As Boolean
    Dim parsed As Double
    IsNumber = TryNumber(v, parsed)
End Function


' Shared validation/conversion path, independent of the host's text locale.
Private Function TryNumber(ByVal v As Variant, ByRef parsed As Double) As Boolean
    Dim s As String, pos As Long, digits As Long, ch As String
    On Error GoTo Invalid
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Or VarType(v) = vbDate Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(CStr(v))
        If Len(s) = 0 Then Exit Function
        pos = 1
        ch = Mid$(s, pos, 1)
        If ch = "+" Or ch = "-" Then pos = pos + 1
        digits = ReadDigits(s, pos)
        If Mid$(s, pos, 1) = "." Then
            pos = pos + 1
            digits = digits + ReadDigits(s, pos)
        End If
        If digits = 0 Then Exit Function
        If UCase$(Mid$(s, pos, 1)) = "E" Then
            pos = pos + 1
            ch = Mid$(s, pos, 1)
            If ch = "+" Or ch = "-" Then pos = pos + 1
            If ReadDigits(s, pos) = 0 Then Exit Function
        End If
        If pos <> Len(s) + 1 Then Exit Function
        parsed = Val(s)
    Else
        If Not IsNumeric(v) Then Exit Function
        parsed = CDbl(v)
    End If
    TryNumber = True
    Exit Function
Invalid:
    Err.Clear
    TryNumber = False
End Function


Private Function ReadDigits(ByVal s As String, ByRef pos As Long) As Long
    Dim first As Long, ch As String
    first = pos
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch < "0" Or ch > "9" Then Exit Do
        pos = pos + 1
    Loop
    ReadDigits = pos - first
End Function


Public Function ToBool(ByVal v As Variant, ByVal defaultValue As Boolean) As Boolean
    Dim s As String
    If IsBlank(v) Or IsError(v) Then
        ToBool = defaultValue
        Exit Function
    End If
    If VarType(v) = vbBoolean Then
        ToBool = CBool(v)
        Exit Function
    End If
    s = UCase$(Trim$(CStr(v)))
    Select Case s
        Case "TRUE", "YES", "Y", "1", "-1"
            ToBool = True
        Case "FALSE", "NO", "N", "0"
            ToBool = False
        Case Else
            ToBool = defaultValue
    End Select
End Function


Public Function ToText(ByVal v As Variant) As String
    If IsBlank(v) Or IsError(v) Then
        ToText = ""
    Else
        ToText = UCase$(Trim$(CStr(v)))
    End If
End Function


' Bucket label: integers for numbered risk classes, upper-case text otherwise.
' A non-integral number is returned as a label no bucket matches, so that
' 8.6 is rejected instead of being rounded into bucket 9.
Public Function BucketLabel(ByVal v As Variant) As String
    Dim x As Double
    If IsNumber(v) Then
        x = ToDouble(v, 0#)
        If x = Int(x) Then
            BucketLabel = CStr(CLng(x))
        Else
            BucketLabel = "NON-INTEGER " & CStr(v)
        End If
    Else
        BucketLabel = ToText(v)
    End If
End Function


' True for a blank cell or a number: optional numeric inputs (tenor,
' maturity, LGD override) must not read text as zero.
Public Function IsBlankOrNumber(ByVal v As Variant) As Boolean
    IsBlankOrNumber = IsBlank(v) Or IsNumber(v)
End Function


' True for a blank cell, a Boolean, or the text TRUE or FALSE (any case,
' surrounding spaces ignored). An optional flag accepts nothing else: aliases
' such as YES, N or 1 are rejected, not read as a flag (#144).
Public Function IsBlankOrBool(ByVal v As Variant) As Boolean
    If IsBlank(v) Then
        IsBlankOrBool = True
    ElseIf IsError(v) Then
        IsBlankOrBool = False
    ElseIf VarType(v) = vbBoolean Then
        IsBlankOrBool = True
    ElseIf VarType(v) = vbString Then
        Select Case UCase$(Trim$(v))
            Case "TRUE", "FALSE"
                IsBlankOrBool = True
            Case Else
                IsBlankOrBool = False
        End Select
    Else
        IsBlankOrBool = False
    End If
End Function


' The same rule for a cell. A workbook saved by LibreOffice stores a Boolean
' cell as a number formatted to show TRUE or FALSE, which Excel reads as that
' number: such a cell is accepted only when it shows TRUE or FALSE. A cell
' holding a plain 1 or 0 is still rejected (#144).
Public Function IsBlankOrBoolCell(ByVal cell As Object) As Boolean
    Dim v As Variant
    v = cell.Value
    If IsBlankOrBool(v) Then
        IsBlankOrBoolCell = True
    ElseIf IsError(v) Then
        IsBlankOrBoolCell = False
    ElseIf IsNumeric(v) And VarType(v) <> vbString Then
        Select Case UCase$(Trim$(CStr(cell.Text)))
            Case "TRUE", "FALSE"
                IsBlankOrBoolCell = True
            Case Else
                IsBlankOrBoolCell = False
        End Select
    Else
        IsBlankOrBoolCell = False
    End If
End Function

' The flag a cell shows, under the IsBlankOrBoolCell rule: FALSE when blank.
' A number is read from what the cell shows, so a Boolean-formatted 2 that
' shows TRUE is TRUE, as the check accepted it (#144).
Public Function CellBool(ByVal cell As Object) As Boolean
    Dim v As Variant
    v = cell.Value
    If IsBlank(v) Or IsError(v) Then
        CellBool = False
    ElseIf VarType(v) = vbBoolean Then
        CellBool = CBool(v)
    ElseIf VarType(v) = vbString Then
        CellBool = (UCase$(Trim$(v)) = "TRUE")
    Else
        CellBool = (UCase$(Trim$(CStr(cell.Text))) = "TRUE")
    End If
End Function


Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function


Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function


Public Function NearlyEqual(ByVal a As Double, ByVal b As Double) As Boolean
    NearlyEqual = (Abs(a - b) <= 0.0000001 * MaxD(1#, MaxD(Abs(a), Abs(b))))
End Function


' Locale-independent key for a maturity or tenor in years.
Public Function TenorKey(ByVal t As Double) As String
    TenorKey = CStr(CLng(Round(t * 10000#, 0)))
End Function


' Parses a required list such as "0.25,0.5,1" into a 1-based Double array;
' returns the count. Raises an error naming the source for an empty list, a
' blank entry ("0.25,,1" or a trailing comma) or an entry that is not a
' number, so that no vertex is dropped silently.
Public Function ParseNumberList(ByVal s As String, ByRef values() As Double, ByVal source As String) As Long
    Dim parts As Variant, i As Long, n As Long, part As String
    If Trim$(s) = "" Then
        Err.Raise vbObjectError + ERR_APP, "CORE_Util.ParseNumberList", "Missing list " & source
    End If
    parts = Split(s, ",")
    ReDim values(1 To UBound(parts) - LBound(parts) + 1)
    n = 0
    For i = LBound(parts) To UBound(parts)
        part = Trim$(CStr(parts(i)))
        If part = "" Then
            Err.Raise vbObjectError + ERR_APP, "CORE_Util.ParseNumberList", "Blank entry in " & source
        ElseIf Not IsNumber(part) Then
            Err.Raise vbObjectError + ERR_APP, "CORE_Util.ParseNumberList", _
                      "Invalid number '" & part & "' in " & source
        End If
        n = n + 1
        values(n) = ToDouble(part, 0#)
    Next i
    ParseNumberList = n
End Function


Public Function InNumberList(ByVal x As Double, ByRef values() As Double, ByVal n As Long) As Boolean
    Dim i As Long
    InNumberList = False
    For i = 1 To n
        If Abs(values(i) - x) < 0.000001 Then
            InNumberList = True
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
'          SORTING OF BUCKET LABELS (NUMERIC WHEN BOTH ARE NUMBERS)
'------------------------------------------------------------------------------
Public Function BucketLess(ByVal a As String, ByVal b As String) As Boolean
    If IsNumber(a) And IsNumber(b) Then
        BucketLess = (ToDouble(a, 0#) < ToDouble(b, 0#))
    Else
        BucketLess = (a < b)
    End If
End Function


Public Sub SortLabels(ByRef labels() As String, ByVal n As Long)
    Dim i As Long, j As Long, tmp As String
    For i = 2 To n
        tmp = labels(i)
        j = i - 1
        Do While j >= 1
            If Not BucketLess(tmp, labels(j)) Then Exit Do
            labels(j + 1) = labels(j)
            j = j - 1
        Loop
        labels(j + 1) = tmp
    Next i
End Sub

'------------------------------------------------------------------------------
'                    TABLE FORMATTING IN THE WORKBOOK PALETTE
'------------------------------------------------------------------------------
' Writes a header row starting at column C and formats it: white bold text
' on headerColor (COLOR_INPUT_HEADER for inputs, COLOR_OUTPUT_HEADER for
' outputs).
Public Sub WriteTableHeader(ByVal ws As Object, ByVal r As Long, ByVal headers As Variant, _
                            ByVal headerColor As Long)
    Dim i As Long, k As Long
    k = 0
    For i = LBound(headers) To UBound(headers)
        k = k + 1
        TableCell(ws, r, k).Value = HeaderCaption(headers(i))
    Next i
    On Error Resume Next
    With ws.Range(TableCell(ws, r, 1), TableCell(ws, r, k))
        .Font.Bold = True
        .Font.Color = COLOR_WHITE
        .Interior.Color = headerColor
        .VerticalAlignment = -4108                           'xlCenter
        .HorizontalAlignment = -4108
        .WrapText = True
    End With
    TableCell(ws, r, 1).HorizontalAlignment = -4131          'xlLeft for the key column
    ws.Rows(r).RowHeight = 42
    On Error GoTo 0
    FormatBorders ws.Range(TableCell(ws, r, 1), TableCell(ws, r, k)), COLOR_GRID, 1
    ExtendPrintArea ws, r, LAYOUT_FIRST_COL + k - 1
End Sub


' Formats rows r1..r2 of a table of nCols columns: thin grey borders, 10 pt,
' an optional number format for columns firstNumCol..nCols, and the input
' fill when the cells are meant to be edited.
Public Sub FormatDataRows(ByVal ws As Object, ByVal r1 As Long, ByVal r2 As Long, ByVal nCols As Long, _
                          ByVal firstNumCol As Long, ByVal numberFormat As String, ByVal isInput As Boolean)
    If r2 < r1 Then Exit Sub
    On Error Resume Next
    With ws.Range(TableCell(ws, r1, 1), TableCell(ws, r2, nCols))
        .Font.Size = 10
        .RowHeight = 20.1
        If isInput Then .Interior.Color = COLOR_INPUT_CELL
    End With
    FormatBorders ws.Range(TableCell(ws, r1, 1), TableCell(ws, r2, nCols)), COLOR_GRID, 1
    If firstNumCol > 0 And numberFormat <> "" Then
        ws.Range(TableCell(ws, r1, firstNumCol), TableCell(ws, r2, nCols)).NumberFormat = numberFormat
    End If
    On Error GoTo 0
    ExtendPrintArea ws, r2, LAYOUT_FIRST_COL + nCols - 1
End Sub


' Grow bounded print areas with the explicitly formatted tables, rather than
' UsedRange (whole-column/row formatting may make it excessively large).
' Output sheets remain unbounded because their results grow at runtime.
Public Sub ExtendPrintArea(ByVal ws As Object, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim current As Object, oldLast As Long
    If ws.Name = SH_RESULTS Or ws.Name = SH_LOG Or Left$(ws.Name, 4) = "Out_" Then Exit Sub
    On Error Resume Next
    If ws.PageSetup.PrintArea <> "" Then
        Set current = ws.Range(ws.PageSetup.PrintArea)
        oldLast = current.Row + current.Rows.Count - 1
        If oldLast > lastRow Then lastRow = oldLast
        oldLast = current.Column + current.Columns.Count - 1
        If oldLast > lastCol Then lastCol = oldLast
    End If
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 2), ws.Cells(lastRow, lastCol)).Address
    On Error GoTo 0
End Sub


' Address each edge explicitly: collection-wide border setters are not
' consistently implemented by spreadsheet compatibility hosts.
Public Sub FormatBorders(ByVal target As Object, ByVal color As Long, ByVal lineStyle As Long)
    Dim edge As Long
    On Error Resume Next
    For edge = 7 To 12
        With target.Borders(edge)
            .Weight = 2                                     'xlThin
            .LineStyle = lineStyle
            .Color = color
        End With
    Next edge
    On Error GoTo 0
End Sub


' Only displayed column captions change; engine keys and positions stay fixed.
Public Function HeaderCaption(ByVal value As Variant) As Variant
    Select Case CStr(value)
        Case "RiskClass": HeaderCaption = "Risk class"
        Case "SubType": HeaderCaption = "Sub-type"
        Case "OptionMaturity": HeaderCaption = "Option maturity (years)"
        Case "UnderlyingMaturity": HeaderCaption = "Underlying maturity (years)"
        Case "CreditQuality": HeaderCaption = "Credit quality"
        Case "MarketValue": HeaderCaption = "Market value"
        Case "GrossNotional": HeaderCaption = "Gross notional"
        Case "CVR_Up": HeaderCaption = "CVR up"
        Case "CVR_Down": HeaderCaption = "CVR down"
        Case "RefReportingCcy": HeaderCaption = "References reporting ccy"
        Case "NetLong": HeaderCaption = "Net long"
        Case "NetShort": HeaderCaption = "Net short"
        Case "WeightedLong": HeaderCaption = "Weighted long"
        Case "WeightedShort": HeaderCaption = "Weighted short"
        Case "RiskFactors": HeaderCaption = "Risk factors"
        Case "AltSbUsed": HeaderCaption = "Alternative Sb used"
        Case "AbsDiff": HeaderCaption = "Absolute difference"
        Case Else: HeaderCaption = value
    End Select
End Function

'------------------------------------------------------------------------------
'                           OUTPUT PLACEHOLDERS
'------------------------------------------------------------------------------
' Writes the empty state of an output sheet (Results, Out_*, Log) below a
' cleared table area: the column headers in the header row (on Out_DRC also
' the bucket table from column J) and a note in the first data row. The
' build, Clear outputs and a failed run all show this state.
Public Sub WriteOutputPlaceholder(ByVal ws As Object, _
                                  Optional ByVal note As String = "No results yet: press Run FRTB-SA.", _
                                  Optional ByVal isAlert As Boolean = False)
    Dim headers As Variant, i As Long
    Select Case ws.Name
        Case SH_RESULTS: headers = Array("Risk class", "Component", "Low", "Medium", "High")
        Case SH_OUT_SUMMARY: headers = Array("Item", "Value")
        Case SH_OUT_CLASS: headers = Array("RiskClass", "Component", "Scenario", "Capital", "Buckets", "AltSbUsed")
        Case SH_OUT_BUCKETS: headers = Array("RiskClass", "Component", "Scenario", "Bucket", "Kb", "Sb", "RiskFactors", "Note")
        Case SH_OUT_DRC: headers = Array("Obligor", "Bucket", "CreditQuality", "RW", "NetLong", "NetShort")
        Case SH_LOG: headers = Array("Level", "Source", "Message")
        Case Else: Err.Raise vbObjectError + ERR_APP, "CORE_Util.WriteOutputPlaceholder", "Not an output sheet: " & ws.Name
    End Select
    WriteTableHeader ws, LAYOUT_HEADER_ROW, headers, COLOR_OUTPUT_HEADER
    If ws.Name = SH_OUT_DRC Then
        headers = Array("Bucket", "NetLong", "NetShort", "HBR", "WeightedLong", "WeightedShort", "DRC")
        For i = 0 To 6
            ws.Cells(LAYOUT_HEADER_ROW, 10 + i).Value = HeaderCaption(headers(i))
        Next i
        On Error Resume Next
        With ws.Range("J5:P5")
            .Font.Bold = True
            .Font.Color = COLOR_WHITE
            .Interior.Color = COLOR_OUTPUT_HEADER
            .WrapText = True
            .HorizontalAlignment = -4108
        End With
        On Error GoTo 0
        FormatBorders ws.Range("J5:P5"), COLOR_GRID, 1
    End If
    TableCell(ws, LAYOUT_FIRST_DATA_ROW, 1).Value = note
    On Error Resume Next
    ws.Rows(LAYOUT_FIRST_DATA_ROW).RowHeight = 20.1
    With TableCell(ws, LAYOUT_FIRST_DATA_ROW, 1)
        .Font.Italic = True
        If isAlert Then
            .Font.Color = COLOR_FAIL_TEXT
            .Interior.Color = COLOR_FAIL_FILL
        Else
            .Font.Color = COLOR_NOTE
        End If
    End With
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
'                                 RUN LOG
'------------------------------------------------------------------------------
Public Sub ResetLog()
    Dim ws As Object
    Set ws = GetSheet(SH_LOG)
    ClearTableArea ws
    WriteTableHeader ws, LAYOUT_HEADER_ROW, Array("Level", "Source", "Message"), COLOR_OUTPUT_HEADER
    mLogRow = LAYOUT_HEADER_ROW
    mErrorCount = 0
    mWarningCount = 0
End Sub


Public Sub LogMessage(ByVal level As String, ByVal source As String, ByVal message As String)
    Dim ws As Object
    If mLogRow < LAYOUT_HEADER_ROW Then mLogRow = LAYOUT_HEADER_ROW
    Set ws = GetSheet(SH_LOG)
    mLogRow = mLogRow + 1
    TableCell(ws, mLogRow, 1).Value = level
    TableCell(ws, mLogRow, 2).Value = source
    TableCell(ws, mLogRow, 3).Value = message
    FormatDataRows ws, mLogRow, mLogRow, 3, 0, "", False
    On Error Resume Next
    Select Case level
        Case "ERROR"
            TableCell(ws, mLogRow, 1).Interior.Color = COLOR_FAIL_FILL
            TableCell(ws, mLogRow, 1).Font.Color = COLOR_FAIL_TEXT
        Case "WARN"
            TableCell(ws, mLogRow, 1).Interior.Color = COLOR_WARN_FILL
            TableCell(ws, mLogRow, 1).Font.Color = COLOR_WARN_TEXT
    End Select
    On Error GoTo 0
    If level = "ERROR" Then mErrorCount = mErrorCount + 1
    If level = "WARN" Then mWarningCount = mWarningCount + 1
End Sub


Public Function LogErrorCount() As Long
    LogErrorCount = mErrorCount
End Function


Public Function LogWarningCount() As Long
    LogWarningCount = mWarningCount
End Function
