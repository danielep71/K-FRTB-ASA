Attribute VB_Name = "CORE_Util"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Provides sheet access, tagged parameter tables, keyed lookups on Collections, value
'   conversions, table formatting in the workbook palette, the run log and the write
'   windows that lift and restore the sheet protection. No regulatory logic lives here.
'
' WHY THIS EXISTS
'   The engine and workbook modules read sheets, tables and values and write the Log
'   through these helpers, so each of those rules is written once.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Offers these groups of procedures:
'     - sheets: look up, read a table cell, clear the table area
'     - tagged tables: find a table by its tag, read it, read key and value pairs and
'       numbers
'     - keyed lookups on Collections
'     - conversions: blank and number tests, numbers, flags, text, bucket labels,
'       tenor keys and number lists
'     - sorting of bucket labels
'     - formatting of table headers, data rows and borders, and the print area
'     - output placeholders
'     - the run log and its error and warning counters
'     - sheet and workbook protection write windows
'
' ERROR POLICY
'   GetSheet and ReadTable raise vbObjectError + ERR_APP for a missing sheet or table,
'   WriteOutputPlaceholder for a sheet that is not an output; KVNumber, TableNumber and
'   ParseNumberList raise it for a required parameter that is blank or not a number.
'   Formatting is cosmetic: a host that rejects a format property leaves the cell
'   unformatted instead of stopping the run. BeginWrite raises vbObjectError + ERR_APP,
'   naming the sheets, when a sheet or the workbook structure has a password of its own.
'
' DEPENDENCIES
'   CORE_Config for sheet names, layout and palette.
'
' NOTES
'   A tagged table starts with a heading cell in column C whose text begins with
'   "[TAG]"; the next row is the header and data rows continue to the next tag or
'   used-range end. Blank keys and internal blank rows raise rather than silently
'   truncating the table.
'
'   The sheets and the workbook structure are protected without a password; a command
'   that writes to locked cells opens a write window with BeginWrite and closes it with
'   EndWrite (#349).
'
'   The module also runs in LibreOffice Calc with VBA support for the automated tests.
'   It needs no references beyond the defaults.
'
' PUBLIC SURFACE
'   Sheets: SheetExists, GetSheet, TableCell, ClearTableArea.
'   Tagged tables: FindTagRow, ReadTable, ReadKeyValues, KVNumber, KVText, KVBool,
'   TableNumber.
'   Keyed lookups on Collections: KeyExists, KeyIndex, AddKey.
'   Conversions: IsBlank, IsBlankOrNumber, IsBlankOrBool, IsBlankOrBoolCell, CellBool,
'   ToDouble, IsNumber, ToBool, ToText, BucketLabel, MaxD, MinD, NearlyEqual, TenorKey,
'   ParseNumberList, InNumberList.
'   Sorting: BucketLess, SortLabels.
'   Formatting: WriteTableHeader, FormatDataRows, HeaderCaption, FormatBorders,
'   ExtendPrintArea.
'   Output placeholders: WriteOutputPlaceholder.
'   Log: ResetLog, LogMessage, LogErrorCount, LogWarningCount.
'   Protection: BeginWrite, EndWrite, CloseWriteWindows.
'
' STATE OWNERSHIP
'   Owns the log cursor and the error and warning counters of the current run (mLogRow,
'   mErrorCount, mWarningCount); ResetLog resets them. Owns the depth of the open write
'   windows (mWriteDepth): the protection comes back when the outermost window closes.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const TAG_SCAN_ROWS     As Long = 1000  'Rows searched for a table tag

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    Public Const ERR_APP            As Long = 2048  'Offset from vbObjectError

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mLogRow                 As Long         'Last row written on the Log sheet
    Private mErrorCount             As Long         'ERROR lines of the current run
    Private mWarningCount           As Long         'WARN lines of the current run
    Private mWriteDepth             As Long         'Open write windows; the sheets are unprotected above 0
    Private mScreenHeld             As Boolean      'True while the outermost write window holds screen updating off
    Private mScreenBeforeWrite      As Boolean      'ScreenUpdating before the outermost write window opened


'
'------------------------------------------------------------------------------
'
'                                    SHEETS
'
'------------------------------------------------------------------------------
'

Public Function SheetExists(ByVal sheetName As String) As Boolean

'
'------------------------------------------------------------------------------
'                         TEST WHETHER A SHEET EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the workbook has a worksheet with the given name.
'
' WHY THIS EXISTS
'   Callers test for a sheet before reading it, without handling the error of a missing
'   one.
'
' INPUTS
'   sheetName
'     Name of the worksheet to look for
'
' RETURNS
'   True when the worksheet exists, False when it does not.
'
' BEHAVIOR
'   Looks the worksheet up by name with errors suppressed and reports whether one was
'   found.
'
' ERROR POLICY
'   A missing sheet gives False and does not raise; the error of the lookup is cleared.
'
' DEPENDENCIES
'   ThisWorkbook.Worksheets
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
    Const PROC_NAME             As String = "CORE_Util.SheetExists"

    Dim FoundSheet              As Object   'Worksheet with the requested name, or Nothing

'------------------------------------------------------------------------------
' LOOK UP SHEET
'------------------------------------------------------------------------------
    'Start with no sheet
        Set FoundSheet = Nothing
    'Look the sheet up by name; a missing sheet leaves Nothing
        On Error Resume Next
        Set FoundSheet = ThisWorkbook.Worksheets(sheetName)
        Err.Clear
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Report whether a sheet was found
        SheetExists = Not (FoundSheet Is Nothing)

End Function

Public Function GetSheet(ByVal sheetName As String) As Object

'
'------------------------------------------------------------------------------
'                            RETURN A SHEET BY NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the worksheet of the given name.
'
' WHY THIS EXISTS
'   Callers get the sheet or a descriptive error naming the missing sheet.
'
' INPUTS
'   sheetName
'     Name of the worksheet
'
' RETURNS
'   The worksheet.
'
' BEHAVIOR
'   Checks that the sheet exists, then returns it.
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP, "Sheet not found: " followed by the name, when the
'   sheet does not exist.
'
' DEPENDENCIES
'   SheetExists, ThisWorkbook.Worksheets
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
    Const PROC_NAME             As String = "CORE_Util.GetSheet"

'------------------------------------------------------------------------------
' VALIDATE SHEET
'------------------------------------------------------------------------------
    'Reject a missing sheet
        If Not SheetExists(sheetName) Then
            Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                "Sheet not found: " & sheetName
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the worksheet
        Set GetSheet = ThisWorkbook.Worksheets(sheetName)

End Function

Public Function TableCell(ByVal ws As Object, ByVal r As Long, ByVal k As Long) As Object

'
'------------------------------------------------------------------------------
'                             RETURN A TABLE CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the cell in column k of a table that starts in LAYOUT_FIRST_COL, in row r.
'
' WHY THIS EXISTS
'   Tables start in column C, so callers address their columns by position in the table.
'
' INPUTS
'   ws
'     Sheet that holds the table
'
'   r
'     Row of the cell
'
'   k
'     Column of the cell within the table; 1 is column C
'
' RETURNS
'   The cell, as an Object.
'
' BEHAVIOR
'   Adds k - 1 to LAYOUT_FIRST_COL and returns that cell of the row.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   LAYOUT_FIRST_COL (CORE_Config)
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
    'Return the cell at the table column
        Set TableCell = ws.Cells(r, LAYOUT_FIRST_COL + k - 1)

End Function

Public Sub ClearTableArea(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                             CLEAR THE TABLE AREA
'------------------------------------------------------------------------------
' PURPOSE
'   Clears everything from the table header row down, keeping the banner rows.
'
' WHY THIS EXISTS
'   Clear resets the cells to the Normal style, Arial at the bottom of the cell, while
'   every cell of the workbook is Aptos and vertically centered (#206).
'
' INPUTS
'   ws
'     Sheet to clear
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears columns 1 to 60 from LAYOUT_HEADER_ROW to the last row, then sets the font
'   name to FONT_NAME and the vertical alignment to center for the same columns.
'
' ERROR POLICY
'   Restoring the font and the alignment ignores errors from the host (On Error Resume
'   Next); the clear itself is not guarded.
'
' DEPENDENCIES
'   LAYOUT_HEADER_ROW, FONT_NAME (CORE_Config)
'
' NOTES
'   The vertical alignment value -4108 is xlCenter.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR CELLS
'------------------------------------------------------------------------------
    'Clear the cells from the header row down, in the first 60 columns
        ws.Range(ws.Cells(LAYOUT_HEADER_ROW, 1), ws.Cells(ws.Rows.Count, 60)).Clear

'------------------------------------------------------------------------------
' RESTORE FORMAT
'------------------------------------------------------------------------------
    'Restore the font and the vertical centering by column, ignoring host errors
        On Error Resume Next
        With ws.Range(ws.Columns(1), ws.Columns(60))
            .Font.Name = FONT_NAME
            .VerticalAlignment = -4108
        End With
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                                TAGGED TABLES
'
'------------------------------------------------------------------------------
'

Public Function FindTagRow(ByVal ws As Object, ByVal tag As String) As Long

'
'------------------------------------------------------------------------------
'                         FIND THE ROW OF A TABLE TAG
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the row of the heading cell of a tagged table.
'
' WHY THIS EXISTS
'   A tagged table is found by the text of its heading in column C, not by a fixed
'   position.
'
' INPUTS
'   ws
'     Sheet to search
'
'   tag
'     Tag text without brackets; the case is ignored
'
' RETURNS
'   The row number, or 0 when no heading in the first TAG_SCAN_ROWS rows begins with the
'   tag.
'
' BEHAVIOR
'   Builds the marker "[TAG]" in upper case, then reads column C from row 1 to
'   TAG_SCAN_ROWS and returns the first row whose trimmed upper-case text begins with
'   the marker.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   TAG_SCAN_ROWS; LAYOUT_FIRST_COL (CORE_Config)
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
    Const PROC_NAME             As String = "CORE_Util.FindTagRow"

    Dim Marker                  As String   'Upper-case tag in brackets
    Dim RowIndex                As Long     'Row being tested

'------------------------------------------------------------------------------
' BUILD MARKER
'------------------------------------------------------------------------------
    'Build the upper-case marker and assume the tag is absent
        Marker = "[" & UCase$(tag) & "]"
        FindTagRow = 0

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Return the first row whose heading begins with the marker
        For RowIndex = 1 To TAG_SCAN_ROWS
            If Left$(UCase$(Trim$(CStr(ws.Cells(RowIndex, LAYOUT_FIRST_COL).Value))), Len(Marker)) = Marker Then
                FindTagRow = RowIndex
                Exit Function
            End If
        Next RowIndex

End Function

Public Function ReadTable( _
    ByVal sheetName As String, _
    ByVal tag As String, _
    ByRef nRows As Long, _
    ByRef nCols As Long) As Variant
'
'------------------------------------------------------------------------------
'                             READ A TAGGED TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the data rows of a tagged table as a 1-based 2D array (rows x columns) and
'   sets nRows and nCols.
'
' WHY THIS EXISTS
'   Parameter tables are read by their tag and must not be truncated silently by a blank
'   key or a gap.
'
' INPUTS
'   sheetName
'     Name of the sheet that holds the table
'
'   tag
'     Tag of the table, without brackets
'
'   nRows
'     Receives the number of data rows
'
'   nCols
'     Receives the number of columns of the header row
'
' RETURNS
'   The array of cell values, or Empty when the table has no data rows or no columns.
'
' BEHAVIOR
'   In order:
'     - gets the sheet and the row of the tag
'     - counts the header cells in the row below the tag until the first blank
'     - walks the rows below the header to the next row whose first cell starts with
'       "[" or to the end of the used range, counting the rows that have data
'     - copies the data rows into the array
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP when the sheet or the table is not found, when a data
'   row has a blank first cell (blank key) and when a blank data row precedes a row with
'   data. The messages name the table, the data row and the sheet.
'
' DEPENDENCIES
'   GetSheet, FindTagRow, TableCell, ToText, IsBlank, ERR_APP
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
    Const PROC_NAME             As String = "CORE_Util.ReadTable"

    Dim SourceSheet             As Object   'Sheet that holds the table
    Dim TagRow                  As Long     'Row of the tag
    Dim RowIndex                As Long     'Row being read
    Dim ColumnIndex             As Long     'Column being read
    Dim LastRow                 As Long     'Last row of the used range
    Dim FirstBlankRow           As Long     'Data row number of the first blank row, 0 when none
    Dim RowHasData              As Boolean  'True when the row has a non-blank cell
    Dim TableData()             As Variant  'Data rows to return

'------------------------------------------------------------------------------
' LOCATE TABLE
'------------------------------------------------------------------------------
    'Get the sheet and the row of the tag
        Set SourceSheet = GetSheet(sheetName)
        TagRow = FindTagRow(SourceSheet, tag)
    'Reject a table that is not on the sheet
        If TagRow = 0 Then
            Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                "Table [" & tag & "] not found on sheet " & sheetName
        End If

'------------------------------------------------------------------------------
' COUNT COLUMNS
'------------------------------------------------------------------------------
    'Count the header cells below the tag until the first blank
        nCols = 0
        Do While Trim$(CStr(TableCell(SourceSheet, TagRow + 1, nCols + 1).Value)) <> vbNullString
            nCols = nCols + 1
        Loop

'------------------------------------------------------------------------------
' COUNT ROWS
'------------------------------------------------------------------------------
    'Count the data rows to the next tag or the end of the used range
        nRows = 0
        FirstBlankRow = 0
        LastRow = SourceSheet.UsedRange.Row + SourceSheet.UsedRange.Rows.Count - 1
        For RowIndex = TagRow + 2 To LastRow
            If Left$(ToText(TableCell(SourceSheet, RowIndex, 1).Value), 1) = "[" Then Exit For
            RowHasData = False
            For ColumnIndex = 1 To nCols
                If Not IsBlank(TableCell(SourceSheet, RowIndex, ColumnIndex).Value) Then RowHasData = True
            Next ColumnIndex
            If RowHasData Then
                If IsBlank(TableCell(SourceSheet, RowIndex, 1).Value) Then
                    Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                        "Blank key in [" & tag & "] data row " & RowIndex - TagRow - 1 & _
                        " (" & sheetName & ")"
                End If
                If FirstBlankRow > 0 Then
                    Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                        "Blank data row in [" & tag & "] data row " & FirstBlankRow & _
                        " (" & sheetName & ")"
                End If
                nRows = nRows + 1
            ElseIf FirstBlankRow = 0 Then
                FirstBlankRow = RowIndex - TagRow - 1
            End If
        Next RowIndex

'------------------------------------------------------------------------------
' HANDLE EMPTY TABLE
'------------------------------------------------------------------------------
    'Return Empty when there are no data rows or no columns
        If nRows = 0 Or nCols = 0 Then
            ReadTable = Empty
            Exit Function
        End If

'------------------------------------------------------------------------------
' COPY DATA
'------------------------------------------------------------------------------
    'Copy the data rows into the array
        ReDim TableData(1 To nRows, 1 To nCols)
        For RowIndex = 1 To nRows
            For ColumnIndex = 1 To nCols
                TableData(RowIndex, ColumnIndex) = TableCell(SourceSheet, TagRow + 1 + RowIndex, ColumnIndex).Value
            Next ColumnIndex
        Next RowIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the array
        ReadTable = TableData

End Function

Public Function ReadKeyValues(ByVal sheetName As String, ByVal tag As String) As Collection

'
'------------------------------------------------------------------------------
'                           READ KEY AND VALUE PAIRS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a [TAG] Key | Value table into a Collection keyed by upper-case key.
'
' WHY THIS EXISTS
'   Parameters are looked up by name; the Collection gives that lookup.
'
' INPUTS
'   sheetName
'     Name of the sheet that holds the table
'
'   tag
'     Tag of the table, without brackets
'
' RETURNS
'   A Collection whose items are the second column and whose keys are the trimmed
'   upper-case first column.
'
' BEHAVIOR
'   Reads the table with ReadTable, then adds one item per row. A key that is already in
'   the Collection is not added again.
'
' ERROR POLICY
'   A duplicate key is logged as WARN (key and table named) and ignored. Errors of
'   ReadTable are not handled here.
'
' DEPENDENCIES
'   ReadTable, LogMessage
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
    Const PROC_NAME             As String = "CORE_Util.ReadKeyValues"

    Dim TableValues             As Variant          'Table read by ReadTable
    Dim RowCount                As Long             'Number of data rows
    Dim ColumnCount             As Long             'Number of columns
    Dim RowIndex                As Long             'Row being added
    Dim KeyValues               As New Collection   'Items keyed by upper-case key

'------------------------------------------------------------------------------
' READ TABLE
'------------------------------------------------------------------------------
    'Read the table as an array
        TableValues = ReadTable(sheetName, tag, RowCount, ColumnCount)

'------------------------------------------------------------------------------
' ADD PAIRS
'------------------------------------------------------------------------------
    'Add each row under its upper-case key; a duplicate key is logged and ignored
        For RowIndex = 1 To RowCount
            On Error Resume Next
            KeyValues.Add TableValues(RowIndex, 2), UCase$(Trim$(CStr(TableValues(RowIndex, 1))))
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                LogMessage "WARN", sheetName, "Duplicate key ignored in [" & tag & "]: " & _
                    CStr(TableValues(RowIndex, 1)), , CStr(TableValues(RowIndex, 1))
            End If
            On Error GoTo 0
        Next RowIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the Collection
        Set ReadKeyValues = KeyValues

End Function

Public Function KVNumber(ByVal col As Collection, ByVal key As String, ByVal source As String) As Double

'
'------------------------------------------------------------------------------
'                            READ A REQUIRED NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the value of a required key as a Double.
'
' WHY THIS EXISTS
'   A parameter that is missing or not numeric must stop the run and name the key.
'
' INPUTS
'   col
'     Collection built by ReadKeyValues
'
'   key
'     Key to read; the case is ignored
'
'   source
'     Name of the table, used in the error messages
'
' RETURNS
'   The numeric value of the key.
'
' BEHAVIOR
'   Checks that the key exists, reads its value, checks that it is a number and converts
'   it.
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP when the key is missing ("Missing parameter") or its
'   value is not a number ("Invalid numeric parameter"); the message names the key and
'   the source.
'
' DEPENDENCIES
'   KeyExists, IsNumber, ToDouble
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
    Const PROC_NAME             As String = "CORE_Util.KVNumber"

    Dim ParameterValue          As Variant  'Value stored under the key

'------------------------------------------------------------------------------
' VALIDATE KEY
'------------------------------------------------------------------------------
    'Reject a missing key
        If Not KeyExists(col, key) Then
            Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                "Missing parameter " & key & " (" & source & ")"
        End If

'------------------------------------------------------------------------------
' READ VALUE
'------------------------------------------------------------------------------
    'Read the value stored under the key
        ParameterValue = col.Item(UCase$(key))
    'Reject a value that is not a number
        If Not IsNumber(ParameterValue) Then
            Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                "Invalid numeric parameter " & key & " (" & source & ")"
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the value as a Double
        KVNumber = ToDouble(ParameterValue, 0#)

End Function

Public Function TableNumber( _
    ByRef t As Variant, _
    ByVal r As Long, _
    ByVal c As Long, _
    ByVal sheetName As String, _
    ByVal tag As String) As Double
'
'------------------------------------------------------------------------------
'                         READ A REQUIRED TABLE NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number in data row r, column c of a table returned by ReadTable.
'
' WHY THIS EXISTS
'   A parameter that is blank or not a number must never be read as zero; the error
'   names where it is.
'
' INPUTS
'   t
'     Array returned by ReadTable
'
'   r
'     Data row
'
'   c
'     Column
'
'   sheetName
'     Sheet of the table, used in the error message
'
'   tag
'     Tag of the table, used in the error message
'
' RETURNS
'   The value converted to a Double.
'
' BEHAVIOR
'   Checks that the cell value is a number, then converts it.
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP naming the table, the data row, the column, the sheet
'   and the value (error value, blank or the text) when the value is not a number.
'
' DEPENDENCIES
'   IsNumber, ToDouble, DescribeValue
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
    Const PROC_NAME             As String = "CORE_Util.TableNumber"

'------------------------------------------------------------------------------
' VALIDATE VALUE
'------------------------------------------------------------------------------
    'Reject a value that is blank or not a number
        If Not IsNumber(t(r, c)) Then
            Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                "Invalid numeric parameter in [" & tag & "] data row " & r & ", column " & c & _
                " (" & sheetName & "): " & DescribeValue(t(r, c))
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the value as a Double
        TableNumber = ToDouble(t(r, c), 0#)

End Function

Private Function DescribeValue(ByVal v As Variant) As String

'
'------------------------------------------------------------------------------
'                        DESCRIBE A VALUE FOR A MESSAGE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns "error value" for an error, "blank" for a blank value and the text between
'   single quotes otherwise.
'
' INPUTS
'   v
'     Value to describe
'
' RETURNS
'   The description.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DESCRIBE
'------------------------------------------------------------------------------
    'Name an error value or a blank, quote any other value
        If IsError(v) Then
            DescribeValue = "error value"
        ElseIf IsBlank(v) Then
            DescribeValue = "blank"
        Else
            DescribeValue = "'" & CStr(v) & "'"
        End If

End Function

Public Function KVText(ByVal col As Collection, ByVal key As String, ByVal defaultValue As String) As String

'
'------------------------------------------------------------------------------
'                            READ AN OPTIONAL TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the text stored under a key, or a default when the key is absent.
'
' WHY THIS EXISTS
'   A text parameter may be left out of its table.
'
' INPUTS
'   col
'     Collection built by ReadKeyValues
'
'   key
'     Key to read; the case is ignored
'
'   defaultValue
'     Text returned when the key is absent
'
' RETURNS
'   The trimmed text of the item, or defaultValue.
'
' BEHAVIOR
'   Returns the item converted to text and trimmed when the key exists, otherwise the
'   default.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   KeyExists
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
    'Return the stored text or the default
        If KeyExists(col, key) Then
            KVText = Trim$(CStr(col.Item(UCase$(key))))
        Else
            KVText = defaultValue
        End If

End Function

Public Function KVBool(ByVal col As Collection, ByVal key As String, ByVal defaultValue As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'                            READ AN OPTIONAL FLAG
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the flag stored under a key, or a default when the key is absent.
'
' WHY THIS EXISTS
'   A flag parameter may be left out of its table.
'
' INPUTS
'   col
'     Collection built by ReadKeyValues
'
'   key
'     Key to read; the case is ignored
'
'   defaultValue
'     Flag returned when the key is absent
'
' RETURNS
'   The flag read by ToBool, or defaultValue.
'
' BEHAVIOR
'   Reads the item with ToBool, using the default for an unrecognized value, when the
'   key exists; otherwise returns the default.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   KeyExists, ToBool
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
    'Return the stored flag or the default
        If KeyExists(col, key) Then
            KVBool = ToBool(col.Item(UCase$(key)), defaultValue)
        Else
            KVBool = defaultValue
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                       COLLECTIONS AS KEY -> VALUE MAPS
'
'------------------------------------------------------------------------------
'

Public Function KeyExists(ByVal col As Collection, ByVal key As String) As Boolean

'
'------------------------------------------------------------------------------
'                          TEST WHETHER A KEY EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the Collection holds an item under the key.
'
' WHY THIS EXISTS
'   A Collection has no test for a key; reading a missing key raises.
'
' INPUTS
'   col
'     Collection to look in
'
'   key
'     Key to look for; the case is ignored
'
' RETURNS
'   True when the key exists, False when it does not.
'
' BEHAVIOR
'   Reads the item under the upper-case key; an error jumps to the handler, which
'   returns False.
'
' ERROR POLICY
'   The error of a missing key is the signal and is not raised.
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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.KeyExists"

    Dim Stored                  As Variant  'Item read under the key

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo Missing

'------------------------------------------------------------------------------
' LOOK UP KEY
'------------------------------------------------------------------------------
    'Read the item under the upper-case key
        Stored = col.Item(UCase$(key))
    'Report the key as found
        KeyExists = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Missing:
    'Report the key as missing
        KeyExists = False

End Function

Public Function KeyIndex(ByVal col As Collection, ByVal key As String) As Long

'
'------------------------------------------------------------------------------
'                          RETURN THE INDEX OF A KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the index stored under a key, or 0 when the key is absent.
'
' WHY THIS EXISTS
'   Collections of positions let callers find an item by name.
'
' INPUTS
'   col
'     Collection to look in
'
'   key
'     Key to look for; the case is ignored
'
' RETURNS
'   The item converted to a Long, or 0 when the key is absent.
'
' BEHAVIOR
'   Reads the item under the upper-case key and converts it; an error jumps to the
'   handler, which returns 0.
'
' ERROR POLICY
'   The error of a missing key is the signal and is not raised.
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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.KeyIndex"

    Dim Stored                  As Variant  'Item read under the key

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo Missing

'------------------------------------------------------------------------------
' LOOK UP KEY
'------------------------------------------------------------------------------
    'Read the item under the upper-case key
        Stored = col.Item(UCase$(key))
    'Return the item as a Long
        KeyIndex = CLng(Stored)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Missing:
    'Return 0 for a missing key
        KeyIndex = 0

End Function

Public Sub AddKey(ByVal col As Collection, ByVal key As String, ByVal value As Variant)

'
'------------------------------------------------------------------------------
'                           ADD AN ITEM UNDER A KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Adds an item to the Collection under the upper-case key.
'
' WHY THIS EXISTS
'   Keys are kept in upper case so that KeyExists and KeyIndex ignore the case.
'
' INPUTS
'   col
'     Collection to add to
'
'   key
'     Key of the item
'
'   value
'     Item to add
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds the item under the upper-case key.
'
' ERROR POLICY
'   A key that is already in the Collection raises the error of Collection.Add; it is
'   not handled here.
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
' ADD ITEM
'------------------------------------------------------------------------------
    'Add the item under the upper-case key
        col.Add value, UCase$(key)

End Sub

'
'------------------------------------------------------------------------------
'
'                           CONVERSIONS AND NUMERICS
'
'------------------------------------------------------------------------------
'

Public Function IsBlank(ByVal v As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                            TEST FOR A BLANK VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a value is Empty, Null or text that is empty after trimming.
'
' WHY THIS EXISTS
'   Cells read as blank must be recognised the same way everywhere.
'
' INPUTS
'   v
'     Value to test
'
' RETURNS
'   True for Empty, Null or blank text; False for an error value and for any other
'   value.
'
' BEHAVIOR
'   Tests Empty and Null, then error values, then the trimmed text.
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
    'Classify the value
        If IsEmpty(v) Or IsNull(v) Then
            IsBlank = True
        ElseIf IsError(v) Then
            IsBlank = False
        Else
            IsBlank = (Trim$(CStr(v)) = vbNullString)
        End If

End Function

Public Function ToDouble(ByVal v As Variant, ByVal defaultValue As Double) As Double

'
'------------------------------------------------------------------------------
'                         CONVERT A VALUE TO A DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a value as a Double, or the default for a blank or error value.
'
' WHY THIS EXISTS
'   Numeric cells keep their value; text must match the complete invariant decimal
'   grammar, so the result does not depend on the text locale of the host.
'
' INPUTS
'   v
'     Value to convert
'
'   defaultValue
'     Value returned for a blank or error value
'
' RETURNS
'   The number, or defaultValue for a blank or error value.
'
' BEHAVIOR
'   Returns the default for a blank or error value, otherwise parses the value with
'   TryNumber.
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP ("Invalid numeric value; use a numeric cell or
'   dot-decimal text") when the value is not a number.
'
' DEPENDENCIES
'   IsBlank, TryNumber
'
' NOTES
'   Text is an optional sign, digits, a dot and an E exponent; there is no grouping and
'   no comma.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.ToDouble"

    Dim ParsedValue             As Double   'Number read from the value

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Use the default for a blank or error value, otherwise parse the value
        If IsBlank(v) Or IsError(v) Then
            ToDouble = defaultValue
        ElseIf TryNumber(v, ParsedValue) Then
            ToDouble = ParsedValue
        Else
            Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                "Invalid numeric value; use a numeric cell or dot-decimal text"
        End If

End Function

Public Function IsNumber(ByVal v As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                              TEST FOR A NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a value is a number under the rule of ToDouble.
'
' WHY THIS EXISTS
'   Validation and conversion use the same rule, independent of the text locale of the
'   host.
'
' INPUTS
'   v
'     Value to test
'
' RETURNS
'   True when TryNumber accepts the value.
'
' BEHAVIOR
'   Calls TryNumber and discards the parsed number.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   TryNumber
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
    Const PROC_NAME             As String = "CORE_Util.IsNumber"

    Dim ParsedValue             As Double   'Number read from the value, not used

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Test the value with the shared parser
        IsNumber = TryNumber(v, ParsedValue)

End Function

Private Function TryNumber(ByVal v As Variant, ByRef parsed As Double) As Boolean

'
'------------------------------------------------------------------------------
'                                PARSE A NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Tests whether a value is a number and returns it in parsed.
'
' WHY THIS EXISTS
'   It is the one validation and conversion path, independent of the text locale of the
'   host.
'
' INPUTS
'   v
'     Value to test
'
'   parsed
'     Receives the number when the function returns True
'
' RETURNS
'   True when the value is a number. False for Empty, Null, error, Boolean and Date
'   values and for text that is not a number.
'
' BEHAVIOR
'   For text, trims it and checks the whole grammar:
'     - an optional sign
'     - digits, with an optional dot and more digits
'     - at least one digit
'     - an optional E (any case) with an optional sign and at least one digit
'     - nothing else
'   Then reads the number with Val. For any other type, requires IsNumeric and converts
'   with CDbl.
'
' ERROR POLICY
'   Does not raise: an error jumps to the handler, which clears it and returns False.
'
' DEPENDENCIES
'   ReadDigits
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
    Const PROC_NAME             As String = "CORE_Util.TryNumber"

    Dim TrimmedText             As String   'Trimmed text of the value
    Dim Position                As Long     'Position being read in the text
    Dim DigitCount              As Long     'Digits found in the mantissa
    Dim CurrentChar             As String   'Character being tested

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo Invalid

'------------------------------------------------------------------------------
' REJECT BY TYPE
'------------------------------------------------------------------------------
    'Reject Empty, Null and error values
        If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    'Reject Boolean and Date values
        If VarType(v) = vbBoolean Or VarType(v) = vbDate Then Exit Function

'------------------------------------------------------------------------------
' READ NUMBER
'------------------------------------------------------------------------------
    'Parse text by the grammar, convert any other value as a number
        If VarType(v) = vbString Then
            TrimmedText = Trim$(CStr(v))
            If Len(TrimmedText) = 0 Then Exit Function
            ' Optional sign
            Position = 1
            CurrentChar = Mid$(TrimmedText, Position, 1)
            If CurrentChar = "+" Or CurrentChar = "-" Then Position = Position + 1
            ' Digits, then an optional dot with more digits
            DigitCount = ReadDigits(TrimmedText, Position)
            If Mid$(TrimmedText, Position, 1) = "." Then
                Position = Position + 1
                DigitCount = DigitCount + ReadDigits(TrimmedText, Position)
            End If
            If DigitCount = 0 Then Exit Function
            ' Optional exponent: E, an optional sign and at least one digit
            If UCase$(Mid$(TrimmedText, Position, 1)) = "E" Then
                Position = Position + 1
                CurrentChar = Mid$(TrimmedText, Position, 1)
                If CurrentChar = "+" Or CurrentChar = "-" Then Position = Position + 1
                If ReadDigits(TrimmedText, Position) = 0 Then Exit Function
            End If
            ' Nothing may follow
            If Position <> Len(TrimmedText) + 1 Then Exit Function
            parsed = Val(TrimmedText)
        Else
            If Not IsNumeric(v) Then Exit Function
            parsed = CDbl(v)
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Report the value as a number
        TryNumber = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Invalid:
    'Clear the error and report the value as not a number
        Err.Clear
        TryNumber = False

End Function

Private Function ReadDigits(ByVal s As String, ByRef pos As Long) As Long

'
'------------------------------------------------------------------------------
'                        COUNT THE DIGITS AT A POSITION
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the digits from position pos of s and moves pos past them.
'
' INPUTS
'   s
'     Text to read
'
'   pos
'     Position to start at; receives the position after the digits
'
' RETURNS
'   The number of digits read, 0 when the character at pos is not a digit.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.ReadDigits"

    Dim StartPosition           As Long     'Position at the start
    Dim CurrentChar             As String   'Character being tested

'------------------------------------------------------------------------------
' COUNT DIGITS
'------------------------------------------------------------------------------
    'Remember the start, then step over the digits
        StartPosition = pos
        Do While pos <= Len(s)
            CurrentChar = Mid$(s, pos, 1)
            If CurrentChar < "0" Or CurrentChar > "9" Then Exit Do
            pos = pos + 1
        Loop
    'Return the number of digits read
        ReadDigits = pos - StartPosition

End Function

Public Function ToBool(ByVal v As Variant, ByVal defaultValue As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'                          CONVERT A VALUE TO A FLAG
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a value as a Boolean, or the default for a blank, error or unrecognized
'   value.
'
' WHY THIS EXISTS
'   Flags come from cells that hold a Boolean, text or a number.
'
' INPUTS
'   v
'     Value to convert
'
'   defaultValue
'     Flag returned for a blank, error or unrecognized value
'
' RETURNS
'   The flag.
'
' BEHAVIOR
'   Returns the default for a blank or error value and the value itself for a Boolean.
'   Any other value is converted to trimmed upper-case text and read as:
'     - TRUE, YES, Y, 1 or -1: True
'     - FALSE, NO, N or 0: False
'     - anything else: the default
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsBlank
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
    Const PROC_NAME             As String = "CORE_Util.ToBool"

    Dim NormalizedText          As String   'Trimmed upper-case text of the value

'------------------------------------------------------------------------------
' HANDLE BLANK AND BOOLEAN
'------------------------------------------------------------------------------
    'Return the default for a blank or error value
        If IsBlank(v) Or IsError(v) Then
            ToBool = defaultValue
            Exit Function
        End If
    'Return a Boolean as it is
        If VarType(v) = vbBoolean Then
            ToBool = CBool(v)
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ TEXT
'------------------------------------------------------------------------------
    'Read the trimmed upper-case text as a flag
        NormalizedText = UCase$(Trim$(CStr(v)))
        Select Case NormalizedText
            Case "TRUE", "YES", "Y", "1", "-1"
                ToBool = True
            Case "FALSE", "NO", "N", "0"
                ToBool = False
            Case Else
                ToBool = defaultValue
        End Select

End Function

Public Function ToText(ByVal v As Variant) As String

'
'------------------------------------------------------------------------------
'                           CONVERT A VALUE TO TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a value as trimmed upper-case text, or an empty string for a blank or error
'   value.
'
' WHY THIS EXISTS
'   Codes read from cells are compared in one normal form.
'
' INPUTS
'   v
'     Value to convert
'
' RETURNS
'   The text, or an empty string.
'
' BEHAVIOR
'   Tests IsBlank and IsError, then trims the text of the value and converts it to upper
'   case.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsBlank
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
    'Return an empty string for a blank or error value, otherwise the normalized text
        If IsBlank(v) Or IsError(v) Then
            ToText = vbNullString
        Else
            ToText = UCase$(Trim$(CStr(v)))
        End If

End Function

Public Function BucketLabel(ByVal v As Variant) As String

'
'------------------------------------------------------------------------------
'                           NORMALIZE A BUCKET LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a bucket label: an integer for a number, upper-case text otherwise.
'
' WHY THIS EXISTS
'   A non-integral number must be rejected, not rounded into another bucket: 8.6 must
'   not become bucket 9.
'
' INPUTS
'   v
'     Bucket value read from a cell
'
' RETURNS
'   The integer as text, "NON-INTEGER " followed by the value for a number that is not
'   an integer, or the upper-case text.
'
' BEHAVIOR
'   For a number, tests whether it equals its integer part. Any other value goes through
'   ToText.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsNumber, ToDouble, ToText
'
' NOTES
'   The returned label for a non-integral number matches no bucket.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.BucketLabel"

    Dim NumericValue            As Double   'Value as a number

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Label a number by its integer value, any other value by its text
        If IsNumber(v) Then
            NumericValue = ToDouble(v, 0#)
            If NumericValue = Int(NumericValue) Then
                BucketLabel = CStr(CLng(NumericValue))
            Else
                BucketLabel = "NON-INTEGER " & CStr(v)
            End If
        Else
            BucketLabel = ToText(v)
        End If

End Function

Public Function IsBlankOrNumber(ByVal v As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                         TEST FOR A BLANK OR A NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True for a blank cell or a number.
'
' WHY THIS EXISTS
'   Optional numeric inputs (tenor, maturity, LGD override) must not read text as zero.
'
' INPUTS
'   v
'     Value to test
'
' RETURNS
'   True when IsBlank or IsNumber is True.
'
' BEHAVIOR
'   Combines IsBlank and IsNumber.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsBlank, IsNumber
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
    'Accept a blank or a number
        IsBlankOrNumber = IsBlank(v) Or IsNumber(v)

End Function

Public Function IsBlankOrBool(ByVal v As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                        TEST FOR A BLANK OR A BOOLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True for a blank cell, a Boolean, or the text TRUE or FALSE (any case,
'   surrounding spaces ignored).
'
' WHY THIS EXISTS
'   An optional flag accepts nothing else: aliases such as YES, N or 1 are rejected, not
'   read as a flag (#144).
'
' INPUTS
'   v
'     Value to test
'
' RETURNS
'   True for the accepted values, False for an error value and for anything else.
'
' BEHAVIOR
'   Tests blank, error, Boolean and string in this order; a string is accepted when its
'   trimmed upper-case text is TRUE or FALSE.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsBlank
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
    'Accept a blank, a Boolean, or the text TRUE or FALSE
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

Public Function IsBlankOrBoolCell(ByVal cell As Object) As Boolean

'
'------------------------------------------------------------------------------
'                     TEST A CELL FOR A BLANK OR A BOOLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True for the cells accepted by IsBlankOrBool, and for a numeric cell that
'   shows TRUE or FALSE.
'
' WHY THIS EXISTS
'   A workbook saved by LibreOffice stores a Boolean cell as a number formatted to show
'   TRUE or FALSE, which Excel reads as that number. A cell holding a plain 1 or 0 is
'   still rejected (#144).
'
' INPUTS
'   cell
'     Cell to test
'
' RETURNS
'   True when the cell is accepted.
'
' BEHAVIOR
'   Applies IsBlankOrBool to the cell value. For a number that is not a string, accepts
'   the cell when its displayed text is TRUE or FALSE.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsBlankOrBool
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
    Const PROC_NAME             As String = "CORE_Util.IsBlankOrBoolCell"

    Dim CellValue               As Variant  'Value of the cell

'------------------------------------------------------------------------------
' READ CELL
'------------------------------------------------------------------------------
    'Read the value of the cell
        CellValue = cell.Value

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Accept the cells IsBlankOrBool accepts, and numbers that show TRUE or FALSE
        If IsBlankOrBool(CellValue) Then
            IsBlankOrBoolCell = True
        ElseIf IsError(CellValue) Then
            IsBlankOrBoolCell = False
        ElseIf IsNumeric(CellValue) And VarType(CellValue) <> vbString Then
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

Public Function CellBool(ByVal cell As Object) As Boolean

'
'------------------------------------------------------------------------------
'                           READ THE FLAG OF A CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the flag a cell shows, under the IsBlankOrBoolCell rule: FALSE when blank.
'
' WHY THIS EXISTS
'   A number is read from what the cell shows, so a Boolean-formatted 2 that shows TRUE
'   is TRUE, as the check accepted it (#144).
'
' INPUTS
'   cell
'     Cell to read
'
' RETURNS
'   The flag.
'
' BEHAVIOR
'   Reads the cell value and returns:
'     - False for a blank or error value
'     - the value for a Boolean
'     - True when the trimmed upper-case text of a string is TRUE
'     - True for any other value when the displayed text is TRUE
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsBlank
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
    Const PROC_NAME             As String = "CORE_Util.CellBool"

    Dim CellValue               As Variant  'Value of the cell

'------------------------------------------------------------------------------
' READ CELL
'------------------------------------------------------------------------------
    'Read the value of the cell
        CellValue = cell.Value

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Read the flag from the value, or from the displayed text for a number
        If IsBlank(CellValue) Or IsError(CellValue) Then
            CellBool = False
        ElseIf VarType(CellValue) = vbBoolean Then
            CellBool = CBool(CellValue)
        ElseIf VarType(CellValue) = vbString Then
            CellBool = (UCase$(Trim$(CellValue)) = "TRUE")
        Else
            CellBool = (UCase$(Trim$(CStr(cell.Text))) = "TRUE")
        End If

End Function

Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double

'
'------------------------------------------------------------------------------
'                           RETURN THE LARGER DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the larger of two Double values.
'
' WHY THIS EXISTS
'   Code that needs a maximum calls one helper.
'
' INPUTS
'   a, b
'     The values to compare
'
' RETURNS
'   The larger value; b when they are equal.
'
' BEHAVIOR
'   Compares a with b.
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
    'Return the larger value
        If a > b Then
            MaxD = a
        Else
            MaxD = b
        End If

End Function

Public Function MinD(ByVal a As Double, ByVal b As Double) As Double

'
'------------------------------------------------------------------------------
'                          RETURN THE SMALLER DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the smaller of two Double values.
'
' WHY THIS EXISTS
'   Code that needs a minimum calls one helper.
'
' INPUTS
'   a, b
'     The values to compare
'
' RETURNS
'   The smaller value; b when they are equal.
'
' BEHAVIOR
'   Compares a with b.
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
    'Return the smaller value
        If a < b Then
            MinD = a
        Else
            MinD = b
        End If

End Function

Public Function NearlyEqual(ByVal a As Double, ByVal b As Double) As Boolean

'
'------------------------------------------------------------------------------
'                      TEST TWO DOUBLES FOR NEAR EQUALITY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when two values differ by at most 0.0000001 times the larger of 1 and
'   their absolute values.
'
' WHY THIS EXISTS
'   Results computed by different routes are compared with a relative tolerance.
'
' INPUTS
'   a, b
'     The values to compare
'
' RETURNS
'   True when the values are nearly equal.
'
' BEHAVIOR
'   Compares the absolute difference with the tolerance.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   MaxD
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
    'Compare the difference with the relative tolerance
        NearlyEqual = (Abs(a - b) <= 0.0000001 * MaxD(1#, MaxD(Abs(a), Abs(b))))

End Function

Public Function TenorKey(ByVal t As Double) As String

'
'------------------------------------------------------------------------------
'                              BUILD A TENOR KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a locale-independent key for a maturity or tenor in years.
'
' WHY THIS EXISTS
'   A tenor is used as a Collection key, and the text of a Double depends on the locale.
'
' INPUTS
'   t
'     Maturity or tenor in years
'
' RETURNS
'   The tenor times 10000, rounded, as an integer in text.
'
' BEHAVIOR
'   Multiplies by 10000, rounds to 0 decimals and converts to Long, then to text.
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
    'Return the tenor in ten-thousandths of a year as text
        TenorKey = CStr(CLng(Round(t * 10000#, 0)))

End Function

Public Function ParseNumberList(ByVal s As String, ByRef values() As Double, ByVal source As String) As Long

'
'------------------------------------------------------------------------------
'                           PARSE A LIST OF NUMBERS
'------------------------------------------------------------------------------
' PURPOSE
'   Parses a required list such as "0.25,0.5,1" into a 1-based Double array and returns
'   the count.
'
' WHY THIS EXISTS
'   No vertex may be dropped silently: an empty list, a blank entry or an entry that is
'   not a number stops the run.
'
' INPUTS
'   s
'     Comma-separated list
'
'   values
'     Receives the numbers, indexed from 1
'
'   source
'     Name of the list, used in the error messages
'
' RETURNS
'   The number of values.
'
' BEHAVIOR
'   Splits the list at commas, sizes values to the number of parts, then trims each
'   part, rejects it when blank or not a number and converts it.
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP naming the source for an empty list, a blank entry
'   (for example "0.25,,1" or a trailing comma) or an entry that is not a number.
'
' DEPENDENCIES
'   IsNumber, ToDouble
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
    Const PROC_NAME             As String = "CORE_Util.ParseNumberList"

    Dim Entries                 As Variant  'Parts of the list
    Dim EntryIndex              As Long     'Part being read
    Dim EntryCount              As Long     'Values stored
    Dim Entry                   As String   'Trimmed text of the part

'------------------------------------------------------------------------------
' VALIDATE LIST
'------------------------------------------------------------------------------
    'Reject an empty list
        If Trim$(s) = vbNullString Then
            Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                "Missing list " & source
        End If

'------------------------------------------------------------------------------
' SPLIT LIST
'------------------------------------------------------------------------------
    'Split the list and size the result
        Entries = Split(s, ",")
        ReDim values(1 To UBound(Entries) - LBound(Entries) + 1)
        EntryCount = 0

'------------------------------------------------------------------------------
' CONVERT ENTRIES
'------------------------------------------------------------------------------
    'Reject a blank or non-numeric entry, convert the others
        For EntryIndex = LBound(Entries) To UBound(Entries)
            Entry = Trim$(CStr(Entries(EntryIndex)))
            If Entry = vbNullString Then
                Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                    "Blank entry in " & source
            ElseIf Not IsNumber(Entry) Then
                Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                    "Invalid number '" & Entry & "' in " & source
            End If
            EntryCount = EntryCount + 1
            values(EntryCount) = ToDouble(Entry, 0#)
        Next EntryIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of values
        ParseNumberList = EntryCount

End Function

Public Function InNumberList(ByVal x As Double, ByRef values() As Double, ByVal n As Long) As Boolean

'
'------------------------------------------------------------------------------
'                      TEST WHETHER A NUMBER IS IN A LIST
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when x is within 0.000001 of one of the first n values.
'
' WHY THIS EXISTS
'   Callers test a value against a list read by ParseNumberList.
'
' INPUTS
'   x
'     Number to look for
'
'   values
'     List indexed from 1
'
'   n
'     Number of values to test
'
' RETURNS
'   True when a value matches.
'
' BEHAVIOR
'   Walks values 1 to n and stops at the first one whose absolute difference with x is
'   below 0.000001.
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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.InNumberList"

    Dim ValueIndex              As Long 'Value being tested

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Assume no match
        InNumberList = False
    'Stop at the first value within the tolerance
        For ValueIndex = 1 To n
            If Abs(values(ValueIndex) - x) < 0.000001 Then
                InNumberList = True
                Exit Function
            End If
        Next ValueIndex

End Function

'
'------------------------------------------------------------------------------
'
'           SORTING OF BUCKET LABELS (NUMERIC WHEN BOTH ARE NUMBERS)
'
'------------------------------------------------------------------------------
'

Public Function BucketLess(ByVal a As String, ByVal b As String) As Boolean

'
'------------------------------------------------------------------------------
'                          COMPARE TWO BUCKET LABELS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when label a sorts before label b.
'
' WHY THIS EXISTS
'   Numbered buckets must sort by value (2 before 10), other labels by text.
'
' INPUTS
'   a, b
'     The labels to compare
'
' RETURNS
'   True when a sorts before b.
'
' BEHAVIOR
'   Compares the numbers when both labels are numbers, otherwise the text.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsNumber, ToDouble
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
    'Compare as numbers when both are numbers, otherwise as text
        If IsNumber(a) And IsNumber(b) Then
            BucketLess = (ToDouble(a, 0#) < ToDouble(b, 0#))
        Else
            BucketLess = (a < b)
        End If

End Function

Public Sub SortLabels(ByRef labels() As String, ByVal n As Long)

'
'------------------------------------------------------------------------------
'                              SORT BUCKET LABELS
'------------------------------------------------------------------------------
' PURPOSE
'   Sorts the first n labels in place, in ascending order of BucketLess.
'
' WHY THIS EXISTS
'   Bucket labels are listed in a stable, numeric-aware order.
'
' INPUTS
'   labels
'     Labels indexed from 1
'
'   n
'     Number of labels to sort
'
' RETURNS
'   Nothing. The array is sorted in place.
'
' BEHAVIOR
'   Insertion sort: each label moves left past the labels that sort after it.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   BucketLess
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
    Const PROC_NAME             As String = "CORE_Util.SortLabels"

    Dim OuterIndex              As Long     'Label being inserted
    Dim InnerIndex              As Long     'Position being compared
    Dim HeldLabel               As String   'Label being inserted

'------------------------------------------------------------------------------
' SORT
'------------------------------------------------------------------------------
    'Insert each label into the sorted part before it
        For OuterIndex = 2 To n
            HeldLabel = labels(OuterIndex)
            InnerIndex = OuterIndex - 1
            Do While InnerIndex >= 1
                If Not BucketLess(HeldLabel, labels(InnerIndex)) Then Exit Do
                labels(InnerIndex + 1) = labels(InnerIndex)
                InnerIndex = InnerIndex - 1
            Loop
            labels(InnerIndex + 1) = HeldLabel
        Next OuterIndex

End Sub

'
'------------------------------------------------------------------------------
'
'                   TABLE FORMATTING IN THE WORKBOOK PALETTE
'
'------------------------------------------------------------------------------
'

Public Sub WriteTableHeader( _
    ByVal ws As Object, _
    ByVal r As Long, _
    ByVal headers As Variant, _
    ByVal headerColor As Long)
'
'------------------------------------------------------------------------------
'                             WRITE A TABLE HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a header row starting at column C and formats it: white bold text on
'   headerColor.
'
' WHY THIS EXISTS
'   Input and output tables share one header look; the color tells them apart
'   (COLOR_INPUT_HEADER for inputs, COLOR_OUTPUT_HEADER for outputs).
'
' INPUTS
'   ws
'     Sheet to write to
'
'   r
'     Header row
'
'   headers
'     Array of column captions
'
'   headerColor
'     Fill color of the header
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - writes each caption through HeaderCaption
'     - sets bold, size, white font, fill, centered alignment and wrap for the header
'       cells, the first cell left-aligned, and the row height to 42
'     - draws the borders and extends the print area
'
' ERROR POLICY
'   Formatting ignores errors from the host (On Error Resume Next).
'
' DEPENDENCIES
'   TableCell, HeaderCaption, FormatBorders, ExtendPrintArea, HEADER_FONT_SIZE,
'   COLOR_WHITE, COLOR_GRID
'
' NOTES
'   The alignment value -4108 is xlCenter and -4131 is xlLeft.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.WriteTableHeader"

    Dim HeaderIndex             As Long 'Caption being written
    Dim ColumnCount             As Long 'Columns written

'------------------------------------------------------------------------------
' WRITE CAPTIONS
'------------------------------------------------------------------------------
    'Write the caption of each header
        ColumnCount = 0
        For HeaderIndex = LBound(headers) To UBound(headers)
            ColumnCount = ColumnCount + 1
            TableCell(ws, r, ColumnCount).Value = HeaderCaption(headers(HeaderIndex))
        Next HeaderIndex

'------------------------------------------------------------------------------
' FORMAT HEADER
'------------------------------------------------------------------------------
    'Format the header cells, ignoring host errors; the key column is left-aligned
        On Error Resume Next
        With ws.Range(TableCell(ws, r, 1), TableCell(ws, r, ColumnCount))
            .Font.Bold = True
            .Font.Size = HEADER_FONT_SIZE
            .Font.Color = COLOR_WHITE
            .Interior.Color = headerColor
            .VerticalAlignment = -4108
            .HorizontalAlignment = -4108
            .WrapText = True
        End With
        TableCell(ws, r, 1).HorizontalAlignment = -4131
        ws.Rows(r).RowHeight = 42
        On Error GoTo 0

'------------------------------------------------------------------------------
' DRAW BORDERS
'------------------------------------------------------------------------------
    'Draw the borders and extend the print area
        FormatBorders ws.Range(TableCell(ws, r, 1), TableCell(ws, r, ColumnCount)), COLOR_GRID, 1
        ExtendPrintArea ws, r, LAYOUT_FIRST_COL + ColumnCount - 1

End Sub

Public Sub FormatDataRows( _
    ByVal ws As Object, _
    ByVal r1 As Long, _
    ByVal r2 As Long, _
    ByVal nCols As Long, _
    ByVal firstNumCol As Long, _
    ByVal numberFormat As String, _
    ByVal isInput As Boolean)
'
'------------------------------------------------------------------------------
'                            FORMAT TABLE DATA ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Formats rows r1 to r2 of a table of nCols columns: thin grey borders, 10 pt, an
'   optional number format for columns firstNumCol to nCols, and the input fill when the
'   cells are meant to be edited.
'
' WHY THIS EXISTS
'   Every table shares one data row look, and the input cells stay editable on a
'   protected sheet (#349).
'
' INPUTS
'   ws
'     Sheet to format
'
'   r1, r2
'     First and last row; nothing happens when r2 < r1
'
'   nCols
'     Number of table columns
'
'   firstNumCol
'     First column that takes the number format; 0 for none
'
'   numberFormat
'     Number format text; blank for none
'
'   isInput
'     True to fill the cells with COLOR_INPUT_CELL and unlock them
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Sets the font size to 10 and the row height to 20.1 for the range, then the input
'   fill and unlocked state if isInput, draws the borders, applies the number format and
'   extends the print area.
'
' ERROR POLICY
'   Formatting ignores errors from the host (On Error Resume Next).
'
' DEPENDENCIES
'   TableCell, FormatBorders, ExtendPrintArea, COLOR_INPUT_CELL, COLOR_GRID
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE RANGE
'------------------------------------------------------------------------------
    'Do nothing for an empty row range
        If r2 < r1 Then Exit Sub

'------------------------------------------------------------------------------
' FORMAT ROWS
'------------------------------------------------------------------------------
    'Format the cells, ignoring host errors; an input cell stays editable on a protected sheet (#349)
        On Error Resume Next
        With ws.Range(TableCell(ws, r1, 1), TableCell(ws, r2, nCols))
            .Font.Size = 10
            .RowHeight = 20.1
            If isInput Then
                .Interior.Color = COLOR_INPUT_CELL
                .Locked = False
            End If
        End With
    'Draw the borders
        FormatBorders ws.Range(TableCell(ws, r1, 1), TableCell(ws, r2, nCols)), COLOR_GRID, 1
    'Apply the number format to the numeric columns
        If firstNumCol > 0 And numberFormat <> vbNullString Then
            ws.Range(TableCell(ws, r1, firstNumCol), TableCell(ws, r2, nCols)).NumberFormat = numberFormat
        End If
        On Error GoTo 0

'------------------------------------------------------------------------------
' EXTEND PRINT AREA
'------------------------------------------------------------------------------
    'Grow the print area to the formatted rows
        ExtendPrintArea ws, r2, LAYOUT_FIRST_COL + nCols - 1

End Sub

Public Sub ExtendPrintArea(ByVal ws As Object, ByVal lastRow As Long, ByVal lastCol As Long)

'
'------------------------------------------------------------------------------
'                            EXTEND THE PRINT AREA
'------------------------------------------------------------------------------
' PURPOSE
'   Grows a bounded print area with the explicitly formatted tables.
'
' WHY THIS EXISTS
'   UsedRange is not used because whole-column or whole-row formatting may make it
'   excessively large. Output sheets remain unbounded because their results grow at
'   runtime.
'
' INPUTS
'   ws
'     Sheet to update
'
'   lastRow
'     Last row to include
'
'   lastCol
'     Last column to include
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Does nothing for Results, Log and the sheets whose name begins with "Out_".
'   Otherwise, when the sheet has a print area, enlarges lastRow and lastCol to cover
'   it, then sets the print area from cell B1 to the last row and column.
'
' ERROR POLICY
'   Errors from the host are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   ws.PageSetup, SH_RESULTS, SH_LOG (CORE_Config)
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
    Const PROC_NAME             As String = "CORE_Util.ExtendPrintArea"

    Dim CurrentArea             As Object   'Existing print area
    Dim OldLast                 As Long     'Last row or column of the existing print area

'------------------------------------------------------------------------------
' SKIP OUTPUT SHEETS
'------------------------------------------------------------------------------
    'Leave the output sheets unbounded
        If ws.Name = SH_RESULTS Or ws.Name = SH_LOG Or Left$(ws.Name, 4) = "Out_" Then Exit Sub

'------------------------------------------------------------------------------
' SET PRINT AREA
'------------------------------------------------------------------------------
    'Keep an existing print area inside the new one, ignoring host errors
        On Error Resume Next
        If ws.PageSetup.PrintArea <> vbNullString Then
            Set CurrentArea = ws.Range(ws.PageSetup.PrintArea)
            OldLast = CurrentArea.Row + CurrentArea.Rows.Count - 1
            If OldLast > lastRow Then lastRow = OldLast
            OldLast = CurrentArea.Column + CurrentArea.Columns.Count - 1
            If OldLast > lastCol Then lastCol = OldLast
        End If
    'Set the print area from column B to the last row and column
        ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 2), ws.Cells(lastRow, lastCol)).Address
        On Error GoTo 0

End Sub

Public Sub FormatBorders(ByVal target As Object, ByVal color As Long, ByVal lineStyle As Long)

'
'------------------------------------------------------------------------------
'                         DRAW THE BORDERS OF A RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Draws thin borders on every edge of a range, addressing each edge explicitly.
'
' WHY THIS EXISTS
'   Collection-wide border setters are not consistently implemented by spreadsheet
'   compatibility hosts.
'
' INPUTS
'   target
'     Range to format
'
'   color
'     Border color
'
'   lineStyle
'     Border line style
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   For each border index 7 to 12, sets weight 2 (xlThin), the line style and the color.
'
' ERROR POLICY
'   Errors from the host are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   target.Borders
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
    Const PROC_NAME             As String = "CORE_Util.FormatBorders"

    Dim EdgeIndex               As Long 'Border index being set

'------------------------------------------------------------------------------
' DRAW BORDERS
'------------------------------------------------------------------------------
    'Draw each edge, ignoring host errors
        On Error Resume Next
        For EdgeIndex = 7 To 12
            With target.Borders(EdgeIndex)
                .Weight = 2
                .LineStyle = lineStyle
                .Color = color
            End With
        Next EdgeIndex
        On Error GoTo 0

End Sub

Public Function HeaderCaption(ByVal value As Variant) As Variant

'
'------------------------------------------------------------------------------
'                           RETURN A HEADER CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the caption displayed for a column key.
'
' WHY THIS EXISTS
'   Only displayed column captions change; engine keys and positions stay fixed.
'
' INPUTS
'   value
'     Column key
'
' RETURNS
'   The caption for a known key, the value itself for any other.
'
' BEHAVIOR
'   Maps the keys RiskClass, SubType, OptionMaturity, UnderlyingMaturity, CreditQuality,
'   MarketValue, GrossNotional, CVR_Up, CVR_Down, RefReportingCcy, NetLong, NetShort,
'   WeightedLong, WeightedShort, RiskFactors, AltSbUsed and AbsDiff to their captions.
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
    'Translate the key into its caption
        Select Case CStr(value)
            Case "RiskClass"
                HeaderCaption = "Risk class"
            Case "SubType"
                HeaderCaption = "Sub-type"
            Case "OptionMaturity"
                HeaderCaption = "Option maturity (years)"
            Case "UnderlyingMaturity"
                HeaderCaption = "Underlying maturity (years)"
            Case "CreditQuality"
                HeaderCaption = "Credit quality"
            Case "MarketValue"
                HeaderCaption = "Market value"
            Case "GrossNotional"
                HeaderCaption = "Gross notional"
            Case "CVR_Up"
                HeaderCaption = "CVR up"
            Case "CVR_Down"
                HeaderCaption = "CVR down"
            Case "RefReportingCcy"
                HeaderCaption = "References reporting ccy"
            Case "NetLong"
                HeaderCaption = "Net long"
            Case "NetShort"
                HeaderCaption = "Net short"
            Case "WeightedLong"
                HeaderCaption = "Weighted long"
            Case "WeightedShort"
                HeaderCaption = "Weighted short"
            Case "RiskFactors"
                HeaderCaption = "Risk factors"
            Case "AltSbUsed"
                HeaderCaption = "Alternative Sb used"
            Case "AbsDiff"
                HeaderCaption = "Absolute difference"
            Case Else
                HeaderCaption = value
        End Select

End Function

'
'------------------------------------------------------------------------------
'
'                             OUTPUT PLACEHOLDERS
'
'------------------------------------------------------------------------------
'

Public Sub WriteOutputPlaceholder( _
    ByVal ws As Object, _
    Optional ByVal note As String = "No results yet: press Run FRTB-SA.", _
    Optional ByVal isAlert As Boolean = False)
'
'------------------------------------------------------------------------------
'                         WRITE AN OUTPUT PLACEHOLDER
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the empty state of an output sheet (Results, Out_*, Log) below a cleared
'   table area: the column headers in the header row (on Out_DRC also the bucket table
'   from column J) and a note in the first data row.
'
' WHY THIS EXISTS
'   The build, Clear outputs and a failed run all show this state.
'
' INPUTS
'   ws
'     Output sheet
'
'   note
'     Text for the first data row; the default asks to press Run FRTB-SA
'
'   isAlert
'     True to show the note with the fail colors
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - chooses the column headers of the sheet
'     - writes the header row
'     - on Out_DRC writes and formats the bucket table header in J5:P5
'     - writes the note in italics in the first data row, in the note color or, for an
'       alert, the fail colors
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP ("Not an output sheet") for any other sheet.
'   Formatting ignores errors from the host (On Error Resume Next).
'
' DEPENDENCIES
'   WriteTableHeader, HeaderCaption, FormatBorders, TableCell, the SH_* names and
'   COLOR_* constants (CORE_Config)
'
' NOTES
'   The alignment value -4108 is xlCenter.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.WriteOutputPlaceholder"

    Dim Headers                 As Variant  'Column headers of the sheet
    Dim ColumnIndex             As Long     'Bucket header column being written

'------------------------------------------------------------------------------
' CHOOSE HEADERS
'------------------------------------------------------------------------------
    'Pick the column headers of the output sheet
        Select Case ws.Name
            Case SH_RESULTS
                Headers = Array("Risk class", "Component", "Low", "Medium", "High")
            Case SH_OUT_SUMMARY
                Headers = Array("Item", "Value")
            Case SH_OUT_CLASS
                Headers = Array("RiskClass", "Component", "Scenario", "Capital", "Buckets", "AltSbUsed")
            Case SH_OUT_BUCKETS
                Headers = Array("RiskClass", "Component", "Scenario", "Bucket", "Kb", "Sb", "RiskFactors", "Note")
            Case SH_OUT_DRC
                Headers = Array("Obligor", "Bucket", "CreditQuality", "RW", "NetLong", "NetShort")
            Case SH_LOG
                Headers = Array("Level", "Source", "Message", "Row", "Reference")
            Case Else
                Err.Raise vbObjectError + ERR_APP, PROC_NAME, _
                    "Not an output sheet: " & ws.Name
        End Select

'------------------------------------------------------------------------------
' WRITE HEADER
'------------------------------------------------------------------------------
    'Write the header row
        WriteTableHeader ws, LAYOUT_HEADER_ROW, Headers, COLOR_OUTPUT_HEADER

'------------------------------------------------------------------------------
' WRITE BUCKET HEADER
'------------------------------------------------------------------------------
    'On Out_DRC, write the bucket table header from column J
        If ws.Name = SH_OUT_DRC Then
            Headers = Array("Bucket", "NetLong", "NetShort", "HBR", "WeightedLong", "WeightedShort", "DRC")
            For ColumnIndex = 0 To 6
                ws.Cells(LAYOUT_HEADER_ROW, 10 + ColumnIndex).Value = HeaderCaption(Headers(ColumnIndex))
            Next ColumnIndex
            ' Aligned as the obligor table (#206); host errors are ignored
            On Error Resume Next
            With ws.Range("J5:P5")
                .Font.Bold = True
                .Font.Size = HEADER_FONT_SIZE
                .Font.Color = COLOR_WHITE
                .Interior.Color = COLOR_OUTPUT_HEADER
                .WrapText = True
                .VerticalAlignment = -4108
                .HorizontalAlignment = -4108
            End With
            On Error GoTo 0
            FormatBorders ws.Range("J5:P5"), COLOR_GRID, 1
        End If

'------------------------------------------------------------------------------
' WRITE NOTE
'------------------------------------------------------------------------------
    'Write the note in the first data row
        TableCell(ws, LAYOUT_FIRST_DATA_ROW, 1).Value = note
    'Show the note in italics, in the fail colors for an alert, ignoring host errors
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

'
'------------------------------------------------------------------------------
'
'                                   RUN LOG
'
'------------------------------------------------------------------------------
'

Public Sub ResetLog()

'
'------------------------------------------------------------------------------
'                              RESET THE RUN LOG
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the Log sheet, writes its header and resets the log cursor and the counters.
'
' WHY THIS EXISTS
'   Each run starts with an empty Log and zero errors and warnings.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Gets the Log sheet, clears its table area, writes the header (Level, Source,
'   Message, Row, Reference) and sets mLogRow to LAYOUT_HEADER_ROW and the two counters
'   to 0.
'
' ERROR POLICY
'   A missing Log sheet raises the error of GetSheet; it is not handled here.
'
' DEPENDENCIES
'   GetSheet, ClearTableArea, WriteTableHeader
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
    Const PROC_NAME             As String = "CORE_Util.ResetLog"

    Dim LogSheet                As Object   'Log sheet

'------------------------------------------------------------------------------
' RESET LOG SHEET
'------------------------------------------------------------------------------
    'Get the Log sheet, clear it and write the header
        Set LogSheet = GetSheet(SH_LOG)
        ClearTableArea LogSheet
        WriteTableHeader LogSheet, LAYOUT_HEADER_ROW, _
            Array("Level", "Source", "Message", "Row", "Reference"), COLOR_OUTPUT_HEADER

'------------------------------------------------------------------------------
' RESET STATE
'------------------------------------------------------------------------------
    'Reset the log cursor and the counters
        mLogRow = LAYOUT_HEADER_ROW
        mErrorCount = 0
        mWarningCount = 0

End Sub

Public Sub LogMessage( _
    ByVal level As String, _
    ByVal source As String, _
    ByVal message As String, _
    Optional ByVal findingRow As Variant, _
    Optional ByVal findingReference As Variant)
'
'------------------------------------------------------------------------------
'                               WRITE A LOG LINE
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one line to the Log sheet and counts it as an error or a warning.
'
' WHY THIS EXISTS
'   findingRow and findingReference name the finding's row and its ID or key (#350).
'   They are Variants because LibreOffice does not give an omitted typed argument its
'   default.
'
' INPUTS
'   level
'     ERROR or WARN; other levels are written without a count or a color
'
'   source
'     Origin of the message, for example a sheet name or "In_Delta row 7"
'
'   message
'     Message text
'
'   findingRow
'     Optional sheet row of the finding
'
'   findingReference
'     Optional ID or key of the finding
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - reads the optional arguments; when no row is given and the source ends in "
'       row N" (for example "In_Delta row 7"), takes N as the row and, when no reference
'       is given, the ID in the first column of that row of the sheet named before " row
'       "
'     - adds a line under the last one and writes level, source, message, row and
'       reference
'     - stores the reference as text
'     - formats the line, with the fail colors for ERROR and the warn colors for WARN
'     - adds 1 to the error or the warning counter
'
' ERROR POLICY
'   Formatting ignores errors from the host (On Error Resume Next). A missing Log sheet
'   raises the error of GetSheet.
'
' DEPENDENCIES
'   GetSheet, TableCell, RowReference, FormatDataRows, SH_LOG and the COLOR_* constants
'   (CORE_Config)
'
' NOTES
'   Text after " row " that is not numeric leaves the row empty.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.LogMessage"

    Dim LogSheet                As Object   'Log sheet
    Dim MarkerPosition          As Long     'Position of " row " in the source, 0 when absent
    Dim SheetRow                As Long     'Sheet row of the finding, 0 when none
    Dim ReferenceText           As String   'ID or key of the finding

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Start the cursor at the header row when it is not set
        If mLogRow < LAYOUT_HEADER_ROW Then mLogRow = LAYOUT_HEADER_ROW
    'Read the optional row and reference
        If Not IsMissing(findingRow) Then SheetRow = CLng(findingRow)
        If Not IsMissing(findingReference) Then ReferenceText = CStr(findingReference)

'------------------------------------------------------------------------------
' DERIVE ROW
'------------------------------------------------------------------------------
    'Without a row, read it from a source that ends in " row N", and the ID of that row
        If SheetRow = 0 Then
            MarkerPosition = InStr(source, " row ")
            If MarkerPosition > 0 Then
                If IsNumeric(Mid$(source, MarkerPosition + 5)) Then
                    SheetRow = CLng(Mid$(source, MarkerPosition + 5))
                    If ReferenceText = vbNullString Then
                        ReferenceText = RowReference(Left$(source, MarkerPosition - 1), SheetRow)
                    End If
                End If
            End If
        End If

'------------------------------------------------------------------------------
' WRITE LINE
'------------------------------------------------------------------------------
    'Add a line under the last one and write the level, source and message
        Set LogSheet = GetSheet(SH_LOG)
        mLogRow = mLogRow + 1
        TableCell(LogSheet, mLogRow, 1).Value = level
        TableCell(LogSheet, mLogRow, 2).Value = source
        TableCell(LogSheet, mLogRow, 3).Value = message
    'Write the row when there is one
        If SheetRow > 0 Then TableCell(LogSheet, mLogRow, 4).Value = SheetRow
    'Write the reference as text, so that an ID such as 00123 stays text
        If ReferenceText <> vbNullString Then
            On Error Resume Next
            TableCell(LogSheet, mLogRow, 5).NumberFormat = "@"
            On Error GoTo 0
            TableCell(LogSheet, mLogRow, 5).Value = ReferenceText
        End If

'------------------------------------------------------------------------------
' FORMAT LINE
'------------------------------------------------------------------------------
    'Format the line
        FormatDataRows LogSheet, mLogRow, mLogRow, 5, 0, vbNullString, False
    'Color the level cell of an error or a warning, ignoring host errors
        On Error Resume Next
        Select Case level
            Case "ERROR"
                TableCell(LogSheet, mLogRow, 1).Interior.Color = COLOR_FAIL_FILL
                TableCell(LogSheet, mLogRow, 1).Font.Color = COLOR_FAIL_TEXT
            Case "WARN"
                TableCell(LogSheet, mLogRow, 1).Interior.Color = COLOR_WARN_FILL
                TableCell(LogSheet, mLogRow, 1).Font.Color = COLOR_WARN_TEXT
            Case Else
                'Other levels keep the default colors; nothing is done
        End Select
        On Error GoTo 0

'------------------------------------------------------------------------------
' COUNT LINE
'------------------------------------------------------------------------------
    'Count the line as an error or a warning
        If level = "ERROR" Then mErrorCount = mErrorCount + 1
        If level = "WARN" Then mWarningCount = mWarningCount + 1

End Sub

Private Function RowReference(ByVal sheetName As String, ByVal sheetRow As Long) As String

'
'------------------------------------------------------------------------------
'                             READ THE ID OF A ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the ID in the first column of a row of an input sheet, trimmed; blank when
'   there is none.
'
' INPUTS
'   sheetName
'     Name of the sheet
'
'   sheetRow
'     Row to read
'
' RETURNS
'   The trimmed text of the first table cell of the row, or an empty string when the
'   sheet does not exist or the cell is blank.
'
' ERROR POLICY
'   Errors are ignored (On Error Resume Next).
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ REFERENCE
'------------------------------------------------------------------------------
    'Read the first table cell of the row when the sheet exists, ignoring errors
        On Error Resume Next
        If SheetExists(sheetName) Then
            If Not IsBlank(TableCell(GetSheet(sheetName), sheetRow, 1).Value) Then
                RowReference = Trim$(CStr(TableCell(GetSheet(sheetName), sheetRow, 1).Value))
            End If
        End If
        On Error GoTo 0

End Function

Public Function LogErrorCount() As Long

'
'------------------------------------------------------------------------------
'                            RETURN THE ERROR COUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of ERROR lines logged in the current run.
'
' WHY THIS EXISTS
'   The total stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The count since the last ResetLog.
'
' BEHAVIOR
'   Returns mErrorCount.
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
    'Return the error count
        LogErrorCount = mErrorCount

End Function

Public Function LogWarningCount() As Long

'
'------------------------------------------------------------------------------
'                           RETURN THE WARNING COUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of WARN lines logged in the current run.
'
' WHY THIS EXISTS
'   The total stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The count since the last ResetLog.
'
' BEHAVIOR
'   Returns mWarningCount.
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
    'Return the warning count
        LogWarningCount = mWarningCount

End Function

'
'------------------------------------------------------------------------------
'
'                                  PROTECTION
'
'------------------------------------------------------------------------------
'

Public Sub BeginWrite()

'
'------------------------------------------------------------------------------
'                             OPEN A WRITE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Opens a write window: the sheets and the workbook structure lose their protection
'   until the outermost window closes.
'
' WHY THIS EXISTS
'   The sheets and the workbook structure are protected, and a command that writes to
'   locked cells needs them unprotected (#349).
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   At the outermost level, turns screen updating off, so that the sheets do not flicker as
'   they are unprotected and protected again, then unprotects through SetProtection with
'   errors suppressed and keeps the error. When that failed, protects again what the call
'   removed, gives screen updating back and raises the kept error. Then adds 1 to the depth
'   of open windows.
'
' ERROR POLICY
'   Raises the error of SetProtection, naming the sheets, when one of them or the
'   workbook structure has a password of its own, and leaves everything as it found it.
'
' DEPENDENCIES
'   SetProtection
'
' NOTES
'   Windows can nest; only the first one unprotects and holds screen updating off. A caller
'   that turned it off before, as a run does, finds it off when the window closes.
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.BeginWrite"

    Dim ErrorNumber             As Long     'Number of the error of SetProtection
    Dim ErrorSource             As String   'Source of the error of SetProtection
    Dim ErrorDescription        As String   'Description of the error of SetProtection

'------------------------------------------------------------------------------
' OPEN WINDOW
'------------------------------------------------------------------------------
    'At the outermost window, hold the screen, lift the protection and keep any error
        If mWriteDepth = 0 Then
            On Error Resume Next
            mScreenBeforeWrite = Application.ScreenUpdating
            Application.ScreenUpdating = False
            mScreenHeld = True
            SetProtection False
            ErrorNumber = Err.Number
            ErrorSource = Err.Source
            ErrorDescription = Err.Description
            ' Put back what this call removed
            If ErrorNumber <> 0 Then
                SetProtection True
                Application.ScreenUpdating = mScreenBeforeWrite
                mScreenHeld = False
            End If
            On Error GoTo 0
            If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription
        End If
    'Count the open window
        mWriteDepth = mWriteDepth + 1

End Sub

Public Sub EndWrite()

'
'------------------------------------------------------------------------------
'                             CLOSE A WRITE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Closes a write window and, with the outermost one, protects the sheets and the
'   workbook structure again.
'
' WHY THIS EXISTS
'   A command whose body raised calls it before it raises, so it never raises itself: a
'   host that refuses the protection leaves the workbook unprotected, which the Excel
'   check reports.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Subtracts 1 from the depth of open windows when it is above 0. When the depth is
'   then 0, protects through SetProtection and gives screen updating back the value it
'   had before the window opened.
'
' ERROR POLICY
'   Never raises: the errors of SetProtection are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   SetProtection
'
' NOTES
'   Screen updating is touched only when an outermost window held it off.
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLOSE WINDOW
'------------------------------------------------------------------------------
    'Count the closed window
        If mWriteDepth > 0 Then mWriteDepth = mWriteDepth - 1
    'With the outermost window, protect again and give the screen back, ignoring errors
        If mWriteDepth = 0 Then
            On Error Resume Next
            SetProtection True
            If mScreenHeld Then
                Application.ScreenUpdating = mScreenBeforeWrite
                mScreenHeld = False
            End If
            On Error GoTo 0
        End If

End Sub

Public Sub CloseWriteWindows()

'
'------------------------------------------------------------------------------
'                           CLOSE ALL WRITE WINDOWS
'------------------------------------------------------------------------------
' PURPOSE
'   Closes every open write window at once and protects the sheets again.
'
' WHY THIS EXISTS
'   A caller that caught an error between a BeginWrite and its EndWrite needs the
'   protection back.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Sets the depth of open windows to 1 and calls EndWrite.
'
' ERROR POLICY
'   As EndWrite: never raises.
'
' DEPENDENCIES
'   EndWrite
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLOSE WINDOWS
'------------------------------------------------------------------------------
    'Collapse the open windows to one, then close it
        mWriteDepth = 1
        EndWrite

End Sub

Private Sub SetProtection(ByVal turnOn As Boolean)

'
'------------------------------------------------------------------------------
'                      PROTECT OR UNPROTECT THE WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Protects or unprotects every managed sheet and the workbook structure, all without a
'   password.
'
' WHY THIS EXISTS
'   BeginWrite and EndWrite open and close the write windows through it.
'
' INPUTS
'   turnOn
'     True to protect, False to unprotect
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   For each sheet of ManagedSheets that exists:
'     - unprotects it when protected
'     - when turning on, unlocks its input cells and protects it
'     - records the sheet when an error occurred or, when turning off, when it is
'       still protected
'   Then does the same for the workbook structure and raises when anything was recorded.
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP, with the source "CORE_Util", naming the sheets, and
'   the workbook structure, that have a password of their own. Errors of the host calls
'   are collected, not raised one by one.
'
' DEPENDENCIES
'   ManagedSheets (CORE_Config), SheetExists, GetSheet, UnlockInputCells, ProtectSheet
'
' NOTES
'   The message asks to remove the password (Review > Unprotect Sheet or Protect
'   Workbook) and run again: the workbook itself needs no password.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.SetProtection"

    Dim SheetNames              As Variant  'Names of the managed sheets
    Dim SheetIndex              As Long     'Sheet being processed
    Dim TargetSheet             As Object   'Sheet being processed
    Dim BlockedNames            As String   'Names of the sheets that refused, separated by commas
    Dim ErrorText               As String   'Message for the sheets and the structure that refused

'------------------------------------------------------------------------------
' LIST SHEETS
'------------------------------------------------------------------------------
    'Get the managed sheets
        SheetNames = ManagedSheets()

'------------------------------------------------------------------------------
' PROCESS SHEETS
'------------------------------------------------------------------------------
    'Unprotect each sheet, and protect it again when turning on; record a sheet that refuses
        For SheetIndex = LBound(SheetNames) To UBound(SheetNames)
            If SheetExists(CStr(SheetNames(SheetIndex))) Then
                Set TargetSheet = GetSheet(CStr(SheetNames(SheetIndex)))
                On Error Resume Next
                Err.Clear
                ' An empty password never prompts
                If TargetSheet.ProtectContents Then TargetSheet.Unprotect ""
                If turnOn And Err.Number = 0 Then
                    UnlockInputCells TargetSheet
                    ProtectSheet TargetSheet
                End If
                If Err.Number <> 0 Then
                    BlockedNames = BlockedNames & IIf(BlockedNames = vbNullString, "", ", ") & TargetSheet.Name
                ElseIf Not turnOn Then
                    If TargetSheet.ProtectContents Then
                        BlockedNames = BlockedNames & IIf(BlockedNames = vbNullString, "", ", ") & TargetSheet.Name
                    End If
                End If
                On Error GoTo 0
            End If
        Next SheetIndex

'------------------------------------------------------------------------------
' PROCESS WORKBOOK STRUCTURE
'------------------------------------------------------------------------------
    'Unprotect the workbook structure, and protect it again when turning on
        On Error Resume Next
        Err.Clear
        If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect ""
        If turnOn And Err.Number = 0 Then ThisWorkbook.Protect Structure:=True, Windows:=False
    'Record a structure that refuses
        If Err.Number <> 0 Then
            ErrorText = "The workbook structure has a password of its own. "
        ElseIf Not turnOn Then
            If ThisWorkbook.ProtectStructure Then
                ErrorText = "The workbook structure has a password of its own. "
            End If
        End If
        On Error GoTo 0

'------------------------------------------------------------------------------
' RAISE FOR REFUSALS
'------------------------------------------------------------------------------
    'Name the sheets that refuse
        If BlockedNames <> vbNullString Then
            ErrorText = ErrorText & "The sheet(s) " & BlockedNames & " have a password of their own. "
        End If
    'Raise one error for everything that refused
        If ErrorText <> vbNullString Then
            Err.Raise vbObjectError + ERR_APP, "CORE_Util", ErrorText & _
                "Remove it (Review > Unprotect Sheet or Protect Workbook) and run again: the workbook itself needs no password."
        End If

End Sub

Private Sub UnlockInputCells(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                            UNLOCK THE INPUT CELLS
'------------------------------------------------------------------------------
' PURPOSE
'   Unlocks the cells in the input fill on the sheets that hold input cells, once,
'   before the first protection.
'
' WHY THIS EXISTS
'   A workbook that the sync updated keeps its own cell formats, and one made before the
'   protection has every cell locked.
'
' INPUTS
'   ws
'     Sheet to unlock
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Does nothing for a sheet that is not Control, In_* or P_*. Reads the lock state of
'   the used range and leaves a sheet with an unlocked cell already alone, so a built
'   workbook pays one property read. Otherwise unlocks each cell of the sheet whose fill
'   is COLOR_INPUT_CELL, covering the pre-formatted input rows of an In_* sheet.
'
' ERROR POLICY
'   Not handled here; the caller (SetProtection) runs it with errors suppressed.
'
' DEPENDENCIES
'   SH_CONTROL, LAYOUT_FIRST_DATA_ROW, INPUT_ROWS, COLOR_INPUT_CELL (CORE_Config)
'
' NOTES
'   The pre-formatted input rows are empty, which a host may leave out of the used
'   range.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.UnlockInputCells"

    Dim InputCell               As Object   'Cell being tested
    Dim LockState               As Variant  'Lock state of the used range; Null when mixed
    Dim LastRow                 As Long     'Last row to test
    Dim LastColumn              As Long     'Last column to test

'------------------------------------------------------------------------------
' SELECT SHEETS
'------------------------------------------------------------------------------
    'Leave the sheets that hold no input cells alone
        If Not (ws.Name = SH_CONTROL Or Left$(ws.Name, 3) = "In_" Or Left$(ws.Name, 2) = "P_") Then Exit Sub
    'Leave a sheet alone when some cells are unlocked already (Null) or none are locked
        LockState = ws.UsedRange.Locked
        If VarType(LockState) <> vbBoolean Then Exit Sub
        If Not LockState Then Exit Sub

'------------------------------------------------------------------------------
' FIND RANGE
'------------------------------------------------------------------------------
    'Find the last row and column of the used range
        LastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        LastColumn = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    'Include the pre-formatted input rows, which a host may leave out of the used range
        If Left$(ws.Name, 3) = "In_" And LastRow < LAYOUT_FIRST_DATA_ROW + INPUT_ROWS - 1 Then
            LastRow = LAYOUT_FIRST_DATA_ROW + INPUT_ROWS - 1
        End If

'------------------------------------------------------------------------------
' UNLOCK CELLS
'------------------------------------------------------------------------------
    'Unlock each cell in the input fill
        For Each InputCell In ws.Range(ws.Cells(1, 1), ws.Cells(LastRow, LastColumn)).Cells
            If InputCell.Interior.Color = COLOR_INPUT_CELL Then InputCell.Locked = False
        Next InputCell

End Sub

Private Sub ProtectSheet(ByVal ws As Object)

'
'------------------------------------------------------------------------------
'                               PROTECT A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Protects a sheet without a password: the input sheets accept inserted rows, sorting
'   and the autofilter on their unlocked cells; every other sheet is protected
'   completely.
'
' INPUTS
'   ws
'     Sheet to protect
'
' ERROR POLICY
'   Not handled here; the caller (SetProtection) runs it with errors suppressed.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
    'Protect the sheet, allowing rows, sorting and filtering on the In_ sheets
        If Left$(ws.Name, 3) = "In_" Then
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
        Else
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If

End Sub
