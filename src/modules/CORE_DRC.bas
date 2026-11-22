Attribute VB_Name = "CORE_DRC"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_DRC
'------------------------------------------------------------------------------
' PURPOSE
'   Calculates the default risk capital (DRC) for non-securitizations: gross
'   jump-to-default, maturity scaling, obligor netting by seniority, hedge benefit
'   ratio and the risk-weighted bucket charge, summed over the buckets.
'
' WHY THIS EXISTS
'   The DRC is a capital component of its own, next to the sensitivities-based
'   method and the residual risk add-on, so it has its own engine module.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Reads the positions of In_DRC, writes the obligor and bucket tables to Out_DRC,
'   logs the rows it rejects and keeps the total of the last run.
'
' ERROR POLICY
'   Invalid rows are logged as ERROR and skipped. An obligor split across buckets is
'   logged as ERROR and skipped; one with several credit qualities takes the
'   worst and is logged as WARN.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, CORE_Params (the parameters must be loaded).
'
' NOTES
'   Reference: Basel Framework MAR22.10-MAR22.26.
'
'   The module also runs in LibreOffice Calc with VBA support for the automated
'   tests. It needs no references beyond the defaults.
'
' PUBLIC SURFACE
'   RunDRC; DrcTotal.
'
' STATE OWNERSHIP
'   Owns the total of the last run. Reads In_DRC; writes Out_DRC and logs rejected
'   rows.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DRC_BUCKET_COL    As Long = 8 'Bucket table: 8th table column (J)

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mTotal  As Double   'DRC of the last run

Public Function DrcTotal() As Double

'
'------------------------------------------------------------------------------
'                          RETURN THE LAST DRC TOTAL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DRC total of the last RunDRC call.
'
' WHY THIS EXISTS
'   The total stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The total, or 0 when RunDRC has not run.
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
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the total of the last run
        DrcTotal = mTotal

End Function

Public Sub RunDRC()

'
'------------------------------------------------------------------------------
'                              CALCULATE THE DRC
'------------------------------------------------------------------------------
' PURPOSE
'   Calculates the DRC of the positions of In_DRC, writes the obligor and bucket tables
'   to Out_DRC and stores the total.
'
' WHY THIS EXISTS
'   The run needs the charge for the capital total, and a position that cannot be
'   used must show in the Log and not stop the run.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The total is read with DrcTotal.
'
' BEHAVIOR
'   In order:
'     - counts the rows of In_DRC
'     - validates each position and computes its gross jump-to-default scaled by
'       maturity (MAR22.11-MAR22.18)
'     - groups the positions by obligor: one bucket each (MAR22.22), the worst credit
'       quality when positions differ (MAR22.24)
'     - nets long and short per obligor by seniority (MAR22.19-MAR22.21)
'     - sums the obligors into buckets, applies the hedge benefit ratio and adds the
'       bucket charges (MAR22.22-MAR22.26)
'
' ERROR POLICY
'   Does not raise for a bad row. Errors from the callees (for example a missing
'   sheet) are not handled here.
'
' DEPENDENCIES
'   CORE_Util (GetSheet, TableCell, LogMessage and the other helpers), CORE_Params (the
'   Drc* accessors), WriteDrcOutput, WriteBucketHeader, FormatBucketRows, MaxL.
'
' NOTES
'   The parameters must be loaded before the call.
'
'   Ranks run from the most junior (1) to the most senior; unused shorts carry
'   upward.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_DRC.RunDRC"

    'Input rows
    Dim InputSheet              As Object        'In_DRC sheet
    Dim RowIndex                As Long          'Row being read
    Dim Source                  As String        'Row reference for the Log
    Dim RowIsValid              As Boolean       'True while the row is usable
    Dim Found                   As Boolean       'True when a lookup found its key
    Dim PositionCount           As Long          'Rows, then valid positions
    Dim Index                   As Long          'Loop counter
    Dim KeptCount               As Long          'Positions or obligors kept
    Dim ObligorIndex            As Long          'Obligor being read
    Dim ObligorCount            As Long          'Number of obligors
    Dim RankIndex               As Long          'Seniority rank being read
    Dim MaxRank                 As Long          'Highest seniority rank
    Dim Obligor                 As String        'Obligor of the row
    Dim Bucket                  As String        'Bucket of the row
    Dim Quality                 As String        'Credit quality of the row
    Dim Seniority               As String        'Seniority of the row
    Dim Notional                As Double        'Signed notional
    Dim MarketValue             As Double        'Signed market value
    Dim Maturity                As Double        'Maturity in years
    Dim Lgd                     As Double        'Loss given default
    Dim Jtd                     As Double        'Jump-to-default amount
    Dim MaturityWeight          As Double        'Maturity scaling factor

    'Positions after validation
    Dim PosObligor()            As String        'Obligor of each position
    Dim PosBucket()             As String        'Bucket of each position
    Dim PosQuality()            As String        'Credit quality of each position
    Dim PosRank()               As Long          'Seniority rank of each position
    Dim PosJtd()                As Double        'Scaled jump-to-default of each position

    'Obligors
    Dim ObligorName()           As String        'Obligor name
    Dim ObligorBucket()         As String        'Bucket of the obligor
    Dim ObligorQuality()        As String        'Worst credit quality of the obligor
    Dim ObligorRW()             As Double        'Risk weight of the obligor
    Dim ObligorLong()           As Double        'Net long jump-to-default
    Dim ObligorShort()          As Double        'Net short jump-to-default
    Dim ObligorKeys             As Collection    'Index of the obligors by name
    Dim ObligorSplit()          As Boolean       'True when positions sit in several buckets
    Dim ObligorMixed()          As Boolean       'True when credit qualities differ
    Dim LongByRank()            As Double        'Longs of one obligor by rank
    Dim ShortByRank()           As Double        'Shorts of one obligor by rank
    Dim Carry                   As Double        'Shorts carried up the ranks
    Dim Offset                  As Double        'Long offset by the carried shorts

    'Buckets
    Dim BucketCount             As Long          'Number of buckets
    Dim BucketLong()            As Double        'Net long of each bucket
    Dim BucketShort()           As Double        'Net short of each bucket
    Dim BucketWeightedLong()    As Double        'Risk-weighted long of each bucket
    Dim BucketWeightedShort()   As Double        'Risk-weighted short of each bucket
    Dim Hbr                     As Double        'Hedge benefit ratio
    Dim BucketCharge            As Double        'Charge of one bucket
    Dim BucketSheet             As Object        'Out_DRC sheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the total at zero
        mTotal = 0#
    'Locate the input sheet
        Set InputSheet = GetSheet(SH_IN_DRC)
    'Read the highest seniority rank
        MaxRank = DrcMaxRank()

'------------------------------------------------------------------------------
' COUNT ROWS
'------------------------------------------------------------------------------
    'Count the rows up to the first row whose first two columns are blank
        PositionCount = 0
        RowIndex = LAYOUT_FIRST_DATA_ROW
        Do While Not (IsBlank(TableCell(InputSheet, RowIndex, 1).Value) And _
            IsBlank(TableCell(InputSheet, RowIndex, 2).Value))
            PositionCount = PositionCount + 1
            RowIndex = RowIndex + 1
        Loop
        ReDim PosObligor(1 To MaxL(PositionCount, 1))
        ReDim PosBucket(1 To MaxL(PositionCount, 1))
        ReDim PosQuality(1 To MaxL(PositionCount, 1))
        ReDim PosRank(1 To MaxL(PositionCount, 1))
        ReDim PosJtd(1 To MaxL(PositionCount, 1))

'------------------------------------------------------------------------------
' VALIDATE POSITIONS
'------------------------------------------------------------------------------
    'Gross JTD per position (MAR22.11-MAR22.13) and maturity scaling (MAR22.15-MAR22.18)
        KeptCount = 0
        For RowIndex = LAYOUT_FIRST_DATA_ROW To LAYOUT_FIRST_DATA_ROW + PositionCount - 1
            Source = SH_IN_DRC & " row " & RowIndex
            Obligor = ToText(TableCell(InputSheet, RowIndex, 2).Value)
            Bucket = ToText(TableCell(InputSheet, RowIndex, 3).Value)
            Quality = ToText(TableCell(InputSheet, RowIndex, 4).Value)
            Seniority = ToText(TableCell(InputSheet, RowIndex, 5).Value)
            RowIsValid = True
            If Obligor = vbNullString Then
                LogMessage "ERROR", Source, "Obligor is required; row skipped"
                RowIsValid = False
            ElseIf Not IsValidDrcBucket(Bucket) Then
                LogMessage "ERROR", Source, "Invalid DRC bucket '" & Bucket & "'; row skipped"
                RowIsValid = False
            ElseIf DrcSeniorityRank(Seniority) = 0 Then
                LogMessage "ERROR", Source, "Unknown seniority '" & Seniority & "'; row skipped"
                RowIsValid = False
            ElseIf Not IsNumber(TableCell(InputSheet, RowIndex, 6).Value) Or _
                Not IsNumber(TableCell(InputSheet, RowIndex, 7).Value) Then
                LogMessage "ERROR", Source, "Notional and market value must be numeric; row skipped"
                RowIsValid = False
            ElseIf Not IsBlankOrNumber(TableCell(InputSheet, RowIndex, 8).Value) Then
                ' Blank means one year; text must not be read as zero and floored to three months
                LogMessage "ERROR", Source, "Maturity is not numeric; row skipped"
                RowIsValid = False
            ElseIf Not IsBlankOrNumber(TableCell(InputSheet, RowIndex, 9).Value) Then
                LogMessage "ERROR", Source, "LGD override is not numeric; row skipped"
                RowIsValid = False
            ElseIf Seniority = "EQUITY" And Not IsAllowedEquityMaturity(TableCell(InputSheet, RowIndex, 8).Value) Then
                ' MAR22.16 and its FAQ: cash equity is three months or at least one year
                LogMessage "ERROR", Source, "Cash equity maturity must be three months or at least one year (MAR22.16); row skipped"
                RowIsValid = False
            Else
                DrcRiskWeight Quality, Found
                If Not Found Then
                    LogMessage "ERROR", Source, "Unknown credit quality '" & Quality & "'; row skipped"
                    RowIsValid = False
                End If
            End If
            If RowIsValid Then
                Notional = ToDouble(TableCell(InputSheet, RowIndex, 6).Value, 0#)
                MarketValue = ToDouble(TableCell(InputSheet, RowIndex, 7).Value, 0#)
                Maturity = ToDouble(TableCell(InputSheet, RowIndex, 8).Value, DrcMaxMaturity())
                If IsNumber(TableCell(InputSheet, RowIndex, 9).Value) Then
                    Lgd = ToDouble(TableCell(InputSheet, RowIndex, 9).Value, 0#)
                Else
                    Lgd = DrcLGD(Seniority, Found)
                End If
                ' JTD = LGD x notional + P&L, with P&L = market value - notional (signed amounts)
                Jtd = Lgd * Notional + (MarketValue - Notional)
                If Notional >= 0# Then
                    Jtd = MaxD(Jtd, 0#)
                Else
                    Jtd = MinD(Jtd, 0#)
                End If
                MaturityWeight = MinD(MaxD(Maturity, DrcMinMaturity()), DrcMaxMaturity()) / DrcMaxMaturity()
                KeptCount = KeptCount + 1
                PosObligor(KeptCount) = Obligor
                PosBucket(KeptCount) = Bucket
                PosQuality(KeptCount) = Quality
                PosRank(KeptCount) = DrcSeniorityRank(Seniority)
                PosJtd(KeptCount) = Jtd * MaturityWeight
            End If
        Next RowIndex
        PositionCount = KeptCount

'------------------------------------------------------------------------------
' GROUP BY OBLIGOR
'------------------------------------------------------------------------------
    'One bucket per obligor (MAR22.22), and the worst credit quality, the one with the
    'highest risk weight, when positions differ (MAR22.24)
        Set ObligorKeys = New Collection
        ObligorCount = 0
        ReDim ObligorName(1 To MaxL(PositionCount, 1))
        ReDim ObligorBucket(1 To MaxL(PositionCount, 1))
        ReDim ObligorQuality(1 To MaxL(PositionCount, 1))
        ReDim ObligorSplit(1 To MaxL(PositionCount, 1))
        ReDim ObligorMixed(1 To MaxL(PositionCount, 1))
        For Index = 1 To PositionCount
            ObligorIndex = KeyIndex(ObligorKeys, PosObligor(Index))
            If ObligorIndex = 0 Then
                ObligorCount = ObligorCount + 1
                ObligorName(ObligorCount) = PosObligor(Index)
                ObligorBucket(ObligorCount) = PosBucket(Index)
                ObligorQuality(ObligorCount) = PosQuality(Index)
                AddKey ObligorKeys, PosObligor(Index), ObligorCount
            Else
                If ObligorBucket(ObligorIndex) <> PosBucket(Index) Then ObligorSplit(ObligorIndex) = True
                If ObligorQuality(ObligorIndex) <> PosQuality(Index) Then
                    ObligorMixed(ObligorIndex) = True
                    If DrcRiskWeight(PosQuality(Index), Found) > _
                        DrcRiskWeight(ObligorQuality(ObligorIndex), Found) Then
                        ObligorQuality(ObligorIndex) = PosQuality(Index)
                    End If
                End If
            End If
        Next Index
    'Drop split obligors; their positions match no remaining obligor below
        KeptCount = 0
        For ObligorIndex = 1 To ObligorCount
            If ObligorSplit(ObligorIndex) Then
                LogMessage "ERROR", SH_IN_DRC, "Obligor " & ObligorName(ObligorIndex) & _
                    " has positions in several buckets (MAR22.22); its positions are skipped", , _
                    ObligorName(ObligorIndex)
            Else
                If ObligorMixed(ObligorIndex) Then
                    LogMessage "WARN", SH_IN_DRC, "Obligor " & ObligorName(ObligorIndex) & _
                        " has several credit qualities; the worst, " & ObligorQuality(ObligorIndex) & ", is used", , _
                        ObligorName(ObligorIndex)
                End If
                KeptCount = KeptCount + 1
                ObligorName(KeptCount) = ObligorName(ObligorIndex)
                ObligorBucket(KeptCount) = ObligorBucket(ObligorIndex)
                ObligorQuality(KeptCount) = ObligorQuality(ObligorIndex)
            End If
        Next ObligorIndex
        ObligorCount = KeptCount

'------------------------------------------------------------------------------
' NET LONG AND SHORT
'------------------------------------------------------------------------------
    'Net long/short per obligor (MAR22.19-22.21): a short offsets a long only when it ranks
    'the same as, or junior to, the long. Ranks run from the most junior (1) to the most
    'senior; unused shorts carry upward.
        ReDim ObligorRW(1 To MaxL(ObligorCount, 1))
        ReDim ObligorLong(1 To MaxL(ObligorCount, 1))
        ReDim ObligorShort(1 To MaxL(ObligorCount, 1))
        For ObligorIndex = 1 To ObligorCount
            ReDim LongByRank(1 To MaxRank)
            ReDim ShortByRank(1 To MaxRank)
            For Index = 1 To PositionCount
                If PosObligor(Index) = ObligorName(ObligorIndex) Then
                    If PosJtd(Index) >= 0# Then
                        LongByRank(PosRank(Index)) = LongByRank(PosRank(Index)) + PosJtd(Index)
                    Else
                        ShortByRank(PosRank(Index)) = ShortByRank(PosRank(Index)) - PosJtd(Index)
                    End If
                End If
            Next Index
            Carry = 0#
            For RankIndex = 1 To MaxRank
                Carry = Carry + ShortByRank(RankIndex)
                Offset = MinD(LongByRank(RankIndex), Carry)
                ObligorLong(ObligorIndex) = ObligorLong(ObligorIndex) + LongByRank(RankIndex) - Offset
                Carry = Carry - Offset
            Next RankIndex
            ObligorShort(ObligorIndex) = Carry
            ObligorRW(ObligorIndex) = DrcRiskWeight(ObligorQuality(ObligorIndex), Found)
        Next ObligorIndex

'------------------------------------------------------------------------------
' AGGREGATE THE BUCKETS
'------------------------------------------------------------------------------
    'Sum the obligors into buckets (MAR22.22-22.26)
        BucketCount = DrcBucketCount()
        ReDim BucketLong(1 To BucketCount)
        ReDim BucketShort(1 To BucketCount)
        ReDim BucketWeightedLong(1 To BucketCount)
        ReDim BucketWeightedShort(1 To BucketCount)
        For ObligorIndex = 1 To ObligorCount
            Index = DrcBucketIndex(ObligorBucket(ObligorIndex))
            BucketLong(Index) = BucketLong(Index) + ObligorLong(ObligorIndex)
            BucketShort(Index) = BucketShort(Index) + ObligorShort(ObligorIndex)
            BucketWeightedLong(Index) = BucketWeightedLong(Index) + _
                ObligorRW(ObligorIndex) * ObligorLong(ObligorIndex)
            BucketWeightedShort(Index) = BucketWeightedShort(Index) + _
                ObligorRW(ObligorIndex) * ObligorShort(ObligorIndex)
        Next ObligorIndex

'------------------------------------------------------------------------------
' WRITE THE OUTPUT
'------------------------------------------------------------------------------
    'Write the obligor table, then the bucket table with its charge
        WriteDrcOutput ObligorCount, ObligorName, ObligorBucket, ObligorQuality, ObligorRW, ObligorLong, ObligorShort
        Set BucketSheet = GetSheet(SH_OUT_DRC)
        WriteBucketHeader BucketSheet
        For Index = 1 To BucketCount
            If BucketLong(Index) + BucketShort(Index) > 0# Then
                Hbr = BucketLong(Index) / (BucketLong(Index) + BucketShort(Index))
            Else
                Hbr = 0#
            End If
            BucketCharge = MaxD(BucketWeightedLong(Index) - Hbr * BucketWeightedShort(Index), 0#)
            mTotal = mTotal + BucketCharge
            TableCell(BucketSheet, LAYOUT_HEADER_ROW + Index, DRC_BUCKET_COL + 0).Value = DrcBucketName(Index)
            TableCell(BucketSheet, LAYOUT_HEADER_ROW + Index, DRC_BUCKET_COL + 1).Value = BucketLong(Index)
            TableCell(BucketSheet, LAYOUT_HEADER_ROW + Index, DRC_BUCKET_COL + 2).Value = BucketShort(Index)
            TableCell(BucketSheet, LAYOUT_HEADER_ROW + Index, DRC_BUCKET_COL + 3).Value = Hbr
            TableCell(BucketSheet, LAYOUT_HEADER_ROW + Index, DRC_BUCKET_COL + 4).Value = BucketWeightedLong(Index)
            TableCell(BucketSheet, LAYOUT_HEADER_ROW + Index, DRC_BUCKET_COL + 5).Value = BucketWeightedShort(Index)
            TableCell(BucketSheet, LAYOUT_HEADER_ROW + Index, DRC_BUCKET_COL + 6).Value = BucketCharge
        Next Index
        FormatBucketRows BucketSheet, BucketCount

End Sub

Private Sub WriteDrcOutput(ByVal nObl As Long, ByRef oName() As String, ByRef oBucket() As String, _
                           ByRef oQuality() As String, ByRef oRW() As Double, _
                           ByRef oLong() As Double, ByRef oShort() As Double)
'
'------------------------------------------------------------------------------
'                           WRITE THE OBLIGOR TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the obligor table of Out_DRC: obligor, bucket, credit quality, risk weight, net
'   long and net short.
'
' WHY THIS EXISTS
'   The obligor table shows how the positions net before the buckets add them.
'
' INPUTS
'   nObl
'     Number of obligors to write
'
'   oName, oBucket, oQuality, oRW, oLong, oShort
'     The obligor arrays, indexed 1 to nObl
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the table area of Out_DRC, writes the header and one row per obligor, then
'   formats the rows and shows the risk weight as a percentage.
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   GetSheet, ClearTableArea, WriteTableHeader, TableCell, FormatDataRows (CORE_Util).
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
    Const PROC_NAME             As String = "CORE_DRC.WriteDrcOutput"

    Dim OutputSheet             As Object        'Out_DRC sheet
    Dim ObligorIndex            As Long          'Obligor being written

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Locate the output sheet
        Set OutputSheet = GetSheet(SH_OUT_DRC)
    'Clear the table area
        ClearTableArea OutputSheet
    'Write the header
        WriteTableHeader OutputSheet, LAYOUT_HEADER_ROW, _
            Array("Obligor", "Bucket", "CreditQuality", "RW", "NetLong", "NetShort"), COLOR_OUTPUT_HEADER

'------------------------------------------------------------------------------
' WRITE ROWS
'------------------------------------------------------------------------------
    'Write one row per obligor
        For ObligorIndex = 1 To nObl
            TableCell(OutputSheet, LAYOUT_HEADER_ROW + ObligorIndex, 1).Value = oName(ObligorIndex)
            TableCell(OutputSheet, LAYOUT_HEADER_ROW + ObligorIndex, 2).Value = oBucket(ObligorIndex)
            TableCell(OutputSheet, LAYOUT_HEADER_ROW + ObligorIndex, 3).Value = oQuality(ObligorIndex)
            TableCell(OutputSheet, LAYOUT_HEADER_ROW + ObligorIndex, 4).Value = oRW(ObligorIndex)
            TableCell(OutputSheet, LAYOUT_HEADER_ROW + ObligorIndex, 5).Value = oLong(ObligorIndex)
            TableCell(OutputSheet, LAYOUT_HEADER_ROW + ObligorIndex, 6).Value = oShort(ObligorIndex)
        Next ObligorIndex

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Format the data rows
        FormatDataRows OutputSheet, LAYOUT_FIRST_DATA_ROW, LAYOUT_HEADER_ROW + nObl, 6, 5, "#,##0", False
    'Show the risk weight as a percentage; ignore formatting errors
        On Error Resume Next
        If nObl > 0 Then
            OutputSheet.Range(TableCell(OutputSheet, LAYOUT_FIRST_DATA_ROW, 4), _
                TableCell(OutputSheet, LAYOUT_HEADER_ROW + nObl, 4)).NumberFormat = "0.0%"
        End If
        On Error GoTo 0

End Sub

Private Sub WriteBucketHeader(ByVal ws As Object)
'
'------------------------------------------------------------------------------
'                        WRITE THE BUCKET TABLE HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Writes and formats the header of the bucket table of Out_DRC.
'
' WHY THIS EXISTS
'   The bucket table sits to the right of the obligor table, from column
'   DRC_BUCKET_COL of the table area (column J), and needs its own header.
'
' INPUTS
'   ws
'     Out_DRC sheet
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the seven column captions, then applies the header font, fill, alignment and
'   borders.
'
' ERROR POLICY
'   Formatting errors are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   TableCell, HeaderCaption, FormatBorders (CORE_Util).
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
    Const PROC_NAME             As String = "CORE_DRC.WriteBucketHeader"

    Dim Headers                 As Variant       'Column captions
    Dim ColumnIndex             As Long          'Column being written

'------------------------------------------------------------------------------
' WRITE CAPTIONS
'------------------------------------------------------------------------------
    'Write the caption of each of the seven columns
        Headers = Array("Bucket", "NetLong", "NetShort", "HBR", "WeightedLong", "WeightedShort", "DRC")
        For ColumnIndex = 0 To 6
            TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL + ColumnIndex).Value = HeaderCaption(Headers(ColumnIndex))
        Next ColumnIndex

'------------------------------------------------------------------------------
' FORMAT HEADER
'------------------------------------------------------------------------------
    'Format the header cells; ignore formatting errors, then draw the borders
        On Error Resume Next
        With ws.Range(TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL), _
            TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL + 6))
            .Font.Bold = True
            .Font.Size = HEADER_FONT_SIZE
            .Font.Color = COLOR_WHITE
            .Interior.Color = COLOR_OUTPUT_HEADER
            .VerticalAlignment = -4108
            .HorizontalAlignment = -4108
            .WrapText = True
        End With
        On Error GoTo 0
        FormatBorders ws.Range(TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL), _
            TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL + 6)), COLOR_GRID, 1

End Sub

Private Sub FormatBucketRows(ByVal ws As Object, ByVal nB As Long)
'
'------------------------------------------------------------------------------
'                            FORMAT THE BUCKET ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Formats the data rows of the bucket table of Out_DRC.
'
' WHY THIS EXISTS
'   The bucket rows need their own number formats next to the obligor table.
'
' INPUTS
'   ws
'     Out_DRC sheet
'
'   nB
'     Number of bucket rows
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
    'Do nothing when there are no bucket rows; format the others and ignore formatting errors
        If nB < 1 Then Exit Sub
        On Error Resume Next
        With ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, DRC_BUCKET_COL), _
            TableCell(ws, LAYOUT_HEADER_ROW + nB, DRC_BUCKET_COL + 6))
            .Font.Size = 10
            .RowHeight = 20.1
            .NumberFormat = "#,##0"
        End With
        FormatBorders ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, DRC_BUCKET_COL), _
            TableCell(ws, LAYOUT_HEADER_ROW + nB, DRC_BUCKET_COL + 6)), COLOR_GRID, 1
        ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, DRC_BUCKET_COL + 3), _
            TableCell(ws, LAYOUT_HEADER_ROW + nB, DRC_BUCKET_COL + 3)).NumberFormat = "0.0000"
        ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, DRC_BUCKET_COL), _
            TableCell(ws, LAYOUT_HEADER_ROW + nB, DRC_BUCKET_COL)).NumberFormat = "@"
        On Error GoTo 0

End Sub

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
'
'------------------------------------------------------------------------------
'                            RETURN THE LARGER LONG
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the larger of two Long values.
'
' INPUTS
'   a, b
'     The values to compare
'
' RETURNS
'   The larger value.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the larger value
        If a > b Then
            MaxL = a
        Else
            MaxL = b
        End If

End Function

Private Function IsAllowedEquityMaturity(ByVal v As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                      TEST THE MATURITY OF A CASH EQUITY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the maturity of a cash equity position is allowed: blank (one
'   year), the three-month floor or at least one year; nothing in between (MAR22.16).
'
' INPUTS
'   v
'     Maturity cell value
'
' RETURNS
'   True when the maturity is allowed.
'
' NOTES
'   The three-month floor is compared with a tolerance of 0.000001.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_DRC.IsAllowedEquityMaturity"

    Dim MaturityYears           As Double        'Maturity in years

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
    'A blank maturity means one year, which is allowed
        If IsBlank(v) Then
            IsAllowedEquityMaturity = True
            Exit Function
        End If
    'Allow the three-month floor or at least one year
        MaturityYears = ToDouble(v, DrcMaxMaturity())
        IsAllowedEquityMaturity = (Abs(MaturityYears - DrcMinMaturity()) < 0.000001) Or _
            (MaturityYears >= DrcMaxMaturity())

End Function
