Attribute VB_Name = "CORE_DRC"
'==============================================================================
' MODULE: CORE_DRC
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the default risk capital for non-securitizations: gross
'   jump-to-default, maturity scaling, obligor netting by seniority, hedge
'   benefit ratio and risk-weighted bucket charge, summed over buckets.
'
' PUBLIC SURFACE
'   RunDRC; DrcTotal.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, CORE_Params (parameters must be loaded).
'
' STATE OWNERSHIP
'   Owns the total of the last run. Reads In_DRC; writes Out_DRC and logs
'   rejected rows.
'
' ERROR POLICY
'   Invalid rows are logged as ERROR and skipped. An obligor split across
'   buckets is logged as ERROR and skipped; one with several credit
'   qualities takes the worst and is logged as WARN.
'
' REFERENCE
'   Basel Framework MAR22.10-MAR22.26.
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
Private Const DRC_BUCKET_COL As Long = 8                     'Bucket table: 8th table column (J)

'------------------------------------------------------------------------------
'                               MODULE STATE
'------------------------------------------------------------------------------
Private mTotal As Double                                     'DRC of the last run

Public Function DrcTotal() As Double
    DrcTotal = mTotal
End Function

Public Sub RunDRC()
    Dim ws As Object, r As Long, src As String, ok As Boolean, found As Boolean
    Dim n As Long, i As Long, j As Long, o As Long, nObl As Long, rk As Long, maxRank As Long
    Dim obligor As String, bucket As String, quality As String, seniority As String
    Dim notional As Double, mv As Double, maturity As Double, lgd As Double, jtd As Double, w As Double

    ' Positions after validation
    Dim pObl() As String, pBucket() As String, pQuality() As String, pRank() As Long, pJtd() As Double

    ' Obligors
    Dim oName() As String, oBucket() As String, oQuality() As String, oRW() As Double
    Dim oLong() As Double, oShort() As Double, oKeys As Collection
    Dim oSplit() As Boolean, oMixed() As Boolean
    Dim longByRank() As Double, shortByRank() As Double, carry As Double, offset As Double

    ' Buckets
    Dim nB As Long, bLong() As Double, bShort() As Double, bWLong() As Double, bWShort() As Double
    Dim hbr As Double, drcB As Double

    mTotal = 0#
    Set ws = GetSheet(SH_IN_DRC)
    maxRank = DrcMaxRank()

    ' Count rows
    n = 0
    r = LAYOUT_FIRST_DATA_ROW
    Do While Not (IsBlank(TableCell(ws, r, 1).Value) And IsBlank(TableCell(ws, r, 2).Value))
        n = n + 1
        r = r + 1
    Loop
    ReDim pObl(1 To MaxL(n, 1))
    ReDim pBucket(1 To MaxL(n, 1))
    ReDim pQuality(1 To MaxL(n, 1))
    ReDim pRank(1 To MaxL(n, 1))
    ReDim pJtd(1 To MaxL(n, 1))

    ' 1. Gross JTD per position (MAR22.11-MAR22.13) and maturity scaling (MAR22.15-MAR22.18)
    j = 0
    For r = LAYOUT_FIRST_DATA_ROW To LAYOUT_FIRST_DATA_ROW + n - 1
        src = SH_IN_DRC & " row " & r
        obligor = ToText(TableCell(ws, r, 2).Value)
        bucket = ToText(TableCell(ws, r, 3).Value)
        quality = ToText(TableCell(ws, r, 4).Value)
        seniority = ToText(TableCell(ws, r, 5).Value)
        ok = True
        If obligor = "" Then
            LogMessage "ERROR", src, "Obligor is required; row skipped"
            ok = False
        ElseIf Not IsValidDrcBucket(bucket) Then
            LogMessage "ERROR", src, "Invalid DRC bucket '" & bucket & "'; row skipped"
            ok = False
        ElseIf DrcSeniorityRank(seniority) = 0 Then
            LogMessage "ERROR", src, "Unknown seniority '" & seniority & "'; row skipped"
            ok = False
        ElseIf Not IsNumber(TableCell(ws, r, 6).Value) Or Not IsNumber(TableCell(ws, r, 7).Value) Then
            LogMessage "ERROR", src, "Notional and market value must be numeric; row skipped"
            ok = False
        ElseIf Not IsBlankOrNumber(TableCell(ws, r, 8).Value) Then
            ' Blank means one year; text must not be read as zero and floored to three months
            LogMessage "ERROR", src, "Maturity is not numeric; row skipped"
            ok = False
        ElseIf Not IsBlankOrNumber(TableCell(ws, r, 9).Value) Then
            LogMessage "ERROR", src, "LGD override is not numeric; row skipped"
            ok = False
        ElseIf seniority = "EQUITY" And Not IsAllowedEquityMaturity(TableCell(ws, r, 8).Value) Then
            ' MAR22.16 and its FAQ: cash equity is three months or at least one year
            LogMessage "ERROR", src, "Cash equity maturity must be three months or at least one year (MAR22.16); row skipped"
            ok = False
        Else
            DrcRiskWeight quality, found
            If Not found Then
                LogMessage "ERROR", src, "Unknown credit quality '" & quality & "'; row skipped"
                ok = False
            End If
        End If
        If ok Then
            notional = ToDouble(TableCell(ws, r, 6).Value, 0#)
            mv = ToDouble(TableCell(ws, r, 7).Value, 0#)
            maturity = ToDouble(TableCell(ws, r, 8).Value, DrcMaxMaturity())
            If IsNumber(TableCell(ws, r, 9).Value) Then
                lgd = ToDouble(TableCell(ws, r, 9).Value, 0#)
            Else
                lgd = DrcLGD(seniority, found)
            End If
            ' JTD = LGD x notional + P&L, with P&L = market value - notional (signed amounts)
            jtd = lgd * notional + (mv - notional)
            If notional >= 0# Then jtd = MaxD(jtd, 0#) Else jtd = MinD(jtd, 0#)
            w = MinD(MaxD(maturity, DrcMinMaturity()), DrcMaxMaturity()) / DrcMaxMaturity()
            j = j + 1
            pObl(j) = obligor
            pBucket(j) = bucket
            pQuality(j) = quality
            pRank(j) = DrcSeniorityRank(seniority)
            pJtd(j) = jtd * w
        End If
    Next r
    n = j

    ' 2. Obligors: one bucket each (MAR22.22), and the worst credit quality,
    ' the one with the highest risk weight, when positions differ (MAR22.24)
    Set oKeys = New Collection
    nObl = 0
    ReDim oName(1 To MaxL(n, 1))
    ReDim oBucket(1 To MaxL(n, 1))
    ReDim oQuality(1 To MaxL(n, 1))
    ReDim oSplit(1 To MaxL(n, 1))
    ReDim oMixed(1 To MaxL(n, 1))
    For i = 1 To n
        o = KeyIndex(oKeys, pObl(i))
        If o = 0 Then
            nObl = nObl + 1
            oName(nObl) = pObl(i)
            oBucket(nObl) = pBucket(i)
            oQuality(nObl) = pQuality(i)
            AddKey oKeys, pObl(i), nObl
        Else
            If oBucket(o) <> pBucket(i) Then oSplit(o) = True
            If oQuality(o) <> pQuality(i) Then
                oMixed(o) = True
                If DrcRiskWeight(pQuality(i), found) > DrcRiskWeight(oQuality(o), found) Then oQuality(o) = pQuality(i)
            End If
        End If
    Next i

    ' Drop split obligors; their positions match no remaining obligor below
    j = 0
    For o = 1 To nObl
        If oSplit(o) Then
            LogMessage "ERROR", SH_IN_DRC, "Obligor " & oName(o) & " has positions in several buckets (MAR22.22); its positions are skipped"
        Else
            If oMixed(o) Then
                LogMessage "WARN", SH_IN_DRC, "Obligor " & oName(o) & " has several credit qualities; the worst, " & oQuality(o) & ", is used"
            End If
            j = j + 1
            oName(j) = oName(o)
            oBucket(j) = oBucket(o)
            oQuality(j) = oQuality(o)
        End If
    Next o
    nObl = j

    ' 3. Net long/short per obligor (MAR22.19-22.21): a short offsets a long
    ' only when it ranks the same as, or junior to, the long. Ranks run from
    ' the most junior (1) to the most senior; unused shorts carry upward.
    ReDim oRW(1 To MaxL(nObl, 1))
    ReDim oLong(1 To MaxL(nObl, 1))
    ReDim oShort(1 To MaxL(nObl, 1))
    For o = 1 To nObl
        ReDim longByRank(1 To maxRank)
        ReDim shortByRank(1 To maxRank)
        For i = 1 To n
            If pObl(i) = oName(o) Then
                If pJtd(i) >= 0# Then
                    longByRank(pRank(i)) = longByRank(pRank(i)) + pJtd(i)
                Else
                    shortByRank(pRank(i)) = shortByRank(pRank(i)) - pJtd(i)
                End If
            End If
        Next i
        carry = 0#
        For rk = 1 To maxRank
            carry = carry + shortByRank(rk)
            offset = MinD(longByRank(rk), carry)
            oLong(o) = oLong(o) + longByRank(rk) - offset
            carry = carry - offset
        Next rk
        oShort(o) = carry
        oRW(o) = DrcRiskWeight(oQuality(o), found)
    Next o

    ' 4. Bucket charge (MAR22.22-22.26)
    nB = DrcBucketCount()
    ReDim bLong(1 To nB)
    ReDim bShort(1 To nB)
    ReDim bWLong(1 To nB)
    ReDim bWShort(1 To nB)
    For o = 1 To nObl
        i = DrcBucketIndex(oBucket(o))
        bLong(i) = bLong(i) + oLong(o)
        bShort(i) = bShort(i) + oShort(o)
        bWLong(i) = bWLong(i) + oRW(o) * oLong(o)
        bWShort(i) = bWShort(i) + oRW(o) * oShort(o)
    Next o

    WriteDrcOutput nObl, oName, oBucket, oQuality, oRW, oLong, oShort

    Dim wsOut As Object
    Set wsOut = GetSheet(SH_OUT_DRC)
    WriteBucketHeader wsOut
    For i = 1 To nB
        If bLong(i) + bShort(i) > 0# Then
            hbr = bLong(i) / (bLong(i) + bShort(i))
        Else
            hbr = 0#
        End If
        drcB = MaxD(bWLong(i) - hbr * bWShort(i), 0#)
        mTotal = mTotal + drcB
        TableCell(wsOut, LAYOUT_HEADER_ROW + i, DRC_BUCKET_COL + 0).Value = DrcBucketName(i)
        TableCell(wsOut, LAYOUT_HEADER_ROW + i, DRC_BUCKET_COL + 1).Value = bLong(i)
        TableCell(wsOut, LAYOUT_HEADER_ROW + i, DRC_BUCKET_COL + 2).Value = bShort(i)
        TableCell(wsOut, LAYOUT_HEADER_ROW + i, DRC_BUCKET_COL + 3).Value = hbr
        TableCell(wsOut, LAYOUT_HEADER_ROW + i, DRC_BUCKET_COL + 4).Value = bWLong(i)
        TableCell(wsOut, LAYOUT_HEADER_ROW + i, DRC_BUCKET_COL + 5).Value = bWShort(i)
        TableCell(wsOut, LAYOUT_HEADER_ROW + i, DRC_BUCKET_COL + 6).Value = drcB
    Next i
    FormatBucketRows wsOut, nB
End Sub

Private Sub WriteDrcOutput(ByVal nObl As Long, ByRef oName() As String, ByRef oBucket() As String, _
                           ByRef oQuality() As String, ByRef oRW() As Double, _
                           ByRef oLong() As Double, ByRef oShort() As Double)
    Dim ws As Object, o As Long
    Set ws = GetSheet(SH_OUT_DRC)
    ClearTableArea ws
    WriteTableHeader ws, LAYOUT_HEADER_ROW, _
        Array("Obligor", "Bucket", "CreditQuality", "RW", "NetLong", "NetShort"), COLOR_OUTPUT_HEADER
    For o = 1 To nObl
        TableCell(ws, LAYOUT_HEADER_ROW + o, 1).Value = oName(o)
        TableCell(ws, LAYOUT_HEADER_ROW + o, 2).Value = oBucket(o)
        TableCell(ws, LAYOUT_HEADER_ROW + o, 3).Value = oQuality(o)
        TableCell(ws, LAYOUT_HEADER_ROW + o, 4).Value = oRW(o)
        TableCell(ws, LAYOUT_HEADER_ROW + o, 5).Value = oLong(o)
        TableCell(ws, LAYOUT_HEADER_ROW + o, 6).Value = oShort(o)
    Next o
    FormatDataRows ws, LAYOUT_FIRST_DATA_ROW, LAYOUT_HEADER_ROW + nObl, 6, 5, "#,##0", False
    On Error Resume Next
    If nObl > 0 Then
        ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, 4), TableCell(ws, LAYOUT_HEADER_ROW + nObl, 4)).NumberFormat = "0.0%"
    End If
    On Error GoTo 0
End Sub


' The bucket table sits to the right of the obligor table, from column
' DRC_BUCKET_COL of the table area (column J).
Private Sub WriteBucketHeader(ByVal ws As Object)
    Dim headers As Variant, i As Long
    headers = Array("Bucket", "NetLong", "NetShort", "HBR", "WeightedLong", "WeightedShort", "DRC")
    For i = 0 To 6
        TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL + i).Value = HeaderCaption(headers(i))
    Next i
    On Error Resume Next
    With ws.Range(TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL), TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL + 6))
        .Font.Bold = True
        .Font.Color = COLOR_WHITE
        .Interior.Color = COLOR_OUTPUT_HEADER
        .VerticalAlignment = -4108
        .HorizontalAlignment = -4108
        .WrapText = True
    End With
    On Error GoTo 0
    FormatBorders ws.Range(TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL), TableCell(ws, LAYOUT_HEADER_ROW, DRC_BUCKET_COL + 6)), COLOR_GRID, 1
End Sub


Private Sub FormatBucketRows(ByVal ws As Object, ByVal nB As Long)
    If nB < 1 Then Exit Sub
    On Error Resume Next
    With ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, DRC_BUCKET_COL), TableCell(ws, LAYOUT_HEADER_ROW + nB, DRC_BUCKET_COL + 6))
        .Font.Size = 10
        .RowHeight = 20.1
        .NumberFormat = "#,##0"
    End With
    FormatBorders ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, DRC_BUCKET_COL), TableCell(ws, LAYOUT_HEADER_ROW + nB, DRC_BUCKET_COL + 6)), COLOR_GRID, 1
    ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, DRC_BUCKET_COL + 3), TableCell(ws, LAYOUT_HEADER_ROW + nB, DRC_BUCKET_COL + 3)).NumberFormat = "0.0000"
    ws.Range(TableCell(ws, LAYOUT_FIRST_DATA_ROW, DRC_BUCKET_COL), TableCell(ws, LAYOUT_HEADER_ROW + nB, DRC_BUCKET_COL)).NumberFormat = "@"
    On Error GoTo 0
End Sub

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

' MAR22.16: a blank maturity (one year), the three-month floor or at least
' one year; nothing in between.
Private Function IsAllowedEquityMaturity(ByVal v As Variant) As Boolean
    Dim m As Double
    If IsBlank(v) Then
        IsAllowedEquityMaturity = True
        Exit Function
    End If
    m = ToDouble(v, DrcMaxMaturity())
    IsAllowedEquityMaturity = (Abs(m - DrcMinMaturity()) < 0.000001) Or (m >= DrcMaxMaturity())
End Function
