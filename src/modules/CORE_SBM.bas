Attribute VB_Name = "CORE_SBM"
'==============================================================================
' MODULE: CORE_SBM
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the sensitivities-based method (MAR21): delta, vega and
'   curvature capital for GIRR, CSR non-securitization, equity, commodity
'   and FX under the low, medium and high correlation scenarios.
'
'   Flow per component:
'     1. Read and validate the input sheet, net identical risk factors.
'     2. Per risk class and bucket: weighted sensitivities, Kb and Sb.
'     3. Across buckets: aggregate with gamma (alternative Sb when needed).
'
' PUBLIC SURFACE
'   RunSBM; SbmCapital(risk class, component, scenario);
'   SbmScenarioTotal(scenario).
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, CORE_Params (parameters must be loaded).
'
' STATE OWNERSHIP
'   Owns the netted risk factors of the component being processed and the
'   capital of the last run. Reads In_Delta, In_Vega and In_Curvature;
'   writes Out_Buckets and Out_RiskClass and logs rejected rows.
'
' ERROR POLICY
'   Invalid input rows are logged as ERROR and skipped, never counted as
'   zero. Missing parameters raise from CORE_Params.
'
' REFERENCE
'   Basel Framework MAR21.4-MAR21.101.
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
'                               MODULE STATE
'------------------------------------------------------------------------------
' Netted risk factors of the component being processed
Private mN As Long
Private mRC() As String
Private mBucket() As String
Private mQual() As String
Private mSub() As String
Private mT1() As Double      ' delta: tenor; vega: option maturity
Private mT2() As Double      ' vega: GIRR underlying maturity
Private mA1() As Double      ' delta/vega: sensitivity; curvature: CVR up
Private mA2() As Double      ' curvature: CVR down
Private mWS() As Double      ' weighted sensitivity (delta/vega)
Private mKeys As Collection

' Results: risk class x component x scenario
Private mCapital(1 To 5, 1 To 3, 1 To 3) As Double
Private mAltUsed(1 To 5, 1 To 3, 1 To 3) As Boolean
Private mBucketCount(1 To 5, 1 To 3) As Long
Private mOutRow As Long

'------------------------------------------------------------------------------
' Entry point
'------------------------------------------------------------------------------
Public Sub RunSBM()
    Dim i As Long, c As Long, s As Long
    For i = 1 To RC_COUNT
        For c = 1 To COMP_COUNT
            mBucketCount(i, c) = 0
            For s = 1 To SCEN_COUNT
                mCapital(i, c, s) = 0#
                mAltUsed(i, c, s) = False
            Next s
        Next c
    Next i
    ResetBucketOutput

    LoadDelta
    ComputeComponent COMP_DELTA
    LoadVega
    ComputeComponent COMP_VEGA
    LoadCurvature
    ComputeComponent COMP_CURV

    FormatDataRows GetSheet(SH_OUT_BUCKETS), LAYOUT_FIRST_DATA_ROW, mOutRow, 8, 5, "#,##0", False
    WriteRiskClassOutput
End Sub

Public Function SbmCapital(ByVal rcIdx As Long, ByVal comp As Long, ByVal scen As Long) As Double
    SbmCapital = mCapital(rcIdx, comp, scen)
End Function

' Simple sum of delta, vega and curvature across risk classes (MAR21.7).
Public Function SbmScenarioTotal(ByVal scen As Long) As Double
    Dim i As Long, c As Long, total As Double
    total = 0#
    For i = 1 To RC_COUNT
        For c = 1 To COMP_COUNT
            total = total + mCapital(i, c, scen)
        Next c
    Next i
    SbmScenarioTotal = total
End Function

'------------------------------------------------------------------------------
' Risk factor store
'------------------------------------------------------------------------------
Private Sub ResetStore()
    mN = 0
    ReDim mRC(1 To 16)
    ReDim mBucket(1 To 16)
    ReDim mQual(1 To 16)
    ReDim mSub(1 To 16)
    ReDim mT1(1 To 16)
    ReDim mT2(1 To 16)
    ReDim mA1(1 To 16)
    ReDim mA2(1 To 16)
    ReDim mWS(1 To 16)
    Set mKeys = New Collection
End Sub

Private Sub GrowStore()
    Dim n As Long
    n = 2 * UBound(mRC)
    ReDim Preserve mRC(1 To n)
    ReDim Preserve mBucket(1 To n)
    ReDim Preserve mQual(1 To n)
    ReDim Preserve mSub(1 To n)
    ReDim Preserve mT1(1 To n)
    ReDim Preserve mT2(1 To n)
    ReDim Preserve mA1(1 To n)
    ReDim Preserve mA2(1 To n)
    ReDim Preserve mWS(1 To n)
End Sub

' Adds a risk factor or nets the amounts into an existing one (MAR21.4(2)).
Private Sub AddRiskFactor(ByVal rc As String, ByVal bucket As String, ByVal qual As String, _
                          ByVal subType As String, ByVal t1 As Double, ByVal t2 As Double, _
                          ByVal a1 As Double, ByVal a2 As Double)
    Dim key As String, idx As Long
    key = rc & "|" & bucket & "|" & qual & "|" & subType & "|" & TenorKey(t1) & "|" & TenorKey(t2)
    idx = KeyIndex(mKeys, key)
    If idx = 0 Then
        mN = mN + 1
        If mN > UBound(mRC) Then GrowStore
        idx = mN
        mRC(idx) = rc
        mBucket(idx) = bucket
        mQual(idx) = qual
        mSub(idx) = subType
        mT1(idx) = t1
        mT2(idx) = t2
        mA1(idx) = 0#
        mA2(idx) = 0#
        AddKey mKeys, key, idx
    End If
    mA1(idx) = mA1(idx) + a1
    mA2(idx) = mA2(idx) + a2
End Sub

'------------------------------------------------------------------------------
' Input loading and validation
'------------------------------------------------------------------------------
Private Function ValidateCommon(ByVal rc As String, ByVal bucket As String, ByVal src As String) As Boolean
    ValidateCommon = False
    Select Case rc
        Case RC_GIRR, RC_CSR, RC_EQ, RC_COM, RC_FX
        Case "CSR_SEC_NCTP", "CSR_SEC_CTP"
            LogMessage "ERROR", src, "Risk class " & rc & " is not implemented in this prototype; row skipped"
            Exit Function
        Case Else
            LogMessage "ERROR", src, "Unknown risk class '" & rc & "'; row skipped"
            Exit Function
    End Select
    If Not IsValidBucket(rc, bucket) Then
        LogMessage "ERROR", src, "Invalid bucket '" & bucket & "' for " & rc & "; row skipped"
        Exit Function
    End If
    If rc = RC_FX And bucket = ReportingCcy() Then
        LogMessage "ERROR", src, "FX bucket equals the reporting currency " & bucket & "; row skipped"
        Exit Function
    End If
    ValidateCommon = True
End Function

Private Sub LoadDelta()
    Dim ws As Object, r As Long, src As String
    Dim rc As String, bucket As String, qual As String, subType As String, tenor As Double
    Dim ok As Boolean

    ResetStore
    Set ws = GetSheet(SH_IN_DELTA)
    r = LAYOUT_FIRST_DATA_ROW
    Do While Not (IsBlank(TableCell(ws, r, 1).Value) And IsBlank(TableCell(ws, r, 2).Value))
        src = SH_IN_DELTA & " row " & r
        rc = ToText(TableCell(ws, r, 2).Value)
        bucket = BucketLabel(TableCell(ws, r, 3).Value)
        qual = ToText(TableCell(ws, r, 4).Value)
        subType = ToText(TableCell(ws, r, 5).Value)
        ok = ValidateCommon(rc, bucket, src)
        If ok And Not IsNumber(TableCell(ws, r, 7).Value) Then
            LogMessage "ERROR", src, "Sensitivity is not numeric; row skipped"
            ok = False
        End If
        If ok And Not IsBlankOrNumber(TableCell(ws, r, 6).Value) Then
            LogMessage "ERROR", src, "Tenor is not numeric; row skipped"
            ok = False
        End If
        If ok Then
            tenor = ToDouble(TableCell(ws, r, 6).Value, 0#)
            Select Case rc
                Case RC_GIRR
                    If subType = "" Then subType = "YIELD"
                    Select Case subType
                        Case "YIELD"
                            If Not IsValidGirrTenor(tenor) Then
                                LogMessage "ERROR", src, "GIRR tenor " & tenor & " is not a regulatory vertex; row skipped"
                                ok = False
                            End If
                        Case "INFLATION", "XCCY"
                            tenor = 0#
                        Case Else
                            LogMessage "ERROR", src, "GIRR sub-type must be YIELD, INFLATION or XCCY; row skipped"
                            ok = False
                    End Select
                Case RC_CSR
                    If subType = "" Then subType = "BOND"
                    If subType <> "BOND" And subType <> "CDS" Then
                        LogMessage "ERROR", src, "CSR sub-type must be BOND or CDS; row skipped"
                        ok = False
                    ElseIf Not IsValidCsrTenor(tenor) Then
                        LogMessage "ERROR", src, "CSR tenor " & tenor & " is not a regulatory vertex; row skipped"
                        ok = False
                    End If
                Case RC_EQ
                    If subType = "" Then subType = "SPOT"
                    If subType <> "SPOT" And subType <> "REPO" Then
                        LogMessage "ERROR", src, "Equity sub-type must be SPOT or REPO; row skipped"
                        ok = False
                    End If
                    tenor = 0#
                Case RC_COM
                    If Not IsValidComTenor(tenor) Then
                        LogMessage "ERROR", src, "Commodity tenor " & tenor & " is not a regulatory vertex; row skipped"
                        ok = False
                    End If
                Case RC_FX
                    If subType <> "" Then
                        LogMessage "ERROR", src, "FX takes no sub-type; row skipped"
                        ok = False
                    End If
                    qual = bucket
                    subType = ""
                    tenor = 0#
            End Select
            If ok And qual = "" Then
                LogMessage "ERROR", src, "Qualifier (curve, issuer or commodity) is required; row skipped"
                ok = False
            End If
        End If
        If ok Then
            AddRiskFactor rc, bucket, qual, subType, tenor, 0#, ToDouble(TableCell(ws, r, 7).Value, 0#), 0#
        End If
        r = r + 1
    Loop

    For r = 1 To mN
        mWS(r) = DeltaRiskWeight(r) * mA1(r)
    Next r
End Sub

Private Function DeltaRiskWeight(ByVal i As Long) As Double
    Select Case mRC(i)
        Case RC_GIRR: DeltaRiskWeight = GirrRiskWeight(mBucket(i), mSub(i), mT1(i))
        Case RC_CSR: DeltaRiskWeight = CsrRiskWeight(mBucket(i))
        Case RC_EQ: DeltaRiskWeight = EqRiskWeight(mBucket(i), mSub(i))
        Case RC_COM: DeltaRiskWeight = ComRiskWeight(mBucket(i))
        Case RC_FX: DeltaRiskWeight = FxRiskWeight(mBucket(i))
    End Select
End Function

' Vega and curvature take the sub-types of the delta that are the same risk
' factor: the bond or CDS curve of a CSR issuer, and the equity spot price.
' Equity repo rates have no vega or curvature (MAR21.12); a commodity
' sub-type is the delivery location and stays free text.
Private Function IsNameSubType(ByVal rc As String, ByVal subType As String, _
                               ByVal src As String, ByVal kind As String) As Boolean
    IsNameSubType = True
    Select Case rc
        Case RC_CSR
            If subType <> "" And subType <> "BOND" And subType <> "CDS" Then
                LogMessage "ERROR", src, "CSR " & kind & " sub-type must be BOND, CDS or none; row skipped"
                IsNameSubType = False
            End If
        Case RC_EQ
            If subType <> "" And subType <> "SPOT" Then
                LogMessage "ERROR", src, "Equity " & kind & " takes sub-type SPOT or none: repo rates have no " & kind & "; row skipped"
                IsNameSubType = False
            End If
    End Select
End Function

Private Sub LoadVega()
    Dim ws As Object, r As Long, src As String
    Dim rc As String, bucket As String, qual As String, subType As String
    Dim optMat As Double, undMat As Double, ok As Boolean

    ResetStore
    Set ws = GetSheet(SH_IN_VEGA)
    r = LAYOUT_FIRST_DATA_ROW
    Do While Not (IsBlank(TableCell(ws, r, 1).Value) And IsBlank(TableCell(ws, r, 2).Value))
        src = SH_IN_VEGA & " row " & r
        rc = ToText(TableCell(ws, r, 2).Value)
        bucket = BucketLabel(TableCell(ws, r, 3).Value)
        qual = ToText(TableCell(ws, r, 4).Value)
        subType = ToText(TableCell(ws, r, 5).Value)
        ok = ValidateCommon(rc, bucket, src)
        If ok And Not IsNumber(TableCell(ws, r, 8).Value) Then
            LogMessage "ERROR", src, "Vega sensitivity is not numeric; row skipped"
            ok = False
        End If
        If ok And Not (IsBlankOrNumber(TableCell(ws, r, 6).Value) And IsBlankOrNumber(TableCell(ws, r, 7).Value)) Then
            LogMessage "ERROR", src, "Option or underlying maturity is not numeric; row skipped"
            ok = False
        End If
        If ok Then
            optMat = ToDouble(TableCell(ws, r, 6).Value, 0#)
            undMat = ToDouble(TableCell(ws, r, 7).Value, 0#)
        End If
        If ok And Not IsValidVegaMaturity(optMat) Then
            LogMessage "ERROR", src, "Option maturity " & optMat & " is not a regulatory vertex; row skipped"
            ok = False
        End If
        If ok Then
            Select Case rc
                Case RC_GIRR
                    If subType = "" Then subType = "YIELD"
                    Select Case subType
                        Case "YIELD"
                            ' One vega surface per currency: option x underlying maturity
                            If Not IsValidVegaMaturity(undMat) Then
                                LogMessage "ERROR", src, "GIRR underlying maturity " & undMat & " is not a regulatory vertex; row skipped"
                                ok = False
                            End If
                            qual = ""
                        Case "INFLATION", "XCCY"
                            ' Inflation and cross-currency basis vega: the option
                            ' maturity only, one risk factor per curve (MAR21.8 FAQ4)
                            If Not IsBlank(TableCell(ws, r, 7).Value) Then
                                LogMessage "ERROR", src, "GIRR " & subType & " vega takes no underlying maturity; row skipped"
                                ok = False
                            End If
                            undMat = 0#
                        Case Else
                            LogMessage "ERROR", src, "GIRR sub-type must be YIELD, INFLATION or XCCY; row skipped"
                            ok = False
                    End Select
                Case RC_CSR, RC_EQ, RC_COM
                    ' Vega risk factors are the underlying name and the option
                    ' maturity: no bond/CDS basis or delivery location (MAR21.94 FAQ)
                    If Not IsNameSubType(rc, subType, src, "vega") Then ok = False
                    subType = ""
                    undMat = 0#
                Case RC_FX
                    If subType <> "" Then
                        LogMessage "ERROR", src, "FX takes no sub-type; row skipped"
                        ok = False
                    End If
                    qual = bucket
                    subType = ""
                    undMat = 0#
            End Select
            If ok And qual = "" And rc <> RC_GIRR Then
                LogMessage "ERROR", src, "Qualifier (issuer or commodity) is required; row skipped"
                ok = False
            End If
        End If
        If ok Then
            AddRiskFactor rc, bucket, qual, subType, optMat, undMat, ToDouble(TableCell(ws, r, 8).Value, 0#), 0#
        End If
        r = r + 1
    Loop

    For r = 1 To mN
        mWS(r) = VegaRiskWeight(mRC(r), mBucket(r)) * mA1(r)
    Next r
End Sub

Private Sub LoadCurvature()
    Dim ws As Object, r As Long, src As String
    Dim rc As String, bucket As String, qual As String, subType As String
    Dim up As Double, dn As Double, ok As Boolean, refCcy As Boolean

    ResetStore
    Set ws = GetSheet(SH_IN_CURV)
    r = LAYOUT_FIRST_DATA_ROW
    Do While Not (IsBlank(TableCell(ws, r, 1).Value) And IsBlank(TableCell(ws, r, 2).Value))
        src = SH_IN_CURV & " row " & r
        rc = ToText(TableCell(ws, r, 2).Value)
        bucket = BucketLabel(TableCell(ws, r, 3).Value)
        qual = ToText(TableCell(ws, r, 4).Value)
        subType = ToText(TableCell(ws, r, 5).Value)
        ok = ValidateCommon(rc, bucket, src)
        If ok And (Not IsNumber(TableCell(ws, r, 6).Value) Or Not IsNumber(TableCell(ws, r, 7).Value)) Then
            LogMessage "ERROR", src, "CVR up and CVR down must both be numeric; row skipped"
            ok = False
        End If
        If ok And Not IsBlankOrBoolCell(TableCell(ws, r, 8)) Then
            LogMessage "ERROR", src, "References reporting ccy (RefReportingCcy) must be TRUE, FALSE or blank; row skipped"
            ok = False
        End If
        If ok Then
            up = ToDouble(TableCell(ws, r, 6).Value, 0#)
            dn = ToDouble(TableCell(ws, r, 7).Value, 0#)
            refCcy = CellBool(TableCell(ws, r, 8))
            Select Case rc
                Case RC_GIRR
                    ' All curves of a currency are shifted together: one risk factor
                    ' per bucket, with no inflation or cross-currency basis curvature (MAR21.8)
                    If subType <> "" And subType <> "YIELD" Then
                        LogMessage "ERROR", src, "GIRR curvature takes sub-type YIELD or none: inflation and cross-currency basis have no curvature; row skipped"
                        ok = False
                    End If
                    qual = ""
                    subType = ""
                Case RC_CSR, RC_EQ, RC_COM
                    ' One curvature risk factor per issuer or commodity: bond and
                    ' CDS curves, and delivery locations, are not separate (MAR21.8-MAR21.14)
                    If Not IsNameSubType(rc, subType, src, "curvature") Then ok = False
                    subType = ""
                Case RC_FX
                    If subType <> "" Then
                        LogMessage "ERROR", src, "FX takes no sub-type; row skipped"
                        ok = False
                    End If
                    qual = bucket
                    subType = ""
                    up = up / FxCurvatureDivisor(refCcy)
                    dn = dn / FxCurvatureDivisor(refCcy)
            End Select
            If ok And qual = "" And rc <> RC_GIRR Then
                LogMessage "ERROR", src, "Qualifier (issuer or commodity) is required; row skipped"
                ok = False
            End If
        End If
        If ok Then AddRiskFactor rc, bucket, qual, subType, 0#, 0#, up, dn
        r = r + 1
    Loop
End Sub

'------------------------------------------------------------------------------
' Correlations within a bucket (medium scenario, before the transform)
'------------------------------------------------------------------------------
Private Function DeltaCorrelation(ByVal rc As String, ByVal bucket As String, _
                                  ByVal i As Long, ByVal j As Long) As Double
    Dim rho As Double
    Select Case rc
        Case RC_GIRR
            If mSub(i) = "XCCY" Or mSub(j) = "XCCY" Then
                rho = GirrRhoXccy()
            ElseIf mSub(i) = "INFLATION" And mSub(j) = "INFLATION" Then
                rho = GirrRhoDiffCurve()
            ElseIf mSub(i) = "INFLATION" Or mSub(j) = "INFLATION" Then
                rho = GirrRhoInflation()
            Else
                rho = GirrTenorCorrelation(mT1(i), mT1(j))
                If mQual(i) <> mQual(j) Then rho = rho * GirrRhoDiffCurve()
            End If
        Case RC_CSR
            rho = CsrIssuerCorrelation(bucket, i, j)
            If Abs(mT1(i) - mT1(j)) > 0.000001 Then rho = rho * CsrRhoTenor()
        Case RC_EQ
            rho = EqIssuerCorrelation(bucket, i, j)
            If mSub(i) <> mSub(j) Then rho = rho * EqRhoSpotRepo()
        Case RC_COM
            rho = ComUnderlyingCorrelation(bucket, i, j)
            If Abs(mT1(i) - mT1(j)) > 0.000001 Then rho = rho * ComRhoTenor()
        Case Else
            rho = 1#
    End Select
    DeltaCorrelation = rho
End Function

' Name x basis components for CSR (tenor excluded).
Private Function CsrIssuerCorrelation(ByVal bucket As String, ByVal i As Long, ByVal j As Long) As Double
    Dim rho As Double
    If mQual(i) = mQual(j) Then rho = 1# Else rho = CsrNameCorrelation(bucket)
    If mSub(i) <> mSub(j) Then rho = rho * CsrRhoBasis()
    CsrIssuerCorrelation = rho
End Function

Private Function EqIssuerCorrelation(ByVal bucket As String, ByVal i As Long, ByVal j As Long) As Double
    If mQual(i) = mQual(j) Then
        EqIssuerCorrelation = 1#
    Else
        EqIssuerCorrelation = EqNameCorrelation(bucket)
    End If
End Function

' Commodity x delivery location components (tenor excluded).
Private Function ComUnderlyingCorrelation(ByVal bucket As String, ByVal i As Long, ByVal j As Long) As Double
    Dim rho As Double
    If mQual(i) = mQual(j) Then rho = 1# Else rho = ComCommodityCorrelation(bucket)
    If mSub(i) <> mSub(j) Then rho = rho * ComRhoBasis()
    ComUnderlyingCorrelation = rho
End Function

Private Function MaturityCorrelation(ByVal ta As Double, ByVal tb As Double) As Double
    If Abs(ta - tb) < 0.000001 Then
        MaturityCorrelation = 1#
    Else
        MaturityCorrelation = Exp(-VegaAlpha() * Abs(ta - tb) / MinD(ta, tb))
    End If
End Function

' Delta correlation restricted to the dimension that vega and curvature risk
' factors share with delta: the name for CSR and equity, the commodity for
' commodity (MAR21.94 and MAR21.100 with their FAQs).
Private Function UnderlyingCorrelation(ByVal rc As String, ByVal bucket As String, _
                                       ByVal i As Long, ByVal j As Long) As Double
    If mQual(i) = mQual(j) Then
        UnderlyingCorrelation = 1#
    Else
        Select Case rc
            Case RC_CSR: UnderlyingCorrelation = CsrNameCorrelation(bucket)
            Case RC_EQ: UnderlyingCorrelation = EqNameCorrelation(bucket)
            Case RC_COM: UnderlyingCorrelation = ComCommodityCorrelation(bucket)
            Case Else: UnderlyingCorrelation = 1#
        End Select
    End If
End Function


' Delta correlation of the GIRR delta risk factors two vega risk factors
' correspond to (MAR21.94, MAR21.46-49): the underlying maturity between two
' yield surfaces, otherwise the sub-type and curve correlations.
Private Function GirrVegaUnderlyingCorrelation(ByVal i As Long, ByVal j As Long) As Double
    If mSub(i) = "YIELD" And mSub(j) = "YIELD" Then
        GirrVegaUnderlyingCorrelation = MaturityCorrelation(mT2(i), mT2(j))
    ElseIf mSub(i) = mSub(j) And mQual(i) = mQual(j) Then
        GirrVegaUnderlyingCorrelation = 1#
    ElseIf mSub(i) = "XCCY" Or mSub(j) = "XCCY" Then
        GirrVegaUnderlyingCorrelation = GirrRhoXccy()
    ElseIf mSub(i) = "INFLATION" And mSub(j) = "INFLATION" Then
        GirrVegaUnderlyingCorrelation = GirrRhoDiffCurve()
    Else
        GirrVegaUnderlyingCorrelation = GirrRhoInflation()
    End If
End Function


' Vega (MAR21.93-95): min(rho_option_maturity x rho_underlying, 1).
Private Function VegaCorrelation(ByVal rc As String, ByVal bucket As String, _
                                 ByVal i As Long, ByVal j As Long) As Double
    Dim rhoUnd As Double
    If rc = RC_GIRR Then
        rhoUnd = GirrVegaUnderlyingCorrelation(i, j)
    Else
        rhoUnd = UnderlyingCorrelation(rc, bucket, i, j)
    End If
    VegaCorrelation = MinD(MaturityCorrelation(mT1(i), mT1(j)) * rhoUnd, 1#)
End Function


' Curvature (MAR21.100): square of the underlying delta correlation.
Private Function CurvatureCorrelation(ByVal rc As String, ByVal bucket As String, _
                                      ByVal i As Long, ByVal j As Long) As Double
    Dim rho As Double
    rho = UnderlyingCorrelation(rc, bucket, i, j)
    CurvatureCorrelation = rho * rho
End Function


Private Function PairCorrelation(ByVal comp As Long, ByVal rc As String, ByVal bucket As String, _
                                 ByVal i As Long, ByVal j As Long, ByVal scen As Long) As Double
    Dim rho As Double
    Select Case comp
        Case COMP_DELTA: rho = DeltaCorrelation(rc, bucket, i, j)
        Case COMP_VEGA: rho = VegaCorrelation(rc, bucket, i, j)
        Case Else: rho = CurvatureCorrelation(rc, bucket, i, j)
    End Select
    PairCorrelation = ScenarioCorrelation(rho, scen)
End Function

Private Function CrossBucketCorrelation(ByVal comp As Long, ByVal rc As String, _
                                        ByVal b As String, ByVal c As String, ByVal scen As Long) As Double
    Dim g As Double
    g = BucketGamma(rc, b, c)
    If comp = COMP_CURV Then g = g * g
    CrossBucketCorrelation = ScenarioCorrelation(g, scen)
End Function

Private Function Psi(ByVal x As Double, ByVal y As Double) As Double
    If x < 0# And y < 0# Then Psi = 0# Else Psi = 1#
End Function

'------------------------------------------------------------------------------
' Aggregation
'------------------------------------------------------------------------------
Private Sub ComputeComponent(ByVal comp As Long)
    Dim rcIdx As Long
    For rcIdx = 1 To RC_COUNT
        ComputeRiskClass comp, rcIdx
    Next rcIdx
End Sub

Private Sub ComputeRiskClass(ByVal comp As Long, ByVal rcIdx As Long)
    Dim rc As String, i As Long, j As Long, k As Long, l As Long, b As Long, c As Long
    Dim nB As Long, labels() As String, found As Boolean
    Dim counts() As Long, members() As Long, maxCount As Long
    Dim isOther() As Boolean, kb() As Double, sb() As Double, note() As String
    Dim scen As Long, sumSq As Double, total As Double, otherSum As Double
    Dim kUp As Double, kDn As Double, sUp As Double, sDn As Double
    Dim x As Double, y As Double, rho As Double, alt As Boolean

    rc = RiskClassCode(rcIdx)

    ' Distinct buckets of this risk class, sorted
    nB = 0
    ReDim labels(1 To 1)
    For i = 1 To mN
        If mRC(i) = rc Then
            found = False
            For b = 1 To nB
                If labels(b) = mBucket(i) Then
                    found = True
                    Exit For
                End If
            Next b
            If Not found Then
                nB = nB + 1
                ReDim Preserve labels(1 To nB)
                labels(nB) = mBucket(i)
            End If
        End If
    Next i
    If nB = 0 Then Exit Sub
    SortLabels labels, nB
    mBucketCount(rcIdx, comp) = nB

    ' Risk factor indices per bucket
    ReDim counts(1 To nB)
    maxCount = 0
    For b = 1 To nB
        For i = 1 To mN
            If mRC(i) = rc And mBucket(i) = labels(b) Then counts(b) = counts(b) + 1
        Next i
        If counts(b) > maxCount Then maxCount = counts(b)
    Next b
    ReDim members(1 To nB, 1 To maxCount)
    ReDim isOther(1 To nB)
    For b = 1 To nB
        k = 0
        For i = 1 To mN
            If mRC(i) = rc And mBucket(i) = labels(b) Then
                k = k + 1
                members(b, k) = i
            End If
        Next i
        isOther(b) = IsOtherBucket(rc, labels(b))
    Next b

    ReDim kb(1 To nB)
    ReDim sb(1 To nB)
    ReDim note(1 To nB)

    For scen = 1 To SCEN_COUNT
        ' Bucket level
        For b = 1 To nB
            note(b) = ""
            If comp = COMP_CURV Then
                kUp = 0#
                kDn = 0#
                sUp = 0#
                sDn = 0#
                For k = 1 To counts(b)
                    i = members(b, k)
                    sUp = sUp + mA1(i)
                    sDn = sDn + mA2(i)
                    If isOther(b) Then
                        kUp = kUp + MaxD(mA1(i), 0#)
                        kDn = kDn + MaxD(mA2(i), 0#)
                    Else
                        kUp = kUp + MaxD(mA1(i), 0#) ^ 2
                        kDn = kDn + MaxD(mA2(i), 0#) ^ 2
                        For l = 1 To counts(b)
                            If l <> k Then
                                j = members(b, l)
                                rho = PairCorrelation(comp, rc, labels(b), i, j, scen)
                                kUp = kUp + rho * mA1(i) * mA1(j) * Psi(mA1(i), mA1(j))
                                kDn = kDn + rho * mA2(i) * mA2(j) * Psi(mA2(i), mA2(j))
                            End If
                        Next l
                    End If
                Next k
                If Not isOther(b) Then
                    kUp = Sqr(MaxD(kUp, 0#))
                    kDn = Sqr(MaxD(kDn, 0#))
                End If
                ' Scenario selection (MAR21.5(3)): larger K, ties broken by larger sum
                If NearlyEqual(kUp, kDn) Then
                    If sUp > sDn Then note(b) = "UP" Else note(b) = "DOWN"
                ElseIf kUp > kDn Then
                    note(b) = "UP"
                Else
                    note(b) = "DOWN"
                End If
                If note(b) = "UP" Then
                    kb(b) = kUp
                    sb(b) = sUp
                Else
                    kb(b) = kDn
                    sb(b) = sDn
                End If
            Else
                kb(b) = 0#
                sb(b) = 0#
                For k = 1 To counts(b)
                    i = members(b, k)
                    sb(b) = sb(b) + mWS(i)
                    If isOther(b) Then
                        kb(b) = kb(b) + Abs(mWS(i))
                    Else
                        kb(b) = kb(b) + mWS(i) ^ 2
                        For l = 1 To counts(b)
                            If l <> k Then
                                j = members(b, l)
                                kb(b) = kb(b) + PairCorrelation(comp, rc, labels(b), i, j, scen) * mWS(i) * mWS(j)
                            End If
                        Next l
                    End If
                Next k
                If Not isOther(b) Then kb(b) = Sqr(MaxD(kb(b), 0#))
            End If
            If isOther(b) Then
                If note(b) = "" Then note(b) = "OTHER" Else note(b) = note(b) & ";OTHER"
            End If
        Next b

        ' Across buckets
        alt = False
        sumSq = AcrossBuckets(comp, rc, nB, labels, isOther, kb, sb, scen, False)
        If sumSq < 0# And comp <> COMP_CURV Then
            ' Alternative specification Sb = max(min(sum WS, Kb), -Kb) (MAR21.4(5)(b))
            alt = True
            sumSq = AcrossBuckets(comp, rc, nB, labels, isOther, kb, sb, scen, True)
        End If
        otherSum = 0#
        For b = 1 To nB
            If isOther(b) Then otherSum = otherSum + kb(b)
        Next b
        total = Sqr(MaxD(sumSq, 0#)) + otherSum
        mCapital(rcIdx, comp, scen) = total
        mAltUsed(rcIdx, comp, scen) = alt

        For b = 1 To nB
            If alt And Not isOther(b) Then
                x = MaxD(MinD(sb(b), kb(b)), -kb(b))
                WriteBucketRow rc, comp, scen, labels(b), kb(b), x, counts(b), _
                               IIf(note(b) = "", "ALT_SB", note(b) & ";ALT_SB")
            Else
                WriteBucketRow rc, comp, scen, labels(b), kb(b), sb(b), counts(b), note(b)
            End If
        Next b
    Next scen
End Sub

Private Function AcrossBuckets(ByVal comp As Long, ByVal rc As String, ByVal nB As Long, _
                               ByRef labels() As String, ByRef isOther() As Boolean, _
                               ByRef kb() As Double, ByRef sb() As Double, _
                               ByVal scen As Long, ByVal useAlt As Boolean) As Double
    Dim b As Long, c As Long, total As Double, x As Double, y As Double
    total = 0#
    For b = 1 To nB
        If Not isOther(b) Then
            total = total + kb(b) ^ 2
            For c = 1 To nB
                If c <> b And Not isOther(c) Then
                    x = sb(b)
                    y = sb(c)
                    If useAlt Then
                        x = MaxD(MinD(x, kb(b)), -kb(b))
                        y = MaxD(MinD(y, kb(c)), -kb(c))
                    End If
                    If comp = COMP_CURV Then
                        total = total + CrossBucketCorrelation(comp, rc, labels(b), labels(c), scen) * x * y * Psi(x, y)
                    Else
                        total = total + CrossBucketCorrelation(comp, rc, labels(b), labels(c), scen) * x * y
                    End If
                End If
            Next c
        End If
    Next b
    AcrossBuckets = total
End Function

'------------------------------------------------------------------------------
' Output
'------------------------------------------------------------------------------
Private Sub ResetBucketOutput()
    Dim ws As Object
    Set ws = GetSheet(SH_OUT_BUCKETS)
    ClearTableArea ws
    WriteTableHeader ws, LAYOUT_HEADER_ROW, _
        Array("RiskClass", "Component", "Scenario", "Bucket", "Kb", "Sb", "RiskFactors", "Note"), COLOR_OUTPUT_HEADER
    mOutRow = LAYOUT_HEADER_ROW
End Sub

Private Sub WriteBucketRow(ByVal rc As String, ByVal comp As Long, ByVal scen As Long, ByVal bucket As String, _
                           ByVal kb As Double, ByVal sb As Double, ByVal nRF As Long, ByVal note As String)
    Dim ws As Object
    Set ws = GetSheet(SH_OUT_BUCKETS)
    mOutRow = mOutRow + 1
    TableCell(ws, mOutRow, 1).Value = rc
    TableCell(ws, mOutRow, 2).Value = ComponentName(comp)
    TableCell(ws, mOutRow, 3).Value = ScenarioName(scen)
    TableCell(ws, mOutRow, 4).Value = bucket
    TableCell(ws, mOutRow, 5).Value = kb
    TableCell(ws, mOutRow, 6).Value = sb
    TableCell(ws, mOutRow, 7).Value = nRF
    TableCell(ws, mOutRow, 8).Value = note
End Sub

Private Sub WriteRiskClassOutput()
    Dim ws As Object, r As Long, i As Long, c As Long, s As Long
    Set ws = GetSheet(SH_OUT_CLASS)
    ClearTableArea ws
    WriteTableHeader ws, LAYOUT_HEADER_ROW, _
        Array("RiskClass", "Component", "Scenario", "Capital", "Buckets", "AltSbUsed"), COLOR_OUTPUT_HEADER
    r = LAYOUT_HEADER_ROW
    For i = 1 To RC_COUNT
        For c = 1 To COMP_COUNT
            For s = 1 To SCEN_COUNT
                r = r + 1
                TableCell(ws, r, 1).Value = RiskClassCode(i)
                TableCell(ws, r, 2).Value = ComponentName(c)
                TableCell(ws, r, 3).Value = ScenarioName(s)
                TableCell(ws, r, 4).Value = mCapital(i, c, s)
                TableCell(ws, r, 5).Value = mBucketCount(i, c)
                TableCell(ws, r, 6).Value = IIf(mAltUsed(i, c, s), "Y", "N")
            Next s
        Next c
    Next i
    FormatDataRows ws, LAYOUT_FIRST_DATA_ROW, r, 6, 4, "#,##0", False
End Sub
