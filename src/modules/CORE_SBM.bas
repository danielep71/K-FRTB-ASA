Attribute VB_Name = "CORE_SBM"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_SBM
'------------------------------------------------------------------------------
' PURPOSE
'   Calculates the sensitivities-based method (MAR21): delta, vega and curvature capital for
'   GIRR, CSR non-securitization, equity, commodity and FX under the low, medium and high
'   correlation scenarios.
'
' WHY THIS EXISTS
'   The sensitivities-based method is a capital component of its own, next to the default
'   risk charge and the residual risk add-on, so it has its own engine module.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Flow per component:
'     - reads and validates the input sheet and nets identical risk factors
'     - per risk class and bucket: weighted sensitivities, Kb and Sb
'     - across buckets: aggregates with gamma (alternative Sb when needed)
'
' ERROR POLICY
'   Invalid input rows are logged as ERROR and skipped, never counted as zero. Missing
'   parameters raise from CORE_Params.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, CORE_Params (the parameters must be loaded).
'
' NOTES
'   The module also runs in LibreOffice Calc with VBA support for the automated
'   tests. It needs no references beyond the defaults.
'
' PUBLIC SURFACE
'   RunSBM; SbmCapital(risk class, component, scenario); SbmScenarioTotal(scenario).
'
' STATE OWNERSHIP
'   Owns the netted risk factors of the component being processed and the capital of
'   the last run. Reads In_Delta, In_Vega and In_Curvature; writes Out_Buckets and
'   Out_RiskClass and logs rejected rows.
'
' REFERENCE
'   Basel Framework MAR21.4-MAR21.101.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'Netted risk factors of the component being processed
    Private mN          As Long         'Number of netted risk factors
    Private mRC()       As String       'Risk class of each risk factor
    Private mBucket()   As String       'Bucket of each risk factor
    Private mQual()     As String       'Qualifier (curve, issuer or commodity) of each risk factor
    Private mSub()      As String       'Sub-type of each risk factor
    Private mT1()       As Double       'Delta: tenor; vega: option maturity
    Private mT2()       As Double       'Vega: GIRR underlying maturity
    Private mA1()       As Double       'Delta and vega: sensitivity; curvature: CVR up
    Private mA2()       As Double       'Curvature: CVR down
    Private mWS()       As Double       'Weighted sensitivity (delta and vega)
    Private mKeys       As Collection   'Index of the risk factors by key

    'Results: risk class x component x scenario
    Private mCapital(1 To 5, 1 To 3, 1 To 3)    As Double   'Capital of the last run
    Private mAltUsed(1 To 5, 1 To 3, 1 To 3)    As Boolean  'True when the alternative Sb was used
    Private mBucketCount(1 To 5, 1 To 3)        As Long     'Number of buckets by risk class and component
    Private mOutRow                             As Long     'Last row written to the bucket output table


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunSBM()

'
'------------------------------------------------------------------------------
'                          CALCULATE THE SBM CAPITAL
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the sensitivities-based method: loads the delta, vega and curvature inputs, computes
'   the capital of each risk class, component and scenario and writes the bucket and risk
'   class output tables.
'
' WHY THIS EXISTS
'   Callers need one entry point that runs the whole method and leaves the capital
'   available through SbmCapital and SbmScenarioTotal.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The results are read with SbmCapital and SbmScenarioTotal.
'
' BEHAVIOR
'   In order:
'     - clears the capital, the alternative-Sb flags and the bucket counts, and resets the
'       bucket output table
'     - for delta, vega and curvature in turn: loads and validates the input sheet, then
'       computes the component
'     - formats the bucket output table and writes the risk class output table
'
' ERROR POLICY
'   Does not handle errors. Invalid input rows are logged as ERROR and skipped by the
'   loaders; errors from the callees (for example a missing parameter) are not caught
'   here.
'
' DEPENDENCIES
'   ResetBucketOutput, LoadDelta, LoadVega, LoadCurvature, ComputeComponent,
'   FormatBucketOutput, WriteRiskClassOutput; RC_COUNT, COMP_COUNT, SCEN_COUNT,
'   COMP_DELTA, COMP_VEGA, COMP_CURV (CORE_Config).
'
' NOTES
'   The parameters must be loaded before the call.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.RunSBM"

    Dim ClassIndex              As Long          'Risk class being reset
    Dim ComponentIndex          As Long          'Component being reset
    Dim ScenarioIndex           As Long          'Scenario being reset

'------------------------------------------------------------------------------
' RESET RESULTS
'------------------------------------------------------------------------------
    'Clear the capital, the alternative-Sb flags and the bucket counts of the last run
        For ClassIndex = 1 To RC_COUNT
            For ComponentIndex = 1 To COMP_COUNT
                mBucketCount(ClassIndex, ComponentIndex) = 0
                For ScenarioIndex = 1 To SCEN_COUNT
                    mCapital(ClassIndex, ComponentIndex, ScenarioIndex) = 0#
                    mAltUsed(ClassIndex, ComponentIndex, ScenarioIndex) = False
                Next ScenarioIndex
            Next ComponentIndex
        Next ClassIndex
    'Clear the bucket output table
        ResetBucketOutput

'------------------------------------------------------------------------------
' COMPUTE COMPONENTS
'------------------------------------------------------------------------------
    'Load the delta rows and compute the delta capital
        LoadDelta
        ComputeComponent COMP_DELTA
    'Load the vega rows and compute the vega capital
        LoadVega
        ComputeComponent COMP_VEGA
    'Load the curvature rows and compute the curvature capital
        LoadCurvature
        ComputeComponent COMP_CURV

'------------------------------------------------------------------------------
' WRITE OUTPUT
'------------------------------------------------------------------------------
    'Format the bucket output table
        FormatBucketOutput
    'Write the risk class output table
        WriteRiskClassOutput

End Sub

Public Function SbmCapital(ByVal rcIdx As Long, ByVal comp As Long, ByVal scen As Long) As Double

'
'------------------------------------------------------------------------------
'                         RETURN THE CAPITAL OF A CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the capital of one risk class, component and scenario from the last run.
'
' WHY THIS EXISTS
'   The results stay private to the module; callers read them here.
'
' INPUTS
'   rcIdx
'     Risk class index (1 to RC_COUNT)
'
'   comp
'     Component: COMP_DELTA, COMP_VEGA or COMP_CURV
'
'   scen
'     Scenario: SCEN_LOW, SCEN_MEDIUM or SCEN_HIGH
'
' RETURNS
'   The capital stored by the last RunSBM call.
'
' BEHAVIOR
'   Returns the element of the capital array at (rcIdx, comp, scen).
'
' ERROR POLICY
'   Does not check the indexes: an index outside the stored range raises the VBA
'   subscript error.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored capital
        SbmCapital = mCapital(rcIdx, comp, scen)

End Function

Public Function SbmScenarioTotal(ByVal scen As Long) As Double

'
'------------------------------------------------------------------------------
'                        SUM THE CAPITAL OF A SCENARIO
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the simple sum of delta, vega and curvature capital across risk classes for one
'   scenario (MAR21.7).
'
' WHY THIS EXISTS
'   The total of the method for a scenario is read by callers; it is not stored.
'
' INPUTS
'   scen
'     Scenario: SCEN_LOW, SCEN_MEDIUM or SCEN_HIGH
'
' RETURNS
'   The sum of the capital of all risk classes and components for the scenario.
'
' BEHAVIOR
'   Adds the capital of every risk class and component for the scenario.
'
' ERROR POLICY
'   Does not check the scenario: an index outside the stored range raises the VBA
'   subscript error.
'
' DEPENDENCIES
'   RC_COUNT, COMP_COUNT (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.SbmScenarioTotal"

    Dim ClassIndex              As Long          'Risk class being added
    Dim ComponentIndex          As Long          'Component being added
    Dim TotalCapital            As Double        'Running sum of the capital

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the sum at zero
        TotalCapital = 0#

'------------------------------------------------------------------------------
' SUM CAPITAL
'------------------------------------------------------------------------------
    'Add the capital of every risk class and component
        For ClassIndex = 1 To RC_COUNT
            For ComponentIndex = 1 To COMP_COUNT
                TotalCapital = TotalCapital + mCapital(ClassIndex, ComponentIndex, scen)
            Next ComponentIndex
        Next ClassIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the sum
        SbmScenarioTotal = TotalCapital

End Function

'
'------------------------------------------------------------------------------
'
'                              RISK FACTOR STORE
'
'------------------------------------------------------------------------------
'

Private Sub ResetStore()
'
'------------------------------------------------------------------------------
'                         EMPTY THE RISK FACTOR STORE
'------------------------------------------------------------------------------
' PURPOSE
'   Empties the risk factor store and allocates its arrays with room for 16 risk factors.
'
' BEHAVIOR
'   Sets the risk factor count to zero, allocates the nine arrays with bounds 1 to 16 and
'   starts a new key index.
'
' NOTES
'   GrowStore doubles the arrays when they are full.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESET STORE
'------------------------------------------------------------------------------
    'Start with no risk factors
        mN = 0
    'Allocate the arrays
        ReDim mRC(1 To 16)
        ReDim mBucket(1 To 16)
        ReDim mQual(1 To 16)
        ReDim mSub(1 To 16)
        ReDim mT1(1 To 16)
        ReDim mT2(1 To 16)
        ReDim mA1(1 To 16)
        ReDim mA2(1 To 16)
        ReDim mWS(1 To 16)
    'Start a new key index
        Set mKeys = New Collection

End Sub

Private Sub GrowStore()
'
'------------------------------------------------------------------------------
'                           DOUBLE THE STORE ARRAYS
'------------------------------------------------------------------------------
' PURPOSE
'   Doubles the size of the risk factor arrays and keeps their content.
'
' BEHAVIOR
'   Sets the new upper bound to twice the current one and resizes the nine arrays.
'
' NOTES
'   Called by AddRiskFactor when the arrays are full.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.GrowStore"

    Dim NewSize                 As Long          'New upper bound of the arrays

'------------------------------------------------------------------------------
' GROW ARRAYS
'------------------------------------------------------------------------------
    'Compute the doubled size
        NewSize = 2 * UBound(mRC)
    'Resize the arrays and keep their content
        ReDim Preserve mRC(1 To NewSize)
        ReDim Preserve mBucket(1 To NewSize)
        ReDim Preserve mQual(1 To NewSize)
        ReDim Preserve mSub(1 To NewSize)
        ReDim Preserve mT1(1 To NewSize)
        ReDim Preserve mT2(1 To NewSize)
        ReDim Preserve mA1(1 To NewSize)
        ReDim Preserve mA2(1 To NewSize)
        ReDim Preserve mWS(1 To NewSize)

End Sub

Private Sub AddRiskFactor( _
    ByVal rc As String, _
    ByVal bucket As String, _
    ByVal qual As String, _
    ByVal subType As String, _
    ByVal t1 As Double, _
    ByVal t2 As Double, _
    ByVal a1 As Double, _
    ByVal a2 As Double)
'
'------------------------------------------------------------------------------
'                           ADD OR NET A RISK FACTOR
'------------------------------------------------------------------------------
' PURPOSE
'   Adds a risk factor to the store, or nets the amounts into an existing identical one
'   (MAR21.4(2)).
'
' WHY THIS EXISTS
'   Identical risk factors are netted before the bucket calculation.
'
' INPUTS
'   rc
'     Risk class code
'
'   bucket
'     Bucket label
'
'   qual
'     Qualifier (curve, issuer or commodity)
'
'   subType
'     Sub-type
'
'   t1
'     Delta: tenor; vega: option maturity; 0 for curvature
'
'   t2
'     Vega: GIRR underlying maturity; 0 otherwise
'
'   a1
'     Delta and vega: sensitivity; curvature: CVR up
'
'   a2
'     Curvature: CVR down; 0 otherwise
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds a key from the risk class, bucket, qualifier, sub-type and the two maturities and
'   looks it up. When the key is new, appends an entry with zero amounts, growing the
'   arrays when they are full. Then adds a1 and a2 to the entry.
'
' ERROR POLICY
'   Does not handle errors; errors from the callees are not caught here.
'
' DEPENDENCIES
'   TenorKey, KeyIndex, AddKey (CORE_Util); GrowStore.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.AddRiskFactor"

    Dim FactorKey               As String        'Key of the risk factor
    Dim FactorIndex             As Long          'Index of the risk factor in the store

'------------------------------------------------------------------------------
' FIND RISK FACTOR
'------------------------------------------------------------------------------
    'Build the key that identifies the risk factor
        FactorKey = rc & "|" & bucket & "|" & qual & "|" & subType & "|" & TenorKey(t1) & "|" & TenorKey(t2)
    'Look the key up
        FactorIndex = KeyIndex(mKeys, FactorKey)

'------------------------------------------------------------------------------
' ADD RISK FACTOR
'------------------------------------------------------------------------------
    'Append a new entry with zero amounts when the key is not stored yet
        If FactorIndex = 0 Then
            mN = mN + 1
            If mN > UBound(mRC) Then GrowStore
            FactorIndex = mN
            mRC(FactorIndex) = rc
            mBucket(FactorIndex) = bucket
            mQual(FactorIndex) = qual
            mSub(FactorIndex) = subType
            mT1(FactorIndex) = t1
            mT2(FactorIndex) = t2
            mA1(FactorIndex) = 0#
            mA2(FactorIndex) = 0#
            AddKey mKeys, FactorKey, FactorIndex
        End If

'------------------------------------------------------------------------------
' NET AMOUNTS
'------------------------------------------------------------------------------
    'Add the amounts to the entry
        mA1(FactorIndex) = mA1(FactorIndex) + a1
        mA2(FactorIndex) = mA2(FactorIndex) + a2

End Sub

'
'------------------------------------------------------------------------------
'
'                         INPUT LOADING AND VALIDATION
'
'------------------------------------------------------------------------------
'

Private Function ValidateCommon( _
    ByVal rc As String, _
    ByVal bucket As String, _
    ByVal src As String) As Boolean
'
'------------------------------------------------------------------------------
'                        VALIDATE RISK CLASS AND BUCKET
'------------------------------------------------------------------------------
' PURPOSE
'   Checks the risk class and the bucket of an input row and logs the reason when the row is
'   rejected.
'
' WHY THIS EXISTS
'   The delta, vega and curvature loaders share these checks.
'
' INPUTS
'   rc
'     Risk class code read from the row
'
'   bucket
'     Bucket label read from the row
'
'   src
'     Source reference for the Log message (sheet and row)
'
' RETURNS
'   True when the row passes the checks; False when it is rejected.
'
' BEHAVIOR
'   Rejects, in this order:
'     - CSR_SEC_NCTP and CSR_SEC_CTP, which are not implemented in this prototype
'     - any other unknown risk class
'     - a bucket that IsValidBucket does not accept for the risk class
'     - an FX bucket equal to the reporting currency
'
' ERROR POLICY
'   Does not raise. Each rejection is logged as ERROR with the text row skipped.
'
' DEPENDENCIES
'   IsValidBucket, ReportingCcy (CORE_Params); LogMessage (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Reject the row until every check has passed
        ValidateCommon = False

'------------------------------------------------------------------------------
' VALIDATE RISK CLASS
'------------------------------------------------------------------------------
    'Accept the implemented risk classes and reject the others
        Select Case rc
            Case RC_GIRR, RC_CSR, RC_EQ, RC_COM, RC_FX
                ' Implemented risk classes: no rejection here
            Case "CSR_SEC_NCTP", "CSR_SEC_CTP"
                LogMessage "ERROR", src, "Risk class " & rc & " is not implemented in this prototype; row skipped"
                Exit Function
            Case Else
                LogMessage "ERROR", src, "Unknown risk class '" & rc & "'; row skipped"
                Exit Function
        End Select

'------------------------------------------------------------------------------
' VALIDATE BUCKET
'------------------------------------------------------------------------------
    'Reject a bucket that is not valid for the risk class
        If Not IsValidBucket(rc, bucket) Then
            LogMessage "ERROR", src, "Invalid bucket '" & bucket & "' for " & rc & "; row skipped"
            Exit Function
        End If
    'Reject an FX bucket equal to the reporting currency
        If rc = RC_FX And bucket = ReportingCcy() Then
            LogMessage "ERROR", src, "FX bucket equals the reporting currency " & bucket & "; row skipped"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Accept the row
        ValidateCommon = True

End Function

Private Sub LoadDelta()
'
'------------------------------------------------------------------------------
'                             LOAD THE DELTA ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Loads the delta sensitivities of In_Delta into the risk factor store and computes their
'   weighted sensitivities.
'
' WHY THIS EXISTS
'   The delta capital is calculated from netted, validated and weighted risk factors.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The risk factors are left in the module store.
'
' BEHAVIOR
'   Resets the store, then reads the rows from the first data row to the first row whose
'   first two columns are blank. For each row:
'     - reads the risk class (column 2), bucket (3), qualifier (4) and sub-type (5), and
'       checks them with ValidateCommon
'     - rejects a sensitivity (column 7) that is not numeric and a tenor (column 6) that is
'       neither blank nor numeric
'     - per risk class: defaults and checks the sub-type, checks the tenor against the
'       regulatory vertices (GIRR yield, CSR and commodity) or sets it to 0 (GIRR inflation
'       and cross-currency basis, equity); FX takes no sub-type and uses the bucket as
'       qualifier
'     - rejects a row without a qualifier
'     - adds the valid row as a risk factor, which nets identical ones
'   After the rows, sets the weighted sensitivity of each risk factor to its delta risk
'   weight times its amount.
'
' ERROR POLICY
'   Invalid rows are logged as ERROR and skipped, never counted as zero. Errors from the
'   callees (for example a missing sheet or parameter) are not handled here.
'
' DEPENDENCIES
'   GetSheet, TableCell, IsBlank, IsNumber, IsBlankOrNumber, ToText, ToDouble,
'   BucketLabel, LogMessage (CORE_Util); IsValidGirrTenor, IsValidCsrTenor,
'   IsValidComTenor (CORE_Params); ResetStore, ValidateCommon, AddRiskFactor,
'   DeltaRiskWeight.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.LoadDelta"

    Dim InputSheet              As Object        'In_Delta sheet
    Dim RowIndex                As Long          'Row being read, then risk factor index
    Dim Source                  As String        'Row reference for the Log
    Dim RiskClass               As String        'Risk class of the row
    Dim Bucket                  As String        'Bucket of the row
    Dim Qualifier               As String        'Qualifier of the row
    Dim SubType                 As String        'Sub-type of the row
    Dim Tenor                   As Double        'Tenor of the row in years
    Dim RowIsValid              As Boolean       'True while the row is usable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Empty the risk factor store
        ResetStore
    'Locate the input sheet
        Set InputSheet = GetSheet(SH_IN_DELTA)
    'Start at the first data row
        RowIndex = LAYOUT_FIRST_DATA_ROW

'------------------------------------------------------------------------------
' READ AND VALIDATE ROWS
'------------------------------------------------------------------------------
    'Read and validate each row until the first row whose first two columns are blank
        Do While Not (IsBlank(TableCell(InputSheet, RowIndex, 1).Value) And _
            IsBlank(TableCell(InputSheet, RowIndex, 2).Value))
            Source = SH_IN_DELTA & " row " & RowIndex
            RiskClass = ToText(TableCell(InputSheet, RowIndex, 2).Value)
            Bucket = BucketLabel(TableCell(InputSheet, RowIndex, 3).Value)
            Qualifier = ToText(TableCell(InputSheet, RowIndex, 4).Value)
            SubType = ToText(TableCell(InputSheet, RowIndex, 5).Value)
            RowIsValid = ValidateCommon(RiskClass, Bucket, Source)
            If RowIsValid And Not IsNumber(TableCell(InputSheet, RowIndex, 7).Value) Then
                LogMessage "ERROR", Source, "Sensitivity is not numeric; row skipped"
                RowIsValid = False
            End If
            If RowIsValid And Not IsBlankOrNumber(TableCell(InputSheet, RowIndex, 6).Value) Then
                LogMessage "ERROR", Source, "Tenor is not numeric; row skipped"
                RowIsValid = False
            End If
            If RowIsValid Then
                Tenor = ToDouble(TableCell(InputSheet, RowIndex, 6).Value, 0#)
                Select Case RiskClass
                    Case RC_GIRR
                        If SubType = vbNullString Then SubType = "YIELD"
                        Select Case SubType
                            Case "YIELD"
                                If Not IsValidGirrTenor(Tenor) Then
                                    LogMessage "ERROR", Source, "GIRR tenor " & Tenor & " is not a regulatory vertex; row skipped"
                                    RowIsValid = False
                                End If
                            Case "INFLATION", "XCCY"
                                Tenor = 0#
                            Case Else
                                LogMessage "ERROR", Source, "GIRR sub-type must be YIELD, INFLATION or XCCY; row skipped"
                                RowIsValid = False
                        End Select
                    Case RC_CSR
                        If SubType = vbNullString Then SubType = "BOND"
                        If SubType <> "BOND" And SubType <> "CDS" Then
                            LogMessage "ERROR", Source, "CSR sub-type must be BOND or CDS; row skipped"
                            RowIsValid = False
                        ElseIf Not IsValidCsrTenor(Tenor) Then
                            LogMessage "ERROR", Source, "CSR tenor " & Tenor & " is not a regulatory vertex; row skipped"
                            RowIsValid = False
                        End If
                    Case RC_EQ
                        If SubType = vbNullString Then SubType = "SPOT"
                        If SubType <> "SPOT" And SubType <> "REPO" Then
                            LogMessage "ERROR", Source, "Equity sub-type must be SPOT or REPO; row skipped"
                            RowIsValid = False
                        End If
                        Tenor = 0#
                    Case RC_COM
                        If Not IsValidComTenor(Tenor) Then
                            LogMessage "ERROR", Source, "Commodity tenor " & Tenor & " is not a regulatory vertex; row skipped"
                            RowIsValid = False
                        End If
                    Case RC_FX
                        If SubType <> vbNullString Then
                            LogMessage "ERROR", Source, "FX takes no sub-type; row skipped"
                            RowIsValid = False
                        End If
                        Qualifier = Bucket
                        SubType = vbNullString
                        Tenor = 0#
                    Case Else
                        ' Any other value: nothing is done
                End Select
                If RowIsValid And Qualifier = vbNullString Then
                    LogMessage "ERROR", Source, "Qualifier (curve, issuer or commodity) is required; row skipped"
                    RowIsValid = False
                End If
            End If
            If RowIsValid Then
                AddRiskFactor RiskClass, Bucket, Qualifier, SubType, Tenor, 0#, _
                    ToDouble(TableCell(InputSheet, RowIndex, 7).Value, 0#), 0#
            End If
            RowIndex = RowIndex + 1
        Loop

'------------------------------------------------------------------------------
' WEIGHT SENSITIVITIES
'------------------------------------------------------------------------------
    'Weight each risk factor by its delta risk weight
        For RowIndex = 1 To mN
            mWS(RowIndex) = DeltaRiskWeight(RowIndex) * mA1(RowIndex)
        Next RowIndex

End Sub

Private Function DeltaRiskWeight(ByVal i As Long) As Double
'
'------------------------------------------------------------------------------
'                         RETURN THE DELTA RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the delta risk weight of one risk factor of the store.
'
' INPUTS
'   i
'     Index of the risk factor in the store
'
' RETURNS
'   The risk weight of the risk factor, from CORE_Params; 0 when its risk class is none of
'   the five handled.
'
' BEHAVIOR
'   Selects the risk weight function by risk class: GirrRiskWeight (bucket, sub-type and
'   tenor), CsrRiskWeight, EqRiskWeight (bucket and sub-type), ComRiskWeight, FxRiskWeight.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Pick the risk weight by risk class
        Select Case mRC(i)
            Case RC_GIRR
                DeltaRiskWeight = GirrRiskWeight(mBucket(i), mSub(i), mT1(i))
            Case RC_CSR
                DeltaRiskWeight = CsrRiskWeight(mBucket(i))
            Case RC_EQ
                DeltaRiskWeight = EqRiskWeight(mBucket(i), mSub(i))
            Case RC_COM
                DeltaRiskWeight = ComRiskWeight(mBucket(i))
            Case RC_FX
                DeltaRiskWeight = FxRiskWeight(mBucket(i))
            Case Else
                ' Any other value: nothing is done
        End Select

End Function

Private Function IsNameSubType( _
    ByVal rc As String, _
    ByVal subType As String, _
    ByVal src As String, _
    ByVal kind As String) As Boolean
'
'------------------------------------------------------------------------------
'                    VALIDATE A VEGA OR CURVATURE SUB-TYPE
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that the sub-type of a vega or curvature row is one the delta sub-types allow
'   for the same risk factor.
'
' WHY THIS EXISTS
'   Vega and curvature take the sub-types of the delta that are the same risk factor:
'   the bond or CDS curve of a CSR issuer, and the equity spot price. Equity repo
'   rates have no vega or curvature (MAR21.12); a commodity sub-type is the delivery
'   location and stays free text.
'
' INPUTS
'   rc
'     Risk class code
'
'   subType
'     Sub-type read from the row
'
'   src
'     Source reference for the Log message
'
'   kind
'     Component name used in the message (vega or curvature)
'
' RETURNS
'   True when the sub-type is accepted; False when the row is rejected.
'
' BEHAVIOR
'   For CSR accepts a blank sub-type, BOND or CDS. For equity accepts a blank sub-type or
'   SPOT. Any other risk class is accepted.
'
' ERROR POLICY
'   Does not raise. A rejection is logged as ERROR with the text row skipped.
'
' DEPENDENCIES
'   LogMessage (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Accept the sub-type until a check rejects it
        IsNameSubType = True

'------------------------------------------------------------------------------
' VALIDATE SUB-TYPE
'------------------------------------------------------------------------------
    'Check the sub-type by risk class
        Select Case rc
            Case RC_CSR
                If subType <> vbNullString And subType <> "BOND" And subType <> "CDS" Then
                    LogMessage "ERROR", src, "CSR " & kind & " sub-type must be BOND, CDS or none; row skipped"
                    IsNameSubType = False
                End If
            Case RC_EQ
                If subType <> vbNullString And subType <> "SPOT" Then
                    LogMessage "ERROR", src, "Equity " & kind & " takes sub-type SPOT or none: repo rates have no " & kind & "; row skipped"
                    IsNameSubType = False
                End If
            Case Else
                ' Any other value: nothing is done
        End Select

End Function

Private Sub LoadVega()
'
'------------------------------------------------------------------------------
'                              LOAD THE VEGA ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Loads the vega sensitivities of In_Vega into the risk factor store and computes their
'   weighted sensitivities.
'
' WHY THIS EXISTS
'   The vega capital is calculated from netted, validated and weighted risk factors.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The risk factors are left in the module store.
'
' BEHAVIOR
'   Resets the store, then reads the rows from the first data row to the first row whose
'   first two columns are blank. For each row:
'     - reads the risk class (column 2), bucket (3), qualifier (4) and sub-type (5), and
'       checks them with ValidateCommon
'     - rejects a vega sensitivity (column 8) that is not numeric, an option (6) or underlying
'       (7) maturity that is neither blank nor numeric, and an option maturity that is not
'       a regulatory vertex
'     - per risk class: GIRR yield keys the risk factor by option and underlying maturity
'       (checked against the vertices) and clears the qualifier; GIRR inflation and
'       cross-currency basis take the option maturity only; CSR, equity and commodity
'       check the sub-type with IsNameSubType and drop it; FX takes no sub-type and uses
'       the bucket as qualifier
'     - rejects a row without a qualifier, except for GIRR
'     - adds the valid row as a risk factor, which nets identical ones
'   After the rows, sets the weighted sensitivity of each risk factor to its vega risk
'   weight times its amount.
'
' ERROR POLICY
'   Invalid rows are logged as ERROR and skipped, never counted as zero. Errors from the
'   callees (for example a missing sheet or parameter) are not handled here.
'
' DEPENDENCIES
'   GetSheet, TableCell, IsBlank, IsNumber, IsBlankOrNumber, ToText, ToDouble,
'   BucketLabel, LogMessage (CORE_Util); IsValidVegaMaturity, VegaRiskWeight
'   (CORE_Params); ResetStore, ValidateCommon, IsNameSubType, AddRiskFactor.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.LoadVega"

    Dim InputSheet              As Object        'In_Vega sheet
    Dim RowIndex                As Long          'Row being read, then risk factor index
    Dim Source                  As String        'Row reference for the Log
    Dim RiskClass               As String        'Risk class of the row
    Dim Bucket                  As String        'Bucket of the row
    Dim Qualifier               As String        'Qualifier of the row
    Dim SubType                 As String        'Sub-type of the row
    Dim OptionMaturity          As Double        'Option maturity of the row in years
    Dim UnderlyingMaturity      As Double        'Underlying maturity of the row in years
    Dim RowIsValid              As Boolean       'True while the row is usable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Empty the risk factor store
        ResetStore
    'Locate the input sheet
        Set InputSheet = GetSheet(SH_IN_VEGA)
    'Start at the first data row
        RowIndex = LAYOUT_FIRST_DATA_ROW

'------------------------------------------------------------------------------
' READ AND VALIDATE ROWS
'------------------------------------------------------------------------------
    'Read and validate each row until the first row whose first two columns are blank
        Do While Not (IsBlank(TableCell(InputSheet, RowIndex, 1).Value) And _
            IsBlank(TableCell(InputSheet, RowIndex, 2).Value))
            Source = SH_IN_VEGA & " row " & RowIndex
            RiskClass = ToText(TableCell(InputSheet, RowIndex, 2).Value)
            Bucket = BucketLabel(TableCell(InputSheet, RowIndex, 3).Value)
            Qualifier = ToText(TableCell(InputSheet, RowIndex, 4).Value)
            SubType = ToText(TableCell(InputSheet, RowIndex, 5).Value)
            RowIsValid = ValidateCommon(RiskClass, Bucket, Source)
            If RowIsValid And Not IsNumber(TableCell(InputSheet, RowIndex, 8).Value) Then
                LogMessage "ERROR", Source, "Vega sensitivity is not numeric; row skipped"
                RowIsValid = False
            End If
            If RowIsValid And Not (IsBlankOrNumber(TableCell(InputSheet, RowIndex, 6).Value) And _
                IsBlankOrNumber(TableCell(InputSheet, RowIndex, 7).Value)) Then
                LogMessage "ERROR", Source, "Option or underlying maturity is not numeric; row skipped"
                RowIsValid = False
            End If
            If RowIsValid Then
                OptionMaturity = ToDouble(TableCell(InputSheet, RowIndex, 6).Value, 0#)
                UnderlyingMaturity = ToDouble(TableCell(InputSheet, RowIndex, 7).Value, 0#)
            End If
            If RowIsValid And Not IsValidVegaMaturity(OptionMaturity) Then
                LogMessage "ERROR", Source, "Option maturity " & OptionMaturity & " is not a regulatory vertex; row skipped"
                RowIsValid = False
            End If
            If RowIsValid Then
                Select Case RiskClass
                    Case RC_GIRR
                        If SubType = vbNullString Then SubType = "YIELD"
                        Select Case SubType
                            Case "YIELD"
                                ' One vega surface per currency: option x underlying maturity
                                If Not IsValidVegaMaturity(UnderlyingMaturity) Then
                                    LogMessage "ERROR", Source, "GIRR underlying maturity " & UnderlyingMaturity & " is not a regulatory vertex; row skipped"
                                    RowIsValid = False
                                End If
                                Qualifier = vbNullString
                            Case "INFLATION", "XCCY"
                                ' Inflation and cross-currency basis vega: the option
                                ' maturity only, one risk factor per curve (MAR21.8 FAQ4)
                                If Not IsBlank(TableCell(InputSheet, RowIndex, 7).Value) Then
                                    LogMessage "ERROR", Source, "GIRR " & SubType & " vega takes no underlying maturity; row skipped"
                                    RowIsValid = False
                                End If
                                UnderlyingMaturity = 0#
                            Case Else
                                LogMessage "ERROR", Source, "GIRR sub-type must be YIELD, INFLATION or XCCY; row skipped"
                                RowIsValid = False
                        End Select
                    Case RC_CSR, RC_EQ, RC_COM
                        ' Vega risk factors are the underlying name and the option
                        ' maturity: no bond/CDS basis or delivery location (MAR21.94 FAQ)
                        If Not IsNameSubType(RiskClass, SubType, Source, "vega") Then RowIsValid = False
                        SubType = vbNullString
                        UnderlyingMaturity = 0#
                    Case RC_FX
                        If SubType <> vbNullString Then
                            LogMessage "ERROR", Source, "FX takes no sub-type; row skipped"
                            RowIsValid = False
                        End If
                        Qualifier = Bucket
                        SubType = vbNullString
                        UnderlyingMaturity = 0#
                    Case Else
                        ' Any other value: nothing is done
                End Select
                If RowIsValid And Qualifier = vbNullString And RiskClass <> RC_GIRR Then
                    LogMessage "ERROR", Source, "Qualifier (issuer or commodity) is required; row skipped"
                    RowIsValid = False
                End If
            End If
            If RowIsValid Then
                AddRiskFactor RiskClass, Bucket, Qualifier, SubType, OptionMaturity, UnderlyingMaturity, _
                    ToDouble(TableCell(InputSheet, RowIndex, 8).Value, 0#), 0#
            End If
            RowIndex = RowIndex + 1
        Loop

'------------------------------------------------------------------------------
' WEIGHT SENSITIVITIES
'------------------------------------------------------------------------------
    'Weight each risk factor by its vega risk weight
        For RowIndex = 1 To mN
            mWS(RowIndex) = VegaRiskWeight(mRC(RowIndex), mBucket(RowIndex)) * mA1(RowIndex)
        Next RowIndex

End Sub

Private Sub LoadCurvature()
'
'------------------------------------------------------------------------------
'                           LOAD THE CURVATURE ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Loads the curvature risk positions (CVR up and CVR down) of In_Curvature into the risk
'   factor store.
'
' WHY THIS EXISTS
'   The curvature capital is calculated from netted and validated risk factors.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. The risk factors are left in the module store.
'
' BEHAVIOR
'   Resets the store, then reads the rows from the first data row to the first row whose
'   first two columns are blank. For each row:
'     - reads the risk class (column 2), bucket (3), qualifier (4) and sub-type (5), and
'       checks them with ValidateCommon
'     - rejects CVR up (column 6) or CVR down (7) that is not numeric, and a reporting
'       currency flag (8) that is not TRUE, FALSE or blank
'     - per risk class: GIRR takes sub-type YIELD or none and one risk factor per bucket;
'       CSR, equity and commodity check the sub-type with IsNameSubType and drop it; FX
'       takes no sub-type, uses the bucket as qualifier and divides both amounts by
'       FxCurvatureDivisor
'     - rejects a row without a qualifier, except for GIRR
'     - adds the valid row as a risk factor, which nets identical ones
'
' ERROR POLICY
'   Invalid rows are logged as ERROR and skipped, never counted as zero. Errors from the
'   callees (for example a missing sheet or parameter) are not handled here.
'
' DEPENDENCIES
'   GetSheet, TableCell, IsBlank, IsNumber, IsBlankOrBoolCell, CellBool, ToText, ToDouble,
'   BucketLabel, LogMessage (CORE_Util); FxCurvatureDivisor (CORE_Params); ResetStore,
'   ValidateCommon, IsNameSubType, AddRiskFactor.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.LoadCurvature"

    Dim InputSheet              As Object        'In_Curvature sheet
    Dim RowIndex                As Long          'Row being read
    Dim Source                  As String        'Row reference for the Log
    Dim RiskClass               As String        'Risk class of the row
    Dim Bucket                  As String        'Bucket of the row
    Dim Qualifier               As String        'Qualifier of the row
    Dim SubType                 As String        'Sub-type of the row
    Dim CvrUp                   As Double        'Curvature risk position, up shock
    Dim CvrDown                 As Double        'Curvature risk position, down shock
    Dim RowIsValid              As Boolean       'True while the row is usable
    Dim RefersToReportingCcy    As Boolean       'Reporting currency flag of the row

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Empty the risk factor store
        ResetStore
    'Locate the input sheet
        Set InputSheet = GetSheet(SH_IN_CURV)
    'Start at the first data row
        RowIndex = LAYOUT_FIRST_DATA_ROW

'------------------------------------------------------------------------------
' READ AND VALIDATE ROWS
'------------------------------------------------------------------------------
    'Read and validate each row until the first row whose first two columns are blank
        Do While Not (IsBlank(TableCell(InputSheet, RowIndex, 1).Value) And _
            IsBlank(TableCell(InputSheet, RowIndex, 2).Value))
            Source = SH_IN_CURV & " row " & RowIndex
            RiskClass = ToText(TableCell(InputSheet, RowIndex, 2).Value)
            Bucket = BucketLabel(TableCell(InputSheet, RowIndex, 3).Value)
            Qualifier = ToText(TableCell(InputSheet, RowIndex, 4).Value)
            SubType = ToText(TableCell(InputSheet, RowIndex, 5).Value)
            RowIsValid = ValidateCommon(RiskClass, Bucket, Source)
            If RowIsValid And (Not IsNumber(TableCell(InputSheet, RowIndex, 6).Value) Or _
                Not IsNumber(TableCell(InputSheet, RowIndex, 7).Value)) Then
                LogMessage "ERROR", Source, "CVR up and CVR down must both be numeric; row skipped"
                RowIsValid = False
            End If
            If RowIsValid And Not IsBlankOrBoolCell(TableCell(InputSheet, RowIndex, 8)) Then
                LogMessage "ERROR", Source, "References reporting ccy (RefReportingCcy) must be TRUE, FALSE or blank; row skipped"
                RowIsValid = False
            End If
            If RowIsValid Then
                CvrUp = ToDouble(TableCell(InputSheet, RowIndex, 6).Value, 0#)
                CvrDown = ToDouble(TableCell(InputSheet, RowIndex, 7).Value, 0#)
                RefersToReportingCcy = CellBool(TableCell(InputSheet, RowIndex, 8))
                Select Case RiskClass
                    Case RC_GIRR
                        ' All curves of a currency are shifted together: one risk factor
                        ' per bucket, with no inflation or cross-currency basis curvature (MAR21.8)
                        If SubType <> vbNullString And SubType <> "YIELD" Then
                            LogMessage "ERROR", Source, "GIRR curvature takes sub-type YIELD or none: inflation and cross-currency basis have no curvature; row skipped"
                            RowIsValid = False
                        End If
                        Qualifier = vbNullString
                        SubType = vbNullString
                    Case RC_CSR, RC_EQ, RC_COM
                        ' One curvature risk factor per issuer or commodity: bond and
                        ' CDS curves, and delivery locations, are not separate (MAR21.8-MAR21.14)
                        If Not IsNameSubType(RiskClass, SubType, Source, "curvature") Then RowIsValid = False
                        SubType = vbNullString
                    Case RC_FX
                        If SubType <> vbNullString Then
                            LogMessage "ERROR", Source, "FX takes no sub-type; row skipped"
                            RowIsValid = False
                        End If
                        Qualifier = Bucket
                        SubType = vbNullString
                        CvrUp = CvrUp / FxCurvatureDivisor(RefersToReportingCcy)
                        CvrDown = CvrDown / FxCurvatureDivisor(RefersToReportingCcy)
                    Case Else
                        ' Any other value: nothing is done
                End Select
                If RowIsValid And Qualifier = vbNullString And RiskClass <> RC_GIRR Then
                    LogMessage "ERROR", Source, "Qualifier (issuer or commodity) is required; row skipped"
                    RowIsValid = False
                End If
            End If
            If RowIsValid Then AddRiskFactor RiskClass, Bucket, Qualifier, SubType, 0#, 0#, CvrUp, CvrDown
            RowIndex = RowIndex + 1
        Loop

End Sub

'
'------------------------------------------------------------------------------
'
'                         CORRELATIONS WITHIN A BUCKET
'
'------------------------------------------------------------------------------
'

Private Function DeltaCorrelation( _
    ByVal rc As String, _
    ByVal bucket As String, _
    ByVal i As Long, _
    ByVal j As Long) As Double
'
'------------------------------------------------------------------------------
'                         RETURN THE DELTA CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the delta correlation of two risk factors of the same bucket, for the medium
'   scenario and before the scenario transform.
'
' WHY THIS EXISTS
'   The bucket aggregation of the delta component needs the correlation of every pair
'   of risk factors.
'
' INPUTS
'   rc
'     Risk class code
'
'   bucket
'     Bucket label
'
'   i, j
'     Indexes of the two risk factors in the store
'
' RETURNS
'   The correlation of the pair, before the scenario transform.
'
' BEHAVIOR
'   By risk class:
'     - GIRR: GirrRhoXccy when either is cross-currency basis; GirrRhoDiffCurve when both
'       are inflation; GirrRhoInflation when one is inflation; otherwise the tenor
'       correlation, times GirrRhoDiffCurve when the curves differ
'     - CSR: CsrIssuerCorrelation, times CsrRhoTenor when the tenors differ
'     - equity: EqIssuerCorrelation, times EqRhoSpotRepo when the sub-types differ
'     - commodity: ComUnderlyingCorrelation, times ComRhoTenor when the tenors differ
'     - any other risk class: 1
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Params correlation accessors (Girr*, Csr*, Eq*, Com*); CsrIssuerCorrelation,
'   EqIssuerCorrelation, ComUnderlyingCorrelation.
'
' NOTES
'   Tenors count as different when they differ by more than 0.000001.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.DeltaCorrelation"

    Dim Correlation             As Double        'Correlation of the pair

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Select the correlation by risk class
        Select Case rc
            Case RC_GIRR
                If mSub(i) = "XCCY" Or mSub(j) = "XCCY" Then
                    Correlation = GirrRhoXccy()
                ElseIf mSub(i) = "INFLATION" And mSub(j) = "INFLATION" Then
                    Correlation = GirrRhoDiffCurve()
                ElseIf mSub(i) = "INFLATION" Or mSub(j) = "INFLATION" Then
                    Correlation = GirrRhoInflation()
                Else
                    Correlation = GirrTenorCorrelation(mT1(i), mT1(j))
                    If mQual(i) <> mQual(j) Then Correlation = Correlation * GirrRhoDiffCurve()
                End If
            Case RC_CSR
                Correlation = CsrIssuerCorrelation(bucket, i, j)
                If Abs(mT1(i) - mT1(j)) > 0.000001 Then Correlation = Correlation * CsrRhoTenor()
            Case RC_EQ
                Correlation = EqIssuerCorrelation(bucket, i, j)
                If mSub(i) <> mSub(j) Then Correlation = Correlation * EqRhoSpotRepo()
            Case RC_COM
                Correlation = ComUnderlyingCorrelation(bucket, i, j)
                If Abs(mT1(i) - mT1(j)) > 0.000001 Then Correlation = Correlation * ComRhoTenor()
            Case Else
                Correlation = 1#
        End Select

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the correlation
        DeltaCorrelation = Correlation

End Function

Private Function CsrIssuerCorrelation( _
    ByVal bucket As String, _
    ByVal i As Long, _
    ByVal j As Long) As Double
'
'------------------------------------------------------------------------------
'                      RETURN THE CSR ISSUER CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the name by basis correlation of two CSR risk factors, tenor excluded.
'
' INPUTS
'   bucket
'     Bucket label
'
'   i, j
'     Indexes of the two risk factors in the store
'
' RETURNS
'   1 for the same qualifier, otherwise CsrNameCorrelation; times CsrRhoBasis when the
'   sub-types differ.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.CsrIssuerCorrelation"

    Dim Correlation             As Double        'Correlation of the pair

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Use 1 for the same qualifier, else the name correlation of the bucket
        If mQual(i) = mQual(j) Then
            Correlation = 1#
        Else
            Correlation = CsrNameCorrelation(bucket)
        End If
    'Apply the basis correlation when the sub-types differ
        If mSub(i) <> mSub(j) Then Correlation = Correlation * CsrRhoBasis()

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the correlation
        CsrIssuerCorrelation = Correlation

End Function

Private Function EqIssuerCorrelation( _
    ByVal bucket As String, _
    ByVal i As Long, _
    ByVal j As Long) As Double
'
'------------------------------------------------------------------------------
'                     RETURN THE EQUITY ISSUER CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the issuer correlation of two equity risk factors.
'
' INPUTS
'   bucket
'     Bucket label
'
'   i, j
'     Indexes of the two risk factors in the store
'
' RETURNS
'   1 for the same qualifier, otherwise EqNameCorrelation.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return 1 for the same qualifier, else the name correlation of the bucket
        If mQual(i) = mQual(j) Then
            EqIssuerCorrelation = 1#
        Else
            EqIssuerCorrelation = EqNameCorrelation(bucket)
        End If

End Function

Private Function ComUnderlyingCorrelation( _
    ByVal bucket As String, _
    ByVal i As Long, _
    ByVal j As Long) As Double
'
'------------------------------------------------------------------------------
'                       RETURN THE COMMODITY CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the commodity by delivery location correlation of two commodity risk factors,
'   tenor excluded.
'
' INPUTS
'   bucket
'     Bucket label
'
'   i, j
'     Indexes of the two risk factors in the store
'
' RETURNS
'   1 for the same qualifier, otherwise ComCommodityCorrelation; times ComRhoBasis when the
'   sub-types differ.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.ComUnderlyingCorrelation"

    Dim Correlation             As Double        'Correlation of the pair

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Use 1 for the same qualifier, else the commodity correlation of the bucket
        If mQual(i) = mQual(j) Then
            Correlation = 1#
        Else
            Correlation = ComCommodityCorrelation(bucket)
        End If
    'Apply the basis correlation when the sub-types differ
        If mSub(i) <> mSub(j) Then Correlation = Correlation * ComRhoBasis()

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the correlation
        ComUnderlyingCorrelation = Correlation

End Function

Private Function MaturityCorrelation(ByVal ta As Double, ByVal tb As Double) As Double
'
'------------------------------------------------------------------------------
'                       RETURN THE MATURITY CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the correlation of two maturities.
'
' INPUTS
'   ta, tb
'     The two maturities
'
' RETURNS
'   1 when the maturities differ by less than 0.000001, otherwise
'   Exp(-VegaAlpha() * Abs(ta - tb) / MinD(ta, tb)).
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the maturity correlation
        If Abs(ta - tb) < 0.000001 Then
            MaturityCorrelation = 1#
        Else
            MaturityCorrelation = Exp(-VegaAlpha() * Abs(ta - tb) / MinD(ta, tb))
        End If

End Function

Private Function UnderlyingCorrelation( _
    ByVal rc As String, _
    ByVal bucket As String, _
    ByVal i As Long, _
    ByVal j As Long) As Double
'
'------------------------------------------------------------------------------
'                      RETURN THE UNDERLYING CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the delta correlation restricted to the dimension that vega and curvature risk
'   factors share with delta: the name for CSR and equity, the commodity for commodity
'   (MAR21.94 and MAR21.100 with their FAQs).
'
' INPUTS
'   rc
'     Risk class code
'
'   bucket
'     Bucket label
'
'   i, j
'     Indexes of the two risk factors in the store
'
' RETURNS
'   1 for the same qualifier. Otherwise the name correlation of the bucket for CSR and
'   equity, the commodity correlation for commodity, and 1 for any other risk class.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return 1 for the same qualifier, else the correlation of the risk class
        If mQual(i) = mQual(j) Then
            UnderlyingCorrelation = 1#
        Else
            Select Case rc
                Case RC_CSR
                    UnderlyingCorrelation = CsrNameCorrelation(bucket)
                Case RC_EQ
                    UnderlyingCorrelation = EqNameCorrelation(bucket)
                Case RC_COM
                    UnderlyingCorrelation = ComCommodityCorrelation(bucket)
                Case Else
                    UnderlyingCorrelation = 1#
            End Select
        End If

End Function

Private Function GirrVegaUnderlyingCorrelation(ByVal i As Long, ByVal j As Long) As Double
'
'------------------------------------------------------------------------------
'                 RETURN THE GIRR VEGA UNDERLYING CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the delta correlation of the GIRR delta risk factors two vega risk factors
'   correspond to (MAR21.94, MAR21.46-49).
'
' INPUTS
'   i, j
'     Indexes of the two risk factors in the store
'
' RETURNS
'   The underlying maturity correlation between two yield surfaces; otherwise 1 for the same
'   sub-type and qualifier, GirrRhoXccy when either is cross-currency basis,
'   GirrRhoDiffCurve when both are inflation, and GirrRhoInflation in the remaining case.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Select the correlation by the sub-types and qualifiers of the pair
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

Private Function VegaCorrelation( _
    ByVal rc As String, _
    ByVal bucket As String, _
    ByVal i As Long, _
    ByVal j As Long) As Double
'
'------------------------------------------------------------------------------
'                         RETURN THE VEGA CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the vega correlation of two risk factors (MAR21.93-95): min(rho_option_maturity
'   x rho_underlying, 1).
'
' INPUTS
'   rc
'     Risk class code
'
'   bucket
'     Bucket label
'
'   i, j
'     Indexes of the two risk factors in the store
'
' RETURNS
'   The correlation, before the scenario transform.
'
' BEHAVIOR
'   Takes the underlying correlation from GirrVegaUnderlyingCorrelation for GIRR and from
'   UnderlyingCorrelation otherwise, multiplies it by the correlation of the option
'   maturities and caps the product at 1.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.VegaCorrelation"

    Dim UnderlyingRho           As Double        'Correlation of the underlyings

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Resolve the correlation of the underlyings
        If rc = RC_GIRR Then
            UnderlyingRho = GirrVegaUnderlyingCorrelation(i, j)
        Else
            UnderlyingRho = UnderlyingCorrelation(rc, bucket, i, j)
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Multiply by the option maturity correlation and cap at 1
        VegaCorrelation = MinD(MaturityCorrelation(mT1(i), mT1(j)) * UnderlyingRho, 1#)

End Function

Private Function CurvatureCorrelation( _
    ByVal rc As String, _
    ByVal bucket As String, _
    ByVal i As Long, _
    ByVal j As Long) As Double
'
'------------------------------------------------------------------------------
'                       RETURN THE CURVATURE CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the curvature correlation of two risk factors (MAR21.100): the square of the
'   underlying delta correlation.
'
' INPUTS
'   rc
'     Risk class code
'
'   bucket
'     Bucket label
'
'   i, j
'     Indexes of the two risk factors in the store
'
' RETURNS
'   The correlation, before the scenario transform.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.CurvatureCorrelation"

    Dim UnderlyingRho           As Double        'Delta correlation of the underlyings

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Resolve the correlation of the underlyings
        UnderlyingRho = UnderlyingCorrelation(rc, bucket, i, j)

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return its square
        CurvatureCorrelation = UnderlyingRho * UnderlyingRho

End Function

Private Function PairCorrelation( _
    ByVal comp As Long, _
    ByVal rc As String, _
    ByVal bucket As String, _
    ByVal i As Long, _
    ByVal j As Long, _
    ByVal scen As Long) As Double
'
'------------------------------------------------------------------------------
'                         RETURN THE PAIR CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the correlation of two risk factors of a bucket for a component and a scenario.
'
' INPUTS
'   comp
'     Component: COMP_DELTA, COMP_VEGA or COMP_CURV
'
'   rc
'     Risk class code
'
'   bucket
'     Bucket label
'
'   i, j
'     Indexes of the two risk factors in the store
'
'   scen
'     Scenario: SCEN_LOW, SCEN_MEDIUM or SCEN_HIGH
'
' RETURNS
'   The correlation of the component, transformed for the scenario.
'
' BEHAVIOR
'   Takes DeltaCorrelation, VegaCorrelation or CurvatureCorrelation by component, then
'   applies ScenarioCorrelation.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.PairCorrelation"

    Dim Correlation             As Double        'Correlation of the pair, medium scenario

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Select the correlation by component
        Select Case comp
            Case COMP_DELTA
                Correlation = DeltaCorrelation(rc, bucket, i, j)
            Case COMP_VEGA
                Correlation = VegaCorrelation(rc, bucket, i, j)
            Case Else
                Correlation = CurvatureCorrelation(rc, bucket, i, j)
        End Select

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Transform the correlation for the scenario
        PairCorrelation = ScenarioCorrelation(Correlation, scen)

End Function

Private Function CrossBucketCorrelation( _
    ByVal comp As Long, _
    ByVal rc As String, _
    ByVal b As String, _
    ByVal c As String, _
    ByVal scen As Long) As Double
'
'------------------------------------------------------------------------------
'                     RETURN THE CROSS-BUCKET CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the correlation between two buckets of a risk class for a component and a
'   scenario.
'
' INPUTS
'   comp
'     Component: COMP_DELTA, COMP_VEGA or COMP_CURV
'
'   rc
'     Risk class code
'
'   b, c
'     The two bucket labels
'
'   scen
'     Scenario: SCEN_LOW, SCEN_MEDIUM or SCEN_HIGH
'
' RETURNS
'   The bucket gamma, squared for curvature, transformed for the scenario.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.CrossBucketCorrelation"

    Dim Gamma                   As Double        'Gamma between the two buckets

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Read the gamma between the two buckets
        Gamma = BucketGamma(rc, b, c)
    'Square it for curvature
        If comp = COMP_CURV Then Gamma = Gamma * Gamma

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Transform the correlation for the scenario
        CrossBucketCorrelation = ScenarioCorrelation(Gamma, scen)

End Function

Private Function Psi(ByVal x As Double, ByVal y As Double) As Double
'
'------------------------------------------------------------------------------
'                       RETURN THE CURVATURE SIGN FACTOR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 0 when both values are negative, otherwise 1.
'
' INPUTS
'   x, y
'     The two values
'
' RETURNS
'   0 when x and y are both below zero, otherwise 1.
'
' NOTES
'   Used in the curvature terms, which are cut when both amounts are negative.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return 0 when both values are negative, else 1
        If x < 0# And y < 0# Then
            Psi = 0#
        Else
            Psi = 1#
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                                 AGGREGATION
'
'------------------------------------------------------------------------------
'

Private Sub ComputeComponent(ByVal comp As Long)
'
'------------------------------------------------------------------------------
'                             COMPUTE A COMPONENT
'------------------------------------------------------------------------------
' PURPOSE
'   Computes one component for every risk class.
'
' INPUTS
'   comp
'     Component: COMP_DELTA, COMP_VEGA or COMP_CURV
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls ComputeRiskClass for each risk class, using the risk factors in the store.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.ComputeComponent"

    Dim ClassIndex              As Long          'Risk class being computed

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Compute each risk class
        For ClassIndex = 1 To RC_COUNT
            ComputeRiskClass comp, ClassIndex
        Next ClassIndex

End Sub

Private Sub ComputeRiskClass(ByVal comp As Long, ByVal rcIdx As Long)
'
'------------------------------------------------------------------------------
'                             COMPUTE A RISK CLASS
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the capital of one risk class for one component under the three correlation
'   scenarios: Kb and Sb per bucket, then the aggregation across buckets, and writes one
'   bucket row per bucket and scenario.
'
' WHY THIS EXISTS
'   Each risk class and component needs its bucket charges and its cross-bucket
'   aggregation for the three correlation scenarios.
'
' INPUTS
'   comp
'     Component: COMP_DELTA, COMP_VEGA or COMP_CURV
'
'   rcIdx
'     Risk class index (1 to RC_COUNT)
'
' RETURNS
'   Nothing. The capital, the alternative-Sb flags and the bucket count are stored in the
'   module results and the bucket rows are written to Out_Buckets.
'
' BEHAVIOR
'   In order:
'     - lists the distinct buckets of the risk class among the stored risk factors and
'       sorts them; returns when there are none
'     - lists the risk factors of each bucket and flags the other-sector buckets
'       (IsOtherBucket)
'     - for each scenario and bucket: delta and vega take Sb as the sum of the weighted
'       sensitivities and Kb from the correlated sum of squares (the sum of the absolute
'       weighted sensitivities for an other bucket); curvature computes K for the up and
'       the down amounts, keeps the larger (ties broken by the larger sum, MAR21.5(3)) and
'       takes the matching sum as Sb
'     - aggregates across buckets with AcrossBuckets; when the sum is negative for delta
'       or vega, repeats it with the alternative Sb = max(min(sum WS, Kb), -Kb)
'       (MAR21.4(5)(b))
'     - stores the capital: the square root of the aggregate, floored at 0, plus the Kb of
'       the other buckets
'     - writes one Out_Buckets row per bucket with Kb, Sb, the number of risk factors and
'       a note (UP, DOWN, OTHER, ALT_SB)
'
' ERROR POLICY
'   Does not handle errors; errors from the callees (for example a missing parameter)
'   are not caught here.
'
' DEPENDENCIES
'   RiskClassCode (CORE_Config); SortLabels, MinD, MaxD, NearlyEqual (CORE_Util);
'   IsOtherBucket (CORE_Params); PairCorrelation, AcrossBuckets, WriteBucketRow.
'
' NOTES
'   Kb and Sb follow the notation of MAR21.4.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.ComputeRiskClass"

    'Risk class and buckets
    Dim RiskClass               As String        'Risk class code
    Dim BucketCount             As Long          'Number of buckets
    Dim Labels()                As String        'Sorted bucket labels
    Dim Found                   As Boolean       'True when a bucket label is already listed
    Dim BucketIndex             As Long          'Bucket being processed

    'Risk factors per bucket
    Dim FactorIndex             As Long          'Risk factor being read
    Dim OtherFactorIndex        As Long          'Second risk factor of a pair
    Dim MemberIndex             As Long          'Position in the member list of the bucket
    Dim OtherMemberIndex        As Long          'Second member position of a pair
    Dim MemberCounts()          As Long          'Number of risk factors in each bucket
    Dim BucketMembers()         As Long          'Risk factor indexes of each bucket
    Dim MaxMemberCount          As Long          'Largest number of risk factors in a bucket
    Dim BucketIsOther()         As Boolean       'True for an other-sector bucket

    'Bucket results
    Dim BucketKb()              As Double        'Kb of each bucket
    Dim BucketSb()              As Double        'Sb of each bucket
    Dim BucketNote()            As String        'Note of each bucket
    Dim KbUp                    As Double        'Curvature K, up amounts
    Dim KbDown                  As Double        'Curvature K, down amounts
    Dim SbUp                    As Double        'Curvature sum, up amounts
    Dim SbDown                  As Double        'Curvature sum, down amounts
    Dim Correlation             As Double        'Correlation of a pair of risk factors

    'Scenario results
    Dim ScenarioIndex           As Long          'Scenario being computed
    Dim SumSquares              As Double        'Aggregate under the square root
    Dim RiskClassCapital        As Double        'Capital of the risk class
    Dim OtherBucketSum          As Double        'Sum of Kb over the other buckets
    Dim AltSb                   As Double        'Alternative Sb of a bucket
    Dim AltUsed                 As Boolean       'True when the alternative Sb was used

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the code of the risk class
        RiskClass = RiskClassCode(rcIdx)

'------------------------------------------------------------------------------
' LIST BUCKETS
'------------------------------------------------------------------------------
    'Collect the distinct buckets of this risk class
        BucketCount = 0
        ReDim Labels(1 To 1)
        For FactorIndex = 1 To mN
            If mRC(FactorIndex) = RiskClass Then
                Found = False
                For BucketIndex = 1 To BucketCount
                    If Labels(BucketIndex) = mBucket(FactorIndex) Then
                        Found = True
                        Exit For
                    End If
                Next BucketIndex
                If Not Found Then
                    BucketCount = BucketCount + 1
                    ReDim Preserve Labels(1 To BucketCount)
                    Labels(BucketCount) = mBucket(FactorIndex)
                End If
            End If
        Next FactorIndex
    'Stop when the risk class has no risk factors
        If BucketCount = 0 Then Exit Sub
    'Sort the labels and store the bucket count
        SortLabels Labels, BucketCount
        mBucketCount(rcIdx, comp) = BucketCount

'------------------------------------------------------------------------------
' INDEX BUCKET MEMBERS
'------------------------------------------------------------------------------
    'Count the risk factors of each bucket
        ReDim MemberCounts(1 To BucketCount)
        MaxMemberCount = 0
        For BucketIndex = 1 To BucketCount
            For FactorIndex = 1 To mN
                If mRC(FactorIndex) = RiskClass And mBucket(FactorIndex) = Labels(BucketIndex) Then
                    MemberCounts(BucketIndex) = MemberCounts(BucketIndex) + 1
                End If
            Next FactorIndex
            If MemberCounts(BucketIndex) > MaxMemberCount Then MaxMemberCount = MemberCounts(BucketIndex)
        Next BucketIndex
    'List the risk factors of each bucket and flag the other buckets
        ReDim BucketMembers(1 To BucketCount, 1 To MaxMemberCount)
        ReDim BucketIsOther(1 To BucketCount)
        For BucketIndex = 1 To BucketCount
            MemberIndex = 0
            For FactorIndex = 1 To mN
                If mRC(FactorIndex) = RiskClass And mBucket(FactorIndex) = Labels(BucketIndex) Then
                    MemberIndex = MemberIndex + 1
                    BucketMembers(BucketIndex, MemberIndex) = FactorIndex
                End If
            Next FactorIndex
            BucketIsOther(BucketIndex) = IsOtherBucket(RiskClass, Labels(BucketIndex))
        Next BucketIndex
    'Allocate the Kb, Sb and note arrays
        ReDim BucketKb(1 To BucketCount)
        ReDim BucketSb(1 To BucketCount)
        ReDim BucketNote(1 To BucketCount)

'------------------------------------------------------------------------------
' COMPUTE SCENARIOS
'------------------------------------------------------------------------------
    'For each scenario: Kb and Sb per bucket, then the capital across buckets
        For ScenarioIndex = 1 To SCEN_COUNT
            ' Bucket level
            For BucketIndex = 1 To BucketCount
                BucketNote(BucketIndex) = vbNullString
                If comp = COMP_CURV Then
                    KbUp = 0#
                    KbDown = 0#
                    SbUp = 0#
                    SbDown = 0#
                    For MemberIndex = 1 To MemberCounts(BucketIndex)
                        FactorIndex = BucketMembers(BucketIndex, MemberIndex)
                        SbUp = SbUp + mA1(FactorIndex)
                        SbDown = SbDown + mA2(FactorIndex)
                        If BucketIsOther(BucketIndex) Then
                            KbUp = KbUp + MaxD(mA1(FactorIndex), 0#)
                            KbDown = KbDown + MaxD(mA2(FactorIndex), 0#)
                        Else
                            KbUp = KbUp + MaxD(mA1(FactorIndex), 0#) ^ 2
                            KbDown = KbDown + MaxD(mA2(FactorIndex), 0#) ^ 2
                            For OtherMemberIndex = 1 To MemberCounts(BucketIndex)
                                If OtherMemberIndex <> MemberIndex Then
                                    OtherFactorIndex = BucketMembers(BucketIndex, OtherMemberIndex)
                                    Correlation = PairCorrelation(comp, RiskClass, Labels(BucketIndex), FactorIndex, _
                                        OtherFactorIndex, ScenarioIndex)
                                    KbUp = KbUp + Correlation * mA1(FactorIndex) * mA1(OtherFactorIndex) * _
                                        Psi(mA1(FactorIndex), mA1(OtherFactorIndex))
                                    KbDown = KbDown + Correlation * mA2(FactorIndex) * mA2(OtherFactorIndex) * _
                                        Psi(mA2(FactorIndex), mA2(OtherFactorIndex))
                                End If
                            Next OtherMemberIndex
                        End If
                    Next MemberIndex
                    If Not BucketIsOther(BucketIndex) Then
                        KbUp = Sqr(MaxD(KbUp, 0#))
                        KbDown = Sqr(MaxD(KbDown, 0#))
                    End If
                    ' Scenario selection (MAR21.5(3)): larger K, ties broken by larger sum
                    If NearlyEqual(KbUp, KbDown) Then
                        If SbUp > SbDown Then
                            BucketNote(BucketIndex) = "UP"
                        Else
                            BucketNote(BucketIndex) = "DOWN"
                        End If
                    ElseIf KbUp > KbDown Then
                        BucketNote(BucketIndex) = "UP"
                    Else
                        BucketNote(BucketIndex) = "DOWN"
                    End If
                    If BucketNote(BucketIndex) = "UP" Then
                        BucketKb(BucketIndex) = KbUp
                        BucketSb(BucketIndex) = SbUp
                    Else
                        BucketKb(BucketIndex) = KbDown
                        BucketSb(BucketIndex) = SbDown
                    End If
                Else
                    BucketKb(BucketIndex) = 0#
                    BucketSb(BucketIndex) = 0#
                    For MemberIndex = 1 To MemberCounts(BucketIndex)
                        FactorIndex = BucketMembers(BucketIndex, MemberIndex)
                        BucketSb(BucketIndex) = BucketSb(BucketIndex) + mWS(FactorIndex)
                        If BucketIsOther(BucketIndex) Then
                            BucketKb(BucketIndex) = BucketKb(BucketIndex) + Abs(mWS(FactorIndex))
                        Else
                            BucketKb(BucketIndex) = BucketKb(BucketIndex) + mWS(FactorIndex) ^ 2
                            For OtherMemberIndex = 1 To MemberCounts(BucketIndex)
                                If OtherMemberIndex <> MemberIndex Then
                                    OtherFactorIndex = BucketMembers(BucketIndex, OtherMemberIndex)
                                    BucketKb(BucketIndex) = BucketKb(BucketIndex) + PairCorrelation(comp, RiskClass, _
                                        Labels(BucketIndex), FactorIndex, OtherFactorIndex, ScenarioIndex) * _
                                        mWS(FactorIndex) * mWS(OtherFactorIndex)
                                End If
                            Next OtherMemberIndex
                        End If
                    Next MemberIndex
                    If Not BucketIsOther(BucketIndex) Then BucketKb(BucketIndex) = Sqr(MaxD(BucketKb(BucketIndex), 0#))
                End If
                If BucketIsOther(BucketIndex) Then
                    If BucketNote(BucketIndex) = vbNullString Then
                        BucketNote(BucketIndex) = "OTHER"
                    Else
                        BucketNote(BucketIndex) = BucketNote(BucketIndex) & ";OTHER"
                    End If
                End If
            Next BucketIndex
            ' Across buckets
            AltUsed = False
            SumSquares = AcrossBuckets(comp, RiskClass, BucketCount, Labels, BucketIsOther, BucketKb, BucketSb, _
                ScenarioIndex, False)
            If SumSquares < 0# And comp <> COMP_CURV Then
                ' Alternative specification Sb = max(min(sum WS, Kb), -Kb) (MAR21.4(5)(b))
                AltUsed = True
                SumSquares = AcrossBuckets(comp, RiskClass, BucketCount, Labels, BucketIsOther, BucketKb, BucketSb, _
                    ScenarioIndex, True)
            End If
            OtherBucketSum = 0#
            For BucketIndex = 1 To BucketCount
                If BucketIsOther(BucketIndex) Then OtherBucketSum = OtherBucketSum + BucketKb(BucketIndex)
            Next BucketIndex
            RiskClassCapital = Sqr(MaxD(SumSquares, 0#)) + OtherBucketSum
            mCapital(rcIdx, comp, ScenarioIndex) = RiskClassCapital
            mAltUsed(rcIdx, comp, ScenarioIndex) = AltUsed
            For BucketIndex = 1 To BucketCount
                If AltUsed And Not BucketIsOther(BucketIndex) Then
                    AltSb = MaxD(MinD(BucketSb(BucketIndex), BucketKb(BucketIndex)), -BucketKb(BucketIndex))
                    WriteBucketRow RiskClass, comp, ScenarioIndex, Labels(BucketIndex), BucketKb(BucketIndex), _
                        AltSb, MemberCounts(BucketIndex), _
                        IIf(BucketNote(BucketIndex) = vbNullString, "ALT_SB", BucketNote(BucketIndex) & ";ALT_SB")
                Else
                    WriteBucketRow RiskClass, comp, ScenarioIndex, Labels(BucketIndex), BucketKb(BucketIndex), _
                        BucketSb(BucketIndex), MemberCounts(BucketIndex), BucketNote(BucketIndex)
                End If
            Next BucketIndex
        Next ScenarioIndex

End Sub

Private Function AcrossBuckets( _
    ByVal comp As Long, _
    ByVal rc As String, _
    ByVal nB As Long, _
    ByRef labels() As String, _
    ByRef isOther() As Boolean, _
    ByRef kb() As Double, _
    ByRef sb() As Double, _
    ByVal scen As Long, _
    ByVal useAlt As Boolean) As Double
'
'------------------------------------------------------------------------------
'                           AGGREGATE ACROSS BUCKETS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the sum under the square root of the aggregation across the buckets of a risk
'   class: the squared Kb of each bucket plus the cross-bucket terms.
'
' WHY THIS EXISTS
'   The aggregation is done twice when the first sum is negative, with the Sb of the
'   buckets and with the alternative Sb.
'
' INPUTS
'   comp
'     Component: COMP_DELTA, COMP_VEGA or COMP_CURV
'
'   rc
'     Risk class code
'
'   nB
'     Number of buckets
'
'   labels
'     Bucket labels, indexed 1 to nB
'
'   isOther
'     True for an other-sector bucket, indexed 1 to nB
'
'   kb, sb
'     Kb and Sb of each bucket, indexed 1 to nB
'
'   scen
'     Scenario: SCEN_LOW, SCEN_MEDIUM or SCEN_HIGH
'
'   useAlt
'     True to replace Sb by max(min(Sb, Kb), -Kb)
'
' RETURNS
'   The sum before the square root; it can be negative.
'
' BEHAVIOR
'   Skips the other buckets. For each remaining bucket adds Kb squared, then for each other
'   remaining bucket adds the cross-bucket correlation times the two Sb values. For
'   curvature the product is also multiplied by Psi.
'
' ERROR POLICY
'   Does not handle errors; errors from the callees are not caught here.
'
' DEPENDENCIES
'   MinD, MaxD (CORE_Util); CrossBucketCorrelation, Psi.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.AcrossBuckets"

    Dim BucketIndex             As Long          'First bucket of a pair
    Dim OtherBucketIndex        As Long          'Second bucket of a pair
    Dim Aggregate               As Double        'Running sum
    Dim SbOfBucket              As Double        'Sb of the first bucket, capped when useAlt is True
    Dim SbOfOther               As Double        'Sb of the second bucket, capped when useAlt is True

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the sum at zero
        Aggregate = 0#

'------------------------------------------------------------------------------
' AGGREGATE
'------------------------------------------------------------------------------
    'Add the Kb terms and the cross-bucket terms of the buckets that are not other
        For BucketIndex = 1 To nB
            If Not isOther(BucketIndex) Then
                Aggregate = Aggregate + kb(BucketIndex) ^ 2
                For OtherBucketIndex = 1 To nB
                    If OtherBucketIndex <> BucketIndex And Not isOther(OtherBucketIndex) Then
                        SbOfBucket = sb(BucketIndex)
                        SbOfOther = sb(OtherBucketIndex)
                        If useAlt Then
                            SbOfBucket = MaxD(MinD(SbOfBucket, kb(BucketIndex)), -kb(BucketIndex))
                            SbOfOther = MaxD(MinD(SbOfOther, kb(OtherBucketIndex)), -kb(OtherBucketIndex))
                        End If
                        If comp = COMP_CURV Then
                            Aggregate = Aggregate + CrossBucketCorrelation(comp, rc, labels(BucketIndex), _
                                labels(OtherBucketIndex), scen) * SbOfBucket * SbOfOther * Psi(SbOfBucket, SbOfOther)
                        Else
                            Aggregate = Aggregate + CrossBucketCorrelation(comp, rc, labels(BucketIndex), _
                                labels(OtherBucketIndex), scen) * SbOfBucket * SbOfOther
                        End If
                    End If
                Next OtherBucketIndex
            End If
        Next BucketIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the sum
        AcrossBuckets = Aggregate

End Function

'
'------------------------------------------------------------------------------
'
'                                    OUTPUT
'
'------------------------------------------------------------------------------
'

Private Sub ResetBucketOutput()
'
'------------------------------------------------------------------------------
'                           RESET THE BUCKET OUTPUT
'------------------------------------------------------------------------------
' PURPOSE
'   Clears Out_Buckets and writes its header row.
'
' BEHAVIOR
'   Clears the table area, writes the eight column captions and sets the last output row to
'   the header row.
'
' DEPENDENCIES
'   GetSheet, ClearTableArea, WriteTableHeader (CORE_Util).
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.ResetBucketOutput"

    Dim OutputSheet             As Object        'Out_Buckets sheet

'------------------------------------------------------------------------------
' RESET OUTPUT
'------------------------------------------------------------------------------
    'Locate the output sheet
        Set OutputSheet = GetSheet(SH_OUT_BUCKETS)
    'Clear the table area
        ClearTableArea OutputSheet
    'Write the header
        WriteTableHeader OutputSheet, LAYOUT_HEADER_ROW, _
            Array("RiskClass", "Component", "Scenario", "Bucket", "Kb", "Sb", "RiskFactors", "Note"), COLOR_OUTPUT_HEADER
    'Start the output row at the header row
        mOutRow = LAYOUT_HEADER_ROW

End Sub

Private Sub FormatBucketOutput()
'
'------------------------------------------------------------------------------
'                           FORMAT THE BUCKET OUTPUT
'------------------------------------------------------------------------------
' PURPOSE
'   Formats the rows written to Out_Buckets.
'
' BEHAVIOR
'   Formats the data rows, then centers the bucket column. Does nothing more when no row
'   was written.
'
' ERROR POLICY
'   Errors while centering the bucket column are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   GetSheet, TableCell, FormatDataRows (CORE_Util).
'
' NOTES
'   Bucket codes are text (EUR) or numbers (16): one alignment for the column (#206).
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.FormatBucketOutput"

    Dim OutputSheet             As Object        'Out_Buckets sheet

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Locate the output sheet
        Set OutputSheet = GetSheet(SH_OUT_BUCKETS)
    'Format the data rows
        FormatDataRows OutputSheet, LAYOUT_FIRST_DATA_ROW, mOutRow, 8, 5, "#,##0", False
    'Stop when no row was written
        If mOutRow < LAYOUT_FIRST_DATA_ROW Then Exit Sub
    'Center the bucket column; ignore formatting errors
        On Error Resume Next
        OutputSheet.Range(TableCell(OutputSheet, LAYOUT_FIRST_DATA_ROW, 4), _
            TableCell(OutputSheet, mOutRow, 4)).HorizontalAlignment = -4108
        On Error GoTo 0

End Sub

Private Sub WriteBucketRow( _
    ByVal rc As String, _
    ByVal comp As Long, _
    ByVal scen As Long, _
    ByVal bucket As String, _
    ByVal kb As Double, _
    ByVal sb As Double, _
    ByVal nRF As Long, _
    ByVal note As String)
'
'------------------------------------------------------------------------------
'                              WRITE A BUCKET ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one row to Out_Buckets.
'
' INPUTS
'   rc
'     Risk class code
'
'   comp
'     Component: COMP_DELTA, COMP_VEGA or COMP_CURV
'
'   scen
'     Scenario: SCEN_LOW, SCEN_MEDIUM or SCEN_HIGH
'
'   bucket
'     Bucket label
'
'   kb, sb
'     Kb and Sb of the bucket
'
'   nRF
'     Number of risk factors in the bucket
'
'   note
'     Note of the row
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Advances the output row and writes the risk class, component name, scenario name,
'   bucket, Kb, Sb, number of risk factors and note in columns 1 to 8.
'
' DEPENDENCIES
'   GetSheet, TableCell (CORE_Util); ComponentName, ScenarioName (CORE_Config).
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.WriteBucketRow"

    Dim OutputSheet             As Object        'Out_Buckets sheet

'------------------------------------------------------------------------------
' WRITE ROW
'------------------------------------------------------------------------------
    'Locate the output sheet
        Set OutputSheet = GetSheet(SH_OUT_BUCKETS)
    'Advance to the next output row
        mOutRow = mOutRow + 1
    'Write the eight columns of the row
        TableCell(OutputSheet, mOutRow, 1).Value = rc
        TableCell(OutputSheet, mOutRow, 2).Value = ComponentName(comp)
        TableCell(OutputSheet, mOutRow, 3).Value = ScenarioName(scen)
        TableCell(OutputSheet, mOutRow, 4).Value = bucket
        TableCell(OutputSheet, mOutRow, 5).Value = kb
        TableCell(OutputSheet, mOutRow, 6).Value = sb
        TableCell(OutputSheet, mOutRow, 7).Value = nRF
        TableCell(OutputSheet, mOutRow, 8).Value = note

End Sub

Private Sub WriteRiskClassOutput()
'
'------------------------------------------------------------------------------
'                         WRITE THE RISK CLASS OUTPUT
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the capital of every risk class, component and scenario to Out_RiskClass.
'
' WHY THIS EXISTS
'   The risk class table shows the result of the last run by risk class, component and
'   scenario, with the number of buckets and whether the alternative Sb was used.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the table area, writes the header, then one row per risk class, component and
'   scenario with the capital, the number of buckets and a Y or N flag for the alternative
'   Sb. Formats the rows and centers the flag column.
'
' ERROR POLICY
'   Errors while centering the flag column are ignored (On Error Resume Next).
'
' DEPENDENCIES
'   GetSheet, ClearTableArea, WriteTableHeader, TableCell, FormatDataRows (CORE_Util);
'   RiskClassCode, ComponentName, ScenarioName (CORE_Config).
'
' NOTES
'   The flag column is centered because of #206.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_SBM.WriteRiskClassOutput"

    Dim OutputSheet             As Object        'Out_RiskClass sheet
    Dim OutputRow               As Long          'Row being written
    Dim ClassIndex              As Long          'Risk class being written
    Dim ComponentIndex          As Long          'Component being written
    Dim ScenarioIndex           As Long          'Scenario being written

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Locate the output sheet
        Set OutputSheet = GetSheet(SH_OUT_CLASS)
    'Clear the table area
        ClearTableArea OutputSheet
    'Write the header
        WriteTableHeader OutputSheet, LAYOUT_HEADER_ROW, _
            Array("RiskClass", "Component", "Scenario", "Capital", "Buckets", "AltSbUsed"), COLOR_OUTPUT_HEADER
    'Start the output row at the header row
        OutputRow = LAYOUT_HEADER_ROW

'------------------------------------------------------------------------------
' WRITE ROWS
'------------------------------------------------------------------------------
    'Write one row per risk class, component and scenario
        For ClassIndex = 1 To RC_COUNT
            For ComponentIndex = 1 To COMP_COUNT
                For ScenarioIndex = 1 To SCEN_COUNT
                    OutputRow = OutputRow + 1
                    TableCell(OutputSheet, OutputRow, 1).Value = RiskClassCode(ClassIndex)
                    TableCell(OutputSheet, OutputRow, 2).Value = ComponentName(ComponentIndex)
                    TableCell(OutputSheet, OutputRow, 3).Value = ScenarioName(ScenarioIndex)
                    TableCell(OutputSheet, OutputRow, 4).Value = mCapital(ClassIndex, ComponentIndex, ScenarioIndex)
                    TableCell(OutputSheet, OutputRow, 5).Value = mBucketCount(ClassIndex, ComponentIndex)
                    TableCell(OutputSheet, OutputRow, 6).Value = _
                        IIf(mAltUsed(ClassIndex, ComponentIndex, ScenarioIndex), "Y", "N")
                Next ScenarioIndex
            Next ComponentIndex
        Next ClassIndex

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Format the data rows
        FormatDataRows OutputSheet, LAYOUT_FIRST_DATA_ROW, OutputRow, 6, 4, "#,##0", False
    'Center the flag column (#206); ignore formatting errors
        On Error Resume Next
        OutputSheet.Range(TableCell(OutputSheet, LAYOUT_FIRST_DATA_ROW, 6), _
            TableCell(OutputSheet, OutputRow, 6)).HorizontalAlignment = -4108
        On Error GoTo 0

End Sub
