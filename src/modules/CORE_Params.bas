Attribute VB_Name = "CORE_Params"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Params
'------------------------------------------------------------------------------
' PURPOSE
'   Loads the run settings (Control sheet) and every regulatory parameter table (P_*
'   sheets) into module storage, and exposes typed accessors: risk weights, correlation
'   building blocks, cross-bucket correlations, liquidity horizons, DRC and RRAO
'   parameters.
'
' WHY THIS EXISTS
'   The P_* sheets are the source of the parameters. The module reads them at each run and
'   the engines read them through typed accessors.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   LoadAllParameters reads the settings, the general tables, the GIRR, CSR, equity,
'   commodity and FX tables, and the DRC and RRAO tables into module variables. The
'   accessors return those values, and the validity and classification checks answer from
'   them.
'
' ERROR POLICY
'   A missing table, or a required key, table value or vertex list entry that is blank or
'   not a number, raises (CORE_Util); an unknown tenor or liquidity horizon raises
'   vbObjectError + ERR_APP. The run stops before any capital is written.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util.
'
' NOTES
'   Each ReadKeyValues(SH_*, "TAG") call stays on one line, exactly as written: the
'   workbook tools read these calls by pattern to list the key tables.
'
'   The module also runs in LibreOffice Calc with VBA support for the automated tests. It
'   needs no references beyond the defaults.
'
' PUBLIC SURFACE
'   LoadAllParameters; ReportingCcy, FxCurvatureDivisor, ScenarioCorrelation, VegaAlpha,
'   VegaRiskWeight; IsValid* vertex and bucket checks, IsOtherBucket; *RiskWeight;
'   correlation accessors; BucketGamma; Drc* and RraoRiskWeight.
'
' STATE OWNERSHIP
'   Owns the parameters of the current run (module variables m*), replaced by every
'   LoadAllParameters call. Parameters are never hard-coded here: the P_* sheets are the
'   source, and M_Setup only writes their defaults.
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'Settings
    Private mReportingCcy       As String   'Reporting currency, upper case
    Private mGirrSpecifiedOn    As Boolean  'True when the GIRR specified-currency discretion is on
    Private mFxSpecifiedOn      As Boolean  'True when the FX specified-pair discretion is on
    Private mFxCurvDivisorOn    As Boolean  'True when the FX curvature divisor applies to options not referencing the reporting currency
    Private mFxCurvDivisorAllOn As Boolean  'True when the FX curvature divisor applies to all FX instruments

    'General
    Private mHighMult       As Double   'Multiplier of the high correlation scenario
    Private mLowMult        As Double   'Multiplier of the low correlation scenario
    Private mVegaRWSigma    As Double   'Vega risk weight sigma
    Private mVegaAlpha      As Double   'Vega alpha
    Private mLHCount        As Long     'Number of liquidity horizon rows
    Private mLHClass()      As String   'Risk class of each liquidity horizon row
    Private mLHBucket()     As String   'Bucket label of each liquidity horizon row
    Private mLHValue()      As Double   'Liquidity horizon of each row
    Private mGirrTenors()   As Double   'GIRR vertex tenors
    Private mGirrTenorN     As Long     'Number of GIRR vertex tenors
    Private mCsrTenors()    As Double   'CSR vertex tenors
    Private mCsrTenorN      As Long     'Number of CSR vertex tenors
    Private mComTenors()    As Double   'Commodity vertex tenors
    Private mComTenorN      As Long     'Number of commodity vertex tenors
    Private mVegaMats()     As Double   'Vega vertex maturities
    Private mVegaMatN       As Long     'Number of vega vertex maturities

    'GIRR
    Private mGirrRWTenor()      As Double       'Tenor of each GIRR risk weight row
    Private mGirrRWValue()      As Double       'Risk weight of each GIRR risk weight row
    Private mGirrRWCount        As Long         'Number of GIRR risk weight rows
    Private mGirrTheta          As Double       'Theta of the tenor correlation (THETA)
    Private mGirrRhoFloor       As Double       'Floor of the tenor correlation (RHO_FLOOR)
    Private mGirrRhoDiffCurve   As Double       'Correlation between different curves (RHO_DIFF_CURVE)
    Private mGirrRhoInflation   As Double       'Correlation of the inflation risk factor (RHO_INFLATION)
    Private mGirrRhoXccy        As Double       'Correlation of the cross-currency basis risk factor (RHO_XCCY)
    Private mGirrRWInflation    As Double       'Inflation risk weight (RW_INFLATION)
    Private mGirrRWXccy         As Double       'Cross-currency basis risk weight (RW_XCCY)
    Private mGirrGamma          As Double       'Gamma between GIRR buckets (GAMMA)
    Private mGirrSpecDivisor    As Double       'Divisor of the specified-currency discretion (SPECIFIED_DIVISOR)
    Private mGirrSpecCcy        As Collection   'Specified currencies, keyed by currency

    'CSR non-securitization
    Private mCsrN               As Long         'Highest CSR bucket number
    Private mCsrQuality()       As String       'Credit quality of each bucket
    Private mCsrGroup()         As String       'Sector group of each bucket
    Private mCsrType()          As String       'Type of each bucket
    Private mCsrRW()            As Double       'Risk weight of each bucket
    Private mCsrRhoName         As Double       'Name correlation (RHO_NAME)
    Private mCsrRhoNameIndex    As Double       'Name correlation of index buckets (RHO_NAME_INDEX)
    Private mCsrRhoTenor        As Double       'Tenor correlation (RHO_TENOR)
    Private mCsrRhoBasis        As Double       'Basis correlation (RHO_BASIS)
    Private mCsrGammaRating     As Double       'Rating factor between buckets of different credit quality (GAMMA_RATING)
    Private mCsrGroupIdx        As Collection   'Row of each sector group in the sector matrix
    Private mCsrSector()        As Double       'Sector gamma matrix

    'Equity
    Private mEqN            As Long     'Highest equity bucket number
    Private mEqType()       As String   'Type of each bucket
    Private mEqRWSpot()     As Double   'Spot risk weight of each bucket
    Private mEqRWRepo()     As Double   'Repo risk weight of each bucket
    Private mEqRhoName()    As Double   'Name correlation of each bucket
    Private mEqRhoSpotRepo  As Double   'Spot-repo correlation (RHO_SPOT_REPO)
    Private mEqGamma()      As Double   'Cross-bucket gamma matrix

    'Commodity
    Private mComN           As Long     'Highest commodity bucket number
    Private mComRW()        As Double   'Risk weight of each bucket
    Private mComRhoCty()    As Double   'Commodity correlation of each bucket
    Private mComRhoTenor    As Double   'Tenor correlation (RHO_TENOR)
    Private mComRhoBasis    As Double   'Basis correlation (RHO_BASIS)
    Private mComGamma()     As Double   'Cross-bucket gamma matrix

    'FX
    Private mFxRW           As Double       'FX risk weight (RW)
    Private mFxGamma        As Double       'Gamma between FX buckets (GAMMA)
    Private mFxSpecDivisor  As Double       'Divisor of the specified-pair discretion (SPECIFIED_DIVISOR)
    Private mFxCurvDivisor  As Double       'Divisor of the curvature discretion (CURVATURE_DIVISOR)
    Private mFxSpecPairs    As Collection   'Specified currency pairs, keyed by pair

    'DRC
    Private mDrcRW              As Collection   'Risk weight by credit quality
    Private mDrcLGD             As Collection   'Loss given default by seniority
    Private mDrcRank            As Collection   'Rank by seniority
    Private mDrcBuckets         As Collection   'Position of each bucket, keyed by bucket name
    Private mDrcBucketNames()   As String       'Bucket names in table order
    Private mDrcMinMat          As Double       'Minimum maturity (MIN_MATURITY)
    Private mDrcMaxMat          As Double       'Maximum maturity (MAX_MATURITY)

    'RRAO
    Private mRraoRW As Collection   'Risk weight by category


'
'------------------------------------------------------------------------------
'
'                                   LOADING
'
'------------------------------------------------------------------------------
'

Public Sub LoadAllParameters()

'
'------------------------------------------------------------------------------
'                             LOAD ALL PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Loads the run settings and every parameter table into module storage.
'
' WHY THIS EXISTS
'   Every run starts from the current content of the Control and P_* sheets, so the
'   parameters are read again at each call and replace those of the previous run.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls the loaders in order:
'     - LoadSettings
'     - LoadGeneral
'     - LoadGirr
'     - LoadCsr
'     - LoadEquity
'     - LoadCommodity
'     - LoadFx
'     - LoadDrc
'     - LoadRrao
'
' ERROR POLICY
'   Does not handle errors: a missing table or an invalid value raises from the loaders
'   (CORE_Util).
'
' DEPENDENCIES
'   LoadSettings, LoadGeneral, LoadGirr, LoadCsr, LoadEquity, LoadCommodity, LoadFx,
'   LoadDrc, LoadRrao.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' LOAD PARAMETERS
'------------------------------------------------------------------------------
    'Read the run settings of the Control sheet
        LoadSettings
    'Read the general parameters
        LoadGeneral
    'Read the GIRR parameters
        LoadGirr
    'Read the CSR parameters
        LoadCsr
    'Read the equity parameters
        LoadEquity
    'Read the commodity parameters
        LoadCommodity
    'Read the FX parameters
        LoadFx
    'Read the DRC parameters
        LoadDrc
    'Read the RRAO parameters
        LoadRrao

End Sub

Private Sub LoadSettings()

'
'------------------------------------------------------------------------------
'                            LOAD THE RUN SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the run settings of the Control sheet: the reporting currency and the discretion
'   switches.
'
' WHY THIS EXISTS
'   The reporting currency and the switches decide how the specified-currency and
'   specified-pair weights and the FX curvature divisor are applied.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the SETTINGS table of Control as key and value pairs, then stores:
'     - REPORTING_CCY in upper case, EUR when the key is absent
'     - GIRR_SPECIFIED_CCY_DISCRETION, FX_SPECIFIED_PAIR_DISCRETION, FX_CURVATURE_DIVISOR
'       and FX_CURVATURE_DIVISOR_ALL, each False when the key is absent
'
' ERROR POLICY
'   None of its own. ReadKeyValues raises when the SETTINGS table is missing.
'
' DEPENDENCIES
'   ReadKeyValues, KVText, KVBool (CORE_Util); SH_CONTROL (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadSettings"

    Dim KeyValues               As Collection   'Key and value pairs of the SETTINGS table

'------------------------------------------------------------------------------
' READ SETTINGS
'------------------------------------------------------------------------------
    'Read the SETTINGS table of the Control sheet
        Set KeyValues = ReadKeyValues(SH_CONTROL, "SETTINGS")
    'Store the reporting currency in upper case
        mReportingCcy = UCase$(KVText(KeyValues, "REPORTING_CCY", "EUR"))
    'Store the discretion and divisor switches
        mGirrSpecifiedOn = KVBool(KeyValues, "GIRR_SPECIFIED_CCY_DISCRETION", False)
        mFxSpecifiedOn = KVBool(KeyValues, "FX_SPECIFIED_PAIR_DISCRETION", False)
        mFxCurvDivisorOn = KVBool(KeyValues, "FX_CURVATURE_DIVISOR", False)
        mFxCurvDivisorAllOn = KVBool(KeyValues, "FX_CURVATURE_DIVISOR_ALL", False)

End Sub

Private Sub LoadGeneral()

'
'------------------------------------------------------------------------------
'                         LOAD THE GENERAL PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the general parameters of P_General: the correlation scenario multipliers, the
'   vega parameters, the vega liquidity horizons and the vertex lists.
'
' WHY THIS EXISTS
'   ScenarioCorrelation, VegaRiskWeight and the vertex checks (IsValidGirrTenor,
'   IsValidCsrTenor, IsValidComTenor, IsValidVegaMaturity) answer from these values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - reads HIGH_MULTIPLIER and LOW_MULTIPLIER from CORR_SCENARIOS
'     - reads RW_SIGMA and ALPHA from VEGA
'     - reads the VEGA_LH table into one array per column: risk class, bucket label and
'       horizon
'     - reads the GIRR, CSR and commodity tenor lists and the vega maturity list from
'       VERTICES
'
' ERROR POLICY
'   Raises (CORE_Util) when a table or a required key is missing, a value is blank or not
'   a number, or a vertex list is empty or has a blank or non-numeric entry.
'
' DEPENDENCIES
'   ReadKeyValues, ReadTable, KVNumber, KVText, TableNumber, ToText, BucketLabel,
'   ParseNumberList (CORE_Util); SH_P_GENERAL (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadGeneral"

    Dim KeyValues               As Collection   'Key and value pairs of the table being read
    Dim TableData               As Variant      'Rows of the VEGA_LH table
    Dim RowCount                As Long         'Rows of the table
    Dim ColumnCount             As Long         'Columns of the table
    Dim RowIndex                As Long         'Row being read

'------------------------------------------------------------------------------
' READ CORRELATION SCENARIOS
'------------------------------------------------------------------------------
    'Read the multipliers of the high and low correlation scenarios
        Set KeyValues = ReadKeyValues(SH_P_GENERAL, "CORR_SCENARIOS")
        mHighMult = KVNumber(KeyValues, "HIGH_MULTIPLIER", SH_P_GENERAL)
        mLowMult = KVNumber(KeyValues, "LOW_MULTIPLIER", SH_P_GENERAL)

'------------------------------------------------------------------------------
' READ VEGA PARAMETERS
'------------------------------------------------------------------------------
    'Read the vega risk weight sigma and alpha
        Set KeyValues = ReadKeyValues(SH_P_GENERAL, "VEGA")
        mVegaRWSigma = KVNumber(KeyValues, "RW_SIGMA", SH_P_GENERAL)
        mVegaAlpha = KVNumber(KeyValues, "ALPHA", SH_P_GENERAL)

'------------------------------------------------------------------------------
' READ LIQUIDITY HORIZONS
'------------------------------------------------------------------------------
    'Read the liquidity horizon table and size one array per column
        TableData = ReadTable(SH_P_GENERAL, "VEGA_LH", RowCount, ColumnCount)
        mLHCount = RowCount
        ReDim mLHClass(1 To RowCount)
        ReDim mLHBucket(1 To RowCount)
        ReDim mLHValue(1 To RowCount)
    'Store the risk class, bucket label and horizon of each row
        For RowIndex = 1 To RowCount
            mLHClass(RowIndex) = ToText(TableData(RowIndex, 1))
            mLHBucket(RowIndex) = BucketLabel(TableData(RowIndex, 2))
            mLHValue(RowIndex) = TableNumber(TableData, RowIndex, 3, SH_P_GENERAL, "VEGA_LH")
        Next RowIndex

'------------------------------------------------------------------------------
' READ VERTEX LISTS
'------------------------------------------------------------------------------
    'Read the tenor lists of the GIRR, CSR and commodity risk classes and the vega maturities
        Set KeyValues = ReadKeyValues(SH_P_GENERAL, "VERTICES")
        mGirrTenorN = ParseNumberList(KVText(KeyValues, "GIRR_TENORS", vbNullString), mGirrTenors, _
            "GIRR_TENORS (" & SH_P_GENERAL & ")")
        mCsrTenorN = ParseNumberList(KVText(KeyValues, "CSR_TENORS", vbNullString), mCsrTenors, _
            "CSR_TENORS (" & SH_P_GENERAL & ")")
        mComTenorN = ParseNumberList(KVText(KeyValues, "COM_TENORS", vbNullString), mComTenors, _
            "COM_TENORS (" & SH_P_GENERAL & ")")
        mVegaMatN = ParseNumberList(KVText(KeyValues, "VEGA_MATURITIES", vbNullString), mVegaMats, _
            "VEGA_MATURITIES (" & SH_P_GENERAL & ")")

End Sub

Private Sub LoadGirr()

'
'------------------------------------------------------------------------------
'                           LOAD THE GIRR PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the GIRR parameters of P_GIRR: the risk weights by tenor, the correlation and
'   weight parameters and the specified currencies.
'
' WHY THIS EXISTS
'   GirrRiskWeight, GirrTenorCorrelation, BucketGamma and the GIRR correlation accessors
'   answer from these values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - reads the GIRR_RW table into a tenor array and a risk weight array
'     - reads THETA, RHO_FLOOR, RHO_DIFF_CURVE, RHO_INFLATION, RHO_XCCY, RW_INFLATION,
'       RW_XCCY, GAMMA and SPECIFIED_DIVISOR from GIRR_PARAMS
'     - rebuilds the specified currencies from GIRR_SPECIFIED_CCY, adding each currency
'       once
'
' ERROR POLICY
'   Raises (CORE_Util) when a table or a required key is missing or a value is blank or
'   not a number.
'
' DEPENDENCIES
'   ReadKeyValues, ReadTable, KVNumber, TableNumber, ToText, KeyExists, AddKey
'   (CORE_Util); SH_P_GIRR (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadGirr"

    Dim KeyValues               As Collection   'Key and value pairs of the GIRR_PARAMS table
    Dim TableData               As Variant      'Rows of the table being read
    Dim RowCount                As Long         'Rows of the table
    Dim ColumnCount             As Long         'Columns of the table
    Dim RowIndex                As Long         'Row being read

'------------------------------------------------------------------------------
' READ RISK WEIGHTS
'------------------------------------------------------------------------------
    'Read the risk weight table and size the tenor and weight arrays
        TableData = ReadTable(SH_P_GIRR, "GIRR_RW", RowCount, ColumnCount)
        mGirrRWCount = RowCount
        ReDim mGirrRWTenor(1 To RowCount)
        ReDim mGirrRWValue(1 To RowCount)
    'Store the tenor and the risk weight of each row
        For RowIndex = 1 To RowCount
            mGirrRWTenor(RowIndex) = TableNumber(TableData, RowIndex, 1, SH_P_GIRR, "GIRR_RW")
            mGirrRWValue(RowIndex) = TableNumber(TableData, RowIndex, 2, SH_P_GIRR, "GIRR_RW")
        Next RowIndex

'------------------------------------------------------------------------------
' READ PARAMETERS
'------------------------------------------------------------------------------
    'Read the correlation, weight, gamma and divisor parameters
        Set KeyValues = ReadKeyValues(SH_P_GIRR, "GIRR_PARAMS")
        mGirrTheta = KVNumber(KeyValues, "THETA", SH_P_GIRR)
        mGirrRhoFloor = KVNumber(KeyValues, "RHO_FLOOR", SH_P_GIRR)
        mGirrRhoDiffCurve = KVNumber(KeyValues, "RHO_DIFF_CURVE", SH_P_GIRR)
        mGirrRhoInflation = KVNumber(KeyValues, "RHO_INFLATION", SH_P_GIRR)
        mGirrRhoXccy = KVNumber(KeyValues, "RHO_XCCY", SH_P_GIRR)
        mGirrRWInflation = KVNumber(KeyValues, "RW_INFLATION", SH_P_GIRR)
        mGirrRWXccy = KVNumber(KeyValues, "RW_XCCY", SH_P_GIRR)
        mGirrGamma = KVNumber(KeyValues, "GAMMA", SH_P_GIRR)
        mGirrSpecDivisor = KVNumber(KeyValues, "SPECIFIED_DIVISOR", SH_P_GIRR)

'------------------------------------------------------------------------------
' READ SPECIFIED CURRENCIES
'------------------------------------------------------------------------------
    'Rebuild the specified currencies, keeping each currency once
        Set mGirrSpecCcy = New Collection
        TableData = ReadTable(SH_P_GIRR, "GIRR_SPECIFIED_CCY", RowCount, ColumnCount)
        For RowIndex = 1 To RowCount
            If Not KeyExists(mGirrSpecCcy, ToText(TableData(RowIndex, 1))) Then
                AddKey mGirrSpecCcy, ToText(TableData(RowIndex, 1)), True
            End If
        Next RowIndex

End Sub

Private Sub LoadCsr()

'
'------------------------------------------------------------------------------
'                           LOAD THE CSR PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the CSR non-securitization parameters of P_CSR: the bucket table, the
'   correlation parameters and the sector matrix.
'
' WHY THIS EXISTS
'   CsrRiskWeight, CsrNameCorrelation, BucketGamma, IsValidBucket and IsOtherBucket answer
'   from these values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - reads CSR_BUCKETS: the highest bucket number sets the array size, then the credit
'       quality, sector group, type and risk weight of each bucket are stored by bucket
'       number
'     - reads RHO_NAME, RHO_NAME_INDEX, RHO_TENOR, RHO_BASIS and GAMMA_RATING from
'       CSR_PARAMS
'     - reads the CSR_GAMMA_SECTOR matrix: the first column is the sector group of each
'       row, and the groups are indexed by row
'
' ERROR POLICY
'   Raises (CORE_Util) when a table or a required key is missing or a value is blank or
'   not a number.
'
' DEPENDENCIES
'   ReadKeyValues, ReadTable, KVNumber, TableNumber, ToText, BucketLabel, AddKey
'   (CORE_Util); SH_P_CSR (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadCsr"

    Dim KeyValues               As Collection   'Key and value pairs of the CSR_PARAMS table
    Dim TableData               As Variant      'Rows of the table being read
    Dim RowCount                As Long         'Rows of the table
    Dim ColumnCount             As Long         'Columns of the table
    Dim RowIndex                As Long         'Row being read
    Dim ColumnIndex             As Long         'Column of the sector matrix being read
    Dim BucketNumber            As Long         'Bucket number of the row

'------------------------------------------------------------------------------
' READ BUCKETS
'------------------------------------------------------------------------------
    'Read the bucket table and find the highest bucket number
        TableData = ReadTable(SH_P_CSR, "CSR_BUCKETS", RowCount, ColumnCount)
        mCsrN = 0
        For RowIndex = 1 To RowCount
            BucketNumber = CLng(TableNumber(TableData, RowIndex, 1, SH_P_CSR, "CSR_BUCKETS"))
            If BucketNumber > mCsrN Then mCsrN = BucketNumber
        Next RowIndex
    'Size the bucket arrays
        ReDim mCsrQuality(1 To mCsrN)
        ReDim mCsrGroup(1 To mCsrN)
        ReDim mCsrType(1 To mCsrN)
        ReDim mCsrRW(1 To mCsrN)
    'Store the quality, sector group, type and risk weight of each bucket
        For RowIndex = 1 To RowCount
            BucketNumber = CLng(TableNumber(TableData, RowIndex, 1, SH_P_CSR, "CSR_BUCKETS"))
            mCsrQuality(BucketNumber) = ToText(TableData(RowIndex, 3))
            mCsrGroup(BucketNumber) = BucketLabel(TableData(RowIndex, 4))
            mCsrType(BucketNumber) = ToText(TableData(RowIndex, 5))
            mCsrRW(BucketNumber) = TableNumber(TableData, RowIndex, 6, SH_P_CSR, "CSR_BUCKETS")
        Next RowIndex

'------------------------------------------------------------------------------
' READ PARAMETERS
'------------------------------------------------------------------------------
    'Read the correlation parameters and the rating factor
        Set KeyValues = ReadKeyValues(SH_P_CSR, "CSR_PARAMS")
        mCsrRhoName = KVNumber(KeyValues, "RHO_NAME", SH_P_CSR)
        mCsrRhoNameIndex = KVNumber(KeyValues, "RHO_NAME_INDEX", SH_P_CSR)
        mCsrRhoTenor = KVNumber(KeyValues, "RHO_TENOR", SH_P_CSR)
        mCsrRhoBasis = KVNumber(KeyValues, "RHO_BASIS", SH_P_CSR)
        mCsrGammaRating = KVNumber(KeyValues, "GAMMA_RATING", SH_P_CSR)

'------------------------------------------------------------------------------
' READ SECTOR MATRIX
'------------------------------------------------------------------------------
    'Sector matrix: first column holds the row group, header holds the column groups
        TableData = ReadTable(SH_P_CSR, "CSR_GAMMA_SECTOR", RowCount, ColumnCount)
    'Index the sector groups by row
        Set mCsrGroupIdx = New Collection
        For RowIndex = 1 To RowCount
            AddKey mCsrGroupIdx, BucketLabel(TableData(RowIndex, 1)), RowIndex
        Next RowIndex
    'Store the gamma of each pair of sector groups
        ReDim mCsrSector(1 To RowCount, 1 To RowCount)
        For RowIndex = 1 To RowCount
            For ColumnIndex = 1 To RowCount
                mCsrSector(RowIndex, ColumnIndex) = TableNumber(TableData, RowIndex, ColumnIndex + 1, SH_P_CSR, _
                    "CSR_GAMMA_SECTOR")
            Next ColumnIndex
        Next RowIndex

End Sub

Private Sub LoadEquity()

'
'------------------------------------------------------------------------------
'                          LOAD THE EQUITY PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the equity parameters of P_EQ: the bucket table, the spot-repo correlation and
'   the cross-bucket gamma matrix.
'
' WHY THIS EXISTS
'   EqRiskWeight, EqNameCorrelation, EqRhoSpotRepo, BucketGamma, IsValidBucket and
'   IsOtherBucket answer from these values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - reads EQ_BUCKETS: the highest bucket number sets the array size, then the type,
'       spot risk weight, repo risk weight and name correlation of each bucket are stored
'       by bucket number
'     - reads RHO_SPOT_REPO from EQ_PARAMS
'     - reads the EQ_GAMMA matrix, one row per bucket number in its first column
'
' ERROR POLICY
'   Raises (CORE_Util) when a table or a required key is missing or a value is blank or
'   not a number.
'
' DEPENDENCIES
'   ReadKeyValues, ReadTable, KVNumber, TableNumber, ToText (CORE_Util); SH_P_EQ
'   (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadEquity"

    Dim KeyValues               As Collection   'Key and value pairs of the EQ_PARAMS table
    Dim TableData               As Variant      'Rows of the table being read
    Dim RowCount                As Long         'Rows of the table
    Dim ColumnCount             As Long         'Columns of the table
    Dim RowIndex                As Long         'Row being read
    Dim ColumnIndex             As Long         'Gamma column being read
    Dim BucketNumber            As Long         'Bucket number of the row

'------------------------------------------------------------------------------
' READ BUCKETS
'------------------------------------------------------------------------------
    'Read the bucket table and find the highest bucket number
        TableData = ReadTable(SH_P_EQ, "EQ_BUCKETS", RowCount, ColumnCount)
        mEqN = 0
        For RowIndex = 1 To RowCount
            BucketNumber = CLng(TableNumber(TableData, RowIndex, 1, SH_P_EQ, "EQ_BUCKETS"))
            If BucketNumber > mEqN Then mEqN = BucketNumber
        Next RowIndex
    'Size the bucket arrays
        ReDim mEqType(1 To mEqN)
        ReDim mEqRWSpot(1 To mEqN)
        ReDim mEqRWRepo(1 To mEqN)
        ReDim mEqRhoName(1 To mEqN)
    'Store the type, risk weights and name correlation of each bucket
        For RowIndex = 1 To RowCount
            BucketNumber = CLng(TableNumber(TableData, RowIndex, 1, SH_P_EQ, "EQ_BUCKETS"))
            mEqType(BucketNumber) = ToText(TableData(RowIndex, 3))
            mEqRWSpot(BucketNumber) = TableNumber(TableData, RowIndex, 4, SH_P_EQ, "EQ_BUCKETS")
            mEqRWRepo(BucketNumber) = TableNumber(TableData, RowIndex, 5, SH_P_EQ, "EQ_BUCKETS")
            mEqRhoName(BucketNumber) = TableNumber(TableData, RowIndex, 6, SH_P_EQ, "EQ_BUCKETS")
        Next RowIndex

'------------------------------------------------------------------------------
' READ PARAMETERS
'------------------------------------------------------------------------------
    'Read the spot-repo correlation
        Set KeyValues = ReadKeyValues(SH_P_EQ, "EQ_PARAMS")
        mEqRhoSpotRepo = KVNumber(KeyValues, "RHO_SPOT_REPO", SH_P_EQ)

'------------------------------------------------------------------------------
' READ GAMMA MATRIX
'------------------------------------------------------------------------------
    'Read the gamma matrix and store each row under its bucket number
        TableData = ReadTable(SH_P_EQ, "EQ_GAMMA", RowCount, ColumnCount)
        ReDim mEqGamma(1 To mEqN, 1 To mEqN)
        For RowIndex = 1 To RowCount
            BucketNumber = CLng(TableNumber(TableData, RowIndex, 1, SH_P_EQ, "EQ_GAMMA"))
            For ColumnIndex = 1 To mEqN
                mEqGamma(BucketNumber, ColumnIndex) = TableNumber(TableData, RowIndex, ColumnIndex + 1, SH_P_EQ, _
                    "EQ_GAMMA")
            Next ColumnIndex
        Next RowIndex

End Sub

Private Sub LoadCommodity()

'
'------------------------------------------------------------------------------
'                        LOAD THE COMMODITY PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the commodity parameters of P_COM: the bucket table, the tenor and basis
'   correlations and the cross-bucket gamma matrix.
'
' WHY THIS EXISTS
'   ComRiskWeight, ComCommodityCorrelation, ComRhoTenor, ComRhoBasis, BucketGamma and
'   IsValidBucket answer from these values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - reads COM_BUCKETS: the highest bucket number sets the array size, then the risk
'       weight and the commodity correlation of each bucket are stored by bucket number
'     - reads RHO_TENOR and RHO_BASIS from COM_PARAMS
'     - reads the COM_GAMMA matrix, one row per bucket number in its first column
'
' ERROR POLICY
'   Raises (CORE_Util) when a table or a required key is missing or a value is blank or
'   not a number.
'
' DEPENDENCIES
'   ReadKeyValues, ReadTable, KVNumber, TableNumber (CORE_Util); SH_P_COM (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadCommodity"

    Dim KeyValues               As Collection   'Key and value pairs of the COM_PARAMS table
    Dim TableData               As Variant      'Rows of the table being read
    Dim RowCount                As Long         'Rows of the table
    Dim ColumnCount             As Long         'Columns of the table
    Dim RowIndex                As Long         'Row being read
    Dim ColumnIndex             As Long         'Gamma column being read
    Dim BucketNumber            As Long         'Bucket number of the row

'------------------------------------------------------------------------------
' READ BUCKETS
'------------------------------------------------------------------------------
    'Read the bucket table and find the highest bucket number
        TableData = ReadTable(SH_P_COM, "COM_BUCKETS", RowCount, ColumnCount)
        mComN = 0
        For RowIndex = 1 To RowCount
            BucketNumber = CLng(TableNumber(TableData, RowIndex, 1, SH_P_COM, "COM_BUCKETS"))
            If BucketNumber > mComN Then mComN = BucketNumber
        Next RowIndex
    'Size the bucket arrays
        ReDim mComRW(1 To mComN)
        ReDim mComRhoCty(1 To mComN)
    'Store the risk weight and commodity correlation of each bucket
        For RowIndex = 1 To RowCount
            BucketNumber = CLng(TableNumber(TableData, RowIndex, 1, SH_P_COM, "COM_BUCKETS"))
            mComRW(BucketNumber) = TableNumber(TableData, RowIndex, 3, SH_P_COM, "COM_BUCKETS")
            mComRhoCty(BucketNumber) = TableNumber(TableData, RowIndex, 4, SH_P_COM, "COM_BUCKETS")
        Next RowIndex

'------------------------------------------------------------------------------
' READ PARAMETERS
'------------------------------------------------------------------------------
    'Read the tenor and basis correlations
        Set KeyValues = ReadKeyValues(SH_P_COM, "COM_PARAMS")
        mComRhoTenor = KVNumber(KeyValues, "RHO_TENOR", SH_P_COM)
        mComRhoBasis = KVNumber(KeyValues, "RHO_BASIS", SH_P_COM)

'------------------------------------------------------------------------------
' READ GAMMA MATRIX
'------------------------------------------------------------------------------
    'Read the gamma matrix and store each row under its bucket number
        TableData = ReadTable(SH_P_COM, "COM_GAMMA", RowCount, ColumnCount)
        ReDim mComGamma(1 To mComN, 1 To mComN)
        For RowIndex = 1 To RowCount
            BucketNumber = CLng(TableNumber(TableData, RowIndex, 1, SH_P_COM, "COM_GAMMA"))
            For ColumnIndex = 1 To mComN
                mComGamma(BucketNumber, ColumnIndex) = TableNumber(TableData, RowIndex, ColumnIndex + 1, SH_P_COM, _
                    "COM_GAMMA")
            Next ColumnIndex
        Next RowIndex

End Sub

Private Sub LoadFx()

'
'------------------------------------------------------------------------------
'                            LOAD THE FX PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the FX parameters of P_FX: the risk weight, the gamma, the two divisors and the
'   specified currency pairs.
'
' WHY THIS EXISTS
'   FxRiskWeight, FxCurvatureDivisor, BucketGamma and the specified-pair test answer from
'   these values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - reads RW, GAMMA, SPECIFIED_DIVISOR and CURVATURE_DIVISOR from FX_PARAMS
'     - rebuilds the specified pairs from FX_SPECIFIED_PAIRS, adding each pair once
'
' ERROR POLICY
'   Raises (CORE_Util) when a table or a required key is missing or a value is blank or
'   not a number.
'
' DEPENDENCIES
'   ReadKeyValues, ReadTable, KVNumber, ToText, KeyExists, AddKey (CORE_Util); SH_P_FX
'   (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadFx"

    Dim KeyValues               As Collection   'Key and value pairs of the FX_PARAMS table
    Dim TableData               As Variant      'Rows of the FX_SPECIFIED_PAIRS table
    Dim RowCount                As Long         'Rows of the table
    Dim ColumnCount             As Long         'Columns of the table
    Dim RowIndex                As Long         'Row being read

'------------------------------------------------------------------------------
' READ PARAMETERS
'------------------------------------------------------------------------------
    'Read the risk weight, the gamma and the two divisors
        Set KeyValues = ReadKeyValues(SH_P_FX, "FX_PARAMS")
        mFxRW = KVNumber(KeyValues, "RW", SH_P_FX)
        mFxGamma = KVNumber(KeyValues, "GAMMA", SH_P_FX)
        mFxSpecDivisor = KVNumber(KeyValues, "SPECIFIED_DIVISOR", SH_P_FX)
        mFxCurvDivisor = KVNumber(KeyValues, "CURVATURE_DIVISOR", SH_P_FX)

'------------------------------------------------------------------------------
' READ SPECIFIED PAIRS
'------------------------------------------------------------------------------
    'Rebuild the specified pairs, keeping each pair once
        Set mFxSpecPairs = New Collection
        TableData = ReadTable(SH_P_FX, "FX_SPECIFIED_PAIRS", RowCount, ColumnCount)
        For RowIndex = 1 To RowCount
            If Not KeyExists(mFxSpecPairs, ToText(TableData(RowIndex, 1))) Then
                AddKey mFxSpecPairs, ToText(TableData(RowIndex, 1)), True
            End If
        Next RowIndex

End Sub

Private Sub LoadDrc()

'
'------------------------------------------------------------------------------
'                           LOAD THE DRC PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the DRC parameters of P_DRC: the risk weights by credit quality, the seniority
'   ranks and loss given default, the buckets and the maturity limits.
'
' WHY THIS EXISTS
'   The Drc* accessors answer from these values.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - reads DRC_RW into a collection of risk weights by credit quality
'     - reads DRC_SENIORITY into a rank collection and a loss given default collection by
'       seniority
'     - reads DRC_BUCKETS into a collection of positions by bucket name and an array of
'       the names
'     - reads MIN_MATURITY and MAX_MATURITY from DRC_PARAMS
'
' ERROR POLICY
'   Raises (CORE_Util) when a table or a required key is missing or a value is blank or
'   not a number.
'
' DEPENDENCIES
'   ReadKeyValues, ReadTable, KVNumber, TableNumber, ToText, AddKey (CORE_Util); SH_P_DRC
'   (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadDrc"

    Dim KeyValues               As Collection   'Key and value pairs of the DRC_PARAMS table
    Dim TableData               As Variant      'Rows of the table being read
    Dim RowCount                As Long         'Rows of the table
    Dim ColumnCount             As Long         'Columns of the table
    Dim RowIndex                As Long         'Row being read

'------------------------------------------------------------------------------
' READ RISK WEIGHTS
'------------------------------------------------------------------------------
    'Read the risk weight of each credit quality
        Set mDrcRW = New Collection
        TableData = ReadTable(SH_P_DRC, "DRC_RW", RowCount, ColumnCount)
        For RowIndex = 1 To RowCount
            AddKey mDrcRW, ToText(TableData(RowIndex, 1)), TableNumber(TableData, RowIndex, 2, SH_P_DRC, "DRC_RW")
        Next RowIndex

'------------------------------------------------------------------------------
' READ SENIORITIES
'------------------------------------------------------------------------------
    'Read the rank and the loss given default of each seniority
        Set mDrcLGD = New Collection
        Set mDrcRank = New Collection
        TableData = ReadTable(SH_P_DRC, "DRC_SENIORITY", RowCount, ColumnCount)
        For RowIndex = 1 To RowCount
            AddKey mDrcRank, ToText(TableData(RowIndex, 1)), _
                CLng(TableNumber(TableData, RowIndex, 2, SH_P_DRC, "DRC_SENIORITY"))
            AddKey mDrcLGD, ToText(TableData(RowIndex, 1)), TableNumber(TableData, RowIndex, 3, SH_P_DRC, "DRC_SENIORITY")
        Next RowIndex

'------------------------------------------------------------------------------
' READ BUCKETS
'------------------------------------------------------------------------------
    'Read the bucket names and index them by position
        Set mDrcBuckets = New Collection
        TableData = ReadTable(SH_P_DRC, "DRC_BUCKETS", RowCount, ColumnCount)
        ReDim mDrcBucketNames(1 To RowCount)
        For RowIndex = 1 To RowCount
            AddKey mDrcBuckets, ToText(TableData(RowIndex, 1)), RowIndex
            mDrcBucketNames(RowIndex) = ToText(TableData(RowIndex, 1))
        Next RowIndex

'------------------------------------------------------------------------------
' READ PARAMETERS
'------------------------------------------------------------------------------
    'Read the minimum and maximum maturity
        Set KeyValues = ReadKeyValues(SH_P_DRC, "DRC_PARAMS")
        mDrcMinMat = KVNumber(KeyValues, "MIN_MATURITY", SH_P_DRC)
        mDrcMaxMat = KVNumber(KeyValues, "MAX_MATURITY", SH_P_DRC)

End Sub

Private Sub LoadRrao()

'
'------------------------------------------------------------------------------
'                           LOAD THE RRAO PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the RRAO risk weight of each category from P_RRAO.
'
' WHY THIS EXISTS
'   RraoRiskWeight answers from this collection.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the RRAO_RW table and stores the risk weight of each category in a collection
'   keyed by category.
'
' ERROR POLICY
'   Raises (CORE_Util) when the table is missing or a risk weight is blank or not a
'   number.
'
' DEPENDENCIES
'   ReadTable, TableNumber, ToText, AddKey (CORE_Util); SH_P_RRAO (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.LoadRrao"

    Dim TableData               As Variant  'Rows of the RRAO_RW table
    Dim RowCount                As Long     'Rows of the table
    Dim ColumnCount             As Long     'Columns of the table
    Dim RowIndex                As Long     'Row being read

'------------------------------------------------------------------------------
' READ RISK WEIGHTS
'------------------------------------------------------------------------------
    'Read the risk weight of each RRAO category
        Set mRraoRW = New Collection
        TableData = ReadTable(SH_P_RRAO, "RRAO_RW", RowCount, ColumnCount)
        For RowIndex = 1 To RowCount
            AddKey mRraoRW, ToText(TableData(RowIndex, 1)), TableNumber(TableData, RowIndex, 2, SH_P_RRAO, "RRAO_RW")
        Next RowIndex

End Sub

'
'------------------------------------------------------------------------------
'
'                        SETTINGS AND GENERAL ACCESSORS
'
'------------------------------------------------------------------------------
'

Public Function ReportingCcy() As String

'
'------------------------------------------------------------------------------
'                        RETURN THE REPORTING CURRENCY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the reporting currency of the run.
'
' WHY THIS EXISTS
'   Callers compare currencies with the reporting currency; the value stays private to the
'   module.
'
' INPUTS
'   None
'
' RETURNS
'   The reporting currency in upper case, EUR when the setting is absent.
'
' BEHAVIOR
'   Returns the currency stored by LoadSettings.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        ReportingCcy = mReportingCcy

End Function

Public Function FxCurvatureDivisor(ByVal refReportingCcy As Boolean) As Double

'
'------------------------------------------------------------------------------
'                       RESOLVE THE FX CURVATURE DIVISOR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the divisor that applies to an FX curvature charge under the supervisory
'   discretion.
'
' WHY THIS EXISTS
'   The discretion depends on two switches of the Control sheet and on whether the
'   instrument references the reporting currency.
'
' INPUTS
'   refReportingCcy
'     True when the instrument references the reporting currency
'
' RETURNS
'   The FX curvature divisor when the discretion applies, otherwise 1.
'
' BEHAVIOR
'   Returns the divisor when:
'     - the switch for all FX instruments is on, or
'     - the switch for options is on and the instrument does not reference the reporting
'       currency
'   Returns 1 otherwise.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   MAR21.98: options not referencing the reporting currency may be divided by 1.5; all FX
'   instruments only with supervisory approval (the _ALL switch).
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Apply the divisor to all FX instruments, or to options not referencing the reporting currency
        If mFxCurvDivisorAllOn Or (mFxCurvDivisorOn And Not refReportingCcy) Then
            FxCurvatureDivisor = mFxCurvDivisor
        Else
            FxCurvatureDivisor = 1#
        End If

End Function

Public Function ScenarioCorrelation(ByVal rho As Double, ByVal scen As Long) As Double

'
'------------------------------------------------------------------------------
'                     TRANSFORM A CORRELATION BY SCENARIO
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a correlation transformed for a correlation scenario.
'
' WHY THIS EXISTS
'   Callers hold a medium-scenario correlation and need the one of the scenario they are
'   calculating; the transform is kept in one place.
'
' INPUTS
'   rho
'     Correlation in the medium scenario
'
'   scen
'     Correlation scenario, one of the SCEN_* constants
'
' RETURNS
'   The correlation for the scenario; rho itself for the medium scenario or any other
'   value.
'
' BEHAVIOR
'   Transforms rho by scenario:
'     - SCEN_HIGH: min(HIGH_MULTIPLIER * rho, 1)
'     - SCEN_LOW: max(2 * rho - 1, LOW_MULTIPLIER * rho)
'     - any other value: rho
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   MinD, MaxD (CORE_Util); SCEN_HIGH, SCEN_LOW (CORE_Config).
'
' NOTES
'   Correlation scenario transform (MAR21.6): high = min(1.25 rho, 1), low = max(2 rho -
'   1, 0.75 rho), medium = rho. The multipliers 1.25 and 0.75 are read from P_General.
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Transform the correlation for the scenario
        Select Case scen
            Case SCEN_HIGH
                ScenarioCorrelation = MinD(mHighMult * rho, 1#)
            Case SCEN_LOW
                ScenarioCorrelation = MaxD(2# * rho - 1#, mLowMult * rho)
            Case Else
                ScenarioCorrelation = rho
        End Select

End Function

Public Function VegaAlpha() As Double

'
'------------------------------------------------------------------------------
'                            RETURN THE VEGA ALPHA
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the vega alpha parameter.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The ALPHA value of the VEGA table of P_General.
'
' BEHAVIOR
'   Returns the value stored by LoadGeneral.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        VegaAlpha = mVegaAlpha

End Function

Public Function VegaRiskWeight(ByVal rc As String, ByVal bucket As String) As Double

'
'------------------------------------------------------------------------------
'                         RESOLVE THE VEGA RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the vega risk weight of a risk class and bucket.
'
' WHY THIS EXISTS
'   The vega risk weight depends on the liquidity horizon of the risk class and bucket.
'
' INPUTS
'   rc
'     Risk class code, one of the RC_* constants
'
'   bucket
'     Bucket label
'
' RETURNS
'   The vega risk weight, at most 1.
'
' BEHAVIOR
'   In order:
'     - looks for a liquidity horizon row with the same risk class and bucket
'     - otherwise looks for a row of the risk class whose bucket is "*", which matches
'       every bucket
'     - raises when neither exists
'     - returns min(RW_sigma * sqrt(LH / 10), 1)
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP, with the source "VegaRiskWeight", when no liquidity
'   horizon is found.
'
' DEPENDENCIES
'   MinD (CORE_Util); ERR_APP (CORE_Util).
'
' NOTES
'   Vega risk weight: min(RW_sigma * sqrt(LH / 10), 100%) with the liquidity horizon
'   looked up by risk class and bucket ("*" matches every bucket).
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.VegaRiskWeight"

    Dim RowIndex                As Long     'Liquidity horizon row being tested
    Dim LiquidityHorizon        As Double   'Liquidity horizon found
    Dim HorizonFound            As Boolean  'True when a row matched

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start without a matching horizon
        HorizonFound = False

'------------------------------------------------------------------------------
' FIND THE EXACT BUCKET
'------------------------------------------------------------------------------
    'Look for a row with the same risk class and bucket
        For RowIndex = 1 To mLHCount
            If mLHClass(RowIndex) = rc And mLHBucket(RowIndex) = bucket Then
                LiquidityHorizon = mLHValue(RowIndex)
                HorizonFound = True
                Exit For
            End If
        Next RowIndex

'------------------------------------------------------------------------------
' FIND THE WILDCARD BUCKET
'------------------------------------------------------------------------------
    'Fall back to the row of the risk class that matches every bucket
        If Not HorizonFound Then
            For RowIndex = 1 To mLHCount
                If mLHClass(RowIndex) = rc And mLHBucket(RowIndex) = "*" Then
                    LiquidityHorizon = mLHValue(RowIndex)
                    HorizonFound = True
                    Exit For
                End If
            Next RowIndex
        End If

'------------------------------------------------------------------------------
' VALIDATE HORIZON
'------------------------------------------------------------------------------
    'Reject a risk class and bucket without a liquidity horizon
        If Not HorizonFound Then
            Err.Raise vbObjectError + ERR_APP, "VegaRiskWeight", _
                "No liquidity horizon for " & rc & " bucket " & bucket
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Scale the sigma by the square root of the horizon, capped at 100%
        VegaRiskWeight = MinD(mVegaRWSigma * Sqr(LiquidityHorizon / 10#), 1#)

End Function

Public Function IsValidGirrTenor(ByVal t As Double) As Boolean

'
'------------------------------------------------------------------------------
'                           TEST A GIRR VERTEX TENOR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a tenor is one of the GIRR vertex tenors.
'
' WHY THIS EXISTS
'   Inputs are checked against the vertex list of the P_General sheet.
'
' INPUTS
'   t
'     Tenor or maturity in years
'
' RETURNS
'   True when t matches a listed vertex, False otherwise.
'
' BEHAVIOR
'   Compares t with each loaded vertex, with a tolerance of 0.000001.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   InNumberList (CORE_Util).
'
' NOTES
'   The list is loaded by LoadGeneral from GIRR_TENORS.
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Test the tenor against the loaded vertex list
        IsValidGirrTenor = InNumberList(t, mGirrTenors, mGirrTenorN)

End Function

Public Function IsValidCsrTenor(ByVal t As Double) As Boolean

'
'------------------------------------------------------------------------------
'                           TEST A CSR VERTEX TENOR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a tenor is one of the CSR vertex tenors.
'
' WHY THIS EXISTS
'   Inputs are checked against the vertex list of the P_General sheet.
'
' INPUTS
'   t
'     Tenor or maturity in years
'
' RETURNS
'   True when t matches a listed vertex, False otherwise.
'
' BEHAVIOR
'   Compares t with each loaded vertex, with a tolerance of 0.000001.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   InNumberList (CORE_Util).
'
' NOTES
'   The list is loaded by LoadGeneral from CSR_TENORS.
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Test the tenor against the loaded vertex list
        IsValidCsrTenor = InNumberList(t, mCsrTenors, mCsrTenorN)

End Function

Public Function IsValidComTenor(ByVal t As Double) As Boolean

'
'------------------------------------------------------------------------------
'                        TEST A COMMODITY VERTEX TENOR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a tenor is one of the commodity vertex tenors.
'
' WHY THIS EXISTS
'   Inputs are checked against the vertex list of the P_General sheet.
'
' INPUTS
'   t
'     Tenor or maturity in years
'
' RETURNS
'   True when t matches a listed vertex, False otherwise.
'
' BEHAVIOR
'   Compares t with each loaded vertex, with a tolerance of 0.000001.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   InNumberList (CORE_Util).
'
' NOTES
'   The list is loaded by LoadGeneral from COM_TENORS.
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Test the tenor against the loaded vertex list
        IsValidComTenor = InNumberList(t, mComTenors, mComTenorN)

End Function

Public Function IsValidVegaMaturity(ByVal t As Double) As Boolean

'
'------------------------------------------------------------------------------
'                         TEST A VEGA VERTEX MATURITY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a tenor is one of the vega vertex maturities.
'
' WHY THIS EXISTS
'   Inputs are checked against the vertex list of the P_General sheet.
'
' INPUTS
'   t
'     Tenor or maturity in years
'
' RETURNS
'   True when t matches a listed vertex, False otherwise.
'
' BEHAVIOR
'   Compares t with each loaded vertex, with a tolerance of 0.000001.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   InNumberList (CORE_Util).
'
' NOTES
'   The list is loaded by LoadGeneral from VEGA_MATURITIES.
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Test the tenor against the loaded vertex list
        IsValidVegaMaturity = InNumberList(t, mVegaMats, mVegaMatN)

End Function

'
'------------------------------------------------------------------------------
'
'                     BUCKET VALIDATION AND CLASSIFICATION
'
'------------------------------------------------------------------------------
'

Public Function IsValidBucket(ByVal rc As String, ByVal bucket As String) As Boolean

'
'------------------------------------------------------------------------------
'                             TEST A BUCKET LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a bucket label is valid for a risk class.
'
' WHY THIS EXISTS
'   Each risk class labels its buckets differently, so the check depends on the risk class
'   and on the loaded bucket tables.
'
' INPUTS
'   rc
'     Risk class code, one of the RC_* constants
'
'   bucket
'     Bucket label
'
' RETURNS
'   True for a valid bucket, False otherwise, including any other risk class.
'
' BEHAVIOR
'   By risk class:
'     - GIRR and FX: the label is three characters long
'     - CSR and equity: the label is a number from 1 to the highest loaded bucket, and the
'       bucket has a type
'     - commodity: the label is a number from 1 to the highest loaded bucket
'     - any other risk class: False
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsNumber (CORE_Util); RC_GIRR, RC_FX, RC_CSR, RC_EQ, RC_COM (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.IsValidBucket"

    Dim BucketNumber            As Long 'Numeric value of the bucket label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start with an invalid bucket
        IsValidBucket = False

'------------------------------------------------------------------------------
' VALIDATE BUCKET
'------------------------------------------------------------------------------
    'Validate the label according to the risk class
        Select Case rc
            Case RC_GIRR, RC_FX
                IsValidBucket = (Len(bucket) = 3)
            Case RC_CSR, RC_EQ, RC_COM
                If Not IsNumber(bucket) Then Exit Function
                BucketNumber = CLng(Val(bucket))
                If BucketNumber < 1 Then Exit Function
                Select Case rc
                    Case RC_CSR
                        If BucketNumber <= mCsrN Then IsValidBucket = (mCsrType(BucketNumber) <> vbNullString)
                    Case RC_EQ
                        If BucketNumber <= mEqN Then IsValidBucket = (mEqType(BucketNumber) <> vbNullString)
                    Case RC_COM
                        IsValidBucket = (BucketNumber <= mComN)
                    Case Else
                        'Nothing is done for any other risk class
                End Select
            Case Else
                'Nothing is done for any other risk class
        End Select

End Function

Public Function IsOtherBucket(ByVal rc As String, ByVal bucket As String) As Boolean

'
'------------------------------------------------------------------------------
'                       TEST FOR AN OTHER SECTOR BUCKET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a bucket is the "other sector" bucket of its risk class.
'
' WHY THIS EXISTS
'   The other sector bucket is aggregated without diversification, so callers need to tell
'   it apart.
'
' INPUTS
'   rc
'     Risk class code, one of the RC_* constants
'
'   bucket
'     Bucket label, a valid number for CSR and equity
'
' RETURNS
'   True when the bucket type is OTHER for CSR or equity, False otherwise.
'
' BEHAVIOR
'   Looks up the type of the bucket for CSR and equity and compares it with OTHER; any
'   other risk class returns False.
'
' ERROR POLICY
'   None of its own; a CSR or equity bucket outside the loaded range raises a VBA
'   subscript error.
'
' DEPENDENCIES
'   RC_CSR, RC_EQ (CORE_Config).
'
' NOTES
'   "Other sector" buckets (CSR 16, equity 11) are aggregated without diversification and
'   added outside the cross-bucket square root.
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start with a regular bucket
        IsOtherBucket = False

'------------------------------------------------------------------------------
' CLASSIFY BUCKET
'------------------------------------------------------------------------------
    'Compare the bucket type with OTHER for CSR and equity
        Select Case rc
            Case RC_CSR
                IsOtherBucket = (mCsrType(CLng(Val(bucket))) = "OTHER")
            Case RC_EQ
                IsOtherBucket = (mEqType(CLng(Val(bucket))) = "OTHER")
            Case Else
                'Nothing is done for any other risk class
        End Select

End Function

'
'------------------------------------------------------------------------------
'
'                              DELTA RISK WEIGHTS
'
'------------------------------------------------------------------------------
'

Public Function GirrRiskWeight(ByVal ccy As String, ByVal subType As String, ByVal tenor As Double) As Double

'
'------------------------------------------------------------------------------
'                         RESOLVE THE GIRR RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the GIRR delta risk weight of a currency, sub-type and tenor.
'
' WHY THIS EXISTS
'   The weight depends on the sub-type and the tenor, and a divisor applies to the
'   specified currencies and the reporting currency when the discretion is on.
'
' INPUTS
'   ccy
'     Currency code
'
'   subType
'     Risk factor sub-type: INFLATION, XCCY or any other value for a tenor weight
'
'   tenor
'     Tenor in years, used for the tenor weights
'
' RETURNS
'   The risk weight, divided by the specified divisor when the discretion applies.
'
' BEHAVIOR
'   In order:
'     - INFLATION returns the inflation weight and XCCY the cross-currency basis weight
'     - any other sub-type looks up the tenor in the GIRR_RW table, with a tolerance of
'       0.000001, and raises when it is not listed
'     - when the GIRR specified-currency discretion is on and the currency is a specified
'       currency or the reporting currency, divides the weight by the specified divisor
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP, with the source "GirrRiskWeight", when the tenor has
'   no risk weight.
'
' DEPENDENCIES
'   KeyExists (CORE_Util); ERR_APP (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.GirrRiskWeight"

    Dim TenorIndex              As Long     'Risk weight row being tested
    Dim Weight                  As Double   'Risk weight being resolved

'------------------------------------------------------------------------------
' RESOLVE WEIGHT
'------------------------------------------------------------------------------
    'Pick the weight of the sub-type, or look the tenor up in the risk weight table
        Select Case subType
            Case "INFLATION"
                Weight = mGirrRWInflation
            Case "XCCY"
                Weight = mGirrRWXccy
            Case Else
                Weight = -1#
                For TenorIndex = 1 To mGirrRWCount
                    If Abs(mGirrRWTenor(TenorIndex) - tenor) < 0.000001 Then
                        Weight = mGirrRWValue(TenorIndex)
                        Exit For
                    End If
                Next TenorIndex
                If Weight < 0# Then
                    Err.Raise vbObjectError + ERR_APP, "GirrRiskWeight", _
                        "No GIRR risk weight for tenor " & tenor
                End If
        End Select

'------------------------------------------------------------------------------
' APPLY SPECIFIED DIVISOR
'------------------------------------------------------------------------------
    'MAR21.44: "the above risk weights" covers the tenor, inflation and cross-currency
    'basis weights of MAR21.42-43 (#29)
        If mGirrSpecifiedOn Then
            If KeyExists(mGirrSpecCcy, ccy) Or ccy = mReportingCcy Then Weight = Weight / mGirrSpecDivisor
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the weight
        GirrRiskWeight = Weight

End Function

Public Function CsrRiskWeight(ByVal bucket As String) As Double

'
'------------------------------------------------------------------------------
'                         RESOLVE THE CSR RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the CSR delta risk weight of a bucket.
'
' WHY THIS EXISTS
'   The bucket risk weights are loaded from P_CSR and stay private to the module.
'
' INPUTS
'   bucket
'     Bucket label, a valid number
'
' RETURNS
'   The risk weight of the bucket.
'
' BEHAVIOR
'   Returns the weight stored for the bucket number.
'
' ERROR POLICY
'   None of its own; a bucket outside the loaded range raises a VBA subscript error.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Look the value up by bucket number
        CsrRiskWeight = mCsrRW(CLng(Val(bucket)))

End Function

Public Function EqRiskWeight(ByVal bucket As String, ByVal subType As String) As Double

'
'------------------------------------------------------------------------------
'                        RESOLVE THE EQUITY RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the equity delta risk weight of a bucket and sub-type.
'
' WHY THIS EXISTS
'   Equity has a spot and a repo risk weight for each bucket, both loaded from P_EQ.
'
' INPUTS
'   bucket
'     Bucket label, a valid number
'
'   subType
'     Risk factor sub-type: REPO for the repo weight, any other value for the spot weight
'
' RETURNS
'   The repo risk weight for REPO, otherwise the spot risk weight of the bucket.
'
' BEHAVIOR
'   Returns the repo weight when subType is REPO, otherwise the spot weight, by bucket
'   number.
'
' ERROR POLICY
'   None of its own; a bucket outside the loaded range raises a VBA subscript error.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the repo weight for REPO, the spot weight otherwise
        If subType = "REPO" Then
            EqRiskWeight = mEqRWRepo(CLng(Val(bucket)))
        Else
            EqRiskWeight = mEqRWSpot(CLng(Val(bucket)))
        End If

End Function

Public Function ComRiskWeight(ByVal bucket As String) As Double

'
'------------------------------------------------------------------------------
'                      RESOLVE THE COMMODITY RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the commodity delta risk weight of a bucket.
'
' WHY THIS EXISTS
'   The bucket risk weights are loaded from P_COM and stay private to the module.
'
' INPUTS
'   bucket
'     Bucket label, a valid number
'
' RETURNS
'   The risk weight of the bucket.
'
' BEHAVIOR
'   Returns the weight stored for the bucket number.
'
' ERROR POLICY
'   None of its own; a bucket outside the loaded range raises a VBA subscript error.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Look the value up by bucket number
        ComRiskWeight = mComRW(CLng(Val(bucket)))

End Function

Public Function FxRiskWeight(ByVal ccy As String) As Double

'
'------------------------------------------------------------------------------
'                          RESOLVE THE FX RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the FX delta risk weight of a currency.
'
' WHY THIS EXISTS
'   The weight is divided for a specified currency pair when the discretion is on.
'
' INPUTS
'   ccy
'     Currency code, paired with the reporting currency
'
' RETURNS
'   The FX risk weight, divided by the specified divisor when the pair is specified.
'
' BEHAVIOR
'   Starts from the FX risk weight; when the specified-pair discretion is on and the pair
'   of ccy and the reporting currency is specified, divides it by the specified divisor.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsSpecifiedFxPair.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.FxRiskWeight"

    Dim Weight                  As Double   'Risk weight being resolved

'------------------------------------------------------------------------------
' RESOLVE WEIGHT
'------------------------------------------------------------------------------
    'Start from the FX risk weight
        Weight = mFxRW
    'Divide the weight when the pair with the reporting currency is a specified pair
        If mFxSpecifiedOn Then
            If IsSpecifiedFxPair(ccy, mReportingCcy) Then Weight = Weight / mFxSpecDivisor
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the weight
        FxRiskWeight = Weight

End Function

Private Function IsSpecifiedFxPair(ByVal ccyA As String, ByVal ccyB As String) As Boolean

'
'------------------------------------------------------------------------------
'                         TEST FOR A SPECIFIED FX PAIR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a currency pair is a specified pair: it is listed in either
'   direction, or both currencies are paired with USD.
'
' INPUTS
'   ccyA
'     First currency code
'
'   ccyB
'     Second currency code
'
' RETURNS
'   True when the pair is specified.
'
' NOTES
'   A pair is specified when it is listed, or when it is a first-order cross of two
'   currencies that are each listed against USD (MAR21.88, fn 23).
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
    'A listed pair is specified in either direction; otherwise both currencies must be paired with USD
        If KeyExists(mFxSpecPairs, ccyA & "/" & ccyB) Or KeyExists(mFxSpecPairs, ccyB & "/" & ccyA) Then
            IsSpecifiedFxPair = True
        Else
            IsSpecifiedFxPair = IsUsdPaired(ccyA) And IsUsdPaired(ccyB)
        End If

End Function

Private Function IsUsdPaired(ByVal ccy As String) As Boolean

'
'------------------------------------------------------------------------------
'                     TEST FOR A CURRENCY PAIRED WITH USD
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a currency is USD or is listed against USD in either direction.
'
' INPUTS
'   ccy
'     Currency code
'
' RETURNS
'   True when ccy is USD or USD/ccy or ccy/USD is a specified pair.
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
    'Test for USD itself, then for the pair in either direction
        IsUsdPaired = (ccy = "USD") Or KeyExists(mFxSpecPairs, "USD/" & ccy) Or KeyExists(mFxSpecPairs, ccy & "/USD")

End Function

'
'------------------------------------------------------------------------------
'
'                      DELTA CORRELATION BUILDING BLOCKS
'
'------------------------------------------------------------------------------
'

Public Function GirrTenorCorrelation(ByVal ta As Double, ByVal tb As Double) As Double

'
'------------------------------------------------------------------------------
'                      RESOLVE THE GIRR TENOR CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the medium-scenario correlation between two GIRR tenors.
'
' WHY THIS EXISTS
'   The tenor correlation decays with the distance between the tenors, down to a floor.
'
' INPUTS
'   ta
'     First tenor in years
'
'   tb
'     Second tenor in years
'
' RETURNS
'   1 for equal tenors, otherwise the correlation, not below the floor.
'
' BEHAVIOR
'   Returns:
'     - 1 when the tenors differ by less than 0.000001
'     - otherwise max(exp(-theta * |ta - tb| / min(ta, tb)), floor), with theta and the
'       floor read from GIRR_PARAMS
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   MaxD, MinD (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return 1 for equal tenors, otherwise the decayed correlation with its floor
        If Abs(ta - tb) < 0.000001 Then
            GirrTenorCorrelation = 1#
        Else
            GirrTenorCorrelation = MaxD(Exp(-mGirrTheta * Abs(ta - tb) / MinD(ta, tb)), mGirrRhoFloor)
        End If

End Function

Public Function GirrRhoDiffCurve() As Double

'
'------------------------------------------------------------------------------
'                 RETURN THE GIRR DIFFERENT-CURVE CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the GIRR correlation between different curves.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The RHO_DIFF_CURVE value of GIRR_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadGirr.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        GirrRhoDiffCurve = mGirrRhoDiffCurve

End Function

Public Function GirrRhoInflation() As Double

'
'------------------------------------------------------------------------------
'                    RETURN THE GIRR INFLATION CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the GIRR correlation of the inflation risk factor.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The RHO_INFLATION value of GIRR_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadGirr.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        GirrRhoInflation = mGirrRhoInflation

End Function

Public Function GirrRhoXccy() As Double

'
'------------------------------------------------------------------------------
'                  RETURN THE GIRR CROSS-CURRENCY CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the GIRR correlation of the cross-currency basis risk factor.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The RHO_XCCY value of GIRR_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadGirr.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        GirrRhoXccy = mGirrRhoXccy

End Function

Public Function CsrNameCorrelation(ByVal bucket As String) As Double

'
'------------------------------------------------------------------------------
'                       RESOLVE THE CSR NAME CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the CSR correlation between two names of one bucket.
'
' WHY THIS EXISTS
'   Index buckets use their own name correlation.
'
' INPUTS
'   bucket
'     Bucket label, a valid number
'
' RETURNS
'   The index name correlation for an INDEX bucket, otherwise the name correlation.
'
' BEHAVIOR
'   Looks up the type of the bucket: INDEX returns RHO_NAME_INDEX, any other type returns
'   RHO_NAME.
'
' ERROR POLICY
'   None of its own; a bucket outside the loaded range raises a VBA subscript error.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the index correlation for an index bucket, the name correlation otherwise
        If mCsrType(CLng(Val(bucket))) = "INDEX" Then
            CsrNameCorrelation = mCsrRhoNameIndex
        Else
            CsrNameCorrelation = mCsrRhoName
        End If

End Function

Public Function CsrRhoTenor() As Double

'
'------------------------------------------------------------------------------
'                       RETURN THE CSR TENOR CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the CSR correlation between tenors.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The RHO_TENOR value of CSR_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadCsr.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        CsrRhoTenor = mCsrRhoTenor

End Function

Public Function CsrRhoBasis() As Double

'
'------------------------------------------------------------------------------
'                       RETURN THE CSR BASIS CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the CSR correlation between bases.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The RHO_BASIS value of CSR_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadCsr.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        CsrRhoBasis = mCsrRhoBasis

End Function

Public Function EqNameCorrelation(ByVal bucket As String) As Double

'
'------------------------------------------------------------------------------
'                     RESOLVE THE EQUITY NAME CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the equity correlation between two names of one bucket.
'
' WHY THIS EXISTS
'   The bucket name correlations are loaded from P_EQ and stay private to the module.
'
' INPUTS
'   bucket
'     Bucket label, a valid number
'
' RETURNS
'   The name correlation of the bucket.
'
' BEHAVIOR
'   Returns the value stored for the bucket number.
'
' ERROR POLICY
'   None of its own; a bucket outside the loaded range raises a VBA subscript error.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Look the value up by bucket number
        EqNameCorrelation = mEqRhoName(CLng(Val(bucket)))

End Function

Public Function EqRhoSpotRepo() As Double

'
'------------------------------------------------------------------------------
'                   RETURN THE EQUITY SPOT-REPO CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the equity correlation between the spot and repo risk factors.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The RHO_SPOT_REPO value of EQ_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadEquity.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        EqRhoSpotRepo = mEqRhoSpotRepo

End Function

Public Function ComCommodityCorrelation(ByVal bucket As String) As Double

'
'------------------------------------------------------------------------------
'                      RESOLVE THE COMMODITY CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the commodity correlation between two commodities of one bucket.
'
' WHY THIS EXISTS
'   The bucket correlations are loaded from P_COM and stay private to the module.
'
' INPUTS
'   bucket
'     Bucket label, a valid number
'
' RETURNS
'   The commodity correlation of the bucket.
'
' BEHAVIOR
'   Returns the value stored for the bucket number.
'
' ERROR POLICY
'   None of its own; a bucket outside the loaded range raises a VBA subscript error.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Look the value up by bucket number
        ComCommodityCorrelation = mComRhoCty(CLng(Val(bucket)))

End Function

Public Function ComRhoTenor() As Double

'
'------------------------------------------------------------------------------
'                    RETURN THE COMMODITY TENOR CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the commodity correlation between tenors.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The RHO_TENOR value of COM_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadCommodity.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        ComRhoTenor = mComRhoTenor

End Function

Public Function ComRhoBasis() As Double

'
'------------------------------------------------------------------------------
'                    RETURN THE COMMODITY BASIS CORRELATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the commodity correlation between bases.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The RHO_BASIS value of COM_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadCommodity.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        ComRhoBasis = mComRhoBasis

End Function

'
'------------------------------------------------------------------------------
'
'                          CROSS-BUCKET CORRELATIONS
'
'------------------------------------------------------------------------------
'

Public Function BucketGamma(ByVal rc As String, ByVal b As String, ByVal c As String) As Double

'
'------------------------------------------------------------------------------
'                        RESOLVE THE CROSS-BUCKET GAMMA
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the medium-scenario correlation (gamma) between two buckets of a risk class.
'
' WHY THIS EXISTS
'   Each risk class stores its cross-bucket correlations differently: a single value, or a
'   matrix by bucket.
'
' INPUTS
'   rc
'     Risk class code, one of the RC_* constants
'
'   b
'     First bucket label
'
'   c
'     Second bucket label
'
' RETURNS
'   The gamma of the two buckets, 0 for any other risk class.
'
' BEHAVIOR
'   By risk class:
'     - GIRR and FX: the single gamma of the class
'     - CSR: the sector matrix value of the two bucket groups, times the rating factor
'       when both buckets are of type NAME and their credit qualities differ
'     - equity and commodity: the matrix value of the two bucket numbers
'     - any other risk class: 0
'
' ERROR POLICY
'   None of its own; a bucket outside the loaded range raises a VBA subscript error.
'
' DEPENDENCIES
'   KeyIndex (CORE_Util); RC_GIRR, RC_FX, RC_CSR, RC_EQ, RC_COM (CORE_Config).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.BucketGamma"

    Dim IndexB                  As Long     'Bucket number of the first bucket
    Dim IndexC                  As Long     'Bucket number of the second bucket
    Dim RatingFactor            As Double   'Rating factor of the CSR pair

'------------------------------------------------------------------------------
' RESOLVE GAMMA
'------------------------------------------------------------------------------
    'Resolve the gamma according to the risk class
        Select Case rc
            Case RC_GIRR
                BucketGamma = mGirrGamma
            Case RC_FX
                BucketGamma = mFxGamma
            Case RC_CSR
                IndexB = CLng(Val(b))
                IndexC = CLng(Val(c))
                ' Rating factor applies only between two buckets of 1-15 (MAR21.57)
                If mCsrType(IndexB) = "NAME" And mCsrType(IndexC) = "NAME" And _
                    mCsrQuality(IndexB) <> mCsrQuality(IndexC) Then
                    RatingFactor = mCsrGammaRating
                Else
                    RatingFactor = 1#
                End If
                BucketGamma = RatingFactor * mCsrSector(KeyIndex(mCsrGroupIdx, mCsrGroup(IndexB)), _
                    KeyIndex(mCsrGroupIdx, mCsrGroup(IndexC)))
            Case RC_EQ
                BucketGamma = mEqGamma(CLng(Val(b)), CLng(Val(c)))
            Case RC_COM
                BucketGamma = mComGamma(CLng(Val(b)), CLng(Val(c)))
            Case Else
                BucketGamma = 0#
        End Select

End Function

'
'------------------------------------------------------------------------------
'
'                            DRC AND RRAO ACCESSORS
'
'------------------------------------------------------------------------------
'

Public Function DrcRiskWeight(ByVal quality As String, ByRef found As Boolean) As Double

'
'------------------------------------------------------------------------------
'                         RESOLVE THE DRC RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DRC risk weight of a credit quality.
'
' WHY THIS EXISTS
'   The risk weight is stored by credit quality and the caller needs to know whether the
'   quality is listed.
'
' INPUTS
'   quality
'     Credit quality
'
'   found
'     Set to True when the quality is listed, False otherwise
'
' RETURNS
'   The risk weight of the credit quality, 0 when it is not listed.
'
' BEHAVIOR
'   Sets found, then returns the risk weight stored under the upper-case quality, or 0
'   when it is not listed.
'
' ERROR POLICY
'   None. A collection that is not loaded counts as an unlisted key.
'
' DEPENDENCIES
'   KeyExists (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Report whether the quality is listed
        found = KeyExists(mDrcRW, quality)
    'Return the value when it is listed, 0 otherwise
        If found Then
            DrcRiskWeight = CDbl(mDrcRW.Item(UCase$(quality)))
        Else
            DrcRiskWeight = 0#
        End If

End Function

Public Function DrcLGD(ByVal seniority As String, ByRef found As Boolean) As Double

'
'------------------------------------------------------------------------------
'                      RESOLVE THE DRC LOSS GIVEN DEFAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DRC loss given default of a seniority.
'
' WHY THIS EXISTS
'   The loss given default is stored by seniority and the caller needs to know whether the
'   seniority is listed.
'
' INPUTS
'   seniority
'     Seniority
'
'   found
'     Set to True when the seniority is listed, False otherwise
'
' RETURNS
'   The loss given default of the seniority, 0 when it is not listed.
'
' BEHAVIOR
'   Sets found, then returns the loss given default stored under the upper-case seniority,
'   or 0 when it is not listed.
'
' ERROR POLICY
'   None. A collection that is not loaded counts as an unlisted key.
'
' DEPENDENCIES
'   KeyExists (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Report whether the seniority is listed
        found = KeyExists(mDrcLGD, seniority)
    'Return the value when it is listed, 0 otherwise
        If found Then
            DrcLGD = CDbl(mDrcLGD.Item(UCase$(seniority)))
        Else
            DrcLGD = 0#
        End If

End Function

Public Function DrcSeniorityRank(ByVal seniority As String) As Long

'
'------------------------------------------------------------------------------
'                          RESOLVE THE SENIORITY RANK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DRC rank of a seniority.
'
' WHY THIS EXISTS
'   The ranks are loaded from DRC_SENIORITY and stay private to the module.
'
' INPUTS
'   seniority
'     Seniority
'
' RETURNS
'   The rank of the seniority, 0 when it is not listed.
'
' BEHAVIOR
'   Looks the seniority up in the rank collection.
'
' ERROR POLICY
'   None. A collection that is not loaded counts as an unlisted key.
'
' DEPENDENCIES
'   KeyIndex (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Look the rank up by seniority
        DrcSeniorityRank = KeyIndex(mDrcRank, seniority)

End Function

Public Function DrcMaxRank() As Long

'
'------------------------------------------------------------------------------
'                      RESOLVE THE HIGHEST SENIORITY RANK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the highest DRC seniority rank.
'
' WHY THIS EXISTS
'   Callers size their rank arrays by the highest rank.
'
' INPUTS
'   None
'
' RETURNS
'   The highest rank, 0 when no seniority is loaded.
'
' BEHAVIOR
'   Scans the rank collection and keeps the largest value.
'
' ERROR POLICY
'   None of its own; before LoadAllParameters the collection is not set and the call
'   fails.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Params.DrcMaxRank"

    Dim RankIndex               As Long 'Collection item being read
    Dim MaxRank                 As Long 'Highest rank found

'------------------------------------------------------------------------------
' SCAN RANKS
'------------------------------------------------------------------------------
    'Start at 0 and keep the largest rank in the collection
        MaxRank = 0
        For RankIndex = 1 To mDrcRank.Count
            If CLng(mDrcRank.Item(RankIndex)) > MaxRank Then MaxRank = CLng(mDrcRank.Item(RankIndex))
        Next RankIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the highest rank
        DrcMaxRank = MaxRank

End Function

Public Function IsValidDrcBucket(ByVal bucket As String) As Boolean

'
'------------------------------------------------------------------------------
'                              TEST A DRC BUCKET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a bucket name is a loaded DRC bucket.
'
' WHY THIS EXISTS
'   The buckets are loaded from DRC_BUCKETS and stay private to the module.
'
' INPUTS
'   bucket
'     Bucket name
'
' RETURNS
'   True when the bucket is listed, False otherwise.
'
' BEHAVIOR
'   Looks the bucket up in the bucket collection.
'
' ERROR POLICY
'   None. A collection that is not loaded counts as an unlisted key.
'
' DEPENDENCIES
'   KeyExists (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Look the bucket up by name
        IsValidDrcBucket = KeyExists(mDrcBuckets, bucket)

End Function

Public Function DrcBucketCount() As Long

'
'------------------------------------------------------------------------------
'                            COUNT THE DRC BUCKETS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of loaded DRC buckets.
'
' WHY THIS EXISTS
'   Callers size their bucket arrays by the bucket count.
'
' INPUTS
'   None
'
' RETURNS
'   The number of buckets of DRC_BUCKETS.
'
' BEHAVIOR
'   Returns the item count of the bucket collection.
'
' ERROR POLICY
'   None of its own; before LoadAllParameters the collection is not set and the call
'   fails.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of buckets
        DrcBucketCount = mDrcBuckets.Count

End Function

Public Function DrcBucketIndex(ByVal bucket As String) As Long

'
'------------------------------------------------------------------------------
'                         RESOLVE THE DRC BUCKET INDEX
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the position of a DRC bucket in the bucket table.
'
' WHY THIS EXISTS
'   The position is the index of the bucket in the bucket arrays of the DRC engine.
'
' INPUTS
'   bucket
'     Bucket name
'
' RETURNS
'   The position from 1, 0 when the bucket is not listed.
'
' BEHAVIOR
'   Looks the bucket up in the bucket collection.
'
' ERROR POLICY
'   None. A collection that is not loaded counts as an unlisted key.
'
' DEPENDENCIES
'   KeyIndex (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Look the position up by bucket name
        DrcBucketIndex = KeyIndex(mDrcBuckets, bucket)

End Function

Public Function DrcBucketName(ByVal i As Long) As String

'
'------------------------------------------------------------------------------
'                         RESOLVE THE DRC BUCKET NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the name of a DRC bucket from its position.
'
' WHY THIS EXISTS
'   It is the inverse of DrcBucketIndex.
'
' INPUTS
'   i
'     Position of the bucket in the bucket table, from 1
'
' RETURNS
'   The bucket name.
'
' BEHAVIOR
'   Returns the name stored at the position.
'
' ERROR POLICY
'   None of its own; a position outside the loaded table raises a VBA subscript error.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the name at the position
        DrcBucketName = mDrcBucketNames(i)

End Function

Public Function DrcMinMaturity() As Double

'
'------------------------------------------------------------------------------
'                       RETURN THE DRC MINIMUM MATURITY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the minimum maturity of the DRC maturity scaling.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The MIN_MATURITY value of DRC_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadDrc.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        DrcMinMaturity = mDrcMinMat

End Function

Public Function DrcMaxMaturity() As Double

'
'------------------------------------------------------------------------------
'                       RETURN THE DRC MAXIMUM MATURITY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the maximum maturity of the DRC maturity scaling.
'
' WHY THIS EXISTS
'   The value stays private to the module; callers read it here.
'
' INPUTS
'   None
'
' RETURNS
'   The MAX_MATURITY value of DRC_PARAMS.
'
' BEHAVIOR
'   Returns the value stored by LoadDrc.
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
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the stored value
        DrcMaxMaturity = mDrcMaxMat

End Function

Public Function RraoRiskWeight(ByVal category As String, ByRef found As Boolean) As Double

'
'------------------------------------------------------------------------------
'                         RESOLVE THE RRAO RISK WEIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the RRAO risk weight of a category.
'
' WHY THIS EXISTS
'   The risk weight is stored by category and the caller needs to know whether the
'   category is listed.
'
' INPUTS
'   category
'     RRAO category
'
'   found
'     Set to True when the category is listed, False otherwise
'
' RETURNS
'   The risk weight of the category, 0 when it is not listed.
'
' BEHAVIOR
'   Sets found, then returns the risk weight stored under the upper-case category, or 0
'   when it is not listed.
'
' ERROR POLICY
'   None. A collection that is not loaded counts as an unlisted key.
'
' DEPENDENCIES
'   KeyExists (CORE_Util).
'
' NOTES
'   None
'
' UPDATED
'   2026-10-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Report whether the category is listed
        found = KeyExists(mRraoRW, category)
    'Return the risk weight when it is listed, 0 otherwise
        If found Then
            RraoRiskWeight = CDbl(mRraoRW.Item(UCase$(category)))
        Else
            RraoRiskWeight = 0#
        End If

End Function
