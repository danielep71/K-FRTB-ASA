Attribute VB_Name = "CORE_Config"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Config
'------------------------------------------------------------------------------
' PURPOSE
'   Holds the workbook-wide constants: product name, sheet names, sheet layout,
'   color palette, risk class codes, SBM components and correlation scenarios,
'   and the pure lookups that translate them.
'
' WHY THIS EXISTS
'   Every other module refers to these names only, so a sheet name, a color or a
'   risk class code changes in one place.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Declares constants and six lookup functions that depend on nothing else.
'
' ERROR POLICY
'   No procedure raises.
'
' DEPENDENCIES
'   None
'
' NOTES
'   The test tools read some constants straight from this file (the sheet names, the
'   tab colors, the risk class codes, LAYOUT_FIRST_DATA_ROW and INPUT_ROWS), so keep
'   each of them on one line, as 'Public Const NAME As Type = value' with any
'   spacing around As.
'
'   The module also runs in LibreOffice Calc with VBA support for the automated
'   tests. It needs no references beyond the defaults.
'
' PUBLIC SURFACE
'   Constants; ManagedSheets, RiskClassCode, RiskClassIndex, ComponentName,
'   ScenarioName, HasNumericBuckets.
'
' STATE OWNERSHIP
'   None: constants and pure functions only.
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'Product
    Public Const APP_NAME       As String = "K-FRTB-ASA"                    'Product name in the brand bar
    Public Const APP_TITLE      As String = "FRTB standardized approach"    'Product title
    Public Const APP_VERSION    As String = "0.2.0-prototype"               'Version the workbook reports

    'Sheet names
    Public Const SH_README      As String = "README"        'README sheet
    Public Const SH_HOME        As String = "Home"          'Home sheet, the start sheet
    Public Const SH_CONTROL     As String = "Control"       'Control sheet
    Public Const SH_RESULTS     As String = "Results"       'Results sheet
    Public Const SH_IN_DELTA    As String = "In_Delta"      'Delta input sheet
    Public Const SH_IN_VEGA     As String = "In_Vega"       'Vega input sheet
    Public Const SH_IN_CURV     As String = "In_Curvature"  'Curvature input sheet
    Public Const SH_IN_DRC      As String = "In_DRC"        'DRC input sheet
    Public Const SH_IN_RRAO     As String = "In_RRAO"       'RRAO input sheet
    Public Const SH_P_GENERAL   As String = "P_General"     'General parameters sheet
    Public Const SH_P_GIRR      As String = "P_GIRR"        'GIRR parameters sheet
    Public Const SH_P_CSR       As String = "P_CSR"         'CSR parameters sheet
    Public Const SH_P_EQ        As String = "P_EQ"          'Equity parameters sheet
    Public Const SH_P_COM       As String = "P_COM"         'Commodity parameters sheet
    Public Const SH_P_FX        As String = "P_FX"          'FX parameters sheet
    Public Const SH_P_DRC       As String = "P_DRC"         'DRC parameters sheet
    Public Const SH_P_RRAO      As String = "P_RRAO"        'RRAO parameters sheet
    Public Const SH_OUT_CLASS   As String = "Out_RiskClass" 'Risk class output sheet
    Public Const SH_OUT_BUCKETS As String = "Out_Buckets"   'Bucket output sheet
    Public Const SH_OUT_DRC     As String = "Out_DRC"       'DRC output sheet
    Public Const SH_OUT_SUMMARY As String = "Out_Summary"   'Summary output sheet
    Public Const SH_LOG         As String = "Log"           'Log sheet
    Public Const SH_TEST        As String = "Test_Expected" 'Expected values and tests sheet

    'Sheet layout shared by every sheet: brand bar (row 1), title bar (row 2), subtitle (row 3),
    'spacer (row 4); the product name and the title sit in column C with the subtitle and the
    'tables, right of the bar (column A) and the gutter (column B)
    Public Const LAYOUT_FIRST_COL       As Long = 3     'Column C
    Public Const LAYOUT_NAV_WIDTH       As Long = 18    'Column A: the navigation bar sits here (#347)
    Public Const LAYOUT_GUTTER_WIDTH    As Long = 2     'Column B: the gutter between the bar and the table
    Public Const LAYOUT_SUBTITLE_ROW    As Long = 3     'Subtitle row
    Public Const LAYOUT_HEADER_ROW      As Long = 5     'Table header on input and output sheets
    Public Const INPUT_ROWS             As Long = 200   'Pre-formatted input rows per input sheet
    Public Const LAYOUT_FIRST_DATA_ROW  As Long = 6     'First data row below the header
    Public Const LAYOUT_LEAD_ROW_HEIGHT As Double = 46  'Height of the first row of text on Home and README

    'The two triangles in B1 that hide and show the navigation bar (#353)
    Public Const NAV_TOGGLE_HIDE As String = "NavToggleHide"  'Shown while the bar is shown
    Public Const NAV_TOGGLE_SHOW As String = "NavToggleShow"  'Shown while the bar is hidden

    'Palette (Excel BGR color values; hex RGB in the trailing comments)
    Public Const COLOR_BRAND_BAR        As Long = 0         '000000 brand bar
    Public Const COLOR_BRAND_TEXT       As Long = 49407     'FFC000 product name
    Public Const COLOR_TITLE_BAR        As Long = 10050086  '265A99 sheet title bar
    Public Const COLOR_SUBTITLE         As Long = 5855577   '595959 subtitle text
    Public Const COLOR_INPUT_HEADER     As Long = 6567967   '1F3864 input table header, section headings
    Public Const COLOR_OUTPUT_HEADER    As Long = 6567967   '1F3864 output table header, the input navy (#206)
    Public Const COLOR_INPUT_CELL       As Long = 13431551  'FFF2CC cells meant to be edited
    Public Const COLOR_GRID             As Long = 12566463  'BFBFBF thin cell borders
    Public Const COLOR_NOTE             As Long = 24703     '7F6000 status notes
    Public Const COLOR_TOTAL_FILL       As Long = 16247773  'DDEBF7 total rows
    Public Const COLOR_PASS_FILL        As Long = 14348258  'E2EFDA pass fill
    Public Const COLOR_PASS_TEXT        As Long = 2315831   '375623 pass text
    Public Const COLOR_FAIL_FILL        As Long = 13551615  'FFC7CE fail fill
    Public Const COLOR_FAIL_TEXT        As Long = 393372    '9C0006 fail text
    Public Const COLOR_WARN_FILL        As Long = 10284031  'FFEB9C warning fill
    Public Const COLOR_WARN_TEXT        As Long = 22428     '9C5700 warning text
    Public Const COLOR_WHITE            As Long = 16777215  'FFFFFF white
    Public Const COLOR_NAV_ENTRY        As Long = 2500134   '262626 navigation entry on the black column A (#347)
    Public Const COLOR_NAV_TEXT         As Long = 14277081  'D9D9D9 navigation entry text
    Public Const TAB_REFERENCE          As Long = 6567967   '1F3864 README
    Public Const TAB_INPUT              As Long = 37055     'BF9000 inputs and parameters
    Public Const TAB_OUTPUT             As Long = 11957550  '2E75B6 outputs
    Public Const TAB_TEST               As Long = 3506772   '548235 tests
    Public Const FONT_NAME              As String = "Aptos" 'Workbook font
    Public Const HEADER_FONT_SIZE       As Long = 11        'Table headers (#351); data rows stay at 10

    'Risk classes implemented by the SBM engine (in reporting order)
    Public Const RC_GIRR    As String = "GIRR"      'General interest rate risk
    Public Const RC_CSR     As String = "CSR_NS"    'Credit spread risk, non-securitization
    Public Const RC_EQ      As String = "EQ"        'Equity risk
    Public Const RC_COM     As String = "COM"       'Commodity risk
    Public Const RC_FX      As String = "FX"        'Foreign exchange risk
    Public Const RC_COUNT   As Long = 5             'Number of risk classes

    'SBM components
    Public Const COMP_DELTA As Long = 1 'Delta
    Public Const COMP_VEGA  As Long = 2 'Vega
    Public Const COMP_CURV  As Long = 3 'Curvature
    Public Const COMP_COUNT As Long = 3 'Number of components

    'Correlation scenarios (MAR21.6)
    Public Const SCEN_LOW       As Long = 1 'Low correlations
    Public Const SCEN_MEDIUM    As Long = 2 'Medium correlations
    Public Const SCEN_HIGH      As Long = 3 'High correlations
    Public Const SCEN_COUNT     As Long = 3 'Number of scenarios


'
'------------------------------------------------------------------------------
'
'                                 CODE LOOKUPS
'
'------------------------------------------------------------------------------
'

Public Function ManagedSheets() As Variant

'
'------------------------------------------------------------------------------
'                           LIST THE MANAGED SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the names of the sheets the workbook manages, in workbook order.
'
' WHY THIS EXISTS
'   The build, the commands and the checks walk the same list of sheets.
'
' INPUTS
'   None
'
' RETURNS
'   An array of 23 sheet names, starting with Home and ending with README.
'
' BEHAVIOR
'   Builds the array from the SH_* constants.
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
    'Build the sheet list in workbook order
        ManagedSheets = Array(SH_HOME, SH_CONTROL, SH_RESULTS, _
            SH_IN_DELTA, SH_IN_VEGA, SH_IN_CURV, SH_IN_DRC, SH_IN_RRAO, _
            SH_OUT_SUMMARY, SH_OUT_CLASS, SH_OUT_BUCKETS, SH_OUT_DRC, SH_LOG, _
            SH_P_GENERAL, SH_P_GIRR, SH_P_CSR, SH_P_EQ, SH_P_COM, SH_P_FX, _
            SH_P_DRC, SH_P_RRAO, SH_TEST, SH_README)

End Function

Public Function RiskClassCode(ByVal i As Long) As String

'
'------------------------------------------------------------------------------
'                           RESOLVE RISK CLASS CODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the code of the risk class at a position of the reporting order.
'
' WHY THIS EXISTS
'   Loops over the risk classes need the code that belongs to each position.
'
' INPUTS
'   i
'     Position of the risk class, 1 to RC_COUNT
'
' RETURNS
'   The RC_* code, or an empty string when the position is out of range.
'
' BEHAVIOR
'   Maps 1 to 5 to GIRR, CSR_NS, EQ, COM and FX.
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
    'Translate the position into its code
        Select Case i
            Case 1
                RiskClassCode = RC_GIRR
            Case 2
                RiskClassCode = RC_CSR
            Case 3
                RiskClassCode = RC_EQ
            Case 4
                RiskClassCode = RC_COM
            Case 5
                RiskClassCode = RC_FX
            Case Else
                RiskClassCode = vbNullString
        End Select

End Function

Public Function RiskClassIndex(ByVal rc As String) As Long

'
'------------------------------------------------------------------------------
'                           RESOLVE RISK CLASS INDEX
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the position of a risk class code in the reporting order.
'
' WHY THIS EXISTS
'   It is the inverse of RiskClassCode, for code that holds a code and needs its position.
'
' INPUTS
'   rc
'     Risk class code, one of the RC_* constants
'
' RETURNS
'   The position from 1 to RC_COUNT, or 0 when the code is not a risk class.
'
' BEHAVIOR
'   Starts at 0, then tests the codes in order and stops at the first match.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   RiskClassCode
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
    Const PROC_NAME             As String = "CORE_Config.RiskClassIndex"

    Dim ClassIndex              As Long         'Position being tested

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Return 0 unless a code matches
        RiskClassIndex = 0
    'Test each risk class code in reporting order
        For ClassIndex = 1 To RC_COUNT
            If RiskClassCode(ClassIndex) = rc Then
                RiskClassIndex = ClassIndex
                Exit Function
            End If
        Next ClassIndex

End Function

Public Function ComponentName(ByVal comp As Long) As String

'
'------------------------------------------------------------------------------
'                            RESOLVE COMPONENT NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the name of an SBM component.
'
' WHY THIS EXISTS
'   Logs, results and the Test_Expected sheet show the component by name.
'
' INPUTS
'   comp
'     Component, one of the COMP_* constants
'
' RETURNS
'   DELTA, VEGA or CURVATURE, or an empty string for any other value.
'
' BEHAVIOR
'   Maps each COMP_* constant to its upper-case name.
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
    'Translate the component into its name
        Select Case comp
            Case COMP_DELTA
                ComponentName = "DELTA"
            Case COMP_VEGA
                ComponentName = "VEGA"
            Case COMP_CURV
                ComponentName = "CURVATURE"
            Case Else
                ComponentName = vbNullString
        End Select

End Function

Public Function ScenarioName(ByVal scen As Long) As String

'
'------------------------------------------------------------------------------
'                            RESOLVE SCENARIO NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the name of a correlation scenario.
'
' WHY THIS EXISTS
'   Results and the Test_Expected sheet show the scenario by name.
'
' INPUTS
'   scen
'     Scenario, one of the SCEN_* constants
'
' RETURNS
'   LOW, MEDIUM or HIGH, or an empty string for any other value.
'
' BEHAVIOR
'   Maps each SCEN_* constant to its upper-case name (MAR21.6).
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
    'Translate the scenario into its name
        Select Case scen
            Case SCEN_LOW
                ScenarioName = "LOW"
            Case SCEN_MEDIUM
                ScenarioName = "MEDIUM"
            Case SCEN_HIGH
                ScenarioName = "HIGH"
            Case Else
                ScenarioName = vbNullString
        End Select

End Function

Public Function HasNumericBuckets(ByVal rc As String) As Boolean

'
'------------------------------------------------------------------------------
'                          TEST FOR NUMBERED BUCKETS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True for the risk classes whose buckets are numbered: CSR, equity and
'   commodity.
'
' WHY THIS EXISTS
'   GIRR and FX buckets are currency codes, so code that reads a bucket needs to know
'   which kind a risk class uses.
'
' INPUTS
'   rc
'     Risk class code, one of the RC_* constants
'
' RETURNS
'   True for RC_CSR, RC_EQ and RC_COM, False for any other value.
'
' BEHAVIOR
'   Compares the code with the three numbered risk classes.
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
    'Flag the risk classes with numbered buckets
        HasNumericBuckets = (rc = RC_CSR Or rc = RC_EQ Or rc = RC_COM)

End Function
