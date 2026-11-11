Attribute VB_Name = "CORE_Config"
'==============================================================================
' MODULE: CORE_Config
'------------------------------------------------------------------------------
' PURPOSE
'   Hold the workbook-wide constants: product name, sheet names, sheet
'   layout, color palette, risk class codes, SBM components and correlation
'   scenarios. Every other module refers to these names only.
'
' PUBLIC SURFACE
'   Constants; RiskClassCode, RiskClassIndex, ComponentName, ScenarioName,
'   HasNumericBuckets.
'
' DEPENDENCIES
'   None.
'
' STATE OWNERSHIP
'   None: constants and pure functions only.
'
' ERROR POLICY
'   No procedure raises.
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
Public Const APP_NAME As String = "K-FRTB-ASA"                'Product name in the brand bar
Public Const APP_TITLE As String = "FRTB standardized approach"
Public Const APP_VERSION As String = "0.2.0-prototype"

' Sheet names
Public Const SH_README As String = "README"
Public Const SH_CONTROL As String = "Control"
Public Const SH_RESULTS As String = "Results"
Public Const SH_IN_DELTA As String = "In_Delta"
Public Const SH_IN_VEGA As String = "In_Vega"
Public Const SH_IN_CURV As String = "In_Curvature"
Public Const SH_IN_DRC As String = "In_DRC"
Public Const SH_IN_RRAO As String = "In_RRAO"
Public Const SH_P_GENERAL As String = "P_General"
Public Const SH_P_GIRR As String = "P_GIRR"
Public Const SH_P_CSR As String = "P_CSR"
Public Const SH_P_EQ As String = "P_EQ"
Public Const SH_P_COM As String = "P_COM"
Public Const SH_P_FX As String = "P_FX"
Public Const SH_P_DRC As String = "P_DRC"
Public Const SH_P_RRAO As String = "P_RRAO"
Public Const SH_OUT_CLASS As String = "Out_RiskClass"
Public Const SH_OUT_BUCKETS As String = "Out_Buckets"
Public Const SH_OUT_DRC As String = "Out_DRC"
Public Const SH_OUT_SUMMARY As String = "Out_Summary"
Public Const SH_LOG As String = "Log"
Public Const SH_TEST As String = "Test_Expected"

' Sheet layout shared by every sheet: brand bar (row 1), title bar (row 2),
' subtitle (row 3), spacer (row 4); tables start in column C.
Public Const LAYOUT_FIRST_COL As Long = 3                    'Column C
Public Const LAYOUT_SUBTITLE_ROW As Long = 3
Public Const LAYOUT_HEADER_ROW As Long = 5                   'Table header on input and output sheets
Public Const LAYOUT_FIRST_DATA_ROW As Long = 6               'First data row below the header

' Palette (Excel BGR color values; hex RGB in the comments)
Public Const COLOR_BRAND_BAR As Long = 0                     '000000 brand bar
Public Const COLOR_BRAND_TEXT As Long = 49407                'FFC000 product name
Public Const COLOR_TITLE_BAR As Long = 10050086              '265A99 sheet title bar
Public Const COLOR_SUBTITLE As Long = 5855577                '595959 subtitle text
Public Const COLOR_INPUT_HEADER As Long = 6567967            '1F3864 input table header, section headings
Public Const COLOR_OUTPUT_HEADER As Long = 11957550          '2E75B6 output table header
Public Const COLOR_INPUT_CELL As Long = 13431551             'FFF2CC cells meant to be edited
Public Const COLOR_GRID As Long = 12566463                   'BFBFBF thin cell borders
Public Const COLOR_NOTE As Long = 24703                      '7F6000 status notes
Public Const COLOR_TOTAL_FILL As Long = 16247773             'DDEBF7 total rows
Public Const COLOR_PASS_FILL As Long = 14348258              'E2EFDA
Public Const COLOR_PASS_TEXT As Long = 2315831               '375623
Public Const COLOR_FAIL_FILL As Long = 13551615              'FFC7CE
Public Const COLOR_FAIL_TEXT As Long = 393372                '9C0006
Public Const COLOR_WARN_FILL As Long = 10284031              'FFEB9C
Public Const COLOR_WARN_TEXT As Long = 22428                 '9C5700
Public Const COLOR_WHITE As Long = 16777215                  'FFFFFF
Public Const TAB_REFERENCE As Long = 6567967                 '1F3864 README
Public Const TAB_INPUT As Long = 37055                       'BF9000 inputs and parameters
Public Const TAB_OUTPUT As Long = 11957550                   '2E75B6 outputs
Public Const TAB_TEST As Long = 3506772                      '548235 tests
Public Const FONT_NAME As String = "Aptos"

' Risk classes implemented by the SBM engine (in reporting order)
Public Const RC_GIRR As String = "GIRR"
Public Const RC_CSR As String = "CSR_NS"
Public Const RC_EQ As String = "EQ"
Public Const RC_COM As String = "COM"
Public Const RC_FX As String = "FX"
Public Const RC_COUNT As Long = 5

' SBM components
Public Const COMP_DELTA As Long = 1
Public Const COMP_VEGA As Long = 2
Public Const COMP_CURV As Long = 3
Public Const COMP_COUNT As Long = 3

' Correlation scenarios (MAR21.6)
Public Const SCEN_LOW As Long = 1
Public Const SCEN_MEDIUM As Long = 2
Public Const SCEN_HIGH As Long = 3
Public Const SCEN_COUNT As Long = 3

'------------------------------------------------------------------------------
'                              CODE LOOKUPS
'------------------------------------------------------------------------------
Public Function RiskClassCode(ByVal i As Long) As String
    Select Case i
        Case 1: RiskClassCode = RC_GIRR
        Case 2: RiskClassCode = RC_CSR
        Case 3: RiskClassCode = RC_EQ
        Case 4: RiskClassCode = RC_COM
        Case 5: RiskClassCode = RC_FX
        Case Else: RiskClassCode = ""
    End Select
End Function


Public Function RiskClassIndex(ByVal rc As String) As Long
    Dim i As Long
    RiskClassIndex = 0
    For i = 1 To RC_COUNT
        If RiskClassCode(i) = rc Then
            RiskClassIndex = i
            Exit Function
        End If
    Next i
End Function


Public Function ComponentName(ByVal comp As Long) As String
    Select Case comp
        Case COMP_DELTA: ComponentName = "DELTA"
        Case COMP_VEGA: ComponentName = "VEGA"
        Case COMP_CURV: ComponentName = "CURVATURE"
        Case Else: ComponentName = ""
    End Select
End Function


Public Function ScenarioName(ByVal scen As Long) As String
    Select Case scen
        Case SCEN_LOW: ScenarioName = "LOW"
        Case SCEN_MEDIUM: ScenarioName = "MEDIUM"
        Case SCEN_HIGH: ScenarioName = "HIGH"
        Case Else: ScenarioName = ""
    End Select
End Function


' True for risk classes whose buckets are numbered (CSR, equity, commodity);
' GIRR and FX buckets are currency codes.
Public Function HasNumericBuckets(ByVal rc As String) As Boolean
    HasNumericBuckets = (rc = RC_CSR Or rc = RC_EQ Or rc = RC_COM)
End Function
