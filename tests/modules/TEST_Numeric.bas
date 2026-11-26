Attribute VB_Name = "TEST_Numeric"

Option Explicit

'
'==============================================================================
' MODULE: TEST_Numeric
'------------------------------------------------------------------------------
' PURPOSE
'   Asserts the numeric validation and conversion contracts of CORE_Util with
'   synthetic values:
'   invariant numeric text, typed numbers, rejected text and required
'   parameters that fail with
'   their key and source.
'
' WHY THIS EXISTS
'   The numeric contracts decide which cell and parameter values count as
'   numbers, so they
'   are pinned by a test of their own.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Test_NumericContracts_Silent returns 1 when every contract holds and raises
'   on the first
'   one that fails.
'
' ERROR POLICY
'   A failed assertion raises vbObjectError + ERR_APP; no dialogs are shown.
'
' DEPENDENCIES
'   CORE_Util.
'
' NOTES
'   Development build only: tools/Build-Workbook.ps1 -Production leaves this
'   module out.
'
'   The module also runs in LibreOffice Calc with VBA support for the automated
'   tests. It needs no
'   references beyond the defaults.
'
' PUBLIC SURFACE
'   Test_NumericContracts_Silent.
'
' STATE OWNERSHIP
'   None: no sheet or workbook state is read or changed.
'
' UPDATED
'   2026-10-21
'==============================================================================


'
'------------------------------------------------------------------------------
'
'                                PUBLIC MACROS
'
'------------------------------------------------------------------------------
'

Public Function Test_NumericContracts_Silent() As Long
'
'------------------------------------------------------------------------------
'                         CHECK THE NUMERIC CONTRACTS
'------------------------------------------------------------------------------
' PURPOSE
'   Asserts the numeric validation and conversion contracts of CORE_Util with
'   synthetic values.
'
' WHY THIS EXISTS
'   The workbook reads numbers from cells and parameters as text or as typed
'   numbers; these contracts say which
'   texts are numbers and which are not.
'
' INPUTS
'   None
'
' RETURNS
'   1 when every contract holds.
'
' BEHAVIOR
'   Checks in order:
'     - valid numeric text (invariant decimal point, exponent, signs, blanks
'     around) is a number with the
'       expected amount
'     - a typed number is accepted
'     - invalid text (comma decimal, thousands separator, trailing junk,
'     overflow, symbols, NaN, hex, percent)
'       is not a number and cannot be converted
'     - a required parameter that is text, blank, Empty, Null, an error or
'     Boolean fails with its key and source
'     - a Boolean and a date are not numbers
'     - an optional value that is Empty or an error returns its default
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP at the first contract that fails; shows no
'   dialog.
'
' DEPENDENCIES
'   CORE_Util (IsNumber, ToDouble, KVNumber), NearlyEqual,
'   AssertInvalidRequired.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Numeric.Test_NumericContracts_Silent"

    Dim Inputs                  As Variant       'Valid numeric texts
    Dim ExpectedAmounts         As Variant       'Amounts the texts stand for
    Dim InvalidValues           As Variant       'Values that are not numbers
    Dim Item                    As Variant       'Value being checked
    Dim Index                   As Long          'Loop counter
    Dim ParsedValue             As Double        'Result of a conversion
    Dim CapturedError           As Long          'Error number of a conversion

'------------------------------------------------------------------------------
' PREPARE THE FIXTURES
'------------------------------------------------------------------------------
    'Prepare the valid texts and their amounts
        Inputs = Array("0.5", ".5", "1.", "-1.25e2", "+2E-1", " 0.5 ")
        ExpectedAmounts = Array(0.5, 0.5, 1#, -125#, 0.2, 0.5)

'------------------------------------------------------------------------------
' CHECK VALID TEXT
'------------------------------------------------------------------------------
    'Check that each valid text is a number with the expected amount
        For Index = LBound(Inputs) To UBound(Inputs)
            If Not IsNumber(Inputs(Index)) Then
                Err.Raise vbObjectError + ERR_APP, "TEST_Numeric", "Valid text rejected"
            End If
            If Not NearlyEqual(ToDouble(Inputs(Index), -99#), ExpectedAmounts(Index)) Then
                Err.Raise vbObjectError + ERR_APP, "TEST_Numeric", "Numeric text amount changed"
            End If
        Next Index

'------------------------------------------------------------------------------
' CHECK TYPED NUMBERS
'------------------------------------------------------------------------------
    'Check that a typed number is accepted
        If Not IsNumber(0.5) Or ToDouble(0.5, -99#) <> 0.5 Then
            Err.Raise vbObjectError + ERR_APP, "TEST_Numeric", "Typed numeric cell rejected"
        End If

'------------------------------------------------------------------------------
' CHECK INVALID TEXT
'------------------------------------------------------------------------------
    'Check that each invalid text is rejected and cannot be converted
        InvalidValues = Array("abc", "0,5", "1,5", "1,000", "1.2junk", "1e", "1e309", "1 2", "$1", "NaN", _
            "Infinity", "&H10", "1%")
        For Each Item In InvalidValues
            If IsNumber(Item) Then Err.Raise vbObjectError + ERR_APP, "TEST_Numeric", "Invalid text accepted"
            On Error Resume Next
            ParsedValue = ToDouble(Item, -99#)
            CapturedError = Err.Number
            Err.Clear
            On Error GoTo 0
            If CapturedError = 0 Then Err.Raise vbObjectError + ERR_APP, "TEST_Numeric", "Invalid text converted"
        Next Item

'------------------------------------------------------------------------------
' CHECK REQUIRED PARAMETERS
'------------------------------------------------------------------------------
    'Check that each bad required parameter fails with its key and source
        InvalidValues = Array("abc", "", Empty, Null, CVErr(2015), False)
        For Each Item In InvalidValues
            AssertInvalidRequired Item
        Next Item

'------------------------------------------------------------------------------
' CHECK BOOLEANS AND DATES
'------------------------------------------------------------------------------
    'Check that a Boolean and a date are not numbers
        If IsNumber(True) Or IsNumber(Now) Then
            Err.Raise vbObjectError + ERR_APP, "TEST_Numeric", "Boolean or date accepted as number"
        End If

'------------------------------------------------------------------------------
' CHECK OPTIONAL DEFAULTS
'------------------------------------------------------------------------------
    'Check that an empty or error value returns its default
        If ToDouble(Empty, 99#) <> 99# Or ToDouble(CVErr(2015), 99#) <> 99# Then
            Err.Raise vbObjectError + ERR_APP, "TEST_Numeric", "Optional default changed"
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Report that every contract holds
        Test_NumericContracts_Silent = 1

End Function


'
'------------------------------------------------------------------------------
'
'                                  ASSERTIONS
'
'------------------------------------------------------------------------------
'

Private Sub AssertInvalidRequired(ByVal item As Variant)
'
'------------------------------------------------------------------------------
'                    ASSERT A BAD REQUIRED PARAMETER FAILS
'------------------------------------------------------------------------------
' PURPOSE
'   Asserts that a required parameter with a bad value fails with its key and
'   source in the error.
'
' INPUTS
'   item
'     The bad value of the parameter
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Raises vbObjectError + ERR_APP when no error is raised or its text lacks
'   the key RW or the source
'   NUMERIC_FIXTURE.
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Numeric.AssertInvalidRequired"

    Dim KeyValues               As New Collection 'Fixture with one key
    Dim ParsedValue             As Double        'Result of the read
    Dim CapturedError           As Long          'Error number of the read
    Dim ErrorDetail             As String        'Error description of the read

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Read the parameter, keeping the error
        KeyValues.Add item, "RW"
        On Error Resume Next
        ParsedValue = KVNumber(KeyValues, "RW", "NUMERIC_FIXTURE")
        CapturedError = Err.Number
        ErrorDetail = Err.Description
        Err.Clear
        On Error GoTo 0
    'Require an error that names the key and the source
        If CapturedError = 0 Or InStr(ErrorDetail, "RW") = 0 Or InStr(ErrorDetail, "NUMERIC_FIXTURE") = 0 Then
            Err.Raise vbObjectError + ERR_APP, "TEST_Numeric", "Required parameter did not fail with context"
        End If

End Sub
