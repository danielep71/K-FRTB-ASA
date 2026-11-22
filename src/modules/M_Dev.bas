Attribute VB_Name = "M_Dev"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_Dev
'------------------------------------------------------------------------------
' PURPOSE
'   Exports the standard modules of this workbook back to the repository, so that
'   changes made in the VBA editor can be reviewed and committed as text: TEST_*
'   to tests/modules, every other standard module to src/modules.
'
' WHY THIS EXISTS
'   Changes made in the VBA editor live in the workbook only; the repository holds
'   the text of the modules.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Writes one .bas file for each standard module; changes nothing in the workbook.
'
' ERROR POLICY
'   Shows a message and exits when the repository folders are not found next to or
'   above the workbook.
'
' DEPENDENCIES
'   CORE_Config for APP_NAME. Needs "Trust access to the VBA project object model".
'
' NOTES
'   Excel for Windows only (VBIDE object model, late bound).
'
' PUBLIC SURFACE
'   Dev_ExportModules.
'
' STATE OWNERSHIP
'   Writes .bas files; changes nothing in the workbook.
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const VBEXT_STD_MODULE  As Long = 1 'vbext_ct_StdModule


'
'------------------------------------------------------------------------------
'
'                                PUBLIC MACROS
'
'------------------------------------------------------------------------------
'

Public Sub Dev_ExportModules()

'
'------------------------------------------------------------------------------
'                              EXPORT THE MODULES
'------------------------------------------------------------------------------
' PURPOSE
'   Exports every standard module of this workbook to a .bas file in the repository
'   folders.
'
' WHY THIS EXISTS
'   A change made in the VBA editor reaches the repository only through a file;
'   this command writes the files in the places the repository expects.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. A message reports the number of modules exported.
'
' BEHAVIOR
'   Looks for src\modules next to the workbook, then one level above it (the build
'   script saves the workbook in dist/). Then, for each standard module:
'     - TEST_* modules are exported to tests\modules
'     - every other standard module is exported to src\modules
'
' ERROR POLICY
'   Shows a message and exits when src\modules is not found. Other errors, for
'   example the VBA project access being off, are not handled here.
'
' DEPENDENCIES
'   ThisWorkbook.VBProject (VBIDE, late bound), Application.PathSeparator, APP_NAME
'   (CORE_Config).
'
' NOTES
'   Overwrites the .bas files of the repository folders it finds.
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Dev.Dev_ExportModules"

    Dim RootFolder              As String       'Repository root
    Dim PathSep                 As String       'Path separator of the host
    Dim Component               As Object       'VBComponent being exported
    Dim ExportedCount           As Long         'Number of modules exported
    Dim TargetFolder            As String       'Folder that receives the module

'------------------------------------------------------------------------------
' LOCATE THE REPOSITORY
'------------------------------------------------------------------------------
    'Read the path separator of the host
        PathSep = Application.PathSeparator
    'The build script saves the workbook in dist/, one level below the root
        RootFolder = ThisWorkbook.Path
    'Try the folder above when src\modules is not next to the workbook
        If Dir(RootFolder & PathSep & "src" & PathSep & "modules", vbDirectory) = vbNullString Then
            RootFolder = RootFolder & PathSep & ".."
        End If
    'Stop with a message when the repository is not found
        If Dir(RootFolder & PathSep & "src" & PathSep & "modules", vbDirectory) = vbNullString Then
            MsgBox "Folder src\modules not found next to or above " & ThisWorkbook.Path, _
                vbExclamation, APP_NAME
            Exit Sub
        End If

'------------------------------------------------------------------------------
' EXPORT MODULES
'------------------------------------------------------------------------------
    'Start the count at zero
        ExportedCount = 0
    'Export each standard module to the folder of its kind
        For Each Component In ThisWorkbook.VBProject.VBComponents
            If Component.Type = VBEXT_STD_MODULE Then
                If VBA.Left$(Component.Name, 5) = "TEST_" Then
                    TargetFolder = RootFolder & PathSep & "tests" & PathSep & "modules"
                Else
                    TargetFolder = RootFolder & PathSep & "src" & PathSep & "modules"
                End If
                Component.Export TargetFolder & PathSep & Component.Name & ".bas"
                ExportedCount = ExportedCount + 1
            End If
        Next Component

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Tell how many modules were exported
        MsgBox ExportedCount & " module(s) exported under " & RootFolder, vbInformation, APP_NAME

End Sub
