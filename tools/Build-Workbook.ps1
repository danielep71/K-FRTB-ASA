<#
.SYNOPSIS
    Builds the FRTB-SA prototype workbook (.xlsm) from the template and the VBA sources.

.DESCRIPTION
    Starts Excel through COM and opens a copy of the macro-free template
    src\workbook\FRTB_Template.xlsx, imports the modules of src\modules and
    (unless -Production) tests\modules, runs M_Setup.FinishWorkbook and
    M_Setup.AddNavigationBars to add the sheet buttons, reapply saved view
    settings and add the navigation bars, and saves the result as a
    macro-enabled workbook. With -RunTests it also runs the calculation and
    the in-workbook regression test and fails when the test does not pass.

    -FromSource builds every sheet from the modules instead
    (M_Setup.BuildWorkbook in a new workbook); -Template does the same and
    saves the result as a macro-free template for review.

    Requirements: Windows with desktop Excel, and "Trust access to the VBA
    project object model" enabled (File > Options > Trust Center > Trust
    Center Settings > Macro Settings).

.PARAMETER OutputPath
    Target path. Defaults to dist\K-FRTB-ASA.xlsm, or dist\FRTB_Template.xlsx
    with -Template, under the repository root.

.PARAMETER RunTests
    Run RunFRTB_Silent and Test_RunRegression_Silent after the build and check
    the result on the Test_Expected sheet.

.PARAMETER FromSource
    Build every sheet from the modules (M_Setup.BuildWorkbook) instead of
    starting from the template.

.PARAMETER Template
    Build from source and save a macro-free template, without sheet buttons
    and with personal information removed. Review it before replacing
    src\workbook\FRTB_Template.xlsx. Cannot be combined with -RunTests.

.PARAMETER Production
    Leave out the test modules of tests\modules (the "Run regression test"
    button then has no macro). Cannot be combined with -RunTests.

.PARAMETER Ribbon
    After the save, add the FRTB-SA ribbon tab (src\ribbon\customUI14.xml) to the
    workbook with tools\Add-RibbonTab.ps1. Cannot be combined with -Template.

.PARAMETER Force
    Overwrite an existing output file.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Build-Workbook.ps1 -RunTests
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [switch]$RunTests,
    [switch]$FromSource,
    [switch]$Production,
    [switch]$Template,
    [switch]$Ribbon,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$templatePath = Join-Path $repoRoot 'src\workbook\FRTB_Template.xlsx'
$sourceDirs = @(Join-Path $repoRoot 'src\modules')
if (-not $Production) {
    $sourceDirs += (Join-Path $repoRoot 'tests\modules')
}
if ($Production -and $RunTests) {
    throw '-RunTests needs the test modules; do not combine it with -Production'
}
if ($Template -and $RunTests) {
    throw '-Template saves the workbook before any run; do not combine it with -RunTests'
}
if ($Template -and $Ribbon) {
    throw '-Template saves a macro-free template; the ribbon tab needs a macro-enabled workbook'
}
$buildFromSource = $FromSource -or $Template
if (-not $buildFromSource -and -not (Test-Path $templatePath)) {
    throw "Template not found: $templatePath (use -FromSource to build without it)"
}
if (-not $OutputPath) {
    if ($Template) {
        $OutputPath = Join-Path $repoRoot 'dist\FRTB_Template.xlsx'
    } else {
        $OutputPath = Join-Path $repoRoot 'dist\K-FRTB-ASA.xlsm'
    }
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

if ((Test-Path $OutputPath) -and -not $Force) {
    throw "Output exists: $OutputPath (use -Force to overwrite)"
}
$modules = @(foreach ($dir in $sourceDirs) { Get-ChildItem -Path $dir -Filter '*.bas' | Sort-Object Name })
if ($modules.Count -eq 0) {
    throw "No .bas modules found in $($sourceDirs -join ', ')"
}
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputPath) | Out-Null

# The VBA editor expects CRLF line endings; normalize into a temporary folder.
# The template is copied there too, so the committed file is never modified.
$staging = Join-Path ([System.IO.Path]::GetTempPath()) ("frtb-sa-build-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $staging | Out-Null
$encoding = [System.Text.Encoding]::GetEncoding(1252)
foreach ($m in $modules) {
    $text = [System.IO.File]::ReadAllText($m.FullName, $encoding)
    $text = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $staging $m.Name), $text, $encoding)
}

$excel = $null
$workbook = $null
$exitCode = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    if ($buildFromSource) {
        $workbook = $excel.Workbooks.Add()
    } else {
        $templateCopy = Join-Path $staging 'FRTB_Template.xlsx'
        Copy-Item -Path $templatePath -Destination $templateCopy
        $workbook = $excel.Workbooks.Open($templateCopy)
        Write-Host "Opened a copy of $templatePath"
    }

    try {
        $components = $workbook.VBProject.VBComponents
    } catch {
        throw "Excel denied access to the VBA project. Enable 'Trust access to the VBA project object model' and retry."
    }
    foreach ($m in $modules) {
        [void]$components.Import((Join-Path $staging $m.Name))
        Write-Host "Imported $($m.Name)"
    }

    $prefix = "'" + $workbook.Name + "'!"
    if ($buildFromSource) {
        [void]$excel.Run($prefix + 'BuildWorkbook')
        Write-Host 'Workbook layout, parameters and sample data created from source'
    } else {
        [void]$excel.Run($prefix + 'FinishWorkbook')
        Write-Host 'Buttons added and saved view settings reapplied'
    }
    # The navigation bars are shapes, which only Excel adds; the template save below drops them
    [void]$excel.Run($prefix + 'AddNavigationBars')
    Write-Host 'Navigation bars added'

    if ($RunTests) {
        [void]$excel.Run($prefix + 'RunFRTB_Silent')
        [void]$excel.Run($prefix + 'Test_RunRegression_Silent')
        $result = [string]$workbook.Worksheets.Item('Test_Expected').Range('C3').Value2
        Write-Host "Regression test: $result"
        if (-not $result.StartsWith('Result: PASS')) {
            $exitCode = 1
        }
    }

    # The workbook's macros turn Excel alerts back on when they finish. A
    # hidden Excel would then wait on a dialog nobody sees: saving the
    # template as .xlsx asks whether to drop the VB project (#302)
    $excel.DisplayAlerts = $false

    if ($Template) {
        # A template holds no buttons: FinishWorkbook adds them at build time.
        # The sheets are protected without a password (#349): the write window
        # lifts it for the deletion and the macro that closes it restores it.
        [void]$excel.Run($prefix + 'OpenWriteWindow')
        foreach ($sheet in $workbook.Worksheets) {
            foreach ($shape in @($sheet.Shapes)) { $shape.Delete() }
        }
        [void]$excel.Run($prefix + 'CloseWriteWindow')
        # 51 = xlOpenXMLWorkbook: the saved file carries no VBA project
        $workbook.RemovePersonalInformation = $true
        $workbook.SaveAs($OutputPath, 51)
    } else {
        # 52 = xlOpenXMLWorkbookMacroEnabled
        $workbook.SaveAs($OutputPath, 52)
    }
    Write-Host "Saved $OutputPath"
} finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Remove-Item -Recurse -Force $staging -ErrorAction SilentlyContinue
}
# Excel saves no custom UI part: add it to the closed file (#353)
if ($Ribbon -and $exitCode -eq 0) {
    & (Join-Path $PSScriptRoot 'Add-RibbonTab.ps1') -Workbook $OutputPath
}
exit $exitCode
