<#
.SYNOPSIS
    Exports the VBA components of an FRTB-SA workbook to a folder, for comparison
    with the repository source.

.DESCRIPTION
    Opens the workbook read-only in a separate Excel instance with macros
    disabled, exports every VBA component with the VBE's own Export (standard
    modules as .bas, class and document modules as .cls, forms as .frm) and
    closes the workbook without saving. The folder must be new or empty, so
    that an earlier export is never mixed in. Compare the export with
    tools/compare_vba_export.py.

    The script also writes export-manifest.json into the folder: the file name
    and SHA-256 of the workbook, hashed before it is opened and again after it
    is closed, and the VB_Name, file, component type and SHA-256 of every
    exported component, a form's .frx included. compare_vba_export.py --workbook checks the workbook
    and the files against it. The manifest is a record, not a signature.

    Unlike M_Dev.Dev_ExportModules, the export never writes into the
    repository and never changes the workbook.

    Requirements: Windows with desktop Excel, "Trust access to the VBA
    project object model" enabled (File > Options > Trust Center > Trust
    Center Settings > Macro Settings) and a VBA project that is not
    password-locked.

.PARAMETER WorkbookPath
    The .xlsm workbook to export. It is opened read-only.

.PARAMETER OutDir
    A new or empty folder for the exported files and the manifest.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\Export-FRTB-VBA.ps1 -WorkbookPath dist\K-FRTB-SA.xlsm -OutDir C:\FRTB\export
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoAutomationSecurityForceDisable = 3
$exportManifestName = 'export-manifest.json'
$exportFormat = 'k-frtb-sa-vba-export-1'

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-ExportExtension {
    # The file extension the VBE gives a component type, or $null for a type
    # it does not export.
    param([Parameter(Mandatory = $true)][int]$ComponentType)

    switch ($ComponentType) {
        1 { return '.bas' }
        2 { return '.cls' }
        3 { return '.frm' }
        100 { return '.cls' }
    }
    return $null
}

$excel = $null
$workbook = $null

try {
    $workbookFullPath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $WorkbookPath).Path)
    $outFullPath = [System.IO.Path]::GetFullPath($OutDir)
    if (Test-Path -LiteralPath $outFullPath -PathType Container) {
        if (@(Get-ChildItem -LiteralPath $outFullPath -Force).Count -gt 0) {
            throw "OutDir is not empty: $outFullPath. Use a new folder."
        }
    } else {
        $null = New-Item -ItemType Directory -Path $outFullPath
    }

    $workbookHash = Get-Sha256 -Path $workbookFullPath

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    # Read-only: the export never changes the workbook.
    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $true)
    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    } catch {
        throw ("Cannot access the VBA project. Enable 'Trust access to the VBA project object model' " +
            'and make sure the VBA project is not password-locked.')
    }

    $count = 0
    $components = New-Object System.Collections.Generic.List[object]
    foreach ($component in $vbProject.VBComponents) {
        $extension = Get-ExportExtension -ComponentType $component.Type
        if ($null -eq $extension) {
            Write-Warning "Not exported: $($component.Name) (component type $($component.Type))"
            continue
        }
        $target = Join-Path $outFullPath ($component.Name + $extension)
        $component.Export($target)
        # The type tells a document module (100), whose export carries no
        # VB_Base line, from a class module (2)
        $components.Add([ordered]@{
                name   = $component.Name
                file   = $component.Name + $extension
                type   = [int]$component.Type
                sha256 = Get-Sha256 -Path $target
            })
        # A form's binary properties go to a .frx beside its .frm.
        $resource = [System.IO.Path]::ChangeExtension($target, '.frx')
        if ($extension -eq '.frm' -and (Test-Path -LiteralPath $resource -PathType Leaf)) {
            $components.Add([ordered]@{
                    name   = $component.Name
                    file   = $component.Name + '.frx'
                    type   = [int]$component.Type
                    sha256 = Get-Sha256 -Path $resource
                })
        }
        Write-Host ('  {0}{1}' -f $component.Name, $extension)
        $count++
    }

    # Close before hashing again: the manifest records the workbook only if
    # the export left it unchanged.
    $workbook.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    $workbook = $null
    if ((Get-Sha256 -Path $workbookFullPath) -ne $workbookHash) {
        throw "The workbook changed while it was exported: $workbookFullPath. Export it again."
    }

    $manifest = [ordered]@{
        format     = $exportFormat
        workbook   = [ordered]@{
            file   = [System.IO.Path]::GetFileName($workbookFullPath)
            sha256 = $workbookHash
        }
        components = $components.ToArray()
    }
    [System.IO.File]::WriteAllText(
        (Join-Path $outFullPath $exportManifestName),
        ($manifest | ConvertTo-Json -Depth 4),
        (New-Object System.Text.UTF8Encoding($false)))

    Write-Host ''
    Write-Host "Workbook: $workbookFullPath"
    Write-Host "SHA-256:  $workbookHash"
    Write-Host "Exported: $count component(s) to $outFullPath, with $exportManifestName"
    Write-Host "Next:     python tools/compare_vba_export.py --export `"$outFullPath`" --mode production|development --candidate-sha <commit> --workbook `"$workbookFullPath`""
} catch {
    Write-Error $_ -ErrorAction Continue
    exit 1
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
