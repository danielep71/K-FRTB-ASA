<#
.SYNOPSIS
    Synchronizes the repository VBA modules into an existing FRTB-SA .xlsm workbook.

.DESCRIPTION
    Updates VBA code only, in a workbook built from the template by
    tools\Build-Workbook.ps1 or by hand. A full synchronization replaces the
    standard modules of one of two workbooks, and asks which one unless -Mode
    gives the answer:

    - Production: the modules of src\modules (engine, workbook macros and
      the export helper), for a workbook that is used to calculate.
    - Development: the production modules and the test modules of
      tests\modules, for a workbook that runs the regression test. The
      regression test compares the outputs of the last run with the expected
      results of the sample portfolio and writes only to Test_Expected.

    With -SourceDir only the modules found in that folder are imported, for
    example modules downloaded for a test, without copying them into the
    repository. A file is identified only by the single Attribute VB_Name it
    declares, never by its file name; a file without exactly one valid
    declaration is ignored. It is taken when that name is a module of the
    repository, or a new module whose name has a project prefix (CORE_, M_,
    TEST_), listed apart. When several files carry the same name (for example
    "CORE_SBM (1).bas"), the newest is taken. The script lists the files and
    asks for confirmation unless -Force is given, converts their line endings
    to CRLF and imports them. Every other module in the workbook is left as
    it is.

    Each imported module replaces the standard module of the same name. The
    script never writes cells, formulas, names, formatting, buttons or
    sheets, saves the workbook with its own calculation settings, and stops
    without saving if the sheet structure changes. A copy
    of the workbook is saved beside it first (*.pre-vba-sync.<time>.xlsm) and
    removed after a successful save; it is kept when the synchronization
    fails. A template change needs a new build, not a synchronization.

    Exit codes: 0 success, 1 failure (the workbook is not saved), 2 cancelled
    (the workbook is not changed).

    Requirements: Windows with desktop Excel, "Trust access to the VBA
    project object model" enabled (File > Options > Trust Center > Trust
    Center Settings > Macro Settings) and a VBA project that is not
    password-locked.

.PARAMETER WorkbookPath
    The .xlsm workbook to update. Close it in Excel first.

.PARAMETER RepoRoot
    The repository root. Defaults to the parent of this script's folder.

.PARAMETER Mode
    Production or Development, for a full synchronization. Not allowed with
    -SourceDir.

.PARAMETER SourceDir
    Import only the repository modules found in this folder.

.PARAMETER Force
    With -SourceDir, import without asking for confirmation.

.PARAMETER NoBackup
    Do not save a safety copy of the workbook first.

.PARAMETER Visible
    Show the Excel instance while it works.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\Sync-FRTB-VBA.ps1 -WorkbookPath dist\K-FRTB-ASA.xlsm -Mode Development

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\Sync-FRTB-VBA.ps1 -WorkbookPath C:\FRTB\K-FRTB-ASA.xlsm -SourceDir "$env:USERPROFILE\Downloads"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,
    [string]$RepoRoot,
    [ValidateSet('Production', 'Development')]
    [string]$Mode,
    [string]$SourceDir,
    [switch]$Force,
    [switch]$NoBackup,
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# $PSScriptRoot is not reliable as a parameter default in every Windows
# PowerShell invocation mode, so the root is resolved after binding.
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw 'Cannot determine the script path. Supply -RepoRoot explicitly.'
    }
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $scriptPath)
}

$vbextStdModule = 1
$xlCalculationManual = -4135
$msoAutomationSecurityForceDisable = 3

# With -SourceDir, a module the repository does not have yet is offered for
# import when its name has a project prefix; any other unknown file is ignored.
$NewModulePattern = '^(CORE|M|TEST)_[A-Za-z0-9_]+$'

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$PathValue)
    return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $PathValue).Path)
}

function Get-DeclaredVbaName {
    # The identity a file declares: exactly one Attribute VB_Name line with a
    # valid VBA name. The file name is never used instead, so a "CORE_SBM.bas"
    # without the attribute cannot replace CORE_SBM. Returns the name and,
    # when there is none, why.
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"\s*$') {
            $names.Add($Matches[1])
        } elseif ($line -match '^\s*Attribute\s+VB_Name\b') {
            return [pscustomobject]@{ Name = ''; Problem = 'malformed Attribute VB_Name line' }
        }
    }
    if ($names.Count -eq 0) {
        return [pscustomobject]@{ Name = ''; Problem = 'no Attribute VB_Name line; the file name is not used as its identity' }
    }
    if ($names.Count -gt 1) {
        return [pscustomobject]@{ Name = ''; Problem = 'several Attribute VB_Name lines' }
    }
    $name = $names[0]
    if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]{0,30}$') {
        return [pscustomobject]@{ Name = ''; Problem = "VB_Name '$name' is not a valid VBA name" }
    }
    return [pscustomobject]@{ Name = $name; Problem = '' }
}

function Copy-WithCrlf {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )

    # ISO-8859-1 maps every byte to one character and back, so the cp1252
    # bytes of an exported module are kept exactly; only line endings change.
    $encoding = [System.Text.Encoding]::GetEncoding(28591)
    $text = $encoding.GetString([System.IO.File]::ReadAllBytes($SourcePath))
    $text = $text.Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", "`r`n")
    [System.IO.File]::WriteAllBytes($TargetPath, $encoding.GetBytes($text))
}

function ConvertTo-SyncMode {
    # The answer to the production/development question: Production,
    # Development, or $null for anything else.
    param([AllowNull()][AllowEmptyString()][string]$Answer)

    switch -Regex (([string]$Answer).Trim()) {
        '^(?i)(p|prod|production)$' { return 'Production' }
        '^(?i)(d|dev|development)$' { return 'Development' }
    }
    return $null
}

function Get-RepositoryModules {
    # The standard modules of one workbook, as tools\Build-Workbook.ps1
    # imports them: src\modules, plus tests\modules for Development. Every
    # module must declare its name, and no name may appear twice.
    param(
        [Parameter(Mandatory = $true)][string]$RepoPath,
        [Parameter(Mandatory = $true)][ValidateSet('Production', 'Development')][string]$SyncMode
    )

    $folders = @(Join-Path (Join-Path $RepoPath 'src') 'modules')
    if ($SyncMode -eq 'Development') {
        $folders += (Join-Path (Join-Path $RepoPath 'tests') 'modules')
    }
    $modules = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($folder in $folders) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            throw "Module folder not found: $folder"
        }
        foreach ($file in @(Get-ChildItem -LiteralPath $folder -File -Filter '*.bas' | Sort-Object Name)) {
            $declared = Get-DeclaredVbaName -SourcePath $file.FullName
            if (-not [string]::IsNullOrEmpty($declared.Problem)) {
                throw "Repository module $($file.FullName): $($declared.Problem)"
            }
            $key = $declared.Name.ToLowerInvariant()
            if ($seen.ContainsKey($key)) {
                throw "Duplicate module name '$($declared.Name)' in the repository: $($seen[$key]) and $($file.FullName)"
            }
            $seen[$key] = $file.FullName
            $modules.Add([pscustomobject]@{ Name = $declared.Name; Path = $file.FullName })
        }
    }
    if ($modules.Count -eq 0) {
        throw "No .bas modules found in $($folders -join ', ')"
    }
    return $modules.ToArray()
}

function Select-FolderModules {
    # The modules to import from a folder: each .bas file whose declared name
    # is a repository module or a new project-prefixed module, the newest copy
    # of each name. Returns the picked files, the new names and why every
    # other file was ignored; prints nothing.
    param(
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [Parameter(Mandatory = $true)][hashtable]$RepositoryNames
    )

    $picked = @{}
    $newNames = @{}
    $ignored = New-Object System.Collections.Generic.List[string]
    $candidates = @(Get-ChildItem -LiteralPath $FolderPath -File | Where-Object { $_.Extension -ieq '.bas' })
    foreach ($file in $candidates) {
        $declared = Get-DeclaredVbaName -SourcePath $file.FullName
        if (-not [string]::IsNullOrEmpty($declared.Problem)) {
            $ignored.Add("$($file.Name) ($($declared.Problem))")
            continue
        }
        $name = $declared.Name
        $key = $name.ToLowerInvariant()
        if (-not $RepositoryNames.ContainsKey($key)) {
            if ($name -notmatch $NewModulePattern) {
                $ignored.Add("$($file.Name) (VB_Name '$name' is not a repository module)")
                continue
            }
            $newNames[$key] = $true
        }
        if ($picked.ContainsKey($key)) {
            if ($file.LastWriteTime -gt $picked[$key].File.LastWriteTime) {
                $ignored.Add("$($picked[$key].File.Name) (older copy of '$name')")
                $picked[$key] = [pscustomobject]@{ Name = $name; File = $file }
            } else {
                $ignored.Add("$($file.Name) (older copy of '$name')")
            }
        } else {
            $picked[$key] = [pscustomobject]@{ Name = $name; File = $file }
        }
    }
    $selection = @(foreach ($key in ($picked.Keys | Sort-Object)) {
            [pscustomobject]@{
                Name  = $picked[$key].Name
                File  = $picked[$key].File
                IsNew = $newNames.ContainsKey($key)
            }
        })
    return [pscustomobject]@{ Selection = $selection; Ignored = $ignored.ToArray() }
}

function Find-VbaComponent {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    foreach ($component in $VBProject.VBComponents) {
        if ([string]::Equals([string]$component.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $component
        }
    }
    return $null
}

function Get-SheetSnapshot {
    param([Parameter(Mandatory = $true)]$Workbook)

    $items = New-Object System.Collections.Generic.List[string]
    foreach ($sheet in $Workbook.Sheets) {
        $codeName = ''
        try { $codeName = [string]$sheet.CodeName } catch { $codeName = '<no-codename>' }
        $items.Add("$($sheet.Name)|$codeName")
    }
    return $items.ToArray()
}

function Write-ModuleList {
    param(
        [Parameter(Mandatory = $true)][string]$Heading,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Items
    )

    if ($Items.Count -eq 0) { return }
    Write-Host $Heading
    foreach ($item in $Items) {
        Write-Host ('  {0,-22} {1,-30} {2:yyyy-MM-dd HH:mm}' -f $item.Name, $item.File.Name, $item.File.LastWriteTime)
    }
}

$excel = $null
$workbook = $null
$saved = $false
$backupPath = $null
$stagePath = $null

try {
    $workbookFullPath = Resolve-FullPath -PathValue $WorkbookPath
    $repoFullPath = Resolve-FullPath -PathValue $RepoRoot
    if ([System.IO.Path]::GetExtension($workbookFullPath) -ine '.xlsm') {
        throw 'WorkbookPath must point to an .xlsm file.'
    }
    if (-not [string]::IsNullOrWhiteSpace($SourceDir) -and -not [string]::IsNullOrWhiteSpace($Mode)) {
        throw 'Use -Mode only for a full synchronization, not with -SourceDir.'
    }

    # The development set lists every repository module; it also tells which
    # names a folder may supply.
    $repositoryNames = @{}
    foreach ($module in (Get-RepositoryModules -RepoPath $repoFullPath -SyncMode Development)) {
        $repositoryNames[$module.Name.ToLowerInvariant()] = $module.Path
    }

    $stagePath = Join-Path ([System.IO.Path]::GetTempPath()) ('frtb-sa-vba-sync-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $stagePath
    $imports = New-Object System.Collections.Generic.List[object]
    $syncMode = $null

    if ([string]::IsNullOrWhiteSpace($SourceDir)) {
        $syncMode = $Mode
        if ([string]::IsNullOrWhiteSpace($syncMode)) {
            Write-Host 'Which workbook is this?'
            Write-Host '  P = production: engine, workbook macros and export helper (src\modules)'
            Write-Host '  D = development: production modules and the test modules (tests\modules);'
            Write-Host '      the tests change the input sheets'
            $syncMode = ConvertTo-SyncMode -Answer (Read-Host 'Production or development? [P/D]')
            if ($null -eq $syncMode) {
                Write-Host 'Cancelled: the workbook was not changed.'
                exit 2
            }
        }
        $sourceLabel = $repoFullPath
        foreach ($module in (Get-RepositoryModules -RepoPath $repoFullPath -SyncMode $syncMode)) {
            $imports.Add($module)
        }
        Write-Host "Mode: $syncMode ($($imports.Count) modules)"
    } else {
        $sourceDirFull = Resolve-FullPath -PathValue $SourceDir
        if (-not (Test-Path -LiteralPath $sourceDirFull -PathType Container)) {
            throw "SourceDir is not a folder: $sourceDirFull"
        }
        $sourceLabel = $sourceDirFull
        $found = Select-FolderModules -FolderPath $sourceDirFull -RepositoryNames $repositoryNames
        if ($found.Ignored.Count -gt 0) {
            Write-Host 'Ignored:'
            foreach ($line in $found.Ignored) { Write-Host "  $line" }
        }
        if ($found.Selection.Count -eq 0) {
            throw "No repository VBA module found in $sourceDirFull."
        }
        Write-ModuleList -Heading "Modules to import from ${sourceDirFull}:" -Items @($found.Selection | Where-Object { -not $_.IsNew })
        Write-ModuleList -Heading 'New modules, not in the repository (imported as new standard modules):' -Items @($found.Selection | Where-Object { $_.IsNew })
        if (-not $Force) {
            $answer = Read-Host "Import these $($found.Selection.Count) module(s) into the workbook? [Y/N]"
            if ($answer -notmatch '^\s*(y|yes|s|si)\s*$') {
                Write-Host 'Cancelled: the workbook was not changed.'
                exit 2
            }
        }
        foreach ($item in $found.Selection) {
            $imports.Add([pscustomobject]@{ Name = $item.Name; Path = $item.File.FullName })
        }
    }
    Write-Host ''

    # Every module is imported from a CRLF copy named after its VB_Name.
    foreach ($module in $imports) {
        $staged = Join-Path $stagePath ($module.Name + '.bas')
        Copy-WithCrlf -SourcePath $module.Path -TargetPath $staged
        $module | Add-Member -NotePropertyName Staged -NotePropertyValue $staged
    }

    if (-not $NoBackup) {
        $directory = [System.IO.Path]::GetDirectoryName($workbookFullPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($workbookFullPath)
        $backupPath = Join-Path $directory ("$baseName.pre-vba-sync." + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlsm')
        Copy-Item -LiteralPath $workbookFullPath -Destination $backupPath
        Write-Host "Backup:     $backupPath"
    }
    Write-Host "Workbook:   $workbookFullPath"
    Write-Host "Source:     $sourceLabel"
    Write-Host "Modules:    $($imports.Count)"
    Write-Host ''

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    # Some Excel builds reject Calculation changes until a workbook is open.
    # The first workbook opened sets the calculation mode, which is saved
    # with it: keep the workbook's own settings and restore them before the
    # save, so that only the VBA changes.
    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $false)
    if ($workbook.ReadOnly) {
        throw 'The workbook opened read-only. Close it in every other Excel instance.'
    }
    $originalCalculation = $excel.Calculation
    $originalCalculateBeforeSave = $excel.CalculateBeforeSave
    $excel.Calculation = $xlCalculationManual
    $excel.CalculateBeforeSave = $false
    $before = Get-SheetSnapshot -Workbook $workbook

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    } catch {
        throw ("Cannot access the VBA project. Enable 'Trust access to the VBA project object model' " +
            'and make sure the VBA project is not password-locked.')
    }

    # A production synchronization never removes modules: test modules
    # already in the workbook stay, and the user is told.
    if ($syncMode -eq 'Production') {
        $testModules = @(foreach ($component in $vbProject.VBComponents) {
                if ($component.Name -like 'TEST_*') { $component.Name }
            })
        if ($testModules.Count -gt 0) {
            Write-Warning ('This workbook holds test modules, which a production synchronization does not remove: ' +
                ($testModules -join ', ') + '. For a production workbook, build it with tools\Build-Workbook.ps1 -Production.')
        }
    }

    foreach ($module in $imports) {
        $existing = Find-VbaComponent -VBProject $vbProject -Name $module.Name
        if ($null -ne $existing) {
            if ([int]$existing.Type -ne $vbextStdModule) {
                throw "'$($module.Name)' exists in the workbook but is not a standard module."
            }
            $vbProject.VBComponents.Remove($existing)
        }
        $imported = $vbProject.VBComponents.Import($module.Staged)
        if ($null -eq $existing) {
            Write-Host ('{0,-22} new module imported' -f $imported.Name)
        } else {
            Write-Host ('{0,-22} replaced' -f $imported.Name)
        }
    }

    $after = Get-SheetSnapshot -Workbook $workbook
    if (($before -join "`n") -cne ($after -join "`n")) {
        throw 'The sheet structure changed during the synchronization. The workbook was not saved.'
    }

    $excel.Calculation = $originalCalculation
    $excel.CalculateBeforeSave = $originalCalculateBeforeSave
    $workbook.Save()
    $saved = $true

    # The backup is only a safety copy; remove it after a successful save so
    # that repeated synchronizations do not accumulate files.
    if ($null -ne $backupPath -and (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Remove-Item -LiteralPath $backupPath -Force
    }
    Write-Host ''
    Write-Host 'Synchronized. Open the workbook and run Debug > Compile VBAProject.'
} catch {
    Write-Error $_ -ErrorAction Continue
    if ($null -ne $workbook -and -not $saved) {
        Write-Warning 'The synchronization failed; the workbook is closed without saving.'
    }
    if ($null -ne $backupPath -and (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Write-Warning "Safety copy kept: $backupPath"
    }
    exit 1
} finally {
    if ($null -ne $stagePath -and (Test-Path -LiteralPath $stagePath -PathType Container)) {
        try { Remove-Item -LiteralPath $stagePath -Recurse -Force } catch {}
    }
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
