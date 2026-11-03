<#
.SYNOPSIS
    Records an Excel run of K-FRTB-SA for the evidence record, in one command.

.DESCRIPTION
    Runs every stage of .github/excel-evidence-policy.json on the commit that
    is checked out, and writes the evidence bundle test-results\excel-run-<sha7>
    (ignored by Git):

    1. session.log: the candidate, the start time, the operator, and the
       environment read from Excel and the registry: the Excel version,
       build and bitness, the Windows version and architecture, the decimal
       and list separators, the macro policy and the VBA project access.
    2. The development build with the regression test (build-development.log).
    3. The development compile: Debug > Compile VBAProject run through the
       VBA editor, and the references of the project, in session.log.
    4. The development export and comparison (round-trip-development.log).
    5. The production build, compile, export and comparison.
    6. The finish time; then tools\record_excel_run.py writes host.json from
       the logs and runs tools\check_excel_evidence.py on it.

    Nothing is typed by hand. A failing stage keeps its log, and the stages
    that depend on it are recorded as NOT_RUN.

    Requirements: Windows with desktop Excel, Git and Python 3 on PATH, a
    clean checkout of the candidate, and "Trust access to the VBA project
    object model" on (File > Options > Trust Center > Trust Center Settings >
    Macro Settings). The script changes no setting.

    If a compile error is found, Excel shows it in a message box: read it,
    click OK, and the run goes on with that compile recorded as FAIL.

.PARAMETER TrustChange
    If you turned on "Trust access to the VBA project object model" only for
    this run, say so here, for example "VBA project access turned on for the
    run and off again after it". It is recorded as a Trust Center change.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\Record-FRTB-ExcelRun.ps1
#>
[CmdletBinding()]
param(
    [string]$TrustChange
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

function Write-Session {
    param([Parameter(Mandatory = $true)][string]$Line)
    Add-Content -Path $script:session -Value $Line -Encoding UTF8
}

function Get-Timestamp {
    $now = [DateTimeOffset]::Now
    return $now.ToString('yyyy-MM-ddTHH:mm:ss') + $now.ToString('zzz')
}

function Release-Com {
    param($Object)
    if ($null -ne $Object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function Get-RegistryValue {
    param([string]$Path, [string]$Name)
    $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) { return $null }
    return $item.$Name
}

function Get-MacroPolicy {
    $names = @{ 1 = 'Enable all macros'; 2 = 'Disable VBA macros with notification';
                3 = 'Disable VBA macros except digitally signed macros'; 4 = 'Disable VBA macros without notification' }
    foreach ($path in 'HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Security',
                      'HKCU:\Software\Microsoft\Office\16.0\Excel\Security') {
        $value = Get-RegistryValue $path 'VBAWarnings'
        if ($null -ne $value) {
            $label = $names[[int]$value]
            if (-not $label) { $label = "VBAWarnings=$value" }
            if ($path -like '*Policies*') { $label += ' (set by policy)' }
            return $label
        }
    }
    return 'Disable VBA macros with notification (Excel default, no registry value)'
}

function Get-VbaProjectAccess {
    foreach ($path in 'HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Security',
                      'HKCU:\Software\Microsoft\Office\16.0\Excel\Security') {
        $value = Get-RegistryValue $path 'AccessVBOM'
        if ($null -ne $value) {
            if ([int]$value -eq 1) { return 'On' } else { return 'Off' }
        }
    }
    return 'Off (no registry value)'
}

function Get-WindowsName {
    $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $product = Get-RegistryValue $key 'ProductName'
    $display = Get-RegistryValue $key 'DisplayVersion'
    $build = Get-RegistryValue $key 'CurrentBuild'
    $revision = Get-RegistryValue $key 'UBR'
    # Windows 11 keeps "Windows 10" in ProductName; the build number tells them apart
    if ($build -and [int]$build -ge 22000 -and $product -like 'Windows 10*') {
        $product = $product -replace '^Windows 10', 'Windows 11'
    }
    return "$product $display (build $build.$revision)"
}

function Write-Environment {
    $excel = New-Object -ComObject Excel.Application
    try {
        $excel.Visible = $false
        $version = [string]$excel.Version
        $build = Get-RegistryValue 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' 'VersionToReport'
        if (-not $build) { $build = "$version.$($excel.Build)" }
        $platform = Get-RegistryValue 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' 'Platform'
        if ($platform -eq 'x64') {
            $bitness = '64-bit'
        } elseif ($platform -eq 'x86') {
            $bitness = '32-bit'
        } elseif ([string]$excel.Path -like '*Program Files (x86)*') {
            $bitness = '32-bit'
        } else {
            $bitness = '64-bit'
        }
        # 3 = xlDecimalSeparator, 5 = xlListSeparator: the separators Excel uses
        $decimal = [string]$excel.International(3)
        $list = [string]$excel.International(5)
        if (-not $excel.UseSystemSeparators) {
            $decimal = "$([string]$excel.DecimalSeparator) (Excel option, not the system one)"
        }
        Write-Session "EXCEL_VERSION=$version"
        Write-Session "EXCEL_BUILD=$build"
        Write-Session "OFFICE_BITNESS=$bitness"
        Write-Session "OS=$(Get-WindowsName)"
        Write-Session "OS_ARCHITECTURE=$env:PROCESSOR_ARCHITECTURE"
        Write-Session "DECIMAL_SEPARATOR=$decimal"
        Write-Session "LIST_SEPARATOR=$list"
        Write-Session "MACRO_POLICY=$(Get-MacroPolicy)"
        Write-Session "VBA_PROJECT_ACCESS=$(Get-VbaProjectAccess)"
    } finally {
        $excel.Quit()
        Release-Com $excel
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Invoke-Logged {
    # Runs a command, writes its output to a log in the bundle, returns its exit code
    param([string]$Log, [string[]]$Header, [scriptblock]$Command)
    # Windows PowerShell 5.1 turns redirected stderr of a child process into
    # error records, which 'Stop' would make terminating: the stage would end
    # without its log. Here they are only logged; the exit code decides.
    $ErrorActionPreference = 'Continue'
    $lines = @($Header)
    $lines += @(& $Command 2>&1 | ForEach-Object { [string]$_ })
    $code = $LASTEXITCODE
    $lines | Out-File -FilePath (Join-Path $script:bundle $Log) -Encoding utf8
    return $code
}

# The VBA editor's objects (VBE, its windows and command bars) give Windows
# PowerShell no type information: a property read through the normal adapter
# returns nothing. These helpers call them through IDispatch by name. The
# leading comma returns a collection (CommandBars, References) as the object
# itself: PowerShell would otherwise unroll it into an array of its items.
function Get-ComProperty {
    param($Object, [string]$Name, [object[]]$Arguments = @())
    return ,([System.__ComObject].InvokeMember($Name, [System.Reflection.BindingFlags]::GetProperty,
                                               $null, $Object, $Arguments))
}

function Set-ComProperty {
    param($Object, [string]$Name, $Value)
    [void][System.__ComObject].InvokeMember($Name, [System.Reflection.BindingFlags]::SetProperty,
                                            $null, $Object, @($Value))
}

function Invoke-ComMethod {
    param($Object, [string]$Name, [object[]]$Arguments = @())
    return ,([System.__ComObject].InvokeMember($Name, [System.Reflection.BindingFlags]::InvokeMethod,
                                               $null, $Object, $Arguments))
}

function Find-CompileControl {
    param($Vbe)
    # 578 is the ID of Debug > Compile VBAProject; 1 = msoControlButton
    $bars = Get-ComProperty $Vbe 'CommandBars'
    $control = Invoke-ComMethod $bars 'FindControl' @(1, 578)
    if ($null -eq $control) {
        $control = Invoke-ComMethod $bars 'FindControl' @([System.Reflection.Missing]::Value, 578)
    }
    return $control
}

function Write-References {
    # The references of the workbook's VBA project, recorded apart from the
    # compile so that a failed compile check still records them
    param([string]$Key, $Workbook)
    try {
        $project = $Workbook.VBProject
        $references = Get-ComProperty $project 'References'
        $count = [int](Get-ComProperty $references 'Count')
        $names = @(for ($index = 1; $index -le $count; $index++) {
            [string](Get-ComProperty (Invoke-ComMethod $references 'Item' @($index)) 'Description')
        })
        Write-Session "REFERENCES_$($Key.Replace('COMPILE_', ''))=$($names -join '|')"
        if ($Key -eq 'COMPILE_DEVELOPMENT') { Write-Session "REFERENCES=$($names -join '|')" }
    } catch {
        Write-Session "REFERENCES_$($Key.Replace('COMPILE_', ''))_ERROR=$($_.Exception.Message)"
    }
}

function Test-Compile {
    # Compiles the workbook's VBA project with the VBA editor's own command and
    # records PASS when the command is then disabled, as Excel does after a
    # compile that found no error; records the project's references first.
    param([string]$Key, [string]$WorkbookPath)
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.DisplayAlerts = $false
        # 3 = msoAutomationSecurityForceDisable: no macro runs while it is open
        $excel.AutomationSecurity = 3
        $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
        $excel.Visible = $true
        Write-References $Key $workbook
        $vbe = Get-ComProperty $excel 'VBE'
        if ($null -eq $vbe) { throw 'Excel returned no VBA editor object' }
        try { Set-ComProperty (Get-ComProperty $vbe 'MainWindow') 'Visible' $true } catch { }
        $control = Find-CompileControl $vbe
        if ($null -eq $control) {
            Write-Session "${Key}=FAIL"
            Write-Session "${Key}_DETAIL=Debug > Compile VBAProject not found in the VBA editor"
            return
        }
        $wasEnabled = [bool](Get-ComProperty $control 'Enabled')
        if ($wasEnabled) { [void](Invoke-ComMethod $control 'Execute') }
        if (-not $wasEnabled) {
            # Excel disables the command for a project saved fully compiled
            Write-Session "${Key}=PASS"
            Write-Session "${Key}_DETAIL=Debug > Compile VBAProject: the project opened fully compiled"
        } elseif ([bool](Get-ComProperty $control 'Enabled')) {
            Write-Session "${Key}=FAIL"
            Write-Session "${Key}_DETAIL=Debug > Compile VBAProject reported a compile error"
        } else {
            Write-Session "${Key}=PASS"
            Write-Session "${Key}_DETAIL=Debug > Compile VBAProject: no error"
        }
    } catch {
        Write-Session "${Key}=FAIL"
        Write-Session "${Key}_DETAIL=Compile check stopped: $($_.Exception.Message)"
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false); Release-Com $workbook }
        if ($null -ne $excel) { $excel.Quit(); Release-Com $excel }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Invoke-Variant {
    # Builds, compiles, exports and compares one workbook
    param([string]$Mode)
    $label = $Mode.ToLowerInvariant()
    $workbook = Join-Path $repoRoot "dist\K-FRTB-SA-$label.xlsm"
    $buildArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'tools\Build-Workbook.ps1',
                   '-OutputPath', $workbook, '-Force')
    if ($Mode -eq 'Development') { $buildArgs += '-RunTests' } else { $buildArgs += '-Production' }
    Write-Host "== $Mode build"
    $code = Invoke-Logged "build-$label.log" @("CANDIDATE=$sha") { & powershell.exe @buildArgs }
    Write-Session "BUILD_$($Mode.ToUpperInvariant())_EXIT=$code"
    # Build-Workbook.ps1 exits 1 after a failed regression but still saves the
    # workbook: go on, and the regression stage records the failure
    $saved = Select-String -Path (Join-Path $script:bundle "build-$label.log") -Pattern '^Saved \S' -Quiet
    if ($code -gt 1 -or -not $saved) {
        Write-Host "   build failed (exit code $code): see build-$label.log"
        return
    }
    if ($code -ne 0) { Write-Host "   the regression test failed: see build-$label.log" }
    Write-Host "== $Mode compile"
    Test-Compile "COMPILE_$($Mode.ToUpperInvariant())" $workbook
    Write-Host "== $Mode export and comparison"
    $export = Join-Path $script:bundle "export-$label"
    $code = Invoke-Logged "export-$label.txt" @() {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\Export-FRTB-VBA.ps1 -WorkbookPath $workbook -OutDir $export
    }
    Write-Session "EXPORT_$($Mode.ToUpperInvariant())_EXIT=$code"
    [void](Invoke-Logged "round-trip-$label.log" @() {
        & python tools\compare_vba_export.py --export $export --mode $label --candidate-sha $sha --workbook $workbook
    })
}

# --- Candidate and bundle -------------------------------------------------------

foreach ($tool in 'git', 'python') {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
        throw "$tool was not found on PATH"
    }
}
$status = @(& git status --porcelain --untracked-files=no)
if ($status.Count -gt 0) {
    throw "The checkout has local changes; commit, stash or discard them first:`n$($status -join "`n")"
}
$sha = (& git rev-parse HEAD).Trim()
$bundle = Join-Path $repoRoot ("test-results\excel-run-" + $sha.Substring(0, 7))
if (Test-Path $bundle) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    Rename-Item -Path $bundle -NewName ((Split-Path -Leaf $bundle) + "-before-$stamp")
}
New-Item -ItemType Directory -Force -Path $bundle | Out-Null
$session = Join-Path $bundle 'session.log'

if ((Get-VbaProjectAccess) -ne 'On') {
    throw "Turn on 'Trust access to the VBA project object model' (File > Options > Trust Center > Trust Center Settings > Macro Settings), then run again. If you turn it on only for this run, pass -TrustChange with what you changed."
}

Write-Host "Candidate: $sha"
Write-Host "Bundle:    $bundle"
Write-Session "CANDIDATE=$sha"
Write-Session "STARTED=$(Get-Timestamp)"
Write-Session "OPERATOR=$env:USERNAME, $env:COMPUTERNAME"
if ($TrustChange) { Write-Session "TRUST_CHANGES=$TrustChange" }
Write-Host '== Environment'
Write-Environment

Invoke-Variant 'Development'
Invoke-Variant 'Production'

Write-Session "FINISHED=$(Get-Timestamp)"
Write-Host '== Evidence record'
& python tools\record_excel_run.py $bundle --candidate $sha
$result = $LASTEXITCODE
Write-Host ''
Write-Host "Send me the zipped folder $bundle and the validation folder next to it."
exit $result
