<#
.SYNOPSIS
    Runs every Excel check of K-FRTB-ASA on one commit, unattended, and writes
    a summary and a zipped bundle (#274).

.DESCRIPTION
    For the branch or commit given, on the Windows host with desktop Excel:
    1. refuses a checkout with changes to tracked files, fetches the branch
       and checks its commit out (detached);
    2. checks "Trust access to the VBA project object model"; when it is off
       it asks, and turns it on for your Windows user only when you answer Y.
       It also asks once whether to run the Trust Center error check below,
       unless a group policy sets the access: that check then cannot run.
       Nothing else is asked: the rest runs unattended;
    3. runs tools\Record-FRTB-ExcelRun.ps1: both builds, both compiles, the
       regression and both round trips, written to the evidence bundle
       test-results\excel-run-SHORTSHA and validated;
    4. opens a copy of the development workbook with its macros enabled for
       that Excel session only (AutomationSecurity Low; the Trust Center is
       not changed) and runs TEST_Acceptance.Test_RunAcceptanceInExcel_Silent:
       the sample run, the state after a failed run and after Clear outputs
       (#85), numeric text with a comma and with a point, the frozen panes,
       the autofilters, the Control descriptions (#203) and the invalid-row
       variant (#28). It writes acceptance.txt, and pictures of the sheets
       after a failed run, after Clear outputs and after the sample run;
    5. builds the workbook from source with the regression test, and the
       template from source, and records the SHA-256 of that template and of
       the committed one, with their differences (tools\compare_templates.py);
    6. when you answered Y in step 2, checks the build's Trust Center error
       path: it turns the VBA project access off for one build and restores
       your setting as it was. When a group policy sets the access, it reports
       the check as not run;
    7. checks the synchronization (#72) on copies of both workbooks: Production
       and Development from the checkout, each followed by a compile and a
       run (and the regression for Development), with the workbook's
       calculation mode kept; a folder of downloaded modules; the failure path
       while the workbook is open in another Excel; and both launchers;
    8. writes summary.txt, zips the folder test-results\excel-check-SHORTSHA-
       TIME, copies the summary to the clipboard and shows the zip in Explorer.

    Each Excel this script starts is its own instance: your open workbooks
    and add-ins are not touched. A watchdog closes the Excel instances started
    by automation during the run after -TimeoutMinutes, so a dialog nobody
    answers cannot hang the run; an Excel you open yourself is never closed.
    For the run, Microsoft Print to PDF is the default printer, so Excel does
    not wait on a slow printer when it sets page setup; your default printer
    is restored at the end, also when the run fails. Without Microsoft Print
    to PDF the default printer is not changed. The script never pushes and
    changes no other Excel or Windows setting.
    Exit codes: 0 every check passed, 1 a check failed, 2 the run could not
    start (local changes, no Git or Python, the VBA project access refused).

.PARAMETER Branch
    The branch of origin to check out. Default: release/v0.2.0.

.PARAMETER Commit
    An exact commit to check instead of the head of -Branch.

.PARAMETER TimeoutMinutes
    The watchdog limit for the whole run. Default: 60.

.PARAMETER TrustCenterCheck
    Run the Trust Center error check without asking.

.PARAMETER NoPrompt
    Ask nothing: the VBA project access must already be on, and the Trust
    Center error check runs only with -TrustCenterCheck.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\Check-FRTB-Excel.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\Check-FRTB-Excel.ps1 -Branch fix/123-slug
#>
[CmdletBinding()]
param(
    [string]$Branch = 'release/v0.2.0',
    [string]$Commit,
    [int]$TimeoutMinutes = 60,
    [switch]$TrustCenterCheck,
    [switch]$NoPrompt
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$msoAutomationSecurityLow = 1
$msoAutomationSecurityForceDisable = 3
$xlScreen = 1
$xlBitmap = 2
$amountTolerance = 0.01
$script:results = New-Object System.Collections.Generic.List[object]

function Get-Timestamp {
    $now = [DateTimeOffset]::Now
    return $now.ToString('yyyy-MM-ddTHH:mm:ss') + $now.ToString('zzz')
}

function Write-Log {
    param([string]$Text)
    $line = '{0}  {1}' -f (Get-Date).ToString('HH:mm:ss'), $Text
    Write-Host $line
    if ($script:logPath) { Add-Content -Path $script:logPath -Value $line -Encoding UTF8 }
}

function Add-Result {
    # One line of the summary: a check, PASS, FAIL or NOT_RUN, and a detail
    param([string]$Name, [string]$Result, [string]$Detail)
    $script:results.Add([pscustomobject]@{ Name = $Name; Result = $Result; Detail = $Detail })
    Write-Log ('{0,-28} {1,-8} {2}' -f $Name, $Result, $Detail)
}

function Stop-NotStarted {
    param([string]$Reason)
    Write-Host ''
    Write-Host "NOT STARTED: $Reason"
    exit 2
}

function Release-Com {
    param($Object)
    if ($null -ne $Object) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch { }
    }
}

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

function Get-SecurityKey {
    param([string]$Version)
    return "HKCU:\Software\Microsoft\Office\$Version\Excel\Security"
}

function Get-VbaAccess {
    # The value Excel uses: a policy, when set, wins over the user's setting
    param([string]$Version)
    foreach ($root in @('HKCU:\Software\Policies\Microsoft\Office', 'HKCU:\Software\Microsoft\Office')) {
        $item = Get-ItemProperty -LiteralPath "$root\$Version\Excel\Security" -Name 'AccessVBOM' -ErrorAction SilentlyContinue
        if ($null -ne $item) {
            return [pscustomobject]@{ Value = [int]$item.AccessVBOM; Policy = ($root -like '*Policies*') }
        }
    }
    return [pscustomobject]@{ Value = 0; Policy = $false }
}

function Set-VbaAccess {
    param([string]$Version, [int]$Value)
    $key = Get-SecurityKey $Version
    if (-not (Test-Path -LiteralPath $key)) { [void](New-Item -Path $key -Force) }
    Set-ItemProperty -LiteralPath $key -Name 'AccessVBOM' -Value $Value -Type DWord
}

function Get-UserVbaAccess {
    # The user's own setting, or $null when it is absent; a policy is not read
    param([string]$Version)
    $item = Get-ItemProperty -LiteralPath (Get-SecurityKey $Version) -Name 'AccessVBOM' -ErrorAction SilentlyContinue
    if ($null -eq $item) { return $null }
    return [int]$item.AccessVBOM
}

function Restore-UserVbaAccess {
    # Puts the user's setting back as Get-UserVbaAccess found it, absent included
    param([string]$Version, $Value)
    if ($null -eq $Value) {
        Remove-ItemProperty -LiteralPath (Get-SecurityKey $Version) -Name 'AccessVBOM' -ErrorAction SilentlyContinue
    } else {
        Set-VbaAccess $Version ([int]$Value)
    }
}

function Get-DefaultPrinterName {
    $printer = Get-CimInstance -ClassName Win32_Printer -Filter 'Default = TRUE' -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -eq $printer) { return $null }
    return [string]$printer.Name
}

function Set-DefaultPrinter {
    # True when the printer exists and Windows made it the default
    param([string]$Name)
    $wql = $Name.Replace('\', '\\').Replace("'", "\'")
    $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name = '$wql'" -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -eq $printer) { return $false }
    $result = Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter -ErrorAction SilentlyContinue
    return ($null -ne $result -and $result.ReturnValue -eq 0)
}

function Read-YesNo {
    param([string]$Question)
    if ($NoPrompt) { return $false }
    $answer = Read-Host "$Question [Y/N]"
    return ($answer -match '^\s*(y|yes|s|si)\s*$')
}

function Invoke-Tool {
    # Runs a PowerShell tool of the checkout in its own process, writes its
    # output to a log of the bundle and returns the exit code
    param([string]$LogName, [string[]]$Arguments)
    $logFile = Join-Path $script:bundle $LogName
    # A native tool's standard error is output here, not a terminating error
    $ErrorActionPreference = 'Continue'
    # The tool's output is shown as it runs and kept in the log
    $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass @Arguments 2>&1 |
        ForEach-Object { $line = [string]$_; Write-Host "   $line"; $line }
    $code = $LASTEXITCODE
    $output | Set-Content -Path $logFile -Encoding UTF8
    Add-Content -Path $logFile -Value "EXIT=$code" -Encoding UTF8
    return $code
}

function New-Excel {
    # A new Excel instance; Low enables the macros of the workbooks it opens
    # for this instance only, ForceDisable runs none
    param([int]$Security)
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = $Security
    $excel.Visible = $true
    return $excel
}

function Close-Excel {
    param($Excel, $Workbook)
    if ($null -ne $Workbook) { try { $Workbook.Close($false) } catch { }; Release-Com $Workbook }
    if ($null -ne $Excel) { try { $Excel.Quit() } catch { }; Release-Com $Excel }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Invoke-Macro {
    # Application.Run of a macro of that workbook, with up to two arguments
    param($Excel, $Workbook, [string]$Macro, [object[]]$Arguments = @())
    $name = "'" + $Workbook.Name + "'!" + $Macro
    switch ($Arguments.Count) {
        0 { return $Excel.Run($name) }
        1 { return $Excel.Run($name, $Arguments[0]) }
        2 { return $Excel.Run($name, $Arguments[0], $Arguments[1]) }
        default { throw "Invoke-Macro takes at most two arguments, not $($Arguments.Count)" }
    }
}

function Test-ProjectCompiles {
    # Debug > Compile VBAProject through the VBA editor's own command (ID 578):
    # Excel disables it after a compile that found no error
    param($Excel)
    $vbe = Get-ComProperty $Excel 'VBE'
    $bars = Get-ComProperty $vbe 'CommandBars'
    $control = Invoke-ComMethod $bars 'FindControl' @(1, 578)
    if ($null -eq $control) { $control = Invoke-ComMethod $bars 'FindControl' @([System.Reflection.Missing]::Value, 578) }
    if ($null -eq $control) { return $false }
    if ([bool](Get-ComProperty $control 'Enabled')) { [void](Invoke-ComMethod $control 'Execute') }
    return -not [bool](Get-ComProperty $control 'Enabled')
}

function Get-AutomationExcelIds {
    # The process IDs of the Excel instances started through COM, which have
    # /automation on their command line; an Excel opened by hand has not
    @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'EXCEL.EXE'" -ErrorAction SilentlyContinue |
        Where-Object { [string]$_.CommandLine -match '(?i)[/-]automation\b' } |
        ForEach-Object { [int]$_.ProcessId })
}

function Stop-TimedOutLauncher {
    # A launcher past its limit: end its whole process tree (cmd.exe, the
    # powershell.exe it runs) and the Excel instances started by automation
    # since the launcher began, then wait for them to go (#279). An Excel
    # opened by hand, or running before the launcher, is never touched
    param($Process, [int[]]$ExcelBefore)
    # taskkill reports a tree that already ended on standard error: not an error here
    $ErrorActionPreference = 'Continue'
    & taskkill.exe /PID $Process.Id /T /F 2>&1 | Out-Null
    $stray = @(Get-AutomationExcelIds | Where-Object { $ExcelBefore -notcontains $_ })
    foreach ($id in $stray) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
    try { [void]$Process.WaitForExit(30000) } catch { }
    foreach ($id in $stray) { Wait-Process -Id $id -Timeout 30 -ErrorAction SilentlyContinue }
}

function Assert-TimeLeft {
    # No stage starts after the time limit: the watchdog closes the Excel
    # instances still running, and the stages left are reported as not run
    if ((Get-Date) -ge $script:deadline) {
        throw "the $TimeoutMinutes-minute limit was reached; the remaining checks did not run"
    }
}

function Find-TagRow {
    # The row of a tagged table, as CORE_Util.FindTagRow finds it; that
    # module is private to the VBA project, so Application.Run cannot call it
    param($Sheet, [string]$Tag)
    for ($r = 1; $r -le 400; $r++) {
        if (([string]$Sheet.Cells.Item($r, 3).Value2).Trim().ToUpperInvariant().StartsWith("[$Tag]")) { return $r }
    }
    throw "$($Sheet.Name) has no [$Tag] table"
}

function Get-ExpectedTotal {
    # The sample's TOTAL on Test_Expected, written there from the reference
    param($Workbook)
    $ws = $Workbook.Worksheets.Item('Test_Expected')
    for ($r = 6; $r -le 120; $r++) {
        if ([string]$ws.Cells.Item($r, 3).Value2 -eq 'TOTAL') { return [double]$ws.Cells.Item($r, 4).Value2 }
    }
    throw 'Test_Expected has no TOTAL row'
}

function Export-SheetPicture {
    # Best effort: a picture of the top of a sheet, as Excel shows it
    param($Workbook, [string]$Sheet, [string]$Address, [string]$Path)
    try {
        $ws = $Workbook.Worksheets.Item($Sheet)
        [void]$ws.Activate()
        $range = $ws.Range($Address)
        [void]$range.CopyPicture($xlScreen, $xlBitmap)
        $chartObject = $ws.ChartObjects().Add(0, 0, [double]$range.Width + 4, [double]$range.Height + 4)
        [void]$chartObject.Activate()
        [void]$chartObject.Chart.Paste()
        [void]$chartObject.Chart.Export($Path)
        [void]$chartObject.Delete()
        return $true
    } catch {
        Write-Log "   picture of $Sheet not taken: $($_.Exception.Message)"
        return $false
    }
}

# ------------------------------------------------------------------ start
foreach ($tool in @('git', 'python')) {
    if ($null -eq (Get-Command $tool -ErrorAction SilentlyContinue)) { Stop-NotStarted "$tool is not on PATH." }
}
$changes = & git status --porcelain --untracked-files=no
if ($changes) { Stop-NotStarted "the checkout has local changes; commit or stash them first:`n$($changes -join "`n")" }

Write-Host "== Checking out $(if ($Commit) { $Commit } else { "origin/$Branch" })"
& git fetch origin $Branch
if ($LASTEXITCODE -ne 0) { Stop-NotStarted "git fetch origin $Branch failed." }
$target = if ($Commit) { $Commit } else { "origin/$Branch" }
$sha = (& git rev-parse --verify "$target^{commit}").Trim()
if ($LASTEXITCODE -ne 0 -or -not $sha) { Stop-NotStarted "$target is not a commit." }
& git -c advice.detachedHead=false switch --detach $sha
if ($LASTEXITCODE -ne 0) { Stop-NotStarted "git switch --detach $sha failed." }
$short = $sha.Substring(0, 7)

$script:bundle = Join-Path $repoRoot ("test-results\excel-check-$short-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
[void](New-Item -ItemType Directory -Path $script:bundle -Force)
$script:logPath = Join-Path $script:bundle 'check.log'
Write-Log "CANDIDATE=$sha"
Write-Log "BRANCH=$Branch"
Write-Log "START=$(Get-Timestamp)"

# The Excel version tells which registry key holds the VBA project access.
# Excel that cannot start is a check that could not start: exit 2, not 1
try {
    $probe = New-Object -ComObject Excel.Application
    $excelVersion = [string]$probe.Version
    Write-Log "EXCEL=$excelVersion $([string]$probe.Build)"
    Close-Excel $probe $null
} catch {
    Stop-NotStarted "Excel could not be started: $($_.Exception.Message)"
}

$access = Get-VbaAccess $excelVersion
$trustChange = ''
if ($access.Value -ne 1) {
    if ($access.Policy) { Stop-NotStarted 'a group policy turns the VBA project access off.' }
    if (-not (Read-YesNo "Trust access to the VBA project object model is off. Turn it on for your Windows user?")) {
        Stop-NotStarted 'Trust access to the VBA project object model is off (File > Options > Trust Center > Trust Center Settings > Macro Settings).'
    }
    Set-VbaAccess $excelVersion 1
    $trustChange = 'VBA project access turned on by tools\Check-FRTB-Excel.ps1 at the operator''s request'
    Write-Log 'TRUST_CHANGE=VBA project access turned on'
}
$runTrustCheck = [bool]$TrustCenterCheck
if ($access.Policy) {
    # Only the user's setting can be turned off; the policy would keep the access on
    Write-Log 'TRUST_CHECK=not run: a group policy sets the VBA project access'
} elseif (-not $runTrustCheck) {
    $runTrustCheck = Read-YesNo 'Also check the build''s Trust Center error path? It turns the VBA project access off for one build and back on.'
}
Write-Host ''
Write-Host 'The rest runs unattended. Do not use the Excel windows that open.'
Write-Host ''

# Watchdog: once the limit passes, it closes the Excel instances started by
# automation during this run. Excel started through COM, by this script or by
# the tools it runs, has /automation on its command line; an Excel the
# operator opens has not, so it is never closed. It keeps closing them until
# the run stops it, and the run starts no new stage after the limit (#274)
$excelBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$script:deadline = (Get-Date).AddMinutes($TimeoutMinutes)
$watchdog = Start-Job -ArgumentList @($excelBefore, $script:deadline.ToString('o')) -ScriptBlock {
    param($before, $deadlineText)
    $limit = [DateTime]::Parse($deadlineText, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
    while ((Get-Date) -lt $limit) { Start-Sleep -Seconds 5 }
    while ($true) {
        Get-CimInstance -ClassName Win32_Process -Filter "Name = 'EXCEL.EXE'" -ErrorAction SilentlyContinue |
            Where-Object { ($before -notcontains $_.ProcessId) -and ([string]$_.CommandLine -match '(?i)[/-]automation\b') } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        Start-Sleep -Seconds 5
    }
}

$development = Join-Path $repoRoot "dist\K-FRTB-ASA-development.xlsm"
$production = Join-Path $repoRoot "dist\K-FRTB-ASA-production.xlsm"

# Excel queries the default printer whenever a workbook sets its page setup;
# a slow printer shows "Waiting for printer connection" in every build. The
# run uses Microsoft Print to PDF, which answers at once, and puts the
# operator's default printer back in the finally block below (#304). The log
# names no printer of the operator.
$pdfPrinter = 'Microsoft Print to PDF'
$printerBefore = Get-DefaultPrinterName
$printerChanged = $false

try {
    if ($printerBefore -eq $pdfPrinter) {
        Write-Log "PRINTER=$pdfPrinter is already the default printer"
    } elseif (Set-DefaultPrinter $pdfPrinter) {
        $printerChanged = $true
        Write-Log "PRINTER=$pdfPrinter is the default printer for this run"
    } else {
        Write-Log "PRINTER=default printer unchanged: $pdfPrinter is not installed"
    }

    # -------------------------------------------------------- 1. recorded run
    Assert-TimeLeft
    Write-Host '== Recorded run (builds, compiles, regression, round trips)'
    $recordArgs = @('-File', 'tools\Record-FRTB-ExcelRun.ps1')
    if ($trustChange) { $recordArgs += @('-TrustChange', $trustChange) }
    $code = Invoke-Tool 'record.log' $recordArgs
    $recordBundle = Join-Path $repoRoot "test-results\excel-run-$short"
    if ($code -eq 0) {
        Add-Result 'RECORDED_RUN' 'PASS' "bundle test-results\excel-run-$short validated"
    } else {
        Add-Result 'RECORDED_RUN' 'FAIL' "exit code $code; see record.log and $recordBundle"
    }
    if (-not (Test-Path -LiteralPath $development) -or -not (Test-Path -LiteralPath $production)) {
        throw 'the recorded run did not leave both workbooks in dist\'
    }

    # ------------------------------------------------------ 2. acceptance
    Assert-TimeLeft
    Write-Host '== Acceptance checks in Excel'
    $copy = Join-Path $script:bundle 'acceptance.xlsm'
    Copy-Item -LiteralPath $development -Destination $copy -Force
    $excel = $null; $workbook = $null
    try {
        $excel = New-Excel $msoAutomationSecurityLow
        $workbook = $excel.Workbooks.Open($copy)
        [void](Invoke-Macro $excel $workbook 'TEST_Acceptance.Test_RunAcceptanceInExcel_Silent')
        $ws = $workbook.Worksheets.Item('Acceptance')
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add([string]$ws.Range('C2').Value2)
        for ($r = 5; $r -le 40; $r++) {
            $name = [string]$ws.Cells.Item($r, 3).Value2
            if (-not $name) { break }
            $result = [string]$ws.Cells.Item($r, 4).Value2
            $detail = [string]$ws.Cells.Item($r, 5).Value2
            $lines.Add("$name=$result | $detail")
            $summaryResult = if ($result -eq 'SKIPPED') { 'FAIL' } else { $result }
            Add-Result "ACCEPTANCE_$name" $summaryResult $detail
        }
        $lines | Set-Content -Path (Join-Path $script:bundle 'acceptance.txt') -Encoding UTF8
        [void](Export-SheetPicture $workbook 'Acceptance' 'B2:E14' (Join-Path $script:bundle 'acceptance.png'))
    } catch {
        Add-Result 'ACCEPTANCE' 'FAIL' "stopped: $($_.Exception.Message)"
    } finally {
        Close-Excel $excel $workbook
    }

    # Pictures of the states #85 and #203 are about, for the record
    Assert-TimeLeft
    Write-Host '== Pictures'
    $copy = Join-Path $script:bundle 'pictures.xlsm'
    Copy-Item -LiteralPath $development -Destination $copy -Force
    $excel = $null; $workbook = $null
    try {
        $excel = New-Excel $msoAutomationSecurityLow
        $workbook = $excel.Workbooks.Open($copy)
        [void](Invoke-Macro $excel $workbook 'M_Main.RunFRTB_Silent')
        [void](Export-SheetPicture $workbook 'Results' 'A1:H30' (Join-Path $script:bundle 'results.png'))
        [void](Export-SheetPicture $workbook 'Control' 'A1:G16' (Join-Path $script:bundle 'control.png'))
        $girr = $workbook.Worksheets.Item('P_GIRR')
        $tagRow = Find-TagRow $girr 'GIRR_RW'
        $cell = $girr.Cells.Item($tagRow + 2, 4)
        $saved = $cell.Value2
        [void]$cell.ClearContents()
        [void](Invoke-Macro $excel $workbook 'M_Main.RunFRTB_Silent')
        [void](Export-SheetPicture $workbook 'Results' 'A1:H12' (Join-Path $script:bundle 'failed-results.png'))
        [void](Export-SheetPicture $workbook 'Out_Summary' 'A1:F10' (Join-Path $script:bundle 'failed-out-summary.png'))
        [void](Export-SheetPicture $workbook 'Log' 'A1:F12' (Join-Path $script:bundle 'failed-log.png'))
        $cell.Value2 = $saved
        [void](Invoke-Macro $excel $workbook 'M_Main.ClearOutputs')
        [void](Export-SheetPicture $workbook 'Results' 'A1:H12' (Join-Path $script:bundle 'cleared-results.png'))
        Add-Result 'PICTURES' 'PASS' 'results, control, failed-results, failed-out-summary, failed-log, cleared-results'
    } catch {
        Add-Result 'PICTURES' 'FAIL' "stopped: $($_.Exception.Message)"
    } finally {
        Close-Excel $excel $workbook
    }

    # ------------------------------------------------------------ 3. builds
    Assert-TimeLeft
    Write-Host '== Build from source with the regression test'
    $code = Invoke-Tool 'build-from-source.log' @('-File', 'tools\Build-Workbook.ps1', '-FromSource', '-RunTests',
                                                 '-OutputPath', (Join-Path $script:bundle 'from-source.xlsm'), '-Force')
    Add-Result 'BUILD_FROM_SOURCE' $(if ($code -eq 0) { 'PASS' } else { 'FAIL' }) "exit code $code"

    Assert-TimeLeft
    Write-Host '== Template built by Excel'
    $excelTemplate = Join-Path $script:bundle 'FRTB_Template-excel.xlsx'
    $code = Invoke-Tool 'build-template.log' @('-File', 'tools\Build-Workbook.ps1', '-Template',
                                              '-OutputPath', $excelTemplate, '-Force')
    if ($code -eq 0 -and (Test-Path -LiteralPath $excelTemplate)) {
        $committed = Join-Path $repoRoot 'src\workbook\FRTB_Template.xlsx'
        $hashes = "excel=$((Get-FileHash -Algorithm SHA256 $excelTemplate).Hash.ToLowerInvariant()) " +
                  "committed=$((Get-FileHash -Algorithm SHA256 $committed).Hash.ToLowerInvariant())"
        $ErrorActionPreference = 'Continue'
        & python tools\compare_templates.py $committed $excelTemplate 2>&1 |
            ForEach-Object { [string]$_ } | Set-Content -Path (Join-Path $script:bundle 'template-differences.txt') -Encoding UTF8
        $ErrorActionPreference = 'Stop'
        Add-Result 'BUILD_TEMPLATE' 'PASS' "$hashes; differences in template-differences.txt"
    } else {
        Add-Result 'BUILD_TEMPLATE' 'FAIL' "exit code $code"
    }

    # ------------------------------------------------- 4. Trust Center path
    if ($access.Policy) {
        Add-Result 'TRUST_CENTER_ERROR' 'NOT_RUN' `
                   'a group policy sets the VBA project access; the error path cannot be tested without changing the policy'
    } elseif ($runTrustCheck) {
        Assert-TimeLeft
        Write-Host '== Trust Center error path'
        $before = Get-UserVbaAccess $excelVersion
        try {
            Set-VbaAccess $excelVersion 0
            $code = Invoke-Tool 'build-trust-center.log' @('-File', 'tools\Build-Workbook.ps1',
                                                          '-OutputPath', (Join-Path $script:bundle 'trust-test.xlsm'), '-Force')
        } finally {
            Restore-UserVbaAccess $excelVersion $before
        }
        $restored = if ($null -eq $before) { 'absent' } else { $before }
        $named = Select-String -Path (Join-Path $script:bundle 'build-trust-center.log') `
                               -Pattern 'Trust access to the VBA project object model' -SimpleMatch -Quiet
        $ok = ($code -ne 0) -and $named
        Add-Result 'TRUST_CENTER_ERROR' $(if ($ok) { 'PASS' } else { 'FAIL' }) `
                   "exit code $code; message names the setting: $named; your setting restored to $restored"
    } else {
        Add-Result 'TRUST_CENTER_ERROR' 'NOT_RUN' 'not requested'
    }

    # ------------------------------------------------- 5. synchronization
    $syncDir = Join-Path $script:bundle 'sync'
    [void](New-Item -ItemType Directory -Path $syncDir -Force)
    foreach ($mode in @('Production', 'Development')) {
        Assert-TimeLeft
        Write-Host "== Synchronization: $mode"
        $source = if ($mode -eq 'Production') { $production } else { $development }
        $copy = Join-Path $syncDir "sync-$($mode.ToLowerInvariant()).xlsm"
        Copy-Item -LiteralPath $source -Destination $copy -Force
        $excel = $null; $workbook = $null
        try {
            $excel = New-Excel $msoAutomationSecurityForceDisable
            $workbook = $excel.Workbooks.Open($copy, 0, $true)
            $calculationBefore = [int]$excel.Calculation
        } finally { Close-Excel $excel $workbook }
        $code = Invoke-Tool "sync-$($mode.ToLowerInvariant()).log" @('-File', 'tools\Sync-FRTB-VBA.ps1',
                            '-WorkbookPath', $copy, '-RepoRoot', $repoRoot, '-Mode', $mode)
        $excel = $null; $workbook = $null
        try {
            $excel = New-Excel $msoAutomationSecurityLow
            $workbook = $excel.Workbooks.Open($copy)
            $calculationAfter = [int]$excel.Calculation
            $compiled = Test-ProjectCompiles $excel
            [void](Invoke-Macro $excel $workbook 'M_Main.RunFRTB_Silent')
            $total = [double](Invoke-Macro $excel $workbook 'M_Main.OutSummaryValue' @('TOTAL'))
            $expected = Get-ExpectedTotal $workbook
            $detail = "exit code $code; compiled: $compiled; total $([math]::Round($total, 2)) (expected $([math]::Round($expected, 2))); calculation mode $calculationBefore -> $calculationAfter"
            $ok = ($code -eq 0) -and $compiled -and ([math]::Abs($total - $expected) -le $amountTolerance) -and ($calculationBefore -eq $calculationAfter)
            if ($mode -eq 'Development') {
                [void](Invoke-Macro $excel $workbook 'TEST_Regression.Test_RunRegression_Silent')
                $regression = [string]$workbook.Worksheets.Item('Test_Expected').Range('C3').Value2
                $detail += "; $regression"
                $ok = $ok -and ($regression -like 'Result: PASS*')
            }
            Add-Result "SYNC_$($mode.ToUpperInvariant())" $(if ($ok) { 'PASS' } else { 'FAIL' }) $detail
        } catch {
            Add-Result "SYNC_$($mode.ToUpperInvariant())" 'FAIL' "exit code $code; check stopped: $($_.Exception.Message)"
        } finally { Close-Excel $excel $workbook }
    }

    Assert-TimeLeft
    Write-Host '== Synchronization from a folder'
    $folder = Join-Path $syncDir 'downloads'
    [void](New-Item -ItemType Directory -Path $folder -Force)
    Copy-Item -LiteralPath (Join-Path $repoRoot 'src\modules\CORE_Config.bas') -Destination $folder
    $copy = Join-Path $syncDir 'sync-development.xlsm'
    $code = Invoke-Tool 'sync-folder.log' @('-File', 'tools\Sync-FRTB-VBA.ps1', '-WorkbookPath', $copy,
                                           '-RepoRoot', $repoRoot, '-SourceDir', $folder, '-Force')
    $replaced = Select-String -Path (Join-Path $script:bundle 'sync-folder.log') -Pattern 'CORE_Config\s+replaced' -Quiet
    Add-Result 'SYNC_FOLDER' $(if ($code -eq 0 -and $replaced) { 'PASS' } else { 'FAIL' }) "exit code $code; CORE_Config replaced: $replaced"

    Assert-TimeLeft
    Write-Host '== Synchronization failure path (workbook open in another Excel)'
    $locker = $null; $locked = $null
    try {
        $locker = New-Excel $msoAutomationSecurityForceDisable
        $locked = $locker.Workbooks.Open($copy)
        $stamp = (Get-Item -LiteralPath $copy).LastWriteTimeUtc
        $code = Invoke-Tool 'sync-failure.log' @('-File', 'tools\Sync-FRTB-VBA.ps1', '-WorkbookPath', $copy,
                                                '-RepoRoot', $repoRoot, '-Mode', 'Development')
    } finally { Close-Excel $locker $locked }
    $safety = @(Get-ChildItem -LiteralPath $syncDir -Filter 'sync-development.pre-vba-sync.*.xlsm' -ErrorAction SilentlyContinue)
    $unchanged = ((Get-Item -LiteralPath $copy).LastWriteTimeUtc -eq $stamp)
    $ok = ($code -eq 1) -and ($safety.Count -ge 1) -and $unchanged
    Add-Result 'SYNC_FAILURE' $(if ($ok) { 'PASS' } else { 'FAIL' }) "exit code $code; safety copy kept: $($safety.Count -ge 1); workbook unchanged: $unchanged"
    $safety | Remove-Item -Force -ErrorAction SilentlyContinue

    Assert-TimeLeft
    Write-Host '== Synchronization launchers'
    # Each launcher prints its own success line
    foreach ($launcher in @(@{ Name = 'Sync-FRTB-VBA.cmd'; Answers = "D`r`n`r`n"; Success = 'SUCCESS: VBA synchronized.' },
                            @{ Name = 'Sync-FRTB-VBA-Downloads.cmd'; Answers = "$folder`r`nY`r`n`r`n";
                               Success = 'SUCCESS: modules imported from the folder.' })) {
        $answers = Join-Path $syncDir ($launcher.Name + '.answers.txt')
        [System.IO.File]::WriteAllText($answers, $launcher.Answers)
        $logFile = Join-Path $script:bundle ($launcher.Name + '.log')
        # cmd /c drops the outer quotes of its command line, so the launcher
        # and the workbook keep theirs
        $command = "`"`"$(Join-Path $repoRoot "tools\$($launcher.Name)")`" `"$copy`"`""
        $excelBeforeLauncher = Get-AutomationExcelIds
        $process = Start-Process -FilePath 'cmd.exe' -PassThru -NoNewWindow -ArgumentList @('/c', $command) `
                                 -RedirectStandardInput $answers -RedirectStandardOutput $logFile
        # A Process from Start-Process keeps its exit code only when its
        # handle is read while it runs; without it ExitCode is empty (#297)
        $null = $process.Handle
        if (-not $process.WaitForExit(180000)) {
            Stop-TimedOutLauncher $process $excelBeforeLauncher
            Add-Result "LAUNCHER_$($launcher.Name)" 'FAIL' 'no answer after 3 minutes; its processes were stopped'
            continue
        }
        $process.WaitForExit()      # the redirected output is complete
        $succeeded = Select-String -Path $logFile -Pattern $launcher.Success -SimpleMatch -Quiet
        Add-Result "LAUNCHER_$($launcher.Name)" $(if ($process.ExitCode -eq 0 -and $succeeded) { 'PASS' } else { 'FAIL' }) `
                   "exit code $($process.ExitCode)"
    }
} catch {
    Add-Result 'RUN' 'FAIL' "stopped: $($_.Exception.Message)"
} finally {
    if ($printerChanged) {
        if ($printerBefore -and (Set-DefaultPrinter $printerBefore)) {
            Write-Log 'PRINTER=your default printer is restored'
        } else {
            Write-Log 'PRINTER=your default printer could not be restored; set it again in Windows Settings'
        }
    }
    Stop-Job $watchdog -ErrorAction SilentlyContinue
    Remove-Job $watchdog -Force -ErrorAction SilentlyContinue
}

# --------------------------------------------------------------- summary
Write-Log "FINISH=$(Get-Timestamp)"
$failed = @($script:results | Where-Object { $_.Result -eq 'FAIL' })
$verdict = if ($failed.Count -eq 0) { 'RESULT=PASS' } else { "RESULT=FAIL ($($failed.Count) failed)" }
$summary = New-Object System.Collections.Generic.List[string]
$summary.Add("K-FRTB-ASA Excel check, $Branch at $sha")
$summary.Add("Excel $excelVersion; $(Get-Timestamp)")
foreach ($item in $script:results) { $summary.Add(('{0,-28} {1,-8} {2}' -f $item.Name, $item.Result, $item.Detail)) }
$summary.Add($verdict)
$summaryPath = Join-Path $script:bundle 'summary.txt'
$summary | Set-Content -Path $summaryPath -Encoding UTF8
Get-ChildItem -LiteralPath $script:bundle -Filter '*.xlsm' -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
$zip = "$($script:bundle).zip"
Compress-Archive -Path (Join-Path $script:bundle '*') -DestinationPath $zip -Force
try { $summary -join "`r`n" | Set-Clipboard } catch { }
Write-Host ''
$summary | ForEach-Object { Write-Host $_ }
Write-Host ''
Write-Host "Bundle: $zip (the summary is on the clipboard)"
Write-Host "Recorded evidence: test-results\excel-run-$short and test-results\excel-run-$short-validation"
try { Start-Process explorer.exe -ArgumentList "/select,`"$zip`"" } catch { }
if ($failed.Count -eq 0) { exit 0 } else { exit 1 }
