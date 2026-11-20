<#
.SYNOPSIS
    Runs the unattended Excel check for the self-hosted runner and writes its summary.

.DESCRIPTION
    .github/workflows/excel-check.yml starts this script with
    powershell.exe -ExecutionPolicy Bypass, because the temporary script that the
    runner writes for a PowerShell step is refused when the host's execution
    policy is Restricted. The script reads the profile and the branch from the
    environment variables PROFILE and BRANCH, and the commit from GITHUB_SHA: the
    branch name is untrusted, so it never reaches a command line as text. It runs
    tools\Check-FRTB-Excel.ps1 with -NoPrompt (the quick profile unless PROFILE is
    full), copies the summary of the bundle into the job summary, and exits with
    the exit code of the check. The bundle stays under test-results on the host.
#>
$ErrorActionPreference = 'Stop'

$extra = @()
if ($env:PROFILE -ne 'full') { $extra += '-Quick' }

$script = Join-Path $PSScriptRoot 'Check-FRTB-Excel.ps1'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script -Branch $env:BRANCH -Commit $env:GITHUB_SHA -NoPrompt @extra
$code = $LASTEXITCODE

$short = $env:GITHUB_SHA.Substring(0, 7)
$results = Join-Path (Split-Path -Parent $PSScriptRoot) 'test-results'
$folder = Get-ChildItem -Path $results -Directory -Filter "excel-check-$short*" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($env:GITHUB_STEP_SUMMARY) {
    if ($folder -and (Test-Path -LiteralPath (Join-Path $folder.FullName 'summary.txt'))) {
        Add-Content -Path $env:GITHUB_STEP_SUMMARY -Value '```text'
        Get-Content -LiteralPath (Join-Path $folder.FullName 'summary.txt') | Add-Content -Path $env:GITHUB_STEP_SUMMARY
        Add-Content -Path $env:GITHUB_STEP_SUMMARY -Value '```'
    } else {
        Add-Content -Path $env:GITHUB_STEP_SUMMARY -Value 'The check wrote no summary: see the log of the step.'
    }
}
exit $code
