#requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Unit', 'Integration', 'All')][string]$Suite = 'All',
    [string]$Filter = '*',
    [string]$TestRoot,
    [string]$ToolsPath
)

$ErrorActionPreference = 'Stop'
try {
    $repo = Split-Path $PSScriptRoot -Parent
    if (-not $TestRoot) { $TestRoot = Join-Path $repo 'tests' }
    if (-not $ToolsPath) { $ToolsPath = Join-Path $repo '.dev-tools/Modules' }
    . (Join-Path $PSScriptRoot 'DevTools.ps1')
    Import-UpdDevModule -Name Pester -ToolsPath $ToolsPath
    $suites = if ($Suite -eq 'All') { @('Unit', 'Integration') } else { @($Suite) }
    $paths = @()
    $suitePaths = @{}
    foreach ($name in $suites) {
        $directory = Join-Path $TestRoot $name
        $files = @(Get-ChildItem -LiteralPath $directory -Filter '*.Tests.ps1' -File -Recurse -ErrorAction Stop)
        if ($files.Count -eq 0) { throw "No tests found for suite $name." }
        $suitePaths[$name] = @($files.FullName)
        $paths += $files.FullName
    }
    $configuration = New-PesterConfiguration
    $configuration.Run.Path = $paths
    $configuration.Run.PassThru = $true
    $configuration.Run.Exit = $false
    $configuration.Filter.FullName = $Filter
    $configuration.Output.Verbosity = 'Detailed'
    $result = Invoke-Pester -Configuration $configuration
    $summary = ("RESULT suite={0} engine={1} total={2} passed={3} failed={4} skipped={5} not_run={6}" -f
        $Suite, $PSVersionTable.PSVersion, $result.TotalCount, $result.PassedCount,
        $result.FailedCount, $result.SkippedCount, $result.NotRunCount)
    Write-Host $summary
    if ($env:GITHUB_STEP_SUMMARY) { Add-Content -LiteralPath $env:GITHUB_STEP_SUMMARY -Value $summary -Encoding UTF8 }
    if ($result.TotalCount -eq 0 -or ($result.PassedCount + $result.FailedCount) -eq 0) {
        throw 'No tests executed; an empty, excluded, or entirely skipped selection is not a pass.'
    }
    foreach ($name in $suites) {
        $containers = @($result.Containers | Where-Object { $_.Item.FullName -in $suitePaths[$name] })
        $discovered = ($containers | Measure-Object -Property TotalCount -Sum).Sum
        $executed = ($containers | Measure-Object -Property PassedCount -Sum).Sum +
            ($containers | Measure-Object -Property FailedCount -Sum).Sum
        if ($discovered -lt 1) { throw "No tests discovered in suite $name." }
        if ($Filter -eq '*' -and $executed -lt 1) { throw "No tests executed in suite $name." }
    }
    if ($result.Result -ne 'Passed' -or $result.FailedCount -gt 0) { exit 1 }
    exit 0
}
catch {
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
    exit 1
}
