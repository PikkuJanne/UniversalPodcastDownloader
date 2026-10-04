param([Parameter(Mandatory)][string]$ConfigPath)

# Exercise the public entry point in a fresh, noninteractive process. Configuration
# and results live outside the copied archive being checked for zero mutations.
$ErrorActionPreference = 'Stop'
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
. (Join-Path $PSScriptRoot 'WorkerRunProjection.ps1')
$workerExit = 1
$result = [ordered]@{
    Succeeded = $false
    EngineVersion = $PSVersionTable.PSVersion.ToString()
    ErrorMessage = $null
    ErrorId = $null
    Output = @()
}
try {
    $parameters = @{
        Mode = 'All'
        FeedUrl = $config.FeedUrl
        OutputPath = $config.OutputPath
    }
    if ($config.LegacyPath) { $parameters.LegacyPath = $config.LegacyPath }
    if ($config.LegacyAction) { $parameters.LegacyAction = $config.LegacyAction }
    if ($config.LegacyEpisodeId) { $parameters.LegacyEpisodeId = $config.LegacyEpisodeId }
    if ($config.LegacyFile) { $parameters.LegacyFile = $config.LegacyFile }
    if ($config.LegacySha256) { $parameters.LegacySha256 = $config.LegacySha256 }
    if ($config.LegacyCheckpoint) { $parameters.LegacyCheckpoint = $config.LegacyCheckpoint }
    if ($config.PreviewOnly) { $parameters.WhatIf = $true }
    $projection = Get-UpdWorkerRunProjection -Output @(& $config.ProductScript @parameters -PassThru)
    $result.Output = @(if ($projection.RunResult.PSObject.Properties['LegacyResult']) { $projection.RunResult.LegacyResult } else { $projection.RunResult.Plan })
    $result.Succeeded = $projection.Succeeded
    $result.ErrorMessage = $projection.ErrorMessage
    $result.RunResult = $projection.RunResult
    $workerExit = $projection.ExitCode
}
catch {
    $result.ErrorMessage = $_.Exception.Message
    $result.ErrorId = $_.FullyQualifiedErrorId
}
$result | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $config.ResultPath -Encoding UTF8
exit $workerExit
