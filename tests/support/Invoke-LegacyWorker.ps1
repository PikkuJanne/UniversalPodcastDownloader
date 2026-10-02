param([Parameter(Mandatory)][string]$ConfigPath)

# Exercise the public entry point in a fresh, noninteractive process. Configuration
# and results live outside the copied archive being checked for zero mutations.
$ErrorActionPreference = 'Stop'
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
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
    $result.Output = @(& $config.ProductScript @parameters)
    $result.Succeeded = $true
}
catch {
    $result.ErrorMessage = $_.Exception.Message
    $result.ErrorId = $_.FullyQualifiedErrorId
}
$result | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $config.ResultPath -Encoding UTF8
if ($result.Succeeded) { exit 0 }
exit 1
