param([Parameter(Mandatory)][string]$ConfigPath)

$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
. (Join-Path $PSScriptRoot 'WorkerRunProjection.ps1')
$workerExit = 1
# Compile test infrastructure before redirecting the application's log roots.
if ($config.Action -eq 'Decline') {
    Add-Type -Path (Join-Path $PSScriptRoot 'DeclineConfirmationHost.cs')
}
$env:LOCALAPPDATA = Join-Path $config.Root 'local'
$env:TEMP = Join-Path $config.Root 'temp'
$env:TMP = $env:TEMP
$result = [ordered]@{
    Succeeded = $false; ErrorMessage = $null; Output = @()
    ConfirmationCount = 0; ItemCount = 0; Bytes = 0
}
$guard = $null
$runspace = $null
$pipeline = $null
try {
    if ($config.SentinelPath) {
        $guard = [IO.File]::Open($config.SentinelPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    }
    if ($config.Action -in @('Resolve', 'Metadata')) {
        . $config.ProductScript
        if ($config.Action -eq 'Metadata') {
            $response = Invoke-PodcastMetadataRequest -Uri $config.FeedUrl
            $result.Bytes = $response.Bytes
        }
        else {
            $resolved = Resolve-PodcastItems -Feeds @($config.FeedUrl)
            $result.ItemCount = @($resolved.Items).Count
        }
    }
    else {
        $parameters = @{
            FeedUrl = $config.FeedUrl; OutputPath = Join-Path $config.Root 'out'
            Mode = 'All'; DiagnosticExportPath = Join-Path $config.ExportRoot 'diagnostics.json'
            PassThru = $true
        }
        if ($config.Action -eq 'Preview') { $parameters.WhatIf = $true }
        if ($config.Action -eq 'Decline') {
            $parameters.Confirm = $true
            $testHost = New-Object UpdTests.DeclineConfirmationHost
            $runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($testHost)
            $runspace.Open()
            $pipeline = [Management.Automation.PowerShell]::Create()
            $pipeline.Runspace = $runspace
            $null = $pipeline.AddCommand($config.ProductScript).AddParameters($parameters)
            $published = @($pipeline.Invoke())
            $result.ConfirmationCount = $testHost.ConfirmationCount
            if ($pipeline.HadErrors) { throw $pipeline.Streams.Error[0] }
        }
        else { $published = @(& $config.ProductScript @parameters) }
        $projection = Get-UpdWorkerRunProjection -Output $published
        $result.Output = @($projection.RunResult.Plan)
        $result.Succeeded = $projection.Succeeded
        $result.ErrorMessage = $projection.ErrorMessage
        $result.RunResult = $projection.RunResult
        $workerExit = $projection.ExitCode
    }
    if ($config.Action -in @('Resolve', 'Metadata')) { $result.Succeeded = $true; $workerExit = 0 }
}
catch { $result.ErrorMessage = $_.Exception.Message }
finally {
    if ($null -ne $guard) { $guard.Dispose() }
    if ($null -ne $pipeline) { $pipeline.Dispose() }
    if ($null -ne $runspace) { $runspace.Dispose() }
}
$result.RootEntries = @(if ([IO.Directory]::Exists($config.Root)) {
    Get-ChildItem -LiteralPath $config.Root -Recurse -Force | ForEach-Object { $_.FullName.Substring($config.Root.Length) }
})
[IO.File]::WriteAllText($config.ResultPath, ($result | ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
exit $workerExit
