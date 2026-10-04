param([Parameter(Mandatory)][string]$ConfigPath)

# Import and call in a disposable process with only synthetic owned destinations.
$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
. (Join-Path $PSScriptRoot 'WorkerRunProjection.ps1')
$env:LOCALAPPDATA = Join-Path $config.Root 'local'
$report = [ordered]@{ HostSurvived = $false; ImportOutputCount = 0; PromptCount = 0; RunResults = @(); Error = $null; CancelBytes = 0; PartialClosed = $false }
$workerExit = 1
try {
    $imported = @(. $config.ProductScript -FeedUrl $config.FeedUrl -OutputPath $config.OutputPath -Mode All)
    $report.ImportOutputCount = $imported.Count
    $promptState = @{ Count = 0 }
    function Read-Host {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Owned child prompt spy proves NonInteractive never requests input; it never returns synthetic user consent.')]
        [CmdletBinding()]
        param([string]$Prompt)
        $promptState.Count++
        $null = $Prompt
        throw 'An automated callable run unexpectedly requested input.'
    }
    if ($config.Action -eq 'Cancel') {
        $originalRequest = (Get-Command Invoke-PodcastMediaRequest).ScriptBlock
        $cancelState = @{ Bytes = 0 }
        function Invoke-PodcastMediaRequest {
            [CmdletBinding()]
            param([string]$Uri, [IO.Stream]$DestinationStream, $Policy, $Resume, [scriptblock]$OnResponse, [scriptblock]$OnProgress)
            $forward = @{} + $PSBoundParameters
            $originalProgress = $OnProgress
            $forward.OnProgress = {
                param([long]$Bytes)
                if ($originalProgress) { $null = & $originalProgress $Bytes }
                $cancelState.Bytes = $Bytes
                throw [OperationCanceledException]::new('Synthetic private cancellation sentinel.')
            }.GetNewClosure()
            & $originalRequest @forward
        }
    }
    $parameters = @{ OutputPath = $config.OutputPath; NonInteractive = $true; MaxAttempts = 3; BaseDelaySeconds = 0; MaxDelaySeconds = 0 }
    if (-not $config.WithoutFeed) { $parameters.FeedUrl = $config.FeedUrl }
    if ($config.PSObject.Properties['Mode']) { $parameters.Mode = $config.Mode }
    if ($config.PSObject.Properties['CustomCount']) { $parameters.CustomCount = [int]$config.CustomCount }
    if ($config.Preview) { $parameters.WhatIf = $true }
    for ($invocation = 0; $invocation -lt [int]$config.Repeat; $invocation++) {
        $projection = Get-UpdWorkerRunProjection -Output @(Invoke-PodcastRun @parameters)
        $report.RunResults += $projection.RunResult
        $workerExit = $projection.ExitCode
    }
    $report.PromptCount = $promptState.Count
    if ($config.Action -eq 'Cancel') {
        $report.CancelBytes = $cancelState.Bytes
        $partials = @(Get-ChildItem -LiteralPath $config.OutputPath -Recurse -Force -File -Filter '.upd-*.tmp')
        if ($partials.Count -eq 1) {
            $handle = [IO.File]::Open($partials[0].FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
            $handle.Dispose()
            $report.PartialClosed = $true
        }
    }
    $report.HostSurvived = $true
}
catch { $report.Error = $_.Exception.Message }
[IO.File]::WriteAllText($config.ResultPath, ($report | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
exit $workerExit
