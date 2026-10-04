param([Parameter(Mandatory)][string]$ConfigPath)

# Test-only process boundary: all signals and destinations belong to one marked
# temporary root. Cancellation is typed and catchable, after a real body write.
$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
$root = [IO.Path]::GetFullPath($config.Root)
$marker = Join-Path $root '.upd-test-owner'
if (-not [IO.File]::Exists($marker) -or [IO.File]::ReadAllText($marker) -cne $config.Token) {
    throw 'Run safety worker requires its matching owned temporary root.'
}
$env:LOCALAPPDATA = Join-Path $root 'local'
. (Join-Path $PSScriptRoot 'WorkerRunProjection.ps1')
$report = [ordered]@{ HostSurvived = $false; Error = $null; RunResult = $null; CancelBytes = 0; PartialClosed = $false;
    SpaceCalls = 0; ProgressCalls = 0; ResponseLength = $null }
$workerExit = 1
try {
    . $config.ProductScript
    if ($config.Action -eq 'HoldArchive') {
        $held = Enter-PodcastArchiveLock -Root $config.OutputPath
        try {
            [IO.File]::WriteAllText($config.ReadyPath, 'archive-selection-held')
            $watch = [Diagnostics.Stopwatch]::StartNew()
            while (-not [IO.File]::Exists($config.ReleasePath)) {
                if ($watch.Elapsed.TotalSeconds -gt 45) { throw 'Owned archive-lock gate timed out.' }
                Start-Sleep -Milliseconds 25
            }
        }
        finally { $held.Dispose() }
        $workerExit = 0
    }
    else {
        $spaceState = @{ Calls = 0 }
        if ($config.PSObject.Properties['AvailableSpace']) {
            function Get-PodcastAvailableSpace {
                [CmdletBinding()]
                [OutputType([long])]
                param([string]$Root)
                $null = $Root
                $spaceState.Calls++
                if ($null -eq $config.AvailableSpace) { return $null }
                return [long]$config.AvailableSpace
            }
        }
        $transferState = @{ ProgressCalls = 0; ResponseLength = $null }
        $originalRequest = (Get-Command Invoke-PodcastMediaRequest).ScriptBlock
        $cancelState = @{ Bytes = 0 }
        if ($config.Action -eq 'GateCancel') {
            if ([int]$config.ChunkLimit -gt 0) {
                $originalRead = (Get-Command Read-PodcastResponseChunk).ScriptBlock
                function Read-PodcastResponseChunk {
                    [CmdletBinding()]
                    param($Source, [byte[]]$Buffer, [int]$Count, $Policy)
                    # Smaller real reads expose a cancellation between geometric
                    # checkpoints without replacing HTTP or inventing body bytes.
                    & $originalRead -Source $Source -Buffer $Buffer -Count ([Math]::Min($Count, [int]$config.ChunkLimit)) -Policy $Policy
                }
            }
        }
        function Invoke-PodcastMediaRequest {
            [CmdletBinding()]
            param([string]$Uri, [IO.Stream]$DestinationStream, $Policy, $Resume, [scriptblock]$OnResponse, [scriptblock]$OnProgress)
            $forward = @{} + $PSBoundParameters
            $originalResponse = $OnResponse
            $originalProgress = $OnProgress
            $forward.OnResponse = {
                param($Info)
                $transferState.ResponseLength = $Info.ResponseLength
                if ($originalResponse) { $null = & $originalResponse $Info }
            }.GetNewClosure()
            $forward.OnProgress = {
                param([long]$Bytes)
                $transferState.ProgressCalls++
                if ($originalProgress) { $null = & $originalProgress $Bytes }
                if ($config.Action -ne 'GateCancel') { return }
                $cancelState.Bytes = $Bytes
                if ($Bytes -lt [long]$config.CancelAfterBytes) { return }
                [IO.File]::WriteAllText($config.ReadyPath, ($Bytes.ToString([Globalization.CultureInfo]::InvariantCulture)))
                $watch = [Diagnostics.Stopwatch]::StartNew()
                while (-not [IO.File]::Exists($config.ReleasePath)) {
                    if ($watch.Elapsed.TotalSeconds -gt 45) { throw 'Owned cancellation gate timed out.' }
                    Start-Sleep -Milliseconds 25
                }
                throw [OperationCanceledException]::new('Synthetic run safety cancellation sentinel.')
            }.GetNewClosure()
            & $originalRequest @forward
        }
        $parameters = @{ FeedUrl = $config.FeedUrl; OutputPath = $config.OutputPath; Mode = 'All'; NonInteractive = $true;
            MaxAttempts = 1; BaseDelaySeconds = 0; MaxDelaySeconds = 0 }
        if ($config.Preview) { $parameters.WhatIf = $true }
        $projection = Get-UpdWorkerRunProjection -Output @(Invoke-PodcastRun @parameters)
        $report.RunResult = $projection.RunResult
        $report.SpaceCalls = $spaceState.Calls
        $report.ProgressCalls = $transferState.ProgressCalls
        $report.ResponseLength = $transferState.ResponseLength
        $workerExit = $projection.ExitCode
        if ($config.Action -eq 'GateCancel') {
            $report.CancelBytes = $cancelState.Bytes
            $partials = @(Get-ChildItem -LiteralPath $config.OutputPath -Recurse -Force -File -Filter '.upd-*.tmp')
            if ($partials.Count -eq 1) {
                $handle = [IO.File]::Open($partials[0].FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
                $handle.Dispose()
                $report.PartialClosed = $true
            }
        }
    }
    $report.HostSurvived = $true
}
catch { $report.Error = $_.Exception.Message }
[IO.File]::WriteAllText($config.ResultPath, ($report | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
exit $workerExit
