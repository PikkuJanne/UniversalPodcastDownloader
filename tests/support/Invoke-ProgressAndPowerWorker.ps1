param([Parameter(Mandatory)][string]$ConfigPath)

# Owned instrumentation observes the real HTTP, validation, history and native
# power helpers. Forced progress detection exercises rendering calls, not a TTY.
$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
$root = [IO.Path]::GetFullPath($config.Root)
if (-not [IO.File]::Exists((Join-Path $root '.upd-test-owner')) -or
    [IO.File]::ReadAllText((Join-Path $root '.upd-test-owner')) -cne $config.Token) {
    throw 'Progress and power worker requires its matching owned temporary root.'
}
$env:LOCALAPPDATA = Join-Path $root 'local'
$nativeObservation = $null -ne $config.PSObject.Properties['NativeHostObservation'] -and [bool]$config.NativeHostObservation
$forceInteractive = -not $nativeObservation -and [bool]$config.ForceInteractive
. (Join-Path $PSScriptRoot 'WorkerRunProjection.ps1')
$observation = @{ Context = $null; Events = New-Object 'System.Collections.Generic.List[object]';
    Renders = New-Object 'System.Collections.Generic.List[object]'; Power = New-Object 'System.Collections.Generic.List[object]' }
$report = [ordered]@{ HostSurvived = $false; Error = $null; RunResult = $null; Events = @(); Renders = @(); Power = @();
    PartialClosed = $false; LockClosed = $false; CancelBytes = 0; Host = $null; InitialContextEnabled = $null }
$workerExit = 1

function Get-UpdProgressSnapshot {
    param([string]$Stage)
    $accepted = 0
    foreach ($path in @(Get-ChildItem -LiteralPath $config.OutputPath -Recurse -Force -File -Filter 'state.json' -ErrorAction SilentlyContinue)) {
        $state = [IO.File]::ReadAllText($path.FullName) | ConvertFrom-Json
        $accepted += @($state.episodes | Where-Object { $_.status -eq 'transfer_verified' -and $null -ne $_.local_sha256 }).Count
    }
    $context = $observation.Context
    [pscustomobject]@{ Observation = $Stage; Enabled = $context.Enabled; RootId = $context.RootId; TransferId = $context.TransferId;
        TotalEpisodes = $context.TotalEpisodes; EpisodeIndex = $context.EpisodeIndex; ProcessedEpisodes = $context.ProcessedEpisodes;
        VerifiedEpisodes = $context.VerifiedEpisodes; Attempt = $context.Attempt; Offset = $context.Offset; Bytes = $context.Bytes;
        TotalBytes = $context.TotalBytes; Stage = $context.Stage; EpisodeOutcome = $context.EpisodeOutcome;
        RunVerified = $context.RunVerified; EpisodePercent = $context.EpisodePercent; TransferPercent = $context.TransferPercent;
        Closed = $context.Closed; AcceptedHistory = $accepted }
}

function Add-UpdProgressObservation {
    param([string]$Stage)
    if ($observation.Events.Count -ge 256) { throw 'Owned progress observation limit exceeded.' }
    $observation.Events.Add((Get-UpdProgressSnapshot -Stage $Stage))
}

try {
    . $config.ProductScript
    $report.Host = [pscustomobject]@{ Name = $Host.Name; OutputRedirected = [Console]::IsOutputRedirected;
        ErrorRedirected = [Console]::IsErrorRedirected; ActualProgressDetected = Test-PodcastProgressInteractive;
        NativeHostObservation = $nativeObservation; ForcedDetector = $forceInteractive }
    if ($forceInteractive) {
        function Test-PodcastProgressInteractive { return $true }
    }
    $originalContext = (Get-Command New-PodcastProgressContext).ScriptBlock
    function New-PodcastProgressContext {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates and observes only the private in-memory progress context in this owned worker.')]
        [CmdletBinding()]
        param([int]$TotalEpisodes, [switch]$NonInteractive)
        $context = & $originalContext @PSBoundParameters
        $observation.Context = $context
        $report.InitialContextEnabled = $context.Enabled
        return $context
    }
    $nativeProgress = Get-Command Write-Progress -CommandType Cmdlet
    function Write-Progress {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Owned worker observes and forwards the actual renderer to prove history/progress ordering; no product override is installed.')]
        [CmdletBinding()]
        param([string]$Activity, [string]$Status, [int]$Id = 0, [int]$ParentId = -1,
            [string]$CurrentOperation, [int]$PercentComplete = -1, [int]$SecondsRemaining = -1, [switch]$Completed, [int]$SourceId = 0)
        if ($observation.Renders.Count -ge 256) { throw 'Owned renderer observation limit exceeded.' }
        $snapshot = Get-UpdProgressSnapshot -Stage 'render'
        $role = if ($Id -eq $snapshot.RootId) { 'root' } else { 'transfer' }
        $observation.Renders.Add([pscustomobject]@{ Role = $role; Id = $Id; ParentId = $ParentId; Activity = $Activity;
            Status = $Status; CurrentOperation = $CurrentOperation; Percent = $PercentComplete; SecondsRemaining = $SecondsRemaining;
            Completed = [bool]$Completed; SourceId = $SourceId; Snapshot = $snapshot })
        & $nativeProgress @PSBoundParameters
    }
    $originalRequest = (Get-Command Invoke-PodcastMediaRequest).ScriptBlock
    $originalRead = (Get-Command Read-PodcastResponseChunk).ScriptBlock
    function Read-PodcastResponseChunk {
        [CmdletBinding()]
        param($Source, [byte[]]$Buffer, [int]$Count, $Policy)
        & $originalRead -Source $Source -Buffer $Buffer -Count ([Math]::Min($Count, 4096)) -Policy $Policy
    }
    function Invoke-PodcastMediaRequest {
        [CmdletBinding()]
        param([string]$Uri, [IO.Stream]$DestinationStream, $Policy, $Resume, [scriptblock]$OnResponse, [scriptblock]$OnProgress)
        $forward = @{} + $PSBoundParameters
        $responseCallback = $OnResponse
        $progressCallback = $OnProgress
        $forward.OnResponse = {
            param($Info)
            if ($responseCallback) { $null = & $responseCallback $Info }
            Add-UpdProgressObservation -Stage 'response'
        }.GetNewClosure()
        $forward.OnProgress = {
            param([long]$Bytes)
            if ($progressCallback) { $null = & $progressCallback $Bytes }
            Add-UpdProgressObservation -Stage 'bytes'
            if ([long]$config.CancelAfterBytes -gt 0 -and $Bytes -ge [long]$config.CancelAfterBytes) {
                $report.CancelBytes = $Bytes
                throw [OperationCanceledException]::new('Synthetic progress cancellation sentinel.')
            }
        }.GetNewClosure()
        & $originalRequest @forward
    }
    $originalValidation = (Get-Command Test-PodcastMediaFile).ScriptBlock
    function Test-PodcastMediaFile {
        [CmdletBinding()]
        param([string]$LiteralPath, [bool]$TransferCompleted = $false, [Nullable[long]]$HttpContentLength,
            [bool]$ContentLengthAppliesToStoredBytes = $true, [string]$ContentType, [string]$EnclosureContentType,
            [string]$MediaUrl, [Nullable[long]]$EnclosureLength)
        Add-UpdProgressObservation -Stage 'validation'
        & $originalValidation @PSBoundParameters
    }
    if ($config.FailHistoryCompletion) {
        $originalSave = (Get-Command Save-PodcastEpisodeRecord).ScriptBlock
        function Save-PodcastEpisodeRecord {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Forwards real transactional history in this owned root, injecting only the explicit completed-history failure fixture.')]
            [CmdletBinding()]
            param($Context, $Record)
            if ($Record.status -eq 'transfer_verified') { throw [IO.IOException]::new('Synthetic completed history failure sentinel.') }
            & $originalSave @PSBoundParameters
        }
    }
    $originalStart = (Get-Command Start-PodcastKeepAwake).ScriptBlock
    $originalStop = (Get-Command Stop-PodcastKeepAwake).ScriptBlock
    function Start-PodcastKeepAwake {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Owned worker forwards the real temporary native helper and observes its lease; the product owns finally restoration.')]
        [CmdletBinding()]
        param([switch]$Enabled)
        $context = & $originalStart @PSBoundParameters
        $observation.Power.Add([pscustomobject]@{ Context = $context; ActiveBeforeStop = $context.Active; Restored = $false })
        return $context
    }
    function Stop-PodcastKeepAwake {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Owned worker forwards actual native restoration without changing power policy or extending the product lease.')]
        [CmdletBinding()]
        param($Context)
        $status = & $originalStop @PSBoundParameters
        foreach ($item in $observation.Power) {
            if ([object]::ReferenceEquals($item.Context, $Context)) { $item.Restored = $status.Restored }
        }
        return $status
    }
    $parameters = @{ FeedUrl = $config.FeedUrl; OutputPath = $config.OutputPath; Mode = 'All';
        NonInteractive = -not ($forceInteractive -or $nativeObservation); MaxAttempts = [int]$config.MaxAttempts;
        BaseDelaySeconds = 0; MaxDelaySeconds = 0; RetryBudgetSeconds = 10 }
    if ($config.Preview) { $parameters.WhatIf = $true }
    if ($config.KeepAwake) { $parameters.KeepAwake = $true }
    $projection = Get-UpdWorkerRunProjection -Output @(Invoke-PodcastRun @parameters)
    $report.RunResult = $projection.RunResult
    $workerExit = $projection.ExitCode
    if ($null -ne $observation.Context) { Add-UpdProgressObservation -Stage 'returned' }
    foreach ($partial in @(Get-ChildItem -LiteralPath $config.OutputPath -Recurse -Force -File -Filter '.upd-*.tmp' -ErrorAction SilentlyContinue)) {
        $handle = [IO.File]::Open($partial.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $handle.Dispose()
        $report.PartialClosed = $true
    }
    foreach ($lock in @(Get-ChildItem -LiteralPath $config.OutputPath -Recurse -Force -File -Filter 'writer.lock' -ErrorAction SilentlyContinue)) {
        $handle = [IO.File]::Open($lock.FullName, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $handle.Dispose()
        $report.LockClosed = $true
    }
    $report.HostSurvived = $true
}
catch { $report.Error = $_.Exception.Message }
$report.Events = @($observation.Events.ToArray())
$report.Renders = @($observation.Renders.ToArray())
$report.Power = @($observation.Power | ForEach-Object {
    [pscustomobject]@{ Requested = $_.Context.Requested; ActiveBeforeStop = $_.ActiveBeforeStop; Restored = $_.Restored;
        PreviousState = $_.Context.Lease.PreviousState; ActivationThreadId = $_.Context.Lease.ActivationThreadId;
        RestorationThreadId = $_.Context.Lease.RestorationThreadId; NativeActive = $_.Context.Lease.Active;
        NativeRestored = $_.Context.Lease.Restored }
})
[IO.File]::WriteAllText($config.ResultPath, ($report | ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
exit $workerExit
