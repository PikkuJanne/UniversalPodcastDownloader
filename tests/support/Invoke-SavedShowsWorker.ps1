param([Parameter(Mandatory)][string]$ConfigPath)

# Observes actual dispatcher, DPAPI, HTTP and byte callbacks in one owned child.
$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
$root = [IO.Path]::GetFullPath($config.Root)
if (-not [IO.File]::Exists((Join-Path $root '.upd-test-owner')) -or
    [IO.File]::ReadAllText((Join-Path $root '.upd-test-owner')) -cne $config.Token) {
    throw 'Saved-show worker requires its matching owned temporary root.'
}
$prefix = $root.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
foreach ($path in @($config.SavedConfigPath, $config.ResultPath)) {
    if (-not [IO.Path]::GetFullPath($path).StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Saved-show worker paths must remain in its owned root.'
    }
}
$env:LOCALAPPDATA = Join-Path $root 'local'
$env:USERPROFILE = Join-Path $root 'profile'
$options = @{ ConfigPath = [string]$config.SavedConfigPath; NonInteractive = $true }
foreach ($property in $config.Options.PSObject.Properties) { $options[$property.Name] = $property.Value }
foreach ($key in @('CustomCount', 'MaxAttempts', 'MaxFeedPages')) {
    if ($options.ContainsKey($key)) { $options[$key] = [int]$options[$key] }
}
if ($options.ContainsKey('ShowName')) { $options.ShowName = [string[]]@($options.ShowName) }
foreach ($key in @('OutputPath', 'ExportShows')) {
    if ($options.ContainsKey($key) -and -not [IO.Path]::GetFullPath($options[$key]).StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Saved-show output and export paths must remain in its owned root.'
    }
}
$observation = @{ Runs = New-Object 'System.Collections.Generic.List[object]'; Requests = New-Object 'System.Collections.Generic.List[object]';
    Active = 0; Peak = 0; Actor = $null; Names = @{}; Protect = 0; Locks = 0; Power = 0 }
$report = [ordered]@{ Result = $null; Error = $null; HostSurvived = $false; RunTrace = @(); RequestTrace = @(); PeakActiveRuns = 0;
    ProtectCalls = 0; ConfigLockCalls = 0; PowerRequested = 0; CancelBytes = 0; PartialClosed = $false; LockClosed = $false }
$workerExit = 1

function Add-UpdSavedRequest {
    param([string]$Kind, [string]$Uri)
    $target = [uri]$Uri
    if ($target.Host -ne '127.0.0.1' -or $target.Scheme -ne 'http') { throw 'Only owned loopback requests are allowed.' }
    if ($observation.Requests.Count -ge 128) { throw 'Owned request observation limit exceeded.' }
    $observation.Requests.Add([pscustomobject]@{ Kind = $Kind; Show = $observation.Actor; Path = $target.AbsolutePath })
}

try {
    . $config.ProductScript
    if ([IO.File]::Exists($config.SavedConfigPath)) {
        try {
            $stored = Read-PodcastSavedShowConfig -ConfigPath $config.SavedConfigPath
            foreach ($show in $stored.shows) {
                if (-not [IO.Path]::GetFullPath($show.output_path).StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
                    throw 'Stored outputs must remain in the owned root.'
                }
                $observation.Names[$show.output_path] = $show.name
            }
        }
        catch { if ($_.Exception.Message -eq 'Stored outputs must remain in the owned root.') { throw } }
    }
    $originalRun = (Get-Command Invoke-PodcastRun).ScriptBlock
    function Invoke-PodcastRun {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Forwards common confirmation/preview parameters to the real run; a second ShouldProcess here would change product behavior.')]
        [CmdletBinding(SupportsShouldProcess)]
        param([string]$Mode = 'Latest', [int]$CustomCount, [string]$OutputPath, [string]$FeedUrl,
            [switch]$NonInteractive, [switch]$KeepAwake, [string]$LegacyPath, [string]$LegacyAction,
            [string]$LegacyEpisodeId, [string]$LegacyFile, [string]$LegacySha256, [string]$LegacyCheckpoint,
            [string]$DiagnosticExportPath, [int]$MaxFeedPages = 20, [int]$MaxAttempts = 3,
            [double]$HeaderTimeoutSeconds = 30, [double]$IdleTimeoutSeconds = 30, [double]$RetryBudgetSeconds = 120,
            [double]$BaseDelaySeconds = 1, [double]$MaxDelaySeconds = 30)
        $actor = $observation.Names[$OutputPath]
        $observation.Actor = $actor
        $observation.Active++
        $observation.Peak = [Math]::Max($observation.Peak, $observation.Active)
        $observation.Runs.Add([pscustomobject]@{ Stage = 'start'; Name = $actor; NonInteractive = [bool]$NonInteractive;
            WhatIf = [bool]$PSBoundParameters['WhatIf']; Confirm = [bool]$PSBoundParameters['Confirm'] })
        try { & $originalRun @PSBoundParameters }
        finally {
            $observation.Runs.Add([pscustomobject]@{ Stage = 'end'; Name = $actor })
            $observation.Active--
            $observation.Actor = $null
        }
    }
    $originalMetadata = (Get-Command Invoke-PodcastMetadataRequest).ScriptBlock
    function Invoke-PodcastMetadataRequest {
        [CmdletBinding()]
        param([string]$Uri, [long]$MaximumBytes = 8388608, $Policy)
        Add-UpdSavedRequest -Kind 'metadata' -Uri $Uri
        & $originalMetadata @PSBoundParameters
    }
    $originalMedia = (Get-Command Invoke-PodcastMediaRequest).ScriptBlock
    $originalRead = (Get-Command Read-PodcastResponseChunk).ScriptBlock
    function Read-PodcastResponseChunk {
        [CmdletBinding()]
        param($Source, [byte[]]$Buffer, [int]$Count, $Policy)
        $boundedCount = if ([long]$config.CancelAfterBytes -gt 0) { [Math]::Min($Count, 4096) } else { $Count }
        & $originalRead -Source $Source -Buffer $Buffer -Count $boundedCount -Policy $Policy
    }
    function Invoke-PodcastMediaRequest {
        [CmdletBinding()]
        param([string]$Uri, [IO.Stream]$DestinationStream, $Policy, $Resume, [scriptblock]$OnResponse, [scriptblock]$OnProgress)
        Add-UpdSavedRequest -Kind 'media' -Uri $Uri
        $forward = @{} + $PSBoundParameters
        $progressCallback = $OnProgress
        if ([long]$config.CancelAfterBytes -gt 0) {
            $forward.OnProgress = {
                param([long]$Bytes)
                if ($progressCallback) { $null = & $progressCallback $Bytes }
                if ($Bytes -ge [long]$config.CancelAfterBytes) {
                    $report.CancelBytes = $Bytes
                    throw [OperationCanceledException]::new('Synthetic saved-show cancellation sentinel.')
                }
            }.GetNewClosure()
        }
        & $originalMedia @forward
    }
    $originalProtect = (Get-Command Protect-PodcastSavedShowFeed).ScriptBlock
    function Protect-PodcastSavedShowFeed {
        [CmdletBinding()]
        param([string]$FeedUrl)
        $observation.Protect++
        & $originalProtect @PSBoundParameters
    }
    $originalLock = (Get-Command Enter-PodcastSavedShowLock).ScriptBlock
    function Enter-PodcastSavedShowLock {
        [CmdletBinding()]
        param([string]$ConfigPath)
        $observation.Locks++
        & $originalLock @PSBoundParameters
    }
    $originalPower = (Get-Command Start-PodcastKeepAwake).ScriptBlock
    function Start-PodcastKeepAwake {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Observes and forwards only the actual product-owned temporary lease in this synthetic child.')]
        [CmdletBinding()]
        param([switch]$Enabled)
        if ($Enabled) { $observation.Power++ }
        & $originalPower @PSBoundParameters
    }
    $output = @(Invoke-PodcastCommand -Options $options)
    if ($output.Count -ne 1 -or $output[0].Type -notin @('Podcast.ConfigResult', 'Podcast.RunResult', 'Podcast.BatchResult')) {
        throw 'The dispatcher did not return exactly one private result.'
    }
    $report.Result = $output[0]
    $workerExit = [int]$output[0].ExitCode
    foreach ($partial in @(Get-ChildItem -LiteralPath $root -Recurse -Force -File -Filter '.upd-*.tmp' -ErrorAction SilentlyContinue)) {
        $handle = [IO.File]::Open($partial.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $handle.Dispose()
        $report.PartialClosed = $true
    }
    foreach ($lock in @(Get-ChildItem -LiteralPath $root -Recurse -Force -File -Filter 'writer.lock' -ErrorAction SilentlyContinue)) {
        $handle = [IO.File]::Open($lock.FullName, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $handle.Dispose()
        $report.LockClosed = $true
    }
    $report.HostSurvived = $true
}
catch { $report.Error = $_.Exception.Message }
$report.RunTrace = @($observation.Runs.ToArray())
$report.RequestTrace = @($observation.Requests.ToArray())
$report.PeakActiveRuns = $observation.Peak
$report.ProtectCalls = $observation.Protect
$report.ConfigLockCalls = $observation.Locks
$report.PowerRequested = $observation.Power
[IO.File]::WriteAllText($config.ResultPath, ($report | ConvertTo-Json -Depth 15), [Text.UTF8Encoding]::new($false))
exit $workerExit
