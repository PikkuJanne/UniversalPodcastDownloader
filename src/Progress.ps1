#requires -Version 5.1

function Test-PodcastProgressInteractive {
    [CmdletBinding()]
    param()
    try {
        return ($Host.Name -eq 'ConsoleHost' -and -not [Console]::IsOutputRedirected -and -not [Console]::IsErrorRedirected)
    }
    catch {
        if (Test-PodcastCancellation -ErrorObject $_) { throw }
        return $false
    }
}

function New-PodcastProgressContext {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only a private in-memory progress value.')]
    [CmdletBinding()]
    param([ValidateRange(0, 2147483647)][int]$TotalEpisodes, [switch]$NonInteractive, [scriptblock]$Clock)

    if ($null -eq $Clock) { $Clock = { [double][Diagnostics.Stopwatch]::GetTimestamp() * 1000.0 / [Diagnostics.Stopwatch]::Frequency } }
    $reason = if ($NonInteractive) { 'noninteractive' }
        elseif ([string]$ProgressPreference -ne 'Continue') { 'preference' }
        elseif ($TotalEpisodes -eq 0) { 'empty' }
        elseif (-not (Test-PodcastProgressInteractive)) { 'host' }
        else { '' }
    # A private random pair avoids the caller's ordinary progress IDs (0/1).
    $rootId = ([BitConverter]::ToInt32([guid]::NewGuid().ToByteArray(), 0) -band 1073741822) + 1024
    return [pscustomobject]@{
        Enabled = ($reason.Length -eq 0); QuietReason = $reason; RootId = $rootId; TransferId = $rootId + 1
        TotalEpisodes = $TotalEpisodes; EpisodeIndex = 0; ProcessedEpisodes = 0; VerifiedEpisodes = 0
        Attempt = 0; Offset = [long]0; Bytes = [long]0; TotalBytes = $null
        Stage = 'preparing'; EpisodeOutcome = ''; EpisodeFinished = $false; TransferStarted = $false
        RunVerified = $false; EpisodePercent = 0; TransferPercent = -1
        ShownIds = [Collections.Generic.List[int]]::new(); Closed = $false; Clock = $Clock; LastRenderedAt = $null
    }
}

function Write-PodcastProgressContext {
    [CmdletBinding()]
    param([AllowNull()]$Context, [switch]$Force)
    if ($null -eq $Context -or $Context.Closed) { return }
    $Context.EpisodePercent = if ($Context.RunVerified) { 100 }
        elseif ($Context.TotalEpisodes -gt 0) { [Math]::Min(99, [int][Math]::Floor(100.0 * $Context.ProcessedEpisodes / $Context.TotalEpisodes)) }
        else { 0 }
    $Context.TransferPercent = if ($Context.EpisodeOutcome -eq 'downloaded') { 100 }
        elseif ($null -eq $Context.TotalBytes) { -1 }
        elseif ($Context.TotalBytes -eq 0) { 0 }
        else { [int][Math]::Min(99.0, [Math]::Floor(100.0 * [double]$Context.Bytes / [double]$Context.TotalBytes)) }
    if (-not $Context.Enabled) { return }
    if ([string]$ProgressPreference -ne 'Continue') { $Context.Enabled = $false; return }
    try {
        $now = [double](& $Context.Clock)
        if (-not $Force -and $null -ne $Context.LastRenderedAt -and $now - $Context.LastRenderedAt -lt 250.0) { return }
        $index = $Context.EpisodeIndex
        $status = switch ($Context.Stage) {
            'preparing' { 'Preparing: episode {0}' -f $index }
            'verifying' { 'Verifying: episode {0}' -f $index }
            'downloaded' { 'Verified: episode {0}' -f $index }
            'verified_skip' { 'Verified history: episode {0}' -f $index }
            'legacy_unverified' { 'Local media remains unverified: episode {0}' -f $index }
            'conflict' { 'Preserved conflict: episode {0}' -f $index }
            'deferred' { 'Deferred: episode {0}' -f $index }
            'failed' { 'Failed: episode {0}' -f $index }
            'cancelled' { 'Cancelled: episode {0}' -f $index }
            'complete' { 'All selected episodes verified.' }
            'incomplete' { 'Processing ended; verified completion was not established.' }
            default { 'Receiving: episode {0}' -f $index }
        }
        $status += '; processed {0}/{1}; verified {2}' -f $Context.ProcessedEpisodes, $Context.TotalEpisodes, $Context.VerifiedEpisodes
        if (-not $Context.ShownIds.Contains($Context.RootId)) { $Context.ShownIds.Add($Context.RootId) }
        $null = Write-Progress -Id $Context.RootId -Activity 'Podcast downloads' -Status $status `
            -CurrentOperation ('Episode {0} of {1}' -f $index, $Context.TotalEpisodes) -PercentComplete $Context.EpisodePercent
        if ($Context.TransferStarted) {
            $bytes = $Context.Bytes.ToString([Globalization.CultureInfo]::InvariantCulture)
            $byteText = if ($null -eq $Context.TotalBytes) { $bytes + ' bytes (total unknown)' }
                else { $bytes + ' of ' + $Context.TotalBytes.ToString([Globalization.CultureInfo]::InvariantCulture) + ' bytes' }
            $phase = if ($Context.Stage -eq 'verifying') { 'Validating and recording' }
                elseif ($Context.EpisodeOutcome -eq 'downloaded') { 'Verified transfer' }
                elseif ($Context.EpisodeOutcome -eq 'failed') { 'Transfer failed' }
                elseif ($Context.EpisodeOutcome -eq 'cancelled') { 'Transfer cancelled' }
                elseif ($Context.EpisodeOutcome -eq 'deferred') { 'Transfer deferred' }
                else { 'Receiving' }
            if (-not $Context.ShownIds.Contains($Context.TransferId)) { $Context.ShownIds.Add($Context.TransferId) }
            $null = Write-Progress -Id $Context.TransferId -ParentId $Context.RootId -Activity 'Episode media' `
                -Status ('{0}; attempt {1}; {2}' -f $phase, $Context.Attempt, $byteText) -PercentComplete $Context.TransferPercent
        }
        $Context.LastRenderedAt = $now
    }
    catch {
        if (Test-PodcastCancellation -ErrorObject $_) { throw }
        # A host rendering failure must not change the media/history outcome.
        $Context.Enabled = $false
    }
}

function Start-PodcastEpisodeProgress {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Updates only private progress state and owned console records; archive confirmation belongs to the caller.')]
    [CmdletBinding()]
    param([AllowNull()]$Context, [ValidateRange(1, 2147483647)][int]$Index)
    if ($null -eq $Context -or $Context.Closed) { return }
    # Retire the previous episode's byte bar before preparing another episode.
    if ($Context.ShownIds.Contains($Context.TransferId)) {
        try {
            $null = & {
                param($progressContext)
                $ProgressPreference = 'Continue'
                $null = Write-Progress -Id $progressContext.TransferId -Activity 'Episode media' -Completed
            } $Context
        }
        catch {
            if (Test-PodcastCancellation -ErrorObject $_) { throw }
            $Context.Enabled = $false
        }
    }
    $Context.EpisodeIndex = $Index; $Context.Stage = 'preparing'; $Context.EpisodeOutcome = ''
    $Context.EpisodeFinished = $false; $Context.TransferStarted = $false; $Context.Attempt = 0
    $Context.Offset = [long]0; $Context.Bytes = [long]0; $Context.TotalBytes = $null
    Write-PodcastProgressContext -Context $Context -Force
}

function Start-PodcastTransferProgress {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Updates only private progress state and owned console records; archive confirmation belongs to the caller.')]
    [CmdletBinding()]
    param([AllowNull()]$Context, [ValidateRange(0, 9223372036854775807)][long]$Offset = 0,
        [Nullable[long]]$TotalBytes, [ValidateRange(1, 2147483647)][int]$Attempt)
    if ($null -eq $Context -or $Context.Closed) { return }
    if ($null -ne $TotalBytes -and $TotalBytes -lt 0) { throw 'Progress byte counts must be nonnegative.' }
    $Context.Attempt = if ($PSBoundParameters.ContainsKey('Attempt')) { $Attempt } else { $Context.Attempt + 1 }
    $Context.Offset = $Offset; $Context.Bytes = $Offset; $Context.TotalBytes = $TotalBytes
    $Context.Stage = 'response'; $Context.TransferStarted = $true; $Context.EpisodeOutcome = ''
    Write-PodcastProgressContext -Context $Context -Force
}

function Update-PodcastTransferProgress {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Updates only private progress state and owned console records; archive confirmation belongs to the caller.')]
    [CmdletBinding()]
    param([AllowNull()]$Context, [ValidateRange(0, 9223372036854775807)][long]$Bytes,
        [ValidateSet('bytes', 'verifying')][string]$Stage = 'bytes')
    if ($null -eq $Context -or $Context.Closed) { return }
    $force = $Context.Stage -ne $Stage
    $Context.Bytes = $Bytes; $Context.Stage = $Stage
    Write-PodcastProgressContext -Context $Context -Force:$force
}

function Complete-PodcastEpisodeProgress {
    [CmdletBinding()]
    param([AllowNull()]$Context,
        [ValidateSet('downloaded', 'verified_skip', 'legacy_unverified', 'conflict', 'deferred', 'failed', 'cancelled')][string]$Outcome)
    if ($null -eq $Context -or $Context.Closed -or $Context.EpisodeFinished) { return }
    $Context.EpisodeFinished = $true; $Context.EpisodeOutcome = $Outcome; $Context.Stage = $Outcome
    $Context.ProcessedEpisodes++
    if ($Outcome -in @('downloaded', 'verified_skip')) { $Context.VerifiedEpisodes++ }
    Write-PodcastProgressContext -Context $Context -Force
}

function Complete-PodcastRunProgress {
    [CmdletBinding()]
    param([AllowNull()]$Context, [bool]$Verified = $false)
    if ($null -eq $Context -or $Context.Closed) { return }
    $Context.RunVerified = $Verified -and $Context.ProcessedEpisodes -eq $Context.TotalEpisodes -and $Context.VerifiedEpisodes -eq $Context.TotalEpisodes
    $Context.Stage = if ($Context.RunVerified) { 'complete' } else { 'incomplete' }
    Write-PodcastProgressContext -Context $Context -Force
}

function Close-PodcastProgress {
    [CmdletBinding()]
    param([AllowNull()]$Context)
    if ($null -eq $Context -or $Context.Closed) { return }
    $Context.Closed = $true
    $firstCancellation = $null
    # Completion only retires this context's own records. A local preference
    # avoids prompting and clears them even if the caller stopped new rendering.
    $ProgressPreference = 'Continue'
    # Children close before parents; never send a host-wide completion record.
    foreach ($id in @($Context.TransferId, $Context.RootId)) {
        if (-not $Context.ShownIds.Contains($id)) { continue }
        try { $null = Write-Progress -Id $id -Activity 'Podcast downloads' -Completed }
        catch { if ($null -eq $firstCancellation -and (Test-PodcastCancellation -ErrorObject $_)) { $firstCancellation = $_ } }
    }
    if ($null -ne $firstCancellation) { throw $firstCancellation }
}
