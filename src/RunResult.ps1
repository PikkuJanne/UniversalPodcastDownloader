function Resolve-PodcastCliOptions {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Options names the normalized command-line option set.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$BoundParameters,
        [ValidateSet('Latest', 'All', 'Custom')][string]$Mode = 'Latest',
        [int]$CustomCount, [string]$FeedUrl, [switch]$NonInteractive, [switch]$LegacyRequested
    )
    $keys = @($BoundParameters.Keys)
    $hasCount = $keys -contains 'CustomCount'
    $hasMode = $keys -contains 'Mode'
    if ($hasCount -and $CustomCount -lt 1) { throw '-CustomCount must be a positive integer.' }
    if ($hasCount -and $hasMode -and $Mode -ne 'Custom') {
        throw '-CustomCount conflicts with an explicit Latest or All mode.'
    }
    if ($hasCount -and -not $hasMode) { $Mode = 'Custom' }
    if (-not $LegacyRequested -and $Mode -eq 'Custom' -and -not $hasCount) {
        throw "Mode 'Custom' requires -CustomCount with a value >= 1."
    }
    if ($NonInteractive -and ($keys -notcontains 'FeedUrl' -or [string]::IsNullOrWhiteSpace($FeedUrl))) {
        throw '-NonInteractive requires an explicit nonblank -FeedUrl.'
    }
    if ($NonInteractive -and $keys -contains 'Confirm' -and $BoundParameters['Confirm']) {
        throw '-NonInteractive cannot be combined with -Confirm.'
    }
    return [pscustomobject]@{
        Mode = $Mode; CustomCount = $CustomCount
        NeedFeed = ($keys -notcontains 'FeedUrl')
        NeedCount = (-not $NonInteractive -and -not $LegacyRequested -and -not $hasMode -and -not $hasCount)
    }
}

function Test-PodcastCancellation {
    [CmdletBinding()]
    param([object]$ErrorObject)
    $cause = $ErrorObject
    if ($cause -is [Management.Automation.ErrorRecord]) { $cause = $cause.Exception }
    for ($depth = 0; $cause -is [Exception] -and $depth -lt 16; $depth++) {
        if ($cause -is [OperationCanceledException] -or $cause -is [Management.Automation.PipelineStoppedException]) { return $true }
        $cause = $cause.InnerException
    }
    return $false
}

function New-PodcastEpisodeResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only an in-memory result value.')]
    [CmdletBinding()]
    param(
        [string]$EpisodeId,
        [Parameter(Mandatory)][ValidateSet('downloaded', 'verified_skip', 'legacy_unverified', 'conflict', 'deferred', 'failed', 'cancelled', 'planned')][string]$Outcome,
        [Nullable[long]]$Bytes, [string]$Verification, [ValidateRange(0, 10)][int]$Attempts = 0,
        [string]$Message
    )
    return [pscustomobject]@{
        EpisodeId = $EpisodeId; Outcome = $Outcome; Bytes = $Bytes
        Verification = $Verification; Attempts = $Attempts; Message = $Message
    }
}

function New-PodcastRunResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only an in-memory result value.')]
    [CmdletBinding()]
    param(
        [ValidateSet('Latest', 'All', 'Custom')][string]$Mode = 'Latest', [switch]$Preview,
        [object[]]$EpisodeResults = @(), [ValidateRange(0, 2147483647)][int]$Planned = 0,
        [object]$Catalogue, [switch]$Fatal, [switch]$Cancelled, [string]$Message,
        [object[]]$Plan = @(), [object]$LegacyResult
    )
    $EpisodeResults = @($EpisodeResults | Where-Object { $null -ne $_ })
    $Plan = @($Plan | Where-Object { $null -ne $_ })
    $counts = @{}
    foreach ($outcome in @('downloaded', 'verified_skip', 'legacy_unverified', 'conflict', 'failed', 'deferred', 'cancelled')) {
        $counts[$outcome] = @($EpisodeResults | Where-Object { $_.Outcome -eq $outcome }).Count
    }
    $catalogueComplete = $null -eq $Catalogue -or [bool]$Catalogue.Complete
    $exitCode = 0
    $status = if ($Preview) { 'preview' } else { 'success' }
    if (-not $catalogueComplete -or ($counts.failed + $counts.deferred + $counts.conflict + $counts.legacy_unverified) -gt 0) {
        $exitCode = 2; $status = 'incomplete'
    }
    if ($Fatal) { $exitCode = 1; $status = 'fatal' }
    if ($Cancelled -or $counts.cancelled -gt 0) { $exitCode = 130; $status = 'cancelled' }
    $result = [pscustomobject]@{
        Type = 'Podcast.RunResult'; SchemaVersion = 1; Mode = $Mode; Preview = [bool]$Preview
        Complete = ($exitCode -eq 0); ExitCode = $exitCode; Status = $status; Planned = $Planned
        Downloaded = $counts.downloaded; VerifiedSkipped = $counts.verified_skip
        LegacyUnverified = $counts.legacy_unverified; Conflicts = $counts.conflict
        Failed = $counts.failed; Deferred = $counts.deferred; Cancelled = $counts.cancelled
        CatalogueComplete = $catalogueComplete
        CatalogueStopReason = $(if ($null -ne $Catalogue) { $Catalogue.StopReason } else { $null })
        PagesFetched = $(if ($null -ne $Catalogue) { [int]$Catalogue.PagesFetched } else { 0 })
        Episodes = @($EpisodeResults); Message = $Message; Plan = @($Plan)
    }
    # Explicit local legacy review/action data remains private and opt-in.
    if ($null -ne $LegacyResult) { $result | Add-Member -NotePropertyName LegacyResult -NotePropertyValue $LegacyResult }
    return $result
}
