<#
UniversalPodcastDownloader.ps1
Minimal Win11 podcast downloader for personal offline archiving

Author: Janne Vuorela
Target OS: Windows 11
Dependencies: PowerShell 5.1+ or PowerShell 7+, Internet access

SYNOPSIS
    One-preset, no-frills podcast downloader intended for my own workflow:
    archiving favorite shows from any RSS/Atom feed into a clean folder structure,
    with a log file per run for later inspection.

WHAT THIS IS (AND ISN’T)
    - Personal, purpose-built tool for my specific use case.
      It trades advanced features for predictability, robustness, and a simple TUI.
    - Text-UI when launched via the .bat wrapper (double-click).
      Also supports direct PowerShell use with parameters.
    - Universal RSS/Atom client:
        - Can use a direct feed URL (ends in .xml, /feed, etc.)
        - Or a normal “show page” URL and TRY to auto-detect the RSS feed.
    - Focused on reliable downloads and traceability via log files,
      not on fancy post-processing, tagging, or re-encoding.

FEATURES
    - Guided TUI workflow:
        - Explains how to find the RSS feed.
        - Lets you paste either:
            - RSS feed URL, or
            - Show page URL for auto-detection.
        - Asks how many newest episodes to download:
            - Enter  = latest only
            - Number = N newest episodes
            - all    = accessible entries in the bounded feed catalogue
    - Per-podcast subfolders based on feed title:
        - Root:   <OutputPath> (default: %USERPROFILE%\Downloads\Podcasts)
        - Folder: <SafeFeedTitle>-<feed identity hash>
        - Files:  YYYY-MM-DD - Episode title-<episode identity hash>.mp3
    - Robust download loop:
        - Up to 3 attempts per episode with short delay between tries.
        - Skips episodes only after checking recorded size and SHA-256 on disk.
        - Summarizes downloaded / skipped / failed at the end.
    - Interactive episode and byte progress:
        - Unknown body lengths show received bytes without a fabricated total.
        - 100% follows verified transfer/history success; redirected output stays quiet.
    - Optional -KeepAwake temporarily requests Windows system wakefulness:
        - Disabled by default; active only after confirmation and released in finally.
        - Does not keep the display on or change the Windows power plan.
    - Explicit feed pagination:
        - Follows feed-level Atom next/prev-archive links before date selection.
        - MaxFeedPages defaults to 20 (1-100); entry/metadata limits also apply.
        - Gaps, limits and cycles report incomplete after accessible downloads.
        - A complete historical catalogue is not guaranteed.
    - Private per-run UTF-8 diagnostics:
        - Startup: LOCALAPPDATA\UniversalPodcastDownloader\Logs (TEMP fallback).
        - Confirmed downloads also keep a log in the podcast folder.
        - Named: download-<UTC timestamp>-<unique run ID>.log
        - Request displays contain only host and opaque per-run ID.
        - Logs use counts, episode identities and safe error categories.
        - Optional -DiagnosticExportPath saves a restricted JSON summary.
        - WhatIf and legacy Preview create no diagnostic files.
        - See docs/codex/DIAGNOSTICS.md for privacy and retention details.

MY INTENDED USAGE
    - I double-click UniversalPodcastDownloader.bat.
    - I paste either:
        - The podcast’s RSS feed URL, or
        - The public show page URL and let the script auto-detect RSS.
    - I tell it how many newest episodes I want (often “all” once per show).
    - I leave it running, in the morning I check:
        - The per-show folder for MP3s, and
        - The .log file if something failed mid-run.

SETUP
    1) Place these files together in a folder of your choice:
         - UniversalPodcastDownloader.ps1
         - UniversalPodcastDownloader.bat   (wrapper to allow double-click)
         - src/                            (bundled PowerShell helper files)
    2) Optional: pin the .bat to Start or Taskbar for quick access.
    3) Ensure the machine has:
         - Working Internet connection.
         - Permission to write to the default output:
             %USERPROFILE%\Downloads\Podcasts
           or whichever OutputPath you configure.
    4) No external binaries required. Relies on the built-in .NET HTTP client
       and the built-in XML parser in PowerShell.

USAGE
    A) Double-click for TUI (default usage)
        - Double-click UniversalPodcastDownloader.bat.
        - Follow prompts:
            1) Paste RSS feed URL or show page URL.
            2) Confirm or adjust detected feed.
            3) Choose how many newest episodes to download:
                 - Enter  = latest only
                 - Number = N newest
                 - all    = bounded accessible feed catalogue
        - Output:
            - Audio files under:
                %USERPROFILE%\Downloads\Podcasts\<FeedTitle>\
            - Log file:
                download-<UTC timestamp>-<unique run ID>.log

    B) Direct PowerShell (interactive TUI still available)
        - Run without parameters for the same guided TUI:
            .\UniversalPodcastDownloader.ps1
        - You can override defaults:
            .\UniversalPodcastDownloader.ps1 -OutputPath "D:\Podcasts"

    C) Direct PowerShell (non-interactive)
        - Add -NonInteractive with an explicit -FeedUrl to suppress input prompts:
            .\UniversalPodcastDownloader.ps1 
                -NonInteractive
                -FeedUrl "https://example.com/feed.xml" 
                -Mode All 
                -OutputPath "D:\Podcasts"
        - Or: download the 10 newest episodes:
            .\UniversalPodcastDownloader.ps1 
                -FeedUrl "https://example.com/feed.xml" 
                -Mode Custom 
                -CustomCount 10

        - CustomCount alone implies Custom; explicit Latest/All plus a count fails.
        - Exit codes: 0 success, 1 fatal input/setup, 2 incomplete, 130 catchable cancellation.
        - Dot-source, then call Invoke-PodcastRun for a structured result without host exit.
        - -PassThru emits the entry script result; parameterized batch launches never pause.
        - Add -KeepAwake to request temporary Windows wakefulness for confirmed work.
          It is off by default; explicit sleep/lid actions can still take effect.

NOTES
    - Episodes are sorted by UTC publication instant newest first; tied/undated entries keep source order.
    - New folder and file names include deterministic SHA-256 identity suffixes.
    - Windows names are sanitized and shortened to fit the selected output root.
    - Unsafe archive paths fail before archive writes; startup diagnostics use a separate safe location.
    - Recorded transfer skips check local size and SHA-256. Adopted files stay distinct.
    - Use -LegacyPath for a read-only inventory and explicit adoption/redownload choices.
    - Original supported audio bytes are retained; completed byte evidence selects
      canonical MP3/M4A/Ogg/WAVE/FLAC extensions for new media destinations.

LIMITATIONS
    - Resume requires an owned checkpoint and matching representation evidence;
      unknown partials remain preserved for review.
    - No built-in scheduler, this is a manual “run when needed” tool.
    - Assumes feeds expose downloadable audio URLs, video-only or exotic feeds may fail.
    - Uses one fixed naming pattern and folder structure.
    - No content-based deduplication across different feeds or folder names.

TROUBLESHOOTING
    - Script window closes immediately on double-click:
        - Run the .bat from an existing cmd window to see errors.
        - Check ExecutionPolicy or corporate restrictions on PowerShell scripts.
    - “Feed parsed, but no downloadable enclosure URLs were found”:
        - The feed may not expose direct audio URLs, or uses a non-standard format.
        - Try locating the real RSS URL from the hosting provider (Podbean, etc.).
    - “No episodes found in the feed” or XML errors:
        - Double-check the URL is a valid RSS/Atom feed, not just a random web page.
        - Some providers require HTTPS and modern TLS, ensure the system is up to date.
    - Only part of the episodes downloaded:
        - Check the .log file in the podcast folder for:
            - Per-episode errors
            - Timeouts or connectivity issues
            - HTTP status codes from the host

LICENSE / WARRANTY
    - Personal tool, provided as-is, without warranty. Use at your own risk.
#>

<#
UniversalPodcastDownloader.ps1
Minimal Win11 podcast downloader for personal offline archiving

Author: Janne Vuorela
Target OS: Windows 11
Dependencies: PowerShell 5.1+ or PowerShell 7+, Internet access, optional .bat wrapper
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Latest','All','Custom')]
    [string]$Mode = 'Latest',

    [int]$CustomCount,

    [string]$OutputPath = "$env:USERPROFILE\Downloads\Podcasts",

    [string]$FeedUrl,

    [switch]$NonInteractive,

    [switch]$KeepAwake,

    [switch]$PassThru,

    [string]$LegacyPath,

    [ValidateSet('Preview','Adopt','Redownload','Rollback')]
    [string]$LegacyAction = 'Preview',

    [string]$LegacyEpisodeId,

    [string]$LegacyFile,

    [string]$LegacySha256,

    [string]$LegacyCheckpoint,

    [string]$DiagnosticExportPath,

    [ValidateRange(1, 100)][int]$MaxFeedPages = 20,
    [ValidateRange(1, 10)][int]$MaxAttempts = 3,
    [ValidateRange(0.001, 86400)][double]$HeaderTimeoutSeconds = 30,
    [ValidateRange(0.001, 86400)][double]$IdleTimeoutSeconds = 30,
    [ValidateRange(0, 86400)][double]$RetryBudgetSeconds = 120,
    [ValidateRange(0, 3600)][double]$BaseDelaySeconds = 1,
    [ValidateRange(0, 3600)][double]$MaxDelaySeconds = 30
)

. (Join-Path $PSScriptRoot 'src/PublicationDate.ps1')
. (Join-Path $PSScriptRoot 'src/MediaSelection.ps1')
. (Join-Path $PSScriptRoot 'src/Naming.ps1')
. (Join-Path $PSScriptRoot 'src/PathSafety.ps1')
. (Join-Path $PSScriptRoot 'src/Diagnostics.ps1')
. (Join-Path $PSScriptRoot 'src/RunResult.ps1')
. (Join-Path $PSScriptRoot 'src/Progress.ps1')
. (Join-Path $PSScriptRoot 'src/KeepAwake.ps1')
. (Join-Path $PSScriptRoot 'src/HistoryStore.ps1')
. (Join-Path $PSScriptRoot 'src/HistoryIdentity.ps1')
. (Join-Path $PSScriptRoot 'src/HistoryWorkflow.ps1')
. (Join-Path $PSScriptRoot 'src/NetworkPolicy.ps1')
. (Join-Path $PSScriptRoot 'src/FeedXml.ps1')
. (Join-Path $PSScriptRoot 'src/FeedDiscovery.ps1')
. (Join-Path $PSScriptRoot 'src/FeedPagination.ps1')
. (Join-Path $PSScriptRoot 'src/MediaRequest.ps1')
. (Join-Path $PSScriptRoot 'src/MediaValidation.ps1')
. (Join-Path $PSScriptRoot 'src/ResumeStore.ps1')
. (Join-Path $PSScriptRoot 'src/MediaTransfer.ps1')
. (Join-Path $PSScriptRoot 'src/LegacyInventory.ps1')
. (Join-Path $PSScriptRoot 'src/LegacyMigration.ps1')

$script:PodcastTransportPolicy = New-PodcastTransportPolicy -MaxAttempts $MaxAttempts `
    -HeaderTimeoutSeconds $HeaderTimeoutSeconds -IdleTimeoutSeconds $IdleTimeoutSeconds `
    -RetryBudgetSeconds $RetryBudgetSeconds -BaseDelaySeconds $BaseDelaySeconds -MaxDelaySeconds $MaxDelaySeconds
$script:PodcastMaxFeedPages = $MaxFeedPages


function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    Write-PodcastDiagnostic -Message $Message -Level $Level
}

function Invoke-PodcastWebRequest {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    # Metadata stays in memory. Media writes use the confirmed transfer path.
    Invoke-PodcastMetadataRequest -Uri $Uri -Policy $script:PodcastTransportPolicy
}

# --- Shared RSS/Atom resolution and guided input ---
function Get-FeedUrlInteractive {
    Write-Host "==== Universal Podcast Downloader ====" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Paste a direct RSS/Atom feed URL or a podcast show page URL." -ForegroundColor Yellow
    Write-Host "For pages with several feeds, choose the feed number shown below."
    Write-Host ""

    while ($true) {
        $inputUrl = Read-Host "Paste RSS feed URL OR podcast page URL"
        if ([string]::IsNullOrWhiteSpace($inputUrl)) {
            Write-Host "Please paste a URL (or press Ctrl+C to quit)." -ForegroundColor Red
            continue
        }
        try {
            Write-Host "  Fetching URL..." -ForegroundColor DarkCyan
            return (Resolve-PodcastItems -Feeds @($inputUrl) -Interactive)
        } catch {
            Write-Host ("  Could not resolve the feed: " + (Get-PodcastDiagnosticError -Error $_)) -ForegroundColor Red
        }
    }
}

function Resolve-PodcastItems {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Retains the established import-safe helper name for existing callers.')]
    [CmdletBinding()]
    param([string[]]$Feeds, [switch]$Interactive, [AllowNull()]$InitialResolution,
        [ValidateRange(1, 100)][int]$MaxPages = $script:PodcastMaxFeedPages)

    $lastFailure = 'No RSS or Atom feed links were found on the page.'
    $initial = $InitialResolution
    foreach ($u in $Feeds) {
        Write-Verbose ("Trying feed: " + (Get-PodcastSafeUrl -Url $u))
        try {
            if ($null -ne $initial) {
                $source = $initial
                $initial = $null
                if (-not [StringComparer]::Ordinal.Equals([string]$source.Url, $u)) {
                    throw 'The initial feed resolution does not belong to the supplied URL.'
                }
                if ($null -ne $source.PSObject.Properties['Catalogue']) { return $source }
            } else {
                $source = Resolve-PodcastSource -Uri $u
            }
            $candidates = @($source.Candidates)
            if ($source.Kind -eq 'Html') {
                if ($candidates.Count -eq 0) { throw 'No RSS or Atom feed links were found on the page.' }
                if ($candidates.Count -gt 1 -and -not $Interactive) {
                    throw 'Multiple feed links were found. Supply a direct feed URL with -FeedUrl.'
                }
                $index = 0
                if ($candidates.Count -gt 1) {
                    Write-Host 'Multiple feeds found:' -ForegroundColor Yellow
                    for ($i = 0; $i -lt $candidates.Count; $i++) {
                        Write-Host ("  {0}. {1}" -f ($i + 1), (Get-PodcastSafeUrl -Url $candidates[$i]))
                    }
                    while ($true) {
                        $choice = Read-Host ("Choose feed number (1-{0})" -f $candidates.Count)
                        $number = 0
                        if ([int]::TryParse($choice, [ref]$number) -and $number -ge 1 -and $number -le $candidates.Count) {
                            $index = $number - 1
                            break
                        }
                        Write-Host 'Enter one of the listed feed numbers.' -ForegroundColor Red
                    }
                }
                # A page can discover one selected feed, never a recursive crawl.
                $feed = Resolve-PodcastSource -Uri $candidates[$index]
                if ($feed.Kind -eq 'Html') { throw 'The discovered URL did not return an RSS or Atom feed.' }
            } else {
                $feed = $source
            }
            $catalogue = Resolve-PodcastCatalogue -InitialResolution $feed -MaxPages $MaxPages
            if (@($catalogue.Items).Count -eq 0) {
                if (-not $catalogue.Catalogue.Complete) {
                    $gap = [InvalidOperationException]::new('Feed catalogue incomplete; no accessible episodes were found.')
                    $gap.Data['PodcastCatalogue'] = $catalogue.Catalogue
                    throw $gap
                }
                throw 'The RSS or Atom feed is valid but contains no episodes.'
            }
            return [pscustomobject]@{
                Url = $feed.Url
                Xml = $feed.Xml
                Items = @($catalogue.Items)
                Kind = $feed.Kind
                FinalUri = $feed.FinalUri
                Content = $feed.Content
                Candidates = $candidates
                SourceUrl = $source.Url
                Catalogue = $catalogue.Catalogue
            }
        } catch {
            if (Test-PodcastCancellation -ErrorObject $_) { throw }
            if ($null -ne $_.Exception.Data['PodcastCatalogue']) { throw }
            $transportFailure = Get-PodcastTransportFailure -ErrorObject $_
            if ($null -ne $transportFailure) { throw $transportFailure }
            $lastFailure = $_
            Write-Verbose ("Feed resolution stopped: {0} ({1})" -f (Get-PodcastSafeUrl -Url $u), (Get-PodcastDiagnosticError -Error $_))
        }
    }
    throw $lastFailure
}

# --- Robust episode extraction + collision-proof filenames ---
function Get-FirstText {
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Array]) { $Value = $Value | Select-Object -First 1 }
    if ($Value -is [string]) { return $Value }

    try { return ($Value.InnerText) } catch { return "$Value" }
}

function Get-XPathText {
    param(
        [xml.XmlNode]$Node,
        [string]$XPath
    )
    $n = $Node.SelectSingleNode($XPath)
    if ($n) { return $n.InnerText }
    return $null
}

function Get-EpisodeData {
    param([xml.XmlNode]$XmlItem)

    $title = Get-XPathText -Node $XmlItem -XPath './*[local-name()="title"][1]'
    if (-not $title) { $title = Get-FirstText $XmlItem.title }

    $isAtom = $XmlItem.LocalName -ceq 'entry' -and $XmlItem.NamespaceURI -ceq 'http://www.w3.org/2005/Atom'
    $dateStr = $null
    $dateSource = $null
    $dateFields = if ($isAtom) { @('published', 'updated') } else { @('pubDate') }
    $dateNamespace = if ($isAtom) { 'http://www.w3.org/2005/Atom' } else { '' }
    foreach ($field in $dateFields) {
        $candidate = Get-XPathText -Node $XmlItem -XPath ('./*[local-name()="{0}" and namespace-uri()="{1}"][1]' -f $field, $dateNamespace)
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $dateStr = $candidate
            $dateSource = $field
            break
        }
    }
    $pubDate = ConvertTo-PodcastPublicationDate -Value $dateStr

    # Preserve old date priority only for historical filename review hints.
    $legacyDateStr = $null
    foreach ($field in @('pubDate', 'updated', 'published')) {
        $legacyCandidate = Get-XPathText -Node $XmlItem -XPath ('./*[local-name()="{0}"][1]' -f $field)
        if (-not [string]::IsNullOrWhiteSpace($legacyCandidate)) { $legacyDateStr = $legacyCandidate; break }
    }
    $legacyPubDate = Get-PodcastLegacyPublicationDate -Value $legacyDateStr

    $guid = Get-XPathText -Node $XmlItem -XPath './*[local-name()="guid"][1]'
    if (-not $guid) { $guid = Get-FirstText $XmlItem.guid }
    # Retain the established ID lookup so date normalization cannot rebind history.
    $atomId = Get-XPathText -Node $XmlItem -XPath './*[local-name()="id"][1]'

    $candidates = [Collections.Generic.List[object]]::new()
    foreach ($node in $XmlItem.SelectNodes('./*[local-name()="enclosure" or (local-name()="link" and @rel="enclosure")]')) {
        $candidateUrl = if ($node.LocalName -ceq 'enclosure') { $node.GetAttribute('url') } else { $node.GetAttribute('href') }
        if ([string]::IsNullOrWhiteSpace($candidateUrl)) { continue }
        $length = $null
        $parsedLength = 0L
        if ([long]::TryParse($node.GetAttribute('length'), [Globalization.NumberStyles]::None,
                [Globalization.CultureInfo]::InvariantCulture, [ref]$parsedLength)) { $length = $parsedLength }
        $candidates.Add([pscustomobject]@{ Url = $candidateUrl; ContentType = $node.GetAttribute('type'); Length = $length })
    }

    if ($candidates.Count -eq 0) {
        $cands = @(
            Get-XPathText -Node $XmlItem -XPath './*[local-name()="guid"][1]'
            Get-XPathText -Node $XmlItem -XPath './*[local-name()="link"][1]'
            (Get-FirstText $XmlItem.guid)
            (Get-FirstText $XmlItem.link)
        ) | Where-Object { $_ }

        foreach ($cand in $cands) {
            if ($cand -match '\.(mp3|m4a|ogg|opus|wav|flac)(?:$|[?#])') {
                $candidates.Add([pscustomobject]@{ Url = [string]$cand; ContentType = ''; Length = $null })
                break
            }
        }
    }

    $media = Select-PodcastAudioCandidate -Candidates $candidates.ToArray()

    [PSCustomObject]@{
        Title   = ($title -as [string])
        PubDate = $pubDate
        PubDateOriginal = $dateStr
        PubDateSource = $dateSource
        LegacyPubDate = $legacyPubDate
        Url     = $(if ($null -ne $media) { $media.Url } else { $null })
        Guid    = ($guid -as [string])
        AtomId  = ($atomId -as [string])
        EnclosureLength = $(if ($null -ne $media) { $media.Length } else { $null })
        Candidates = @($candidates.ToArray())
        MediaContentType = $(if ($null -ne $media) { $media.ContentType } else { $null })
        MediaExtension = $(if ($null -ne $media) { $media.Extension } else { $null })
        MediaSelectionReason = $(if ($null -ne $media) { $media.Reason } else { 'no_supported_audio_candidate' })
    }
}

function Select-PodcastEpisode {
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Episodes,
        [ValidateSet('Latest','All','Custom')][string]$Mode = 'Latest',
        [int]$CustomCount
    )

    if ($Mode -eq 'Custom' -and $CustomCount -lt 1) {
        throw "Mode 'Custom' requires -CustomCount with a value >= 1."
    }

    $selected = Get-OrderedPodcastEpisode -Episodes $Episodes
    switch ($Mode) {
        'Latest' { $selected = @($selected | Select-Object -First 1) }
        'Custom' { $selected = @($selected | Select-Object -First $CustomCount) }
    }

    # Preserve an array for zero/one/many results when assigned by the caller.
    return ,$selected
}

function Invoke-PodcastRun {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [ValidateSet('Latest','All','Custom')]
        [string]$Mode = 'Latest',

        [int]$CustomCount,

        [string]$OutputPath = "$env:USERPROFILE\Downloads\Podcasts",

        [string]$FeedUrl,

        [switch]$NonInteractive,

        [switch]$KeepAwake,

        [string]$LegacyPath,

        [ValidateSet('Preview','Adopt','Redownload','Rollback')]
        [string]$LegacyAction = 'Preview',

        [string]$LegacyEpisodeId,

        [string]$LegacyFile,

        [string]$LegacySha256,

        [string]$LegacyCheckpoint,

        [string]$DiagnosticExportPath,

        [ValidateRange(1, 100)][int]$MaxFeedPages = 20,
        [ValidateRange(1, 10)][int]$MaxAttempts = 3,
        [ValidateRange(0.001, 86400)][double]$HeaderTimeoutSeconds = 30,
        [ValidateRange(0.001, 86400)][double]$IdleTimeoutSeconds = 30,
        [ValidateRange(0, 86400)][double]$RetryBudgetSeconds = 120,
        [ValidateRange(0, 3600)][double]$BaseDelaySeconds = 1,
        [ValidateRange(0, 3600)][double]$MaxDelaySeconds = 30
    )


    # Per-invocation preferences do not replace the caller's progress policy.
    $ErrorActionPreference = 'Stop'
    if ($NonInteractive) { $ConfirmPreference = 'None' }
    $episodeResults = @()
    $destinationPlan = @()
    $planned = $null
    $resolved = $null
    $total = 0
    $script:PodcastTransportPolicy = New-PodcastTransportPolicy -MaxAttempts $MaxAttempts `
        -HeaderTimeoutSeconds $HeaderTimeoutSeconds -IdleTimeoutSeconds $IdleTimeoutSeconds `
        -RetryBudgetSeconds $RetryBudgetSeconds -BaseDelaySeconds $BaseDelaySeconds -MaxDelaySeconds $MaxDelaySeconds
    $script:PodcastMaxFeedPages = $MaxFeedPages

    $historyLock = $null
    $archiveLock = $null
    $runResources = [pscustomobject]@{ Power = $null; Progress = $null; Result = $null }
    $onConfirmed = {
        if ($KeepAwake -and $null -eq $runResources.Power) {
            $runResources.Power = Start-PodcastKeepAwake -Enabled
            Write-Host ('[*] ' + $runResources.Power.Message)
            Write-PodcastDiagnostic -Message $runResources.Power.Message
        }
        if ($legacyRequested -and $LegacyAction -eq 'Redownload') {
            $runResources.Progress = New-PodcastProgressContext -TotalEpisodes 1 -NonInteractive:$NonInteractive
            Start-PodcastEpisodeProgress -Context $runResources.Progress -Index 1
        }
    }
    $onTransferProgress = {
        param($ProgressEvent)
        if ($ProgressEvent.Stage -eq 'response') {
            Start-PodcastTransferProgress -Context $runResources.Progress -Offset $ProgressEvent.Bytes `
                -TotalBytes $ProgressEvent.TotalBytes -Attempt $ProgressEvent.Attempt
        }
        else {
            Update-PodcastTransferProgress -Context $runResources.Progress -Bytes $ProgressEvent.Bytes -Stage $ProgressEvent.Stage
        }
    }
    $legacyRequested = $PSBoundParameters.ContainsKey('LegacyPath')
    $diagnosticPreview = $WhatIfPreference -or ($legacyRequested -and $LegacyAction -eq 'Preview')
    $diagnosticConfirmation = $ConfirmPreference -in @('Low', 'Medium')

    # --- Main ---
    try {
        # Each invocation owns a fresh in-memory context even when CLI validation fails.
        $null = Initialize-PodcastDiagnostics -Preview
        $cli = Resolve-PodcastCliOptions -BoundParameters $PSBoundParameters -Mode $Mode -CustomCount $CustomCount `
            -FeedUrl $FeedUrl -NonInteractive:$NonInteractive -LegacyRequested:$legacyRequested
        $Mode = $cli.Mode
        $CustomCount = $cli.CustomCount
        $null = Initialize-PodcastDiagnostics -Preview:($diagnosticPreview -or $diagnosticConfirmation)
        foreach ($option in @('LegacyAction', 'LegacyEpisodeId', 'LegacyFile', 'LegacySha256', 'LegacyCheckpoint')) {
            if ($PSBoundParameters.ContainsKey($option) -and -not $legacyRequested) {
                throw 'Legacy options require -LegacyPath and an explicit -FeedUrl.'
            }
        }
        if ($legacyRequested -and ([string]::IsNullOrWhiteSpace($LegacyPath) -or
                -not $PSBoundParameters.ContainsKey('FeedUrl') -or [string]::IsNullOrWhiteSpace($FeedUrl))) {
            throw 'Legacy review requires an existing -LegacyPath and an explicit -FeedUrl.'
        }
        $needFeed = $cli.NeedFeed
        $needCount = $cli.NeedCount

        $resolved = $null
        if ($needFeed) {
            $resolved = Get-FeedUrlInteractive
            $FeedUrl = $resolved.Url
        }

        if ($needCount) {
            Write-Host ""
            Write-Host "How many episodes to download (newest first)?" -ForegroundColor Yellow
            Write-Host "  - Press Enter for only the latest episode."
            Write-Host "  - Enter a number like 5 to download the 5 newest."
            Write-Host "  - Type 'all' to download all accessible entries within the feed limits."
            Write-Host ""

            while ($true) {
                $inputCount = Read-Host "Episodes to download (number / 'all' / Enter = 1)"
                if ([string]::IsNullOrWhiteSpace($inputCount)) {
                    $Mode = 'Latest'
                    break
                }

                $trimmed = $inputCount.Trim()

                if ($trimmed.ToLower() -eq 'all') {
                    $Mode = 'All'
                    break
                }

                $n = 0
                if ([int]::TryParse($trimmed, [ref]$n) -and $n -gt 0) {
                    $Mode        = 'Custom'
                    $CustomCount = $n
                    break
                }

                Write-Host "Invalid input. Enter 'all', a positive number, or just Enter." -ForegroundColor Red
            }
        }

        if (-not $FeedUrl) {
            throw "No feed URL specified. Use -FeedUrl or paste it via the TUI."
        }

        if (-not $legacyRequested -and $Mode -eq 'Custom' -and (-not $CustomCount -or $CustomCount -lt 1)) {
            throw "Mode 'Custom' requires -CustomCount with a value >= 1."
        }

        # Resolve a user-selected relative root once; metadata is never resolved as a path.
        $rootInput = $OutputPath
        if (-not [IO.Path]::IsPathRooted($rootInput)) {
            $rootInput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($rootInput)
        }
        $baseOutputPath = Assert-PodcastDestination -Root $rootInput -Directory
        $candidateFeeds = @($FeedUrl) | Where-Object { $_ } | Select-Object -Unique

        Write-Host "[*] Fetching podcast feed..."
        if ($null -eq $resolved) { $resolved = Resolve-PodcastItems -Feeds $candidateFeeds }
        Write-Host ("    Using feed: " + (Get-PodcastSafeUrl -Url $resolved.Url))
        Write-Host ('    Feed pages retrieved: {0}; repeated entries removed: {1}.' -f $resolved.Catalogue.PagesFetched, $resolved.Catalogue.DuplicateCount)
        Write-Host '    Selection uses accessible entries; a complete historical catalogue is not guaranteed.'
        if (-not $resolved.Catalogue.Complete) {
            $catalogueMessage = Get-PodcastCatalogueMessage -Catalogue $resolved.Catalogue
            Write-Warning $catalogueMessage
            Write-Log $catalogueMessage 'WARN'
        }

        # Feed title, PS5.1-safe
        $feedTitle = $null
        try {
            $node1 = $resolved.Xml.SelectSingleNode('//*[local-name()="rss"]/*[local-name()="channel"]/*[local-name()="title"][1]')
            if ($node1) { $feedTitle = $node1.InnerText }
            if (-not $feedTitle) {
                $node2 = $resolved.Xml.SelectSingleNode('//*[local-name()="feed"]/*[local-name()="title"][1]')
                if ($node2) { $feedTitle = $node2.InnerText }
            }
        } catch {}

        if ($feedTitle) {
            Write-Host '    Feed title available (kept in local archive metadata).'
        } else {
            Write-Host "    Feed title: (unknown)"
        }

        $episodes = @(foreach ($it in $resolved.Items) { Get-EpisodeData $it })
        $unsupportedCount = @($episodes | Where-Object { -not $_.Url }).Count
        if ($unsupportedCount -gt 0) {
            Write-Warning ("Feed entries without a supported audio candidate: {0}." -f $unsupportedCount)
            Write-Log ("Feed entries without a supported audio candidate: {0}." -f $unsupportedCount) 'WARN'
        }
        $episodes = @($episodes | Where-Object { $_.Url })
        $episodeCount = $episodes.Count
        if ($episodeCount -eq 0) {
            throw 'Feed parsed, but no downloadable enclosure URLs were found. No supported audio candidate was declared.'
        }
        # Validate enclosure targets during planning, before any archive write or
        # media request. The transport validates the original target and every hop again.
        foreach ($episode in $episodes) { $null = Get-PodcastRequestUri -Uri $episode.Url }
        $allEpisodes = $episodes
        if ($legacyRequested) {
            $legacyResult = Invoke-PodcastLegacyMigration -Root $baseOutputPath -LegacyRoot $LegacyPath -FeedUrl $resolved.Url `
                -Episodes $allEpisodes -Action $LegacyAction -EpisodeId $LegacyEpisodeId -FileName $LegacyFile `
                -Sha256 $LegacySha256 -Checkpoint $LegacyCheckpoint -Policy $script:PodcastTransportPolicy `
                -OnConfirmed $onConfirmed -OnProgress $onTransferProgress
            $legacyPreview = $WhatIfPreference -or $LegacyAction -eq 'Preview' -or $null -eq $legacyResult.Outcome
            if ($legacyResult.Outcome -eq 'transfer_verified') {
                $episodeResults += New-PodcastEpisodeResult -EpisodeId $legacyResult.EpisodeId -Outcome downloaded `
                    -Message 'Explicit legacy redownload completed; original media preserved.'
                Complete-PodcastEpisodeProgress -Context $runResources.Progress -Outcome downloaded
            }
            $legacyRunResult = New-PodcastRunResult -Mode $Mode -Preview:$legacyPreview -EpisodeResults $episodeResults `
                -Planned $(if ($legacyPreview) { @($legacyResult.Episodes).Count } else { 1 }) `
                -Catalogue $resolved.Catalogue -LegacyResult $legacyResult -Message 'Explicit legacy review or action completed.'
            $runResources.Result = $legacyRunResult
            Complete-PodcastRunProgress -Context $runResources.Progress -Verified:$legacyRunResult.Complete
            return $legacyRunResult
        }

        # Reserve room for both identifiers, separators and an episode's date/extension.
        $folderBudget = [Math]::Min(100, 259 - $baseOutputPath.TrimEnd('\').Length - 2 - 84)
        if ($folderBudget -lt 66) {
            $folderBudget = [Math]::Min(100, 259 - $baseOutputPath.TrimEnd('\').Length - 2 - 70)
        }
        if ($folderBudget -lt 66) { throw 'Output root is too long to retain safe identifiers. Choose a shorter output path.' }
        $archive = Resolve-PodcastArchive -Root $baseOutputPath -FeedTitle $feedTitle -FeedUrl $resolved.Url -MaxFolderLength $folderBudget
        $safeFeedTitle = $archive.FolderName

        $OutputPath = Assert-PodcastDestination -Root $baseOutputPath -RelativePath $safeFeedTitle -Directory
        $episodes = Select-PodcastEpisode -Episodes $allEpisodes -Mode $Mode -CustomCount $CustomCount
        $fileBudget = [Math]::Min(180, 259 - $OutputPath.Length - 1)
        $historyState = $archive.State
        if ($null -eq $historyState) {
            $historyState = New-PodcastHistory -FeedId $archive.FeedId -FeedAliasFingerprint (Get-PodcastNameHash -IdentityKey ('feed:' + $resolved.Url))
        }
        $null = New-PodcastHistoryPlan -Episodes $allEpisodes -State $historyState -MaxFileNameLength $fileBudget
        $destinationPlan = New-PodcastHistoryPlan -Episodes $episodes -State $historyState -MaxFileNameLength $fileBudget

        $legacyReview = Find-PodcastLegacyReview -Root $baseOutputPath -ArchiveRoot $OutputPath -FeedTitle $feedTitle `
            -Episodes $allEpisodes -SelectedEpisodes $episodes -State $historyState -MaxFileNameLength $fileBudget
        if ($null -ne $legacyReview) {
            if ($WhatIfPreference) {
                return New-PodcastRunResult -Mode $Mode -Preview -Planned $destinationPlan.Count `
                    -Catalogue $resolved.Catalogue -LegacyResult $legacyReview -Message 'Local legacy inventory requires review.'
            }
            throw 'Legacy archive requires review; use -LegacyPath with -LegacyAction Preview.'
        }
        if (-not $PSCmdlet.ShouldProcess('Selected podcast archive', 'Update archive history and logs; download selected missing episodes')) {
            # Runtime plan items hold request credentials and mutable history. Return
            # a diagnostic projection instead of serializing those internal objects.
            $publicPlan = @($destinationPlan | ForEach-Object {
                [pscustomobject]@{
                    EpisodeId = $_.EpisodeId
                    IdentitySource = $_.IdentitySource
                    Source = Get-PodcastSafeUrl -Url $_.Episode.Url
                    Recorded = ($null -ne $_.StateRecord)
                }
            })
            return New-PodcastRunResult -Mode $Mode -Preview -Planned $destinationPlan.Count `
                -Catalogue $resolved.Catalogue -Plan $publicPlan -Message 'Archive plan completed without changes.'
        }
        if ($diagnosticConfirmation) { $null = Initialize-PodcastDiagnostics }
        $null = & $onConfirmed
        $null = Invoke-PodcastDestinationPreflight -Root $baseOutputPath
        if (-not (Test-Path -LiteralPath $baseOutputPath)) {
            Write-Host '[*] Creating selected base output directory.'
            $null = Assert-PodcastDestination -Root $baseOutputPath -Directory
            try { $null = [IO.Directory]::CreateDirectory($baseOutputPath) }
            catch {
                if (Test-PodcastCancellation -ErrorObject $_) { throw }
                throw 'The destination is not writable. Check output-folder permissions and retry.'
            }
        }

        # Serialize discovery and first state creation across title changes. This
        # short root lock is released before media transfer; the show lock remains.
        $archiveLock = Enter-PodcastArchiveLock -Root $baseOutputPath
        $archive = Resolve-PodcastArchive -Root $baseOutputPath -FeedTitle $feedTitle -FeedUrl $resolved.Url -MaxFolderLength $folderBudget
        $safeFeedTitle = $archive.FolderName
        $OutputPath = Assert-PodcastDestination -Root $baseOutputPath -RelativePath $safeFeedTitle -Directory
        $fileBudget = [Math]::Min(180, 259 - $OutputPath.Length - 1)
        $historyState = $archive.State
        if ($null -eq $historyState) {
            $historyState = New-PodcastHistory -FeedId $archive.FeedId -FeedAliasFingerprint (Get-PodcastNameHash -IdentityKey ('feed:' + $resolved.Url))
        }
        $destinationPlan = New-PodcastHistoryPlan -Episodes $episodes -State $historyState -MaxFileNameLength $fileBudget
        $legacyReview = Find-PodcastLegacyReview -Root $baseOutputPath -ArchiveRoot $OutputPath -FeedTitle $feedTitle `
            -Episodes $allEpisodes -SelectedEpisodes $episodes -State $historyState -MaxFileNameLength $fileBudget
        if ($null -ne $legacyReview) { throw 'Legacy archive requires review; use -LegacyPath with -LegacyAction Preview.' }
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            Write-Host '[*] Creating podcast folder.'
            $null = Assert-PodcastDestination -Root $baseOutputPath -RelativePath $safeFeedTitle -Directory
            try { $null = [IO.Directory]::CreateDirectory($OutputPath) }
            catch {
                if (Test-PodcastCancellation -ErrorObject $_) { throw }
                throw 'The destination is not writable. Check output-folder permissions and retry.'
            }
        }

        # Lock the established show before logs, ownership or history changes, then
        # reload and replan in case another completed run changed its history.
        $historyLock = Enter-PodcastHistoryLock -Root $OutputPath
        $lockedState = Read-PodcastHistory -Root $OutputPath
        if ($null -ne $lockedState) { $historyState = $lockedState }
        $feedAlias = Get-PodcastNameHash -IdentityKey ('feed:' + $resolved.Url)
        if ($historyState.feed_id -cne $archive.FeedId -or $feedAlias -cnotin $historyState.feed_alias_fingerprints) {
            throw 'Feed identity changed while acquiring archive writer protection; preserving history.'
        }
        $destinationPlan = New-PodcastHistoryPlan -Episodes $episodes -State $historyState -MaxFileNameLength $fileBudget
        $legacyReview = Find-PodcastLegacyReview -Root $baseOutputPath -ArchiveRoot $OutputPath -FeedTitle $feedTitle `
            -Episodes $allEpisodes -SelectedEpisodes $episodes -State $historyState -MaxFileNameLength $fileBudget
        if ($null -ne $legacyReview) { throw 'Legacy archive requires review; use -LegacyPath with -LegacyAction Preview.' }
        $null = Invoke-PodcastDestinationPreflight -Root $OutputPath
        if ($historyState.generation -eq 0) {
            $historyState.generation = 1
            $historyState = Write-PodcastHistory -Lock $historyLock -State $historyState
        }
        $historyContext = @{ Lock = $historyLock; State = $historyState }
        $archiveLock.Dispose()
        $archiveLock = $null

        Set-PodcastDiagnosticArchive -Root $OutputPath
        Write-Log 'UniversalPodcastDownloader log'
        Write-Log ("Feed URL     : {0}" -f (Get-PodcastSafeUrl -Url $FeedUrl))
        Write-Log ("Resolved URL : {0}" -f (Get-PodcastSafeUrl -Url $resolved.Url))
        Write-Log ("Mode         : {0}" -f $Mode)
        Write-Log ("CustomCount  : {0}" -f ($CustomCount -as [string]))
        if (-not $resolved.Catalogue.Complete) { Write-Log (Get-PodcastCatalogueMessage -Catalogue $resolved.Catalogue) 'WARN' }

        Write-Log ("Feed items with valid URLs: {0}" -f $episodeCount)

        $total = $destinationPlan.Count
        $runResources.Progress = New-PodcastProgressContext -TotalEpisodes $total -NonInteractive:$NonInteractive
        Write-Host "[*] Episodes to download: $total"
        Write-Log ("Episodes to download (after mode/filter): {0}" -f $total)

        $downloaded = @()
        $skipped    = @()
        $adopted    = @()
        $failed     = @()

        $index = 0
        foreach ($planned in $destinationPlan) {
            $index++
            $ep = $planned.Episode
            Start-PodcastEpisodeProgress -Context $runResources.Progress -Index $index

            $fileName = $planned.FileName
            $relativeDestination = [IO.Path]::Combine($safeFeedTitle, $fileName)
            $destFile = Assert-PodcastDestination -Root $baseOutputPath -RelativePath $relativeDestination

            $historyAction = Resolve-PodcastHistoryItem -Context $historyContext -Planned $planned
            if ($historyAction -eq 'verified_skip') {
                Write-Host "[-] Skipping (verified history): episode $index"
                $skipped += [PSCustomObject]@{ Title = $ep.Title; File = $destFile }
                $episodeResults += New-PodcastEpisodeResult -EpisodeId $planned.EpisodeId -Outcome verified_skip `
                    -Bytes $planned.StateRecord.bytes -Verification 'local-size-sha256' -Message 'Recorded media verified on disk.'
                Write-Log ("Verified history and on-disk SHA-256: {0}" -f $planned.EpisodeId)
                Complete-PodcastEpisodeProgress -Context $runResources.Progress -Outcome verified_skip
                continue
            }
            if ($historyAction -eq 'adopted_skip') {
                Write-Host "[-] Skipping (owner-adopted local file): episode $index"
                $adopted += [PSCustomObject]@{ Title = $ep.Title; File = $destFile }
                $episodeResults += New-PodcastEpisodeResult -EpisodeId $planned.EpisodeId -Outcome legacy_unverified `
                    -Message 'Owner-adopted local media is unchanged; transfer completeness remains unverified.'
                Write-Log ("Owner-adopted local file is unchanged: {0}; transfer completeness remains unverified." -f $planned.EpisodeId)
                Complete-PodcastEpisodeProgress -Context $runResources.Progress -Outcome legacy_unverified
                continue
            }
            if ($historyAction -eq 'conflict') {
                $message = 'Existing media is unknown or changed; preserved for review.'
                Write-Warning $message
                $failed += [PSCustomObject]@{ EpisodeId = $planned.EpisodeId; Error = $message }
                $episodeResults += New-PodcastEpisodeResult -EpisodeId $planned.EpisodeId -Outcome conflict -Message $message
                Write-Log $message 'ERROR'
                Complete-PodcastEpisodeProgress -Context $runResources.Progress -Outcome conflict
                continue
            }

            Write-Host "[+] Downloading episode $index of $total"
            Write-Verbose ("    URL: " + (Get-PodcastSafeUrl -Url $ep.Url))
            Write-Verbose ("    Episode ID: " + $planned.EpisodeId)

            Write-Log ("Starting download {0}/{1}: {2}" -f $index, $total, $planned.EpisodeId)
            Write-Log ("Source URL : {0}" -f (Get-PodcastSafeUrl -Url $ep.Url))
            if ($ep.PubDate) { Write-Log ("PubDate UTC: {0:yyyy-MM-dd HH:mm:ss}" -f $ep.PubDate) }

            $success   = $false
            $attempt   = 1
            $lastError = $null

            try {
                $transferResult = Invoke-PodcastRecordedTransfer -Context $historyContext -Planned $planned `
                    -Policy $script:PodcastTransportPolicy -OnProgress $onTransferProgress
                $success = $true
                $attempt = $transferResult.Attempts
            } catch {
                if (Test-PodcastCancellation -ErrorObject $_) { throw }
                $lastError = $_
                # Revalidate paths before recording failure. Prepared evidence for
                # an already placed final file remains for the next run to reconcile.
                $null = Assert-PodcastDestination -Root $baseOutputPath -RelativePath $relativeDestination
                $transportFailure = Get-PodcastTransportFailure -ErrorObject $lastError
                if ($null -ne $transportFailure -and $transportFailure.Data.Contains('Attempts')) {
                    $attempt = [int]$transportFailure.Data['Attempts']
                }
                elseif ($lastError.Exception.Data.Contains('PodcastAttempts')) {
                    $attempt = [int]$lastError.Exception.Data['PodcastAttempts']
                }
                $msg = Get-PodcastDiagnosticError -Error $lastError
                Write-Warning ("    Transfer failed: {0}" -f $msg)
                Write-Log ("Transfer failed: {0}" -f $msg) 'WARN'
            }

            if ($success) {
                Write-Host "    Saved episode $index."
                $downloaded += [PSCustomObject]@{ Title = $ep.Title; File = $transferResult.File }
                $episodeResults += New-PodcastEpisodeResult -EpisodeId $planned.EpisodeId -Outcome downloaded `
                    -Bytes $transferResult.Bytes -Verification $transferResult.Verification -Attempts $attempt -Message 'Media transfer validated and recorded.'
                Write-Log ("Download succeeded: {0}" -f $planned.EpisodeId)
                Write-Log ("File size: {0} bytes; validation: {1}" -f $transferResult.Bytes, $transferResult.Verification)
                foreach ($validationWarning in $transferResult.Warnings) {
                    Write-Warning $validationWarning
                    Write-Log $validationWarning 'WARN'
                }
                Complete-PodcastEpisodeProgress -Context $runResources.Progress -Outcome downloaded
            } else {
                $retainedEvidence = @($historyContext.State.episodes | Where-Object {
                    $_.episode_id -ceq $planned.EpisodeId -and $null -ne $_.local_sha256 -and $null -ne $_.bytes
                })
                if ($retainedEvidence.Count -eq 0 -and -not (Test-Path -LiteralPath $destFile)) {
                    $failureRecord = New-PodcastEpisodeRecord -Planned $planned
                    Save-PodcastEpisodeRecord -Context $historyContext -Record $failureRecord
                }
                Write-Warning ("    Giving up after {0} attempts." -f $attempt)
                $errMsg = if ($lastError) { Get-PodcastDiagnosticError -Error $lastError } else { "Unknown error" }
                $failed += [PSCustomObject]@{ EpisodeId = $planned.EpisodeId; Error = $errMsg }
                $outcome = if ($null -ne $transportFailure -and $transportFailure.Data['Kind'] -eq 'Deferred') { 'deferred' } else { 'failed' }
                $episodeResults += New-PodcastEpisodeResult -EpisodeId $planned.EpisodeId -Outcome $outcome -Attempts $attempt -Message $errMsg
                Write-Log ("Giving up after {0} attempts: {1}" -f $attempt, $planned.EpisodeId) 'ERROR'
                Write-Log ("Last error: {0}" -f $errMsg) 'ERROR'
                Complete-PodcastEpisodeProgress -Context $runResources.Progress -Outcome $outcome
            }
        }

        Write-Host ""
        Write-Host "Summary" -ForegroundColor Cyan
        Write-Host "-------"
        Write-Host ("Downloaded : {0}" -f $downloaded.Count)
        Write-Host ("Skipped    : {0}" -f $skipped.Count)
        Write-Host ("Adopted    : {0}" -f $adopted.Count)
        Write-Host ("Failed     : {0}" -f $failed.Count)

        Write-Log ("Summary: Downloaded={0}, Skipped={1}, Failed={2}, Adopted={3}" -f $downloaded.Count, $skipped.Count, $failed.Count, $adopted.Count)

        if ($failed.Count -gt 0) {
            Write-Host ""
            Write-Host "Failed episodes:" -ForegroundColor Yellow
            foreach ($f in $failed) {
                Write-Host (" - {0}  ({1})" -f $f.EpisodeId, $f.Error)
                Write-Log ("Failed: {0} ({1})" -f $f.EpisodeId, $f.Error) 'ERROR'
            }
            # Return the full summary below; episode failures are incomplete, not setup errors.
        }

        $runResult = New-PodcastRunResult -Mode $Mode -Planned $total -EpisodeResults $episodeResults -Catalogue $resolved.Catalogue
        $runResources.Result = $runResult
        Complete-PodcastRunProgress -Context $runResources.Progress -Verified:$runResult.Complete
        if ($runResult.Complete) {
            $runResult.Message = 'Run completed.'
            Write-Log 'Run completed.' 'INFO'
            Write-Host ''
            Write-Host '[OK] Done. Files are in the selected podcast archive.'
        }
        else {
            $runResult.Message = 'Run incomplete; failed, deferred, conflicting or unverified media, or unresolved feed pages remain.'
            Write-Warning $runResult.Message
            Write-Log $runResult.Message 'WARN'
        }
        return $runResult
    }
    catch {
        $primaryError = $_
        $cancelled = Test-PodcastCancellation -ErrorObject $primaryError
        $safeFailure = if ($cancelled) { 'Run cancelled; retained media and recovery evidence were preserved.' } else { Get-PodcastDiagnosticError -Error $primaryError }
        if ($cancelled -and $null -ne $planned -and
            $planned.EpisodeId -cnotin @($episodeResults | ForEach-Object { $_.EpisodeId })) {
            $episodeResults += New-PodcastEpisodeResult -EpisodeId $planned.EpisodeId -Outcome cancelled -Message $safeFailure
        }
        foreach ($remaining in $destinationPlan) {
            if ($remaining.EpisodeId -cnotin @($episodeResults | ForEach-Object { $_.EpisodeId })) {
                $episodeResults += New-PodcastEpisodeResult -EpisodeId $remaining.EpisodeId -Outcome deferred -Message 'Episode was not processed because the run stopped.'
            }
        }
        Write-Host ("ERROR: {0}" -f $safeFailure) -ForegroundColor Red
        Write-Log $safeFailure 'ERROR'
        $catalogue = if ($null -ne $resolved) { $resolved.Catalogue } else { $null }
        $catalogueGap = $primaryError.Exception.Data['PodcastCatalogue']
        if ($null -ne $catalogueGap) { $catalogue = $catalogueGap }
        $failureRunResult = New-PodcastRunResult -Mode $Mode -Preview:$diagnosticPreview -Planned $destinationPlan.Count `
            -EpisodeResults $episodeResults -Catalogue $catalogue -Fatal:($null -eq $catalogueGap -and -not $cancelled) `
            -Cancelled:$cancelled -Message $safeFailure
        $runResources.Result = $failureRunResult
        return $failureRunResult
    }
    finally {
        try { Close-PodcastProgress -Context $runResources.Progress }
        catch {
            if ((Test-PodcastCancellation -ErrorObject $_) -and $null -ne $runResources.Result -and
                $runResources.Result.ExitCode -eq 0 -and -not $runResources.Result.Preview) {
                # All owned display IDs have been attempted. Keep committed
                # episode evidence and continue power/lock cleanup, but expose
                # a new catchable cancellation after otherwise successful work.
                $runResources.Result.ExitCode = 130
                $runResources.Result.Status = 'cancelled'
                $runResources.Result.Complete = $false
                $runResources.Result.Message = 'Run cancelled while clearing progress; completed media was preserved.'
                Write-PodcastDiagnosticFallback -Message $runResources.Result.Message
            }
            else { Write-PodcastDiagnosticFallback -Message 'Progress display cleanup failed; the operation result is preserved.' }
        }
        try {
            $powerCleanup = Stop-PodcastKeepAwake -Context $runResources.Power
            if ($null -ne $runResources.Power -and -not $powerCleanup.Restored) {
                Write-PodcastDiagnosticFallback -Message $powerCleanup.Message
            }
        }
        catch { Write-PodcastDiagnosticFallback -Message 'Temporary keep-awake cleanup failed; the operation result is preserved.' }
        try { if ($null -ne $historyLock) { $historyLock.Stream.Dispose() } }
        catch { Write-PodcastDiagnosticFallback -Message 'History lock cleanup failed; the operation result is preserved.' }
        try { if ($null -ne $archiveLock) { $archiveLock.Dispose() } }
        catch { Write-PodcastDiagnosticFallback -Message 'Archive lock cleanup failed; the operation result is preserved.' }
        if ($DiagnosticExportPath) {
            # Diagnostic export is optional and best effort, including on failure.
            try { Export-PodcastDiagnostics -Path $DiagnosticExportPath }
            catch { Write-PodcastDiagnosticFallback -Message 'Diagnostic export is unavailable; the original operation result is preserved.' }
        }
        Close-PodcastDiagnostics
    }
}

# Importing defines the API without prompts, writes, downloads or process exit.
if ($MyInvocation.InvocationName -eq '.') { return }
$launcherGuided = $env:UPD_LAUNCHER -eq '1' -and $PSBoundParameters.Count -eq 0
$runArguments = @{}
foreach ($key in $PSBoundParameters.Keys) {
    if ($key -ne 'PassThru') { $runArguments[$key] = $PSBoundParameters[$key] }
}
$runResult = Invoke-PodcastRun @runArguments
if ($PassThru) { $runResult }
if ($launcherGuided) {
    try { $null = Read-Host 'Press Enter to close' }
    catch { Write-PodcastDiagnosticFallback -Message 'The closing prompt is unavailable; the operation result is preserved.' }
}
exit $runResult.ExitCode
