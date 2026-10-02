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
            - all    = whole feed
    - Per-podcast subfolders based on feed title:
        - Root:   <OutputPath> (default: %USERPROFILE%\Downloads\Podcasts)
        - Folder: <SafeFeedTitle>-<feed identity hash>
        - Files:  YYYY-MM-DD - Episode title-<episode identity hash>.mp3
    - Robust download loop:
        - Up to 3 attempts per episode with short delay between tries.
        - Skips episodes only after checking recorded size and SHA-256 on disk.
        - Summarizes downloaded / skipped / failed at the end.
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
                 - all    = entire feed
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
        - Use when you already know the RSS URL and desired mode:
            .\UniversalPodcastDownloader.ps1 
                -FeedUrl "https://example.com/feed.xml" 
                -Mode All 
                -OutputPath "D:\Podcasts"
        - Or: download the 10 newest episodes:
            .\UniversalPodcastDownloader.ps1 
                -FeedUrl "https://example.com/feed.xml" 
                -Mode Custom 
                -CustomCount 10

NOTES
    - Episodes are sorted by publication date (PubDate) newest first.
    - New folder and file names include deterministic SHA-256 identity suffixes.
    - Windows names are sanitized and shortened to fit the selected output root.
    - Unsafe archive paths fail before archive writes; startup diagnostics use a separate safe location.
    - Recorded transfer skips check local size and SHA-256. Adopted files stay distinct.
    - Use -LegacyPath for a read-only inventory and explicit adoption/redownload choices.
    - The tool does not transcode or modify audio, it saves whatever the feed serves,
      but uses the .mp3 extension by default for naming.

LIMITATIONS
    - No resume of partially downloaded files, failed downloads are retried from scratch.
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

    [string]$LegacyPath,

    [ValidateSet('Preview','Adopt','Redownload','Rollback')]
    [string]$LegacyAction = 'Preview',

    [string]$LegacyEpisodeId,

    [string]$LegacyFile,

    [string]$LegacySha256,

    [string]$LegacyCheckpoint,

    [string]$DiagnosticExportPath,

    [ValidateRange(1, 10)][int]$MaxAttempts = 3,
    [ValidateRange(0.001, 86400)][double]$HeaderTimeoutSeconds = 30,
    [ValidateRange(0.001, 86400)][double]$IdleTimeoutSeconds = 30,
    [ValidateRange(0, 86400)][double]$RetryBudgetSeconds = 120,
    [ValidateRange(0, 3600)][double]$BaseDelaySeconds = 1,
    [ValidateRange(0, 3600)][double]$MaxDelaySeconds = 30
)

. (Join-Path $PSScriptRoot 'src/Naming.ps1')
. (Join-Path $PSScriptRoot 'src/PathSafety.ps1')
. (Join-Path $PSScriptRoot 'src/Diagnostics.ps1')
. (Join-Path $PSScriptRoot 'src/HistoryStore.ps1')
. (Join-Path $PSScriptRoot 'src/HistoryIdentity.ps1')
. (Join-Path $PSScriptRoot 'src/HistoryWorkflow.ps1')
. (Join-Path $PSScriptRoot 'src/NetworkPolicy.ps1')
. (Join-Path $PSScriptRoot 'src/FeedXml.ps1')
. (Join-Path $PSScriptRoot 'src/FeedDiscovery.ps1')
. (Join-Path $PSScriptRoot 'src/MediaRequest.ps1')
. (Join-Path $PSScriptRoot 'src/MediaValidation.ps1')
. (Join-Path $PSScriptRoot 'src/ResumeStore.ps1')
. (Join-Path $PSScriptRoot 'src/MediaTransfer.ps1')
. (Join-Path $PSScriptRoot 'src/LegacyInventory.ps1')
. (Join-Path $PSScriptRoot 'src/LegacyMigration.ps1')

$script:PodcastTransportPolicy = New-PodcastTransportPolicy -MaxAttempts $MaxAttempts `
    -HeaderTimeoutSeconds $HeaderTimeoutSeconds -IdleTimeoutSeconds $IdleTimeoutSeconds `
    -RetryBudgetSeconds $RetryBudgetSeconds -BaseDelaySeconds $BaseDelaySeconds -MaxDelaySeconds $MaxDelaySeconds

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
    param([string[]]$Feeds, [switch]$Interactive, [AllowNull()]$InitialResolution)

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
            if (@($feed.Items).Count -eq 0) { throw 'The RSS or Atom feed is valid but contains no episodes.' }
            return [pscustomobject]@{
                Url = $feed.Url
                Xml = $feed.Xml
                Items = @($feed.Items)
                Kind = $feed.Kind
                FinalUri = $feed.FinalUri
                Content = $feed.Content
                Candidates = $candidates
                SourceUrl = $source.Url
            }
        } catch {
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

    $dateStr = Get-XPathText -Node $XmlItem -XPath './*[local-name()="pubDate"][1]'
    if (-not $dateStr) { $dateStr = Get-XPathText -Node $XmlItem -XPath './*[local-name()="updated"][1]' }
    if (-not $dateStr) { $dateStr = Get-XPathText -Node $XmlItem -XPath './*[local-name()="published"][1]' }
    if (-not $dateStr) { $dateStr = Get-FirstText $XmlItem.pubDate }

    $pubDate = $null
    if ($dateStr) { try { $pubDate = [datetime]::Parse($dateStr) } catch {} }

    $guid = Get-XPathText -Node $XmlItem -XPath './*[local-name()="guid"][1]'
    if (-not $guid) { $guid = Get-FirstText $XmlItem.guid }
    $atomId = Get-XPathText -Node $XmlItem -XPath './*[local-name()="id"][1]'

    $url = $null
    $mediaNode = $null
    $enc = $XmlItem.SelectSingleNode('./*[local-name()="enclosure"][1]')
    if ($enc) {
        $attr = $enc.Attributes["url"]
        if ($attr) { $url = $attr.Value }
        if (-not $url) { $url = $enc.GetAttribute("url") }
        if ($url) { $mediaNode = $enc }
    }

    if (-not $url) {
        $ln = $XmlItem.SelectSingleNode('./*[local-name()="link" and @rel="enclosure"][1]')
        if ($ln) {
            $url = $ln.GetAttribute("href")
            if ($url) { $mediaNode = $ln }
        }
    }

    if (-not $url) {
        $cands = @(
            Get-XPathText -Node $XmlItem -XPath './*[local-name()="guid"][1]'
            Get-XPathText -Node $XmlItem -XPath './*[local-name()="link"][1]'
            (Get-FirstText $XmlItem.guid)
            (Get-FirstText $XmlItem.link)
        ) | Where-Object { $_ }

        foreach ($cand in $cands) {
            if ($cand -match '\.(mp3|m4a)($|\?)') { $url = "$cand"; break }
        }
    }

    $enclosureLength = $null
    if ($mediaNode) {
        $parsedLength = 0L
        if ([long]::TryParse($mediaNode.GetAttribute('length'), [Globalization.NumberStyles]::None,
                [Globalization.CultureInfo]::InvariantCulture, [ref]$parsedLength)) {
            $enclosureLength = $parsedLength
        }
    }

    [PSCustomObject]@{
        Title   = ($title -as [string])
        PubDate = $pubDate
        Url     = $url
        Guid    = ($guid -as [string])
        AtomId  = ($atomId -as [string])
        EnclosureLength = $enclosureLength
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

    $selected = @($Episodes | Sort-Object PubDate -Descending)
    switch ($Mode) {
        'Latest' { $selected = @($selected | Select-Object -First 1) }
        'Custom' { $selected = @($selected | Select-Object -First $CustomCount) }
    }

    # Preserve an array for zero/one/many results when assigned by the caller.
    return ,$selected
}

# Dot-sourcing exposes helpers for tests without starting the downloader.
if ($MyInvocation.InvocationName -eq '.') { return }

# --- Global config ---
$ErrorActionPreference = 'Stop'
$prevProgress = $global:ProgressPreference
$global:ProgressPreference = 'Continue'

$historyLock = $null
$archiveLock = $null
$legacyRequested = $PSBoundParameters.ContainsKey('LegacyPath')
$diagnosticPreview = $WhatIfPreference -or ($legacyRequested -and $LegacyAction -eq 'Preview')
$diagnosticConfirmation = $ConfirmPreference -in @('Low', 'Medium')
$null = Initialize-PodcastDiagnostics -Preview:($diagnosticPreview -or $diagnosticConfirmation)

# --- Main ---
try {
    foreach ($option in @('LegacyAction', 'LegacyEpisodeId', 'LegacyFile', 'LegacySha256', 'LegacyCheckpoint')) {
        if ($PSBoundParameters.ContainsKey($option) -and -not $legacyRequested) {
            throw 'Legacy options require -LegacyPath and an explicit -FeedUrl.'
        }
    }
    if ($legacyRequested -and ([string]::IsNullOrWhiteSpace($LegacyPath) -or
            -not $PSBoundParameters.ContainsKey('FeedUrl') -or [string]::IsNullOrWhiteSpace($FeedUrl))) {
        throw 'Legacy review requires an existing -LegacyPath and an explicit -FeedUrl.'
    }
    $needFeed  = -not $PSBoundParameters.ContainsKey('FeedUrl')
    $needCount = -not $legacyRequested -and -not $PSBoundParameters.ContainsKey('Mode') -and -not $PSBoundParameters.ContainsKey('CustomCount')

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
        Write-Host "  - Type 'all' to download everything from the feed."
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
    $episodes = @($episodes | Where-Object { $_.Url })
    $episodeCount = $episodes.Count
    if ($episodeCount -eq 0) {
        throw 'Feed parsed, but no downloadable enclosure URLs were found.'
    }
    # Validate enclosure targets during planning, before any archive write or
    # media request. The transport validates the original target and every hop again.
    foreach ($episode in $episodes) { $null = Get-PodcastRequestUri -Uri $episode.Url }
    $allEpisodes = $episodes
    if ($legacyRequested) {
        Invoke-PodcastLegacyMigration -Root $baseOutputPath -LegacyRoot $LegacyPath -FeedUrl $resolved.Url `
            -Episodes $allEpisodes -Action $LegacyAction -EpisodeId $LegacyEpisodeId -FileName $LegacyFile `
            -Sha256 $LegacySha256 -Checkpoint $LegacyCheckpoint -Policy $script:PodcastTransportPolicy
        return
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
        if ($WhatIfPreference) { return $legacyReview }
        throw 'Legacy archive requires review; use -LegacyPath with -LegacyAction Preview.'
    }
    if (-not $PSCmdlet.ShouldProcess('Selected podcast archive', 'Update archive history and logs; download selected missing episodes')) {
        # Runtime plan items hold request credentials and mutable history. Return
        # a diagnostic projection instead of serializing those internal objects.
        return @($destinationPlan | ForEach-Object {
            [pscustomobject]@{
                EpisodeId = $_.EpisodeId
                IdentitySource = $_.IdentitySource
                Source = Get-PodcastSafeUrl -Url $_.Episode.Url
                Recorded = ($null -ne $_.StateRecord)
            }
        })
    }
    if ($diagnosticConfirmation) { $null = Initialize-PodcastDiagnostics }
    if (-not (Test-Path -LiteralPath $baseOutputPath)) {
        Write-Host '[*] Creating selected base output directory.'
        $null = Assert-PodcastDestination -Root $baseOutputPath -Directory
        $null = [IO.Directory]::CreateDirectory($baseOutputPath)
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
        $null = [IO.Directory]::CreateDirectory($OutputPath)
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

    Write-Log ("Feed items with valid URLs: {0}" -f $episodeCount)

    $total = $destinationPlan.Count
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

        $fileName = $planned.FileName
        $relativeDestination = [IO.Path]::Combine($safeFeedTitle, $fileName)
        $destFile = Assert-PodcastDestination -Root $baseOutputPath -RelativePath $relativeDestination

        $historyAction = Resolve-PodcastHistoryItem -Context $historyContext -Planned $planned
        if ($historyAction -eq 'verified_skip') {
            Write-Progress -Activity "Podcast downloads" -Status "Skipping (verified history): episode $index" `
                -PercentComplete ([int](($index/$total)*100)) -CurrentOperation "Episode $index of $total"
            Write-Host "[-] Skipping (verified history): episode $index"
            $skipped += [PSCustomObject]@{ Title = $ep.Title; File = $destFile }
            Write-Log ("Verified history and on-disk SHA-256: {0}" -f $planned.EpisodeId)
            continue
        }
        if ($historyAction -eq 'adopted_skip') {
            Write-Progress -Activity "Podcast downloads" -Status "Skipping (owner-adopted local file): episode $index" `
                -PercentComplete ([int](($index/$total)*100)) -CurrentOperation "Episode $index of $total"
            Write-Host "[-] Skipping (owner-adopted local file): episode $index"
            $adopted += [PSCustomObject]@{ Title = $ep.Title; File = $destFile }
            Write-Log ("Owner-adopted local file is unchanged: {0}; transfer completeness remains unverified." -f $planned.EpisodeId)
            continue
        }
        if ($historyAction -eq 'conflict') {
            $message = 'Existing media is unknown or changed; preserved for review.'
            Write-Warning $message
            $failed += [PSCustomObject]@{ EpisodeId = $planned.EpisodeId; Error = $message }
            Write-Log $message 'ERROR'
            continue
        }

        $pct = [int](($index/$total)*100)
        Write-Progress -Activity "Podcast downloads" -Status "Preparing: episode $index" -PercentComplete $pct `
            -CurrentOperation "Episode $index of $total"

        Write-Host "[+] Downloading episode $index of $total"
        Write-Verbose ("    URL: " + (Get-PodcastSafeUrl -Url $ep.Url))
        Write-Verbose ("    Episode ID: " + $planned.EpisodeId)

        Write-Log ("Starting download {0}/{1}: {2}" -f $index, $total, $planned.EpisodeId)
        Write-Log ("Source URL : {0}" -f (Get-PodcastSafeUrl -Url $ep.Url))
        if ($ep.PubDate) { Write-Log ("PubDate   : {0:yyyy-MM-dd HH:mm:ss}" -f $ep.PubDate) }

        $success   = $false
        $attempt   = 1
        $lastError = $null

        try {
            $transferResult = Invoke-PodcastRecordedTransfer -Context $historyContext -Planned $planned -Policy $script:PodcastTransportPolicy
            $success = $true
            $attempt = $transferResult.Attempts
        } catch {
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
            $downloaded += [PSCustomObject]@{ Title = $ep.Title; File = $destFile }
            Write-Log ("Download succeeded: {0}" -f $planned.EpisodeId)
            Write-Log ("File size: {0} bytes; validation: {1}" -f $transferResult.Bytes, $transferResult.Verification)
            foreach ($validationWarning in $transferResult.Warnings) {
                Write-Warning $validationWarning
                Write-Log $validationWarning 'WARN'
            }
        } else {
            if (-not (Test-Path -LiteralPath $destFile)) {
                $failureRecord = New-PodcastEpisodeRecord -Planned $planned
                Save-PodcastEpisodeRecord -Context $historyContext -Record $failureRecord
            }
            Write-Warning ("    Giving up after {0} attempts." -f $attempt)
            $errMsg = if ($lastError) { Get-PodcastDiagnosticError -Error $lastError } else { "Unknown error" }
            $failed += [PSCustomObject]@{ EpisodeId = $planned.EpisodeId; Error = $errMsg }
            Write-Log ("Giving up after {0} attempts: {1}" -f $attempt, $planned.EpisodeId) 'ERROR'
            Write-Log ("Last error: {0}" -f $errMsg) 'ERROR'
        }
    }

    Write-Progress -Activity "Podcast downloads" -Completed

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
        throw ('Download incomplete: {0} episode(s) failed or require review; existing files were preserved.' -f $failed.Count)
    }

    Write-Log "Run completed." 'INFO'

    Write-Host ""
    Write-Host '[OK] Done. Files are in the selected podcast archive.'
}
catch {
    $primaryError = $_
    $safeFailure = Get-PodcastDiagnosticError -Error $primaryError
    Write-Host ("ERROR: {0}" -f $safeFailure) -ForegroundColor Red
    Write-Log ("Fatal error: {0}" -f $safeFailure) 'ERROR'
    # Keep the exception instance for callers. Public formatting must not reveal
    # its raw message, inner exceptions, target, or original invocation arguments.
    $publicError = [Management.Automation.ErrorRecord]::new($primaryError.Exception,
        'PodcastRunFailed', [Management.Automation.ErrorCategory]::NotSpecified, $null)
    $publicError.ErrorDetails = [Management.Automation.ErrorDetails]::new($safeFailure)
    $PSCmdlet.ThrowTerminatingError($publicError)
}
finally {
    if ($null -ne $historyLock) { $historyLock.Stream.Dispose() }
    if ($null -ne $archiveLock) { $archiveLock.Dispose() }
    if ($DiagnosticExportPath) {
        # Diagnostic export is optional and best effort, including on failure.
        try { Export-PodcastDiagnostics -Path $DiagnosticExportPath }
        catch { Write-PodcastDiagnosticFallback -Message 'Diagnostic export is unavailable; the original operation result is preserved.' }
    }
    Close-PodcastDiagnostics
    $global:ProgressPreference = $prevProgress
}
