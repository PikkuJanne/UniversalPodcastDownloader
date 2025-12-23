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
        - Folder: <FeedTitle> (sanitized)
        - Files:  YYYY-MM-DD - Episode title.mp3
    - Robust download loop:
        - Up to 3 attempts per episode with short delay between tries.
        - Skips episodes where the target file already exists.
        - Summarizes downloaded / skipped / failed at the end.
    - Verbose, per-run log file:
        - Stored in the podcast folder next to the audio files.
        - Named: YYYYMMDD_HHMMSS_<FeedTitle>.log
        - Logs:
            - Feed URL and resolved feed URL
            - Mode, output folder, and counts
            - Each episode download attempt and outcome
            - Final summary, including failures and error messages

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
    2) Optional: pin the .bat to Start or Taskbar for quick access.
    3) Ensure the machine has:
         - Working Internet connection.
         - Permission to write to the default output:
             %USERPROFILE%\Downloads\Podcasts
           or whichever OutputPath you configure.
    4) No external binaries required. Relies on Invoke-WebRequest
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
                YYYYMMDD_HHMMSS_<FeedTitle>.log

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
    - Files are named “YYYY-MM-DD - Title.mp3” when a date is available,
      otherwise “Title.mp3”.
    - Filename and folder names are sanitized for Windows (no : * ? " < > | etc.).
    - The script always logs to the podcast folder for each run, even on failures.
    - Existing files are never overwritten, they are just skipped and logged as such.
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

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Latest','All','Custom')]
    [string]$Mode = 'Latest',

    [int]$CustomCount,

    [string]$OutputPath = "$env:USERPROFILE\Downloads\Podcasts",

    [string]$FeedUrl
)

# --- Global config ---
$ErrorActionPreference = 'Stop'
$prevProgress = $global:ProgressPreference
$global:ProgressPreference = 'Continue'

# Log file path, set later once we know the podcast folder
$script:LogFile = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message

    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line
    }
}

# --- Helpers ---
function Find-RssInHtml {
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][string]$BaseUrl
    )

    # Look for: <link ... type="application/rss+xml" ... href="...">
    $pattern = '<link[^>]+type=["'']application/(rss|atom)\+xml["''][^>]*>'
    $matches = [System.Text.RegularExpressions.Regex]::Matches(
        $Html,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    foreach ($m in $matches) {
        $hrefMatch = [System.Text.RegularExpressions.Regex]::Match(
            $m.Value,
            'href=["''](?<url>[^"\'']+)["'']',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )

        if ($hrefMatch.Success) {
            $href = $hrefMatch.Groups['url'].Value
            try {
                $base = [Uri]$BaseUrl
                $uri  = [Uri]::new($base, $href)
                return $uri.AbsoluteUri
            } catch {
                # If Uri combine fails, fall back to raw href
                return $href
            }
        }
    }

    return $null
}

function Get-FeedUrlInteractive {
    Write-Host "==== Universal Podcast Downloader ====" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "How to find the RSS feed for your podcast:" -ForegroundColor Yellow
    Write-Host "  1. Open the podcast's main page in your browser."
    Write-Host "  2. Look for an icon or link labelled 'RSS', 'Feed', or 'Subscribe'."
    Write-Host "  3. Copy that link (often ends with .xml or /feed)."
    Write-Host ""
    Write-Host "You can also paste a normal podcast web page URL (show page),"
    Write-Host "and I'll TRY to auto-detect the RSS feed from there."
    Write-Host ""

    while ($true) {
        $inputUrl = Read-Host "Paste RSS feed URL OR podcast page URL"
        if ([string]::IsNullOrWhiteSpace($inputUrl)) {
            Write-Host "Please paste a URL (or press Ctrl+C to quit)." -ForegroundColor Red
            continue
        }

        if ($inputUrl -notmatch '^https?://') {
            Write-Host "Note: URLs usually start with http:// or https://. Continuing anyway..." -ForegroundColor DarkYellow
        }

        try {
            Write-Host "  Fetching URL..." -ForegroundColor DarkCyan
            $resp = Invoke-WebRequest -Uri $inputUrl
            $html = $resp.Content

            # If the content already looks like RSS/Atom, assume it's a direct feed
            if ($html -match '<rss' -or $html -match '<feed') {
                Write-Host "  This looks like a direct RSS/Atom feed." -ForegroundColor Green
                return $inputUrl
            }

            # Otherwise, try to detect a <link type="application/rss+xml"...> tag
            $rssUrl = Find-RssInHtml -Html $html -BaseUrl $inputUrl
            if ($rssUrl) {
                Write-Host "  Found RSS candidate: $rssUrl" -ForegroundColor Green
                $ans = Read-Host "Use this feed? (Y/n)"
                if ($ans -match '^(n|no)$') {
                    Write-Host "  Okay, let's try another URL." -ForegroundColor Yellow
                    continue
                }
                return $rssUrl
            }

            Write-Host "  Could not auto-detect an RSS feed on that page." -ForegroundColor Red
            Write-Host "  Try a different URL or copy a direct 'RSS' link." -ForegroundColor Yellow
        } catch {
            Write-Host "  Failed to fetch URL: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

function Resolve-PodcastItems {
    param([string[]]$Feeds)

    foreach ($u in $Feeds) {
        Write-Verbose "Trying feed: $u"
        try {
            $resp = Invoke-WebRequest -Uri $u
            $xml  = [xml]$resp.Content

            # Try RSS first
            $items = $xml.SelectNodes('//rss/channel/item')
            if (-not $items -or $items.Count -eq 0) {
                # Try Atom fallback, items are usually <entry>
                $items = $xml.SelectNodes('//feed/entry')
            }

            if ($items -and $items.Count -gt 0) {
                return [PSCustomObject]@{ Url = $u; Xml = $xml; Items = $items }
            }
        } catch {
            Write-Verbose "Feed failed: $u  ($($_.Exception.Message))"
        }
    }

    throw "No episodes found in the feed. Double-check the RSS URL."
}

function Get-EpisodeData {
    param($XmlItem)

    # Title
    $title = $null
    $titleNode = $XmlItem.title
    if ($titleNode) {
        if ($titleNode -is [string]) {
            $title = $titleNode
        } else {
            $title = $titleNode.InnerText
        }
    }

    # pubDate (RSS) or updated/published (Atom)
    $dateStr = $null
    if ($XmlItem.pubDate) {
        $dateStr = $XmlItem.pubDate
    }
    if (-not $dateStr) {
        $updatedNode = $XmlItem.SelectSingleNode('updated')
        if ($updatedNode) { $dateStr = $updatedNode.InnerText }
    }
    if (-not $dateStr) {
        $publishedNode = $XmlItem.SelectSingleNode('published')
        if ($publishedNode) { $dateStr = $publishedNode.InnerText }
    }

    $pubDate = $null
    if ($dateStr) {
        try { $pubDate = [datetime]::Parse($dateStr) } catch {}
    }

    # Enclosure URL (RSS)
    $url = $null
    $enc = $XmlItem.SelectSingleNode('enclosure')
    if ($enc) {
        $url = $enc.GetAttribute('url')
        if (-not $url -and $enc.url) { $url = $enc.url }
    }

    # Atom enclosure alternative: <link rel="enclosure" href="...">
    if (-not $url) {
        $links = $XmlItem.SelectNodes('link')
        foreach ($ln in $links) {
            $rel = $ln.GetAttribute('rel')
            if ($rel -and $rel -eq 'enclosure') {
                $href = $ln.GetAttribute('href')
                if ($href) { $url = $href; break }
            }
        }
    }

    # Fallback: some feeds put the MP3 in <guid> or a normal <link>
    if (-not $url) {
        $guid = $XmlItem.guid
        $link = $XmlItem.link
        foreach ($cand in @($guid, $link)) {
            if ($cand -and ($cand.ToString() -match '\.mp3($|\?)')) {
                $url = $cand.ToString()
                break
            }
        }
    }

    [PSCustomObject]@{
        Title   = "$title"
        PubDate = $pubDate
        Url     = $url
    }
}

try {
    # --- TUI decisions ---
    $needFeed  = -not $PSBoundParameters.ContainsKey('FeedUrl')
    $needCount = -not $PSBoundParameters.ContainsKey('Mode') -and -not $PSBoundParameters.ContainsKey('CustomCount')

    if ($needFeed) {
        $FeedUrl = Get-FeedUrlInteractive
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

    if ($Mode -eq 'Custom' -and (-not $CustomCount -or $CustomCount -lt 1)) {
        throw "Mode 'Custom' requires -CustomCount with a value >= 1."
    }

    # --- Base output directory ---
    $baseOutputPath = $OutputPath
    if (-not (Test-Path -LiteralPath $baseOutputPath)) {
        Write-Host "[*] Creating base output directory: $baseOutputPath"
        New-Item -ItemType Directory -Path $baseOutputPath -Force | Out-Null
    }

    # --- Fetch & parse feed ---
    $candidateFeeds = @($FeedUrl) | Where-Object { $_ } | Select-Object -Unique

    Write-Host "[*] Fetching podcast feed..."
    Write-Log ("Fetching feed from URL: {0}" -f $FeedUrl)
    $resolved = Resolve-PodcastItems -Feeds $candidateFeeds
    Write-Host "    Using feed: $($resolved.Url)"
    Write-Log ("Resolved feed URL: {0}" -f $resolved.Url)

    # Feed title, for subfolder
    $feedTitle = $null
    try {
        if ($resolved.Xml.rss -and $resolved.Xml.rss.channel) {
            $feedTitle = $resolved.Xml.rss.channel.title
        }
        if (-not $feedTitle -and $resolved.Xml.feed) {
            $feedTitle = $resolved.Xml.feed.title
        }
    } catch {}

    if ($feedTitle) {
        Write-Host "    Feed title: $feedTitle"
    } else {
        Write-Host "    Feed title: (unknown)"
    }

    # --- Per-podcast subfolder ---
    $safeFeedTitle = $feedTitle
    if (-not $safeFeedTitle) { $safeFeedTitle = 'UnknownPodcast' }
    $safeFeedTitle = ($safeFeedTitle -replace '[\\\/\:\*\?\"<>\|]', '_').Trim()
    if (-not $safeFeedTitle) { $safeFeedTitle = 'Podcast' }

    $OutputPath = Join-Path $baseOutputPath $safeFeedTitle
    if (-not (Test-Path -LiteralPath $OutputPath)) {
        Write-Host "[*] Creating podcast folder: $OutputPath"
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    # --- Prepare log file in podcast folder ---
    $dateStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile = Join-Path $OutputPath ("{0}_{1}.log" -f $dateStamp, $safeFeedTitle)

    Set-Content -LiteralPath $script:LogFile -Value "UniversalPodcastDownloader log"
    Write-Log ("Feed URL     : {0}" -f $FeedUrl)
    Write-Log ("Feed title   : {0}" -f ($feedTitle -as [string]))
    Write-Log ("Mode         : {0}" -f $Mode)
    Write-Log ("CustomCount  : {0}" -f ($CustomCount -as [string]))
    Write-Log ("Output folder: {0}" -f $OutputPath)

    # --- Build episode list ---
    $episodes = foreach ($it in $resolved.Items) { Get-EpisodeData $it }

    # Filter out any items with no URL
    $episodes = $episodes | Where-Object { $_.Url }

    Write-Log ("Feed items with valid URLs: {0}" -f $episodes.Count)

    if (-not $episodes -or $episodes.Count -eq 0) {
        throw "Feed parsed, but no downloadable enclosure URLs were found."
    }

    # Sort newest first
    $episodes = $episodes | Sort-Object PubDate -Descending

    switch ($Mode) {
        'Latest' {
            $episodes = $episodes | Select-Object -First 1
        }
        'All' {
            # leave as-is
        }
        'Custom' {
            $episodes = $episodes | Select-Object -First $CustomCount
        }
    }

    $total = $episodes.Count
    Write-Host "[*] Episodes to download: $total"
    Write-Log ("Episodes to download (after mode/filter): {0}" -f $total)

    if ($PSBoundParameters.ContainsKey('Verbose')) {
        $episodes | ForEach-Object {
            Write-Verbose (" - {0:yyyy-MM-dd}  {1}" -f $_.PubDate, $_.Title)
        }
    }

    # --- Download loop with robustness & summary ---
    $downloaded = @()
    $skipped    = @()
    $failed     = @()
    $maxRetries = 3

    $index = 0
    foreach ($ep in $episodes) {
        $index++

        # Build a safe filename
        $safeTitle = ($ep.Title -replace '[\\\/\:\*\?\"<>\|]', '_').Trim()
        if (-not $safeTitle) { $safeTitle = 'Episode' }
        if ($ep.PubDate) {
            $fileName = '{0:yyyy-MM-dd} - {1}.mp3' -f $ep.PubDate, $safeTitle
        } else {
            $fileName = "$safeTitle.mp3"
        }
        $destFile  = Join-Path $OutputPath $fileName

        if (Test-Path -LiteralPath $destFile) {
            Write-Progress -Activity "Podcast downloads" -Status "Skipping (exists): $fileName" `
                -PercentComplete ([int](($index/$total)*100)) -CurrentOperation "Episode $index of $total"
            Write-Host "[-] Skipping (already exists): $fileName"
            $skipped += [PSCustomObject]@{
                Title = $ep.Title
                File  = $destFile
            }
            Write-Log ("Skipping existing file: {0}" -f $destFile)
            continue
        }

        $pct = [int](($index/$total)*100)
        Write-Progress -Activity "Podcast downloads" -Status "Preparing: $fileName" -PercentComplete $pct `
            -CurrentOperation "Episode $index of $total"
        Write-Host "[+] Downloading ($index/$total): $($ep.Title)"
        Write-Verbose "    URL: $($ep.Url)"
        Write-Verbose "    OUT: $destFile"

        Write-Log ("Starting download {0}/{1}: {2}" -f $index, $total, $ep.Title)
        Write-Log ("Target file: {0}" -f $destFile)
        Write-Log ("Source URL : {0}" -f $ep.Url)

        $success   = $false
        $attempt   = 0
        $lastError = $null

        while (-not $success -and $attempt -lt $maxRetries) {
            $attempt++
            Write-Log ("Attempt {0} of {1} for: {2}" -f $attempt, $maxRetries, $ep.Title)

            try {
                Invoke-WebRequest -Uri $ep.Url -OutFile $destFile
                $success = $true
            } catch {
                $lastError = $_
                Write-Warning ("    Attempt {0} of {1} failed: {2}" -f $attempt, $maxRetries, $lastError.Exception.Message)
                Write-Log ("Attempt {0} failed: {1}" -f $attempt, $lastError.Exception.Message) 'WARN'
                if ($attempt -lt $maxRetries) {
                    Start-Sleep -Seconds 3
                }
            }
        }

        if ($success) {
            Write-Host "    Saved: $fileName"
            $downloaded += [PSCustomObject]@{
                Title = $ep.Title
                File  = $destFile
            }
            Write-Log ("Download succeeded: {0}" -f $destFile)
        } else {
            Write-Warning "    Giving up after $maxRetries attempts."
            $failed += [PSCustomObject]@{
                Title = $ep.Title
                File  = $destFile
                Error = if ($lastError) { $lastError.Exception.Message } else { "Unknown error"
                }
            }
            Write-Log ("Giving up after {0} attempts: {1}" -f $maxRetries, $ep.Title) 'ERROR'
        }
    }

    Write-Progress -Activity "Podcast downloads" -Completed

    # --- Summary ---
    Write-Host ""
    Write-Host "Summary" -ForegroundColor Cyan
    Write-Host "-------"
    Write-Host ("Downloaded : {0}" -f $downloaded.Count)
    Write-Host ("Skipped    : {0}" -f $skipped.Count)
    Write-Host ("Failed     : {0}" -f $failed.Count)

    Write-Log ("Summary: Downloaded={0}, Skipped={1}, Failed={2}" -f $downloaded.Count, $skipped.Count, $failed.Count)

    if ($failed.Count -gt 0) {
        Write-Host ""
        Write-Host "Failed episodes:" -ForegroundColor Yellow
        foreach ($f in $failed) {
            Write-Host (" - {0}  ({1})" -f $f.Title, $f.Error)
            Write-Log ("Failed: {0} ({1})" -f $f.Title, $f.Error) 'ERROR'
        }
    }

    Write-Log "Run completed." 'INFO'

    Write-Host ""
    Write-Host "[OK] Done. Files in: $OutputPath"
}
catch {
    Write-Host ("ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Log ("Fatal error: {0}" -f $_.Exception.Message) 'ERROR'
    throw
}
finally {
    $global:ProgressPreference = $prevProgress
}
