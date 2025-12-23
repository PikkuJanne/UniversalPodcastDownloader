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

function Sanitize-ForWindowsName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $safe = ($Name -replace '[\\\/\:\*\?\"<>\|]', '_').Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return $null }
    return $safe
}

# --- RSS autodetect helpers ---
function Find-RssInHtml {
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][string]$BaseUrl
    )

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

            if ($html -match '<rss' -or $html -match '<feed') {
                Write-Host "  This looks like a direct RSS/Atom feed." -ForegroundColor Green
                return $inputUrl
            }

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

# --- Feed parsing ---
function Resolve-PodcastItems {
    param([string[]]$Feeds)

    foreach ($u in $Feeds) {
        Write-Verbose "Trying feed: $u"
        try {
            $resp = Invoke-WebRequest -Uri $u
            $xml  = [xml]$resp.Content

            $items = $xml.SelectNodes('//*[local-name()="rss"]/*[local-name()="channel"]/*[local-name()="item"]')
            if (-not $items -or $items.Count -eq 0) {
                $items = $xml.SelectNodes('//*[local-name()="feed"]/*[local-name()="entry"]')
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

    $url = $null
    $enc = $XmlItem.SelectSingleNode('./*[local-name()="enclosure"][1]')
    if ($enc) {
        $attr = $enc.Attributes["url"]
        if ($attr) { $url = $attr.Value }
        if (-not $url) { $url = $enc.GetAttribute("url") }
    }

    if (-not $url) {
        $ln = $XmlItem.SelectSingleNode('./*[local-name()="link" and @rel="enclosure"][1]')
        if ($ln) { $url = $ln.GetAttribute("href") }
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

    [PSCustomObject]@{
        Title   = ($title -as [string])
        PubDate = $pubDate
        Url     = $url
        Guid    = ($guid -as [string])
    }
}

function New-EpisodeFileName {
    param(
        [Parameter(Mandatory)]$Episode,
        [Parameter(Mandatory)][int]$Index
    )

    $ext = 'mp3'
    if ($Episode.Url -match '\.(mp3|m4a)($|\?)') { $ext = $Matches[1] }

    $titleRaw  = if ([string]::IsNullOrWhiteSpace($Episode.Title)) { 'Episode' } else { $Episode.Title }
    $safeTitle = (Sanitize-ForWindowsName $titleRaw)
    if (-not $safeTitle) { $safeTitle = 'Episode' }

    $prefix = ''
    if ($Episode.PubDate) { $prefix = '{0:yyyy-MM-dd} - ' -f $Episode.PubDate }

    $name = "$prefix$safeTitle.$ext"

    if ($safeTitle -eq 'Episode' -or [string]::IsNullOrWhiteSpace($prefix)) {
        $id = if ($Episode.Guid) { $Episode.Guid } elseif ($Episode.Url) { $Episode.Url } else { "$Index" }
        $bytes = [Text.Encoding]::UTF8.GetBytes($id)

        $sha = [Security.Cryptography.SHA1]::Create()
        try {
            $hash = ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join ''
        } finally {
            $sha.Dispose()
        }

        $hash = $hash.Substring(0,8)
        $name = "$prefix$safeTitle-$hash.$ext"
    }

    return $name
}

# --- Main ---
try {
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

    $baseOutputPath = $OutputPath
    if (-not (Test-Path -LiteralPath $baseOutputPath)) {
        Write-Host "[*] Creating base output directory: $baseOutputPath"
        New-Item -ItemType Directory -Path $baseOutputPath -Force | Out-Null
    }

    $candidateFeeds = @($FeedUrl) | Where-Object { $_ } | Select-Object -Unique

    Write-Host "[*] Fetching podcast feed..."
    $resolved = Resolve-PodcastItems -Feeds $candidateFeeds
    Write-Host "    Using feed: $($resolved.Url)"

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
        Write-Host "    Feed title: $feedTitle"
    } else {
        Write-Host "    Feed title: (unknown)"
    }

    $safeFeedTitle = (Sanitize-ForWindowsName $feedTitle)
    if (-not $safeFeedTitle) { $safeFeedTitle = 'UnknownPodcast' }

    $OutputPath = Join-Path $baseOutputPath $safeFeedTitle
    if (-not (Test-Path -LiteralPath $OutputPath)) {
        Write-Host "[*] Creating podcast folder: $OutputPath"
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $dateStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile = Join-Path $OutputPath ("{0}_{1}.log" -f $dateStamp, $safeFeedTitle)

    Set-Content -LiteralPath $script:LogFile -Value "UniversalPodcastDownloader log"
    Write-Log ("Feed URL     : {0}" -f $FeedUrl)
    Write-Log ("Resolved URL : {0}" -f $resolved.Url)
    Write-Log ("Feed title   : {0}" -f ($feedTitle -as [string]))
    Write-Log ("Mode         : {0}" -f $Mode)
    Write-Log ("CustomCount  : {0}" -f ($CustomCount -as [string]))
    Write-Log ("Output folder: {0}" -f $OutputPath)

    $episodes = foreach ($it in $resolved.Items) { Get-EpisodeData $it }
    $episodes = $episodes | Where-Object { $_.Url }

    Write-Log ("Feed items with valid URLs: {0}" -f $episodes.Count)

    if (-not $episodes -or $episodes.Count -eq 0) {
        throw "Feed parsed, but no downloadable enclosure URLs were found."
    }

    $episodes = $episodes | Sort-Object PubDate -Descending

    switch ($Mode) {
        'Latest' { $episodes = $episodes | Select-Object -First 1 }
        'All'    { }
        'Custom' { $episodes = $episodes | Select-Object -First $CustomCount }
    }

    $total = $episodes.Count
    Write-Host "[*] Episodes to download: $total"
    Write-Log ("Episodes to download (after mode/filter): {0}" -f $total)

    $downloaded = @()
    $skipped    = @()
    $failed     = @()
    $maxRetries = 3

    $index = 0
    foreach ($ep in $episodes) {
        $index++

        $fileName = New-EpisodeFileName -Episode $ep -Index $index
        $destFile = Join-Path $OutputPath $fileName

        if (Test-Path -LiteralPath $destFile) {
            Write-Progress -Activity "Podcast downloads" -Status "Skipping (exists): $fileName" `
                -PercentComplete ([int](($index/$total)*100)) -CurrentOperation "Episode $index of $total"
            Write-Host "[-] Skipping (already exists): $fileName"
            $skipped += [PSCustomObject]@{ Title = $ep.Title; File = $destFile }
            Write-Log ("Skipping existing file: {0}" -f $destFile)
            continue
        }

        $pct = [int](($index/$total)*100)
        Write-Progress -Activity "Podcast downloads" -Status "Preparing: $fileName" -PercentComplete $pct `
            -CurrentOperation "Episode $index of $total"

        Write-Host "[+] Downloading ($index/$total): $($ep.Title)"
        Write-Verbose "    URL: $($ep.Url)"
        Write-Verbose "    OUT: $destFile"

        Write-Log ("Starting download {0}/{1}: {2}" -f $index, $total, ($ep.Title -as [string]))
        Write-Log ("Target file: {0}" -f $destFile)
        Write-Log ("Source URL : {0}" -f $ep.Url)
        if ($ep.Guid) { Write-Log ("GUID      : {0}" -f $ep.Guid) }
        if ($ep.PubDate) { Write-Log ("PubDate   : {0:yyyy-MM-dd HH:mm:ss}" -f $ep.PubDate) }

        $success   = $false
        $attempt   = 0
        $lastError = $null

        while (-not $success -and $attempt -lt $maxRetries) {
            $attempt++
            Write-Log ("Attempt {0} of {1}" -f $attempt, $maxRetries)

            try {
                Invoke-WebRequest -Uri $ep.Url -OutFile $destFile
                $success = $true
            } catch {
                $lastError = $_
                $msg = $lastError.Exception.Message
                Write-Warning ("    Attempt {0} of {1} failed: {2}" -f $attempt, $maxRetries, $msg)
                Write-Log ("Attempt failed: {0}" -f $msg) 'WARN'
                if ($attempt -lt $maxRetries) { Start-Sleep -Seconds 3 }
            }
        }

        if ($success) {
            $size = $null
            try { $size = (Get-Item -LiteralPath $destFile).Length } catch {}
            Write-Host "    Saved: $fileName"
            $downloaded += [PSCustomObject]@{ Title = $ep.Title; File = $destFile }
            Write-Log ("Download succeeded: {0}" -f $destFile)
            if ($size -ne $null) { Write-Log ("File size: {0} bytes" -f $size) }
        } else {
            Write-Warning ("    Giving up after {0} attempts." -f $maxRetries)
            $errMsg = if ($lastError) { $lastError.Exception.Message } else { "Unknown error" }
            $failed += [PSCustomObject]@{ Title = $ep.Title; File = $destFile; Error = $errMsg }
            Write-Log ("Giving up after {0} attempts: {1}" -f $maxRetries, ($ep.Title -as [string])) 'ERROR'
            Write-Log ("Last error: {0}" -f $errMsg) 'ERROR'
        }
    }

    Write-Progress -Activity "Podcast downloads" -Completed

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
            Write-Log ("Failed: {0} ({1})" -f ($f.Title -as [string]), $f.Error) 'ERROR'
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
