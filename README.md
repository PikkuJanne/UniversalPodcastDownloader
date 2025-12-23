# UniversalPodcastDownloader — Universal RSS podcast downloader for Win11 (PowerShell)
Minimal, no-frills podcast downloader I use to archive my favorite shows for offline listening. It’s a personal, purpose-built tool, not a general “podcast manager”. It trades features for a simple TUI, predictable directory structure, and verbose logging so I can see exactly what happened overnight.

**Synopsis**  
- Accepts either:
  - A direct RSS/Atom feed URL, or  
  - A normal “show page” URL and tries to auto-detect the RSS feed.  
- Downloads newest episodes first, with three modes:
  - Latest (1 newest episode)
  - Custom (N newest episodes)
  - All (everything in the feed)
- Creates per-podcast subfolders based on feed title:
  - <OutputPath>\<FeedTitle>\YYYY-MM-DD - Episode title.mp3
- Writes a per-run log file in the podcast folder:
  - YYYYMMDD_HHmmss_<FeedTitle>.log with detailed attempt-by-attempt info.

**Requirements**  
- Windows 11  
- PowerShell (Windows PowerShell or PowerShell 7 is fine)  
- Internet access  
- Optional: .bat wrapper for double-click/TUI usage

**Installation**  
- Place these files together:
  - UniversalPodcastDownloader.ps1
  - UniversalPodcastDownloader.bat (wrapper for double-click)
- Default output root is:
  - %USERPROFILE%\Downloads\Podcasts
- No external binaries are required; the script uses Invoke-WebRequest and PowerShell’s XML parsing.

Usage  
1. TUI via .bat (my default)
   - Double-click UniversalPodcastDownloader.bat.
   - Follow the prompts:
     - Paste either:
       - The podcast’s RSS/Atom URL, or  
       - The public show page URL (the tool will try to auto-detect the RSS feed).
     - Choose how many episodes to download (newest first):
       - Enter = latest only
       - Number = N newest
       - all = entire feed
   - The episodes are saved under:
     - %USERPROFILE%\Downloads\Podcasts\<FeedTitle>\
   - A log file for the run is written next to the audio files.

2. TUI via direct PowerShell
   - Run without parameters for the same guided flow:
    .\UniversalPodcastDownloader.ps1
   - You can override the output root:
    .\UniversalPodcastDownloader.ps1 -OutputPath "D:\Podcasts"

3. Command line (non-interactive)
   - Use when you already know the RSS URL and desired mode:
    #All episodes for a known feed
    .\UniversalPodcastDownloader.ps1 
        -FeedUrl "https://example.com/feed.xml" 
        -Mode All 
        -OutputPath "D:\Podcasts"

    #10 newest episodes
    .\UniversalPodcastDownloader.ps1 
        -FeedUrl "https://example.com/feed.xml" 
        -Mode Custom 
        -CustomCount 10 
        -OutputPath "D:\Podcasts"

**Output layout**  
- Default root:
  - %USERPROFILE%\Downloads\Podcasts
- For each feed:
  - A subfolder named after the feed title, sanitized for Windows:
    - %USERPROFILE%\Downloads\Podcasts\<FeedTitle>\
- Episodes:
  - YYYY-MM-DD - Episode title.mp3 when a publication date is known  
  - Episode title.mp3 if no date is available  
- Existing files are never overwritten:
  - Already present files are skipped and noted in the log.

**Logging**  
- Each run produces one log file in the podcast’s folder:  
  - YYYYMMDD_HHmmss_<FeedTitle>.log
- Logged details include:
  - Feed URL and resolved RSS URL
  - Feed title, mode, and output folder
  - Count of items in the feed and how many were selected
  - For each episode:
    - Target filename and full path
    - Download attempts (up to 3) with errors per attempt
    - Final result: downloaded / skipped / failed
  - End-of-run summary (Downloaded / Skipped / Failed) with error messages for failures
- This is meant for “left it running overnight, what went wrong?” scenarios.

**Batch wrapper (included)**  
- UniversalPodcastDownloader.bat (double-click launcher):
  - Double-click = open the TUI.
  - If you want drag-and-drop support later, you can extend the wrapper to pass %* through to the script.

**Technical details**
- Feed resolution:
  - Direct RSS/Atom content is detected via <rss> / <feed>.
  - For normal HTML pages, the tool scans for:
    - <link type="application/rss+xml" ... href="..."> or Atom equivalents.
- Episode parsing:
  - Reads title and publication date (<title>, <pubDate>, <updated>, <published>).
  - Tries to find an audio URL via:
    - <enclosure url="...">
    - <link rel="enclosure" href="...">
    - Fallback: .mp3 URLs in <guid> or <link>.
- Sorting & selection:
  - Episodes are sorted by publication date, newest first.
  - Modes:
    - Latest = first 1
    - Custom = first N
    - All = everything
- Download robustness:
  - Each episode is attempted up to 3 times.
  - Short sleep between retries.
  - Errors and failures are logged with messages from the remote host.

**Troubleshooting**
- Script window closes immediately:
  - Run UniversalPodcastDownloader.bat from an existing cmd window to see errors.
  - Check PowerShell’s ExecutionPolicy and any corporate restrictions.
- “No episodes found in the feed”:
  - The URL may not be a real RSS/Atom feed.
  - Try copying the RSS link from the host (Apple Podcasts, Podbean, etc.).
- “Feed parsed, but no downloadable enclosure URLs were found”:
  - The feed might not expose direct audio URLs, or it uses a custom format.
  - Some feeds only link to web players, not direct files.
- Only some episodes downloaded:
  - Open the latest .log file in the podcast folder.
  - Look for per-episode errors (timeouts, HTTP 403/404, connection resets).
  - Try rerunning for the same feed; existing files will be skipped.
- Skips too many files:
  - The tool skips when the target filename already exists.
  - If you changed naming patterns or moved files manually, adjust or delete the old files before rerunning.

**Intent & License**
This is a personal tool for a very specific workflow (downloading and archiving podcast episodes I care about, with logs I can read later). It’s provided as-is, without warranty. Use at your own risk. If you want to reuse or adapt it, feel free, just keep in mind it intentionally avoids features to stay simple, predictable, and easy to reason about when something fails at 03:00.