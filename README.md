# UniversalPodcastDownloader — Universal RSS podcast downloader for Win11 (PowerShell)
Minimal, no-frills podcast downloader I use to archive my favorite shows for offline listening. It’s a personal, purpose-built tool, not a general “podcast manager”. It keeps a simple TUI, predictable directory structure, and local logs for reviewing completed runs and failures.

**Synopsis**  
- Accepts either:
  - A direct RSS/Atom feed URL, or  
  - A normal “show page” URL and tries to auto-detect the RSS feed.  
- Downloads newest episodes first, with three modes:
  - Latest (1 newest episode)
  - Custom (N newest episodes)
  - All (everything in the feed)
- Creates per-podcast subfolders based on feed title:
  - `<OutputPath>\<SafeFeedTitle>-<feed hash>\YYYY-MM-DD - Episode title-<episode hash>.mp3`
- Starts a private local log before feed discovery, then continues it in the podcast folder after an ordinary download is confirmed.
  - `download-<UTC timestamp>-<unique run ID>.log` records attempts, safe error categories and the run summary. Preview writes no log.

**Requirements**  
- Windows 11  
- PowerShell (Windows PowerShell or PowerShell 7 is fine)  
- Internet access  
- Optional: .bat wrapper for double-click/TUI usage

**Installation**  
- Place these files together:
  - UniversalPodcastDownloader.ps1
  - UniversalPodcastDownloader.bat (wrapper for double-click)
  - src/ (all included network, XML, naming, path safety, media, identity, history and diagnostic helpers)
- Default output root is:
  - %USERPROFILE%\Downloads\Podcasts
- No external binaries are required; the script uses the built-in .NET HttpClient for pages, feeds and streamed media, and a bounded XML reader for feeds.

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
     - `%USERPROFILE%\Downloads\Podcasts\<SafeFeedTitle>-<feed hash>\`
   - A log file for the run is written next to the audio files when that location is writable. Startup failures have a fallback log as described below.

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
  - A safe title followed by a full SHA-256 suffix derived from the resolved feed URL.
- Episodes:
  - `YYYY-MM-DD - Episode title-<episode hash>.mp3` when a date is known.
  - `Episode title-<episode hash>.mp3` otherwise. Identity uses the RSS GUID, then Atom ID, then the exact media URL, scoped to the feed.
  - Names are shortened to fit Windows path limits; the full identifier and extension remain. A tight budget can omit the date. A root with insufficient room fails with a request for a shorter path.
  - Reserved device names, control characters and trailing dots/spaces are handled safely. Dot/dot-dot and rooted metadata paths are rejected. Junctions and symbolic links in destination paths are refused.
- Existing files are never overwritten:
  - Recorded files are skipped only after their size and SHA-256 match local history. Downloaded files and explicitly adopted files retain separate evidence levels. Changed or unknown files at a planned destination are preserved and reported for review. Missing files with transfer evidence can be downloaded again; missing or changed adopted files require review.
  - A file appearing during transfer is preserved. An owned temporary sibling is moved into place only when the final name is available.
  - A feed title change retains its established folder. A stable RSS GUID or Atom ID retains its recorded filename across episode title, date and signed media URL changes.
  - A legacy folder matching the old title-based name or the proposed archive folder blocks an ordinary new download when unclaimed media needs review. Use the explicit legacy workflow below to inspect and choose one episode at a time. A title match never binds a feed identity automatically.
  - Changed feed URLs require an explicit alias association; titles and redirects do not establish one automatically. Without a publisher identifier, a changed media URL can mean a new episode.
- Downloads are validated before final placement:
  - Each attempt writes to a unique `.upd-<GUID>.tmp` in the destination folder. Streams close before validation and final rename.
  - Empty bodies, text/error pages, unsupported binary signatures, incomplete HTTP bodies and unsolicited partial responses fail. Valid recognizable audio can succeed without Content-Length. A feed enclosure-length mismatch produces a warning.
  - Checks read at most 64 KiB for recognizable MPEG audio, WAV, FLAC, Ogg or MP4 signatures. They do not decode the whole file or prove publisher authenticity; Ogg/MP4 audio tracks are not verified.
  - A caught failure cleans only its own temporary file. A killed process may leave a temporary sibling; rerunning starts a fresh download and preserves that old partial. There is no resume or automatic orphan cleanup.

**Local history and recovery**

- Each established show stores versioned history in `.upd/state.json` and its previous valid generation in `.upd/state.json.bak`. Records contain relative destinations, identity fingerprints, outcomes, measured bytes, SHA-256 and bounded transfer evidence. Request URLs stay exact in memory and are not stored in history.
- Ordinary schema-1 histories remain schema 1. An explicit legacy action promotes the archive to schema 2, which also supports `adopted` and `unverified`. History writes and rollback never downgrade the schema.
- Exclusive file handles serialize writers. The brief output-root `.upd-archive.lock` protects discovery and initial show creation; `.upd/writer.lock` protects that show's run. Lock files remain after exit; the operating system releases their handles on process death. Retry a busy archive after its writer finishes.
- Prepared evidence is saved before final placement. If the process stops after placement but before the completion record, the next run checks those exact bytes before recording completion. A digest proves consistency with recorded bytes, not publisher authenticity.
- Corrupt, unsupported or contradictory state is preserved and stops the run. Keep the primary, backup and media for inspection; there is no automatic reset or backup restoration. Explicit legacy rollback uses a selected checkpoint as described below. An unreadable show history blocks discovery under that output root because its feed aliases cannot be ruled out safely.
- State replacement requires filesystem support for atomic same-volume replacement. Local Windows tests cover process interruption; hardware power-loss durability and live network shares are not guaranteed.

## Review and migrate a legacy archive

Start with a copy of an older archive. `LegacyPath` must name an existing immediate show directory inside `OutputPath`; it can be an absolute path or that folder's name. Supply `FeedUrl` explicitly. Legacy actions use the whole current feed for identity matching and do not ask for a download count.

Set these example paths and URL to the copied archive and its feed:

```powershell
$legacy = @{
    FeedUrl = 'https://example.com/feed.xml'
    OutputPath = 'D:\Podcasts'
    LegacyPath = 'D:\Podcasts\Original show'
}
$preview = & .\UniversalPodcastDownloader.ps1 @legacy
$preview.Episodes | Select-Object EpisodeId, Title, Classification, Candidates, SuggestedPath | Format-List
$preview.Files | Select-Object RelativePath, Bytes, Sha256, Plausible, Classification, Reason, RecordedStatus | Format-List
```

The default action is `Preview`. It may fetch the feed, but creates no media, directories, state, logs, configuration or lock files. It inventories immediate files, hashes readable originals and performs a bounded local signature check. Empty files, text bodies, partial filenames and ambiguous matches remain conflicts. Unknown plausible media remains `unverified`. An existing binding appears as `recorded`, with its evidence level in `RecordedStatus`.

The returned inventory is sensitive local review data: titles and exact filenames are needed to choose an adoption. Keep it local. Diagnostic exports exclude inventory, history and checkpoints.

Historical title/date names and hash suffixes are only matching hints. A suggestion requires one file for one episode. Repeated titles or multiple copies require an explicit file choice. You may choose a differently named plausible file after reviewing it; a file already bound to another episode cannot be adopted for this one.

### Adopt one reviewed local file

Copy a full episode ID and an exact filename from the preview. The digest ties the decision to the reviewed bytes; the operation rejects a file that changed afterward.

```powershell
$episodeId = Read-Host 'Paste the full reviewed EpisodeId'
$fileName = Read-Host 'Paste the exact reviewed RelativePath'
$chosen = @($preview.Files | Where-Object { $_.RelativePath -ceq $fileName })
if ($chosen.Count -ne 1 -or -not $chosen[0].Plausible) {
    throw 'Choose one plausible file from a fresh preview.'
}
$adopt = @{
    LegacyAction = 'Adopt'
    LegacyEpisodeId = $episodeId
    LegacyFile = $fileName
    LegacySha256 = $chosen[0].Sha256
}
& .\UniversalPodcastDownloader.ps1 @legacy @adopt -WhatIf
$result = & .\UniversalPodcastDownloader.ps1 @legacy @adopt
$result
```

Adoption keeps the original path and bytes. Its status is `adopted`, with `owner-approved-local-signature` evidence and `local-signature-only; transfer-completeness-unverified` confidence. A digest and recognizable signature do not prove that the publisher supplied a complete or correct episode. Later runs count unchanged adopted files separately from verified transfer skips. Missing or changed adopted files require another review; they do not trigger an automatic replacement.

### Download one separate replacement

Use this action when you choose to obtain a new transfer for the reviewed episode. It allocates a separate filename, preserves every original and records the observed transfer for the selected episode.

```powershell
& .\UniversalPodcastDownloader.ps1 @legacy -LegacyAction Redownload -LegacyEpisodeId $episodeId -WhatIf
$result = & .\UniversalPodcastDownloader.ps1 @legacy -LegacyAction Redownload -LegacyEpisodeId $episodeId
$result
```

The new file receives `transfer_verified` evidence from the observed transfer and bounded signature check. This does not establish cryptographic completeness or publisher authenticity.

### Roll back a metadata decision

Each changing legacy action returns a `Checkpoint` basename and prints it before its final mutation. Keep that value. The `.upd/legacy-<id>.json` checkpoint contains the archive's full prior history metadata. Rollback restores all records from that snapshot in a new generation, so it can also remove metadata decisions made after the checkpoint. Review its proposed `Operation.RestoreRecords` before applying it.

```powershell
$checkpoint = $result.Checkpoint
$rollbackPreview = & .\UniversalPodcastDownloader.ps1 @legacy -LegacyAction Rollback -LegacyCheckpoint $checkpoint -WhatIf
$rollbackPreview.Operation.RestoreRecords | Format-List
$rollback = & .\UniversalPodcastDownloader.ps1 @legacy -LegacyAction Rollback -LegacyCheckpoint $checkpoint
$rollback
```

Rollback never deletes, renames or restores media bytes. Original files and separate downloads survive; files whose records were removed return to review. Schema 2 and the feed association remain, including when rolling back the first adoption to an empty history. Rollback itself saves another checkpoint. A corrupt checkpoint or one from a different feed is refused; copying `.bak` over live state is not an automatic recovery procedure.

## Preview and network boundaries

An ordinary download supports `-WhatIf`. It may retrieve the feed or selected show page and read local history/media to build a plan. It makes no enclosure request and creates no output folders, media, history, checkpoints, configuration, logs, exports or lock files. It makes no keep-awake or power-setting change and does not print a completed-download banner. The same policy applies to legacy `Preview` and changing legacy actions run with `-WhatIf`.

```powershell
& .\UniversalPodcastDownloader.ps1 -FeedUrl $legacy.FeedUrl -OutputPath $legacy.OutputPath -Mode All -WhatIf
```

Feed, page and enclosure targets must be absolute HTTP or HTTPS URLs without a username/password component. Private-network and loopback hosts are allowed. Up to five redirects are followed after validating each target; an HTTPS-to-HTTP downgrade is refused. Windows credentials and cookies are not automatically sent, and platform TLS verification and proxy defaults remain in use. Signed path/query values stay in the actual requests while diagnostic displays omit them.

Feed/page responses are limited to 8 MiB and must use an uncompressed complete HTTP 200 body. Feed XML rejects DTDs and external resource resolution, with explicit size and structure limits. Oversized or unsupported input fails before archive execution. A supplied metadata URL or permitted redirect can still return unexpected content, including audio; preview reads that response through the bounded metadata path and does not start enclosure transfers. See [input and preview boundaries](docs/codex/INPUT_BOUNDARIES.md) for exact limits and compatibility details.

## Diagnostics and privacy

- Before feed discovery, writable runs try `%LOCALAPPDATA%\UniversalPodcastDownloader\Logs`, then `%TEMP%\UniversalPodcastDownloader\Logs`. If neither works, a safe notice goes to standard error. Logging failure never replaces the operation's original error.
- An ordinary confirmed download continues logging in its show folder if possible, replaying up to 256 recent startup events. The startup copy remains in place. Legacy changes keep their startup log. Logs use UTF-8 without a BOM and a unique, no-overwrite filename: `download-<UTC timestamp>-<32-character run ID>.log`.
- URLs displayed by the downloader use only the hostname and a random request ID. User information, paths, queries and fragments are omitted. Logs and routine console output omit raw titles, local paths, publisher IDs, response headers and raw exception messages. Hostnames and episode identity fingerprints can still reveal or correlate subscriptions; review any log before sharing it.
- `-WhatIf`, legacy `Preview` and declined confirmation write no diagnostic files or exports. Explicit confirmation starts a fresh file log only after acceptance; pre-confirmation events stay in memory and are then discarded.
- Optional `-DiagnosticExportPath` writes a new JSON file in an existing directory. It contains only a run ID, start time, PowerShell version and up to 256 event times, levels and fixed codes. It excludes message text, URLs, logs, inventory, state, checkpoints, media and configuration. Existing files are never overwritten.

```powershell
& .\UniversalPodcastDownloader.ps1 -FeedUrl 'https://example.com/feed.xml' -Mode Latest `
    -DiagnosticExportPath "$env:TEMP\upd-diagnostic-review.json"
```

Nothing is uploaded automatically. Logs, startup copies and exports remain until you delete them; there is no automatic rotation or retention deadline. Older logs may contain credentials and are never read into new exports. Shell input history, transcripts, caller-inspected error objects and the legacy inventory remain sensitive local data. Keep archive history, backups and migration checkpoints for verification and recovery. See the [diagnostic and data retention policy](docs/codex/DIAGNOSTICS.md) for the exact boundaries.

**Batch wrapper (included)**  
- UniversalPodcastDownloader.bat (double-click launcher):
  - Double-click = open the TUI.
  - If you want drag-and-drop support later, you can extend the wrapper to pass %* through to the script.

**Technical details**
- Feed resolution:
  - Direct RSS/Atom content is detected via <rss> / <feed>.
  - For normal HTML pages, the tool scans for:
    - <link type="application/rss+xml" ... href="..."> or Atom equivalents.
  - Relative discovered links use the final page URL after validated redirects. Redirects alone do not change a stored feed identity.
- Episode parsing:
  - Uses explicit XML reader limits with DTD processing prohibited and external resolution disabled.
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
  - Metadata and episodes get up to 3 attempts for transient failures. Permanent HTTP, validation and local-file errors stop immediately.
  - Exponential backoff respects `Retry-After`. A server delay beyond the retry budget reports a deferred failure.
  - Separate 30-second header and body-idle timeouts allow long downloads that keep making progress.
  - Configure `-MaxAttempts`, `-HeaderTimeoutSeconds`, `-IdleTimeoutSeconds`, `-RetryBudgetSeconds`, `-BaseDelaySeconds` and `-MaxDelaySeconds` when needed. See [retry and timeout policy](docs/codex/TRANSPORT_POLICY.md) for defaults and limits.
  - Media failures use local error categories; a run with failed episodes reports an error without an “[OK]” completion message.
  - HTTP Content-Length is checked against bytes received when the platform exposes it. Original media bytes are kept; unexpected HTTP content encodings are rejected.

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
- Input rejected by the network or XML policy:
  - Use an absolute HTTP(S) URL without user information. A feed requiring browser cookies or automatic Windows authentication is unsupported.
  - Redirect loops, more than five redirects, HTTPS-to-HTTP redirects, oversized metadata and XML containing a DTD are refused. Ask the publisher for a direct supported feed if needed.
- Only some episodes downloaded:
  - Open the latest .log file in the podcast folder.
  - Look for failed attempts and safe error categories. Raw server error details are deliberately omitted.
  - Rerun the same feed; unchanged history-backed files are verified and skipped.
- Failure before a podcast folder is available:
  - Look in `%LOCALAPPDATA%\UniversalPodcastDownloader\Logs`, then `%TEMP%\UniversalPodcastDownloader\Logs`. A logging failure is reported on standard error without exposing the rejected path.
- A file needs review:
  - Unknown or changed files are preserved. Use `-LegacyPath` to preview the archive, then explicitly adopt one reviewed file or redownload one episode to a separate target. Keep the originals, state, backup and returned checkpoints.

**Intent & License**
This is a personal tool for a very specific workflow (downloading and archiving podcast episodes I care about, with logs I can read later). It’s provided as-is, without warranty. Use at your own risk. If you want to reuse or adapt it, feel free, just keep in mind it intentionally avoids features to stay simple, predictable, and easy to reason about when something fails at 03:00.
