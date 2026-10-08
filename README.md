# UniversalPodcastDownloader

A local Windows RSS/Atom podcast downloader for keeping original audio with verifiable history. Use the guided text interface, a one-off command, or optional saved shows. There is no runtime package manager, Python, FFmpeg, cloud account, GUI or background service requirement.

Version **1.0.0** is the first stable distribution of the completed downloader. Obtain the portable ZIP, manifest, checksums and notes from the [v1.0.0 release page](https://github.com/PikkuJanne/UniversalPodcastDownloader/releases/tag/v1.0.0); see [package, checksum and source guidance](RELEASE.md). Keep `UniversalPodcastDownloader.ps1`, `UniversalPodcastDownloader.bat` and the **complete `src` directory** together. The [MIT license](LICENSE) covers the software; it does not license podcast recordings.

## Requirements and compatibility

Windows 11, a writable output location and access to the selected feed/media hosts are required. Use Windows PowerShell 5.1 or a supported PowerShell 7 release on Windows; exact exercised versions are below. The `.bat` launcher uses the built-in Windows PowerShell executable; run the `.ps1` in PowerShell 7 to use that engine.

| Environment actually exercised | Engine | Evidence and limits |
| --- | --- | --- |
| Windows 11 Pro, build `10.0.26300.0` | Windows PowerShell `5.1.26100.9444` | Local product suites and copied-runtime checks; [UPD-0401 evidence](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/evidence/UPD-0401.md) |
| Same Windows installation | PowerShell `7.6.5` | Same local suite and copied-runtime scope; [UPD-0401 evidence](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/evidence/UPD-0401.md) |

The [UPD-0402 documentation evidence](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/evidence/UPD-0402.md) records the literal examples and distinguishes current-source checks from prior results. Development test tools are separate from the downloader. These observations do not certify other operating systems, every media file, physical Ctrl+C/sleep/lid behavior, hardware power-loss durability or live network shares.

The supplied launcher uses `-ExecutionPolicy Bypass` for its child process only. It changes no machine or user policy and needs no administrator rights. Organization policy may still prevent execution; use your approved PowerShell workflow. No certificate/TLS exception is installed.

## Quick start

Open PowerShell in the directory containing the script and `src`. Built-in help lists every implemented parameter:

<!-- UPD-0402 example:help -->
```powershell
Get-Help .\UniversalPodcastDownloader.ps1 -Full
```

For the guided interface, double-click `UniversalPodcastDownloader.bat`. Paste a direct RSS/Atom URL or a show-page URL, then choose Enter for the latest episode, a positive count, or `all`. One discovered feed is selected automatically; several require a numbered choice. Noninteractive commands require a direct feed URL when discovery finds several candidates.

The following examples share variables in the **same PowerShell session**. Enter your selected RSS/show-page URL when prompted. A reserved example such as `https://example.invalid/feed.xml` is a placeholder and will fail discovery; it is not a working podcast. Choose a permanent `$root` for real archiving. The displayed temporary root keeps the examples away from your default archive; the downloader does not delete it automatically. Private feed URLs entered at the prompt remain sensitive process/transcript data.

<!-- UPD-0402 example:setup -->
```powershell
$feed = Read-Host 'Podcast RSS or show-page URL'
$root = Join-Path $env:TEMP 'UPD example – ää'
$archive = Join-Path $root 'Archive'
$config = Join-Path $root 'Private settings\shows.json'
$null = New-Item -ItemType Directory -Path $root -Force
```

Preview the accessible catalogue before downloading:

<!-- UPD-0402 example:preview -->
```powershell
& .\UniversalPodcastDownloader.ps1 -NonInteractive -FeedUrl $feed -OutputPath $archive -Mode All -WhatIf
```

Download the latest episode through the launcher. Named arguments are forwarded literally; parameterized launches return without the closing pause. An argument-free launch retains the guided questions and final Enter prompt.

<!-- UPD-0402 example:launcher -->
```powershell
.\UniversalPodcastDownloader.bat -NonInteractive -FeedUrl $feed -OutputPath $archive -Mode Latest
$LASTEXITCODE
```

A positive count alone selects Custom:

<!-- UPD-0402 example:custom -->
```powershell
& .\UniversalPodcastDownloader.ps1 -NonInteractive -FeedUrl $feed -OutputPath $archive -CustomCount 5
$LASTEXITCODE
```

Latest selects one eligible episode, Custom selects up to the count, and All selects the bounded accessible catalogue. `-NonInteractive` without a mode/count selects Latest. Explicit Latest/All together with CustomCount is invalid; Custom requires a positive count. Without an explicit mode/count or NonInteractive, the ordinary interface asks for a selection. Defaults without OutputPath use `%USERPROFILE%\Downloads\Podcasts`.

## Feeds, files and preview

RSS and Atom feeds can be supplied directly or discovered from supported HTML RSS/Atom link elements. Discovery uses the final page URL, a supported HTML base URL and decoded link attributes; it stops at one selected feed. This is not a general website crawler or browser-session importer.

All modes collect supported feed-level Atom `next` / `prev-archive` continuations before selecting episodes. `-MaxFeedPages` defaults to 20, including the first page, and accepts 1–100. The catalogue is also bounded to 10,000 raw entries and 32 MiB of decoded characters. Entry links, HTTP Link headers, provider APIs and other pagination conventions are unsupported. Gaps, cycles, ambiguity or limits retain accessible work and return incomplete. Even a completed supported chain does not prove a complete historical archive. Episodes sort by UTC publication instant; Atom uses published before updated fallback, ties keep feed order and undated entries come last. See [discovery](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/FEED_DISCOVERY.md), [pagination](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/FEED_PAGINATION.md) and [dates](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/PUBLICATION_DATES.md).

Supported audio candidates are MP3, M4A, Ogg Opus/Vorbis/Speex, WAVE and FLAC. The first candidate with a supported audio MIME/URL hint is selected in feed order; only when none is known can the first generic enclosure be selected provisionally. New files receive `.mp3`, `.m4a`, `.ogg`, `.wav` or `.flac` from the received bytes. Checks inspect at most 64 KiB for recognizable audio/container evidence; ambiguous or unsupported containers, text/error pages, empty bodies and invalid HTTP completion fail. Missing Content-Length is supported; feed enclosure lengths are advisory. These checks do not decode a track, prove playability/audio-only content or authenticate a publisher. Original received bytes are retained without transcoding or tagging. See [audio policy](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/AUDIO_FORMATS.md).

New output uses a sanitized show folder and identity suffixes:

```text
<OutputPath>\<SafeFeedTitle>-<feed hash>\
    YYYY-MM-DD - <SafeEpisodeTitle>-<episode hash>.<audio extension>
    .upd\state.json
    .upd\state.json.bak
    download-<UTC timestamp>-<run ID>.log
```

The date prefix is omitted when unavailable. Stable publisher GUID/Atom ID takes precedence over the exact media URL for episode identity. The original selected feed URL identifies the archive; redirects and title matches do not silently bind aliases. Existing recorded paths remain authoritative when titles, dates or extensions change. Without a publisher ID, a changed media URL can mean a new episode; there is no universal cross-feed/content deduplication or alias-management CLI.

Existing media is never overwritten. Verified skips require recorded bytes and SHA-256 to match the actual file. Unknown or changed files remain conflicts for review. Containment/reparse checks, exclusive writer handles, owned temporary files and prepared history before no-overwrite placement protect the archive. `-WhatIf` may fetch bounded metadata and read/hash local evidence, including legacy inventory; it requests no planned enclosure and writes no media, directories, history, checkpoints, configuration, locks, logs or exports. It makes no keep-awake request. A metadata URL can itself serve unexpected bytes, which are read only through the bounded metadata path.

## Optional saved shows

One-off/guided runs do not read or create saved configuration. Saving stores settings without fetching a feed or media. Use a new dedicated private configuration parent; do not pre-create it with shared permissions. These commands use the explicit example configuration and archive:

<!-- UPD-0402 example:saved -->
```powershell
& .\UniversalPodcastDownloader.ps1 -SaveShow demo -FeedUrl $feed -OutputPath $archive -CustomCount 5 -ConfigPath $config
& .\UniversalPodcastDownloader.ps1 -ListShows -ConfigPath $config
& .\UniversalPodcastDownloader.ps1 -ShowName demo -NonInteractive -ConfigPath $config
& .\UniversalPodcastDownloader.ps1 -Batch -NonInteractive -ConfigPath $config
& .\UniversalPodcastDownloader.ps1 -Batch -NonInteractive -WhatIf -ConfigPath $config
& .\UniversalPodcastDownloader.ps1 -RemoveShow demo -WhatIf -ConfigPath $config
$export = Join-Path $root ('shows-summary-' + [guid]::NewGuid().ToString('N') + '.json')
& .\UniversalPodcastDownloader.ps1 -ExportShows $export -ConfigPath $config
```

Names are case-insensitively unique, start with an ASCII letter/digit and contain 1–64 ASCII letters, digits, underscores or hyphens. At most 100 shows are supported. Saving an existing name replaces its settings in place; removing it preserves media/history. The default configuration is `%LOCALAPPDATA%\UniversalPodcastDownloader\saved-shows\shows.json`. ConfigPath requires a saved operation and an absolute supported path. Exactly one saved selector is allowed, except Batch can combine with ShowName. Direct FeedUrl is allowed only while saving; saved operations reject legacy arguments. List/Remove/Export reject download options.

Batch runs sequentially in stored order; an explicit ShowName subset follows its supplied order. A single name works through `.bat`/native `-File`; a multi-element array requires a PowerShell session or in-process API. Named/batch runtime overrides are Mode, CustomCount, OutputPath, KeepAwake, MaxFeedPages and the six transport settings below; they do not change saved settings. Count alone implies Custom. Explicit Latest/All drops a stored Custom count, while explicit Custom needs a positive supplied count unless the saved mode was already Custom. DiagnosticExportPath is not a batch override.

Saved URLs use Windows DPAPI CurrentUser protection. Created directories/files have protected current-user-only owner/DACL permissions. Existing unsafe permissions, corrupt/unsupported schema or unavailable credentials fail while preserving configuration; there is no automatic reset, import, evaluation or permission repair. Mutation re-reads under an exclusive persistent lock and uses owned flushed atomic replacement. Re-save one unreadable credential with its URL under the same Windows identity after inspecting structural errors. A different account/profile must re-save the URLs. List/export/management results omit URL, ciphertext and output path; names can still disclose subscriptions. Export creates a new protected file in an existing directory and never overwrites. It is a sanitized summary, not a credential backup or import format. See [saved-show policy](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/SAVED_SHOWS.md).

## Results, transport and progress

| Exit code | Meaning |
| --- | --- |
| 0 | Successful requested operation or preview; preview is not completed media |
| 1 | Fatal input/setup/state error; batch structural/setup error |
| 2 | Incomplete catalogue, failed/deferred/conflicting/unverified selected media; batch has an isolated fatal/incomplete show |
| 130 | Cancellation was caught and a result could be produced |

Batch continues after an isolated show failure. Cancellation stops remaining shows and retains completed counts/unstarted names; structural failure between children retains processed counts and returns 1. An empty batch succeeds with no selected shows. Hard termination and every physical Ctrl+C path cannot guarantee code 130 or cleanup. `[OK]`/completed-download output is reserved for successful ordinary execution.

`-PassThru` emits the entry script's structured result before it sets process status; native output is PowerShell formatting, not a JSON protocol. Dot-sourcing loads local helpers without prompts, network requests, archive/configuration writes or host exit. `Invoke-PodcastRun` returns one `Podcast.RunResult`; `Invoke-PodcastCommand -Options <dictionary>` also handles saved operations. Ordinary projections use counts, opaque identities and safe messages. `LegacyResult` may contain sensitive local inventory, including an ordinary preview that encounters files requiring review. Keep it local. Batch child projections exclude that inventory. See [CLI/results/quoting policy](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/CLI_RESULTS.md).

<!-- UPD-0402 example:api -->
```powershell
. .\UniversalPodcastDownloader.ps1
$run = Invoke-PodcastRun -NonInteractive -FeedUrl $feed -OutputPath $archive -Mode Latest -WhatIf
$run | Select-Object Status, ExitCode, Preview, Planned, CatalogueComplete
```

The per-run transport defaults are MaxAttempts 3, HeaderTimeoutSeconds 30, IdleTimeoutSeconds 30, RetryBudgetSeconds 120, BaseDelaySeconds 1 and MaxDelaySeconds 30. Transient failures retry with exponential backoff and Retry-After; permanent validation/local errors stop. RetryBudgetSeconds controls starting another attempt or delayed redirect, not the total duration of a progressing download. Headers and body-idle waits are separate. URL targets must be absolute HTTP(S) without user information; private/loopback hosts are permitted. At most five validated redirects are followed; HTTPS-to-HTTP downgrade is refused. No automatic Windows credentials/cookies are sent. Platform TLS/proxy defaults remain. Metadata must be complete, uncompressed HTTP 200 within 8 MiB; bounded XML rejects DTDs and external resolution. See [transport](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/TRANSPORT_POLICY.md) and [input limits](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/INPUT_BOUNDARIES.md).

Interactive ConsoleHost progress uses actual response bytes and validated totals. Unknown totals remain indeterminate; retries reset their offsets and validated resume includes the existing prefix. Progress reaches 100 only after verified history completion. Noninteractive/redirected runs retain readable notices/summaries without animation. `-KeepAwake` is opt-in and requests temporary system wakefulness after confirmation, restoring the owning native thread's prior flags when cleanup is catchable and succeeds. It does not keep the display on or change a power plan. Explicit sleep/lid actions can still apply. See [progress and power limits](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/PROGRESS_AND_POWER.md).

## Review and migrate a legacy archive

Work on an explicit **copy** of an older archive first. For these examples, place the copied show folder directly inside `$root\Legacy archive`, keeping the originals elsewhere. This separate output root avoids a feed already associated with the quick-start archive. LegacyPath selects an existing immediate show folder by name or absolute path and requires an explicit FeedUrl. It uses the bounded accessible catalogue, independent of ordinary download counts. Review every selected episode/file; historical names are hints, not proof of identity or completeness.

<!-- UPD-0402 example:legacy-preview -->
```powershell
$legacyArchive = Join-Path $root 'Legacy archive'
$copiedShow = Read-Host 'Name of a copied show folder directly under Legacy archive'
$legacy = @{ FeedUrl = $feed; OutputPath = $legacyArchive; LegacyPath = $copiedShow; NonInteractive = $true }
$review = & .\UniversalPodcastDownloader.ps1 @legacy -PassThru
$review.LegacyResult.Episodes | Format-List EpisodeId, Title, Classification, Candidates, SuggestedPath
$review.LegacyResult.Files | Format-List RelativePath, Bytes, Sha256, Plausible, Classification, RecordedStatus
```

Default LegacyAction Preview reads/hashes immediate files and creates no downloader files or locks. Partial names, empty/text bodies and ambiguous matches remain conflicts; plausible unknown audio stays unverified. Choose one exact reviewed filename and full EpisodeId. Adoption rechecks the reviewed SHA-256 under writer protection and preserves the original path/bytes:

<!-- UPD-0402 example:legacy-adopt -->
```powershell
$episodeId = Read-Host 'Full EpisodeId from the reviewed inventory'
$fileName = Read-Host 'Exact RelativePath from the reviewed inventory'
$chosen = @($review.LegacyResult.Files | Where-Object { $_.RelativePath -ceq $fileName })
if ($chosen.Count -ne 1 -or -not $chosen[0].Plausible) { throw 'Choose one plausible file from a fresh preview.' }
$adopt = @{ LegacyAction = 'Adopt'; LegacyEpisodeId = $episodeId; LegacyFile = $fileName; LegacySha256 = $chosen[0].Sha256 }
$adoptPreview = & .\UniversalPodcastDownloader.ps1 @legacy @adopt -WhatIf -PassThru
$adoption = & .\UniversalPodcastDownloader.ps1 @legacy @adopt -PassThru
$adoption.LegacyResult
```

An adopted file has local-signature evidence and **unverified transfer completeness**. A completed explicit adoption can return 0 for its metadata action; when selected, later ordinary runs count it separately from verified skips and return 2. Missing/changed adopted files need review, never automatic replacement. A recognizable signature or matching digest does not prove publisher authenticity or a complete correct episode.

Redownload explicitly obtains one episode into a **separate filename**, preserving every original:

<!-- UPD-0402 example:legacy-redownload -->
```powershell
$replacementPreview = & .\UniversalPodcastDownloader.ps1 @legacy -LegacyAction Redownload -LegacyEpisodeId $episodeId -WhatIf -PassThru
$replacement = & .\UniversalPodcastDownloader.ps1 @legacy -LegacyAction Redownload -LegacyEpisodeId $episodeId -PassThru
$replacement.LegacyResult
```

Changing actions print/return a Checkpoint basename. Keep it. A checkpoint snapshots the entire prior history metadata, not media. The following rollback uses the redownload checkpoint; review RestoreRecords before applying it, because an older checkpoint also removes later metadata decisions:

<!-- UPD-0402 example:legacy-rollback -->
```powershell
$checkpoint = $replacement.LegacyResult.Checkpoint
$rollbackPreview = & .\UniversalPodcastDownloader.ps1 @legacy -LegacyAction Rollback -LegacyCheckpoint $checkpoint -WhatIf -PassThru
$rollbackPreview.LegacyResult.Operation.RestoreRecords | Format-List
$rollback = & .\UniversalPodcastDownloader.ps1 @legacy -LegacyAction Rollback -LegacyCheckpoint $checkpoint -PassThru
$rollback.LegacyResult
```

Rollback writes a new history generation, keeps schema 2/feed association and saves another checkpoint. It never deletes, renames or restores media bytes; separate downloads survive and any removed binding needs review. Corrupt or mismatched history/checkpoints are refused. See [state and migration policy](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/STATE_AND_MIGRATION.md).

## Recovery and troubleshooting

Keep `.upd/state.json`, its previous valid `.bak`, migration checkpoints, media and partial/sidecar evidence together. New ordinary archives use schema 1; explicit legacy actions promote to schema 2 and later runs retain it. Prepared size/hash evidence is saved before final placement, then completion is recorded. A restart can reconcile exactly matching placed bytes; hashing proves consistency with recorded bytes, not authenticity. Corrupt/newer/contradictory state is preserved and stops execution. **Do not delete originals, reset history or copy a backup over live state as a repair shortcut.** An unreadable show history can block that output root because alias ownership cannot be ruled out.

Resume is automatic only with an owned valid sidecar, exact prefix size/hash, strong ETag, known total and matching feed/episode/request/path/representation. Appending requires a validated complete-tail 206. Ignored ranges, changed validators or 416 start a new transfer while preserving the old partial. Corrupt sidecars and uncheckpointed crash tails stop for review; a temporary filename alone grants no ownership. Unknown partials are never reused or cleaned up. Catchable cancellation attempts eligible actual-byte checkpoints and independent cleanup, but failed checkpoint/Dispose or hard termination can leave evidence requiring review. See [resume policy](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/RESUME_POLICY.md).

| Symptom | Action |
| --- | --- |
| Writer/configuration busy | Wait for that writer, then retry. Persistent `.upd-archive.lock`, `.upd/writer.lock` and configuration `.lock` files are normal; do not delete them or kill a PID based on their contents. |
| Destination/space failure | Check permissions/free space or choose another root. Confirmed work probes write access and compares known response bytes with available space; unknown lengths/capacity remain explicit and the check cannot guarantee later free space. |
| Empty feed / no eligible audio | Obtain a direct supported RSS/Atom feed; a player page, video-only feed or unsupported enclosure may not provide downloadable audio. |
| Discovery ambiguity / refused input | Supply the direct feed. Browser cookies/login, compressed/oversized metadata, DTD XML, unsafe schemes and TLS downgrade are unsupported. |
| Incomplete run | Review safe error categories and rerun the same feed. Unchanged verified media is hashed/skipped; unresolved catalogue gaps or adopted media remain incomplete. |
| File needs review | Preserve originals and state; use the copied legacy preview and an explicit one-file decision above. |
| Saved settings refused | Preserve the file; inspect schema/permissions. Use a new dedicated private parent rather than broadening existing access. Re-save an unreadable URL under the correct Windows identity only after structural errors are resolved. |
| Failure before show log exists | Inspect `%LOCALAPPDATA%\UniversalPodcastDownloader\Logs`, then `%TEMP%\UniversalPodcastDownloader\Logs`. If logging cannot initialize, a safe notice goes to standard error. |

State/config replacement requires same-volume atomic filesystem support. Local process-interruption tests do not promise hardware power-loss or UNC/share durability. No scheduler, automatic archive cleanup or universal provider/codec support is implemented.

## Diagnostics and privacy

Ordinary writable runs create a unique local UTF-8 startup log and, where writable, a same-run show log; the startup copy remains. Preview and declined confirmation create no logs/exports. Routine output omits raw request URLs, titles, filesystem paths, headers and raw exceptions; URL displays contain only a hostname and random per-run ID. Hostnames and episode fingerprints can still disclose/correlate subscriptions. Inspect logs before sharing; older logs may contain credentials.

One-off `-DiagnosticExportPath` creates a new JSON file in an existing directory without overwriting. It contains only schema/run ID/start time/engine version and up to 256 event times, levels and fixed codes. It excludes messages, URLs/hosts, inventory, configuration, history, checkpoints and media. Nothing is uploaded automatically; selected podcast hosts still receive requests and may log them. Logs/startup copies/exports have no automatic rotation or expiry. Keep archive history/checkpoints for recovery rather than treating them as disposable diagnostics.

DPAPI/ACLs do not protect against another process running as the same user, same-user malware or an administrator. URLs exist in process memory; closing diagnostics clears correlation-map keys, not all immutable strings. Shell history, transcripts, parameter-binding errors, caller-inspected exceptions and private LegacyResult objects can retain sensitive data. See [diagnostic policy](https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/3a45f5096d1952acaf4346a372a626131f08cec4/docs/codex/DIAGNOSTICS.md) for the exact sharing/retention boundaries.
