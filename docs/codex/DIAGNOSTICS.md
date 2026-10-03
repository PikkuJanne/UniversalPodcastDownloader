# Diagnostics and data retention

Implemented by UPD-0106. Diagnostics help review a run while keeping request secrets out of routine output and the optional export. They do not change the actual URL sent to a publisher.

## Log lifecycle

For a writable run, diagnostics initialize before feed discovery. The first writable location wins:

1. `%LOCALAPPDATA%\UniversalPodcastDownloader\Logs` (the Windows local application-data folder is used if the environment value is absent).
2. `%TEMP%\UniversalPodcastDownloader\Logs` (the platform temporary folder is used if that value is absent).

If neither location works, standard error reports that file logging is unavailable. No rejected path or raw sink exception is printed. Diagnostic failure must not replace the original download/discovery error or turn a successful archive operation into a failure.

Files use UTF-8 without a BOM and the name `download-<yyyyMMddTHHmmssfffZ>-<32 lowercase hex run ID>.log`. Timestamps are UTC; each run ID is randomly generated. Creation refuses an existing filename. No URL, feed title or publisher identifier is used in the log basename.

After an ordinary download is confirmed and its show directory is available, the downloader opens a same-run log there and replays up to 256 recent buffered events. Subsequent events use that show log. The startup copy remains where it was created and stops receiving events after the switch. If opening the show log fails, the startup sink continues. Applied legacy actions keep the startup sink.

`-WhatIf` and legacy `-LegacyAction Preview` write no logs or diagnostic exports. Explicit confirmation keeps diagnostic events in memory until accepted. Acceptance starts a fresh run context and file log; the pre-confirmation buffer is discarded. Declining leaves no diagnostic files. Buffers are bounded to 256 recent events, so they are not an unlimited record of a long run.

## Output privacy

URLs are displayed as hostname plus an opaque random request ID, for example `example.com [url <random ID>]`. All user information, path, query and fragment content is discarded. This avoids trying to guess which tokenized path or query names are safe. The same run can correlate a URL while it remains in the bounded request-ID map; no stable URL hash is exported. Exact request URLs remain in process for networking and are never replaced with their display values.

UPD-0401 clears that existing dictionary in Close-PodcastDiagnostics finally, including absent or failing writers. A pristine close creates no diagnostic state. The bounded context/events remain available for an explicit post-close JSON export, and the preview prohibition still applies. Clearing correlation keys does not promise erasure of immutable URL strings throughout process memory, caller variables, shell history or transcripts.

Routine console, verbose and log messages omit raw feed/episode titles, local paths, publisher IDs, headers, response bodies and raw exception messages. Error formatting uses fixed local categories and exact allowlisted application instructions; unrecognized exception text is omitted. Caught failures produce fixed safe console text and result messages; raw exceptions are excluded from returned run/episode results. Text filtering adds a second check to application-authored messages, but it is not a general detector for secrets in arbitrary prose.

Logs can contain hostnames, episode identity fingerprints, dates, counters, attempt numbers, validation categories and outcomes. Hostnames can themselves identify a private subscription, and fingerprints allow correlation. Treat logs as local diagnostic data and review them before sharing. The restricted JSON export below contains fewer fields.

## Optional JSON export

Pass `-DiagnosticExportPath <new-file.json>` to the main script. Its parent directory must already exist and pass destination checks. The exporter creates the file without overwriting; an unavailable destination produces a safe diagnostic notice. WhatIf, legacy Preview and declined confirmation suppress this write too. Export is local and optional; nothing is uploaded.

The schema is an explicit allowlist:

| Field | Contents |
|---|---|
| `schema_version` | `1` |
| `run_id` | Random 32-character run ID |
| `started_utc` | Run start time in UTC |
| `engine_version` | PowerShell version |
| `events` | At most 256 recent accepted event records |
| `events[].timestamp_utc` | Event time in UTC |
| `events[].level` | `INFO`, `WARN`, `ERROR` or `DEBUG` |
| `events[].code` | `run-started`, `archive-log` or `message` |

The exporter omits message text, URL hosts and request IDs as well as raw URLs, headers and exception objects. It never reads log files, traverses archive files, or serializes configuration, history, checkpoints, inventory, media, arbitrary runtime fields or in-memory request maps. Adding data to a runtime object does not add it to the export. The resulting file still reveals when a run occurred, its PowerShell version and event levels; review that metadata before sharing.

## Sensitive local data outside the export

- Legacy inventory and migration results expose titles, exact filenames, identities and checkpoint basenames needed to choose and reverse local actions. UPD-0301 preserves these local review/action objects in RunResult.LegacyResult, including a legacy inventory encountered during WhatIf. Keep these objects local. Ordinary run/episode results contain safe authored messages, counts and opaque identities; no raw exceptions or destination paths.
- Media filenames and history relative paths can contain publisher-provided text. State, backups and checkpoints also retain identity fingerprints and verification evidence. They are not diagnostic exports.
- PowerShell parameter binding, source-loading errors and host behavior occur outside the initialized run boundary. A caller can inspect original exceptions and in-process request state. PowerShell input history, typed commands, transcripts and caller-created dumps may contain credentials. The downloader does not scrub those external records.
- Logs created before UPD-0106 may contain complete URLs, credentials and local paths. They are neither rewritten nor read into current exports. Inspect them locally before any sharing.

## Retention

Startup copies, show logs and exported JSON remain until the owner deletes them. There is no automatic rotation, expiration, upload or remote retention. A failed or abruptly terminated run may leave a diagnostic file; it remains local under the same policy.

Archive state, backups and migration checkpoints have a separate purpose: they support identity binding, verification and explicit metadata recovery. Keep them with the archive. The downloader does not delete or expire them, originals or unknown partial files. Deleting history can remove evidence required for verified skips; deleting checkpoints removes those rollback options. Diagnostic cleanup must not be treated as archive cleanup.

No raw sensitive-debug mode is introduced. Network/XML boundaries and [CLI/launcher results](CLI_RESULTS.md) are implemented; future saved-subscription storage retains its separate privacy task.
