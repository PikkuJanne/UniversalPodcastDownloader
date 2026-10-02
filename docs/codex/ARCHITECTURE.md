# Implementation design and behavior contracts

This document specifies proposed behavior. It is not a description of already implemented functions. Prefer the smallest change that satisfies each task; introduce seams early and consolidate after evidence. Keep root entry-point names and current simple default workflow.

## Suggested structure

```text
UniversalPodcastDownloader.ps1          # Parameter binding, TUI/CLI, exit boundary
UniversalPodcastDownloader.bat          # Double-click/argument-forwarding launcher
src/UniversalPodcastDownloader/         # Small import-safe module, if needed
  UniversalPodcastDownloader.psd1
  UniversalPodcastDownloader.psm1
  Private/                             # Only split where useful to testing
scripts/                               # Test/build developer commands, added in M0/M4
tests/                                 # Product Pester unit/integration suites, added by Codex
tools/codex-handoff/                    # Supplied synthetic helper kit, not app runtime
docs/codex/                             # Bounded task/evidence continuity
```

Do not make a framework out of a small script. A module can be a single small implementation file at first. Imports must not prompt, fetch, create directories or execute the program. Only the script boundary translates result objects into process exit codes. If dot-sourcing is unsupported, document it; never accidentally terminate a parent shell from module helpers.

## Pipeline

1. Bind/validate parameters; establish noninteractive vs guided behavior.
2. Choose diagnostic mode: normal startup logging or in-memory/console-only preview.
3. Resolve input into candidate feed(s) through a shared bounded request adapter.
4. Parse safely into typed episode records; choose a feed explicitly when ambiguous.
5. Normalize identity/date/media candidates, deduplicate, apply documented selection.
6. Read existing state without mutation and build an explicit execution/migration plan.
7. Render plan. If WhatIf/declined confirmation, return without write/media side effects.
8. Acquire archive writer protection, validate paths/state again and execute per-episode transfers.
9. Finalize validated partials without overwrite; update state safely; aggregate outcomes.
10. Dispose resources, release temporary keep-awake/locks, report accurate summary and code.

Do not hold an archive lock during interactive questions, HTML/feed discovery or a dry run. Revalidate the disk/state after acquiring it. A lock for a planned-but-uncreated folder requires careful sequencing: directory creation itself is a confirmed write.

## Proposed internal records

`Episode`: stable feed reference; RSS GUID/Atom ID if present; title; original/parsed publication date; deterministic tie key; array of enclosure candidates; media metadata; source entry position. Keep collection outputs as arrays, including one or zero records.

`DownloadPlanItem`: episode identity; chosen URI held only as needed; safe display URI/opaque request ID; relative target; existing-file classification; requested action; reason. Planning cannot create files or repair history.

`EpisodeResult`: identity; outcome; bytes; local path; verification level; attempts; error category; redacted human message. Suggested outcomes: downloaded, verified_skip, legacy_unverified, conflict, deferred, failed, cancelled. An unknown legacy file is not a verified skip.

`RunResult`: planned, downloaded, verified skips, unverified/conflicts, failed/deferred/cancelled, complete boolean, redacted diagnostics path and exit classification. A partial page fetch makes an All run incomplete even if retrieved episodes all downloaded.

## Transport design

Small early tasks may retain Invoke-WebRequest with safe parsing. The later transport should expose feed retrieval and streamed media transfer separately behind one policy layer. A .NET HttpClient adapter is a reasonable candidate when it simplifies idle timeouts/resume, but verify every API used on .NET Framework/PowerShell 5.1 and the selected PowerShell 7 runtime. Do not maintain two diverging engines without evidence that it is necessary.

Keep connect/headers, idle-body and optional total-budget settings explicit. A slow but progressing long episode should not hit an arbitrary short total request limit. Dispose responses/streams on all paths. Inject retry waits/clock for fast unit tests. Use the fixture server for byte-level integration checks. Preserve default proxy and certificate validation behavior unless the user explicitly configures supported alternatives; never bypass TLS for convenience.

HTTP-specific mechanics must follow the primary references in SOURCES.md [S6]. In brief, partial transfers require consistent ranges and suitable validators; an ignored range must not be appended; a server delay cannot be shortened into an early retry. The detailed cases in ACCEPTANCE_CASES.json are this project's proposed recovery policies, not evidence of existing implementation.

## Untrusted content and privacy

Treat RSS/Atom/HTML, headers, filenames, URLs, saved config and history as untrusted data. Configure explicit XML limits; prohibit DTDs and external resolution [S5]. Use a bounded metadata/HTML response size separately from large streamed audio. Validate URL scheme at entry and redirect boundaries. Respect an original HTTP feed with a clear policy; do not silently disable valid user input, but warn/refuse HTTPS-to-HTTP downgrade according to a documented security policy.

Feed titles and server Content-Disposition names are never arbitrary paths. Canonical path containment is necessary but not sufficient against junction/reparse-point races; fail safely for unsupported cases. Do not claim a perfect sandbox against a concurrent privileged local adversary. Validate output components and existing ancestors at relevant write boundaries [S4].

Private subscriptions may encode secrets anywhere in URLs. Default diagnostics should retain only hostname plus opaque request ID rather than a guessed list of safe path/query fields. Do not log Authorization, cookie values, raw signed URLs, raw response bodies or full exception objects. An explicit local sensitive debug mode, if later added, needs clear consent and must not be included in shareable exports.

## Compatibility details

Keep PowerShell 5.1 syntax throughout runtime code: no null-conditional operator, ternary operator, pipeline chain operators, or ForEach-Object -Parallel. Avoid assuming PowerShell 7-only web cmdlet switches or .NET overloads. Use explicit UTF-8 encoding consistently; consider BOM for PowerShell source containing non-ASCII text on 5.1. Preserve/restore any changed process preferences in finally; avoid global settings when scoped alternatives exist.

Use appropriate generic collections internally rather than repeatedly growing large arrays, but return stable public shapes. Date parsing should use explicit culture and offset handling; UTC chronological comparison plus a documented filename date convention avoids machine-dependent behavior. Missing/tied dates must have a deterministic fallback, and old filenames should remain bound through history rather than renamed on every improved parse.

## Out of scope

No GUI rewrite, playback, transcoding/tag rewriting, hosted media downloader, DRM/auth bypass, podcast-provider scraper, browser automation, mandatory scheduler, cloud accounts, unrelated website deployment or aggressive parallelism. Saved shows, batch, progress and keep-awake are optional user-facing features, not mandatory new steps in the basic workflow.
