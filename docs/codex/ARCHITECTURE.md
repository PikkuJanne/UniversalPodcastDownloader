# Implementation design and behavior contracts

UPD-0401 records the implemented portable architecture below. Earlier design contracts remain applicable where their implementation sections identify the evidence. Keep the original entry points and simple default workflow; release packaging and verified user help remain separate later tasks.

## Implemented structure and API boundaries (UPD-0401)

```text
UniversalPodcastDownloader.ps1          # Parameter binding, TUI/CLI, exit boundary
UniversalPodcastDownloader.bat          # Double-click/argument-forwarding launcher
src/*.ps1                              # 27 bundled import-safe implementation helpers
scripts/                               # Development-only test/analyzer/tool setup commands
tests/                                 # Product Pester unit/integration suites
tools/codex-handoff/                    # Supplied synthetic helper kit, not app runtime
docs/codex/                             # Bounded task/evidence continuity
```

The runtime needs only the two root entry files, the complete relative `src` directory and a supported Windows PowerShell engine. It does not load `scripts`, `tests`, `tools`, documentation or development modules. The script explicitly loads 24 helper files; NetworkPolicy loads TransportPolicy/ResumePolicy and MediaTransfer loads Preflight. Repeated definition-only imports do not allocate run state. There is no installed package manager, module manifest, Python or native-library download requirement. HTTP/DPAPI assemblies and the optional keep-awake bridge initialize lazily on the relevant invocation.

| Boundary | Responsibility |
| --- | --- |
| Root `.ps1` / `.bat` | Parameter binding, guided input, orchestration, literal launcher forwarding and the process exit boundary |
| Commands / Batch / RunResult | Optional command routing, one shared saved-run argument allowlist, strict primitive overrides, sequential aggregation and schema-1 results |
| FeedDiscovery / FeedXml / FeedPagination / PublicationDate / MediaSelection | Bounded source classification/parsing, catalogue collection, UTC ordering and supported audio candidate selection |
| NetworkPolicy / TransportPolicy / MediaRequest | One URL/redirect/client policy, bounded retries/header/idle waits, metadata buffering and media response framing |
| MediaTransfer / MediaValidation / Preflight | Confirmed transactional byte transfer, bounded media indications, owned write/capacity checks |
| HistoryStore / HistoryIdentity / HistoryWorkflow / ResumeStore / ResumePolicy | Exclusive state ownership, stable identities, prepared/completed ordering and strict owned resume evidence |
| LegacyInventory / LegacyMigration | Explicit private review and confirmed metadata/adoption/redownload/rollback actions |
| Naming / PathSafety | Supported filename allocation, containment and reparse checks |
| Diagnostics / Progress / KeepAwake | Bounded safe diagnostic events, real-byte presentation and optional temporary native leases |
| SavedShows | Separate strict configuration, CurrentUser protection, private permissions and atomic mutation |

The stable entry API is `Invoke-PodcastRun` for one show and `Invoke-PodcastCommand -Options` for the same workflow plus opt-in management/named/batch operations. Dot-sourcing loads functions without starting a run or exiting the caller. `Invoke-PodcastBatch` is the lower-level batch API; its primitive override allowlist and configuration policy are unchanged. Lower-level helpers are implementation/testing seams, not independently installable engines. Existing resolver/naming/parser helpers remain callable; standalone requests use fresh default policies and resolvers default to 20 pages. Dot-source parameter values and earlier runs no longer configure hidden script defaults. Pass `Policy`/`MaxPages` explicitly to resolver helpers or use the documented run options.

Only the script boundary translates a result into process status. Run/command/batch helpers return one result and preserve caller preferences and LASTEXITCODE. Management/batch projections omit credential URLs, ciphertext and output paths. Named runs retain the private RunResult contract, including sensitive LegacyResult during a legacy review encountered in preview. See [CLI_RESULTS.md](CLI_RESULTS.md).

`Invoke-PodcastRun` creates one local transport policy and passes it through guided/direct resolution, the selected HTML feed, every catalogue page and legacy/media execution. The operation owns its retry loop; redirect/backoff waits share a helper while retaining the original clock observation, deadlines, rounded sleeps and attempt metadata. Body adapters remain separate because metadata has an 8 MiB memory bound whereas media streams with framing/resume/progress/transaction ownership. Likewise HTML whitespace/deduplication, pagination alias keys, original feed identity, legacy identity hints and resume fingerprints are distinct policies, not interchangeable normalization rules.

Diagnostics retain a bounded run context for deliberate post-close export. Close releases the exact-URL correlation dictionary's keys in finally, including absent/failing writers; it does not erase immutable strings throughout process memory. Progress, native leases and locks retain their private ownership contexts and independent cleanup. Function-local error/confirmation preferences and owned progress cleanup are scoped to the invocation. No product global preference, process execution policy, TLS/certificate, environment, culture or working-directory mutation is needed.

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

UPD-0301 implements `EpisodeResult` with opaque identity, outcome, nullable bytes, verification, attempts and a safe authored message. Outcomes include downloaded, verified_skip, legacy_unverified, conflict, deferred, failed and cancelled. An unknown/adopted legacy file is not a verified skip. Routine results omit local paths and raw error objects.

`Invoke-PodcastRun` returns one schema-1 `Podcast.RunResult`: selection/preview, aggregate episode counts, episode array, catalogue coverage, plan projection, complete/status/exit code and safe message. Any unresolved catalogue makes every mode, including preview, incomplete. Local legacy reviews/actions retain sensitive inventory in `LegacyResult`, including an inventory encountered during ordinary WhatIf. Successful explicit adoption/rollback reports completion of the metadata action without claiming downloaded media; ordinary runs encountering adopted media remain incomplete. Only the entry script exits; dot-sourcing defines the API without starting it. See [CLI_RESULTS.md](CLI_RESULTS.md) for the implemented contract.

## Transport design

UPD-0107 uses the built-in .NET HttpClient for metadata and streamed media behind one URL/redirect policy. `src/NetworkPolicy.ps1` provides `Get-PodcastRequestUri`, `Invoke-PodcastHttpGet` and `Invoke-PodcastMetadataRequest`; `src/MediaRequest.ps1` retains media framing and streaming checks. No runtime package is added. UPD-0201 adds src/TransportPolicy.ps1 for bounded retry and header/idle waits, shared by metadata and transactional media. UPD-0202 adds isolated ResumePolicy/ResumeStore helpers and extends MediaTransfer with validated owned checkpoints, retaining those same transport boundaries on both .NET Framework/PowerShell 5.1 and PowerShell 7. See [RESUME_POLICY.md](RESUME_POLICY.md) for eligibility, ownership, append/restart and recovery rules.

The implemented [transport policy](TRANSPORT_POLICY.md) keeps headers and idle-body limits separate from the elapsed retry-start budget. Each metadata request or episode transaction owns exactly one retry loop. Retry-After delays on failures and redirects share its deadline. Unknown, local-file, validation and history failures stop without retries. A slow but progressing long episode should not hit an arbitrary short total request limit. Dispose responses/streams on all paths. Inject retry waits/clock for fast unit tests. Use the fixture server for byte-level integration checks. Preserve default proxy and certificate validation behavior unless the user explicitly configures supported alternatives; never bypass TLS for convenience.

HTTP-specific mechanics must follow the primary references in SOURCES.md [S6]. In brief, partial transfers require consistent ranges and suitable validators; an ignored range must not be appended; a server delay cannot be shortened into an early retry. The detailed cases in ACCEPTANCE_CASES.json are this project's proposed recovery policies, not evidence of existing implementation.

## Untrusted content and privacy

Treat RSS/Atom/HTML, headers, filenames, URLs, saved config and history as untrusted data. UPD-0107 validates absolute HTTP(S) targets without user information at entry, discovery, enclosure planning and redirect boundaries. Private/loopback hosts and initial HTTP requests remain supported. At most five manually checked redirects are allowed; HTTPS-to-HTTP downgrade is refused. Automatic cookies, default request credentials and automatic redirects are disabled. Platform TLS verification and proxy defaults remain unchanged [S11-S12].

Metadata responses are bounded to 8 MiB before XML parsing. `src/FeedXml.ps1` uses explicit XmlReader settings: DTD prohibited, null resolver, 8,388,608 document characters, and `MaxCharactersFromEntities` set to 1,024. Built-in character references remain subject to the document limit. A first streaming pass checks depth at most 64 and at most 100,000 reader nodes plus attributes; a second pass loads a DOM with the same settings and its own null resolver. HTML discovery also limits direct input to 8,388,608 characters and uses 250 ms regex timeouts. These are explicit input limits, not a claim that the earlier default parser had a demonstrated exploit [S5, S13-S14].

Feed titles and server Content-Disposition names are never arbitrary paths. Canonical path containment is necessary but not sufficient against junction/reparse-point races; fail safely for unsupported cases. Do not claim a perfect sandbox against a concurrent privileged local adversary. Validate output components and existing ancestors at relevant write boundaries [S4].

Private subscriptions may encode secrets anywhere in URLs. Default diagnostics retain only hostname plus a random request ID, omitting all user information, path, query and fragment content. Do not log Authorization, cookie values, raw signed URLs, raw response bodies or full exception objects. An explicit local sensitive debug mode, if later added, needs clear consent and must not be included in shareable exports.

## Implemented preview boundary (UPD-0107)

The script resolves and parses metadata, validates all parsed enclosure targets, reads local state and builds a plan before archive execution. Normal WhatIf returns a diagnostic projection of that plan, or a local legacy inventory when review is required. Legacy Preview returns that review inventory, and changing legacy WhatIf validates the requested choice. Each path exits before writer locks, directory creation, history/checkpoint updates and enclosure transfer. It never reaches the completed-download banner or optional keep-awake activation. Metadata retrieval can receive arbitrary server content, including audio returned by the supplied URL or an allowed redirect; it remains bounded by the metadata reader.

Normal execution passes ShouldProcess before archive writes; legacy changes have their own ShouldProcess decision. After acceptance, both reread and validate relevant state under writer protection. Diagnostics remain in memory for preview, including the finally/export path. Metadata reads and local hashing are allowed, so preview is not an offline operation. Original feed identity is retained across redirects; relative discovered HTML links use the effective page URI. The complete [input and preview policy](INPUT_BOUNDARIES.md) records limits and remaining scope.

## Implemented shared discovery (UPD-0203)

`src/FeedDiscovery.ps1` classifies bounded metadata by the document root, carries fetched content forward and returns all static RSS/Atom discovery candidates. Guided input and CLI use `Resolve-PodcastItems`; guided planning reuses its resolved result. Direct feed identity stays the original URL across redirects, while page discovery resolves against the final page URL and first supported base href. A sole candidate is automatic; multiple candidates require a numbered guided choice or a direct CLI URL. Empty feeds, malformed XML, unsupported roots and no-link pages have distinct fixed messages. [FEED_DISCOVERY.md](FEED_DISCOVERY.md) defines the safe scanner subset, entity decoding, deduplication, one-level discovery and remaining limitations. No runtime dependency is added.

## Implemented publication dates (UPD-0204)

`src/PublicationDate.ps1` parses a bounded invariant Gregorian date subset into UTC DateTimeOffset values. RSS uses pubDate; Atom published precedes updated fallback, with actual date namespaces checked. Numeric UTC instants drive selection, dated entries come first, and explicit source positions resolve equal/missing dates. New filename dates are UTC, while history keeps established paths. A separate old-priority/local-date legacy hint is review-only; publisher-ID extraction remains compatible to prevent history rebinding. [PUBLICATION_DATES.md](PUBLICATION_DATES.md) records grammar, precision, missing-offset assumptions and unsupported inputs.

## Implemented audio selection (UPD-0205)

`src/MediaSelection.ps1` selects the first eligible known audio enclosure in source order, with a generic unknown URL only as fallback. Episode candidates and advisory MIME/extension hints remain in memory. `MediaValidation.ps1` combines completed transport framing with bounded supported audio indications, reports fixed metadata warnings and rejects unknown/text/ambiguous/video probes. New or unfinished allocations resolve canonical extensions before prepared history; established records keep their destinations. Resume ownership remains on the exact provisional path; explicit legacy redownloads retain their separate allocation identity. [AUDIO_FORMATS.md](AUDIO_FORMATS.md) records formats, M4A branding limits and the conservative corrected-path crash window. Original bytes, schemas and preview boundaries are unchanged.

## Implemented diagnostics (UPD-0106)

Bundled `src/Diagnostics.ps1` defines the run context, URL display, safe error formatting, file sinks and restricted JSON export. Importing it has no side effects. Normal startup initializes it before discovery and tries the local application-data log directory, then the temporary log directory. UTF-8 without BOM, UTC timestamps, random full run IDs and CreateNew avoid engine-specific encoding and accidental log replacement. Diagnostic writes are best effort; failure reports a safe notice on standard error and preserves the primary operation error.

After a confirmed ordinary download creates or opens its show directory, the sink switches to a new same-run log there and replays at most 256 recent events. The original startup log remains. If the show sink fails, the startup sink continues. Legacy changes retain their startup sink. Preview and WhatIf keep only a bounded in-memory buffer. Explicit confirmation delays diagnostic writes until acceptance, then initializes a fresh run context and discards the pre-confirmation buffer.

Application-authored messages use safe error categories and omit untrusted titles, publisher IDs, paths and headers. URL correlation IDs are random and scoped to an in-memory bounded run map; exact request strings stay separate from display values. Caught failures produce fixed safe console text and result messages; raw exceptions are excluded from returned run/episode results. PowerShell caller inspection, transcripts, input history and legacy review objects are outside the shareable diagnostic boundary.

`-DiagnosticExportPath` serializes a strict allowlist of run metadata and bounded event times, levels and fixed codes into a new UTF-8 JSON file. It never reads log files or serializes free text, arbitrary runtime properties, local history, inventory, configuration, checkpoints, media, request URLs or exceptions. Preview suppresses this write too. The complete [diagnostic and retention policy](DIAGNOSTICS.md) distinguishes logs from the smaller export and documents remaining hostname/identity correlation, sensitive local data and manual retention.

## Compatibility details

UPD-0303 adds private run/episode progress contexts and optional temporary Windows keep-awake. Media callbacks carry only validated response totals, actual offsets, retry attempts and measured bytes; durable history completion precedes successful episode presentation. Contexts own their progress IDs, preserve caller preferences and close each rendered activity independently. Run cleanup separately attempts progress, power, history and archive release. A typed progress-cleanup cancellation changes an otherwise successful result to 130 while retaining its completed media/counts; an existing failure remains primary.

The keep-awake helper is lazy and off by default. A dedicated managed thread owns a temporary `ES_CONTINUOUS | ES_SYSTEM_REQUIRED` native request and restores its exact previous flags on that same native thread. Confirmation precedes activation, including legacy actions; preview never activates it. Unsupported/native failures are fixed advisories. See [implemented progress and power contract](PROGRESS_AND_POWER.md).

Keep PowerShell 5.1 syntax throughout runtime code: no null-conditional operator, ternary operator, pipeline chain operators, or ForEach-Object -Parallel. Avoid assuming PowerShell 7-only web cmdlet switches or .NET overloads. Use explicit UTF-8 encoding consistently; consider BOM for PowerShell source containing non-ASCII text on 5.1. Preserve/restore any changed process preferences in finally; avoid global settings when scoped alternatives exist.

Use appropriate generic collections internally rather than repeatedly growing large arrays, but return stable public shapes. Date parsing should use explicit culture and offset handling; UTC chronological comparison plus a documented filename date convention avoids machine-dependent behavior. Missing/tied dates must have a deterministic fallback, and old filenames should remain bound through history rather than renamed on every improved parse.

## Implemented UPD-0304 saved shows and batch

`src/Commands.ps1` routes optional saved selectors while forwarding ordinary one-off arguments unchanged. `src/SavedShows.ps1` owns bounded strict schema-1 JSON, current-user DPAPI protection, ACL validation, an exclusive persistent configuration lock and same-directory atomic replacement. Imports perform no configuration I/O or native initialization. Preview/declined management operations stop before directory, lock or protection work. List/export explicitly project metadata; no expression evaluation or portable credential import is supported.

`src/Batch.ps1` validates a bounded selection and allowlisted overrides before executing one existing run at a time. Each child attempts all existing media/history/progress/power cleanup before the next begins; failing disposal/native cleanup remains an explicit limitation. Per-show credential/transfer/catalogue errors remain isolated; structural configuration errors stop safely, and typed cancellation stops with completed counts/unstarted names. Public aggregate children remove sensitive LegacyResult and unknown properties. History/resume/media invariants remain in the existing single-show engine. See [SAVED_SHOWS.md](SAVED_SHOWS.md) and [CLI_RESULTS.md](CLI_RESULTS.md).

## Out of scope

No GUI rewrite, playback, transcoding/tag rewriting, hosted media downloader, DRM/auth bypass, podcast-provider scraper, browser automation, mandatory scheduler, cloud accounts, unrelated website deployment or aggressive parallelism. Saved shows, batch, progress and keep-awake are optional user-facing features, not mandatory new steps in the basic workflow.

## Implemented UPD-0206 pagination

`src/FeedPagination.ps1` collects one supported advertised chain through the shared resolver before identity/date selection. Requested/effective URI aliases detect cycles; source fields and original archive identity remain anchored to the initial feed. Catalogue metadata carries stop reasons and accessible counts. Exact duplicate identities collapse; contradictory metadata stops before writes. Limits and failed later pages preserve accessible work but prevent the completed banner in every mode. Preview remains metadata-only, and guided resolution is reused. See [FEED_PAGINATION.md](FEED_PAGINATION.md) for relation/base, bounds, selection and incomplete-result semantics, and [CLI_RESULTS.md](CLI_RESULTS.md) for the implemented UPD-0301 result/launcher contract.
