# Initial design decisions

These are proposed project defaults for this improvement project. Record any necessary deviation with reason and evidence; do not silently replace the owner's workflow.

| ID | Decision | Rationale |
|---|---|---|
| D001 | Improve existing PowerShell tool; no rewrite. | Preserve a successful simple workflow. |
| D002 | Windows PowerShell 5.1 and PowerShell 7 on Windows. | Existing advertised compatibility remains a requirement. |
| D003 | No new runtime third-party dependencies. | Keep double-click installation approachable. Development test tools are separate. |
| D004 | One task per Codex thread; one feature branch per milestone. | Keep context bounded and GitHub continuation reliable. |
| D005 | Local active-machine testing is authoritative; CI is supplementary. | Windows behavior must be observed, not inferred from Linux helper tests. |
| D006 | Unknown legacy files are unverified, never silently removed/replaced. | Protect the current archive. |
| D007 | Local versioned JSON first; no database server. | Sufficient for modest single-user archives if locking/atomic update is correct. |
| D008 | Sequential downloads by default. | Avoid connection pressure and needless concurrency complexity. |
| D009 | Hashes help identity/integrity checking but do not prove publisher authenticity. | Do not overstate what local hashes establish. |
| D010 | Preview may read feed/HTML but does not request media or persist anything. | Clear safe preview contract. |
| D011 | CustomCount alone implies Custom; contradictory explicit Mode rejected. | Remove current surprising parameter behavior. |
| D012 | Exit codes: 0 success; 1 setup/fatal; 2 incomplete; 130 catchable cancellation. | Make automation interpret outcomes. Module helpers return objects/errors rather than exiting the host. |
| D013 | HTTP response framing and RSS enclosure length have different authority. | Metadata estimates can be stale; avoid false corruption claims. |
| D014 | Website means product/distribution content, not hosted downloading. | No server storage, hosting selection or unrelated-site changes. |
| D015 | Release tags/publication, merges/settings and paid signing remain owner-gated. | Review and explicit publication control. |

## UPD-0002 decisions (2026-10-02)

- **D016:** Dot-source the existing script to load functions without startup/main execution. Relocating startup configuration below a dot-source return preserves parameters, nine function bodies, main block and two-file distribution. No runtime module is needed yet.
- **D017:** Pin Pester 5.7.1/PSScriptAnalyzer 1.24.0 to reviewed Gallery hashes in a local developer cache. These fixed versions were exercised on both Windows engines; they are not claims about the newest releases. Setup is explicit and never runs from the downloader/test runners.
- **D018:** Use a source-context/count baseline for original lint warnings, with zero unmatched findings allowed. Exclude Write-Host for retained TUI/developer summaries. Defect characterizations are separate from future acceptance; PS5.1 BasicParsing injection is a labeled harness aid, not a production fix. See evidence/UPD-0002.md.

## UPD-0101 decisions (2026-10-02)

- **D019:** Route all existing page/feed/media calls through a small `Invoke-PodcastWebRequest` helper that always supplies `UseBasicParsing`. Keep Invoke-WebRequest, optional OutFile and existing request/error behavior. No parser-default, TLS or persistent policy mutation. This supersedes the D018 worker injection; historical evidence remains intact.
- **D020:** Wrap extraction/filtering in arrays and isolate sorting/mode selection in `Select-PodcastEpisode`, returning one array object for zero/one/many entries. Keep the existing explicit empty-feed/no-enclosure errors. This permits direct shape tests while preserving the current main loop, entry-point names and progress calculations. Broader progress UX remains UPD-0303.

## Future decision entry

ID; date; task; problem; considered alternatives; choice; compatibility/security/archive implications; evidence; owner approval where needed.

## UPD-0102 decisions (2026-10-02)

- **D021:** Separate pure naming and Windows path helpers into bundled `src/Naming.ps1` and `src/PathSafety.ps1`. The original entry points remain, but installation now includes `src/`; no runtime package is added. Importing defines helpers without writes or network access.
- **D022:** New feed folders and episode names retain full SHA-256 identity suffixes. Feed identity uses the exact resolved feed URL; episodes use the existing RSS GUID then exact media URL, with a stable metadata fallback for direct helper calls. Case-sensitive full keys and case-insensitive destination maps reject contradictory identity reuse or forced hash/name collisions within the plan. Identical entries share one planned destination. Parser/Atom identity and durable alias/history policy remain assigned to later tasks. Old archives stay untouched and unadopted; new names may cause a separate download. Hashes do not prove transfer completeness or publisher authenticity.
- **D023:** Budget the complete path for both Windows engines: at most 259 UTF-16 code units for files and 247 for directories, without long-path registry changes or device namespaces. Keep full identifiers while shortening title text without splitting surrogate pairs; omit an optional date only when needed. Roots too long for minimum identifiers fail clearly. Reject raw rooted metadata and dot components before normalization; check canonical containment and all existing ancestors/targets for reparse points at writes. These checks do not provide a sandbox against a concurrent privileged local filesystem attacker.
- **D024:** Destination safety requires no-overwrite final placement even if a competing file appears during a request. Each attempt reserves a unique sibling `.upd-<GUID>.tmp`, retains Invoke-WebRequest with safe parsing, rechecks paths, and uses the two-argument File.Move. Only its owned temporary file is cleaned on caught failure, after rechecking safety. This narrow prerequisite does not complete UPD-0103: transfer framing/content validation, crash recovery and incomplete outcomes remain unimplemented/unaccepted. A killed process may leave an unclaimed temporary file, which subsequent runs preserve.
- **D025:** Give each log a unique run ID and create it without overwrite; revalidate its path before every append. Logs retain their existing content pending the assigned privacy task. No source URL or secret is placed in a filename.

## UPD-0103 decisions (2026-10-02)

- **D026:** Stream media through the built-in .NET HttpClient with ResponseHeadersRead, one GET and an exclusively held caller-owned destination stream. Keep page/feed Invoke-WebRequest with UseBasicParsing. This supersedes the media portion of D019/D024 and adds no runtime package. Loopback probes showed PS5.1 returning a short body without throwing, so compare actual bytes with the same response's Content-Length explicitly. Disable automatic decompression, request identity encoding and reject unexpected content encodings/partial responses; preserve platform proxy and TLS defaults. PS5.1 can normalize away Content-Length when chunked framing takes precedence, so raw-header conflict detection is not guaranteed. Header timeout uses the platform default; body idle/retry policy remains UPD-0201.
- **D027:** Validate the closed temporary file before no-overwrite File.Move. Return explicit outcome, bytes, signature, verification and warnings. Read at most 64 KiB, including a bounded seek past ID3 artwork; reject empty/text/unrecognized bodies and obvious header truncations. Recognizable MPEG, WAV, FLAC, Ogg and MP4 signatures are limited evidence, not full decoding or authenticity. Warn for unverified Ogg/MP4 tracks and contradictory MIME. Publisher enclosure lengths are advisory; meaningful HTTP framing lengths are authoritative. Richer media selection/extension policy remains UPD-0205.
- **D028:** Retain CreateNew, same-directory temporary files and path rechecks; keep the destination handle exclusive through transfer and flush before closing. Caught failures delete only the current attempt's owned file. Abrupt process death can leave a partial which later runs preserve and never count as complete or resume. Completion/history reconciliation belongs to UPD-0104; sidecars, adoption and resume are not introduced. Process-local debugger breakpoints test real stream interruption, before/after final rename and a competing file without product test hooks.
- **D029:** A run with exhausted episode failures throws after its failed summary and cannot print the success banner. The current entry point reports error code 1; the full 0/2/1/130 and launcher contract remains UPD-0301. Remove the obsolete post-download size lookup and its two lint allowances because the transfer now supplies measured bytes.

## UPD-0104 decisions (2026-10-02)

- **D030:** Bind each show to its original exact feed URL fingerprint and explicit alias fingerprints in bounded schema-1 history. Discover that state before deriving a new title-based folder. Reject duplicate aliases, corrupt state and reparse discovery paths; no automatic alias inference or recovery reset. Root discovery/initialization uses a brief exclusive lock, closing the simultaneous first-run/title-change race. Per-show locks then permit independent transfers. No index or new runtime dependency is needed.
- **D031:** Prefer RSS GUID, Atom ID, then exact media URL, with a source fingerprint and a feed-scoped full episode hash. Preserve recorded destinations across mutable metadata changes. Check the entire parsed snapshot for contradictory reuse before selection. Cross-refresh publisher ID reuse and fallback signed-URL changes remain inherently ambiguous. Existing archives without these records are unadopted; migration remains UPD-0105.
- **D032:** Commit prepared size/SHA-256/transfer evidence before no-overwrite final placement, then completed evidence afterward. Reruns hash disk bytes, reconcile prepared finals, safely redownload missing files and report changed/unknown destinations as conflicts. Persist only relative paths and fingerprints; request URLs stay exact in memory. Local hashes establish consistency, not authenticity. Unclaimed partials remain untouched.
- **D033:** Use held FileShare.None handles, strict versioned JSON validation, generation checks, owned flushed metadata temporaries and File.Replace with the preceding valid state as backup. Corrupt primary/backup, newer schema and unsupported atomic replacement fail without reset or fallback. Process crashes are tested; power-loss/UNC durability is not claimed. Preserve persistent lock filenames and release handles in finally. Add the UTF-8 BOM required for the existing script's non-ASCII source on PS5.1 and remove its obsolete lint allowance.

## UPD-0105 decisions (2026-10-02)

- **D034:** Inventory existing immediate show directories without writes. Original names and full modern hash suffixes are candidate hints only. Hash and inspect local signatures under a read handle; classify unknown plausible files as unverified and empty, partial, unsafe or ambiguous files as conflicts. Explicit adoption selects one episode, one exact filename and its reviewed SHA-256; recheck that digest while denying writes through the metadata commit. Title-only legacy matches stop ordinary downloads for review rather than implicitly creating a second archive.
- **D035:** Promote to schema 2 only for an explicit migration operation. Version 1 continues unchanged for ordinary archives. Schema 2 adds unverified and adopted, with adopted requiring owner-approved-local-signature evidence and an explicit transfer-completeness-unverified note. Matching adopted files remain adopted on rerun; missing or changed adopted files require another owner decision. Never promote local observations into completed-transfer evidence. No schema downgrade is allowed.
- **D036:** Explicit redownload selects one episode and allocates a separate bounded filename. Preserve all original bytes and paths, including occupied modern destinations. Save unique, flushed full-state checkpoints before each migration action; rollback restores only that checkpoint's metadata at a new generation, retains schema 2 and saves an undo checkpoint. Never remove media, unknown partials or checkpoints. Corrupt primary/backup still requires manual inspection; rollback is not automatic corruption recovery.
- **D037:** Fix the existing unused ShouldProcess boundary now because migration preview requires actual no-write behavior. Both normal WhatIf and migration Preview read feeds and local evidence without creating roots, logs, history, locks or media. Explicit actions honor WhatIf and Confirm. Broader request/XML policy and later A022 acceptance remain UPD-0107.
