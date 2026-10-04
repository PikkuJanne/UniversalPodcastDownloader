# Source audit and review traceability

Reviewed main commit: `2ac82614493be7196c9ebee116f23fec07368b50`. Verified again on 2026-10-01 through GitHub branch/tree reads. This is static review, not a Windows runtime or penetration-test result. Source locations are function anchors; line spans are approximate and must be checked against the actual fetched source. The original 23 recommendations remain in scope.

Pinned source: https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/2ac82614493be7196c9ebee116f23fec07368b50/UniversalPodcastDownloader.ps1

Pinned launcher: https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/2ac82614493be7196c9ebee116f23fec07368b50/UniversalPodcastDownloader.bat

Pinned README: https://github.com/PikkuJanne/UniversalPodcastDownloader/blob/2ac82614493be7196c9ebee116f23fec07368b50/README.md

## R01 — Incomplete file protection

**Evidence anchor:** Main foreach and Invoke-WebRequest -OutFile destination; existing-file Test-Path skip (roughly lines 545–614).

**Assessment:** Observed source logic can leave final-name partial files; actual interruption must be reproduced locally.

**Implementation:** UPD-0103, UPD-0105.

## R02 — Stable identity and collision handling

**Evidence anchor:** New-EpisodeFileName/Get-EpisodeData; hash suffix conditional on fallback title or missing date (roughly 331–420).

**Assessment:** Observed naming collision path; state/history proposed, not already present.

**Implementation:** UPD-0102, UPD-0104, UPD-0105.

## R03 — Singleton collection compatibility

**Evidence anchor:** Episode selection/count/progress (roughly 516–565).

**Assessment:** Source uses singleton PSCustomObject Count; 5.1 behavior documented [S2]; test both engines.

**Implementation:** UPD-0101.

## R04 — Safe web parsing

**Evidence anchor:** Get-FeedUrlInteractive, Resolve-PodcastItems and media transfer.

**Assessment:** Source uses Invoke-WebRequest without UseBasicParsing; Microsoft documents the 5.1 security prompt/change [S3].

**Implementation:** UPD-0101.

## R05 — Path/name safety

**Evidence anchor:** Sanitize-ForWindowsName and Join-Path using feed title.

**Assessment:** Dot/dot-dot, reserved/control/length cases are not fully handled; write-boundary protection needed [S4].

**Implementation:** UPD-0102.

## R06 — Truthful outcomes

**Evidence anchor:** Final summary and [OK] output; batch wrapper unconditional Done.

**Assessment:** Observed final success-looking message despite failed episode list; define process result contract.

**Implementation:** UPD-0301.

## R07 — Early/private logs

**Evidence anchor:** LogFile initialized after feed/title/folder; URL logs in main.

**Assessment:** Observed early-log gap and full URL output; secret leakage depends on actual URL data.

**Implementation:** UPD-0106.

## R08 — Real WhatIf

**Evidence anchor:** SupportsShouldProcess declaration and unguarded download loop.

**Assessment:** No explicit program-level ShouldProcess gate in inspected workflow; underlying cmdlet behavior is not a complete dry-run contract [S7].

**Implementation:** UPD-0107.

## R09 — URL/XML boundaries

**Evidence anchor:** Interactive scheme warning then continue; [xml] conversion in Resolve-PodcastItems.

**Assessment:** Hardening recommendation. No exploit was demonstrated; parser defaults vary by runtime.

**Implementation:** UPD-0107.

## R10 — Retry/timeout policy

**Evidence anchor:** Three attempts, fixed 3-second delay, no explicit request timeouts.

**Assessment:** Observed policy limitations; new bounded policy proposed.

**Implementation:** UPD-0201.

## R11 — Resume

**Evidence anchor:** Script LIMITATIONS plus transfer loop.

**Assessment:** No resume implementation observed; build on temporary/identity foundations.

**Implementation:** UPD-0202.

## R12 — Shared discovery

**Evidence anchor:** Get-FeedUrlInteractive vs direct FeedUrl path; Find-RssInHtml first candidate.

**Assessment:** Observed different interactive/CLI paths and limited HTML handling.

**Implementation:** UPD-0203.

## R13 — Dates

**Evidence anchor:** Get-EpisodeData date parsing and Atom field fallback order.

**Assessment:** Observed current-culture DateTime.Parse and updated before published; proposed original-publish ordering [S9].

**Implementation:** UPD-0204.

## R14 — Audio formats

**Evidence anchor:** First enclosure; extension handling mp3/m4a else mp3.

**Assessment:** Observed format-selection/naming limits; no conversion needed.

**Implementation:** UPD-0205.

## R15 — Bounded All/pagination

**Evidence anchor:** Resolve-PodcastItems returns one fetched document.

**Assessment:** Single-document processing observed; inaccessible historical episodes cannot be assumed available.

**Implementation:** UPD-0206.

## R16 — Overnight controls

**Evidence anchor:** Main progress and transfer loop.

**Assessment:** Preflight/cancellation/concurrency/keep-awake are proposed enhancements; do not conflate with already implemented retries.

**Implementation:** UPD-0302, UPD-0303.

## R17 — CLI/launcher

**Evidence anchor:** needCount condition, default Mode Latest; launcher does not forward args.

**Assessment:** Observed CustomCount-alone inconsistency and wrapper argument omission.

**Implementation:** UPD-0301.

## R18 — Saved shows/batch

**Evidence anchor:** Current parameter block and single-feed main.

**Assessment:** Optional convenience, not a mandatory replacement interface.

**Implementation:** UPD-0304.

## R19 — Automated tests

**Evidence anchor:** Reviewed recursive repository tree: seven root files only.

**Assessment:** No tests/workflow files in that reviewed tree; newer branches must be checked.

**Implementation:** UPD-0002, UPD-0401, UPD-0501.

## R20 — Separation/maintainability

**Evidence anchor:** Functions plus monolithic main orchestration.

**Assessment:** Incremental import-safe seams then consolidation, not wholesale rewrite.

**Implementation:** UPD-0002, UPD-0401.

## R21 — Documentation/help

**Evidence anchor:** README commands; large introductory block; source behavior differs in details.

**Assessment:** Align examples/help with actual implemented behavior and tested engines.

**Implementation:** UPD-0402.

## R22 — Releases/provenance

**Evidence anchor:** MIT LICENSE present; previous release API read returned empty.

**Assessment:** Release status must be rechecked at implementation; preserve MIT; signing remains optional/gated.

**Implementation:** UPD-0403, UPD-0404, UPD-0502.

## R23 — Website product page

**Evidence anchor:** User request for a website for tools.

**Assessment:** Prepare portable distribution/support assets; no hosted download service or chosen host assumed.

**Implementation:** UPD-0405, UPD-0502.

## Review refinements carried into implementation

Keep proof levels precise: a completed network transfer is not a guarantee that a publisher served the intended entire episode; enclosure length is advisory; short suffix hashes are not mathematically collision-proof; arbitrary URL secrets cannot all be redacted with a short query-key list; PowerShell 5.1 runtime behavior needs direct testing; WhatIf can legitimately read a feed but must have a documented no-media/no-write boundary. No security exploit or end-to-end product success is claimed by this handoff.
