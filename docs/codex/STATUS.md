# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0103 only**.

- UPD-0001, UPD-0002, UPD-0101 and UPD-0102: done; historical evidence preserved.
- UPD-0103: **in_progress**; implementation and all local checks complete, push/CI checkpoint pending.
- Next after completion: **UPD-0104 — stable identities and durable local history**; unstarted.
- Branch/upstream: `codex/upd-m1-safety` / `origin/codex/upd-m1-safety`.
- M1 [draft PR #2](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/2) remains stacked on `codex/upd-m0-foundation`; M0 draft PR #1 remains unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0103 evidence](evidence/UPD-0103.md), [decisions](DECISIONS.md).

## Verified repository and behavior

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Original `UniversalPodcastDownloader-main` snapshot remains unmodified. Saved origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/remote/PR head was `7a8cebff73607515c948828b012e67025904273f`; clean, fetch divergence 0/0. Predecessor [CI 36994440854](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/36994440854) passed both Windows jobs at that exact SHA. Use the established authenticated HTTPS fallback and command-scoped author identity from UPD-0001 evidence; no persistent settings changed.

Each media attempt holds a unique CreateNew sibling exclusively, streams one HTTP response, checks outcome/length, closes the writer and validates bounded media signatures before no-overwrite final placement. Invalid empty/text/partial/truncated bodies fail; valid lengthless recognizable audio succeeds. Enclosure-length estimates are advisory. Caught failures clean only the attempt's temporary file. Process interruption leaves an unclaimed partial which later runs preserve while starting fresh. Finalization races preserve the competing file. Failed episode runs produce an error without the success banner.

Safe names/full identity suffixes, canonical/reparse checks, Windows path budgets and collision planning from UPD-0102 remain. Installation includes every bundled `src/` helper. Original entry points, modes and unchanged audio remain. Page/feed requests keep explicit safe parsing; media uses built-in HttpClient, with no new runtime package.

## Actual local checks

Windows 10.0.26300.0; Python 3.14.7; Git 2.56.0.windows.1; gh 2.97.0; Pester 5.7.1; PSScriptAnalyzer 1.24.0.

| Engine | Focused A011-A013 units | Integration | Full All | Analysis |
|---|---|---|---|---|
| PowerShell 7.6.5 | 68/0/0/169 | 40/0/0/0 | 277/0/0/0 | 25 files; 0 parse errors/new findings; 6 baseline warnings |
| Windows PowerShell 5.1.26100.9444 | 68/0/0/169 | 40/0/0/0 | 277/0/0/0 | 25 files; 0 parse errors/new findings; 5 baseline warnings |

Counts are passed/failed/skipped/not_run. Full suite: 227 product units, ten runner guards and 40 loopback integration cases. Three parser defect characterizations remain. Python fixture-helper selftests: 20/0/0, separate from product acceptance. A001-A013 passed; A014-A060 remain not_run. The obsolete size lookup's two lint allowances were removed.

## Continuity and limits

- Signature/HTTP completion is bounded evidence, not full decoding, publisher completeness or authenticity. Ogg/MP4 tracks remain unverified. PS5.1 can normalize away Content-Length when chunked framing takes precedence; universal raw-header conflict detection is not claimed.
- Existing final files/legacy archives remain untouched and unadopted; skips are unverified. New names or changed feed URLs/titles can create separate downloads. Durable identity/history and completion/history reconciliation are UPD-0104; resume, orphan cleanup and migration are not implemented.
- Platform default header timeout is 100 seconds; body idle/retry policy remains UPD-0201. Full exit-code/launcher/cancellation semantics remain UPD-0301. Logs retain legacy URL content pending UPD-0106.
- Junctions and actual crash/finalization races were exercised. Privileged file symlinks/live UNC and hardware power-loss durability were not. There is no perfect concurrent-adversary sandbox claim.
- Only synthetic/mock/loopback data and marked owned temporary outputs were used. No real archive/private feeds, manual launcher, merges, history rewriting, branch/issue deletion, persistent settings changes, releases, purchases or deployments.

## GitHub checkpoint

Implementation push and its exact-head CI are pending. Keep UPD-0103 in_progress until these are verified; do not start UPD-0104 yet. Final documentation checkpoint and its CI will be recorded separately in draft PR #2/final handoff, since this file cannot embed its own commit SHA.
