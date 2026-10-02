# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0201 only**.

- UPD-0001, UPD-0002 and UPD-0101 through UPD-0107: done; historical evidence preserved.
- UPD-0201: **in_progress**; implementation and focused checks complete; full matrix verification pending.
- Next after completion: **UPD-0202 — add validator-aware safe resume**; not started.
- Branch: `codex/upd-m2-network-feeds`, created from verified M1 tip `82f4ca88e3b304e7cb85449758b09f0e9d892ce9`.
- M1 draft PR #2 and M0 draft PR #1 remain open/unmerged. M2 draft PR will be stacked on `codex/upd-m1-safety`.
- [Commands](DEVELOPMENT.md), [UPD-0201 evidence](evidence/UPD-0201.md), [retry and timeout policy](TRANSPORT_POLICY.md), [input boundaries](INPUT_BOUNDARIES.md), [diagnostics](DIAGNOSTICS.md), [history and migration](STATE_AND_MIGRATION.md).

## Verified checkout and changes

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Preserved `UniversalPodcastDownloader-main` snapshot untouched. Origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/fetched remote/PR HEAD `82f4ca88e3b304e7cb85449758b09f0e9d892ce9` was clean with 0/0 divergence. [Predecessor CI 37013964074](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37013964074) passed both jobs at that exact SHA. Continue command-scoped authenticated HTTPS and public author identity without persistent settings changes.

One shared transport policy now bounds attempts, connection/header waits, idle body reads and the elapsed retry window. Metadata retries from empty memory; recorded media retries with a fresh owned temporary per attempt. The old fixed episode retry loop is removed. Transient HTTP/network failures may retry with exponential backoff; permanent HTTP, TLS/policy, validation, filesystem and history failures stop. Retry-After dates/seconds on errors and redirects never shorten the server delay; over-budget waits defer.

Defaults are three total attempts, 30-second headers, 30-second idle reads, 120-second retry window, one-second base backoff and 30-second local cap. Script parameters expose all six values. A progressing body can outlast the header timeout and retry window; there is no short total media deadline. Transport failure categories produce fixed private messages. Original audio, history schemas, prepared completion evidence, no-overwrite placement, locks, preview and legacy protections remain. No runtime dependency was added.

## Checks and limits

Focused A023-A026 unit selection passed **126/0/0/425 per engine**. Initial broad PS7 unit snapshot passed **550/0/0/0** before the final shared-deadline case was added. PS7 transport integration passed 42/0/0/111. Final analyzers pass 54 files with zero parse errors/new findings and 5/4 existing warnings. A real native truncated-body classification issue was corrected at the source-read boundary; new focused units pass 51/0/0/501 on both engines. Full local All, native transport rerun and CI are pending; evidence records intermediate fixture corrections without treating those failures as passes.

Local Windows 10.0.26300.0; PowerShell 7.6.5 / 5.1.26100.9444; Pester 5.7.1; PSScriptAnalyzer 1.24.0; bundled Python 3.12.14. CI Python stays pinned to 3.14.7. Use the bundled Python directory via process-only PATH because the ambient alias is unusable. Tests isolate logs and archives in marked owned temporary data.

A001-A024 passed; A025-A060 remain not_run until verification completes. Retry time is an attempt-start budget, not a total transfer deadline or an operating-system cleanup guarantee. Unknown failures stop conservatively. No resume/range append, richer feed parsing, launcher/exit-code redesign, manual launcher verification, real-archive test, merge, history rewrite, persistent setting change, release or publication is claimed.
