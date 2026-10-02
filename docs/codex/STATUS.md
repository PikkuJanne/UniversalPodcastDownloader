# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0201 only**.

- UPD-0001, UPD-0002 and UPD-0101 through UPD-0107: done; historical evidence preserved.
- UPD-0201: **done**; implementation pushed and verified; full local coverage and both GitHub jobs passed.
- Next, ready: **UPD-0202 — add validator-aware safe resume**; not started.
- Branch: `codex/upd-m2-network-feeds`, created from verified M1 tip `82f4ca88e3b304e7cb85449758b09f0e9d892ce9`.
- [M2 draft PR #3](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/3) is stacked on `codex/upd-m1-safety`. M1 PR #2 and M0 PR #1 remain open/unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0201 evidence](evidence/UPD-0201.md), [retry and timeout policy](TRANSPORT_POLICY.md), [input boundaries](INPUT_BOUNDARIES.md), [diagnostics](DIAGNOSTICS.md), [history and migration](STATE_AND_MIGRATION.md).

## Verified checkout and changes

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Preserved `UniversalPodcastDownloader-main` snapshot untouched. Origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/fetched remote/PR HEAD `82f4ca88e3b304e7cb85449758b09f0e9d892ce9` was clean with 0/0 divergence. [Predecessor CI 37013964074](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37013964074) passed both jobs at that exact SHA. Continue command-scoped authenticated HTTPS and public author identity without persistent settings changes.

One shared transport policy now bounds attempts, connection/header waits, idle body reads and the elapsed retry window. Metadata retries from empty memory; recorded media retries with a fresh owned temporary per attempt. The old fixed episode retry loop is removed. Transient HTTP/network failures may retry with exponential backoff; permanent HTTP, TLS/policy, validation, filesystem and history failures stop. Retry-After dates/seconds on errors and redirects never shorten the server delay; over-budget waits defer.

Defaults are three total attempts, 30-second headers, 30-second idle reads, 120-second retry window, one-second base backoff and 30-second local cap. Script parameters expose all six values. A progressing body can outlast the header timeout and retry window; there is no short total media deadline. Transport failure categories produce fixed private messages. Original audio, history schemas, prepared completion evidence, no-overwrite placement, locks, preview and legacy protections remain. No runtime dependency was added.

## Checks and limits

Settled suite: **705 checks** (542 product units, ten runner guards, 153 integrations), including two remaining parser characterizations. Local All snapshots passed **704/0/0/0 on both engines** (PS7 965.71 seconds; native 574.94 seconds). They discovered the suite before the final Framework unit was added; all integration workers used the corrected runtime. Final Unit then passed **552/0/0/0 per engine**, covering that added case and the settled source. Both CI jobs passed **705/0/0/0** at implementation HEAD. Analysis: **54 files, zero parse errors/new findings**, 5/4 existing warnings. Focused transport integrations passed 42/0/0/111 each; helper tests passed 34/0 separately. Evidence preserves initial fixture failures and the genuine Framework truncated-body correction.

Local Windows 10.0.26300.0; PowerShell 7.6.5 / 5.1.26100.9444; Pester 5.7.1; PSScriptAnalyzer 1.24.0; bundled Python 3.12.14. CI Python stays pinned to 3.14.7. Use the bundled Python directory via process-only PATH because the ambient alias is unusable. Tests isolate logs and archives in marked owned temporary data.

A001-A026 passed; A027-A060 remain not_run. Retry time is an attempt-start budget, not a total transfer deadline or an operating-system cleanup guarantee. Unknown failures stop conservatively. No resume/range append, richer feed parsing, launcher/exit-code redesign, manual launcher verification, real-archive test, merge, history rewrite, persistent setting change, release or publication is claimed.


## GitHub checkpoint and continuation

Implementation `22632b891c455f82a9bba9b9a7b8d48f9f3de5c6` is pushed and independently matches the remote. [CI 37016749029](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37016749029) passed both Windows Server 2022 jobs: 705/0/0/0 each, 54 analyzed files, zero parse errors/new findings and 5/4 baseline warnings. CI engines: PowerShell 7.6.6 and Windows PowerShell 5.1.20348.5622. Final checkpoint adds verification records plus a stronger synthetic XML assertion; its SHA and CI are recorded in draft PR #3/final handoff without a self-referential commit loop.

Stop after UPD-0201. NEXT_THREAD_PROMPT.md describes UPD-0202 on the same M2 branch and draft PR; resume is unstarted.
