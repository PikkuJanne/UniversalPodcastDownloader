# Current implementation status

Updated 2026-10-03 (Europe/Berlin). Scope: **UPD-0301 only**.

- UPD-0001/0002, UPD-0101 through UPD-0107, UPD-0201 through UPD-0206 and UPD-0301: done; historical evidence preserved.
- UPD-0301: **done**; bound CLI choices, private structured results and the Windows launcher verified on both engines.
- Next, ready: **UPD-0302 — polish cancellation, preflight and concurrent-run safety**; unstarted.
- Actual checkout `D:\projects\UniversalPodcastDownloader`; original `UniversalPodcastDownloader-main` snapshot untouched.
- Branch `codex/upd-m3-cli-results`; [M3 draft PR](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/4) stacked on `codex/upd-m2-network-feeds`. Earlier M0/M1/M2 PRs remain open, draft and unmerged.

## Verified behavior

Startup local/fetched remote/M2 PR head `6f19d939da3fb1c6f1616dfdfb31d4362a030342` matched, clean with 0/0 divergence. [Predecessor exact-head CI 37053135970](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37053135970) passed both 1,272-check jobs. Implementation checkpoint `8b963f1c01b5bede34afeb0a2c275fd176837786` contains explicitly named source/test/policy files; final documentation uses a separate commit. Final independently matched pushed SHA and exact-head full CI are recorded in the draft PR/final handoff, avoiding self-referential commits.

A positive CustomCount alone selects Custom; explicitly combining a count with Latest/All fails before discovery or input. NonInteractive requires a nonblank explicit feed, defaults to Latest, rejects enabled Confirm and locally suppresses inherited confirmation. Guided feed/count input and retained-resolution reuse remain available.

Dot-source then call Invoke-PodcastRun for one schema-1 Podcast.RunResult without host exit. Safe episode outcomes distinguish downloaded, verified skip, unverified adopted media, conflict, deferred, failed and cancelled. Only the entry script exits: 0 success/clean preview, 1 fatal input/setup/run failure, 2 incomplete media/catalogue, 130 catchable cancellation. Partial/empty catalogue previews stay 2 without writes. Summary cancellation cannot double-count completed media; request cleanup cannot replace an existing primary failure. Completed explicit legacy metadata actions can succeed without promoting media into transfer evidence. Local LegacyResult inventories remain private, including an inventory encountered during WhatIf.

Batch forwards arguments directly once with quoted engine/script paths and local delayed-expansion handling; it preserves exact child status and reports a missing script visibly. Only an argument-free launcher call retains the final Enter prompt. Actual owned Windows terminal observations confirmed success 0/incomplete 2 guided flows, one initial feed request, original media hashes and a retained closing prompt; parameterized preview/missing-script paths did not pause. See [CLI_RESULTS.md](CLI_RESULTS.md) and [launcher evidence](evidence/UPD-0301-LAUNCHER.md).

Media bytes, naming, history schemas 1/2, locks, no-overwrite, strict owned resume, bounded discovery/pagination/date/audio policies and privacy/preview boundaries remain protected. No runtime package, progress/keep-awake/preflight or release work is added. The development CI deadline is 40 minutes because the measured concurrent local PS7 integration plus Unit suites exceeded 36 minutes before setup/analysis. Product request/retry limits and required assertions remain unchanged.

## Verification and limits

Settled inventory: **1,364 checks** = **1,075 units** (1,065 product plus ten runner guards) + **289 integrations**. Full local Unit passed **1,075/0/0/0 each**, PS7 194.23 s/native 138.73 s. New CLI/result integration passed **21/0/0/268 each**, PS7 105.00 s/native 58.99 s. Eight launcher plumbing cases and four real Windows terminal/process observations are recorded separately. Broad local integration snapshots were PS7 **266/23/0/0**, 1968.36 s/native **266/23/0/0**, 1113.33 s; stale cached message/exit expectations were migrated without weakening physical preservation checks. Settled affected reruns passed PS7 **42/0/0/247**, 275.38 s/native **197/0/0/92**, 765.2 s. Exact final-source full All CI is recorded separately in the PR/final handoff. Counts are passed/failed/skipped/not_run.

Final analyzers: **78 files**, parse/new findings zero, baseline **2 PS7/1 native** warnings. Fixture helpers: **58/0**, separate from acceptance. Local Windows **10.0.26300.0**, PowerShell **7.6.5 / 5.1.26100.9444**, Pester **5.7.1**, PSScriptAnalyzer **1.24.0**, bundled Python **3.12.14**. [UPD-0301 evidence](evidence/UPD-0301.md) records exact commands, snapshots, limits and checkpoints.

A001-A041 passed; A042-A060 remain not_run. Mocked/synthetic/owned loopback output only. No real archive/private feed, graphical Explorer double-click/accessibility, physical Ctrl+C universal guarantee, arbitrary PID kill, power-loss guarantee, merge/history rewrite, account/persistent settings change or publication was performed/tested. One owned native engine cache spill from the terminal probe was preserved under ignored .dev-tools. Stop after UPD-0301; NEXT_THREAD_PROMPT.md describes UPD-0302.
