# Current implementation status

Updated 2026-10-03 (Europe/Berlin). Scope: **UPD-0302 only**.

- UPD-0001/0002, UPD-0101 through UPD-0107, UPD-0201 through UPD-0206 and UPD-0301/0302: done; historical evidence preserved.
- UPD-0302: **done**; A042/A043 verified on both supported Windows engines.
- Next, ready: **UPD-0303 — improve progress and optional temporary keep-awake**; unstarted.
- Actual checkout `D:\projects\UniversalPodcastDownloader`; original `UniversalPodcastDownloader-main` snapshot untouched.
- Branch `codex/upd-m3-cli-results`; [M3 draft PR](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/4) stacked on `codex/upd-m2-network-feeds`. Earlier M0/M1/M2 PRs remain open, draft and unmerged.

## Verified behavior

Startup local/fetched origin, independently read remote branch and PR head matched `d11aa960fbee3acb795508c08fba6e3bdc8e47ae`, clean with 0/0 divergence. [Predecessor final-source CI 37104893521](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37104893521) passed both full 1,364-check jobs; its synthetic PR checkout tree matched the pushed tree. Named implementation checkpoint `a0eedde417f80a0ad3470afb6f232bb146d4a6fa` contains 25 source/test/policy files; completion documentation uses a separate commit. Final independently matched pushed SHA and full exact-source CI are recorded in the draft PR/final handoff.

Existing OS writer protection now distinguishes same-show/output-root sharing conflicts from inaccessible locks using fixed private guidance. Root selection remains brief and independent shows can transfer concurrently afterward. Stale persistent lock contents never authorize a PID lookup, kill or file deletion; exclusive reopening after handle release is the recovery mechanism.

After confirmation, bounded owned CreateNew/DeleteOnClose probes check destination write access, followed by an actual-show check under the writer lock. A missing root probes its existing ordinary ancestor; separately denied directory creation gets useful fixed guidance. Preview and declined confirmation bypass probes and capacity queries. AvailableFreeSpace is compared with validated remaining response bytes before body copying, including a resume tail. Enclosure sizes stay advisory; arbitrary UNC/unavailable capacity and unknown response size are explicitly reported. No extra media HEAD/GET is added, and capacity can change during transfer.

Catchable cancellation retains partials and forces eligible actual-byte checkpoints before stream closure. Failed checkpoint replacement retains prior evidence for strict review without replacing cancellation. Unknown-length/uncheckpointed partials survive without invented sidecars. Cleanup attempts all legacy resources independently, preserves an existing primary error and retains typed cancellation after otherwise successful disposal. Legacy inspection and resume-prefix hashing also preserve typed cancellation while closing real guards. No universal failing-Dispose/physical Ctrl+C/kill/power-loss guarantee is claimed.

UPD-0301 CLI/result/launcher behavior remains: count alone selects Custom, NonInteractive requires a feed and suppresses prompts, Invoke-PodcastRun returns one private schema-1 result without host exit, and only the entry script sets 0 success/1 fatal/2 incomplete/130 catchable cancellation. Original bytes, recorded destinations, history schemas 1/2, prepared/completed ordering, no-overwrite, strict owned resume, catalogue gaps and privacy boundaries remain intact. No progress/keep-awake/release work or runtime dependency was added.

## Verification and limits

Settled inventory: **1,460 checks** = **1,158 units** (1,148 product plus ten runner guards) + **302 integrations**. Full local Unit passed **1158/0/0/0 each**, PS7 **216.48 s**/native **148.61 s**. Required new integrations passed **13/0/0/289 each**, PS7 **130.39 s**/native **74.76 s**. Existing A015 integrations passed **13/0/0/287** on an earlier 300-case snapshot. Final analysis: **87 files**, parse/new zero, baseline **2 PS7/1 native**. Broad local Integration was launched; its eventual results and final-source full All CI are recorded separately in the PR/final handoff, not presumed passed here. Counts are passed/failed/skipped/not_run.

Local Windows **10.0.26300.0**, PowerShell **7.6.5 / 5.1.26100.9444**, Pester **5.7.1**, PSScriptAnalyzer **1.24.0**, bundled Python **3.12.14**. [UPD-0302 evidence](evidence/UPD-0302.md) records exact commands, desired failing baselines, fixture corrections, source checkpoint and limits. No fixture helper changed; predecessor 58/0 helper results remain separate historical evidence.

A001-A043 passed; A044-A060 remain not_run. Mocked/synthetic/marked owned loopback outputs and ACLs only. No real archive/private feed, arbitrary PID kill, persistent settings change, publication or history rewrite. Actual physical Ctrl+C and universal forced termination/power-loss/resource release remain unclaimed. Stop after UPD-0302; NEXT_THREAD_PROMPT.md describes UPD-0303.
