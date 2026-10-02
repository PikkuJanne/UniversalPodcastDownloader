# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0101 only**.

- UPD-0001 and UPD-0002: done; historical evidence preserved.
- UPD-0101: **done**; implementation pushed/verified, local and GitHub checks passed.
- Next, ready and unstarted: **UPD-0102 — safe deterministic Windows destinations**.
- Branch/upstream: `codex/upd-m1-safety` / `origin/codex/upd-m1-safety`, created from verified M0 tip `7be83d07bd10fa02185b127239acddb7720aa150`.
- M0 remains [draft PR #1](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/1), base `main`. M1 [draft PR #2](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/2) is stacked on `codex/upd-m0-foundation`.
- [Commands](DEVELOPMENT.md) and [UPD-0101 evidence](evidence/UPD-0101.md).

## Verified repository state

Actual root: `D:\projects\UniversalPodcastDownloader`. Original `UniversalPodcastDownloader-main` snapshot remains untouched and was not accessed. Origin fetch/push remains `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`.

Starting local/remote/PR M0 HEAD was `7be83d07bd10fa02185b127239acddb7720aa150`; clean tree. Fetch found 0 ahead/0 behind. Main remains `2ac82614493be7196c9ebee116f23fec07368b50`. No M1 branches/PRs or newer conflicting work existed. UPD-0002 completion independently confirmed with both successful jobs in [run 36989042791](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/36989042791).

Use the existing authenticated GitHub CLI HTTPS fallback and command-scoped author settings from UPD-0001 evidence while SSH is unavailable. Saved origin and persistent settings stay unchanged.

## Implemented in UPD-0101

Extraction/filtering retain arrays; `Select-PodcastEpisode` returns an array for zero/one/many selections in Latest, Custom and All. Empty/no-enclosure input retains explicit errors before episode progress/media requests. Counts and existing progress arithmetic now work for singleton PSCustomObjects in PS5.1.

All page/feed/media requests use `Invoke-PodcastWebRequest` with explicit `UseBasicParsing`. The existing network engine, retries and entry-point names remain. Both PS5.1 failure characterizations are now desired-behavior regressions; harness parsing injection is removed.

## Actual local checks

Windows 10.0.26300.0; Python 3.14.7; Git 2.56.0.windows.1; gh 2.97.0; Pester 5.7.1; PSScriptAnalyzer 1.24.0.

| Engine | Focused A006/A007 units | Integration | Full All | Analysis |
|---|---|---|---|---|
| PowerShell 7.6.5 | 39/0/0/29 | 14/0/0/0 | 82/0/0/0 | 14 files; 0 parse errors/new findings; 10 baseline warnings |
| Windows PowerShell 5.1.26100.9444 | 39/0/0/29 | 14/0/0/0 | 82/0/0/0 | 14 files; 0 parse errors/new findings; 9 baseline warnings |

Counts are passed/failed/skipped/not_run. Full All includes 58 product unit checks (six unchanged defect characterizations), ten runner guards and 14 real loopback integration checks. A006/A007 pass locally and in CI. A008-A060 remain not_run.

## Observed GitHub checks

[Successful implementation run 36990787628](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/36990787628), exact head `12be5abbe0bc3741a275b06a90e22b70bfa7253a`, completed 2026-10-02 09:37:54 UTC. Both Windows Server 2022 jobs passed tests, analysis and runtime inventory:

- PS7 7.6.6: 82 passed / 0 failed / 0 skipped / 0 not_run; analysis 14 files, 0 parse errors/new findings, 10 baseline warnings.
- Windows PS5.1.20348.5622: 82 passed / 0 failed / 0 skipped / 0 not_run; analysis 14 files, 0 parse errors/new findings, 9 baseline warnings.

OS 10.0.20348.0; pinned Python 3.14.7, Pester 5.7.1, PSScriptAnalyzer 1.24.0. Local PS7 is the installed 7.6.5; CI exercised the current 7.6.6 update. No system runtime upgrade was made.

## Known limitations and continuity

- Filename collisions, reserved/dot names, Atom date/identity and first-enclosure defects remain for their assigned tasks.
- Analysis retains original warnings. PS7 reports one additional Write-Log built-in-command-profile warning. No baseline allowance was added.
- Local PS5.1 uses process-only policy Bypass and a native child module path. No persistent policy/TLS changes, archive/private subscriptions or live podcast hosts.
- Manual launcher, live feeds and broader recovery/security/release acceptance remain unrun.
- No runtime package/module, transport redesign or entry-point change. No merge/rebase/reset/force-push, deletion, settings, tags/publication, purchase or deployment.

Implementation SHA `12be5abbe0bc3741a275b06a90e22b70bfa7253a` was independently verified equal to the remote branch before PR creation at approximately 11:36 Europe/Berlin. This final documentation checkpoint is pushed and verified separately. Stop after UPD-0101. Final pushed HEAD and CI are recorded in the draft PR/final handoff after the documentation checkpoint to avoid a self-referential commit hash.
