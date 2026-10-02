# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0101 only**.

- UPD-0001 and UPD-0002: done; historical evidence preserved.
- UPD-0101: **in_progress**; implementation and local checks complete, push/PR/CI pending.
- Next after completion: **UPD-0102 — safe deterministic Windows destinations**.
- Branch: `codex/upd-m1-safety`, created from verified M0 tip `7be83d07bd10fa02185b127239acddb7720aa150`.
- M0 remains [draft PR #1](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/1), base `main`. M1 must use a stacked draft PR based on `codex/upd-m0-foundation`.
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

Counts are passed/failed/skipped/not_run. Full All includes 58 product unit checks (six unchanged defect characterizations), ten runner guards and 14 real loopback integration checks. A006/A007 pass locally; CI pending. A008-A060 remain not_run.

## Known limitations and continuity

- Filename collisions, reserved/dot names, Atom date/identity and first-enclosure defects remain for their assigned tasks.
- Analysis retains original warnings. PS7 reports one additional Write-Log built-in-command-profile warning. No baseline allowance was added.
- Local PS5.1 uses process-only policy Bypass and a native child module path. No persistent policy/TLS changes, archive/private subscriptions or live podcast hosts.
- Manual launcher, live feeds and broader recovery/security/release acceptance remain unrun.
- No runtime package/module, transport redesign or entry-point change. No merge/rebase/reset/force-push, deletion, settings, tags/publication, purchase or deployment.

Stop after UPD-0101. Final pushed HEAD and CI will be recorded in the draft PR/final handoff after the documentation checkpoint to avoid a self-referential commit hash.
