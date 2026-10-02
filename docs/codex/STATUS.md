# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0002 only**, following completed UPD-0001.

- UPD-0001: done; [historical evidence](evidence/UPD-0001.md).
- UPD-0002: **in_progress**; implementation and local checks complete, push/CI pending.
- Next after completion: **UPD-0101 — collections and safe Windows web requests**, not started.
- Branch/upstream: `codex/upd-m0-foundation` / `origin/codex/upd-m0-foundation`.
- Reused [draft PR #1](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/1), base `main`.
- [Commands](DEVELOPMENT.md) and [UPD-0002 evidence](evidence/UPD-0002.md).

## Observed repository state

Actual root: `D:\projects\UniversalPodcastDownloader`. Original `UniversalPodcastDownloader-main` snapshot remains untouched. Origin fetch/push: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`.

Starting local/remote/PR HEAD: `9fc4c004a55943fe9e099a93247299921bfa8f1e`; clean tree. Verified fetch found 0 ahead/0 behind. Remote main remains `2ac82614493be7196c9ebee116f23fec07368b50`. No newer work or instruction conflict; no reset/stash/merge/rebase.

SSH remains unavailable; use the existing authenticated GitHub CLI HTTPS fallback and command-scoped author settings from UPD-0001 evidence. Saved origin and persistent credential settings stay unchanged.

## Implemented foundation

Dot-sourcing the original script now loads functions without startup preferences or main execution. Its parameter block, nine function bodies and main block remain unchanged. Batch launcher and six other original files are unchanged; no runtime module/dependency added.

Added explicit repository-local setup, Pester 5.7.1/PSScriptAnalyzer 1.24.0 package hash locks, test/analysis runners, synthetic product tests and a Windows 5.1/7 CI matrix. Python is integration tooling only. Runners never install tools and fail on missing dependencies, empty/all-skipped selections and test failures. Analysis retains known warnings and rejects unmatched findings.

## Actual local checks

Windows 10.0.26300.0; Python 3.14.7; Git 2.56.0.windows.1; gh 2.97.0.

| Engine | Pester All | Analysis |
|---|---|---|
| PowerShell 7.6.5 | 33 passed / 0 failed / 0 skipped / 0 not_run | 12 files; 0 parse errors; 0 new findings; 10 baseline warnings |
| Windows PowerShell 5.1.26100.9444 | 33 passed / 0 failed / 0 skipped / 0 not_run | 12 files; 0 parse errors; 0 new findings; 9 baseline warnings |

Each run includes 19 product unit checks (six explicit defect characterizations), ten runner checks and four integration checks. These are not 33 completed acceptance cases. A004/A005 locally verified; A006-A060 remain not_run. Historical helper evidence stays separate.

## Known limitations

- Unmodified PS5.1 noninteractive entrypoint fails during legacy web parsing/confirmation.
- Harness-only UseBasicParsing exposes the PS5.1 singleton Count/division-by-zero defect. Successful PS5.1 multi-item transfers use that labeled test default; production requests are unchanged.
- Filename collisions, reserved/dot names, Atom date/identity and first-enclosure defects remain for assigned later tasks.
- Analysis retains legacy warnings, including original UTF-8-without-BOM encoding. PS7 reports one additional Write-Log built-in-command-profile warning.
- Local PS5.1 tests use process-only policy Bypass and a native child module path; no persisted changes. No real archive, private subscriptions or external podcast hosts accessed.
- Initial CI found test-message wrapping and multiple-Python discovery issues; tooling corrections are implemented and the next CI run is pending. Manual launcher, live feeds and broader recovery/security/release acceptance remain unrun.

## Continuity and owner gates

Last observed remote feature SHA at task start: `9fc4c004a55943fe9e099a93247299921bfa8f1e`. Local implementation tested before commit. Final checkpoint will record implementation SHA and CI; final HEAD is reported after push rather than embedded in itself.

No approval needed for authorized feature-branch work/draft PR updates. Merges, rebases, resets, force pushes, existing branch/issue deletion, settings, release tags/publication, purchases and deployment remain outside scope. Stop after UPD-0002.
