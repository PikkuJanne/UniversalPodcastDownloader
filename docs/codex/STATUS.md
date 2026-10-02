# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0102 only**.

- UPD-0001, UPD-0002 and UPD-0101: done; historical evidence preserved.
- UPD-0102: **in_progress**; implementation and all local checks complete; GitHub CI/synchronization pending.
- Next: **UPD-0103 — transactional file completion**; unstarted.
- Branch/upstream: `codex/upd-m1-safety` / `origin/codex/upd-m1-safety`.
- M1 [draft PR #2](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/2) remains stacked on `codex/upd-m0-foundation`; M0 [draft PR #1](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/1) remains unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0102 evidence](evidence/UPD-0102.md), [design decisions](DECISIONS.md).

## Verified repository and implementation

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Original `UniversalPodcastDownloader-main` snapshot remains unmodified. Saved origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/remote/PR head was `8b4487891b39e3d4bf5664651587ea8030681b1d`; clean, fetch divergence 0/0. Predecessor final [CI 36991067325](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/36991067325) passed both Windows jobs at that exact SHA. Use the established authenticated HTTPS fallback and command-scoped author identity from UPD-0001 evidence; no persistent settings changed.

New feed/episode destinations retain full SHA-256 identity suffixes after Windows sanitization and dynamic length budgeting. Pure planners detect full-key/hash/case-insensitive collisions. Dot/rooted metadata is rejected; every relevant write checks canonical containment, existing ancestor/leaf kinds and reparse points. Paths fit legacy Windows limits without enabling long-path settings. Bundled `src/` helper files are now part of installation.

Episode plans are checked before podcast folder/log creation. Each media attempt uses a unique owned sibling and no-overwrite final placement; new log files use unique run IDs and CreateNew. Unsafe destination changes stop retries. Transfer content/framing validation and recovery remain UPD-0103. Original entry points, modes, safe web parsing and array behavior remain.

## Actual local checks

Windows 10.0.26300.0; Python 3.14.7; Git 2.56.0.windows.1; gh 2.97.0; Pester 5.7.1; PSScriptAnalyzer 1.24.0.

| Engine | Focused A008-A010 units | Integration | Final full All | Analysis |
|---|---|---|---|---|
| PowerShell 7.6.5 | 109/0/0/60 | 27/0/0/0 | 196/0/0/0 | 18 files; 0 parse errors/new findings; 8 baseline warnings |
| Windows PowerShell 5.1.26100.9444 | 109/0/0/60 | 27/0/0/0 | 196/0/0/0 | 18 files; 0 parse errors/new findings; 7 baseline warnings |

Counts are passed/failed/skipped/not_run. Full suite: 159 product units, ten runner guards and 27 real loopback integration checks. Three parser defect characterizations remain. Python fixture-helper selftests: 20/0/0, separate from product acceptance. Final local/GitHub results are recorded in UPD-0102 evidence after verification.

## Continuity and limits

- Existing title-only archives remain untouched and unadopted. New names may produce separate downloads; changed feed URLs/titles can create a different folder. Persistent identity/aliases, migration and verified skips remain later work.
- File symbolic links are checked via attributes, but integration exercises privilege-free junctions (including at media leaf paths); live UNC shares and privileged file-symlink creation were not tested. No perfect concurrent-adversary sandbox claim.
- Body/audio validation, crash handling, incomplete exit semantics, privacy changes, manual launcher and broader recovery/release cases remain unrun/unimplemented. A011-A060 remain not_run. Logs retain legacy URL content pending their assigned task.
- A killed process may leave a unique temporary sibling; later runs preserve unclaimed files. Do not test against the real archive or private subscriptions.
- No merges, history rewriting, deletions of existing branches/issues, persistent system/settings/security changes, releases, purchases or deployments.
