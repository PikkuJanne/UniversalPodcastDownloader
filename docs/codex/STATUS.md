# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0104 only**.

- UPD-0001, UPD-0002 and UPD-0101 through UPD-0103: done; historical evidence preserved.
- UPD-0104: **done**; implementation pushed/verified, all local checks and both GitHub jobs passed.
- Next, ready: **UPD-0105 — protect and migrate legacy archives**; unstarted.
- Branch/upstream: `codex/upd-m1-safety` / `origin/codex/upd-m1-safety`.
- M1 [draft PR #2](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/2) remains stacked on `codex/upd-m0-foundation`; M0 draft PR #1 remains unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0104 evidence](evidence/UPD-0104.md), [schema and recovery](STATE_AND_MIGRATION.md), [decisions](DECISIONS.md).

## Verified checkout and implementation

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Original `UniversalPodcastDownloader-main` snapshot remains unmodified. Saved origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/remote/PR head was `1e6c6d7dbcb964175f5256075cc3ae0ce35ce1d0`; clean, fetch divergence 0/0. Predecessor [CI 36998204651](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/36998204651) passed both Windows jobs at that exact SHA. Continue the authenticated HTTPS fallback and command-scoped author identity from UPD-0001 evidence without persistent settings changes.

Schema-1 per-show history stores explicit feed aliases, stable feed-scoped RSS GUID/Atom ID/media URL fingerprints, recorded relative filenames, outcomes and byte/hash/transfer evidence. Titles and signed URLs remain mutable without renaming recorded publisher identities. Exact request URLs are kept in memory, not persisted in state. Unknown/corrupt/newer state is preserved and reported.

A brief output-root lock protects archive discovery/initialization; a show lock covers its run. Flushed owned state temporaries replace JSON atomically with a previous-generation backup. Prepared evidence precedes no-overwrite media placement, allowing post-placement crash reconciliation. Reruns hash selected media: unchanged files are verified skips, missing files are safely downloaded again, changed/unknown destinations are preserved conflicts. Original entry points, modes, safe paths, bounded media validation and unchanged audio remain.

## Actual checks

Windows 10.0.26300.0; Python 3.14.7; Git 2.56.0.windows.1; gh 2.97.0; Pester 5.7.1; PSScriptAnalyzer 1.24.0. Both supported engines passed 25 focused identity and 60 focused state checks. Real process/loopback history tests passed in focused runs on both engines. Final All: **379 passed / 0 failed / 0 skipped / 0 not_run on each engine**. Analysis: 32 files, zero parse errors/new findings, 5 PS7 / 4 PS5.1 baseline warnings. A001-A016 passed; A017-A060 remain not_run. The settled suite contains 311 product units, ten runner guards and 58 integrations (379 total). Two parser defect characterizations remain. Python fixture-helper selftests: 21/0/0, separate from product acceptance.

## Continuity and limits

- No automatic legacy adoption, alias-management CLI, backup restoration, schema migration, resume or orphan cleanup. Archives lacking new history remain untouched and may receive separate new downloads. UPD-0105 owns inventory/adoption policy.
- Digests prove consistency with recorded bytes, not publisher authenticity or complete decoding. Cross-refresh GUID reuse cannot reliably be distinguished from edits; changed fallback URLs can produce new identities.
- Corrupt/unreadable history or reparse output subdirectories stop discovery under the root. Atomic File.Replace support is required. Tests cover local process death, not hardware power loss/live UNC or privileged adversarial races.
- Existing logs retain legacy URLs pending UPD-0106. Broader network/launcher/cancellation behavior and later acceptance cases remain unimplemented. A017-A060 remain not_run.
- Tests used only synthetic/mock/loopback data and owned marked temporary outputs. No real archive/private subscriptions, manual launcher, merges, history rewriting, branch/issue deletion, persistent settings changes, releases, purchases or deployments.

## GitHub checkpoint

Implementation head `086aa2316f9d5cb01c6ec10a8bdf5a3d84b26da2` equals the independently verified remote branch. [CI 37000740167](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37000740167) passed both Windows Server 2022 jobs at that exact SHA: 379/0/0/0 each, 32 analyzed files, zero parse errors/new findings, 5/4 baseline warnings. CI engines: PowerShell 7.6.6 and Windows PowerShell 5.1.20348.5622. The final documentation-only checkpoint and its CI are verified separately in draft PR #2/final handoff, since this file cannot embed its own SHA. Stop after UPD-0104; the exact UPD-0105 continuation is in NEXT_THREAD_PROMPT.md.
