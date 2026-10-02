# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0105 only**.

- UPD-0001, UPD-0002 and UPD-0101 through UPD-0104: done; historical evidence preserved.
- UPD-0105: **done**; implementation pushed/verified, focused local checks and both full GitHub jobs passed.
- Next, ready: **UPD-0106 — make startup diagnostics private and reliable**; unstarted.
- Branch/upstream: `codex/upd-m1-safety` / `origin/codex/upd-m1-safety`.
- M1 [draft PR #2](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/2) remains stacked on `codex/upd-m0-foundation`; M0 draft PR #1 remains unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0105 evidence](evidence/UPD-0105.md), [schema and recovery](STATE_AND_MIGRATION.md), [decisions](DECISIONS.md).

## Verified checkout and implementation

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Original `UniversalPodcastDownloader-main` snapshot remains unmodified. Saved origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/remote/PR head was `b67ef91052a2ff2c772d304c3ccde43614077113`; clean, fetch divergence 0/0. Predecessor [CI 37001209085](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37001209085) passed both Windows jobs at that exact SHA. Continue the authenticated HTTPS fallback and command-scoped author identity from UPD-0001 evidence without persistent settings changes.

UPD-0105 adds read-only inventory of existing show folders and explicit per-episode adoption/redownload. Adoption requires an episode ID, exact filename and reviewed SHA-256, then rechecks those bytes under a read guard through commit. Schema 2 records adopted confidence separately from observed transfers. Matching adopted files remain adopted; missing/changed adopted files require review without automatic replacement. Explicit redownload uses a separate safe filename and preserves all originals. Normal runs flag unclaimed legacy media before implicitly downloading duplicates. Both normal WhatIf and legacy Preview avoid filesystem, log, lock and media writes.

Each applied legacy operation saves a unique full-state checkpoint. Explicit rollback restores that snapshot's metadata at a new generation, preserves media/checkpoints, and retains schema 2. Existing version-1 archives remain supported without automatic promotion. Root/show locks, strict JSON validation, atomic state replacement, backups, stable identities and prepared-transfer reconciliation remain in use. Original entry points and modes remain; no runtime dependency was added.

## Actual checks

Windows 10.0.26300.0; PowerShell 7.6.5 / Windows PowerShell 5.1.26100.9444; Python 3.14.7; Git 2.56.0.windows.1; gh 2.97.0; Pester 5.7.1; PSScriptAnalyzer 1.24.0. Focused schema, inventory, migration safety and actual CLI tests pass on both engines. Analysis: 39 files, zero parse errors/new findings, 5 PS7 / 4 PS5.1 baseline warnings. Python fixture-helper selftests: 22 passed, separate from product acceptance. Full implementation CI passed **473 / 0 / 0 / 0** (passed/failed/skipped/not_run) on each engine. Evidence distinguishes intermediate local failures from corrected checks and complete CI results.

The settled suite has 473 checks (382 product units, ten runner guards, 81 integrations). Two parser defect characterizations remain. A017-A018 tests check every original file's bytes/hash, unchanged preview trees, explicit mapping/digest requirements, ambiguity, remote content changes, separate redownloads and metadata-only rollback. Later acceptance A019-A060 remains not_run.

## Continuity and limits

- Digests prove consistency with reviewed/recorded bytes, not publisher authenticity or complete decoding. Adopted files have no observed transfer-completeness proof. Names are candidate hints only; exact owner selection can resolve ambiguous mappings.
- Inventory reads immediate files only. The selected LegacyPath must be an ordinary existing immediate show directory under OutputPath. Title matching blocks accidental duplication but never establishes feed ownership. Changed fallback URLs remain inherently ambiguous; alias management is not implemented.
- Rollback restores the entire chosen metadata snapshot, including removing later records, while preserving all media. Schema 2, feed association and metadata files remain after rollback. Checkpoints accumulate for owner review. No automatic corrupt-state restore, resume or orphan cleanup.
- Corrupt/unreadable history or reparse output subdirectories stop discovery. Atomic File.Replace support is required. Local process-crash tests do not establish hardware power-loss/live UNC or privileged adversarial-race guarantees.
- Existing logs/console/error URL privacy remains UPD-0106. Broader network/XML, launcher/cancellation and later acceptance remain pending. Preview must stay free of diagnostic writes during that next task.
- Tests used only synthetic/mock/loopback data and owned marked outputs. No real archive/private subscriptions, manual launcher, merges, history rewriting, branch/issue deletion, persistent settings changes, releases, purchases or deployments.

## GitHub checkpoint

Implementation `a24b7781b9861d5a51af6ee7b7e7ee5c305c4e6f` is pushed and independently matches the implementation remote checkpoint. [Implementation CI 37003884285](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37003884285) passed both Windows Server 2022 jobs at that exact SHA: 473/0/0/0 each, 39 analyzed files, zero parse errors/new findings, 5/4 baseline warnings. CI engines: PowerShell 7.6.6 and Windows PowerShell 5.1.20348.5622. The final documentation checkpoint and its exact-head CI are verified separately in draft PR #2/final handoff, since this file cannot embed its own SHA. Stop after UPD-0105; the exact UPD-0106 continuation is in NEXT_THREAD_PROMPT.md.
