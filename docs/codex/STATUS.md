# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0106 only**.

- UPD-0001, UPD-0002 and UPD-0101 through UPD-0105: done; historical evidence preserved.
- UPD-0106: **in_progress**; implementation and focused checks complete, final full local/CI verification and synchronization pending.
- Next: **UPD-0107 — enforce preview and untrusted-input boundaries**; not started.
- Branch/upstream: `codex/upd-m1-safety` / `origin/codex/upd-m1-safety`.
- M1 [draft PR #2](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/2) remains stacked on `codex/upd-m0-foundation`; M0 draft PR #1 remains unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0106 evidence](evidence/UPD-0106.md), [diagnostic privacy and retention](DIAGNOSTICS.md), [schema and recovery](STATE_AND_MIGRATION.md).

## Verified checkout and implementation

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Preserved `UniversalPodcastDownloader-main` snapshot untouched. Origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/fetched remote/PR HEAD `cb2ad35a3e940a4be090149127acedb1b7641f4a` was clean with 0/0 divergence. [Predecessor CI 37004676869](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37004676869) passed both jobs at that SHA. Continue UPD-0001's authenticated HTTPS fallback and command-scoped public author identity without persistent settings changes.

Startup diagnostics now try local application data then TEMP before discovery. Logs use UTF-8, UTC and random run IDs; confirmed ordinary downloads retain the startup copy and switch to a same-run show log. Failures fall back safely without replacing the original operation error. URL displays show only host and opaque per-run IDs. Routine output omits raw titles/paths/GUIDs/headers/body/error text, while exact allowed application instructions remain useful. Original exceptions remain available in-process behind safe ErrorDetails.

Optional DiagnosticExportPath creates a restricted JSON from allowed metadata and fixed event codes. It never reads logs, history, checkpoints, config or media. WhatIf/legacy Preview create no diagnostics or exports. Confirmation buffers until accepted and then starts a fresh context. Normal preview objects omit internal request/history records; local legacy inventory still carries exact owner-review filenames and stays outside exports. Startup/show logs and exports are retained until owner deletion.

History schemas, legacy adoption/redownload/rollback, locks, atomic metadata replacement, completion evidence and original media safeguards remain. Original entry points and modes stay; no runtime dependency was added.

## Checks and limits

Local Windows 10.0.26300.0; PowerShell 7.6.5 / 5.1.26100.9444; bundled Python 3.12.14; Git 2.56.0.windows.1; gh 2.97.0; Pester 5.7.1; PSScriptAnalyzer 1.24.0. Focused privacy units and startup integrations pass both engines. Analysis: 44 files, zero parse errors/new findings, 5/4 baseline warnings. Settled suite: 521 checks (420 product units, ten runner guards, 91 integrations); two parser characterizations remain. Full local/CI outcomes are pending and will be recorded before handoff. Later A022-A060 remain not_run.

Use the existing bundled Python via process-only PATH; the ambient Windows alias had no usable runtime. CI Python remains pinned to 3.14.7. The test runner isolates both primary and fallback diagnostic locations in owned marked temporary data.

Logs expose hosts and episode fingerprints; review locally before sharing. Exports omit those fields as well as all free text. Older logs, local legacy review objects, archive paths/history, original exception inspection, shell history/transcripts and pre-initialization host errors can still be sensitive. No universal secret detector, automatic log cleanup, launcher/network/XML/retry/resume changes or new acceptance for those later tasks is claimed. Archive durability limits remain as documented by prior tasks.

Tests used synthetic/mock/loopback data only. No real archive/private feed, merge, history rewrite, persistent settings change, release, purchase or deployment. Stop after UPD-0106; NEXT_THREAD_PROMPT.md describes UPD-0107 after this task's completion.
