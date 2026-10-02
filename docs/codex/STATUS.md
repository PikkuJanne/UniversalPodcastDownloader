# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0106 only**.

- UPD-0001, UPD-0002 and UPD-0101 through UPD-0105: done; historical evidence preserved.
- UPD-0106: **done**; implementation pushed/verified, full local checks and both corrected GitHub jobs passed.
- Next, ready: **UPD-0107 — enforce preview and untrusted-input boundaries**; not started.
- Branch/upstream: `codex/upd-m1-safety` / `origin/codex/upd-m1-safety`.
- M1 [draft PR #2](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/2) remains stacked on `codex/upd-m0-foundation`; M0 draft PR #1 remains unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0106 evidence](evidence/UPD-0106.md), [diagnostic privacy and retention](DIAGNOSTICS.md), [schema and recovery](STATE_AND_MIGRATION.md).

## Verified checkout and implementation

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Preserved `UniversalPodcastDownloader-main` snapshot untouched. Origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/fetched remote/PR HEAD `cb2ad35a3e940a4be090149127acedb1b7641f4a` was clean with 0/0 divergence. [Predecessor CI 37004676869](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37004676869) passed both jobs at that SHA. Continue UPD-0001's authenticated HTTPS fallback and command-scoped public author identity without persistent settings changes.

Startup diagnostics now try local application data then TEMP before discovery. Logs use UTF-8, UTC and random run IDs; confirmed ordinary downloads retain the startup copy and switch to a same-run show log. Failures fall back safely without replacing the original operation error. URL displays show only host and opaque per-run IDs. Routine output omits raw titles/paths/GUIDs/headers/body/error text, while exact allowed application instructions remain useful. Original exceptions remain available in-process behind safe ErrorDetails.

Optional DiagnosticExportPath creates a restricted JSON from allowed metadata and fixed event codes. It never reads logs, history, checkpoints, config or media. WhatIf/legacy Preview create no diagnostics or exports. Confirmation buffers until accepted and then starts a fresh context. Normal preview objects omit internal request/history records; local legacy inventory still carries exact owner-review filenames and stays outside exports. Startup/show logs and exports are retained until owner deletion.

History schemas, legacy adoption/redownload/rollback, locks, atomic metadata replacement, completion evidence and original media safeguards remain. Original entry points and modes stay; no runtime dependency was added.

## Checks and limits

Local Windows 10.0.26300.0; PowerShell 7.6.5 / 5.1.26100.9444; bundled Python 3.12.14; Git 2.56.0.windows.1; gh 2.97.0; Pester 5.7.1; PSScriptAnalyzer 1.24.0. Focused privacy units and startup integrations pass both engines. Analysis: 44 files, zero parse errors/new findings, 5/4 baseline warnings. Settled suite: 521 checks (420 product units, ten runner guards, 91 integrations); two parser characterizations remain. Full local All: **521 passed / 0 failed / 0 skipped / 0 not_run per engine**. Corrected CI passed the same full counts on both supported engines. A001-A021 passed; A022-A060 remain not_run. Evidence retains the initial failed CI and the test-path/module-discovery corrections.

Use the existing bundled Python via process-only PATH; the ambient Windows alias had no usable runtime. CI Python remains pinned to 3.14.7. The test runner isolates both primary and fallback diagnostic locations in owned marked temporary data.

Logs expose hosts and episode fingerprints; review locally before sharing. Exports omit those fields as well as all free text. Older logs, local legacy review objects, archive paths/history, original exception inspection, shell history/transcripts and pre-initialization host errors can still be sensitive. No universal secret detector, automatic log cleanup, launcher/network/XML/retry/resume changes or new acceptance for those later tasks is claimed. Archive durability limits remain as documented by prior tasks.

Tests used synthetic/mock/loopback data only. No real archive/private feed, merge, history rewrite, persistent settings change, release, purchase or deployment. Stop after UPD-0106; NEXT_THREAD_PROMPT.md describes UPD-0107 after this task's completion.

## GitHub checkpoint

Implementation `713fa11615dc3a0a9b0d69f7351ceccd1908794e` is pushed and independently equals the remote. [CI 37009390015](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37009390015) passed both Windows Server 2022 jobs at that exact SHA: 521/0/0/0 each, 44 analyzed files, zero parse errors/new findings and 5/4 baseline warnings. CI engines: PowerShell 7.6.6 and Windows PowerShell 5.1.20348.5622. The final documentation checkpoint and exact-head CI are verified separately in draft PR #2/final handoff; the runtime is unchanged. Stop after UPD-0106 and use NEXT_THREAD_PROMPT.md for UPD-0107.
