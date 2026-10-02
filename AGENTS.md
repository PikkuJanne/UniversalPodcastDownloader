# UniversalPodcastDownloader — repository working rules

Keep the simple Windows downloader: original `.ps1`/`.bat` entry points, Latest/Custom/All, configurable output and unchanged original audio. Windows PowerShell 5.1 and PowerShell 7 on Windows remain supported. No runtime package manager, Python, FFmpeg, GUI, cloud account or daemon requirement.

Start each task with `docs/codex/STATUS.md`, the current `TASKS.json` entry, actual Git state and that task's prompt. One coherent task per thread. Load additional design documents only as linked. Preserve existing owner work and newer instructions. Codex project guidance is described in `docs/codex/SOURCES.md` [S1].

Never test against the real podcast archive. Use synthetic fixtures or explicit copies; cleanup only your own marked temporary test directories. Never overwrite completed/unknown media, publish secrets, weaken TLS, change machine-wide execution policy or require administrator rights for normal use.

Unit tests mock all external network. Integration tests bind loopback and run from temporary output paths. `tools/codex-handoff/` contains helper fixtures, NOT implemented product acceptance. UPD-0002 must establish and document the real Pester/lint runners. Missing tests/runtimes are NOT passes. Record exact commands, versions and pass/fail/skip/not-run counts.

Use verified-origin feature branches, explicit staging, normal commits/pushes and draft PRs. Before ending, update STATUS/TASKS/evidence/NEXT_THREAD_PROMPT, push and verify remote HEAD. No merges, rebases, resets, force-pushes, existing branch/issue deletion, settings changes, release tags/publication, purchases or deployments without explicit owner approval. No credential guessing or blanket `git add .`.

Prefer small reversible changes. Separate planning from side effects; preview must not write/download media. Keep logs safe to share. Record design changes in `docs/codex/DECISIONS.md`; no silent scope expansion or new runtime dependency.
