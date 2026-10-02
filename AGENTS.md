# UniversalPodcastDownloader — repository working rules

Keep the simple Windows downloader: original `.ps1`/`.bat` entry points, Latest/Custom/All, configurable output and unchanged original audio. Windows PowerShell 5.1 and PowerShell 7 on Windows remain supported. No runtime package manager, Python, FFmpeg, GUI, cloud account or daemon requirement.

Start each task with `docs/codex/STATUS.md`, the current `TASKS.json` entry, actual Git state and that task's prompt. One coherent task per thread. Load additional design documents only as linked. Preserve existing owner work and newer instructions. Codex project guidance is described in `docs/codex/SOURCES.md` [S1].

Never test against the real podcast archive. Use synthetic fixtures or explicit copies; cleanup only your own marked temporary test directories. Never overwrite completed/unknown media, publish secrets, weaken TLS, change machine-wide execution policy or require administrator rights for normal use.

Unit tests mock all external network. Integration tests bind loopback and run from temporary output paths. `tools/codex-handoff/` contains helper fixtures, NOT implemented product acceptance. Missing tests/runtimes are NOT passes. Record exact commands, versions and pass/fail/skip/not-run counts.

Development commands from the checkout root: run `pwsh -NoProfile -File ./scripts/Install-DevTools.ps1` once to install the versions and package hashes in `tools/devtools.json` into ignored `.dev-tools/Modules`. Then use `./scripts/Test.ps1 -Suite Unit`, `-Suite Integration`, or `-Suite All`; add `-Filter '*test name*'` for a focused selection. Run `./scripts/Analyze.ps1` in both supported Windows engines. Runners fail on missing tools, empty/all-skipped selections, test failures, parse errors or new lint findings. Existing lint warnings are bounded by rule, message, source context and count in `tools/lint-baseline.json`; reduce that baseline when fixing its causes.

Read `docs/codex/DEVELOPMENT.md` for exact fresh-process commands, Windows PowerShell 5.1 process-only execution-policy/module-path handling and test limitations. Pester/PSScriptAnalyzer are development-only; Python is required only for loopback integration tests. Dot-sourcing `UniversalPodcastDownloader.ps1` loads its helpers without starting the program; the downloader adds no runtime packages. Baseline characterization assertions document existing defects and do not satisfy future product acceptance.

Use verified-origin feature branches, explicit staging, normal commits/pushes and draft PRs. Before ending, update STATUS/TASKS/evidence/NEXT_THREAD_PROMPT, push and verify remote HEAD. No merges, rebases, resets, force-pushes, existing branch/issue deletion, settings changes, release tags/publication, purchases or deployments without explicit owner approval. No credential guessing or blanket `git add .`.

Prefer small reversible changes. Separate planning from side effects; preview must not write/download media. Keep logs safe to share. Record design changes in `docs/codex/DECISIONS.md`; no silent scope expansion or new runtime dependency.
