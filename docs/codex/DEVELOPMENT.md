# Development and test commands

Run these commands from the actual checkout root. The downloader still uses its original PowerShell and batch entry points. Development setup is explicit and is never invoked when launching the downloader.

## Install the development tools once

```powershell
pwsh -NoProfile -File .\scripts\Install-DevTools.ps1
```

The setup script downloads the pinned PowerShell Gallery packages, checks their SHA-256 hashes and extracts them into ignored `.dev-tools/Modules`. It requires no administrator access, global module installation or repository trust change. Existing complete module directories are reused after their manifest versions are checked. A package hash mismatch or incomplete existing directory fails for inspection.

The authoritative download URLs and package hashes are in [tools/devtools.json](../../tools/devtools.json):

| Tool | Version | Official package and source |
| --- | --- | --- |
| Pester | 5.7.1 | [PowerShell Gallery](https://www.powershellgallery.com/packages/Pester/5.7.1), [Pester release](https://github.com/pester/Pester/releases/tag/5.7.1) |
| PSScriptAnalyzer | 1.24.0 | [PowerShell Gallery](https://www.powershellgallery.com/packages/PSScriptAnalyzer/1.24.0), [PowerShell release](https://github.com/PowerShell/PSScriptAnalyzer/releases/tag/1.24.0) |

Runners import these exact repository-local manifests. They fail if a pinned module is missing instead of selecting an unrelated globally installed version. Their output records PowerShell, OS and loaded module versions.

Python is needed only by the synthetic loopback integration server. The historical foundation checks used Python 3.14.7, which CI still pins. Current local checks use the bundled Python 3.12.14 runtime as described below. Unit tests and the downloader do not require Python. No application runtime packages were introduced.

## Focused and full runs

Use fresh PowerShell processes so imports and preferences from another session cannot affect the result:

```powershell
pwsh -NoProfile -File .\scripts\Test.ps1 -Suite Unit
pwsh -NoProfile -File .\scripts\Test.ps1 -Suite Integration
pwsh -NoProfile -File .\scripts\Test.ps1 -Suite All
pwsh -NoProfile -File .\scripts\Test.ps1 -Suite Unit -Filter '*A004*'
pwsh -NoProfile -File .\scripts\Analyze.ps1
```

`-Filter` matches Pester full test names using wildcards. It works with any suite. Missing suites/tools, no selected tests, a selection containing only skipped tests, and test failures all return a nonzero exit code. `RESULT` reports total, passed, failed, skipped and not-run counts; filtered-out cases must not be reported as passes. `-TestRoot` and `-ToolsPath` exist for the runner's synthetic failure checks, not as prerequisites for normal use.

### Windows PowerShell 5.1

Run the same commands under native Windows PowerShell. On this machine, a child launched from PowerShell 7 can inherit incompatible PowerShell 7 module paths. The following limits module discovery to native Windows modules while the child runs; repository-local Pester and PSScriptAnalyzer are still imported explicitly. Both environment and execution-policy changes are process-only.

```powershell
$nativeEngine = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$previousModulePath = $env:PSModulePath
try {
    $env:PSModulePath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\Modules'
    & $nativeEngine -NoProfile -NonInteractive -ExecutionPolicy Bypass -File .\scripts\Test.ps1 -Suite All
    $testExitCode = $LASTEXITCODE
    & $nativeEngine -NoProfile -NonInteractive -ExecutionPolicy Bypass -File .\scripts\Analyze.ps1
    $analysisExitCode = $LASTEXITCODE
    if ($testExitCode -ne 0 -or $analysisExitCode -ne 0) {
        throw "Checks failed: test exit=$testExitCode; analysis exit=$analysisExitCode"
    }
}
finally {
    $env:PSModulePath = $previousModulePath
}
```

Replace `-Suite All` with `Unit`, `Integration`, or a filtered command for a focused run. Do not change machine-wide execution policy or install modules globally to make these checks run.

## What the current suite covers

The UPD-0201 suite contains **705 checks per engine**. Consult [UPD-0201 evidence](evidence/UPD-0201.md) for the completed runs and exact snapshot counts; adding a test does not make an earlier run cover it.

| Group | Count | Scope |
| --- | ---: | --- |
| Product unit checks | 542 | Import safety, parsing, selection/web regressions, naming/containment, streamed requests, body validation, transactional failures, identity, state storage, legacy inventory/schema, migration safety, diagnostic privacy/lifecycle/export, presentation, shared URL/redirect policy, bounded metadata and explicit XML limits, transient/permanent retry policy, shared server-delay budgets, cancellation and private transport categories; includes two remaining parser characterizations |
| Runner guards | 10 | Missing tools, empty/filtered/all-skipped suites, pass/failure exit status, missing analyzer and a new lint warning |
| Product integration checks | 153 | HTML/RSS/Atom, modes, naming/junctions, media validation, interrupted transfers, state/finalization crashes, disk reconciliation, real process locks, legacy previews/adoption/redownload, metadata rollback, startup fallback, UTF-8, concurrent log IDs, preview persistence/request boundaries, redirect/credential policy, metadata limits and rejection of DTD/entity input, real server delays/retry counts, header/idle deadlines, progressing long responses and truncated-response recovery |

Dot-sourcing `. .\UniversalPodcastDownloader.ps1` defines the existing helper functions and returns before startup preferences, logging, prompts and downloads. It is the import seam; no separate runtime module or package is required. A004 tests this against the actual script. Normal invocation with `&` retains the entry-point behavior.

Unit network access is mocked; synthetic external URLs use `.invalid`. Integration tests start an owned Python process bound to `127.0.0.1` on an ephemeral port, use marked temporary output directories and stop only their tracked child processes. They check media lengths/hashes, request counts and preservation on repeat runs. They never select the real archive or a private feed.

The two remaining unit characterizations assert Atom updated-before-published behavior and first-enclosure selection even when it is video. Passing those assertions means the defect was reproduced. UPD-0104 replaced the missing Atom ID characterization with desired-behavior identity regressions; broader date/media parsing acceptance remains pending.

UPD-0101 replaced both PS5.1 failure characterizations with desired-behavior regressions by routing page/feed requests through `Invoke-PodcastWebRequest`, which supplied `-UseBasicParsing` to `Invoke-WebRequest`; that implementation remains recorded in its historical evidence. UPD-0107 retains the helper name but delegates metadata to the built-in .NET HttpClient with explicit HTTP(S), redirect and body limits. Media uses the same request policy and keeps its streamed completion checks from UPD-0103. Product requests no longer use the legacy web DOM parser. Feed XML uses bounded XmlReader settings with DTD prohibited and external resolution disabled. The HTML discovery worker supplies only the exact application UI responses; its hidden child remains `-NonInteractive`. Harness control traffic retains its own safe parsing switch.

`Select-PodcastEpisode` returns an array for zero, one or many entries. Unit cases cover Latest, All and Custom counts of 1, 2 and 5, null input, invalid Custom counts, sorting, URL filtering, counts, download/skip progress, and explicit empty/no-enclosure errors. The entry point still rejects an empty feed or a feed without downloadable URLs before episode progress/media requests. A006 covers valid arithmetic; byte-level progress timing and other UX changes remain UPD-0303.

Focused UPD-0101 commands (use the same native child wrapper above for PS5.1):

```powershell
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Unit -Filter '*A00[67]*'
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Integration
```

Focused UPD-0102 unit coverage uses `-Suite Unit -Filter '*A0[01][089]*'` (A008, A009 and A010); use `-Suite Integration` for the real loopback path checks. The worker's two late-boundary hooks insert only tracked test junctions, at Preparing or after a real media response. These are adversarial filesystem fixtures, not alternate product behavior or network mocks.

Focused UPD-0103 units use `-Suite Unit -Filter '*A01[123]*'`: 68 checks, with 169 not_run. Run `-Suite Integration` for all 40 loopback cases. Transaction tests kill only their owned worker process after an actual stream write or immediately before/after File.Move. Process-local debugger breakpoints also insert a competing final file and verify the temporary stream has closed; product code has no test hook. Reruns preserve abandoned partials and begin a fresh request. Signature validation is bounded and is not full decoding.

UPD-0104 identity units use `-Suite Unit -Filter '*A014*'` (25 checks); state units use `-Suite Unit -Filter '*A015*'` (60 checks). History integrations use `-Suite Integration -Filter '*A01[56]*'` (18 checks). Real child processes exercise both lock scopes and crashes before/after state replacement and final placement. Prepared evidence, changed/deleted media, signed URL refreshes and corrupt history are checked against actual disk and loopback requests. Full `-Suite All` uses the current suite inventory above; historical counts belong to their recorded task snapshots.

UPD-0105 legacy units use `-Suite Unit -Filter '*A01[78]*'` (71 checks); focused integration uses `-Suite Integration -Filter '*A01[78]*'` (25 checks, including the updated A009/A017 historical-folder guard and A016/A017 unknown-destination guard). CLI tests hash every copied original before/after operations, assert unchanged preview trees, require exact reviewed adoption digests, distinguish adoption from observed transfers, preserve remote-changed originals, choose one redownload from a multi-episode feed, and restore only metadata from explicit checkpoints. Ordinary WhatIf now performs no filesystem writes; early invalid feeds/identities leave the output root absent. Schema-1 behavior stays supported alongside explicit schema-2 migration.

UPD-0106 units use `-Suite Unit -Filter '*A020*'` (38 checks); diagnostics integrations use `-Suite Integration -Filter '*A019/A021*'` (10 checks). They cover URL/userinfo/path/query/header redaction, native error formatting while preserving the exception, strict UTF-8 including emoji, unique concurrent log names, startup fallback, failed append/export, restricted export fields and no-write previews. The runner isolates both LOCALAPPDATA and TEMP/TMP in a short, exclusively created and marked temporary directory (under an existing RUNNER_TEMP when supplied by CI, otherwise platform TEMP); it restores the process environment and removes only that verified owned directory. No user diagnostic directory is a test target.

### UPD-0107 focused checks

Use fresh processes and the native child wrapper above for Windows PowerShell 5.1:

```powershell
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Unit -Filter '*A023*'
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Unit -Filter '*A02[34]*'
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Unit -Filter '*A024*'
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Integration -Filter '*A02[234]*'
pwsh -NoProfile -NonInteractive -File .\scripts\Analyze.ps1
python -B -m unittest discover -s .\tools\codex-handoff\selftests -v
```

The current focused checkpoint records these passed/failed/skipped/not-run counts. The early XML result used a smaller intermediate suite; its filtered-out count is preserved rather than recalculated against the settled inventory.

| Check | Engine | Passed / failed / skipped / not_run |
| --- | --- | --- |
| A023 request-policy units | PS7 and PS5.1, each | 66 / 0 / 0 / 435 |
| Combined A023/A024 units | PS7 | 76 / 0 / 0 / 425 |
| A024 XML units, early snapshot | PS5.1 | 10 / 0 / 0 / 432 |
| A022/A023/A024 integrations | PS7 and PS5.1, each | 20 / 0 / 0 / 91 |
| Python fixture-helper self-tests | Bundled Python 3.12.14 | 28 passed; separate from product acceptance |

The new cases check both initial and redirected request targets, disabled request credentials/cookies, unchanged TLS/proxy defaults, bounded metadata and decoding, DTD/entity rejection, XML structure limits and valid-feed compatibility. Actual entry-point previews leave archive/diagnostic trees unchanged, create no export or lock, make no planned enclosure request and print no completion banner. Metadata retrieval can still receive arbitrary content from the supplied URL or a permitted redirect; the no-media guarantee covers planned enclosure requests and media-transfer execution. No keep-awake feature is introduced. The exact policy is in [INPUT_BOUNDARIES.md](INPUT_BOUNDARIES.md).

This task's local Python is the bundled **3.12.14** runtime. The ambient Windows `python` alias did not resolve a usable runtime; prefix the bundled Python directory to the calling process PATH for local integrations. CI remains pinned to Python **3.14.7**. No Python install or global PATH change is required. Exact local path and commands appear in the evidence.

These checks do not validate the whole application, launcher UX, catchable cancellation, private feeds or future acceptance cases. Helper-server self-tests are a separate layer. Historical evidence preserves earlier snapshots. Current task evidence is [UPD-0202](evidence/UPD-0202.md), with the settled 829-check inventory (642 units and 187 integrations), resume recovery/ownership cases, exact commands and local/CI results. [Resume policy](RESUME_POLICY.md) documents the tested recovery scope and conservative limits.

## Static analysis policy

`Analyze.ps1` parses the runtime, runner and test PowerShell files in the selected engine, then runs PSScriptAnalyzer's warning/error rules. [tools/PSScriptAnalyzerSettings.psd1](../../tools/PSScriptAnalyzerSettings.psd1) excludes `PSAvoidUsingWriteHost` intentionally because the existing TUI and developer summaries use host output. Other default rules remain enabled.

[tools/lint-baseline.json](../../tools/lint-baseline.json) retains **5 existing warning allowances** from source commit `2ac82614493be7196c9ebee116f23fec07368b50`. Each allowance matches the exact repository-relative file, rule, message, surrounding source text and maximum occurrence count. The runner permits no new finding or parse error; moving a warning into unrelated source or increasing its count fails. Reduce/remove entries as later tasks fix their causes.

| Rule | Baseline count | Source context |
| --- | ---: | --- |
| `PSAvoidAssignmentToAutomaticVariable` | 1 | Regex result assigned to `$matches` |
| `PSAvoidOverwritingBuiltInCmdlets` | 1 | Existing `Write-Log` function |
| `PSAvoidUsingEmptyCatchBlock` | 2 | Feed-title extraction and date parsing |
| `PSUseSingularNouns` | 1 | `Resolve-PodcastItems` |

The known baseline is 5 warnings on PS7 and 4 on PS5.1, whose analyzer built-in command profile does not emit the `Write-Log` override warning. These remain acknowledged legacy warnings. UPD-0102 removed two naming allowances; UPD-0103 removed the obsolete size lookup and its two allowances; UPD-0104 added the main script's UTF-8 BOM and removed that allowance. Pure helpers use narrow, documented suppressions for their retained names. The UPD-0107 analysis covers 50 PowerShell files. Full analysis has zero parse errors and zero new findings on both engines, with 5 baseline warnings on PS7 and 4 on native PS5.1.

## CI and verified sources

[test.yml](../../.github/workflows/test.yml) runs Windows PowerShell 5.1 and PowerShell 7 on `windows-2022`, with a 15-minute job limit and two-job maximum. It triggers for pull requests, pushes to `main` and manual dispatch. It uses only `contents: read`, disables checkout credential persistence and has no artifact/release/publication step. Test/analysis counts and runtime inventory appear in the job summary. Workflow configuration alone is not evidence of a completed CI run; consult the task evidence and actual GitHub checks.

Action references were verified against the official release tags and source on 2026-10-02:

| Action | Full pinned commit | Official source |
| --- | --- | --- |
| checkout v7.0.1 | `3d3c42e5aac5ba805825da76410c181273ba90b1` | [Release](https://github.com/actions/checkout/releases/tag/v7.0.1), [commit](https://github.com/actions/checkout/commit/3d3c42e5aac5ba805825da76410c181273ba90b1) |
| setup-python v7.0.0 | `5fda3b95a4ea91299a34e894583c3862153e4b97` | [Release](https://github.com/actions/setup-python/releases/tag/v7.0.0), [commit](https://github.com/actions/setup-python/commit/5fda3b95a4ea91299a34e894583c3862153e4b97) |

Both use Node 24 and require an Actions runner version of at least 2.327.1. Reviewed sources include their action metadata, entry points, checkout credential handling and Python version selection/install flow. The pinned [Python 3.14.7 release](https://github.com/actions/python-versions/releases/tag/3.14.7-31064857500) supplies the Windows x64 distribution. CI does not install or upgrade the user's local Python or PowerShell.

## UPD-0201 focused checks

Use fresh processes and the native wrapper above:

```powershell
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A02[56]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A025/A026*'
```

The settled selections contain 51 units and 42 integrations. Units inject clock/delay and simulate token-ignoring streams; integrations use real loopback request counts/timestamps, exact media hashes and failed/completed history evidence. Fixture selftests remain separate helper checks. The final Framework regression permits bare IOException retries only while reading the HTTP response source, leaving local destination/history errors permanent. Retry settings and elapsed-window semantics are in [TRANSPORT_POLICY.md](TRANSPORT_POLICY.md).
