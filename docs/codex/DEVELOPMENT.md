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

Python is needed only by the synthetic loopback integration server. The local foundation checks use Python 3.14.7; CI pins the same version. Unit tests and the downloader do not require Python. No application runtime packages were introduced.

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

After UPD-0103, the full suite contains **277 checks per engine**:

| Group | Count | Scope |
| --- | ---: | --- |
| Product unit checks | 227 | Import safety, parsing, selection/web regressions, naming/containment, streamed requests, bounded body validation and transactional failures; includes three remaining parser characterizations |
| Runner guards | 10 | Missing tools, empty/filtered/all-skipped suites, pass/failure exit status, missing analyzer and a new lint warning |
| Product integration checks | 40 | HTML/RSS/Atom, mode/count/repeat behavior, hostile metadata, naming/junctions, invalid/lengthless bodies, actual interrupted transfers, finalization crashes and competing files |

Dot-sourcing `. .\UniversalPodcastDownloader.ps1` defines the existing helper functions and returns before startup preferences, logging, prompts and downloads. It is the import seam; no separate runtime module or package is required. A004 tests this against the actual script. Normal invocation with `&` retains the entry-point behavior.

Unit network access is mocked; synthetic external URLs use `.invalid`. Integration tests start an owned Python process bound to `127.0.0.1` on an ephemeral port, use marked temporary output directories and stop only their tracked child processes. They check media lengths/hashes, request counts and preservation on repeat runs. They never select the real archive or a private feed.

The three remaining unit characterizations assert Atom updated-before-published behavior, a missing Atom ID and first-enclosure selection even when it is video. Passing those assertions means the defect was reproduced. UPD-0102 converted the three naming defect characterizations into desired-behavior regressions; the parser acceptance cases remain pending.

UPD-0101 replaces both PS5.1 failure characterizations with desired-behavior regressions. Page/feed requests use `Invoke-PodcastWebRequest`, which always passes `-UseBasicParsing` to the existing network cmdlet. UPD-0103 media requests use the built-in .NET HttpClient with one streamed GET and headers from that same response; no DOM parsing occurs. The worker supplies no parsing default. Its HTML discovery check supplies only two exact application UI responses; the hidden child remains `-NonInteractive`, so the web cmdlet's legacy confirmation would fail. Harness control traffic retains its own safe parsing switch.

`Select-PodcastEpisode` returns an array for zero, one or many entries. Unit cases cover Latest, All and Custom counts of 1, 2 and 5, null input, invalid Custom counts, sorting, URL filtering, counts, download/skip progress, and explicit empty/no-enclosure errors. The entry point still rejects an empty feed or a feed without downloadable URLs before episode progress/media requests. A006 covers valid arithmetic; byte-level progress timing and other UX changes remain UPD-0303.

Focused UPD-0101 commands (use the same native child wrapper above for PS5.1):

```powershell
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Unit -Filter '*A00[67]*'
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Integration
```

Focused UPD-0102 unit coverage uses `-Suite Unit -Filter '*A0[01][089]*'` (A008, A009 and A010); use `-Suite Integration` for the real loopback path checks. The worker's two late-boundary hooks insert only tracked test junctions, at Preparing or after a real media response. These are adversarial filesystem fixtures, not alternate product behavior or network mocks.

Focused UPD-0103 units use `-Suite Unit -Filter '*A01[123]*'`: 68 checks, with 169 not_run. Run `-Suite Integration` for all 40 loopback cases. Transaction tests kill only their owned worker process after an actual stream write or immediately before/after File.Move. Process-local debugger breakpoints also insert a competing final file and verify the temporary stream has closed; product code has no test hook. Reruns preserve abandoned partials and begin a fresh request. Signature validation is bounded and is not full decoding.

These checks do not validate the whole application, launcher UX, catchable cancellation, private feeds, durable history/resume or future acceptance cases. Helper-server self-tests are a separate layer. See historical evidence for [UPD-0002](evidence/UPD-0002.md), [UPD-0101](evidence/UPD-0101.md) and [UPD-0102](evidence/UPD-0102.md). Current task evidence is [UPD-0103](evidence/UPD-0103.md).

## Static analysis policy

`Analyze.ps1` parses the runtime, runner and test PowerShell files in the selected engine, then runs PSScriptAnalyzer's warning/error rules. [tools/PSScriptAnalyzerSettings.psd1](../../tools/PSScriptAnalyzerSettings.psd1) excludes `PSAvoidUsingWriteHost` intentionally because the existing TUI and developer summaries use host output. Other default rules remain enabled.

[tools/lint-baseline.json](../../tools/lint-baseline.json) retains **6 existing warning allowances** from source commit `2ac82614493be7196c9ebee116f23fec07368b50`. Each allowance matches the exact repository-relative file, rule, message, surrounding source text and maximum occurrence count. The runner permits no new finding or parse error; moving a warning into unrelated source or increasing its count fails. Reduce/remove entries as later tasks fix their causes.

| Rule | Baseline count | Source context |
| --- | ---: | --- |
| `PSAvoidAssignmentToAutomaticVariable` | 1 | Regex result assigned to `$matches` |
| `PSAvoidOverwritingBuiltInCmdlets` | 1 | Existing `Write-Log` function |
| `PSAvoidUsingEmptyCatchBlock` | 2 | Feed-title extraction and date parsing |
| `PSUseBOMForUnicodeEncodedFile` | 1 | Existing runtime file encoding |
| `PSUseSingularNouns` | 1 | `Resolve-PodcastItems` |

The local PS7 analysis reports 6 known warnings; PS5.1 reports 5 because its analyzer built-in command profile does not emit the `Write-Log` override warning. Both observations have zero new findings and zero parse errors. These remain acknowledged legacy warnings. UPD-0102 removed two naming allowances; UPD-0103 removed the obsolete size lookup and its two allowances. The pure naming helpers use narrow, documented suppressions for their retained names. Analysis includes all bundled `src/` helpers, totaling 25 PowerShell files.

## CI and verified sources

[test.yml](../../.github/workflows/test.yml) runs Windows PowerShell 5.1 and PowerShell 7 on `windows-2022`, with a 15-minute job limit and two-job maximum. It triggers for pull requests, pushes to `main` and manual dispatch. It uses only `contents: read`, disables checkout credential persistence and has no artifact/release/publication step. Test/analysis counts and runtime inventory appear in the job summary. Workflow configuration alone is not evidence of a completed CI run; consult the task evidence and actual GitHub checks.

Action references were verified against the official release tags and source on 2026-10-02:

| Action | Full pinned commit | Official source |
| --- | --- | --- |
| checkout v7.0.1 | `3d3c42e5aac5ba805825da76410c181273ba90b1` | [Release](https://github.com/actions/checkout/releases/tag/v7.0.1), [commit](https://github.com/actions/checkout/commit/3d3c42e5aac5ba805825da76410c181273ba90b1) |
| setup-python v7.0.0 | `5fda3b95a4ea91299a34e894583c3862153e4b97` | [Release](https://github.com/actions/setup-python/releases/tag/v7.0.0), [commit](https://github.com/actions/setup-python/commit/5fda3b95a4ea91299a34e894583c3862153e4b97) |

Both use Node 24 and require an Actions runner version of at least 2.327.1. Reviewed sources include their action metadata, entry points, checkout credential handling and Python version selection/install flow. The pinned [Python 3.14.7 release](https://github.com/actions/python-versions/releases/tag/3.14.7-31064857500) supplies the Windows x64 distribution. CI does not install or upgrade the user's local Python or PowerShell.
