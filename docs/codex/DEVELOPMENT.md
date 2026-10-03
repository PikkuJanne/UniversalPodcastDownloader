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

## Historical suite coverage and current checks

The historical UPD-0201 suite contained **705 checks per engine**. Consult [UPD-0201 evidence](evidence/UPD-0201.md) for that snapshot. The current UPD-0303 inventory is **1,534 checks per engine**, detailed in its section below; adding a test does not make an earlier run cover it. The following table is the historical UPD-0201 breakdown.

| Group | Count | Scope |
| --- | ---: | --- |
| Product unit checks | 542 | Import safety, parsing, selection/web regressions, naming/containment, streamed requests, body validation, transactional failures, identity, state storage, legacy inventory/schema, migration safety, diagnostic privacy/lifecycle/export, presentation, shared URL/redirect policy, bounded metadata and explicit XML limits, transient/permanent retry policy, shared server-delay budgets, cancellation and private transport categories; includes two remaining parser characterizations |
| Runner guards | 10 | Missing tools, empty/filtered/all-skipped suites, pass/failure exit status, missing analyzer and a new lint warning |
| Product integration checks | 153 | HTML/RSS/Atom, modes, naming/junctions, media validation, interrupted transfers, state/finalization crashes, disk reconciliation, real process locks, legacy previews/adoption/redownload, metadata rollback, startup fallback, UTF-8, concurrent log IDs, preview persistence/request boundaries, redirect/credential policy, metadata limits and rejection of DTD/entity input, real server delays/retry counts, header/idle deadlines, progressing long responses and truncated-response recovery |

Dot-sourcing `. .\UniversalPodcastDownloader.ps1` defines the existing helper functions and returns before startup preferences, logging, prompts and downloads. It is the import seam; no separate runtime module or package is required. A004 tests this against the actual script. Normal invocation with `&` retains the entry-point behavior.

Unit network access is mocked; synthetic external URLs use `.invalid`. Integration tests start an owned Python process bound to `127.0.0.1` on an ephemeral port, use marked temporary output directories and stop only their tracked child processes. They check media lengths/hashes, request counts and preservation on repeat runs. They never select the real archive or a private feed.

The UPD-0201 snapshot retained two unit characterizations: Atom updated-before-published behavior and first-enclosure selection even when it is video. UPD-0204 replaced the date characterization with desired-behavior A034 regressions; UPD-0205 replaced the media characterization with desired A035 regressions. UPD-0104 replaced the missing Atom ID characterization with desired-behavior identity regressions. Historical passing characterizations reproduced defects and did not satisfy future acceptance.

UPD-0101 replaced both PS5.1 failure characterizations with desired-behavior regressions by routing page/feed requests through `Invoke-PodcastWebRequest`, which supplied `-UseBasicParsing` to `Invoke-WebRequest`; that implementation remains recorded in its historical evidence. UPD-0107 retains the helper name but delegates metadata to the built-in .NET HttpClient with explicit HTTP(S), redirect and body limits. Media uses the same request policy and keeps its streamed completion checks from UPD-0103. Product requests no longer use the legacy web DOM parser. Feed XML uses bounded XmlReader settings with DTD prohibited and external resolution disabled. The HTML discovery worker supplies only the exact application UI responses; its hidden child remains `-NonInteractive`. Harness control traffic retains its own safe parsing switch.

`Select-PodcastEpisode` returns an array for zero, one or many entries. Unit cases cover Latest, All and Custom counts of 1, 2 and 5, null input, invalid Custom counts, sorting, URL filtering, counts, download/skip progress, and explicit empty/no-enclosure errors. The entry point still rejects an empty feed or a feed without downloadable URLs before episode progress/media requests. A006 covers valid arithmetic; UPD-0303 adds byte-level progress timing and verified-success presentation.

Focused UPD-0101 commands (use the same native child wrapper above for PS5.1):

```powershell
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Unit -Filter '*A00[67]*'
pwsh -NoProfile -NonInteractive -File .\scripts\Test.ps1 -Suite Integration
```

Focused UPD-0102 unit coverage uses `-Suite Unit -Filter '*A0[01][089]*'` (A008, A009 and A010); use `-Suite Integration` for the real loopback path checks. The worker's two late-boundary hooks insert only tracked test junctions, at Preparing or after a real media response. These are adversarial filesystem fixtures, not alternate product behavior or network mocks.

Focused UPD-0103 units use `-Suite Unit -Filter '*A01[123]*'`: 68 checks, with 169 not_run. Run `-Suite Integration` for all 40 loopback cases. Transaction tests kill only their owned worker process after an actual stream write or immediately before/after File.Move. Process-local debugger breakpoints also insert a competing final file and verify the temporary stream has closed; product code has no test hook. Reruns preserve abandoned partials and begin a fresh request. Signature validation is bounded and is not full decoding.

UPD-0104 identity units use `-Suite Unit -Filter '*A014*'` (25 checks); state units use `-Suite Unit -Filter '*A015*'` (60 checks). History integrations use `-Suite Integration -Filter '*A01[56]*'` (18 checks). Real child processes exercise both lock scopes and crashes before/after state replacement and final placement. Prepared evidence, changed/deleted media, signed URL refreshes and corrupt history are checked against actual disk and loopback requests. Full `-Suite All` uses the current inventory in the UPD-0303 section; historical counts belong to their recorded task snapshots.

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

These earlier checks do not validate the whole application, launcher UX, catchable cancellation, private feeds or future acceptance cases. Helper-server self-tests are a separate layer. Historical evidence preserves earlier snapshots. Current task evidence is [UPD-0206](evidence/UPD-0206.md), with 1,272 checks (1,012 units and 260 integrations), bounded pagination/partial-catalogue results, intermediate failures, exact commands and results. [Resume policy](RESUME_POLICY.md) documents the tested recovery scope and conservative limits.

## Static analysis policy

`Analyze.ps1` parses the runtime, runner and test PowerShell files in the selected engine, then runs PSScriptAnalyzer's warning/error rules. [tools/PSScriptAnalyzerSettings.psd1](../../tools/PSScriptAnalyzerSettings.psd1) excludes `PSAvoidUsingWriteHost` intentionally because the existing TUI and developer summaries use host output. Other default rules remain enabled.

[tools/lint-baseline.json](../../tools/lint-baseline.json) retains **2 surviving warning allowances** from the reviewed baseline source `2ac82614493be7196c9ebee116f23fec07368b50`. Each allowance matches the exact repository-relative file, rule, message, surrounding source text and maximum occurrence count. The runner permits no new finding or parse error; moving a warning into unrelated source or increasing its count fails. Reduce/remove entries as later tasks fix their causes.

| Rule | Baseline count | Source context |
| --- | ---: | --- |
| `PSAvoidOverwritingBuiltInCmdlets` | 1 | Existing `Write-Log` function |
| `PSAvoidUsingEmptyCatchBlock` | 1 | Feed-title extraction |

The current baseline emits 2 warnings on PS7 and 1 on native PS5.1, whose analyzer built-in command profile does not emit the `Write-Log` override warning. These remain acknowledged legacy warnings. Earlier tasks removed obsolete naming/size/BOM allowances; UPD-0203 reduced the baseline to 3/2 and UPD-0204 to 2/1. Pure helpers use narrow, documented suppressions for retained names. Final UPD-0205 analysis covers 69 PowerShell files with zero parse errors and zero new findings on both engines.

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

## UPD-0203 focused checks

Use fresh processes and the native/Python wrappers above:

```powershell
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A03[12]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A0[03][067]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A019/A021*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A01[01]:*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite All
pwsh -NoProfile -NonInteractive -File ./scripts/Analyze.ps1
```

Settled inventory is 936: 720 units (710 product units including two parser characterizations plus ten runner guards) and 216 integrations. Final settled Unit passed 720/0/0/0 per engine. A006/A007/A030 focus passed 43/0/0/173 each; repaired diagnostics 10/0/0/206 each; affected archive/transaction units 12/0/0/708 each. Final analyzers passed 63 files with no parse errors/new findings and reduced 3/2 baseline warnings; fixture helpers: 44 passed separately. The 935-case broad local snapshots exposed four stale mock fixtures and three stale diagnostic expectations; evidence records those failures and their corrected reruns separately from full exact-source CI.

The new 78 units and 29 loopback cases cover root classification, empty/malformed/unsupported diagnostics, redirects/base/entities, attribute variants, deduplication, ambiguity and numbered choices, retained response reuse, original identity, static scanner safety and preview media/archive boundaries. Startup sink/export boundaries also retain the broader diagnostics regressions. See [shared discovery policy](FEED_DISCOVERY.md) and [task evidence](evidence/UPD-0203.md) for the explicit scanner subset and exact source/engine results.

## UPD-0204 focused checks

Use fresh processes and the native/Python wrappers above:

```powershell
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A03[34]:*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A033/A034*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A01[5678]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Analyze.ps1
python -B -m unittest discover -s tools/codex-handoff/selftests -v
```

Settled inventory: 1,093 checks, comprising 870 units (860 product units, including one media-selection characterization, plus ten runner guards) and 223 integrations. Final local Unit passed 870/0/0/0 each. A033/A034 units passed 151/0/0/719 each; date integrations 7/0/0/216 each; affected history/legacy integrations 42/0/0/181 each. Final analyzers: 66 files, no parse errors/new findings, reduced 2/1 baseline warnings. Helper selftests: 46/0, separate from acceptance. No local full All run was performed; exact-head CI is recorded separately in the PR/final handoff.

The new 150 units and seven loopback cases cover invariant parsing across en-US/de-DE/fi-FI, UTC offsets/precision/ranges, published-first Atom selection, date namespaces with compatible IDs, explicit tied/missing order, UTC midnight filename dates, empty/old-priority legacy hints, no-write preview and recorded archive preservation. Prior probes failed as intended before implementation. See [date policy](PUBLICATION_DATES.md) and [UPD-0204 evidence](evidence/UPD-0204.md) for commands, versions, snapshots and limitations.

## UPD-0205 focused checks (historical snapshot)

Use fresh processes and the native/Python wrappers above:

```powershell
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A035*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A012 bounded conservative*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A035/A036*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A01[125678]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A02[78]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Analyze.ps1
python -B -m unittest discover -s tools/codex-handoff/selftests -v
```

The UPD-0205 snapshot inventory was 1,184 checks: 940 units (930 product units, ten runner guards) and 244 integrations, with no remaining media-selection characterization. Final local Unit passed 940/0/0/0 each; A035 units 71/0/0/869 each; existing A012 39/0/0/901 each; format integrations 21/0/0/223 each. Affected archive/body/history/legacy selected 54 at the preceding 243-integration snapshot, passing 54/0/0/189 each. Resume selected 34 at 244 integrations; its results are in the task evidence. Final analyzers: 69 files, parse/new findings zero, baseline 2 PS7/1 native; helper selftests 51/0 separate from acceptance. No local full All run was performed; final exact-head full CI is recorded separately in the PR/final handoff.

Audio probes exercise bounded MP3/M4A/Ogg/WAVE/FLAC recognition and exact byte hashes, header/URL conflicts, text/video/ambiguity rejection, original identity/budgets, established path authority, unfinished-allocation correction, separate legacy allocation, no-write preview and prepared/checkpoint/no-overwrite boundaries. M4A/Ogg probes establish structural indications rather than full decoding or integrity. See [audio policy](AUDIO_FORMATS.md) and [UPD-0205 evidence](evidence/UPD-0205.md) for exact commands, intermediate failures, snapshots and remaining limits.

## UPD-0206 pagination checks (2026-10-02)

`src/FeedPagination.ps1` traverses one advertised next/prev-archive chain before mode selection, with 20 pages by default (configurable 1-100), 10,000 raw-entry and 32 MiB accepted-character bounds. Feed-level namespace/rel/type/base validation, cycles, duplicates, contradiction guards, later failures and no-write preview are documented in [FEED_PAGINATION.md](FEED_PAGINATION.md). No runtime package is added; current incomplete errors remain nonzero pending UPD-0301's result/launcher contract.

Settled inventory: **1,272 checks** = **1,012 units** (1,002 product units plus ten runner guards) + **260 integrations**. Full local Unit passed 1,012/0/0/0 each, PS7 231.20 s/native 173.05 s. Pagination units passed 72/0/0/940 each; pagination loopback passed 16/0/0/244 each, PS7 146.11 s/native 105.23 s. Existing A030 discovery passed 29/0/0/231 each, PS7 167.83 s/native 94.01 s. Existing A015/A016 history passed 18/0/0/242 each; durations are in task evidence. A031-A036 date/audio checks passed 28/0/0/218 each at the earlier 246-integration snapshot. Counts are passed/failed/skipped/not_run. Final analyzers: 72 files, parse/new findings zero, baseline 2 PS7/1 native. Fixture helper selftests 57/0 are separate from product acceptance. No local full All was performed; final exact-head full CI is recorded in the PR/final handoff.

Fresh-process focus commands use `-Suite Unit -Filter '*A037/A038*'` and `-Suite Integration -Filter '*A037/A038*'`. Affected filters are `-Suite Integration -Filter '*A030*'`, `'*A01[56]*'` and `'*A03[1-6]*'`. Use the native wrapper/process-only Python PATH above; do not run helpers against real archives. Local Windows 10.0.26300.0, PowerShell 7.6.5/native 5.1.26100.9444, Pester 5.7.1/PSScriptAnalyzer 1.24.0 and bundled Python 3.12.14. [UPD-0206 evidence](evidence/UPD-0206.md) records baseline defects, snapshot changes, exact commands, outcomes and limitations. A001-A038 passed; A039-A060 remain not_run. Stop after UPD-0206; UPD-0301 is next.

## UPD-0301 CLI/results checks (2026-10-03)

Dot-source the root script and call Invoke-PodcastRun for one private RunResult without host exit; use entry-script PassThru only when capturing its boundary result. See [CLI_RESULTS.md](CLI_RESULTS.md). CustomCount alone implies Custom, conflicting explicit Latest/All fails, and NonInteractive forbids missing feed/input prompts. Exit 0/1/2/130 classification preserves catalogue gaps and caught cancellation. Explicit legacy review/action inventory remains private in LegacyResult, including local WhatIf review.

Settled inventory **1,364** = **1,075 units** (1,065 product/ten runner guards) + **289 integrations**. Full local Unit passed 1,075/0/0/0 each, PS7 194.23 s/native 138.73 s. New process/result focus passed 21/0/0/268 each, PS7 105.00 s/native 58.99 s. Broad snapshots were PS7 266/23/0/0 and native 266/23/0/0, retaining cached old presentation assertions; settled affected reruns passed PS7 42/0/0/247 and native 197/0/0/92. Counts are passed/failed/skipped/not_run. See evidence for exact timings and filters. Whole-tree analysis: 78 files, parse/new 0, baseline PS7 2/native 1; helpers 58/0 separate. No local full All was performed; final exact-head full CI is recorded in the PR/final handoff.

Fresh focused commands use -Suite Unit -Filter '*A039*' / '*A040*', and -Suite Integration -Filter '*A039/A040*' / '*real Windows batch launcher plumbing*'. Run full Unit/Integration and Analyze in fresh processes with native module isolation and process-only bundled Python PATH above. Real owned Windows terminal observations establish guided prompts/final Enter and success 0/incomplete 2, while parameterized/missing-script paths do not pause. No graphical double-click or universal physical Ctrl+C claim. The GitHub Windows job deadline is 40 minutes because the concurrent local PS7 integration took 1,968.36 s plus 194.23 s for units, exceeding 36 minutes before setup/analysis; native integration took 1,113.33 s plus 138.73 s for units. Product request/retry limits are unchanged. Required assertions and both engines remain enabled. [Task evidence](evidence/UPD-0301.md) and [launcher observations](evidence/UPD-0301-LAUNCHER.md) distinguish baselines, interim failures, focused passes and manual limitations.

## UPD-0302 focused checks (historical snapshot)

Use fresh processes, native module isolation/process-only Bypass and the bundled Python PATH wrapper above:

```powershell
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A043: *'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A042/A043*writer lock*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A043 graceful cancellation*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A043 *cleanup*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A043 resume prefix guard*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit -Filter '*A017*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Unit
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A042/A043*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite All
pwsh -NoProfile -NonInteractive -File ./scripts/Analyze.ps1
```

UPD-0302 inventory: **1,460** = **1,158 units** (1,148 product + ten runner guards) + **302 integrations**. Full local Unit passed **1158/0/0/0 each**, PS7 216.48 s/native 148.61 s. New real run-safety focus passed **13/0/0/289 each**, PS7 130.39 s/native 74.76 s. Final whole-tree analysis passed **87 files**, parse/new zero, bounded baseline PS7 2/native 1. Broader local Integration was launched; actual final counts and exact-source full All CI are recorded separately in the draft PR/final handoff. Filtered/helper/historical counts are not current full-suite passes.

Tests cover actual concurrent OS writer locks, independent same-root shows, release/stale-content reopening without PID authority, real ACL denial with exact owned restoration, actual-byte cancellation checkpoints/remaining-tail recovery, retained unknown-length partials, known/unknown space observations and preview no-write boundaries. New cleanup/guard units include checkpoint and disposal faults, actual exclusive reopening, primary-error precedence, legacy inspection and resume-prefix cancellation. Earlier broad Unit failures were stale cancelled-partial cleanup expectations and incomplete standalone mock dependencies/metadata; their assertions were strengthened/migrated and final full Unit rerun. No fixture helper changed, so predecessor helper results remain historical. See [UPD-0302 evidence](evidence/UPD-0302.md) for every baseline/snapshot and limitation.

## UPD-0303 progress and temporary keep-awake checks

Use fresh processes and the native/Python process-only wrappers above:

```powershell
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite All -Filter '*A04[45]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A009 rechecks a destination junction inserted*'
pwsh -NoProfile -NonInteractive -File ./scripts/Analyze.ps1
```

Current inventory **1534** = **1224 units** (1214 product/ten runner guards) + **310 integrations**. Fresh settled new focus passed **74/0/0/1460 each**, PS7 79.33 s/native 51.25 s, child/outer 0. Affected junction focus passed **2/0/0/308 each**, PS7 15.78 s/native 9.60 s. Whole-tree analyzers passed **94 files**, parse/new zero, baseline PS7 2/native 1. Earlier broad Unit passed **1221/0/0/0 each**, PS7 187.87 s/native 125.31 s, before the final three cleanup tests/fix; it is not full current-unit coverage. Final exact-head full All CI is recorded separately in the PR/handoff. No local full All or full settled Integration was claimed; filtered-out/helper checks are not passes.

The 66 new units and eight loopback cases cover honest known/unknown response totals, offsets/retries, completion-history ordering, throttling/preferences/owned IDs, quiet/preview, invalid body/history failure, cancellation/resumed final hash/catalogue gaps, lazy native thread-scoped leases and independent cleanup. Actual Windows ConsoleHost/native API observations are separate from mocked/forced presentation. [UPD-0303 evidence](evidence/UPD-0303.md) and [manual observations](evidence/UPD-0303-WINDOWS.md) retain baselines, exact versions, commands and limits. Physical Ctrl+C, sleep/lid behavior and hard-kill restoration remain unclaimed. No fixture helper change or rerun; predecessor helper results remain historical.
