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

The historical UPD-0201 suite contained **705 checks per engine**. Consult [UPD-0201 evidence](evidence/UPD-0201.md) for that snapshot. UPD-0401 had **1,709**, UPD-0402 had **1,717**, UPD-0403 had **1,748**, and UPD-0404 adds43 workflow/candidate-tooling checks for current **1,791**. Adding a test does not make an earlier run cover it. Current counts and source coverage are in the UPD-0404 section below. The following table is historical UPD-0201 coverage.

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

UPD-0104 identity units use `-Suite Unit -Filter '*A014*'` (25 checks); state units use `-Suite Unit -Filter '*A015*'` (60 checks). History integrations use `-Suite Integration -Filter '*A01[56]*'` (18 checks). Real child processes exercise both lock scopes and crashes before/after state replacement and final placement. Prepared evidence, changed/deleted media, signed URL refreshes and corrupt history are checked against actual disk and loopback requests. Full `-Suite All` uses the current inventory in the UPD-0402 section; historical counts belong to their recorded task snapshots.

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

These earlier checks do not validate the whole application, launcher UX, catchable cancellation, private feeds or future acceptance cases. Helper-server self-tests are a separate layer. Historical [UPD-0206 evidence](evidence/UPD-0206.md) preserves its 1,272 checks (1,012 units and 260 integrations), bounded pagination/partial-catalogue results, intermediate failures, exact commands and results. Current [UPD-0401 evidence](evidence/UPD-0401.md) records the full 1,709-check local audit on both engines. [Resume policy](RESUME_POLICY.md) documents the tested recovery scope and conservative limits.

## Static analysis policy

`Analyze.ps1` parses the runtime, runner and test PowerShell files in the selected engine, then runs PSScriptAnalyzer's warning/error rules. [tools/PSScriptAnalyzerSettings.psd1](../../tools/PSScriptAnalyzerSettings.psd1) excludes `PSAvoidUsingWriteHost` intentionally because the existing TUI and developer summaries use host output. Other default rules remain enabled.

[tools/lint-baseline.json](../../tools/lint-baseline.json) now has **no warning allowances**: UPD-0401 fixed the two remaining causes below. The retained source identifier `2ac82614493be7196c9ebee116f23fec07368b50` records the original reviewed baseline. The runner permits no new finding or parse error. Any future explicitly reviewed allowance must match the exact repository-relative file, rule, message, surrounding source text and maximum occurrence count; moving a warning into unrelated source or increasing its count fails.

| Historical rule | Former baseline count | Fixed source context |
| --- | ---: | --- |
| `PSAvoidOverwritingBuiltInCmdlets` | 1 | Existing `Write-Log` function |
| `PSAvoidUsingEmptyCatchBlock` | 1 | Feed-title extraction |

The historical baseline emitted 2 warnings on PS7 and 1 on native PS5.1, whose analyzer built-in command profile did not emit the `Write-Log` override warning. Earlier tasks removed obsolete naming/size/BOM allowances; UPD-0203 reduced the baseline to 3/2 and UPD-0204 to 2/1. Historical UPD-0205 analysis covered 69 PowerShell files with zero parse errors/new findings on both. Current UPD-0401 analysis covers 108 files with zero parse errors, new findings and baseline warnings on both engines. Its 35 retained scoped runtime suppressions have authored justifications for in-memory operations, caller-confirmed resource ownership, stable helper/collection APIs and lookup-only legacy SHA1.

## CI and verified sources

[test.yml](../../.github/workflows/test.yml) runs Windows PowerShell 5.1 and PowerShell 7 on `windows-2022`, with a 40-minute job limit and two-job maximum. It triggers for pull requests, pushes to `main` and manual dispatch. It uses only `contents: read`, disables checkout credential persistence and has no artifact/release/publication step. Test/analysis counts and runtime inventory appear in the job summary. Workflow configuration alone is not evidence of a completed CI run; consult the task evidence and actual GitHub checks.

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

UPD-0303 inventory **1534** = **1224 units** (1214 product/ten runner guards) + **310 integrations**. Fresh settled new focus passed **74/0/0/1460 each**, PS7 79.33 s/native 51.25 s, child/outer 0. Affected junction focus passed **2/0/0/308 each**, PS7 15.78 s/native 9.60 s. Whole-tree analyzers passed **94 files**, parse/new zero, baseline PS7 2/native 1. Earlier broad Unit passed **1221/0/0/0 each**, PS7 187.87 s/native 125.31 s, before the final three cleanup tests/fix; it is not full current-unit coverage. Final exact-head full All CI is recorded separately in the PR/handoff. No local full All or full settled Integration was claimed; filtered-out/helper checks are not passes.

The 66 new units and eight loopback cases cover honest known/unknown response totals, offsets/retries, completion-history ordering, throttling/preferences/owned IDs, quiet/preview, invalid body/history failure, cancellation/resumed final hash/catalogue gaps, lazy native thread-scoped leases and independent cleanup. Actual Windows ConsoleHost/native API observations are separate from mocked/forced presentation. [UPD-0303 evidence](evidence/UPD-0303.md) and [manual observations](evidence/UPD-0303-WINDOWS.md) retain baselines, exact versions, commands and limits. Physical Ctrl+C, sleep/lid behavior and hard-kill restoration remain unclaimed. No fixture helper change or rerun; predecessor helper results remain historical.

## UPD-0304 saved-show and sequential-batch checks

Use fresh processes and the same native/Python process-only wrappers:

```powershell
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite All -Filter '*A04[67]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Test.ps1 -Suite Integration -Filter '*A0[34][019]*'
pwsh -NoProfile -NonInteractive -File ./scripts/Analyze.ps1
```

Historical UPD-0304 inventory **1682 =1359 units (1349 product/ten runner guards)+323 integrations**. Frozen new148 focus passed148/0/0/1534 each, PS7 211.88s/native103.51s; whole analysis102 files parse/new0 baseline2/1; child/outer0. Earlier affected prior entry58 focus passed58/0/0/263 of321 each222.05/120.88s, before two added entry cases. Earlier full Unit1351 passed each297.29/177.53s before the last eight guards/preference/catch corrections; not full current coverage. Final exact-head full All CI is independently recorded in PR/handoff. No local current fullAll/fullIntegration pass is inferred.

135 new units and13 product integrations cover strict schema/DPAPI/protected ACL/atomic update/no-evaluation/safe export, real named repetition, complete WhatIf tree preservation, sequential failure isolation and counts, subset/conflicts and typed real-byte cancellation with strict checkpoints/closed guards/unstarted shows. Eleven owned worker cases import the actual dispatcher; two use the actual product -File parameter/exit boundary. All tests use explicit marked owned configuration/output paths; no real archive/subscriptions. [UPD-0304 evidence](evidence/UPD-0304.md) preserves baseline/fixture/runtime failures, exact commands/versions and limitations. Pinned development tools and CI40-minute deadlines remain unchanged; fixture helper checks remain historical and separate.

## UPD-0401 architecture and compatibility checks (historical snapshot)

Use fresh processes and the same native/Python process-only wrappers:

The local Python PATH wrapper used for UPD-0401 is below. Run it from the actual checkout in a PowerShell7 process. For native5.1, run the earlier native module-path wrapper inside this PATH try/finally instead of the PS7 calls. Each Test/Analyze invocation starts a fresh child; only this parent process's PATH changes. The path is the observed bundled runtime on this machine, not an application prerequisite.

```powershell
$updFixturePythonDir = 'C:\Users\jtvuo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python'
$updPreviousPath = $env:PATH
$updPs7 = Join-Path $PSHOME 'pwsh.exe'
try {
    $env:PATH = $updFixturePythonDir + ';' + $updPreviousPath
    & $updPs7 -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite All
    $updTestExit = $LASTEXITCODE
    & $updPs7 -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Analyze.ps1
    $updAnalysisExit = $LASTEXITCODE
    if ($updTestExit -ne 0 -or $updAnalysisExit -ne 0) {
        throw "Checks failed: test=$updTestExit; analysis=$updAnalysisExit"
    }
}
finally {
    $env:PATH = $updPreviousPath
}
```

```powershell
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite All
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Unit -Filter '*A048 [dp]*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Unit -Filter '*A048 transport context*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Unit -Filter '*A048 shared retry*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Unit -Filter '*A03[1278]*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Integration -Filter '*A048*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Analyze.ps1
```

UPD-0401 inventory **1709 =1383 units (1373 product/ten runner guards)+326 integrations**. Full local All passed1709/0/0/0 each, PS7 2574.54s/native 1436.05s, child/outer0. Whole analysis108 files reports parse/new/known-baseline0 on both. Final focused diagnostic/run12 passed12/0/0/1371 of1383 in6.20/6.37s; source/context8 passed8/0/0/1371 of1379 in3.48/3.87s; affected150 passed150/0/0/1229 of1379 in12.31/14.47s; retry4 passed4/0/0/1379 of1383 in3.97/3.91s; frozen copied-runtime3 passed3/0/0/323 of326 in24.15/15.83s. Historical discovery inventories preceded four later retry checks; they are not current full coverage. Exact final-head CI is independently recorded in PR/handoff.

24 new units/three integrations cover caller state, explicit/fresh policy, response reuse, diagnostics close/export, retry observations/exceptions, copied runtime/import/real media/repeat/preview and resource reopening. Runtime worker PATH/module lookup excludes development dependencies; parent Python is fixture tooling only. Tree snapshots compare names/lengths/hashes, not timestamps or a future release ZIP. Retained35 scoped runtime suppressions remain justified; colliding Write-Log/empty catch causes are fixed and the baseline is empty. [UPD-0401 evidence](evidence/UPD-0401.md) records every baseline/mixed-wiring/harness correction, exact versions/commands and limitations. No fixture helper changed/reran; prior helper results remain historical. Help/manual examples and package/release gates remain not_run.


## UPD-0402 documentation verification (historical snapshot)

Use the same fresh-process/Python PATH/native module wrappers. No app runtime dependency was added:

```powershell
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Unit
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Integration -Filter '*A049/A050*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Integration -Filter '*A04[018]*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Integration -Filter '*A01[78]*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Analyze.ps1
```

UPD-0402 inventory1717=1383 Unit (1373 product/ten runner guards)+334 Integration. Both engines passed Unit1383/0/0/0, docs8/0/0/326, affected CLI/runtime32/0/0/302 and legacy25/0/0/309; final whole analysis110 files parse/new/known0. These are1448 distinct local assertions;269 integrations are filtered out, not passes. Full All for that snapshot is its final-head CI in M4 draft PR #5/final handoff; UPD-0401 full local1709 is historical. Exact versions/durations/failed harness snapshots are in [UPD-0402 evidence](evidence/UPD-0402.md).

Eight new cases run literal11 README fences/four Get-Help examples, all public help parameters, script0/1/2 and genuine argument-free batch input from fresh runtime-only copies with Unicode/spaces and isolated owned profile/config/output. Native5.1 splits multiline examples into Code/Remarks; reconstruction must match authored AST command text. Actual hashes/history prove original preservation and adoption versus transfer verification. Preview/package/default-config boundaries are asserted. Manual A049/A050 are separately reviewed for meaning/rendering; no physical owner double-click/accessibility/keyboard cancellation check or release ZIP is claimed. No fixture helper selftest changed/reran.

## UPD-0403 local packaging and actual ZIP verification

Native diagnostic formatting can wrap phrases onto multiple lines in longer CI temporary paths. Packaging units retain raw child output but normalize formatting whitespace for phrase assertions. The three dirty-source cases and missing-license case exercise word-wrapped captured diagnostics while retaining exit1 and output-absence guards. First-head CI37133849561's four native assertion failures are historical; corrected local runs, final rebuilt artifacts and exact-source CI are recorded separately in evidence/PR.

Build from a clean committed source; Git is a build-tool requirement. No Git, Python, Pester or analyzer is required by the extracted application. Version and exact36-file allowlist are checked in `tools/release-package.json`. Never copy a private archive/configuration into a build. The builder exports immutable regular Git blobs, generates a37-entry ZIP and identical internal/sidecar manifest, validates every entry and produces SHA256SUMS covering ZIP/sidecar. Each output parent receives a new version/full-commit child directory; occupied candidates fail without overwrite. Source/output reparse/root/ancestor guards and marked staging cleanup remain active.

```powershell
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Build-Release.ps1 -OutputDirectory ./artifacts/reproduce-one
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Build-Release.ps1 -OutputDirectory ./artifacts/reproduce-two
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Unit -Filter '*A051/A052*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Integration -Filter '*A051/A052*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Test.ps1 -Suite Integration -Filter '*A04[89]*'
pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ./scripts/Analyze.ps1
```

Run tests with the process-only Python PATH/native module wrappers above, and repeat with native5.1. Build itself needs no Python. New candidate0.1.0-rc.1 is unpublished; tags/releases were inspected empty before selection. Inventory uses ordinal sorting and exact source commit/tree/URL; ZIP timestamps/attributes are fixed. Actual clean-checkpoint four builds have equal repeat bytes per engine and equal manifests across engines, while cross-engine compression bytes differ. Do not infer publisher authentication or universal byte reproducibility from checksums. [RELEASE](../../RELEASE.md) explains unsigned/process-only launch policy, source verification and application-code rollback versus retained archive/schema state.

Current **1748=1410 Unit (1373 existing product+ten runner guards+27 packaging-tooling guards)+338 Integration**. Both local engines passed packaging Unit27/0/0/1383 and actual ZIP Integration4/0/0/334. Affected docs/runtime11/0/0/323 passed at the earlier334-integration snapshot before new ZIP cases were discovered. These cover42 distinct current assertions per engine,1706 not selected locally. Whole final analyzers114 files parse/new/baseline0/0/0 each. No current full local Unit/Integration/All pass is claimed; final exact-source full All CI must be verified in PR before handoff. Existing helper selftests are historical and unchanged. Exact timings/commands/failed harness snapshots and clean source/checksum identifiers are in [UPD-0403 evidence](evidence/UPD-0403.md).

The new units use tiny owned synthetic Git repositories for binary/source/dirty/inventory/revision/path/junction/collision/tamper guards. Four actual ZIP cases use a clean committed copy of real runtime source and ignored private poison, build twice, validate hashes/import closure/MIT/artwork, and freshly extract into Unicode/spaces paths outside checkout. Isolated app workers lack developer tools, render help, run genuine native batch preview/fatal paths, transfer original loopback bytes, verify repeat/preview preservation and closed handles. Separate canonical clean-checkpoint builds and two actual extraction observations per engine support manual A052 source/license review; automated observations do not change its manual classification.

Native Windows host startup can create standard profile/TEMP dirs and mutate exactly `USERPROFILE/AppData/Local/Microsoft/Windows/PowerShell/StartupProfileData-NonInteractive`. Worker host-only observations and cache before/after sizes/hashes are recorded. The fixture prepares only owned standard parents and separates that exact host cache file; all other directories/files, including downloader output/configuration/logs/TEMP, remain in preservation comparisons. This asserts application-controlled preview preservation rather than a host-wide zero-write promise. No runtime behavior, launcher or helper changed.

## UPD-0404 candidate workflow verification

Current1791=1453 Unit+338 Integration. New43 are19 workflow-policy/mutation checks and24 actual synthetic candidate tooling checks, separate from downloader acceptance. Both engines pass focused Unit43/0/0/1410 and affected docs Integration8/0/0/330. Actual package prerequisites4/0/0/334 passed before draft changes.55 distinct local selections;1736 current assertions not selected locally. No current full local All pass is claimed. Both whole116-file analyzers pass parse/new/baseline0/0/0; verified one-off actionlint1.7.12 reports no findings. Initial ten test-fixture lint warnings were corrected without baseline expansion, then43 checks rerun both engines. Exact commands/versions/durations and earlier runs are in [UPD-0404 evidence](evidence/UPD-0404.md).

Use Test.ps1 -Suite Unit -Filter '*A053*', Integration -Filter '*A049*' or '*A051/A052*', and Analyze.ps1 through the fresh-process/native module/Python wrappers above. On a clean committed checkout, invoke Prepare-ReleaseCandidate.ps1 -SourceCommit <full HEAD> -OutputDirectory <new owned parent>; optional -GitHubOutput <existing owned file> exports exactly3 validated paths. Dot-sourcing loads only development validator functions. No runtime imports or publication API are added.

Clean-checkpoint actual wrapper builds on both engines independently verify Git blob bytes,36 payloads/37 ZIP entries, original MIT, manifests/checksums and output identities. Final-head candidates and final downloaded CI artifact are rebuilt/verified separately and recorded with exact source/run/hash identities in PR/handoff. New CI tests/packages the same immutable event head after both complete engine jobs, unlike the historical default PR merge checkout. It does not certify a future merge result. Require final full All1791/0/0/0 and116-file analysis0/0/0 plus actual candidate upload before handoff; never infer them from focused local runs. Review [workflow trust boundaries](RELEASE_WORKFLOW.md) and [approval checklist](../releases/APPROVAL_CHECKLIST.md) before any separate publication decision.
