# UPD-0301 launcher and A041 observations

Observed 3 October 2026 (Europe/Berlin), actual checkout `D:\projects\UniversalPodcastDownloader`, M3 branch `codex/upd-m3-cli-results`. Launcher/source changes were uncommitted at these observations; the main UPD-0301 evidence records the settled implementation checkpoint. The preserved `UniversalPodcastDownloader-main` snapshot was untouched.

## Desired regression and automated plumbing

Before editing the batch file, the desired forwarding regression failed on both engines: **0/1/0/271**, Integration inventory **272**, PS7 **16.72 s**, native **16.75 s**. A harmless colocated child requested code 2 through launcher arguments; the original launcher ignored those arguments, returned 0, emitted an unconditional completion message and required input release. This was a product launcher failure, not a helper-only pass.

The settled eight launcher checks passed **8/0/0/264** per test host at the same 272-integration snapshot: PS7 **6.80 s**, native **6.73 s**. Each runs the real batch file copied beside a harmless synthetic script in a marked owned temporary root. Assertions cover literal PowerShell-expression data with no injected file, exact native child codes **0/1/2/130**, parameterized no-pause, the no-argument marker, a visible missing-script error/code 1 and inherited ERRORLEVEL shadowing. The fixture launcher path contains spaces, Unicode, ampersand, parentheses and exclamation. Both Pester hosts exercise a native Windows PowerShell 5.1 launcher child; the product's direct PS7 process boundary is covered separately.

Focused commands, each in a fresh process:

```powershell
pwsh -NoProfile -File ./scripts/Test.ps1 -Suite Integration -Filter '*A041 launcher forwarding*'
pwsh -NoProfile -File ./scripts/Test.ps1 -Suite Integration -Filter '*real Windows batch launcher plumbing*'
pwsh -NoProfile -File ./scripts/Analyze.ps1 -Path ./tests/Integration/Launcher.Tests.ps1
```

The same Suite/Filter/Path commands ran through the documented native module-isolation wrapper with process-only `-ExecutionPolicy Bypass`. Final analysis of this owned test file: **1 file, zero parse errors/new findings/baseline warnings**, both engines. An initial Pester fixture-variable scope warning was corrected by using explicit script scope; the eight tests were rerun afterward. Earlier first-pass eight-case results were PS7 5.85 s/native 5.89 s and remain separate snapshots.

## Actual terminal observations

The actual repository launcher and product entry point were exercised in owned Windows terminal sessions. The observer supplied a loopback fixture feed only after seeing the real feed prompt, pressed Enter after seeing the count prompt, then waited for and observed the retained closing prompt before pressing Enter again. USERPROFILE, LOCALAPPDATA, TEMP/TMP and native module paths were set for the owned subprocess and restored. Owned archive output and server processes were removed after verification.

| Observation | Observed behavior |
| --- | --- |
| Argument-free success, 06:15 UTC | Real feed/count prompts; Enter selected Latest; one MP3 stored with source SHA-256 unchanged; `[OK]`; `Press Enter to close` remained pending until answered; actual launcher code 0. Initial feed and enclosure requested once each. |
| Argument-free missing continuation, 06:16 UTC | Real feed/count prompts; accessible Latest stored with original hash; catalogue/run warnings; no `[OK]`; closing prompt retained until answered; actual launcher code 2. Initial feed, missing continuation and selected enclosure requested once each. |
| Actual parameterized `-NonInteractive -WhatIf` | Actual launcher code 0 without supplying input; no guided/closing prompt or archive/log/state/lock/checkpoint/media output. Metadata was allowed. |
| Copied launcher with its script absent | Fixed visible `ERROR: UniversalPodcastDownloader.ps1 is missing next to the launcher.`; code 1; no completion message or pause. |

The terminal observer was PS7 **7.6.5**; the real launcher used Windows PowerShell **5.1.26100.9444** on Windows **10.0.26300.0**. Pester **5.7.1**, PSScriptAnalyzer **1.24.0**; loopback tooling used bundled Python **3.12.14**. Success/incomplete terminal probes invoked `.dev-tools/upd0301-launcher-pty.ps1 -Case success` / `-Case incomplete`, then supplied the observed loopback URL, count Enter and closing Enter through their owned terminal input. Parameterized/missing-script observations used `.dev-tools/upd0301-launcher-observe.ps1 -PausePrompt 'Press Enter to close' -SkipGuided`.

Ignored logs/records retain `upd0301-launcher-baseline-*`, `launcher-final-*`, `launcher-settled-*`, `launcher-analysis-*`, `launcher-analysis-settled-*`, `launcher-manual-success.json`, `launcher-manual-incomplete.json`, `launcher-observations.json` and `launcher-observe-parameterized.log`, all under `.dev-tools` with the `upd0301-` prefix. An initial redirected-pipe observer timed out while locating the interactive Read-Host prompt; its owned children were cleaned and terminal observation replaced that instrumentation. It is not counted as a failed product assertion or a passing manual observation.

The first guided observation also generated a native engine module-analysis cache at `Microsoft/Windows/PowerShell/ModuleAnalysisCache` in the checkout. Its `PSMODULECACHE` header, native module paths and 06:15:27 UTC creation/write time identify the 8,246-byte file as an owned probe artifact. The exact file was retained as ignored `.dev-tools/upd0301-owned-native-ModuleAnalysisCache`; only its now-empty original directories were removed. This background engine-cache write is distinct from the product's archive and preview output. [Native module-analysis cache](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_windows_powershell_5.1?view=powershell-5.1)

Explorer double-click visual appearance, accessibility/layout and physical keyboard Ctrl+C were not tested. Terminal observations establish the guided input/closing behavior through the actual Windows launcher; fixture code 130 establishes forwarding only. Catchable product cancellation and full matrix results belong to the main task evidence. No real archive/private feed, persistent setting, machine-wide policy or unrelated process was changed.
