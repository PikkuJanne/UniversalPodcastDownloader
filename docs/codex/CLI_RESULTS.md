# Implemented CLI, result and launcher policy

UPD-0301, 3 October 2026. The Windows script and batch entry points remain portable and use no runtime package.

## Arguments and interactive input

`-CustomCount N` alone selects Custom mode. N must be a positive integer. Explicit `-Mode Custom` requires a count; combining a count with explicit Latest or All fails before metadata retrieval or archive planning. Latest selects one and Custom selects N from the bounded accessible catalogue after UTC publication ordering. Pagination gaps remain incomplete in every mode; reaching N does not establish complete coverage.

`-NonInteractive` requires an explicit nonblank `-FeedUrl`. It suppresses feed/count questions and confirmation prompts; combining it with an explicitly enabled `-Confirm` fails. With no explicit mode or count, noninteractive input selects Latest. Multiple HTML-discovered feeds still require a direct feed URL and cannot trigger an interactive numbered choice. Normal guided usage remains available when noninteractive mode is omitted.

```powershell
pwsh -NoProfile -File .\UniversalPodcastDownloader.ps1 -FeedUrl 'https://example.invalid/feed.xml' -CustomCount 5 -NonInteractive
pwsh -NoProfile -File .\UniversalPodcastDownloader.ps1 -FeedUrl 'https://example.invalid/feed.xml' -NonInteractive -WhatIf
```

These are syntax examples using a reserved test domain. `-WhatIf` may retrieve bounded metadata and read local evidence, but never requests planned enclosure media or writes archive/log/export/lock/history/checkpoint files. An incomplete preview retains `Preview=true` and an incomplete result; it does not claim downloaded media.

`-KeepAwake` optionally requests temporary Windows system sleep prevention after work is confirmed. It is off by default, uses a dedicated native thread, and restores that thread's previous execution-state flags during catchable cleanup. Preview never activates it. Unsupported platforms and native failures produce a fixed advisory and permit ordinary download checks. Explicit Sleep and lid actions remain possible; forced termination cannot guarantee restoration.

Animated progress is limited to an interactive ConsoleHost with unredirected output, `-NonInteractive` absent and the caller's progress preference set to Continue. Validated response totals drive percentages; unknown totals show received bytes. An episode reaches 100% only after verified media and completion history, and the run reaches 100% only when all selected work is verified without catalogue gaps. See [progress and power policy](PROGRESS_AND_POWER.md).

## Callable API and script boundary

Dot-sourcing the root script loads import-safe helpers, including `Invoke-PodcastRun`. Calling that function returns a `Podcast.RunResult` object and never exits its caller's host. The executable script emits the result only with `-PassThru`, then exits with its `ExitCode`. Native `-PassThru` output is PowerShell's textual formatting alongside console messages; it is not a JSON protocol. In-process callers can use the returned object directly.

```powershell
. .\UniversalPodcastDownloader.ps1
$result = Invoke-PodcastRun -FeedUrl 'https://example.invalid/feed.xml' -NonInteractive -WhatIf
$result.ExitCode
```

| Code | Status | Meaning |
| --- | --- | --- |
| 0 | success or preview | The requested operation completed within its documented scope. A preview performs planning only. Exhausting supported feed pagination does not prove a complete historical catalogue. |
| 1 | fatal | Input, setup, history or another run-level operation failed. Retained results and deferred entries describe work already observed. |
| 2 | incomplete | Feed pages remain unresolved, or selected media is failed, deferred, conflicting or unverified. Accessible completed work remains recorded. |
| 130 | cancelled | The application caught cancellation and could produce a result. Recovery evidence and completed media remain governed by the existing transaction policy. |

Cancellation takes precedence over fatal and incomplete results; a run-level fatal error takes precedence over incompleteness. Host termination, forced process kill and every operating-system Ctrl+C path cannot guarantee a catchable result or code 130. Microsoft's native `-File` documentation describes host-specific Ctrl+C behavior; tests and evidence identify the catchable paths actually exercised. [PowerShell executable boundary](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_powershell_exe?view=powershell-5.1)

The run result contains `Type`, `SchemaVersion`, `Mode`, `Preview`, `Complete`, `ExitCode`, `Status`, `Planned`, outcome counts, `CatalogueComplete`, `CatalogueStopReason`, `PagesFetched`, `Episodes`, `Message` and `Plan`. Episodes and Plan remain arrays for zero, one and many entries. Per-episode results contain opaque identity, outcome, available byte/verification/attempt observations and a fixed private message. Attempt counts report available observations; zero does not assert that no network work occurred.

Ordinary result projections omit raw request URLs, titles, headers, publication text, validators, mutable history, checkpoints and filesystem destinations. Local legacy review/action results can include `LegacyResult`, including when ordinary `-WhatIf` encounters an archive requiring review. These objects contain sensitive local inventory and should not be treated as shareable diagnostic exports. Established media/history destinations remain authoritative. A successful explicit adoption or rollback can return 0 for the completed metadata action without counting the adopted media as downloaded. A later ordinary run reports that adoption as `legacy_unverified` and returns 2; it does not turn the record into a verified skip. An unresolved catalogue still makes an explicit action incomplete.

`[OK]` and `Run completed` appear only for successful ordinary execution. Incomplete, fatal, cancelled and preview results retain their own messages without a completed-download banner. A catalogue warning remains visible when episode work also fails.

## Windows batch launcher

`UniversalPodcastDownloader.bat` uses the colocated script and the built-in Windows PowerShell 5.1 executable. It invokes a quoted absolute path with `-File`, passing the original `%*` argument text once. It uses no `-Command`, `CALL`, evaluation or argument-content condition. Delayed expansion is disabled locally. A missing colocated script produces a fixed visible error and code 1. The child's exact process code survives local cleanup, including an inherited environment variable named ERRORLEVEL.

The launcher supplies a process-local `UPD_LAUNCHER=1` marker. Only the script boundary uses it: a launch with no original bound parameters retains the guided closing `Press Enter to close` prompt after the result, then returns the same process code. Every parameterized launcher invocation omits that closing pause. Use `-NonInteractive` or an explicit mode/count when avoiding the ordinary input questions too. The callable API never performs the launcher closing pause.

CMD applies its own quoting and expansion before the batch file receives arguments. Quote paths and values containing spaces or metacharacters; the launcher does not undo caller-side environment expansion or malformed quoting. PowerShell expressions passed as properly quoted argument data remain literal script parameters. `-ExecutionPolicy Bypass` applies only to the child process and does not change the machine's policy. [CMD quoting and expansion](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/cmd), [setlocal](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/setlocal), [ERRORLEVEL semantics](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/if), [exit /b](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/exit), [PowerShell -File](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_powershell_exe?view=powershell-5.1)

Actual Windows terminal observations and launcher fixture checks are recorded in [UPD-0301 launcher evidence](evidence/UPD-0301-LAUNCHER.md). Main task evidence records the full engine matrix and final source checkpoint.
