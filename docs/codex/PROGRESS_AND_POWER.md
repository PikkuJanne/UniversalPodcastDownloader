# Progress and temporary keep-awake

Implemented in UPD-0303; Windows PowerShell 5.1 and PowerShell 7 remain supported. Imports and previews neither render episode progress nor initialize native power helpers. There is no runtime package or new mandatory TUI question.

## Progress

An ordinary unredirected ConsoleHost displays an episode parent and media child. The parent identifies the current episode and shows processed and verified counts separately. Receiving bytes use only the validated response's full length and actual bytes, including an already verified resume prefix. Each new response resets its offset and total; a fresh retry starts at zero. Feed enclosure lengths remain advisory and never establish the displayed HTTP total.

Known-length receiving and validation percentages stop at 99, including when every advertised byte has arrived. Unknown-length bodies show actual received bytes and an indeterminate percentage until verified success. The verifying phase includes media checks, placement and history work. An episode reaches 100 only after the recorded transfer returns with committed `transfer_verified` history. A whole run reaches 100 only when its result is complete and every selected episode is downloaded or a verified history skip. Conflicts, owner-adopted unverified files, failure/deferred/cancelled episodes and catalogue gaps cannot produce whole-run 100.

Byte repainting is bounded to four times per second, with response, phase and outcome changes rendered immediately. Progress IDs belong to the invocation; a previous child is cleared before the next episode and finally attempts every shown child/parent independently. Cleanup clears only owned IDs under a function-local preference and leaves the caller's preference unchanged. Typed rendering cancellation propagates through the existing recovery path. A new typed cancellation while closing progress changes an otherwise successful returned result to 130 after cleanup; an earlier failure remains primary, and already committed episode evidence is retained. Ordinary display failures disable rendering without changing media results.

`-NonInteractive`, redirected stdout/stderr, non-console hosts and caller preferences other than `Continue` suppress animated progress. Existing fixed episode notices and the downloaded/skipped/adopted/failed summary remain readable. The imported API still returns exactly one schema-1 result. Progress carries fixed phase text and numeric counts, never untrusted titles, URLs, paths or raw exceptions. The host may choose its own layout, colours and update frequency; accessible text and redirected output do not depend on a particular terminal theme.

## Opt-in temporary keep-awake

```powershell
pwsh -NoProfile -File .\UniversalPodcastDownloader.ps1 -FeedUrl 'https://example.invalid/feed.xml' -Mode All -NonInteractive -KeepAwake
```

`-KeepAwake` is disabled by default and Windows-only. It starts after `ShouldProcess` accepts ordinary archive execution or an explicit legacy action, so planning, prompts, WhatIf, legacy Preview and declined confirmation do not request native wakefulness. Explicit legacy redownloads receive the same progress and confirmed lifetime. The helper requests system wakefulness, without a display or away-mode request and without changing the power plan.

The lazily compiled built-in helper owns a dedicated managed worker, with native thread affinity around both Win32 calls. It calls `SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)`, retains that thread's exact returned previous flags, and restores them on the same native thread before joining. It never runs a PowerShell callback on the worker. A separate caller thread's existing request remains untouched; concurrent leases have independent ownership. Acquisition and cleanup are bounded. Unavailable/failed power helpers report fixed advisory text and retain the download result; native restoration is not claimed when it cannot be confirmed. Progress, power and archive-lock cleanup are attempted independently during normal completion, exception and catchable cancellation.

Windows may honour explicit sleep, a lid action or platform/policy restrictions despite the request. Physical sleep/lid behaviour, hard termination and power loss were not measured. Finally cannot be promised after an uncatchable hard kill. No power-policy command, display setting, administrator requirement, arbitrary PID lookup/kill or persistent setting change is added. See [Microsoft SetThreadExecutionState](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate) and [Thread.BeginThreadAffinity](https://learn.microsoft.com/en-us/dotnet/api/system.threading.thread.beginthreadaffinity).

## Evidence boundaries

[UPD-0303 evidence](evidence/UPD-0303.md) distinguishes mocked units, real owned loopback processes, actual native API observations and Windows ConsoleHost observations. The console checks used real unredirected Windows pseudo-terminals and the native renderer with the detector unmodified. Separate intercepted render records verify byte/history ordering, without claiming a pixel screenshot or every physical terminal's accessibility. Cancellation was injected as a typed exception after actual bytes; physical keyboard Ctrl+C remains unclaimed. Test artifacts and cleanup belong only to marked temporary roots and tracked fixture children.
