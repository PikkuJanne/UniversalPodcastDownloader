# UPD-0303 — actual Windows console and native power observations

3 October 2026 (Europe/Berlin). Actual checkout `D:\projects\UniversalPodcastDownloader`; marked synthetic loopback outputs only. Windows 10.0.26300.0, PowerShell 7.6.5 and native Windows PowerShell 5.1.26100.9444. No real archive/private feed, machine-wide power policy or display setting was changed.

## Actual console observations

The owned observation script was run in actual Windows pseudo-terminals using `pwsh -NoProfile -File .dev-tools/upd0303-manual-tty.ps1 -Engine PS7` and the same command with `-Engine PS51`. It starts an owned ephemeral loopback server and five fresh matching-engine child processes, invoking the tracked `tests/support/Invoke-ProgressAndPowerWorker.ps1`. `NativeHostObservation=true` uses the product's actual detector; no `ForceInteractive` override is used. A wrapper records and forwards the real native `Write-Progress` cmdlet. Transcript and JSON reports are retained locally under ignored `.dev-tools/upd0303-manual-tty-{ps7,ps51}.{log,json}`.

Both engines reported ConsoleHost, stdout and stderr unredirected, actual interactive detection true and forced detection false. The four ordinary cases enabled the private progress context. The NonInteractive case disabled it despite that same interactive host. Actual renderer output showed preparing/receiving/verifying and episode/byte information; the final fixed summary remained readable. These are native text/renderer observations, not screenshots or a claim about every terminal theme.

| Synthetic case | Both engine exits | Observed progress | Native keep-awake |
| --- | --- | --- | --- |
| Known-length `/transport/feed/progress` | 0 | Actual validated total and received bytes; receiving/verifying below 100; successful 100 only with accepted completion history | Requested, active, restored |
| Unknown-length `/resume/feed/no-length` | 0 | Actual numeric received bytes, total unknown and percent -1 while receiving; verified completion follows history | Off by default; no request |
| Invalid HTML media `/feeds/format-html.xml` | 2 | Failure and no non-completed 100 | Requested, active, restored |
| Catchable cancellation `/resume/feed/valid` | 130 | Typed cancellation after 8,192 actual bytes; no non-completed 100; retained owned partial | Requested, active, restored |
| NonInteractive `/feeds/single.xml` | 0 | Zero rendered records; fixed notices and summary | Off by default; no request |

Both observation scripts and outer tool processes exited 0, recording `observed=5 failed=0 forced_detector=0`. Every requested lease restored its exact previous state `0x80000000` on the same native thread that activated it; the native worker ended inactive with restoration confirmed. Thread IDs and output paths stay in the private ignored reports rather than product presentation.

## Separate actual native API observations

Fresh-process `upd0303-keepawake-smoke.ps1` observations in both engines proved importing/default-off use leaves the native type uncompiled. Success, thrown failure and caught typed cancellation each acquired and restored a real `SetThreadExecutionState` lease on its owning native thread. A separate caller thread held `0x80000001` before, during and after another lease's cleanup, then restored its own original flags in its own finally. Both child/outer exits were 0; ignored logs are `upd0303-keepawake-smoke-{ps7,ps51}.log`.

The helper requests `ES_CONTINUOUS | ES_SYSTEM_REQUIRED`, without display/away mode, OS policy commands, administrator access or a persistent setting. The recorded API behavior accords with [Microsoft's thread-scoped execution-state contract](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate) and the worker's [native thread affinity](https://learn.microsoft.com/en-us/dotnet/api/system.threading.thread.beginthreadaffinity).

## Boundaries

Cancellation was an actual typed exception after streamed bytes, rather than physical keyboard Ctrl+C. Physical automatic sleep prevention, explicit Sleep/lid behavior, hard termination and power loss were not measured. Finally cannot be promised after an uncatchable hard kill. Redirected output, completion-history ordering, resumed final hashes, fresh retry offsets and independent cleanup faults are separately covered by [automated task evidence](UPD-0303.md). Only owned marked outputs and tracked fixture children were removed.
