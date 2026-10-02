# Retry and timeout policy

UPD-0201 uses one built-in .NET transport policy for pages, feeds and media on Windows PowerShell 5.1 and PowerShell 7. No runtime package is required. The URL, redirect, TLS, credential and size rules in [INPUT_BOUNDARIES.md](INPUT_BOUNDARIES.md) still apply.

## Configuration

These script parameters also name the arguments to `New-PodcastTransportPolicy`:

| Parameter | Default | Accepted range | Meaning |
|---|---:|---:|---|
| `MaxAttempts` | 3 | 1–10 | Initial attempt plus retries per metadata request or episode transfer |
| `HeaderTimeoutSeconds` | 30 | 0.001–86,400 | Time allowed for a request to obtain headers, including connection setup, on each redirect hop |
| `IdleTimeoutSeconds` | 30 | 0.001–86,400 | Time allowed for each pending body read to produce bytes or EOF |
| `RetryBudgetSeconds` | 120 | 0–86,400 | Elapsed window in which another attempt or a delayed redirect may start |
| `BaseDelaySeconds` | 1 | 0–3,600 | Initial exponential backoff delay |
| `MaxDelaySeconds` | 30 | 0–3,600 | Maximum locally chosen backoff; does not shorten a server delay |

For example, `-MaxAttempts 2 -HeaderTimeoutSeconds 20 -IdleTimeoutSeconds 60 -RetryBudgetSeconds 90` changes the policy for both metadata and media. The normal defaults need no additional options. `RetryBudgetSeconds 0` allows the initial attempt and prevents retries; a redirect requiring a future server delay is deferred.

## Time boundaries

`HttpClient.Timeout` is disabled in favor of explicit bounded header waits. Each GET uses `ResponseHeadersRead` and a cancellation token. Body reads use `ReadAsync`, a bounded task wait and cancellation plus stream disposal on timeout. The explicit wait also covers .NET Framework streams that do not honor cancellation after a read starts. Final cleanup disposes streams, responses, requests and clients.

The retry window begins before the initial attempt and is shared by retries and redirect delays. It controls starting more work after failure or a server-requested wait; it is not a total download deadline. A transfer that continues producing bytes within the idle limit may outlast both the header timeout and retry window. A later failure after the window ends is reported without another attempt. No arbitrary short total-duration limit kills an active episode. Local filesystem calls and operating-system cleanup are not covered by network deadlines.

## Failure and delay decisions

HTTP 408, 429, 500, 502, 503 and 504 are transient. Supported connection, DNS, socket and premature-end failures, header timeouts, idle timeouts and short framed bodies may retry. Other HTTP statuses, invalid URLs or redirects, certificate/TLS failures, invalid metadata/media, unsupported encoding, destination errors and history/finalization errors stop. Unknown failures are conservative permanent failures.

Backoff doubles from the base delay, up to the local maximum. A valid `Retry-After` integer or HTTP date provides an additional minimum wait, including on redirects. The later of local backoff and server time wins. Large numeric server delays are deferred instead of overflowing or falling back to a short wait. Invalid unrecognized values use local backoff. Waits round upward and recheck the clock; a short wakeup cannot cause an early retry. If the required wait exceeds the remaining budget, or the budget expires while waiting, the request fails with a safe deferred category. It does not schedule background work.

Each metadata retry starts a fresh in-memory body. Each episode retry repeats the transaction with a new exclusively owned temporary file. Failed attempts clean only their own partial; unknown partials and final media remain protected. Only a fully received, validated and recorded transfer is complete. Resume and ranges remain UPD-0202.

## Diagnostics and tests

Transport failures carry fixed categories and attempt counts in memory. User-facing diagnostics map categories to fixed messages and omit raw exceptions, URLs, headers and server dates. Deferred episodes count as unsuccessful in the existing summary; the broader result/exit-code contract remains UPD-0301.

Unit tests inject clock and delay functions. Loopback tests measure actual request counts/times and exercise delayed headers, stalled bodies, continuing slow bodies and recovery after truncation. All archives and logs are synthetic and temporary. [Task evidence](evidence/UPD-0201.md) records exact engine results separately from this policy.
