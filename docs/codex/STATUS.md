# Current implementation status

Updated 2026-10-03 (Europe/Berlin). Scope: **UPD-0303 complete**.

- UPD-0001/0002, UPD-0101 through UPD-0107, UPD-0201 through UPD-0206 and UPD-0301 through UPD-0303: done; historical evidence preserved.
- A001-A045 passed; A046-A060 remain not_run. UPD-0304 is ready and unstarted.
- Actual checkout `D:\projects\UniversalPodcastDownloader`; original `UniversalPodcastDownloader-main` snapshot untouched.
- Branch `codex/upd-m3-cli-results`; [M3 draft PR #4](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/4) stacked on `codex/upd-m2-network-feeds`. Earlier milestone PRs remain open, draft and unmerged.

Startup clean local/fetched origin/independent remote/PR head matched `535054961988d763f141b9fc9001006f824091cd`, 0/0 divergence. [Predecessor final-source CI 37108878506](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37108878506) passed All 1460/0/0/0 each with checkout-tree equality. Named UPD-0303 source/policy/test checkpoint `19c299f5075a17eb258e1347ca520a9a9760caed` contains 19 files; completion evidence is separate. Final pushed SHA, clean remote equality and exact-source full All CI are recorded independently in the draft PR/final handoff.

Interactive ConsoleHost now shows honest episode and streamed-byte progress from validated response totals/actual resume offsets. Receiving/verifying stays below 100; unknown totals show actual bytes. Episode 100 follows verified completion history; run 100 requires complete verified work without catalogue gaps. Each response/retry resets its state. Owned progress IDs, bounded repainting and independent cleanup preserve caller preferences. NonInteractive/redirected/non-console output remains quiet with fixed notices and readable summaries. Progress contains fixed phase/count text rather than private metadata.

Optional KeepAwake is off by default, Windows-only and activates after confirmation. A dedicated native-affinity worker acquires temporary system execution state and restores its exact previous flags on the same native thread. It requests no display/away mode and changes no power policy. Unsupported/native failures are fixed advisories; failed restoration is not claimed successful. Cleanup attempts progress, power and each lock independently. A new typed progress-cleanup cancellation changes an otherwise successful result to 130 while retaining its completed media/counts; earlier failures remain primary.

Previous CLI 0/1/2/130/private schema-1 results, literal batch forwarding, writer protection, owned write/space preflight and cancellation checkpoints remain. Original bytes, no-overwrite finals, contained/reparse-checked destinations, history schemas 1/2, prepared/completed ordering, strict owned resume, bounded metadata/catalogue gaps and privacy boundaries remain intact. No runtime dependency or release work was added.

Settled inventory: **1534** = **1224 units** (1214 product/ten runner guards) + **310 integrations**. Fresh final new focus **74/0/0/1460 each**, PS7 **79.33 s**/native **51.25 s**, child/outer 0. Affected junction integrations **2/0/0/308 each**, PS7 15.78 s/native 9.60 s. Final analysis **94 files**, parse/new 0, baseline 2/1. Earlier broad Unit **1221/0/0/0 each** preceded the last three cleanup tests and fix; it does not cover the settled inventory. No local full All or full settled Integration pass is inferred. Exact final-head full CI belongs in the PR/handoff.

Local Windows 10.0.26300.0, PowerShell 7.6.5/native 5.1.26100.9444, Pester 5.7.1/PSScriptAnalyzer 1.24.0 and bundled Python 3.12.14. [Task evidence](evidence/UPD-0303.md) records commands, failed baselines and snapshot limits. [Windows observations](evidence/UPD-0303-WINDOWS.md) passed five actual unforced console cases per engine plus native API restoration/caller independence. Catchable cancellation was typed after actual bytes; physical keyboard Ctrl+C, physical sleep/lid behavior, hard kill and power loss remain unclaimed.

Only synthetic/mocked/marked owned loopback outputs and tracked fixture children were used. No real archive/private subscriptions, arbitrary PID kill, persistent settings change, merge/history rewrite or publication. Stop after UPD-0303; NEXT_THREAD_PROMPT.md describes UPD-0304.
