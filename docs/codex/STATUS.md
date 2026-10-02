# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0203 only**.

- UPD-0001, UPD-0002, UPD-0101 through UPD-0107, UPD-0201 and UPD-0202: done; historical evidence preserved.
- UPD-0203: **done**; shared discovery/resolution implemented, pushed and verified; settled units and full exact-source CI passed both Windows engines.
- Next, ready: **UPD-0204 — normalize publication dates and stable ordering**; unstarted.
- Actual checkout `D:\projects\UniversalPodcastDownloader`; the original `UniversalPodcastDownloader-main` snapshot remains untouched.
- Branch `codex/upd-m2-network-feeds`; [M2 draft PR #3](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/3) remains stacked on `codex/upd-m1-safety`. PRs #1/#2 remain open, draft and unmerged.

## Verified behavior and preservation

Starting local/fetched remote/PR HEAD `2341bddaabd5c5fa64a64b7dc3450c8b3773b86b` matched with 0/0 divergence; [predecessor CI 37026036037](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37026036037) passed both jobs. Implementation `cddb61d8aba35061996c6e3e068775a0968ec4c7` was pushed normally and independently matched remote HEAD. Command-scoped authenticated HTTPS and public author identity were used without persistent settings changes.

Both guided input and CLI use one bounded source/feed resolver. Direct feeds are classified by actual RSS/Atom roots and fetched content is reused. Static page links support final-response/base URI resolution, entity decoding, attribute variants and deduplicated candidates. One candidate is automatic; several require numbered guided selection or a direct CLI URL. Empty/malformed/unsupported/no-link responses have distinct fixed diagnostics. Direct feed identity remains the original URL across redirects; page archive identity uses the selected feed URL.

No new runtime package, flag, nested retry, archive schema, date/media-selection/pagination or release behavior was added. Preview, safe targets, XML/metadata limits, original media, no-overwrite paths, writer locks, resume checkpoints, history/legacy evidence and private diagnostics remain protected. The scanner has explicit limits and is not a complete browser HTML implementation.

## Verification and limits

Settled inventory: **936 checks** (710 product units, including two parser characterizations, plus ten runner guards and 216 integrations). Final local Unit passed **720/0/0/0 per engine**, PS7 166.78 seconds/native 117.02 seconds. Focused A006/A007/A030 integrations passed **43/0/0/173 each**; corrected diagnostics passed **10/0/0/206 each**. All results use passed/failed/skipped/not_run order.

Broad local All snapshots discovered the preceding **935-check inventory** and completed **928/7/0/0 on each engine**, PS7 1,525.31 seconds/native 900.53 seconds. The seven failures were four dynamically scoped unit mock fixtures and three stale malformed-source diagnostic expectations. Corrections were made while later integration workers continued, and the final prefixed-Atom unit was added after discovery. These snapshots are retained as failed runs. Final settled Unit (720/0/0/0 each) and affected diagnostic focus (10/0/0/206 each) passed after repair; full post-fix results come from the clean-checkout CI below.

Final analyzers passed **63 files, zero parse errors/new findings**, reduced **3 PS7 / 2 native baseline warnings**. Fixture helpers passed **44/0**, separate from product acceptance. Local Windows 10.0.26300.0, PowerShell 7.6.5 / 5.1.26100.9444, Pester 5.7.1, PSScriptAnalyzer 1.24.0 and bundled Python 3.12.14; use process-only PATH/native module handling from DEVELOPMENT.md.

[CI 37031909323](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37031909323) passed **936/0/0/0 on each Windows engine** at implementation `cddb61d8aba35061996c6e3e068775a0968ec4c7`. PowerShell 7.6.6 took 421.79 seconds; Windows PowerShell 5.1.20348.5622 took 597.16 seconds. Both analyzers passed 63 files, zero parse errors/new findings, 3/2 baseline warnings. Pester 5.7.1, PSScriptAnalyzer 1.24.0, Python 3.14.7, Windows Server 2022 (10.0.20348.0), runner image 20260927.320.1. Complete logs: `.dev-tools/upd0203-ci-code-{ps7,ps51}.log`.

A001-A032 passed; A033-A060 remain not_run. No real archive/private feed, live publisher, UNC/hardware power-loss guarantee, merge, history rewrite, persistent setting change or publication was tested/performed. [UPD-0203 evidence](evidence/UPD-0203.md) preserves intermediate failures; [discovery policy](FEED_DISCOVERY.md), [development commands](DEVELOPMENT.md), [resume](RESUME_POLICY.md) and [transport](TRANSPORT_POLICY.md) record exact contracts.

The final documentation checkpoint updates these records and the next task without changing runtime/tests. Its SHA and exact-head CI are verified separately in the PR/final handoff, avoiding a self-referential commit. Stop after UPD-0203; NEXT_THREAD_PROMPT.md describes UPD-0204 on the same M2 branch/draft PR.
