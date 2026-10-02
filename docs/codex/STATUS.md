# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0204 only**.

- UPD-0001, UPD-0002, UPD-0101 through UPD-0107 and UPD-0201 through UPD-0204: done; historical evidence preserved.
- UPD-0204: **done**; culture-independent publication dates and stable ordering implemented and verified on both Windows engines.
- Next, ready: **UPD-0205 — select and preserve supported audio formats**; unstarted.
- Actual checkout `D:\projects\UniversalPodcastDownloader`; preserved original `UniversalPodcastDownloader-main` snapshot untouched.
- Branch `codex/upd-m2-network-feeds`; [M2 draft PR #3](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/3) remains stacked on `codex/upd-m1-safety`. PRs #1/#2 remain open, draft and unmerged.

## Verified behavior

Starting local/fetched remote/PR HEAD `88fba918eedece4148680951b7dcc333f77df10c` matched, clean with 0/0 divergence. [Predecessor exact-head CI 37034063516](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37034063516) passed both jobs. Implementation `b5fe1b194c66fb5d869134afaccc0dac95636c09` contains 21 explicitly named source/test/policy files; final completion records use a separate documentation checkpoint. Its pushed SHA, remote match and exact-head CI are recorded in the PR/final handoff, avoiding self-referential commits.

Publication parsing uses bounded numeric Gregorian components and explicit offsets to return UTC DateTimeOffset/null independent of current culture. RSS uses pubDate; Atom published precedes updated fallback, with actual date namespaces checked. Malformed present published stays undated. Dates compare by UTC ticks; equal/missing dates retain explicit source positions, dated entries first. Latest/Custom/All keep array shapes and source records unchanged.

New filename dates and duplicate metadata comparisons use UTC. Existing stable identities keep recorded relative paths and original media. The established publisher-ID lookup is deliberately compatible so date work cannot rebind history. Separate old-priority/local-day LegacyPubDate is only an unverified filename review hint. Missing, invalid and unsupported date text preserves enclosure/identity data. [PUBLICATION_DATES.md](PUBLICATION_DATES.md) records grammar, precision, UTC/no-zone policy and historical hint limits.

Shared discovery, bounded transport, validator-aware owned resume, preview, safe targets, XML/metadata limits, no-overwrite placement, writer locks, history schemas 1/2, prepared completion evidence, legacy workflows and private diagnostics remain protected. No runtime package, flag, archive schema, media-choice, pagination, launcher or release behavior was added.

## Verification and limits

Settled inventory: **1,093 checks**: 870 units (860 product units including one remaining media-selection characterization plus ten runner guards) and 223 integrations. Local final Unit passed **870/0/0/0 per engine**, PS7 179.23 seconds/native 132.30 seconds. Final A033/A034 units passed **151/0/0/719 each**; final date boundary integrations passed **7/0/0/216 each**, PS7 86.15 seconds/native 46.99 seconds. Affected history/legacy integrations passed **42/0/0/181 each**, PS7 393.32 seconds/native 234.04 seconds. Counts use passed/failed/skipped/not_run; unselected cases are not passes. No local full All run was performed; final complete exact-head CI is recorded separately in the PR/final handoff.

Final analyzers passed **66 files, zero parse errors/new findings**, reduced **2 PS7 / 1 native baseline warnings**. Fixture helpers passed **46/0**, separate from product acceptance. Local Windows 10.0.26300.0, PowerShell 7.6.5 / 5.1.26100.9444, pinned Pester 5.7.1/PSScriptAnalyzer 1.24.0 and bundled Python 3.12.14; process-only PATH/native module handling remains in DEVELOPMENT.md.

Meaningful pre-fix desired probes failed **0/8/0/719 on both engines**, reproducing ambiguous culture dates, DateTime rather than UTC DTO, filename cast failure and updated-before-published ordering. The first implementation focus passed 148 cases before compatibility review added three cases. Initial two/one new analyzer findings were corrected. [UPD-0204 evidence](evidence/UPD-0204.md) preserves each snapshot and commands; earlier task evidence is unchanged.

A001-A034 passed; A035-A060 remain not_run. The remaining media-selection characterization does not satisfy A035. Static date parsing is a documented subset; legacy hints cannot reproduce every prior machine's time zone or unsupported date text. Tests use synthetic/mock/loopback sources and owned outputs only. No real archive/private feed, live publisher, UNC/hardware power-loss guarantee, merge, history rewrite, persistent setting change or publication was tested/performed. Stop after UPD-0204; NEXT_THREAD_PROMPT.md describes UPD-0205 on the same M2 branch/draft PR.
