# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0205 only**.

- UPD-0001, UPD-0002, UPD-0101 through UPD-0107 and UPD-0201 through UPD-0205: done; historical evidence preserved.
- UPD-0205: **done**; supported audio selection and canonical final extensions verified on both Windows engines.
- Next, ready: **UPD-0206 — follow explicit feed pagination safely**; unstarted.
- Actual checkout `D:\projects\UniversalPodcastDownloader`; preserved original `UniversalPodcastDownloader-main` snapshot untouched.
- Branch `codex/upd-m2-network-feeds`; [M2 draft PR #3](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/3) remains stacked on `codex/upd-m1-safety`. PRs #1/#2 remain open, draft and unmerged.

## Verified behavior

Starting local/fetched remote/PR head `207108b227ef0e9effef38313484ce663e5fc7d0` matched, clean with 0/0 divergence. [Predecessor exact-head CI 37039367636](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37039367636) passed both jobs. Implementation `e8b7e5d8562ed12d85e0413bea400f727aaff788` contains 29 explicitly named runtime/test/policy files; final completion records use a separate documentation checkpoint. Final pushed SHA, independent remote match and exact-head full CI are recorded in the PR/final handoff, avoiding self-referential commits.

`MediaSelection.ps1` retains RSS/Atom candidates, selects the first eligible known audio in feed order and uses generic unknown only as fallback. Supported MIME can identify extensionless audio; URI path suffixes are weak hints. Explicit unsupported types are excluded. A later typed alternative does not reorder an earlier eligible URL, preserving URL-only identity. Publisher-ID/date/discovery contracts remain compatible.

Completed transport framing plus bounded signatures determine canonical MP3/M4A/Ogg/WAVE/FLAC extensions. Misleading feed/header/URL metadata yields fixed private warnings. Text, unknown binary, observed video/unsupported MPEG layers and ambiguous containers fail. Original stored bytes remain unchanged. MP4/Ogg checks provide limited structural audio indications; M4A branding can contain video and is not proof of audio-only media. [AUDIO_FORMATS.md](AUDIO_FORMATS.md) documents aliases, first-packet/box limits and the 64 KiB read budget. No decoding/transcoding/integrity guarantee or runtime package is added.

New/unfinished allocations resolve final filenames with full identity/budget before durable prepared history. Established record paths stay authoritative. Prior failed/missing allocations lacking both bytes and digest may correct their eventual extension; explicit legacy Redownload keeps its separate allocation hash and original media. Resolved collisions preserve existing files before preparation; concurrent placement remains no-overwrite. Existing byte evidence survives later failures.

Resume keeps exact provisional-path ownership. Corrected prepared paths reconcile already placed files; a preparation crash before checkpoint retirement can leave a path mismatch that stops safely for review, preserving partial, sidecar and history evidence. Strongly checkpointed rejected bodies remain for review without final audio or verified history. Preview performs no enclosure request or persistent write. Writer locks, history schemas 1/2, legacy/diagnostic protections, dates and discovery remain intact; pagination/launcher/release work is unchanged.

## Verification and limits

Settled inventory: **1,184 checks**: **940 units** (930 product units plus ten runner guards; no remaining media-selection characterization) + **244 integrations**. Final local Unit: **940/0/0/0 per engine**, PS7 200.44 seconds/native 146.43 seconds. Final A035 units **71/0/0/869 each**; existing A012 **39/0/0/901 each**. Final audio loopback focus **21/0/0/223 each**, PS7 267.86 seconds/native 163.18 seconds.

Affected archive/body/history/legacy passed **54/0/0/189 each** at the preceding 243-integration snapshot, PS7 640.52 seconds/native 372.05 seconds. Resume passed **34/0/0/210 each** at 244 integrations, PS7 333 seconds/native 250.79 seconds. Affected runs began before the final unfinished-allocation refinement; final unit/format focus and exact-source full CI establish settled-source coverage. Counts are passed/failed/skipped/not_run. No local full All run was performed; exact-head full CI is recorded separately in the PR/final handoff.

Final analyzers passed **69 files, zero parse errors/new findings**, baseline **2 PS7 / 1 native warnings**. Fixture helpers passed **51/0**, separate from product acceptance. Local Windows **10.0.26300.0**, PowerShell **7.6.5 / 5.1.26100.9444**, Pester **5.7.1**, PSScriptAnalyzer **1.24.0**, bundled Python **3.12.14**; process-only PATH/native module handling stays in DEVELOPMENT.md.

Meaningful pre-fix video-first probes failed on both engines; a later unfinished-allocation probe reproduced the provisional extension bug before its fix. Intermediate failures included overly broad privacy/stale cleanup/legacy race assertions and one test-helper analyzer noun, corrected and rerun. [UPD-0205 evidence](evidence/UPD-0205.md) retains exact commands, versions, snapshot counts, durations and limitations.

A001-A036 passed; A037-A060 remain not_run. Synthetic/mock/loopback fixtures and marked owned outputs only. No real archive/private feed, live publisher, UNC/hardware power-loss guarantee, merge, history rewrite, persistent setting change or publication was tested/performed. Stop after UPD-0205; NEXT_THREAD_PROMPT.md describes UPD-0206 on the same M2 branch/draft PR.
