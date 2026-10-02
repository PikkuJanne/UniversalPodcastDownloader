# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0107 only**.

- UPD-0001, UPD-0002 and UPD-0101 through UPD-0106: done; historical evidence preserved.
- UPD-0107: **done**; implementation pushed and verified; full local checks and both GitHub jobs passed.
- Next, ready: **UPD-0201 — implement bounded retry and timeout policy**; not started.
- Branch/upstream: `codex/upd-m1-safety` / `origin/codex/upd-m1-safety`.
- M1 [draft PR #2](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/2) remains stacked on `codex/upd-m0-foundation`; M0 draft PR #1 remains unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0107 evidence](evidence/UPD-0107.md), [input and preview policy](INPUT_BOUNDARIES.md), [diagnostics](DIAGNOSTICS.md), [history and migration](STATE_AND_MIGRATION.md).

## Verified checkout and changes

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Preserved `UniversalPodcastDownloader-main` snapshot untouched. Origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/fetched remote/PR HEAD `0d05898fee3dc7d2fb7ec154508c68ad17dc3301` was clean. [Predecessor CI 37010095589](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37010095589) passed both jobs at that exact SHA. Continue command-scoped authenticated HTTPS and public author identity without persistent settings changes.

Metadata and media now use a shared HTTP(S) request policy with manual validated redirects, at most five hops, no HTTPS downgrade/userinfo/default credentials/cookies and unchanged TLS/proxy defaults. Private and loopback destinations remain supported. Metadata is streamed into memory with an 8 MiB cap and strict decoding. XML uses explicit DTD prohibition/null resolvers plus character, depth and node/attribute limits. HTML relative links use the effective page URL; feed identity retains the original requested URL.

The existing read/plan/ShouldProcess/locked-revalidation path validates all parsed enclosure targets before archive writes. Preview and declined confirmation cannot create directories, media, state, locks, logs or exports, and cannot report completed downloads. A later real run still downloads and verifies bytes. History schemas, no-overwrite behavior, original audio, modes and entry points remain intact. No runtime dependency was added.

## Checks and limits

Focused A023 units pass **66/0/0/435** on both supported local engines; new integrations pass **20/0/0/91** each. Helper tests pass **28/0** separately. Settled suite contains 612 checks: 491 product units, ten runner guards and 111 integrations, including two remaining parser characterizations. Full local All passed **612/0/0/0 on each engine** (PS7: 761.38 seconds; native: 436.37 seconds). Full CI passed the same counts on both engines. Analysis passes on 50 PowerShell files with zero parse errors/new findings and 5 PS7 / 4 native baseline warnings.

Local Windows 10.0.26300.0; PowerShell 7.6.5 / 5.1.26100.9444; Pester 5.7.1; PSScriptAnalyzer 1.24.0; bundled Python 3.12.14. CI Python stays pinned to 3.14.7. Use the bundled Python directory via process-only PATH because the ambient alias is unusable. Tests isolate logs and archives in marked owned temporary data.

A001-A024 passed; A025-A060 remain not_run. The policy intentionally rejects compressed metadata and unsupported text encodings. A selected metadata URL can itself return audio; preview never initiates planned enclosure transfers. Body/XML size limits do not provide idle/total transfer deadlines; UPD-0201 owns retry/timeout work. No proven prior external XML exploit, real-archive testing, manual launcher verification, merge, history rewrite, persistent settings change, release or publication is claimed.


## GitHub checkpoint and continuation

Implementation `388054d439ac3ee7ee336297148c5dd58683f2fa` is pushed and independently equals the remote. [CI 37012792338](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37012792338) passed both Windows jobs at that exact SHA: 612/0/0/0 each, 50 analyzed files, zero parse errors/new findings and 5/4 baseline warnings. CI engines are PowerShell 7.6.6 and Windows PowerShell 5.1.20348.5622. The final documentation checkpoint and exact-head CI are verified separately in draft PR #2/final handoff; runtime and tests are unchanged.

M1 implementation tasks are complete; the PR remains draft and unmerged. Stop after UPD-0107. NEXT_THREAD_PROMPT.md describes UPD-0201 and the M2 branch/stacked-PR procedure; M2 is unstarted.
