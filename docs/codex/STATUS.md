# Current implementation status

Updated 2026-10-02 (Europe/Berlin). Scope: **UPD-0202 only**.

- UPD-0001, UPD-0002, UPD-0101 through UPD-0107 and UPD-0201: done; historical evidence preserved.
- UPD-0202: **done**; implementation pushed and verified; local coverage and both GitHub jobs passed.
- Next, ready: **UPD-0203 — unify feed discovery and resolution**; unstarted.
- Branch: `codex/upd-m2-network-feeds`.
- [M2 draft PR #3](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/3) remains stacked on `codex/upd-m1-safety`. M1 PR #2 and M0 PR #1 remain open, draft and unmerged.
- [Commands](DEVELOPMENT.md), [UPD-0202 evidence](evidence/UPD-0202.md), [safe resume](RESUME_POLICY.md), [retry policy](TRANSPORT_POLICY.md), [history](STATE_AND_MIGRATION.md), [diagnostics](DIAGNOSTICS.md).

## Verified checkout and changes

Actual checkout: `D:\projects\UniversalPodcastDownloader`. Preserved `UniversalPodcastDownloader-main` snapshot untouched. Origin: `git@github.com:PikkuJanne/UniversalPodcastDownloader.git`. Starting local/fetched remote/PR HEAD `2092c55a2d3aac61bee4742813a44a2c38e2b815` was clean with 0/0 divergence. [Predecessor CI 37018152755](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37018152755) passed both jobs at that exact SHA. Continue command-scoped authenticated HTTPS and public author identity without persistent settings changes.

Resume now requires strict owned sidecar evidence, matching identities and representation, a strong ETag, known total and exact local prefix size/hash. Only validated complete-tail 206 responses append. Ignored or inconsistent ranges, changed validators/representation and 416 use one fresh GET into a new file and preserve old evidence. Corrupt state or uncheckpointed crash tails stop for review. Caught network failures checkpoint their latest bytes for the existing bounded retry loop.

The sidecar is separate from history schemas 1/2. Media and metadata are flushed before atomic checkpoint replacement, under exclusive handles and the archive writer lock. Progress checkpoints are geometrically spaced. Prepared history evidence, no-overwrite final placement, preview, legacy protections, private diagnostics and original audio bytes remain intact. No runtime dependency or option was added.

## Checks and limits

Settled inventory: **829 checks** (632 product units, ten runner guards, 187 integrations), including two remaining parser characterizations. Final local Unit passed **642/0/0/0 on each engine**. Broad integration snapshots ran **186/1/0/0 per engine**: the early A009 cleanup failure was fixed while remaining workers continued. Its focused rerun passed **2/0/0/185 each**; all 34 resume cases passed in both full integration runs. The final clean-checkout CI passed **829/0/0/0 per engine**. Evidence preserves the initial failures and distinguishes these snapshots from a single all-green local run.

Analyzers: **60 files, zero parse errors/new findings**, unchanged 5/4 baseline warnings. Fixture helpers: **41 passed / 0 failed**, separate from product acceptance. Local Windows 10.0.26300.0; PowerShell 7.6.5 / 5.1.26100.9444; Pester 5.7.1; PSScriptAnalyzer 1.24.0; bundled Python 3.12.14. CI uses Python 3.14.7. Use process-only Python PATH and native module-path handling from DEVELOPMENT.md. Tests isolate logs and archives in marked owned temporary data.

A001-A029 passed; A030-A060 remain not_run. Resume deliberately rejects uncertain crash tails and some RFC-permitted response forms. Unknown partials and old sidecar snapshots are retained. No live archive, private feed, power-loss/UNC guarantee, discovery expansion, launcher redesign, merge, history rewrite, persistent settings change, release or publication is claimed.

## GitHub checkpoint and continuation

Initial implementation `6ae67ab95067c75114ae3fe4d034075ab960611b` exposed a fresh-partial cleanup regression at the late destination boundary. Correction `f32238923367419f29f0151220ce417a008a77d9` is pushed and independently matches the remote checkpoint. [CI 37023755907](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37023755907) passed both Windows Server 2022 jobs: **829/0/0/0 each**, 60 analyzed files, zero parse errors/new findings and 5/4 baseline warnings. CI engines: PowerShell 7.6.6 and Windows PowerShell 5.1.20348.5622. The final documentation checkpoint records these results and the next task; its SHA and exact-head CI are recorded in draft PR #3/final handoff without a self-referential commit loop.

Stop after UPD-0202. NEXT_THREAD_PROMPT.md describes UPD-0203 on the same M2 branch and draft PR.
