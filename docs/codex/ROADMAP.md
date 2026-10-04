# Implementation roadmap

Preserve the existing small Windows downloader. Work in the order below, one task per thread. Do not mark future tasks complete because documentation or fixture tooling was added.

| Task | Milestone | Scope | Original review items |
|---|---|---|---|
| UPD-0001 | M0 | Bootstrap handoff and verify repository | Handoff |
| UPD-0002 | M0 | Create efficient regression and CI foundations | 19, 20 |
| UPD-0101 | M1 | Fix collections and safe Windows web requests | 3, 4 |
| UPD-0102 | M1 | Enforce safe deterministic Windows destinations | 5, 2 |
| UPD-0103 | M1 | Make file completion transactional | 1 |
| UPD-0104 | M1 | Add stable identities and durable local history | 2 |
| UPD-0105 | M1 | Protect and migrate legacy archives | 1, 2 |
| UPD-0106 | M1 | Make startup diagnostics private and reliable | 7 |
| UPD-0107 | M1 | Enforce preview and untrusted-input boundaries | 8, 9 |
| UPD-0201 | M2 | Implement bounded retry and timeout policy | 10 |
| UPD-0202 | M2 | Add validator-aware safe resume | 11 |
| UPD-0203 | M2 | Unify feed discovery and resolution | 12 |
| UPD-0204 | M2 | Normalize publication dates and stable ordering | 13 |
| UPD-0205 | M2 | Select and preserve supported audio formats | 14 |
| UPD-0206 | M2 | Follow explicit feed pagination safely | 15 |
| UPD-0301 | M3 | Finalize CLI, results and launcher semantics | 6, 17 |
| UPD-0302 | M3 | Polish cancellation, preflight and concurrent-run safety | 16 |
| UPD-0303 | M3 | Improve progress and optional temporary keep-awake | 16 |
| UPD-0304 | M3 | Add opt-in saved shows and sequential batch mode | 18 |
| UPD-0401 | M4 | Complete maintainability and compatibility audit | 19, 20 |
| UPD-0402 | M4 | Write verified help and user documentation | 21 |
| UPD-0403 | M4 | Build traceable release ZIPs and checksums | 22 |
| UPD-0404 | M4 | Prepare secure GitHub release workflow | 22 |
| UPD-0405 | M4 | Prepare website-ready distribution content | 23 |
| UPD-0501 | M5 | Run release-candidate acceptance from a clean clone | 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23 |
| UPD-0502 | M5 | Owner-review handoff and publication gate | 22, 23 |

## Milestone branches

M0: `codex/upd-m0-foundation`; M1: `codex/upd-m1-safety`; M2: `codex/upd-m2-network-feeds`; M3: `codex/upd-m3-workflow`; M4: `codex/upd-m4-release`; M5: `codex/upd-m5-acceptance`. Reuse an existing correct branch rather than create duplicates. See GITHUB_WORKFLOW.md for unmerged predecessor handling.

## Public-release gates

All mandatory downloader behavior cases require evidence on the declared Windows engine matrix. Manual cases remain outstanding until actually performed. Signing/purchases, repository settings, merges, release tags/publication and website deployment are separate owner gates. The handoff covers every review item but does not authorize unrelated web hosting or a rewrite.
