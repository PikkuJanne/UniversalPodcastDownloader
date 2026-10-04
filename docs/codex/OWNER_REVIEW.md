# UPD-0502 owner-review package

Prepared 4 October 2026 (Europe/Berlin). **NOT_READY; UPD-0502 remains blocked.** This package supplies the candidate identities, evidence, limitations and decision boundaries for review. A059's required condition, all mandatory gates met, is still unmet. A060 has no authorized action or publication outcome. No waiver or change to the acceptance scope has been received.

## Required conditions before a readiness decision

| Case | Current condition | Evidence needed |
| --- | --- | --- |
| A041 | `not_run`: Explorer double-click was not observed | Actual Windows Explorer double-click of the reviewed extracted batch launcher, with source/package identity, quoted spaces/Unicode path, guided input, console lifetime and exit behavior recorded. Historical terminal launcher observations remain component evidence. |
| A045 | `not_run`: physical Ctrl+C was not observed | Actual keyboard Ctrl+C during an owned synthetic streaming transfer under Windows PowerShell 5.1 and PowerShell 7; record partial/sidecar preservation, completion/history behavior, cleanup, opt-in keep-awake restoration and process/exit results. Typed exceptions do not supply this observation. |
| A059 | `not_run`: all-mandatory-gates-met precondition unmet | Resolve the required conditions with matching evidence, or record an explicit owner-approved scope change and narrow the claims; then rerun the readiness gate and prepare a final handoff bound to that exact reviewed candidate. |
| A060 | Reviewed `owner_gate`; canonical `not_run` | Separate explicit authorization for an exact action and candidate, followed by verified resulting remote identifiers. The current request authorizes preparation, normal branch pushes and draft PR updates only. |

The canonical register has **56 passed / 4 not_run**; the reviewed mandatory gate has **56 of 59 satisfied**, with A041/A045/A059 unresolved and A060 separately owner-gated. This package does not change those statuses. Available desktop tools cannot observe or control native Windows Explorer/keyboard interaction in this session. See [readiness reconciliation](RELEASE_READINESS.md), [review matrix](RELEASE_ACCEPTANCE.json) and [current evidence](evidence/UPD-0502.md).

## Immutable reviewed predecessor

The task started from the following independently synchronized M5 candidate. These identifiers describe that immutable predecessor; the documentation handoff's eventual head has its own manifest and artifact identifiers, recorded in [draft PR #6](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/6). A commit cannot embed its own hash. Require the final PR's complete verification record before selecting its newer artifact; never treat the predecessor artifact as a final-head artifact.

| Field | Observed identity |
| --- | --- |
| Repository | `PikkuJanne/UniversalPodcastDownloader` |
| Working checkout | `D:\projects\UniversalPodcastDownloader`; original `UniversalPodcastDownloader-main` snapshot preserved |
| M5 branch / PR | `codex/upd-m5-acceptance` / [#6](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/6), OPEN and DRAFT |
| M5 source | `8725f16a7691e92de4f9db45396d5df6138da571` |
| M5 tree | `298badeb6d83974b056cf93132da8cd8cfdb92ea` |
| Unmerged stacked base | `codex/upd-m4-release`, `6a34eb60486e30e7a2631dc0bdbb4d69e643c2b9`, [draft PR #5](https://github.com/PikkuJanne/UniversalPodcastDownloader/pull/5) |
| Observed main | `2ac82614493be7196c9ebee116f23fec07368b50` |
| Version / status | `0.1.0-rc.1` / `UNRELEASED_CANDIDATE` |
| CI | [37151296136](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37151296136), attempt 1; all three jobs successful at the exact source |
| Artifact | [11284770856](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37151296136/artifacts/11284770856) |
| Artifact name | `candidate-8725f16a7691e92de4f9db45396d5df6138da571-37151296136-1` |
| Outer artifact | 6,217,665 bytes; SHA256 `ca50f142472258294cef69d9f3698dfc8dd5c60531e7c4cffbcc591bdb058841` |
| Portable ZIP SHA256 | `d7a10eb798b1755e2e268e618c7f69b951d77d8adc2270a1144b44d3d8021ba4` |
| Manifest SHA256 | `14d2138a516065975f9e11a1aafeb0e9e2e8852f928c85f9b46aa547944b39be` |
| SHA256SUMS SHA256 | `c8353aece3718a9103443ccd695e524c27a367a040c32a3948b9dfee26303af4` |
| Artifact expiry | 10 October 2026, 22:40:46 Europe/Berlin (20:40:46 UTC); review storage, subject to service retention |

Fresh synchronization compares local HEAD, fetched branch, independent `ls-remote` and PR head with divergence 0/0 and a clean working tree. Tags and releases were empty on inspection. Candidate download verifies the three sidecars, API/upload/download outer digest, identical embedded manifest, 37 safe ZIP entries, 36 raw immutable Git payloads and original MIT bytes. Downloaded product code is not executed in this task.

Predecessor hosted PS5.1 `5.1.20348.5622` and PS7 `7.6.6` each passed **All 1862**, with failed/skipped/inconclusive/not_run **0**, and **118-file analysis**, with parse/new/baseline findings **0**. Hosted Windows Server 2022 observations are separate from the active Windows 11 clean-clone observations. [UPD-0501 evidence](evidence/UPD-0501.md) retains local full All1791/analysis116 at immutable `6a34eb6`, 14 actual extracted-package observations and original-byte preservation. The 71 added M5 checks are development readiness-policy checks. Full-suite counts include both product and development checks.

The committed matrix continues to assess tested source `6a34eb60486e30e7a2631dc0bdbb4d69e643c2b9`, tree `f609d23918e840119d9fb3aee22a85e7dd9e1a48`. Final-source equivalence must independently compare all 29 runtime and 36 payload Git blobs, retain original tested-source/engine/evidence identities and preserve every unresolved condition. Runtime fingerprint: `709367285219c0ac0f4f1b8be2d346d6b91d5d6885b3b3afc27372352f745ca1`. Equal payloads support the declared inherited observations; they do not relabel an earlier full local test as a current-source run. Source-based manifest/ZIP hashes change with the documentation commit even when payload bytes remain equal.

## Limitations to retain in any approval

- All covers the bounded accessible catalogue; a publisher may omit historical episodes. Feed discovery, pagination and media formats are bounded by the documented controls.
- Audio is preserved without transcoding. Container checks do not decode the full recording. Completed or unknown media is never overwritten to force migration; legacy adoption starts as unverified and must follow the documented review.
- Preview boundaries concern application data. PowerShell host startup caches are separate. Recovery/rollback must preserve media, history/backups and owned partial/sidecar state; no blanket deletion, blind backup restore or schema rollback is promised.
- Native keep-awake and catchable cancellation have recorded component evidence. Physical Ctrl+C is outstanding; sleep/lid, hard kill, power loss, arbitrary cleanup failure and live UNC durability remain unverified limitations.
- DPAPI is scoped to the current Windows user; compromise of that user/admin account, live memory and previously recorded inputs are outside its protection. The app adds no telemetry; selected feed/media hosts may log requests. Share redacted diagnostics and keep private subscriptions out of tests and public logs.
- The package is unsigned. Launcher execution policy is process-only. Checksums establish byte equality against an expected value, not publisher authentication. Reviewed action pins do not freeze hosted runner images or every upstream dependency byte.
- Website material is portable content/metadata and accurate capture instructions. No real screenshot asset, physical accessibility/display approval, selected live site, download release link or deployment is supplied. Optional signing and website integration remain separate owner choices.
- The MIT license covers the software, not podcast recordings. No runtime Python, Git, Pester, PSScriptAnalyzer, package manager, FFmpeg, cloud account or daemon requirement is introduced.

## Decisions and action record

Use the [approval checklist](../releases/APPROVAL_CHECKLIST.md) with [local draft notes](../releases/DRAFT-0.1.0-rc.1.md), [website content](../website/PRODUCT_PAGE.md), [workflow trust boundaries](RELEASE_WORKFLOW.md) and [publication requirements](RELEASE_AND_WEBSITE.md). Select and preserve the exact reviewed artifact before expiry; recording it here does not select approved distribution bytes.

| Action | Authority / outcome at handoff |
| --- | --- |
| Named documentation commit, normal feature-branch push, existing draft PR update | Authorized implementation workflow; resulting final head/run/artifact are verified and recorded in PR #6 after completion |
| Acceptance waiver / claim change | None received; A041/A045/A059 remain unresolved |
| Merge / rebase / reset / force push | None authorized or performed |
| Tag creation or push / GitHub draft release / release publication | None authorized or performed; intended tag and approved ZIP unassigned |
| Signing, certificate purchase, trust-store changes | None authorized or performed |
| Settings, secrets, visibility, existing branch/issue deletion | None authorized or performed |
| Domain, hosting, DNS, website deployment or purchases | None authorized or performed |

For a future exact action, first record the human approval text, date/time in Europe/Berlin, action, immutable source/tree/version/ZIP hash and intended tag/target. Inspect changed source and rebuild/retest if necessary. Creating a GitHub release for an absent tag may create the tag; verify a separately authorized existing tag before any release action. After each authorized action, independently verify its remote outcome and record exact commit/tag/release/asset identifiers and downloaded hashes. Authority for one action does not authorize another.

Continue **UPD-0502** using [the scoped next prompt](NEXT_THREAD_PROMPT.md) after its missing conditions can be supplied. No next task is invented and no publication decision is requested for an incomplete candidate.
