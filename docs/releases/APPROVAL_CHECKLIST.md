# Candidate review and publication approval

**v1.0.0 follow-up, 8 October 2026:** the owner explicitly requested: “We are on v1.0.0 now so please make sure v1.0.0 is published on GitHub.” This authorizes the stable version checkpoint, verified tag and publication of the reviewed ZIP, manifest and checksums. The agent selects and verifies those bytes within this request; no additional approval is required for the same action. See [v1.0.0 action record](../codex/evidence/V1.0.0-PUBLICATION.md), [release notes](v1.0.0.md) and the [version-specific GitHub release](https://github.com/PikkuJanne/UniversalPodcastDownloader/releases/tag/v1.0.0). Website assets/deployment, signing, purchases and repository settings remain outside scope.

**Historical 4 October candidate checklist follows.** Its unchecked boxes and NOT GRANTED approval record describe that earlier preparation task, not the later explicit publication request. The frozen candidate acceptance records retain their original source and evidence identities.

Status: **UNRELEASED_CANDIDATE / ready_for_owner_review; no publication approval recorded**. Updated 4 October2026 (Europe/Berlin). UPD-0405 portable website content, UPD-0501 acceptance reconciliation and UPD-0502 [owner-review preparation](../codex/OWNER_REVIEW.md) are complete. Actual human Explorer and both-engine physical Ctrl+C observations resolve A041/A045; all59 mandatory cases, including A059's handoff, are supported. Canonical59 passed/1 not_run. A060 remains a separate action gate; releaseAuthorized=false and published=false.

## Review the concrete candidate

- [ ] Confirm the exact version, source commit/tree, clean branch/PR head and successful two-engine tests/analysis. Record run URLs and actual passed/failed/skipped/not-run counts; do not substitute historical or helper-only results.
- [x] Resolve all mandatory acceptance cases. Current canonical status is59 passed/1 not_run; A041 Explorer double-click, A045 physical Ctrl+C and A059 handoff have supported evidence. [Manual observations](../codex/evidence/UPD-0502-MANUAL.md) are bound to the unchanged extracted ed3 candidate; no waiver or policy change was used. A060 has no authorized/performed publication action. Review current gate results/source binding in the final handoff before selecting an action.
- [ ] Download the exact run's candidate artifact into a new review directory. Record artifact ID/name/digest and run attempt. CI artifacts expire after 7 days; retain an explicitly selected reviewed copy before expiry. Storage of a candidate does not authorize release publication.
- [ ] Independently verify ZIP/manifest hashes from `SHA256SUMS`, the identical embedded manifest, all payload hashes and the committed allowlist. Confirm the full immutable source URL, version and tree against GitHub. The outer GitHub artifact digest and ZIP checksum identify different byte streams.
- [ ] Treat PR/fork-produced artifact data as untrusted. Review the workflow and candidate source, use ordinary hosted runners, and do not execute downloaded candidate files with credentials or in a privileged release workflow. An action SHA or checksum does not make malicious source safe.
- [ ] Verify the chosen final candidate's extraction evidence in a path with spaces and non-ASCII characters on both supported Windows engines. The observed ed3 launcher succeeded by actual Explorer double-click and guided input; numeric current Explorer child exit was unavailable. Both Sandbox ConsoleHosts received human physical Ctrl+C and restored their real guest native leases. Retained partials483328/139264 bytes match synthetic input, but saved65536-byte sidecars mismatch full size/hash: strict resume would refuse and preserve for review. Historical clean-clone launcher/help/repeat/preview/copied-legacy evidence retains its original source. Use synthetic copies only; no host sleep/lid, successful-resume or exit130 guarantee is inferred.
- [ ] Review privacy/security limits, partial catalogue claims, unsigned/process-only launcher behavior, recovery, original license and artwork, and the draft notes. Confirm no private URLs/tokens/configuration/history/media/logs are present in release assets.
- [ ] Choose the exact reviewed ZIP produced by one engine and record its hash. Compression may differ across engines; never replace bytes after approval without a new review.

## Obtain explicit action approval

- [ ] Record the owner's approval text, date/time (Europe/Berlin), exact candidate commit/tree/hash/version, intended tag and requested action. Approval to review a PR does not itself authorize merging, tagging, publishing, purchases, settings changes or deployment.
- [ ] If a merge is authorized, inspect the resulting immutable source and rebuild/retest if it differs from the approved candidate. Do not carry an old head's approval onto changed bytes.
- [ ] Treat tag creation/push as its own authorized side effect. Inspect existing tags/releases first; never replace an existing tag, force-push or invent the target commit.
- [ ] Before creating even a GitHub draft release, verify the exact intended remote tag already exists and targets the approved source. Creating a release for a missing tag can implicitly create that tag. Keep these local draft notes until tag authority is explicit.
- [ ] Publication must name the approved tag and exact asset bytes. The automation added by UPD-0404 has no release-creation or publication step and no contents-write token.
- [ ] Authenticode/signing, certificate purchases, trust-store changes, repository settings/visibility/secrets, branch deletion and live hosting each need their own approval if desired. None are prerequisites created by this checklist.

## Verify an authorized outcome

- [ ] After a separately authorized action, record resulting immutable commit/tag/release/asset URLs and independently verify downloaded asset hashes. Update distribution links to exact version assets only after actual publication.
- [ ] Record any failure or partial action honestly. Preserve the reviewed source/artifact; do not retry with broader authority or change owner data/settings.

## Approval record

| Field | Record |
| --- | --- |
| Owner approval | **NOT GRANTED / NOT REQUESTED by this task** |
| Approved action(s) | None |
| Approval date/time | Not applicable |
| Candidate version | `0.1.0-rc.1` / `UNRELEASED_CANDIDATE` |
| Observed candidate identities | [Owner package](../codex/OWNER_REVIEW.md) and PR #6's final-head verification record; these are review identities, not approved source/assets |
| Approved source/tree/ZIP hash | Unassigned; owner must select the exact reviewed candidate and separately authorize an action. Preparation readiness does not approve distribution bytes. |
| Intended tag | Unassigned; no tag creation/push authorized |
| Outcome/URLs | None; no release, merge, settings or deployment action |

See [draft notes](DRAFT-0.1.0-rc.1.md), [workflow design](../codex/RELEASE_WORKFLOW.md) and [release requirements](../codex/RELEASE_AND_WEBSITE.md).
