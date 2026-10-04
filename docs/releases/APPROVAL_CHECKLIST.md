# Candidate review and publication approval

Status: **UNRELEASED_CANDIDATE / NOT_READY; no publication approval recorded**. Updated 4 October 2026 (Europe/Berlin). UPD-0405 portable website content and UPD-0501 acceptance reconciliation are complete. UPD-0502 supplies the [owner-review package](../codex/OWNER_REVIEW.md), but remains blocked by required A041/A045 observations and A059's all-mandatory-gates-met precondition. A060 remains a separate action gate.

## Review the concrete candidate

- [ ] Confirm the exact version, source commit/tree, clean branch/PR head and successful two-engine tests/analysis. Record run URLs and actual passed/failed/skipped/not-run counts; do not substitute historical or helper-only results.
- [ ] Resolve all mandatory acceptance cases. Current canonical status is56 passed/4 not_run; A041 Explorer double-click, A045 physical Ctrl+C and A059 final gate-qualified handoff are unresolved. A060 has no authorized/performed publication action. Record any explicit owner-approved scope change and narrow feature claims accordingly; the current offline policy rejects deferral flags, so a deliberate claims/policy review would be required.
- [ ] Download the exact run's candidate artifact into a new review directory. Record artifact ID/name/digest and run attempt. CI artifacts expire after 7 days; retain an explicitly selected reviewed copy before expiry. Storage of a candidate does not authorize release publication.
- [ ] Independently verify ZIP/manifest hashes from `SHA256SUMS`, the identical embedded manifest, all payload hashes and the committed allowlist. Confirm the full immutable source URL, version and tree against GitHub. The outer GitHub artifact digest and ZIP checksum identify different byte streams.
- [ ] Treat PR/fork-produced artifact data as untrusted. Review the workflow and candidate source, use ordinary hosted runners, and do not execute downloaded candidate files with credentials or in a privileged release workflow. An action SHA or checksum does not make malicious source safe.
- [ ] Verify fresh extraction in a path with spaces and non-ASCII characters on both supported Windows engines. Review native launcher/help, original audio bytes, repeat/preview behavior and preservation of copied legacy/archive state. Test with synthetic copies only.
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
| Approved source/tree/ZIP hash | Unassigned; select the exact fully reviewed candidate only after unresolved readiness conditions are addressed |
| Intended tag | Unassigned; no tag creation/push authorized |
| Outcome/URLs | None; no release, merge, settings or deployment action |

See [draft notes](DRAFT-0.1.0-rc.1.md), [workflow design](../codex/RELEASE_WORKFLOW.md) and [release requirements](../codex/RELEASE_AND_WEBSITE.md).
