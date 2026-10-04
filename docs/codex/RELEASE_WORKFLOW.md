# Candidate workflow and trust boundaries

UPD-0404 prepares review candidates only. The [workflow](../../.github/workflows/test.yml) runs tests and analysis on both supported Windows engines, then builds/uploads a candidate in a separate fresh hosted Windows job after both matrix jobs succeed. It never creates a tag or GitHub release. The local [draft notes](../releases/DRAFT-0.1.0-rc.1.md) and [approval checklist](../releases/APPROVAL_CHECKLIST.md) are the release material; no GitHub draft release is created for an absent tag.

## Immutable tested source

For `pull_request`, the checkout is the event's full PR head SHA; for `push` or `workflow_dispatch`, it is the event's `github.sha`. A 40-hex check and actual Git HEAD comparison precede execution. The test matrix and candidate job use the same source, and the candidate manifest identifies that exact commit/tree. This replaces the historical default PR merge checkout for new runs. These runs verify the head's candidate; they do not claim an independently tested future merge result. Rebuild and retest changed source after a separately authorized merge.

The workflow has no ref/version/feed-URL inputs. Context values reach PowerShell through environment variables, not generated shell source. Checkout receives an immutable ref and does not persist credentials. Candidate preparation reads only committed package files through the existing builder, then verifies the three-file output and selected source before emitting fixed step-output names. A separate hosted job prevents test-created workspace files from becoming package inputs. No test artifact is downloaded into that job.

## Permissions, triggers and artifacts

Workflow permissions start empty; each source-reading job grants only `contents: read`. No custom secret, environment deployment credential, OIDC, contents-write or release token is supplied. GitHub's job-scoped artifact service credential remains necessary to upload a candidate; it does not grant repository publication authority. Triggers are `pull_request`, main-branch `push` and input-free `workflow_dispatch`. There is no `pull_request_target`, `workflow_run`, release event, tag-push trigger, cache action, self-hosted runner or privileged follow-on consumer. Jobs/concurrency are bounded; the matrix retains independent engine results.

Upload uses an immutable reviewed action, three explicit validated paths, failure on missing files, hidden files disabled, overwrite disabled, outer compression disabled and 7-day retention. Only the versioned ZIP, `manifest.json` and `SHA256SUMS` are stored. Raw stdout/stderr/test results, fixtures, `.git`, downloaded media, history, saved settings, private URLs and development modules are never selected for upload. Hosted test console logs still contain synthetic fixture observations; no real podcast archive or private subscriptions are used.

The artifact name includes full source SHA and run ID/attempt. The upload action's artifact ID/digest and successful run URL identify the outer artifact; the portable ZIP's checksum identifies a different byte stream. Inspect the actual download and compare its manifest, checksums and all ZIP entries with committed source before considering distribution. A successful PR artifact is **UNRELEASED_CANDIDATE**, including artifacts from forks. A PR may change its own code/workflow, so green tests, action pins and checksum consistency do not confer trust on malicious source. Keep untrusted artifacts out of privileged workflows.

## Reviewed pins and remaining dependency trust

Reviewed on 3 October 2026 against the official action repositories' tag refs, commit identities, `action.yml`, input behavior and source entry points:

| Action | Full commit | Reviewed release |
| --- | --- | --- |
| `actions/checkout` | `3d3c42e5aac5ba805825da76410c181273ba90b1` | `v7.0.1` |
| `actions/setup-python` | `5fda3b95a4ea91299a34e894583c3862153e4b97` | `v7.0.0` |
| `actions/upload-artifact` | `043fb46d1a93c77aae656e7c1c64a875d1fc6a0a` | `v7.0.1` |

Checkout/setup-python pins are retained. Python is fixture-only and remains exact `3.14.7`, x64, with `check-latest: false` and no package-dependency cache configured. Pester `5.7.1` and PSScriptAnalyzer `1.24.0` are installed repository-locally from the existing version/URL/SHA-256 lock in [devtools.json](../../tools/devtools.json); no global installation or new app dependency is added.

Review confirms the configured action interfaces and relevant behavior; it is not a formal audit of every bundled transitive byte. SHA pins protect action references against moving tags. They do not freeze the hosted runner image, Git/PowerShell binaries, upstream Python manifest/tool downloads or attest to publisher identity. Runtime inventory and candidate source/hash evidence remain necessary. Recommendations to require checks/review or restrict Actions are owner configuration decisions; this task changes no repository controls.

## Validation and publication gate

Workflow-policy regression tests reject unsafe permission/trigger/ref/action/upload changes, and actual candidate tests exercise refusal and source/checksum boundaries with synthetic committed data. A separate verified one-off `actionlint` checks YAML, GitHub expressions and job dependencies; it is development validation, not an application dependency. Final CI additionally runs full product tests, analysis and actual candidate preparation/upload. Exact commands/counts and review observations belong in [UPD-0404 evidence](evidence/UPD-0404.md) and the final draft PR.

GitHub's release API can create a missing tag even when creating a draft. Keep draft notes local until separately authorized tag creation and exact existing-tag verification. No release creation/publication command is included in the workflow. A checksum alone is not publisher authentication; signing, certificate spending and trusted certificate stores remain optional owner gates. UPD-0405 content and UPD-0501 reconciliation are complete. The [UPD-0502 package](OWNER_REVIEW.md) records NOT_READY: A041/A045/A059 remain unresolved and A060 requires separate exact action authority.

Primary sources checked during this task: [GitHub secure workflow use](https://docs.github.com/en/actions/reference/security/secure-use), [artifact storage](https://docs.github.com/en/actions/tutorials/store-and-share-data), [release API tag behavior](https://docs.github.com/en/rest/releases/releases#create-a-release), and the three official action repositories at the pinned commits above. See [sources](SOURCES.md) for provenance and [acceptance](ACCEPTANCE.md) for task-level distinctions.
