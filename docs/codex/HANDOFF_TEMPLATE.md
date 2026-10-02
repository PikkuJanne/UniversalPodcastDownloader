# Task handoff — UPD-xxxx

## Scope and outcome
Task ID/title, status (done/in_progress/blocked), exact implemented behavior, and explicit exclusions. Cite accepted case IDs, not just a general assertion that tests pass.

## Repository state
Verified owner/repository; branch and PR; tested source commit; last verified remote SHA with observation time; dirty/untracked files remaining and their owner. Final response reports final HEAD after the documentation commit.

## Actual verification
| Command | Engine/tool version | Source commit | Passed | Failed | Skipped/not run | Evidence |
|---|---|---|---:|---:|---:|---|
| Record actual command | Record actual version | Actual SHA | Actual count | Actual count | Actual count and reason | Sanitized path |

State separately: helper tests, product unit tests, product integration tests, manual Windows checks, CI. Do not substitute one for another. Note every unavailable engine and pending acceptance gate.

## Safety review
Existing-archive preservation, overwrite avoidance, credential/log review, WhatIf side effects, runtime dependencies and compatibility. Record new assumptions and decisions.

## Next thread
Exact task ID, relevant file paths, what was completed, what must not be repeated, known blocker, starting branch and recommended first test. Include one ready-to-paste prompt. Do not ask the next thread to rediscover the project.
