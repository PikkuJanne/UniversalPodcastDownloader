# Local and GitHub working agreement

## Authority and boundaries

GitHub is the canonical source for continuation. The active Codex machine performs authoritative development/testing; CI is supplementary. This handoff was prepared read-only and has not created any branch, issue, PR, tag or release.

The user's requested implementation workflow includes local code/docs/tests, feature branches, commits, normal pushes and draft PRs. Do not infer permission for merges, rebases, resets, force pushes, existing branch/issue closure/deletion, changes to protection/visibility/secrets, signing purchases, release tags/publication or a live website deployment. Prepare those as explicit owner gates.

## Start every thread

Read current repository guidance; identify actual root. Check origin identity without publishing embedded credentials. Inspect branch, HEAD, upstream, tracked/untracked changes and unfinished merges. Fetch current remote references without changing local work. Compare branch ancestry and look for an existing PR. Do not automatically stash or clean a dirty tree. Preserve unrelated edits and report conflicting ownership.

Illustrative read-only commands after confirming the directory is correct:

```powershell
git rev-parse --show-toplevel
git status --short
git branch --show-current
git log -1 --format='%H %s'
git remote get-url origin
git fetch origin
git branch -vv
```

Never print a credential-bearing remote URL in committed evidence. `git fetch` updates remote-tracking refs; it is not a substitute for reviewing divergence. A changed main is not a reason to reset to BASELINE.json.

## Branches and draft PRs

Use one branch per milestone, named in ROADMAP.md. Each thread commits a coherent task on that branch. Reuse a relevant existing branch/PR rather than opening duplicates. The first bootstrap task can open the M0 draft PR with handoff files only; its body must not claim runtime fixes.

If the predecessor milestone is merged, create the next branch from fetched/reviewed current main, advancing a clean local main only by a safe fast-forward when appropriate. If it is not merged and work should continue, create the next branch from the predecessor's actual tip and use a **stacked draft PR whose base is that predecessor branch**. State the dependency plainly. Do not merge it yourself merely to start the next milestone. After owner merges a predecessor, inspect GitHub's base handling and retarget only as appropriate; do not rebase/reset/force-push to tidy history without approval.

If GitHub CLI exists and is already authenticated, it can create/update draft PRs. A connected GitHub tool is also acceptable. Use named repository/base/head values and the actual existing PR. Do not assume CLI installation or authentication; check it. Do not create 26 separate issues by default. The versioned task register and milestone PR checklist are sufficient; an umbrella issue is optional if useful and nonduplicative.

## End every thread

Run focused checks and applicable broader gates; inspect the diff and staged files for secrets, archive paths, generated media/logs and unrelated edits. Update STATUS, TASKS, sanitized evidence and the next-thread prompt. Stage explicit intended paths. A mixed pre-existing working tree is not permission for blanket staging.

Commit the coherent unit and push normally to the verified feature branch. Verify remote state, not just the command's success message:

```powershell
git rev-parse HEAD
git ls-remote --heads origin 'refs/heads/codex/upd-m0-foundation'
```

Use the actual current milestone branch, not the example name. The two hashes must agree after a successful push. Return final SHA/branch/PR in the user-facing handoff. If push/authentication/checks fail, report local-only status and exact reason; never claim the task is synchronized. Do not fabricate credentials or broaden access.

The final commit cannot embed its own hash in STATUS.md. Store the tested source commit and last remote observation, then report final HEAD after the documentation checkpoint. Update the draft PR with that final evidence without introducing a self-referential commit loop.

## CI and release safety

Create a small Windows engine matrix in UPD-0002 and pin actually verified dependency versions. Use reviewed full action commit SHAs, least-privilege workflow tokens, bounded jobs and concurrency. Do not execute untrusted PR code with privileged pull_request_target, pass raw feed URLs into shell commands, publish sensitive artifacts, or label an unrun engine as passing. See official GitHub guidance [S10 in SOURCES.md].

No owner settings changes are made automatically. Record recommendations such as required checks separately. Workflows should validate/build candidates, not automatically tag or publish a release whenever code is pushed. Draft release preparation must not cause implicit creation of a release tag. Publication stays separately authorized.
