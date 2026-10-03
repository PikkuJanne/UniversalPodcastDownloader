# Verification strategy

## Layers and honest labels

**Bundle checks** validate this handoff's structure, checksums and helper server. They say nothing about downloader correctness.

**Product unit tests** are Pester tests Codex must implement against import-safe functions. Mock external calls; use synthetic XML/HTML/names/config. Inject clocks/delays for retry tests. Pester mocking is documented in SOURCES.md [S8].

**Product integration tests** use the supplied loopback server or an equivalent Windows-native harness plus temporary output folders. Assert actual bytes, state transitions, request counts, process results and side effects. No live podcast host in the default suite.

**Windows manual checks** cover launcher behavior, cancellation/keep-awake, fresh extraction and accessibility/readability. A Linux or WSL PowerShell run does not prove Windows PowerShell 5.1 behavior.

**Optional live smoke** is a separately enabled small public feed/episode check after local tests. No private subscription, full catalogue or real user archive by default. Publisher outages are not core deterministic-test failures.

## Required engine matrix

Windows PowerShell 5.1 and a currently supported PowerShell 7.x on Windows. Record exact OS, runtime and development module versions. Check support/version details at implementation time; do not upgrade the user's system without permission. Both engines must parse and exercise runtime code. If one is unavailable, record not-run and keep its release gate open. A CI job can supplement but not disguise missing required local evidence.

UPD-0002 chooses and records Pester/PSScriptAnalyzer versions compatible with the matrix, preferably without administrator installation. Development dependencies are not runtime requirements. Do not silently install modules on every downloader launch. Expected runner contracts (to be implemented, not already present):

```powershell
.\scripts\Test.ps1 -Suite Unit
.\scripts\Test.ps1 -Suite Integration
.\scripts\Test.ps1 -Suite All
.\scripts\Analyze.ps1
```

The names are proposed contracts; document actual names in AGENTS.md once implemented. Runners must fail on missing tools or zero unexpectedly discovered tests, propagate failures and report skip counts. They must not claim success because no test files existed.

## Starting from a flawed baseline

Establish baseline smoke and characterize pure-function behavior. Add a focused failing regression when implementing each fix and record that it failed before/ passed after. Do not mark all future acceptance tests as skipped and call the project passing. Future requirements remain `not_run` in ACCEPTANCE_CASES.json until implemented/tested; distinguish current CI baseline from public release readiness.

## Fixture kit

`tools/codex-handoff/fixture_server.py` is Python standard-library **development tooling only**. Python is never required to run the downloader. Where Python is unavailable, Codex may implement an equivalent loopback harness or document that integration prerequisite; do not make the app install it.

The server binds 127.0.0.1 only and serves explicitly defined routes, generated silent MP3 and named synthetic feed fixtures. It never proxies the web or reads arbitrary client-specified paths. Test-only state/control endpoints are local. Use an ephemeral port, owned subprocess and temporary ready-file; ensure cleanup stops only that subprocess. No administrator HTTP listener registration or machine-wide firewall changes.

Read its README and self-tests for tested helper behaviors. A successful helper self-test is not an A011/A027 downloader pass; product tests must actually invoke the downloader against it and assert artifacts/outcomes.

## Safety assertions

Every integration test allocates a new temporary root with an ownership marker; never resolve the user's real default Downloads/Podcasts as test output. Cleanup may remove only those marked test roots. Legacy scenarios operate on generated files or explicit copies; before/after SHA-256 assertions prove that original copies did not change. Junction/reparse tests must be contained and self-owned.

Mocked external URLs use reserved `.invalid` hosts. The XML DTD fixture's external target must never be fetched. Assert no network during parsing and no media requests or persistent filesystem effects during WhatIf. Redaction tests use fake values only and inspect every output channel/export. Include malformed history and unknown schema, corrupted sidecars, zero media length, feed titles dot-dot, same-date collisions, invalid dates and unsupported schemes.

## Evidence records

UPD-0302 run-safety checks use real owned workers and actual streamed bytes. Gate a writer only after its checkpoint callback; reject a same-show contender without another media request, run another show concurrently, release only tracked workers and reopen stale persistent lock files without interpreting their contents. Assert history validity, original hashes, cancellation 130 and exclusive partial reopening. Resume recovery must validate the actual final checkpoint, request only the remaining tail and produce the original complete media hash before a verified repeat skip.

Permission fixtures deny CreateFiles or CreateDirectories only on marked owned directories for the current user, save the original ACL and restore it in finally before cleanup. Prove separate CreateFiles permission when testing denied directory creation. Test-only capacity providers return observed zero or unknown; the actual HTTP pipeline still validates headers/ranges and copies bodies. Assert zero body progress on reliable insufficient-space rejection, no enclosure-metadata hard gate for an unknown body and no probe/capacity/archive work under preview. Cleanup fault units verify primary-error precedence and release actual media guards before injecting disposal errors; a failing Dispose is not a universal release guarantee.

After restoring a fixture ACL, require exact owner/group, resource-manager control, every other descriptor flag and complete ordered DACL/SACL binary contents. Windows may add only the observed DACL automatic-inheritance bookkeeping bit (`AI`, `0x0400`); never permit its removal or changes to protection, inheritance requirements or ACE flags. Literal SDDL equality incorrectly rejected that round-trip on Windows Server 2022. This normalization does not skip permission/preservation assertions or relax restored access rules. [Microsoft control flags](https://learn.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.controlflags), [security descriptor format](https://learn.microsoft.com/en-us/windows/win32/secauthz/security-descriptor-string-format)

For each case record task ID, case ID, source commit, actual command, engine version, date, result and sanitized evidence path. Capture failed/skipped/not-run explicitly. Summaries should be compact enough for future threads; store huge generated logs as ignored local output, not repository history. Re-run impacted cases after architecture changes; prior evidence is not automatically applicable to a changed release.

## Release gate

All mandatory A001–A059 cases must be satisfied or explicitly documented as an owner-approved scope change. A060 applies only when a publication action is authorized and performed; until then mark `owner_gate`/not applicable to implementation readiness, not failed. Optional signing and live website deployment are owner-gated, not hidden implementation claims. `ready_for_owner_review` is different from `released`.
