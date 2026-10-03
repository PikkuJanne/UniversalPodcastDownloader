# 0.1.0-rc.1 — draft release notes

Status: **UNRELEASED_CANDIDATE**. These are local review material, not a GitHub release or a stable-version announcement. No tag or release exists at the preparation observation on 3 October 2026. Final acceptance and publication approval are still outstanding. See the [approval checklist](APPROVAL_CHECKLIST.md) before any distribution decision.

## What the candidate contains

UniversalPodcastDownloader is a portable Windows RSS/Atom podcast archiver. It keeps the original PowerShell and batch entry points, Latest/Custom/All selection, configurable output folders, and original downloaded audio bytes. The application requires Windows PowerShell 5.1 or PowerShell 7 on Windows; no Python, Git, package manager, FFmpeg, cloud account or daemon is required to run it.

The candidate adds bounded feed discovery and catalogue traversal, stable episode identities, Windows-safe naming, verified download history, conservative review of legacy files, and resume of strictly owned partial transfers. Private run results and meaningful exit codes support unattended use. Saved-show batches are sequential, with credentials protected for the current Windows user by DPAPI. Interactive ConsoleHost runs show byte progress when progress is enabled; temporary keep-awake is opt-in. The [README](../../README.md) documents supported commands, recovery, preview and privacy behavior.

The portable ZIP includes both entry points, all 27 runtime helpers, original MIT license and artwork, README, changelog and release guide. Development tools, test fixtures, personal settings, feed credentials, archive history, downloaded media and logs are excluded. The sidecar manifest identifies every payload file and its length/hash, full source commit/tree and repository. `SHA256SUMS` covers the ZIP and sidecar manifest; identical manifest bytes are inside the ZIP. See [source and checksum guidance](../../RELEASE.md).

## Compatibility and recovery

Both supported Windows engines have concrete automated and fresh-extraction evidence. Exact versions, selected counts, failures and unrun cases are in the [task evidence](../codex/evidence/UPD-0404.md) and the final draft PR. CI also records its hosted Windows Server image and installed engine versions. These observations do not certify all Windows builds, podcast providers or audio files.

Extract a candidate into a new directory, keep `src` with the launchers, and read the packaged README before choosing an archive. Keep older ZIPs/manifests/checksums separately. Application-code rollback does not restore archive or saved-configuration schemas. Preserve media, history/backups, partials and resume sidecars; test rollback with explicit synthetic copies. Completed or unknown media must never be overwritten to force a migration.

## Known limitations

- A feed may expose a partial accessible catalogue. All selects the bounded accessible catalogue, not a guarantee of every historical episode.
- Media checks recognize supported containers conservatively; they do not decode the entire audio recording. No transcoding changes the original bytes.
- Requests go to user-selected feed/media hosts, which may log them. The product adds no telemetry. Share redacted diagnostics, and avoid pasting private URLs into shell history or transcripts.
- DPAPI protects credentials at rest for the current Windows user; it does not protect against compromise of that user/admin account, live memory access, or already-recorded input/transcripts.
- Physical Ctrl+C, sleep/lid/hard termination, failed cleanup, hardware power loss and network-share durability retain the limitations described in the README. Catchable cancellation tests do not establish guarantees for those events.
- The script is unsigned. The batch launcher applies process-only execution policy to its child; it does not weaken machine/user policy. Authenticode, certificate purchase and trust-store changes require a separate owner decision.
- Checksums establish byte equality against a trusted expected value; they alone do not authenticate the publisher. Repeated ZIP bytes are checked within one compression implementation; equal bytes across engines are not promised.

The original MIT license covers the software. It does not grant rights to podcast recordings.

## Candidate identity to attach after verification

Use the exact final-head candidate report and successful workflow run, never a moving branch ZIP. Fill this review record with observed values before requesting a publication decision:

| Field | Required observed value |
| --- | --- |
| Version/status | `0.1.0-rc.1` / `UNRELEASED_CANDIDATE` |
| Source commit/tree | Full immutable IDs from the manifest, equal to the tested checkout |
| Test run | Exact successful two-engine run URL and counts |
| Artifact | Exact run ID/name; CI retention is 7 days |
| ZIP and manifest hashes | Values independently checked against `SHA256SUMS` and source |
| Intended tag/action | Owner-selected exact tag and separately authorized action |
| Final readiness | A055-A060 remain outstanding at UPD-0404; resolve or explicitly approve any deferral |

No release download link or tag is invented here. The final PR records concrete candidate identifiers without embedding a document's own commit hash into itself. A CI candidate upload is review storage, not approval to create a GitHub release.
