# Bundle validation — 1 October 2026

## Scope

This report is about the **handoff package and synthetic helper tooling only**. It is not a Windows compatibility report or a claim that the downloader improvements have been implemented. No GitHub write, original downloader change, real podcast download, archive migration, release or website deployment was performed while preparing this bundle.

## Checks actually performed

| Check | Result |
|---|---|
| Current main/tree inspected through connected GitHub | Baseline `2ac82614493be7196c9ebee116f23fec07368b50` confirmed |
| Task and acceptance registers | 26 unique tasks, 60 unique acceptance cases; dependencies and cross-references validated |
| Original improvement coverage | All 23 items map to implementation tasks, not just the final audit |
| Python helper syntax and JSON parsing | Passed |
| Loopback helper self-tests | **20 passed, 0 failed** on Linux / Python 3.13.5 |
| Synthetic MP3 | ffprobe recognizes MP3; generated silence, approximately 0.366 seconds |
| Repository overlay safety | Does not include replacement original root script, batch, README, license or artwork |
| SHA-256 manifest and file inventory | Checked with bundled read-only verifier |
| Verifier negative controls | Altered file, extra file and unsafe manifest path must be rejected; checked during packaging |
| ZIP round trip | Extracted archive revalidated against manifest and checksum inventory |

The helper tests include valid/truncated media, correct and ignored ranges, changed/weak/missing validators, inconsistent range metadata, retry-once/status handling, no-length media, HTML masquerading as audio, safe ready-file handling, loopback-only binding, no arbitrary path serving and query-free request statistics. Raw test names/outcome are preserved in `tools/codex-handoff/SELFTEST_RESULTS.txt` in the repository overlay.

## Explicitly not executed

- The improved downloader: it does not exist yet in this handoff.
- Windows PowerShell 5.1 or PowerShell 7 application tests.
- The optional `Apply-Handoff.ps1` helper under an actual PowerShell runtime. Its source was reviewed; Codex must validate it locally before application.
- Launcher/Windows filename/junction/cancellation/power behavior or actual private/legacy feeds.
- GitHub feature-branch pushes, CI runs, PR creation, tags, signatures, releases or deployment.

The 60 downloader acceptance cases begin `not_run`. Helper self-tests must never be used as evidence that those product cases pass. UPD-0001 checks the Windows handoff application; UPD-0002 starts product test foundations.

## Recheck the extracted package

Run `python -B verify_bundle.py` from the external bundle folder. The optional PowerShell application helper also verifies declared bundle hashes before any copy. Keep the bundle unchanged; run development tests with `python -B` so Python cache files are not added to the immutable bundle. Integrity hashes detect changed bytes but are not an independent publisher signature.
