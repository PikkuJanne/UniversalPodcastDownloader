# Changelog

## 1.0.0 — first stable release

The completed downloader is distributed as v1.0.0 at the owner's request on 8 October 2026. The 29 runtime files are unchanged from the tested, merged development closure; this version updates distribution metadata, user guidance and stable-version packaging validation.

- Provide the portable Windows RSS/Atom downloader with Latest/Custom/All modes, configurable output, original audio bytes, verified history and conservative recovery.
- Include bounded feed discovery/catalogue traversal, safe episode naming, strict owned transfer resume, private run results, optional saved-show batches, interactive byte progress and opt-in keep-awake.
- Ship the original MIT license and artwork with the complete runtime, README, release guide, source/file manifest and SHA-256 checksums.
- Support Windows PowerShell 5.1 and PowerShell 7 on Windows. No application runtime package manager, Python, Git or FFmpeg is required.

Known limits remain documented in [README](README.md) and [RELEASE](RELEASE.md). Human physical Ctrl+C and native cleanup were observed in both Windows Sandbox engines; host sleep/lid, hard-kill/power-loss and network-share durability are not guaranteed. Uncheckpointed partial tails require review. Catalogue traversal is bounded and media recognition does not decode entire recordings.

## 0.1.0-rc.1 — unpublished candidate

No application tags or GitHub releases existed when this candidate version was selected on 3 October 2026. This is the first versioned packaging candidate, not a stable release announcement. The planning bundle's `1.0.0` is unrelated to the application version. Release workflow and local draft material are prepared; final acceptance and publication approval remain outstanding.

- Preserve the original PowerShell/batch entry points, Latest/Custom/All selection, configurable output and original audio bytes.
- Add bounded RSS/Atom discovery and catalogue traversal, deterministic identities and Windows-safe paths, verified history, conservative legacy review and strict owned transfer resume.
- Add private run results and exit codes, sequential saved-show batches with CurrentUser DPAPI protection, explicit transport limits, interactive byte progress and opt-in temporary keep-awake.
- Provide verified standard PowerShell help and README examples, including recovery and privacy limitations.
- Package the complete runtime, original MIT license, artwork and user guidance from one committed source. Include a source/file manifest and SHA-256 checksums; exclude personal data and development material.
- Prepare read-only candidate CI with reviewed SHA-pinned actions, two-engine gates and a narrow artifact inventory. Keep draft notes and approval records local to the source; create no release tags or public releases automatically.

Compatibility observations: Windows 11 Pro build `10.0.26300.0`, Windows PowerShell `5.1.26100.9444` and PowerShell `7.6.5`. CI separately observes Windows Server 2022 and its recorded engine versions. These observations do not certify every Windows version, podcast provider or audio file.

Known limits include bounded accessible catalogues, container recognition rather than full audio decoding, same-user/admin access to credentials or memory, input history/transcripts, and unproven physical Ctrl+C/sleep/lid/hard-kill or hardware power-loss/network-share durability. Preserve archives, state and unknown partials. See [README](README.md) and [package/source/rollback guidance](RELEASE.md).
