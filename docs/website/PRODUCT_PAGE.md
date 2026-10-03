# UniversalPodcastDownloader

A local Windows RSS/Atom podcast downloader that keeps original audio and records evidence for repeat downloads. Use the guided text interface, a one-off command or optional saved shows.

**Current status: `0.1.0-rc.1` — `UNRELEASED_CANDIDATE`.** No public release, stable version or published download is announced here. Download, checksum, manifest and release-note asset links remain unassigned. This is portable draft content for a future distribution page; the source contains the features described below, while final candidate acceptance and publication approval remain outstanding.

## Windows requirements

Use Windows 11 with Windows PowerShell 5.1 or a supported PowerShell 7 release on Windows, a writable archive location and access to your chosen feed/media hosts. Local checks exercised Windows 11 Pro build `10.0.26300.0`, Windows PowerShell `5.1.26100.9444` and PowerShell `7.6.5`; this does not certify other operating systems or every provider/file. [Usage and compatibility details](../../README.md) record the evidence and limits.

The portable application needs no Python, Git, Pester, package manager, FFmpeg, cloud account, GUI or background service. It normally needs no administrator rights. The `.bat` launcher uses built-in Windows PowerShell; run the `.ps1` from PowerShell 7 to use that engine. The script is unsigned. The launcher applies `-ExecutionPolicy Bypass` only to its child process, without changing user/machine policy. Managed Windows policy can still prevent execution; use your organization's approved workflow.

## Get started with a reviewed candidate

1. Obtain a specifically reviewed candidate ZIP, its manifest and `SHA256SUMS`. Public download links are not available on this draft page. Compare the ZIP and manifest SHA-256 values before extraction, following the [package/source guide](../../RELEASE.md). A checksum detects different bytes; it does not authenticate the publisher.
2. Extract into a new ordinary directory. Keep `UniversalPodcastDownloader.ps1`, `UniversalPodcastDownloader.bat` and the complete `src` directory together. Read the included README and view built-in help from that directory:

   ```powershell
   Get-Help .\UniversalPodcastDownloader.ps1 -Full
   ```

3. Double-click `UniversalPodcastDownloader.bat`, enter your selected RSS/Atom or supported show-page URL, and choose Enter, a positive count or `all`. Downloads default to `%USERPROFILE%\Downloads\Podcasts`; use `-OutputPath` for another archive root. The [README examples](../../README.md#quick-start) show command-line selection and `-WhatIf` preview before downloading.

## Choose what to keep

| Mode | Selection from the accessible catalogue |
| --- | --- |
| Latest | One eligible episode, ordered newest first |
| Custom | Up to a positive episode count |
| All | All eligible entries collected within the feed limits |

Episodes sort by UTC publication time; ties keep feed order and undated entries come last. Supported show-page RSS/Atom links can be discovered; several discovered feeds require a choice, or a direct feed URL for noninteractive use. This is not browser-session import or a general website crawler.

All modes retrieve supported feed-level pagination before selection. The default is 20 pages, adjustable from 1 to 100, with additional entry/metadata bounds. Missing pages, cycles, ambiguity and limits remain incomplete. Even a completed supported chain does not prove that the publisher's complete historical archive is accessible.

Recognized audio includes MP3, M4A, Ogg Opus/Vorbis/Speex, WAVE and FLAC. Files keep the received bytes without transcoding or retagging. Checks use bounded container/signature evidence and HTTP completion; they do not decode tracks, prove playability or authenticate an episode. Unsupported or ambiguous responses can be refused.

`-WhatIf` can fetch bounded metadata and read/hash local evidence, but requests no planned enclosure and creates no downloader media, directories, state, settings, locks, logs or exports. Preview is not a completed download. A metadata URL can itself return unexpected bytes through the bounded metadata path.

## Keep the archive and its recovery evidence together

Configurable output roots contain show folders and filenames with identity suffixes. Existing media is not overwritten. A verified repeat skip requires recorded size/SHA-256 to match the actual file; unknown or changed files remain conflicts for review. Owned partials resume only when local bytes and server representation evidence match. Some servers or interruptions require a fresh separate transfer or manual review.

Preserve media, `.upd` history/backups/checkpoints and partial/sidecar evidence together. Do not delete state or overwrite a live backup as a repair shortcut. Review legacy archives on explicit copies first; adoption remains transfer-unverified, and rollback changes reviewed metadata rather than restoring or deleting media. See [recovery and legacy guidance](../../README.md#recovery-and-troubleshooting).

Optional saved shows protect feed URLs with Windows DPAPI CurrentUser and private configuration permissions; one-off/guided runs do not load saved settings. Saved batches run sequentially. Interactive ConsoleHost progress appears when the host and progress preference permit it; noninteractive/redirected runs keep readable notices. `-KeepAwake` is opt-in, makes a temporary system-sleep request and does not keep the display on or change a power plan. Explicit sleep/lid actions, forced termination, cleanup failures and hardware power loss remain limits. Local interruption checks do not establish UNC/share durability.

## Privacy and support

Audio, archive state, optional configuration and diagnostics stay local. There is no product telemetry, automatic upload or update checker. The selected feed/media hosts receive network requests and may log them. Private feed URLs can contain credentials: keep them out of public reports, screenshots, shell history and transcripts.

Ordinary writable runs create local startup/show logs. Routine output uses safe messages and hostname/opaque request IDs; hostnames and fingerprints may still disclose subscriptions. Inspect logs before sharing, especially older logs. Optional diagnostic export contains restricted event metadata rather than URLs, messages or archive inventory; it is not uploaded automatically. Logs/exports have no automatic rotation or expiry. DPAPI/permissions cannot protect against same-user malware or an administrator, and private data can remain in process memory or caller-inspected results. [Full privacy guidance](../../README.md#diagnostics-and-privacy) describes these boundaries.

For help, start with the included [README](../../README.md) and [release/source guide](../../RELEASE.md). Report reproducible problems through [GitHub Issues](https://github.com/PikkuJanne/UniversalPodcastDownloader/issues). Include the exact candidate version/source commit, Windows/PowerShell version, command mode, exit code and a synthetic reproduction. Review any diagnostic export before attaching it; do not attach private URLs, credentials, configuration, archive history, media or unreviewed logs. Existing verified files/state should be preserved while investigating.

No scheduler, automatic archive cleanup, universal provider/codec support or guarantee of catchable Ctrl+C/forced-stop recovery is implemented. See the [changelog](../../CHANGELOG.md) for source changes. Final clean-clone acceptance and owner publication review have not been completed by preparing this page.

## License and real UI

The software is distributed under the original [MIT license](../../LICENSE), with the original supplied artwork retained. That license does not license podcast recordings or grant permission to redistribute them; respect the publisher's terms and your rights to the audio.

No screenshot asset is embedded here. [Screenshot capture instructions](SCREENSHOT_CAPTURE.md) explain how to capture the actual local text interface with synthetic input and record its source/engine. Use an inspected real capture when available; do not substitute generated imagery or describe a sample as a published release.
