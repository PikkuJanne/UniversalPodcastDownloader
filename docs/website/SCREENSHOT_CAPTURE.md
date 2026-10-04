# Capture the real Windows text interface

This is a maintainer capture recipe, not an application prerequisite. It uses the checked-in loopback fixture and an already available Python 3.9+ interpreter. Python and the test helpers are development-only and are excluded from the portable downloader ZIP.

No screenshot asset has been supplied with these instructions. Until a real capture is reviewed, the product page must keep its screenshot placeholder. Use pixels from the implemented Windows console interface; do not generate an illustration, reconstruct terminal output as an image, invent a transfer, or add a completed-download banner.

## Choose the frame

The recommended frame is the guided episode-selection prompt. It shows the actual choices for the latest episode, a positive count and `all`, using a synthetic feed. The recipe runs with `-WhatIf`: it can retrieve the fixture metadata, but it requests no planned media and creates no downloader archive, history, configuration, logs, locks or diagnostic export. A preview is not completed media.

The application intentionally suppresses animated transfer progress in previews. This recipe therefore does not provide a progress-bar screenshot or a 100% completion claim. Interactive, unredirected ConsoleHost downloads can display progress automatically; `-KeepAwake` is a separate opt-in feature. See [the progress policy](../codex/PROGRESS_AND_POWER.md).

Suggested caption for a reviewed capture:

> Universal Podcast Downloader's guided episode selection on Windows, using a local synthetic feed in preview mode. Unreleased candidate; no media downloaded.

## Start an owned local fixture

Open a fresh ordinary **Windows PowerShell 5.1** console or **PowerShell 7** console with no profile. Use the chosen engine for the entire recipe and repeat in a new console for the other engine when both screenshots are wanted. An ordinary Windows Terminal tab running that engine is suitable; host appearance can vary. Do not use PowerShell ISE, redirected output, a transcript, `-NonInteractive`, or a test worker that replaces `Read-Host` for the capture itself.

Change to the actual repository checkout. Keep unrelated terminals, private feeds, personal archive folders, saved subscriptions and account details out of the frame. The following setup uses the existing test harness only to start and track a hidden fixture-server child; it does not replace the application's prompts or renderer.

```powershell
$ErrorActionPreference = 'Stop'
$captureRepo = (Get-Location).Path
$captureProgram = Join-Path $captureRepo 'UniversalPodcastDownloader.ps1'
$captureHarness = Join-Path $captureRepo 'tests/support/IntegrationHarness.ps1'
if (-not (Test-Path -LiteralPath $captureProgram) -or
    -not (Test-Path -LiteralPath $captureHarness)) {
    throw 'Run this recipe from the actual repository checkout.'
}

$captureEngine = (Get-Process -Id $PID).Path
$captureEngineVersion = $PSVersionTable.PSVersion.ToString()
$capturePython = (Get-Command python -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
& $capturePython --version
if ($LASTEXITCODE -ne 0) { throw 'Choose an already available working Python interpreter.' }

. $captureHarness
$captureContext = New-UpdIntegrationContext -RepositoryRoot $captureRepo
try {
    Start-UpdFixtureServer -Context $captureContext -PythonPath $capturePython
}
catch {
    Remove-UpdIntegrationContext -Context $captureContext
    throw
}
$captureFeed = $captureContext.BaseUrl + '/feeds/single.xml'
$captureArchive = Join-Path $captureContext.Root 'Preview archive'
Write-Host ('Synthetic capture feed: ' + $captureFeed)
Write-Host ('Capture engine: ' + $captureEngineVersion)
```

If `python` resolves to a Windows application alias instead of a working interpreter, set `$capturePython` to the full path of an existing Python executable and check `--version` before creating the context. Do not install Python to launch the downloader; this requirement belongs only to the synthetic capture fixture. [Fixture documentation](../../tools/codex-handoff/README.md) describes its routes and boundaries.

Copy the printed synthetic URL for the next prompt. It must start with `http://127.0.0.1:`. The server chooses a fresh local port; do not reuse a port from an example, expose the listener through a proxy or tunnel, or substitute a real/private feed. `/feeds/single.xml` contains one synthetic episode whose media route is `/media/ok.mp3`.

The archive path is a new, initially nonexistent child of the harness's marked temporary root. Do not replace it with your Downloads directory or real archive. Do not supply saved-show, legacy, diagnostic-export or keep-awake arguments.

## Capture the prompt and finish the preview

Save the screenshot outside `$captureContext.Root`, because cleanup removes that owned temporary tree. An appropriate initial review destination is an ignored `artifacts/tui-capture-<unique-id>/` directory in this checkout. Keep the original PNG and any cropped review copy separate; never overwrite an unknown existing file. Do not add an unreviewed image to the product page.

Paste this block in the same parent console:

```powershell
try {
    Clear-Host
    & $captureEngine -NoLogo -NoProfile -ExecutionPolicy Bypass -File $captureProgram -OutputPath $captureArchive -WhatIf
    $captureExitCode = $LASTEXITCODE
    if ($captureExitCode -ne 0) { throw ('Synthetic preview exited with ' + $captureExitCode) }

    $captureStats = Get-UpdFixtureState -Context $captureContext
    if ([int]$captureStats.'/media/ok.mp3' -ne 0) {
        throw 'A media request occurred; do not label this run as a no-media preview.'
    }
    if (Test-Path -LiteralPath $captureArchive) {
        throw 'The preview archive exists; do not claim a no-write preview.'
    }
    Write-Host 'Synthetic preview finished: exit 0, no media request, no preview archive.'
}
finally {
    Remove-UpdIntegrationContext -Context $captureContext
}
```

The nested engine is intentional: the entry script ends with an exit status, while the parent console retains ownership of the fixture and runs cleanup. The `Bypass` execution-policy argument applies only to this child process; no machine-wide or user-wide policy is changed. The child has no `-NonInteractive` option and shares the visible console without stdout/stderr redirection.

While that block runs:

1. At `Paste RSS feed URL OR podcast page URL`, paste the copied loopback URL and press Enter.
2. Wait for `How many episodes to download (newest first)?` and the three displayed selection choices. Stop at `Episodes to download (number / 'all' / Enter = 1)` before answering.
3. Use Windows Snipping Tool or another local screenshot tool to capture the actual console region. Keep the program heading, selection choices and active prompt legible. Exclude unrelated windows, taskbar notifications, user names, repository/private paths and previous terminal history. Cropping can remove surrounding desktop chrome; it must preserve the application's displayed wording.
4. Save the PNG to the review destination outside the temporary root, then return to the console and type `all` followed by Enter. The checked-in feed has one accessible entry. The real application should display its WhatIf action rather than a downloaded-media summary or `[OK] Done` banner.
5. Let the block finish and retain its explicit preview checks. Do not close or force-kill the parent while it owns the fixture. If a setup, preview or cleanup check fails, preserve the observation and do not treat it as a successful capture procedure.

This recipe creates harness readiness files and temporary directories as development setup. Its no-write statement concerns the downloader preview, not those fixture files or Windows host activity. Console startup can create ordinary host cache files; it is not a claim that the whole Windows host performed zero writes. Cleanup stops only the fixture process tracked by this context and removes only its verified marked temporary root. It cannot guarantee cleanup after forced termination.

## Review and record provenance

Before including a capture in distribution content, record these facts next to the reviewed asset:

- Full source commit or the exact candidate `manifest.json` source commit/tree; note any uncommitted runtime changes. A screenshot of an older source is not evidence for a newer candidate.
- Candidate status/version, Windows version, actual engine version and console host; do not describe an unreleased candidate as stable.
- Capture date, image filename, dimensions and whether the review copy was cropped.
- Synthetic input `/feeds/single.xml`, `-WhatIf`, explicit owned output path, exit status, zero `/media/ok.mp3` requests and absent preview archive.
- Whether the image shows episode selection or final WhatIf output. It proves that observed screen, not a download, progress timing, accessibility in every terminal, keep-awake behavior or release readiness.

Suggested filenames after review are `tui-selection-ps51.png` and `tui-selection-ps7.png` under `docs/website/assets/`. Use descriptive alt text such as “Windows PowerShell 5.1 console showing the guided episode-selection choices with a local synthetic feed in preview mode.” Check that the PNG is an actual capture, contains no personal data and matches its caption. Original project artwork may accompany it, but is not a substitute for the real interface.

When a stable release is later approved, capture or recheck the exact approved source and update the caption/provenance accordingly. Do not fabricate a download URL or publish an image/page as part of this recipe. Release links and publication remain subject to the [release approval checklist](../releases/APPROVAL_CHECKLIST.md).
