param([Parameter(Mandatory)][string]$ConfigPath)

# Literal documentation is trusted repository source, copied into this worker's
# config. Runtime requests, credentials, output and every child remain owned.
$ErrorActionPreference = 'Stop'
$documentationConfig = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
$ownedRoot = [IO.Path]::GetFullPath($documentationConfig.Root).TrimEnd('\', '/')
$marker = Join-Path $ownedRoot '.upd-test-owner'
if (-not [IO.File]::Exists($marker) -or [IO.File]::ReadAllText($marker) -cne $documentationConfig.Token) {
    throw 'Documentation worker requires its matching owned root.'
}
foreach ($path in @($ConfigPath, $documentationConfig.ProductScript, $documentationConfig.ResultPath, $documentationConfig.Sample)) {
    if (-not [IO.Path]::GetFullPath($path).StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Documentation paths must remain inside the owned root.'
    }
}
$source = [uri]$documentationConfig.BaseUrl
if ($source.Scheme -ne 'http' -or $source.Host -ne '127.0.0.1') { throw 'Documentation examples require an owned loopback source.' }
$env:LOCALAPPDATA = Join-Path $ownedRoot 'local'
$env:USERPROFILE = $ownedRoot
$env:TEMP = $ownedRoot
$env:TMP = $env:TEMP
$env:PATH = $PSHOME + ';' + (Join-Path $env:WINDIR 'System32') + ';' + $env:WINDIR
$env:PSModulePath = Join-Path $PSHOME 'Modules'
$packageRoot = [IO.Path]::GetDirectoryName($documentationConfig.ProductScript)
Set-Location -LiteralPath $packageRoot
$report = [ordered]@{ Engine = $PSVersionTable.PSVersion.ToString(); Blocks = @(); Help = $null; Error = $null;
    PackageChanged = $false; RuntimePythonAvailable = $false; DefaultConfigCreated = $false;
    PreviewChangedTree = $false; OriginalsPreserved = $false; ObservedVerifiedBytes = $false;
    AdoptionIsUnverified = $false; CallableResult = $null; PassThruResult = $null; RecoveryOutcomes = @();
    RenderedCodeMatchesAuthored = $false;
    ChildExitCode = $null; ChildStdout = $null; ChildStderr = $null; GuidedPromptsObserved = $false }
$workerExit = 1

function Get-UpdDocumentationTree {
    param([string]$Root)
    if (-not [IO.Directory]::Exists($Root)) { return @() }
    return @(Get-ChildItem -LiteralPath $Root -Recurse -Force | Sort-Object FullName | ForEach-Object {
        if ($_.PSIsContainer) { 'directory|' + $_.FullName }
        else { $_.FullName + '|' + $_.Length + '|' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
    })
}

function Invoke-UpdDocumentationChild {
    param([string]$Executable, [string[]]$Arguments, [string]$RawArguments, [string]$InputText)
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $Executable
    $start.Arguments = if ($RawArguments) { $RawArguments } else { ($Arguments | ForEach-Object { '"' + $_ + '"' }) -join ' ' }
    $start.WorkingDirectory = $packageRoot
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    if ($Executable.EndsWith('cmd.exe', [StringComparison]::OrdinalIgnoreCase)) {
        $start.EnvironmentVariables['PSModulePath'] = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules'
    }
    $child = [Diagnostics.Process]::new()
    $child.StartInfo = $start
    try {
        if (-not $child.Start()) { throw 'Could not start the owned documentation child.' }
        $stdout = $child.StandardOutput.ReadToEndAsync()
        $stderr = $child.StandardError.ReadToEndAsync()
        if ($InputText) { $child.StandardInput.Write($InputText) }
        $child.StandardInput.Close()
        if (-not $child.WaitForExit(30000)) { throw 'Owned documentation child exceeded its bounded wait.' }
        return [pscustomobject]@{ ExitCode = $child.ExitCode; Stdout = $stdout.Result; Stderr = $stderr.Result }
    }
    finally {
        if ($child.Id -and -not $child.HasExited) { $child.Kill(); $child.WaitForExit() }
        $child.Dispose()
    }
}

try {
    $packageBefore = Get-UpdDocumentationTree -Root $packageRoot
    $report.RuntimePythonAvailable = $null -ne (Get-Command python -CommandType Application -ErrorAction SilentlyContinue)
    . $documentationConfig.ProductScript
    $exampleRoot = Join-Path $env:TEMP ('UPD example ' + [char]0x2013 + ' ' + [char]0x00e4 + [char]0x00e4)
    $legacyCopy = Join-Path (Join-Path $exampleRoot 'Legacy archive') 'Copied historical show'
    $originals = @()
    if ($documentationConfig.Action -eq 'Readme') {
        $sourceCopy = Join-Path $ownedRoot 'synthetic-source'
        $null = [IO.Directory]::CreateDirectory($sourceCopy)
        $null = [IO.Directory]::CreateDirectory($legacyCopy)
        [IO.File]::Copy($documentationConfig.Sample, (Join-Path $sourceCopy '2026-09-01 - Original episode.mp3'))
        [IO.File]::WriteAllBytes((Join-Path $sourceCopy 'unknown.part'), [byte[]]@(73, 68, 51, 1, 2))
        [IO.File]::WriteAllText((Join-Path $sourceCopy 'owner-notes.txt'), 'Synthetic original notes must survive every recovery command.')
        foreach ($item in Get-ChildItem -LiteralPath $sourceCopy -File) {
            [IO.File]::Copy($item.FullName, (Join-Path $legacyCopy $item.Name))
        }
        $originals = @(Get-ChildItem -LiteralPath $legacyCopy -File | ForEach-Object {
            [pscustomobject]@{ Path = $_.FullName; Bytes = $_.Length; Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
        })
    }
    $promptState = @{ Count = 0 }
    function Read-Host {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'This owned documentation worker supplies only the expected synthetic URL and reviewed copied-file identifiers; it rejects every other prompt and never supplies confirmation consent.')]
        [CmdletBinding()]
        param([string]$Prompt)
        $promptState.Count++
        if ($Prompt -match 'Podcast RSS|show-page URL') { return $documentationConfig.BaseUrl + '/feeds/history.xml' }
        if ($Prompt -match 'folder|directory') { return 'Copied historical show' }
        if ($Prompt -match 'EpisodeId') {
            $feedId = Get-PodcastNameHash -IdentityKey ('feed:' + $documentationConfig.BaseUrl + '/feeds/history.xml')
            $episode = [pscustomobject]@{ Title = 'Original episode'; Guid = 'history-stable-001'; AtomId = $null;
                Url = $documentationConfig.BaseUrl + '/media/history.mp3?signature=original&part=1'; PubDate = [datetime]'2026-09-01T12:00:00Z' }
            return (Get-PodcastEpisodeIdentity -Episode $episode -FeedId $feedId).Id
        }
        if ($Prompt -match 'RelativePath|LegacyFile') { return '2026-09-01 - Original episode.mp3' }
        if ($Prompt -match 'Sha256|digest') { return (Get-FileHash -LiteralPath (Join-Path $legacyCopy '2026-09-01 - Original episode.mp3') -Algorithm SHA256).Hash.ToLowerInvariant() }
        if ($Prompt -match 'checkpoint|basename') {
            return (Get-ChildItem -LiteralPath (Join-Path $legacyCopy '.upd') -File -Filter 'legacy-*.json' | Sort-Object Name | Select-Object -First 1).Name
        }
        throw ('Unexpected documentation input prompt: ' + $Prompt)
    }

    if ($documentationConfig.Action -eq 'Help') {
        $name = if ($documentationConfig.CommandName -eq 'UniversalPodcastDownloader.ps1') { '.\UniversalPodcastDownloader.ps1' } else { $documentationConfig.CommandName }
        $help = Get-Help $name -Full
        $command = Get-Command $name
        $common = @('Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction', 'ProgressAction',
            'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable')
        $missing = @($command.Parameters.Keys | Where-Object { $_ -notin $common } | Where-Object {
            $parameter = @($help.Parameters.Parameter | Where-Object Name -eq $_)
            $parameter.Count -ne 1 -or -not ($parameter[0].Description.Text -join ' ').Trim()
        })
        $examples = @($help.Examples.Example | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Code) })
        $rendered = $help | Out-String -Width 160
        $report.Help = @{ Authored = (-not [string]::IsNullOrWhiteSpace(($help.Description.Text -join ' ')) -and
            $help.Synopsis -notmatch '\[\[-Mode\]' -and $rendered -match 'SYNOPSIS' -and $rendered -match 'PARAMETERS' -and $rendered -match 'EXAMPLE');
            MissingParameters = $missing; ExampleCount = $examples.Count; Rendered = $rendered }
    }
    elseif ($documentationConfig.Action -in @('Readme', 'HelpExamples')) {
        $blocks = @()
        if ($documentationConfig.Action -eq 'Readme') {
            $pattern = '(?s)<!-- UPD-0402 example:(?<id>[a-z-]+) -->\s*```powershell\s*\r?\n(?<code>.*?)\r?\n```'
            $blocks = @([regex]::Matches($documentationConfig.Readme, $pattern) | ForEach-Object {
                [pscustomobject]@{ Id = $_.Groups['id'].Value; Code = $_.Groups['code'].Value }
            })
            if ($blocks.Count -ne [regex]::Matches($documentationConfig.Readme, '```powershell').Count) { throw 'Every README PowerShell example must have a verification ID.' }
        }
        else {
            foreach ($name in @('.\UniversalPodcastDownloader.ps1', 'Invoke-PodcastRun')) {
                $index = 0
                $authoredExamples = @((Get-Command $name).ScriptBlock.Ast.GetHelpContent().Examples)
                foreach ($example in @((Get-Help $name -Full).Examples.Example | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Code) })) {
                    $index++
                    $code = [string]$example.Code
                    # Windows PowerShell 5.1 renders only the first command in
                    # Code and keeps the remaining contiguous command lines in
                    # Remarks, before the blank line introducing explanation.
                    if ($code -notmatch '\r?\n') {
                        $continuation = (($example.Remarks.Text -join "`n") -split '\r?\n\s*\r?\n', 2)[0]
                        if (-not [string]::IsNullOrWhiteSpace($continuation)) { $code += "`n" + $continuation }
                    }
                    $authoredCode = ([string]$authoredExamples[$index - 1] -split '\r?\n\s*\r?\n', 2)[0]
                    $normalizedCode = $code.Replace("`r`n", "`n").Replace("`r", "`n").Trim()
                    $normalizedAuthored = $authoredCode.Replace("`r`n", "`n").Replace("`r", "`n").Trim()
                    if (-not [StringComparer]::Ordinal.Equals($normalizedCode, $normalizedAuthored)) {
                        throw ('Rendered help commands differ from the authored contiguous example: ' + $name + ':' + $index)
                    }
                    $blocks += [pscustomobject]@{ Id = $name + ':' + $index; Code = $code }
                }
                if ($index -ne $authoredExamples.Count) { throw 'Rendered help omitted an authored example.' }
            }
            $report.RenderedCodeMatchesAuthored = $true
        }
        foreach ($block in $blocks) {
            $before = Get-UpdDocumentationTree -Root $exampleRoot
            $modulePathBefore = $env:PSModulePath
            if ($block.Id -eq 'launcher') { $env:PSModulePath = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules' }
            $LASTEXITCODE = 0
            try { $output = @(. ([scriptblock]::Create($block.Code))) }
            finally { $env:PSModulePath = $modulePathBefore }
            $blockExit = $LASTEXITCODE
            $runResults = @($output | Where-Object { $_.PSObject.Properties['Type'] -and $_.Type -eq 'Podcast.RunResult' })
            if ($runResults.Count -gt 0) { $report.PassThruResult = $runResults[-1] }
            if ($block.Id -eq 'api') { $report.CallableResult = $run }
            if ($block.Id -like 'Invoke-PodcastRun:*') { $report.CallableResult = $result }
            if ($block.Id -in @('preview', 'legacy-preview')) {
                $report.PreviewChangedTree = $report.PreviewChangedTree -or @(Compare-Object $before (Get-UpdDocumentationTree -Root $exampleRoot)).Count -gt 0
            }
            if ($block.Id -eq 'legacy-adopt') {
                $state = [IO.File]::ReadAllText((Join-Path $legacyCopy '.upd/state.json')) | ConvertFrom-Json
                $report.AdoptionIsUnverified = @($state.episodes | Where-Object { $_.status -eq 'adopted' -and
                    $_.verification.method -eq 'owner-approved-local-signature' -and
                    $_.verification.notes -contains 'local-signature-only; transfer-completeness-unverified' }).Count -eq 1
                $report.RecoveryOutcomes += $adoption.LegacyResult.Outcome
            }
            if ($block.Id -eq 'legacy-redownload') { $report.RecoveryOutcomes += $replacement.LegacyResult.Outcome }
            if ($block.Id -eq 'legacy-rollback') { $report.RecoveryOutcomes += $rollback.LegacyResult.Outcome }
            $hasher = [Security.Cryptography.SHA256]::Create()
            try {
                $hashCode = $block.Code.Replace("`r`n", "`n").Replace("`r", "`n").Trim()
                $codeHash = [BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($hashCode))).Replace('-', '').ToLowerInvariant()
            }
            finally { $hasher.Dispose() }
            $report.Blocks += [pscustomobject]@{ Id = $block.Id; CodeSha256 = $codeHash; Succeeded = $blockExit -eq 0; ExitCode = $blockExit }
            if ($blockExit -ne 0) { throw ('Literal documentation example failed: ' + $block.Id + ' exit=' + $blockExit) }
        }
        $report.OriginalsPreserved = $originals.Count -gt 0
        foreach ($original in $originals) {
            if (-not [IO.File]::Exists($original.Path) -or (Get-Item -LiteralPath $original.Path).Length -ne $original.Bytes -or
                (Get-FileHash -LiteralPath $original.Path -Algorithm SHA256).Hash -cne $original.Hash) { $report.OriginalsPreserved = $false }
        }
        $sampleHash = (Get-FileHash -LiteralPath $documentationConfig.Sample -Algorithm SHA256).Hash
        $verified = @(Get-ChildItem -LiteralPath $exampleRoot -Recurse -File -Filter 'state.json' | ForEach-Object {
            $state = [IO.File]::ReadAllText($_.FullName) | ConvertFrom-Json
            $archiveRoot = Split-Path (Split-Path $_.FullName -Parent) -Parent
            foreach ($record in $state.episodes) {
                if ($record.status -eq 'transfer_verified' -and $record.verification.method -in @('completed-http-length-and-signature', 'completed-eof-and-signature')) {
                    $mediaPath = Join-Path $archiveRoot $record.relative_path
                    if ((Get-FileHash -LiteralPath $mediaPath -Algorithm SHA256).Hash -eq $sampleHash -and
                        (Get-Item -LiteralPath $mediaPath).Length -eq $record.bytes -and
                        $record.local_sha256 -eq $sampleHash.ToLowerInvariant()) { $true }
                }
            }
        })
        $report.ObservedVerifiedBytes = $verified.Count -gt 0
    }
    else {
        if ($documentationConfig.Action -eq 'Guided') {
            $child = Invoke-UpdDocumentationChild -Executable (Join-Path $env:WINDIR 'System32/cmd.exe') -RawArguments '/d /s /v:off /c "".\UniversalPodcastDownloader.bat""' -InputText ($documentationConfig.BaseUrl + "/feeds/history.xml`r`n1`r`n`r`n")
            $report.ChildStdout = $child.Stdout
            $report.ChildStderr = $child.Stderr
            if ($child.ExitCode -ne 0) { throw ('Guided child failed: ' + $child.Stdout + $child.Stderr) }
            # Native redirected Read-Host omits its prompt label. Observe the
            # actual guided instructions and completed selection instead.
            $report.GuidedPromptsObserved = $child.Stdout -match 'Paste a direct RSS/Atom feed URL' -and
                $child.Stdout -match 'How many episodes to download' -and $child.Stdout -match 'Downloaded\s+: 1'
            if (-not [IO.Directory]::Exists((Join-Path $env:USERPROFILE 'Downloads/Podcasts'))) {
                throw ('Guided child created no synthetic output: ' + $child.Stdout + $child.Stderr)
            }
            $files = @(Get-ChildItem -LiteralPath (Join-Path $env:USERPROFILE 'Downloads/Podcasts') -Recurse -File -Filter '*.mp3')
            $report.ObservedVerifiedBytes = $files.Count -eq 1 -and
                (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $documentationConfig.Sample -Algorithm SHA256).Hash
        }
        else {
            $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
            $feedPath = if ($documentationConfig.ExpectedExit -eq 2) { '/feeds/transaction-html.xml' } else { '/feeds/history.xml' }
            $arguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $documentationConfig.ProductScript,
                '-FeedUrl', ($documentationConfig.BaseUrl + $feedPath), '-OutputPath', (Join-Path $ownedRoot 'process archive'), '-NonInteractive')
            if ($documentationConfig.ExpectedExit -eq 1) { $arguments += @('-Mode', 'All', '-CustomCount', '1') }
            $child = Invoke-UpdDocumentationChild -Executable (Join-Path $PSHOME $engine) -Arguments $arguments
        }
        $report.ChildExitCode = $child.ExitCode
        if ($child.ExitCode -ne $documentationConfig.ExpectedExit) { throw ('Entry child exit differs: ' + $child.ExitCode + '; ' + $child.Stdout + $child.Stderr) }
    }
    $report.PackageChanged = @(Compare-Object $packageBefore (Get-UpdDocumentationTree -Root $packageRoot)).Count -gt 0
    $report.DefaultConfigCreated = [IO.Directory]::Exists((Join-Path $env:LOCALAPPDATA 'UniversalPodcastDownloader/saved-shows'))
    $workerExit = 0
}
catch { $report.Error = $_.Exception.Message }
[IO.File]::WriteAllText($documentationConfig.ResultPath, ($report | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false))
exit $workerExit
