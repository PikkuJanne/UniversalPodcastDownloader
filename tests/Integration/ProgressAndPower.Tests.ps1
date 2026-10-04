BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned progress and power loopback fixtures.' }
    $script:progressPython = $python.Source
    $script:progressAudio = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
    $script:progressLength = (Get-Item -LiteralPath $script:progressAudio).Length
    $script:progressHash = (Get-FileHash -LiteralPath $script:progressAudio -Algorithm SHA256).Hash
    $audio = [IO.File]::ReadAllBytes($script:progressAudio)
    $repeated = New-Object byte[] ($audio.Length * 64)
    for ($offset = 0; $offset -lt $repeated.Length; $offset += $audio.Length) { [Array]::Copy($audio, 0, $repeated, $offset, $audio.Length) }
    $hash = [Security.Cryptography.SHA256]::Create()
    try { $script:progressResumeHash = ([BitConverter]::ToString($hash.ComputeHash($repeated))).Replace('-', '').ToLowerInvariant() }
    finally { $hash.Dispose() }

    function Invoke-UpdProgressWorker {
        param([Parameter(Mandatory)]$Context, [string]$FeedPath = '/feeds/single.xml', [switch]$ForceInteractive,
            [switch]$KeepAwake, [switch]$Preview, [switch]$FailHistoryCompletion, [long]$CancelAfterBytes = 0, [int]$MaxAttempts = 1)
        if ($FeedPath -notmatch '^/(?:feeds/[a-z0-9-]+\.xml|(?:resume|transport)/feed/[a-z0-9-]+)$') {
            throw 'Progress workers require a named owned loopback feed.'
        }
        $identifier = [guid]::NewGuid().ToString('N')
        $configPath = Join-Path $Context.Root ($identifier + '-progress-config.json')
        $config = @{ Root = $Context.Root; Token = $Context.Token; ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1';
            FeedUrl = $Context.BaseUrl + $FeedPath; OutputPath = Join-Path $Context.Root 'output'; ResultPath = Join-Path $Context.Root ($identifier + '-progress-result.json');
            ForceInteractive = [bool]$ForceInteractive; KeepAwake = [bool]$KeepAwake; Preview = [bool]$Preview;
            FailHistoryCompletion = [bool]$FailHistoryCompletion; CancelAfterBytes = $CancelAfterBytes; MaxAttempts = $MaxAttempts }
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $owned = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engine) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
            (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-ProgressAndPowerWorker.ps1'), '-ConfigPath', $configPath)
        if (-not $owned.Process.WaitForExit(40000)) { throw 'Owned progress worker completion timed out.' }
        if (-not [IO.File]::Exists($config.ResultPath)) { throw ('Owned progress worker produced no result: ' + $owned.ErrorOutput.Result) }
        [pscustomobject]@{ ExitCode = $owned.Process.ExitCode; Result = [IO.File]::ReadAllText($config.ResultPath) | ConvertFrom-Json;
            Stdout = $owned.Output.Result; Stderr = $owned.ErrorOutput.Result; OutputPath = $config.OutputPath }
    }

    function Assert-UpdProgressOrdering {
        param([Parameter(Mandatory)]$Run)
        $Run.Result.Error | Should -BeNullOrEmpty -Because ($Run.Stdout + $Run.Stderr)
        $Run.Result.HostSurvived | Should -BeTrue
        @($Run.Result.Renders).Count | Should -BeGreaterThan 0
        foreach ($render in @($Run.Result.Renders | Where-Object { -not $_.Completed -and $_.Percent -eq 100 })) {
            $render.Snapshot.AcceptedHistory | Should -BeGreaterThan 0
            if ($render.Role -eq 'root') {
                $render.Snapshot.RunVerified | Should -BeTrue
                $render.Snapshot.AcceptedHistory | Should -Be $render.Snapshot.TotalEpisodes
            }
            else { $render.Snapshot.EpisodeOutcome | Should -Be 'downloaded' }
        }
        $safeText = $Run.Result.Renders | Select-Object Activity, Status, CurrentOperation | ConvertTo-Json -Depth 3
        $safeText | Should -Not -Match 'https?://|127\.0\.0\.1|[A-Za-z]:\\|Original synthetic audio|Original resume audio|Resume fixture|Transport fixture'
        @($Run.Result.Events | Where-Object Observation -eq 'returned')[0].Closed | Should -BeTrue
    }

    function Assert-UpdPowerRestore {
        param([Parameter(Mandatory)]$Run)
        $requested = @($Run.Result.Power | Where-Object Requested)
        $requested.Count | Should -Be 1
        $requested[0].ActiveBeforeStop | Should -BeTrue
        $requested[0].Restored | Should -BeTrue
        $requested[0].NativeRestored | Should -BeTrue
        $requested[0].NativeActive | Should -BeFalse
        $requested[0].ActivationThreadId | Should -BeGreaterThan 0
        $requested[0].RestorationThreadId | Should -Be $requested[0].ActivationThreadId
        $requested[0].PreviousState | Should -BeGreaterThan 0
    }

    function Get-UpdProgressArchive {
        param([Parameter(Mandatory)]$Run)
        $paths = @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -Force -File -Filter 'state.json')
        $paths.Count | Should -Be 1
        [pscustomobject]@{ Root = Split-Path $paths[0].DirectoryName -Parent;
            State = [IO.File]::ReadAllText($paths[0].FullName) | ConvertFrom-Json }
    }
}

Describe 'A044/A045 progress and temporary power through actual owned processes' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:progressPython
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'shows known real bytes and 100 only after validation and recorded success, then restores actual native keep-awake' {
        $run = Invoke-UpdProgressWorker -Context $context -FeedPath '/transport/feed/progress' -ForceInteractive -KeepAwake
        $run.ExitCode | Should -Be 0
        Assert-UpdProgressOrdering -Run $run
        Assert-UpdPowerRestore -Run $run
        $receiving = @($run.Result.Events | Where-Object Observation -eq 'bytes')
        $receiving.Count | Should -BeGreaterThan 1
        $receiving[-1].Bytes | Should -Be $script:progressLength
        $receiving[-1].TotalBytes | Should -Be $script:progressLength
        @($run.Result.Events | Where-Object Observation -eq 'validation')[0].Stage | Should -Be 'verifying'
        @($run.Result.Events | Where-Object Observation -eq 'validation')[0].AcceptedHistory | Should -Be 0
        @($run.Result.Renders | Where-Object { -not $_.Completed -and $_.Role -eq 'root' -and $_.Percent -eq 100 }).Count | Should -BeGreaterThan 0
        $archive = Get-UpdProgressArchive -Run $run
        (Get-FileHash -LiteralPath (Join-Path $archive.Root $archive.State.episodes[0].relative_path) -Algorithm SHA256).Hash | Should -Be $script:progressHash
    }

    It 'shows actual unknown-length bytes without inventing a total or receiving percentage' {
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $run = Invoke-UpdProgressWorker -Context $context -FeedPath '/resume/feed/no-length' -ForceInteractive
        $run.ExitCode | Should -Be 0
        Assert-UpdProgressOrdering -Run $run
        $receiving = @($run.Result.Events | Where-Object Observation -eq 'bytes')
        $receiving[-1].Bytes | Should -Be ($script:progressLength * 64)
        $receiving[-1].TotalBytes | Should -BeNullOrEmpty
        $renders = @($run.Result.Renders | Where-Object { -not $_.Completed -and $_.Role -eq 'transfer' -and $_.Snapshot.EpisodeOutcome -ne 'downloaded' })
        $renders.Count | Should -BeGreaterThan 0
        foreach ($render in $renders) {
            $render.Percent | Should -Be -1
            $render.Status | Should -Match 'bytes \(total unknown\)'
        }
        @($run.Result.Power | Where-Object Requested).Count | Should -Be 0
    }

    It 'resets honest byte progress for an actual fresh retry after a truncated uncheckpointed response' {
        $run = Invoke-UpdProgressWorker -Context $context -FeedPath '/transport/feed/truncated-once' -ForceInteractive -MaxAttempts 2
        $run.ExitCode | Should -Be 0
        Assert-UpdProgressOrdering -Run $run
        $responses = @($run.Result.Events | Where-Object Observation -eq 'response')
        $responses.Count | Should -Be 2
        $responses[0].Attempt | Should -Be 1
        $responses[1].Attempt | Should -Be 2
        foreach ($response in $responses) { $response.Offset | Should -Be 0; $response.Bytes | Should -Be 0 }
        @($run.Result.Events | Where-Object { $_.Observation -eq 'bytes' -and $_.Attempt -eq 1 })[-1].Bytes | Should -BeLessThan $script:progressLength
        (Get-UpdFixtureState -Context $context).'/transport/media/truncated-once' | Should -Be 2
    }

    It 'never reports 100 for a completed invalid body and restores the actual native lease after failure' {
        $run = Invoke-UpdProgressWorker -Context $context -FeedPath '/feeds/format-html.xml' -ForceInteractive -KeepAwake
        $run.ExitCode | Should -Be 2
        Assert-UpdProgressOrdering -Run $run
        Assert-UpdPowerRestore -Run $run
        @($run.Result.Events | Where-Object Observation -eq 'validation')[0].Stage | Should -Be 'verifying'
        @($run.Result.Renders | Where-Object { -not $_.Completed -and $_.Percent -eq 100 }).Count | Should -Be 0
        $run.Result.RunResult.Failed | Should -Be 1
        $run.Result.PartialClosed | Should -BeTrue
        $run.Result.LockClosed | Should -BeTrue
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
    }

    It 'closes progress and actual native power on caught cancellation, then starts resumed progress at the retained actual offset' {
        $cancelled = Invoke-UpdProgressWorker -Context $context -FeedPath '/resume/feed/valid' -ForceInteractive -KeepAwake -CancelAfterBytes 8192
        $cancelled.ExitCode | Should -Be 130
        Assert-UpdProgressOrdering -Run $cancelled
        Assert-UpdPowerRestore -Run $cancelled
        $cancelled.Result.CancelBytes | Should -Be 8192
        $cancelled.Result.RunResult.Cancelled | Should -Be 1
        $cancelled.Result.PartialClosed | Should -BeTrue
        $cancelled.Result.LockClosed | Should -BeTrue
        @($cancelled.Result.Renders | Where-Object { -not $_.Completed -and $_.Percent -eq 100 }).Count | Should -Be 0
        $sidecars = @(Get-ChildItem -LiteralPath $cancelled.OutputPath -Recurse -Force -File -Filter 'resume-*.json')
        $sidecars.Count | Should -Be 1
        $checkpoint = [IO.File]::ReadAllText($sidecars[0].FullName) | ConvertFrom-Json
        $checkpoint.offset | Should -Be 8192
        $partialPath = Join-Path (Split-Path $sidecars[0].DirectoryName -Parent) $checkpoint.partial_name
        (Get-Item -LiteralPath $partialPath).Length | Should -Be $checkpoint.offset
        (Get-FileHash -LiteralPath $partialPath -Algorithm SHA256).Hash.ToLowerInvariant() | Should -Be $checkpoint.prefix_sha256
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $resumed = Invoke-UpdProgressWorker -Context $context -FeedPath '/resume/feed/valid' -ForceInteractive
        $resumed.ExitCode | Should -Be 0
        Assert-UpdProgressOrdering -Run $resumed
        $response = @($resumed.Result.Events | Where-Object Observation -eq 'response')[0]
        $response.Offset | Should -Be $checkpoint.offset
        $response.Bytes | Should -Be $checkpoint.offset
        $response.TotalBytes | Should -Be ($script:progressLength * 64)
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        $events[1].range | Should -Be ('bytes=' + $checkpoint.offset + '-')
        $archive = Get-UpdProgressArchive -Run $resumed
        $archive.State.episodes[0].status | Should -Be 'transfer_verified'
        $archive.State.episodes[0].local_sha256 | Should -Be $script:progressResumeHash
        (Get-FileHash -LiteralPath (Join-Path $archive.Root $archive.State.episodes[0].relative_path) -Algorithm SHA256).Hash.ToLowerInvariant() | Should -Be $script:progressResumeHash
        (Get-Item -LiteralPath (Join-Path $archive.Root $archive.State.episodes[0].relative_path)).Length | Should -Be ($script:progressLength * 64)
        Test-Path -LiteralPath $sidecars[0].FullName | Should -BeFalse
    }

    It 'keeps progress below 100 when completed-history recording fails after validated original bytes are placed' {
        $run = Invoke-UpdProgressWorker -Context $context -ForceInteractive -FailHistoryCompletion
        $run.ExitCode | Should -Be 2
        Assert-UpdProgressOrdering -Run $run
        @($run.Result.Renders | Where-Object { -not $_.Completed -and $_.Percent -eq 100 }).Count | Should -Be 0
        $archive = Get-UpdProgressArchive -Run $run
        $archive.State.episodes[0].status | Should -Be 'prepared'
        (Get-FileHash -LiteralPath (Join-Path $archive.Root $archive.State.episodes[0].relative_path) -Algorithm SHA256).Hash | Should -Be $script:progressHash
        $run.Result.LockClosed | Should -BeTrue
    }

    It 'keeps overall progress below 100 for an unresolved catalogue while showing recorded accessible transfer success' {
        $run = Invoke-UpdProgressWorker -Context $context -FeedPath '/feeds/pagination-gap.xml' -ForceInteractive
        $run.ExitCode | Should -Be 2
        Assert-UpdProgressOrdering -Run $run
        $run.Result.RunResult.Downloaded | Should -Be 1
        $run.Result.RunResult.CatalogueComplete | Should -BeFalse
        @($run.Result.Renders | Where-Object { -not $_.Completed -and $_.Role -eq 'root' -and $_.Percent -eq 100 }).Count | Should -Be 0
        @($run.Result.Renders | Where-Object { -not $_.Completed -and $_.Role -eq 'transfer' -and $_.Percent -eq 100 }).Count | Should -BeGreaterThan 0
    }

    It 'keeps preview power off and redirected noninteractive output quiet with a readable truthful summary' {
        $preview = Invoke-UpdProgressWorker -Context $context -Preview -KeepAwake
        $preview.ExitCode | Should -Be 0
        @($preview.Result.Power | Where-Object Requested).Count | Should -Be 0
        @($preview.Result.Renders).Count | Should -Be 0
        Test-Path -LiteralPath $preview.OutputPath | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
        $quiet = Invoke-UpdProgressWorker -Context $context
        $quiet.ExitCode | Should -Be 0
        $quiet.Result.Error | Should -BeNullOrEmpty
        @($quiet.Result.Renders).Count | Should -Be 0
        @($quiet.Result.Power | Where-Object Requested).Count | Should -Be 0
        $quiet.Stdout | Should -Match 'Summary[\s\S]*Downloaded\s+: 1[\s\S]*Failed\s+: 0'
        $quiet.Stdout | Should -Not -Match 'Preparing:|Receiving:|Validating and recording|total unknown|Original synthetic audio|http://127\.0\.0\.1|\x1b\['
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
    }
}
