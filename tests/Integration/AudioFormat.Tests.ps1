BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    . (Join-Path $repositoryRoot 'UniversalPodcastDownloader.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned loopback audio-format integration fixtures.' }
    $script:formatPython = $python.Source
    $script:formatAsset = @{}
    foreach ($name in @('silence.mp3', 'format-audio.m4a', 'format-isom-audio.m4a', 'format-audio.wav', 'format-audio.flac', 'format-audio.ogg')) {
        $path = Join-Path $repositoryRoot ('tools/codex-handoff/fixtures/' + $name)
        $script:formatAsset[$name] = [pscustomobject]@{
            Hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
            Bytes = (Get-Item -LiteralPath $path).Length
        }
    }

    function Get-UpdFormatArchive {
        param([Parameter(Mandatory)][string]$OutputPath)
        $states = @(Get-ChildItem -LiteralPath $OutputPath -Recurse -Force -File -Filter 'state.json')
        $states.Count | Should -Be 1
        $archiveRoot = Split-Path $states[0].DirectoryName -Parent
        $state = Read-PodcastHistory -Root $archiveRoot
        @($state.episodes).Count | Should -Be 1
        [pscustomobject]@{
            Root = $archiveRoot
            Path = $states[0].FullName
            State = $state
            Record = $state.episodes[0]
            MediaPath = Join-Path $archiveRoot $state.episodes[0].relative_path
        }
    }

    function Get-UpdFormatFile {
        param([Parameter(Mandatory)][string]$OutputPath)
        @(Get-ChildItem -LiteralPath $OutputPath -Recurse -Force -File | Where-Object {
            $_.Extension -in @('.mp3', '.m4a', '.mp4', '.ogg', '.wav', '.flac', '.bin')
        })
    }
}

Describe 'A035/A036 supported audio at real transfer and history boundaries' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:formatPython
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'selects the audio enclosure after declared video for <Kind>' -TestCases @(
        @{ Kind = 'RSS'; Feed = 'audio-after-video'; IdentitySource = 'rss-guid' }
        @{ Kind = 'Atom'; Feed = 'atom-audio-after-video'; IdentitySource = 'atom-id' }
    ) {
        param($Kind, $Feed, $IdentitySource)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath ('/feeds/format-' + $Feed + '.xml')
        $run.Result.Succeeded | Should -BeTrue -Because ($Kind + ': ' + $run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Stdout | Should -Match 'Downloaded\s+: 1'
        $archive = Get-UpdFormatArchive -OutputPath $run.OutputPath
        $archive.Record.identity_source | Should -Be $IdentitySource
        [IO.Path]::GetExtension($archive.MediaPath) | Should -Be '.mp3'
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset['silence.mp3'].Hash
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/media/format-mp3' | Should -Be 1
        $stats.'/media/format-video.mp4' | Should -BeNullOrEmpty
    }

    It 'uses original <Asset> bytes and the recognized extension for <Feed>, then skips the same file' -TestCases @(
        @{ Feed = 'm4a-generic'; Media = '/media/format-m4a'; Asset = 'format-audio.m4a'; Extension = '.m4a'; Kind = 'mp4-container' }
        @{ Feed = 'isom-audio'; Media = '/media/format-isom-audio'; Asset = 'format-isom-audio.m4a'; Extension = '.m4a'; Kind = 'mp4-container' }
        @{ Feed = 'm4a-wrong-extension'; Media = '/media/format-wrong.mp3'; Asset = 'format-audio.m4a'; Extension = '.m4a'; Kind = 'mp4-container' }
        @{ Feed = 'mp3-wrong-extension'; Media = '/media/format-wrong.m4a'; Asset = 'silence.mp3'; Extension = '.mp3'; Kind = 'mpeg-audio' }
        @{ Feed = 'wave'; Media = '/media/format-wave'; Asset = 'format-audio.wav'; Extension = '.wav'; Kind = 'wave' }
        @{ Feed = 'flac'; Media = '/media/format-flac'; Asset = 'format-audio.flac'; Extension = '.flac'; Kind = 'flac' }
        @{ Feed = 'ogg'; Media = '/media/format-ogg'; Asset = 'format-audio.ogg'; Extension = '.ogg'; Kind = 'ogg-container' }
    ) {
        param($Feed, $Media, $Asset, $Extension, $Kind)
        $feedPath = '/feeds/format-' + $Feed + '.xml'
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $feedPath
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $first.Stdout | Should -Match 'Downloaded\s+: 1'
        $archive = Get-UpdFormatArchive -OutputPath $first.OutputPath
        $archive.Record.status | Should -Be 'transfer_verified'
        $archive.Record.verification.media_kind | Should -Be $Kind
        $archive.Record.verification.method | Should -Be 'completed-http-length-and-signature'
        $archive.Record.relative_path | Should -Be ([IO.Path]::GetFileName($archive.MediaPath))
        [IO.Path]::GetExtension($archive.MediaPath) | Should -Be $Extension
        @(Get-UpdFormatFile -OutputPath $first.OutputPath).Count | Should -Be 1
        $mediaFile = Get-Item -LiteralPath $archive.MediaPath
        $mediaFile.Length | Should -Be $script:formatAsset[$Asset].Bytes
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset[$Asset].Hash
        $stamp = $mediaFile.LastWriteTimeUtc
        if ($Feed -like '*wrong-extension') {
            $first.Stdout | Should -Match 'disagrees with the (detected audio format|recognized binary signature)'
        }
        $repeat = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $feedPath
        $repeat.Result.Succeeded | Should -BeTrue -Because ($repeat.Stdout + $repeat.Stderr + $repeat.Result.ErrorMessage)
        $repeat.Stdout | Should -Match 'Skipped\s+: 1'
        $after = Get-UpdFormatArchive -OutputPath $repeat.OutputPath
        $after.MediaPath | Should -Be $archive.MediaPath
        (Get-Item -LiteralPath $archive.MediaPath).LastWriteTimeUtc | Should -Be $stamp
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset[$Asset].Hash
        (Get-UpdFixtureState -Context $context).PSObject.Properties[$Media].Value | Should -Be 1
    }

    It 'rejects audio-labelled <Feed> bytes with <Category> and no final media' -TestCases @(
        @{ Feed = 'html'; Media = '/media/format-html.mp3'; Category = 'non_audio_text' }
        @{ Feed = 'video'; Media = '/media/format-video.mp4'; Category = 'unsupported_media' }
        @{ Feed = 'ambiguous-mp4'; Media = '/media/format-ambiguous.mp4'; Category = 'ambiguous_media' }
        @{ Feed = 'ambiguous-ogg'; Media = '/media/format-ambiguous.ogg'; Category = 'ambiguous_media' }
    ) {
        param($Feed, $Media, $Category)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath ('/feeds/format-' + $Feed + '.xml')
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Stdout | Should -Match 'Downloaded\s+: 0'
        $run.Stdout | Should -Match 'Failed\s+: 1'
        $run.Stdout | Should -Match $Category
        @(Get-UpdFormatFile -OutputPath $run.OutputPath).Count | Should -Be 0
        # These strong-validator fixtures have a durable owned checkpoint.
        # Rejected bytes remain evidence for review, never accepted final media.
        $run.Result.RemainingTemporaryCount | Should -Be 1
        $archive = Get-UpdFormatArchive -OutputPath $run.OutputPath
        $archive.Record.status | Should -Be 'failed'
        $archive.Record.bytes | Should -BeNullOrEmpty
        $archive.Record.local_sha256 | Should -BeNullOrEmpty
        $sidecars = @(Get-ChildItem -LiteralPath (Join-Path $archive.Root '.upd') -File -Filter 'resume-*.json')
        $sidecars.Count | Should -Be 1
        $checkpoint = Get-Content -LiteralPath $sidecars[0].FullName -Raw | ConvertFrom-Json
        $partialPath = Join-Path $archive.Root $checkpoint.partial_name
        (Get-Item -LiteralPath $partialPath).Length | Should -Be $checkpoint.offset
        $checkpoint.offset | Should -Be $checkpoint.total_length
        (Get-FileHash -LiteralPath $partialPath -Algorithm SHA256).Hash | Should -Be $checkpoint.prefix_sha256
        $probe = [IO.File]::Open($partialPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $probe.Dispose()
        (Get-UpdFixtureState -Context $context).PSObject.Properties[$Media].Value | Should -Be 1
    }

    It 'previews a conflicting extension without requesting enclosures or writing an archive' {
        $preview = Invoke-UpdIntegrationWorker -Context $context -Action Preview -FeedPath '/feeds/format-m4a-wrong-extension.xml'
        $preview.Result.Succeeded | Should -BeTrue -Because ($preview.Stdout + $preview.Stderr + $preview.Result.ErrorMessage)
        Test-Path -LiteralPath $preview.OutputPath | Should -BeFalse
        @(Get-ChildItem -LiteralPath $context.Root -Recurse -Force -File -Filter '*.log').Count | Should -Be 0
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/format-m4a-wrong-extension.xml' | Should -Be 1
        $stats.'/media/format-wrong.mp3' | Should -BeNullOrEmpty
    }

    It 'persists the corrected prepared path and reconciles <Stage>' -TestCases @(
        @{ Stage = 'BeforeFinalizeCrash'; Downloads = 1; Skips = 0; Requests = 2 }
        @{ Stage = 'AfterFinalizeCrash'; Downloads = 0; Skips = 1; Requests = 1 }
    ) {
        param($Stage, $Downloads, $Skips, $Requests)
        $crashed = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml' -TransactionHook $Stage
        $crashed.Result | Should -BeNullOrEmpty
        $crashed.HookMarker.Hook | Should -Be $Stage
        $archive = Get-UpdFormatArchive -OutputPath $crashed.OutputPath
        $archive.Record.status | Should -Be 'prepared'
        [IO.Path]::GetExtension($archive.MediaPath) | Should -Be '.m4a'
        $archive.MediaPath | Should -Be $crashed.HookMarker.Destination
        $archive.Record.bytes | Should -Be $script:formatAsset['format-audio.m4a'].Bytes
        $archive.Record.local_sha256 | Should -Be $script:formatAsset['format-audio.m4a'].Hash
        if ($Stage -eq 'BeforeFinalizeCrash') {
            Test-Path -LiteralPath $archive.MediaPath | Should -BeFalse
            (Get-FileHash -LiteralPath $crashed.HookMarker.Temporary -Algorithm SHA256).Hash | Should -Be $script:formatAsset['format-audio.m4a'].Hash
        }
        else { (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset['format-audio.m4a'].Hash }
        $restart = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml'
        $restart.Result.Succeeded | Should -BeTrue -Because ($restart.Stdout + $restart.Stderr + $restart.Result.ErrorMessage)
        $restart.Stdout | Should -Match ('Downloaded\s+: ' + $Downloads)
        $restart.Stdout | Should -Match ('Skipped\s+: ' + $Skips)
        (Get-UpdFormatArchive -OutputPath $restart.OutputPath).Record.status | Should -Be 'transfer_verified'
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset['format-audio.m4a'].Hash
        (Get-UpdFixtureState -Context $context).'/media/format-wrong.mp3' | Should -Be $Requests
    }

    It 'preserves the completed provisional checkpoint and corrected prepared evidence for review after a crash' {
        $crashed = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml' -TransactionHook AfterPrepareBeforeResumeRetireCrash
        $crashed.Result | Should -BeNullOrEmpty
        $crashed.HookMarker.Hook | Should -Be 'AfterPrepareBeforeResumeRetireCrash'
        $archive = Get-UpdFormatArchive -OutputPath $crashed.OutputPath
        $archive.Record.status | Should -Be 'prepared'
        [IO.Path]::GetExtension($archive.MediaPath) | Should -Be '.m4a'
        Test-Path -LiteralPath $archive.MediaPath | Should -BeFalse
        $resumePath = $crashed.HookMarker.ResumePath
        $checkpoint = Get-Content -LiteralPath $resumePath -Raw | ConvertFrom-Json
        [IO.Path]::GetExtension($checkpoint.relative_path) | Should -Be '.mp3'
        $checkpoint.offset | Should -Be $script:formatAsset['format-audio.m4a'].Bytes
        $sidecarHash = (Get-FileHash -LiteralPath $resumePath -Algorithm SHA256).Hash
        $partialHash = (Get-FileHash -LiteralPath $crashed.HookMarker.Temporary -Algorithm SHA256).Hash
        $partialHash | Should -Be $script:formatAsset['format-audio.m4a'].Hash
        $restart = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml'
        $restart.Result.Succeeded | Should -BeFalse
        $restart.Stdout | Should -Match 'Failed\s+: 1'
        ($restart.Stdout + $restart.Result.ErrorMessage) | Should -Match 'review|inconsistent|checkpoint'
        $after = Get-UpdFormatArchive -OutputPath $restart.OutputPath
        $after.MediaPath | Should -Be $archive.MediaPath
        $after.Record.status | Should -Be 'missing'
        $after.Record.bytes | Should -Be $archive.Record.bytes
        $after.Record.local_sha256 | Should -Be $archive.Record.local_sha256
        Test-Path -LiteralPath $after.MediaPath | Should -BeFalse
        (Get-FileHash -LiteralPath $resumePath -Algorithm SHA256).Hash | Should -Be $sidecarHash
        (Get-FileHash -LiteralPath $crashed.HookMarker.Temporary -Algorithm SHA256).Hash | Should -Be $partialHash
        (Get-UpdFixtureState -Context $context).'/media/format-wrong.mp3' | Should -Be 1
    }

    It 'preserves a concurrent final file at the corrected extension without overwrite' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml' -TransactionHook FinalRace
        $run.Result.Succeeded | Should -BeFalse
        $run.HookMarker.Hook | Should -Be 'FinalRace'
        $run.HookMarker.TemporaryExclusiveOpen | Should -BeTrue
        [IO.Path]::GetExtension($run.HookMarker.Destination) | Should -Be '.m4a'
        [IO.File]::ReadAllText($run.HookMarker.Destination) | Should -Be 'Synthetic concurrent final file. Preserve these exact bytes.'
        $archive = Get-UpdFormatArchive -OutputPath $run.OutputPath
        $archive.MediaPath | Should -Be $run.HookMarker.Destination
        $archive.Record.status | Should -Be 'prepared'
        $run.Result.RemainingTemporaryCount | Should -Be 0
        (Get-UpdFixtureState -Context $context).'/media/format-wrong.mp3' | Should -Be 1
    }

    It 'corrects a provisional suffix in failed history without completed-file evidence' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdFormatArchive -OutputPath $first.OutputPath
        # Seed only this owned fixture's uncompleted allocation. It has no final
        # media, prepared digest or completed-file path to preserve.
        [IO.File]::Delete($archive.MediaPath)
        $provisionalName = [IO.Path]::ChangeExtension($archive.Record.relative_path, '.mp3')
        $historyLock = Enter-PodcastHistoryLock -Root $archive.Root
        try {
            $state = Read-PodcastHistory -Root $archive.Root
            $record = $state.episodes[0]
            $record.relative_path = $provisionalName
            $record.status = 'failed'
            $record.bytes = $null
            $record.local_sha256 = $null
            $record.completed_utc = $null
            $record.verification = [pscustomobject]@{ method = 'none'; media_kind = $null; notes = @() }
            $state.generation++
            $null = Write-PodcastHistory -Lock $historyLock -State $state
        }
        finally { $historyLock.Stream.Dispose() }
        $retry = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml'
        $retry.Result.Succeeded | Should -BeTrue -Because ($retry.Stdout + $retry.Stderr + $retry.Result.ErrorMessage)
        $retry.Stdout | Should -Match 'Downloaded\s+: 1'
        $after = Get-UpdFormatArchive -OutputPath $retry.OutputPath
        $after.Record.episode_id | Should -Be $archive.Record.episode_id
        $after.Record.status | Should -Be 'transfer_verified'
        [IO.Path]::GetExtension($after.MediaPath) | Should -Be '.m4a'
        (Get-FileHash -LiteralPath $after.MediaPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset['format-audio.m4a'].Hash
        Test-Path -LiteralPath (Join-Path $archive.Root $provisionalName) | Should -BeFalse
        @(Get-UpdFormatFile -OutputPath $retry.OutputPath).Count | Should -Be 1
        (Get-UpdFixtureState -Context $context).'/media/format-wrong.mp3' | Should -Be 2
    }

    It 'honors a historical recorded suffix on repeat and missing-file redownload without changing original bytes' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdFormatArchive -OutputPath $first.OutputPath
        $historicalName = [IO.Path]::ChangeExtension($archive.Record.relative_path, '.mp3')
        $historicalPath = Join-Path $archive.Root $historicalName
        [IO.File]::Move($archive.MediaPath, $historicalPath)
        $historyLock = Enter-PodcastHistoryLock -Root $archive.Root
        try {
            $state = Read-PodcastHistory -Root $archive.Root
            $state.episodes[0].relative_path = $historicalName
            $state.generation++
            $null = Write-PodcastHistory -Lock $historyLock -State $state
        }
        finally { $historyLock.Stream.Dispose() }
        $stamp = (Get-Item -LiteralPath $historicalPath).LastWriteTimeUtc
        $repeat = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml'
        $repeat.Result.Succeeded | Should -BeTrue -Because ($repeat.Stdout + $repeat.Stderr + $repeat.Result.ErrorMessage)
        $repeat.Stdout | Should -Match 'Skipped\s+: 1'
        (Get-UpdFormatArchive -OutputPath $repeat.OutputPath).MediaPath | Should -Be $historicalPath
        (Get-Item -LiteralPath $historicalPath).LastWriteTimeUtc | Should -Be $stamp
        (Get-FileHash -LiteralPath $historicalPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset['format-audio.m4a'].Hash
        (Get-UpdFixtureState -Context $context).'/media/format-wrong.mp3' | Should -Be 1
        [IO.File]::Delete($historicalPath)
        $redownload = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/format-m4a-wrong-extension.xml'
        $redownload.Result.Succeeded | Should -BeTrue -Because ($redownload.Stdout + $redownload.Stderr + $redownload.Result.ErrorMessage)
        $redownload.Stdout | Should -Match 'Recorded destination extension retained'
        (Get-UpdFormatArchive -OutputPath $redownload.OutputPath).MediaPath | Should -Be $historicalPath
        (Get-FileHash -LiteralPath $historicalPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset['format-audio.m4a'].Hash
        Test-Path -LiteralPath $archive.MediaPath | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/media/format-wrong.mp3' | Should -Be 2
    }

    It 'keeps URL-derived identity when a later typed alternative is added after a URL-only audio candidate' {
        $feedPath = '/feeds/format-identity-selection.xml'
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $feedPath
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdFormatArchive -OutputPath $first.OutputPath
        $archive.Record.identity_source | Should -Be 'media-url'
        $stamp = (Get-Item -LiteralPath $archive.MediaPath).LastWriteTimeUtc
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing -TimeoutSec 10
        $repeat = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $feedPath
        $repeat.Result.Succeeded | Should -BeTrue -Because ($repeat.Stdout + $repeat.Stderr + $repeat.Result.ErrorMessage)
        $repeat.Stdout | Should -Match 'Skipped\s+: 1'
        $after = Get-UpdFormatArchive -OutputPath $repeat.OutputPath
        $after.Record.episode_id | Should -Be $archive.Record.episode_id
        $after.MediaPath | Should -Be $archive.MediaPath
        (Get-Item -LiteralPath $archive.MediaPath).LastWriteTimeUtc | Should -Be $stamp
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:formatAsset['silence.mp3'].Hash
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/media/format-stable.mp3' | Should -Be 1
        $stats.'/media/format-m4a' | Should -BeNullOrEmpty
    }
}
