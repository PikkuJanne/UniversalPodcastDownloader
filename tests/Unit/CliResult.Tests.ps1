BeforeAll {
    $script:CliRepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:CliDownloaderPath = Join-Path $script:CliRepositoryRoot 'UniversalPodcastDownloader.ps1'
    Mock Read-Host { throw 'CLI units must not prompt unless a test supplies guided input.' }
    Mock Invoke-WebRequest { throw 'CLI units must not contact the external network.' }
    . $script:CliDownloaderPath -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'CLI units must not create a network client.' }
    $script:CliMediaFixture = [IO.File]::ReadAllBytes((Join-Path $script:CliRepositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
}

Describe 'A039: command-line selection validation' -Tag 'Unit', 'A039' {
    BeforeEach {
        $script:CliOutputRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $cliMetadataFixture = [pscustomobject]@{ Content = '<rss><channel><title>CLI unit show</title>' +
            '<item><title>One</title><guid>one</guid><pubDate>2026-10-01</pubDate><enclosure url="https://media.example.invalid/one.mp3" /></item>' +
            '<item><title>Two</title><guid>two</guid><pubDate>2026-10-02</pubDate><enclosure url="https://media.example.invalid/two.mp3" /></item>' +
            '<item><title>Three</title><guid>three</guid><pubDate>2026-10-03</pubDate><enclosure url="https://media.example.invalid/three.mp3" /></item>' +
            '</channel></rss>' }
        Mock Write-Host {}
        Mock Write-Progress {}
        Mock Invoke-PodcastMetadataRequest { $cliMetadataFixture }
        $cliTransferFixture = $script:CliMediaFixture
        Mock Invoke-PodcastMediaRequest {
            $DestinationStream.Write($cliTransferFixture, 0, $cliTransferFixture.Length)
            [pscustomobject]@{
                Completed = $true; Bytes = $cliTransferFixture.Length
                ContentLength = $cliTransferFixture.Length; ContentType = 'audio/mpeg'
            }
        }
    }

    It 'infers Custom from CustomCount alone and downloads exactly that many episodes' {
        $result = Invoke-PodcastRun -FeedUrl 'https://feeds.example.invalid/cli.xml' -CustomCount 2 -OutputPath $script:CliOutputRoot
        $result.Type | Should -BeExactly 'Podcast.RunResult'
        $result.Mode | Should -Be 'Custom'
        $result.ExitCode | Should -Be 0
        $result.Planned | Should -Be 2
        $result.Downloaded | Should -Be 2
        @(Get-ChildItem -LiteralPath $script:CliOutputRoot -Filter '*.mp3' -Recurse).Count | Should -Be 2
        Should -Invoke Invoke-PodcastMediaRequest -Times 2 -Exactly
        Should -Invoke Read-Host -Times 0 -Exactly
    }

    It 'rejects conflicting explicit <Mode> with CustomCount before prompting or metadata' -ForEach @(
        @{ Mode = 'Latest' }, @{ Mode = 'All' }
    ) {
        $result = Invoke-PodcastRun -FeedUrl 'https://feeds.example.invalid/cli.xml' -Mode $Mode -CustomCount 2 -OutputPath $script:CliOutputRoot
        $result.ExitCode | Should -Be 1
        $result.Status | Should -Be 'fatal'
        $result.Message | Should -BeExactly '-CustomCount conflicts with an explicit Latest or All mode.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
        Should -Invoke Read-Host -Times 0 -Exactly
        Test-Path -LiteralPath $script:CliOutputRoot | Should -BeFalse
    }

    It 'uses Latest without any prompt when noninteractive has only an explicit feed' {
        $result = Invoke-PodcastRun -FeedUrl 'https://feeds.example.invalid/cli.xml' -NonInteractive -OutputPath $script:CliOutputRoot
        $result.ExitCode | Should -Be 0
        $result.Mode | Should -Be 'Latest'
        $result.Downloaded | Should -Be 1
        @($result.Episodes).Count | Should -Be 1
        Should -Invoke Read-Host -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://media.example.invalid/three.mp3' }
    }

    It 'rejects <Kind> noninteractive feed input before prompts requests or writes' -ForEach @(
        @{ Kind = 'missing'; Extra = @{} }, @{ Kind = 'blank'; Extra = @{ FeedUrl = '  ' } }
    ) {
        $result = Invoke-PodcastRun -NonInteractive -OutputPath $script:CliOutputRoot @Extra
        $result.ExitCode | Should -Be 1
        $result.Message | Should -BeExactly '-NonInteractive requires an explicit nonblank -FeedUrl.'
        Should -Invoke Read-Host -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
        Test-Path -LiteralPath $script:CliOutputRoot | Should -BeFalse
    }

    It 'rejects noninteractive confirmation before a ShouldProcess prompt' {
        $result = Invoke-PodcastRun -FeedUrl 'https://feeds.example.invalid/cli.xml' -NonInteractive -Confirm -OutputPath $script:CliOutputRoot
        $result.ExitCode | Should -Be 1
        $result.Message | Should -BeExactly '-NonInteractive cannot be combined with -Confirm.'
        Should -Invoke Read-Host -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
        Test-Path -LiteralPath $script:CliOutputRoot | Should -BeFalse
    }

    It 'returns a catchable metadata cancellation without throwing or creating an archive' {
        Mock Invoke-PodcastMetadataRequest { throw [OperationCanceledException]::new('privateCancellationCanary') }
        $result = Invoke-PodcastRun -FeedUrl 'https://feeds.example.invalid/cli.xml' -NonInteractive -OutputPath $script:CliOutputRoot
        $result.ExitCode | Should -Be 130
        $result.Status | Should -Be 'cancelled'
        $result.Complete | Should -BeFalse
        ($result | ConvertTo-Json -Depth 8) | Should -Not -Match 'privateCancellationCanary|Exception|StackTrace'
        Should -Invoke Read-Host -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
        Test-Path -LiteralPath $script:CliOutputRoot | Should -BeFalse
    }

    It 'does not ask for a feed choice in noninteractive HTML ambiguity' {
        Mock Invoke-PodcastMetadataRequest {
            [pscustomobject]@{ Content = '<html><link type="application/rss+xml" href="one.xml"><link type="application/atom+xml" href="two.xml"></html>' }
        }
        $result = Invoke-PodcastRun -FeedUrl 'https://feeds.example.invalid/page' -NonInteractive -OutputPath $script:CliOutputRoot
        $result.ExitCode | Should -Be 1
        $result.Message | Should -BeExactly 'Multiple feed links were found. Supply a direct feed URL with -FeedUrl.'
        Should -Invoke Read-Host -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
        Test-Path -LiteralPath $script:CliOutputRoot | Should -BeFalse
    }

    It 'retains guided feed and count inputs while reusing the feed already fetched' {
        $guidedInputs = [Collections.Generic.Queue[string]]::new()
        $guidedInputs.Enqueue('https://feeds.example.invalid/cli.xml')
        $guidedInputs.Enqueue('2')
        Mock Read-Host { $guidedInputs.Dequeue() }
        $result = Invoke-PodcastRun -OutputPath $script:CliOutputRoot
        $result.ExitCode | Should -Be 0
        $result.Mode | Should -Be 'Custom'
        $result.Planned | Should -Be 2
        $result.Downloaded | Should -Be 2
        $guidedInputs.Count | Should -Be 0
        Should -Invoke Read-Host -Times 2 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 2 -Exactly
    }

    It 'preserves a completed episode and reports cancellation plus remaining deferred work' {
        $cancelMediaFixture = $script:CliMediaFixture
        Mock Invoke-PodcastMediaRequest {
            if ($Uri -eq 'https://media.example.invalid/two.mp3') {
                $DestinationStream.WriteByte(255)
                throw [OperationCanceledException]::new('privateMediaCancellationCanary')
            }
            $DestinationStream.Write($cancelMediaFixture, 0, $cancelMediaFixture.Length)
            [pscustomobject]@{ Completed = $true; Bytes = $cancelMediaFixture.Length; ContentLength = $cancelMediaFixture.Length; ContentType = 'audio/mpeg' }
        }
        $result = Invoke-PodcastRun -FeedUrl 'https://feeds.example.invalid/cli.xml' -Mode All -NonInteractive -OutputPath $script:CliOutputRoot
        $result.ExitCode | Should -Be 130
        $result.Planned | Should -Be 3
        $result.Downloaded | Should -Be 1
        $result.Cancelled | Should -Be 1
        $result.Deferred | Should -Be 1
        @($result.Episodes.Outcome) -join ',' | Should -BeExactly 'downloaded,cancelled,deferred'
        $media = @(Get-ChildItem -LiteralPath $script:CliOutputRoot -Filter '*.mp3' -Recurse)
        $media.Count | Should -Be 1
        [BitConverter]::ToString([IO.File]::ReadAllBytes($media[0].FullName)) | Should -BeExactly ([BitConverter]::ToString($cancelMediaFixture))
        @(Get-ChildItem -LiteralPath $script:CliOutputRoot -Filter '*.tmp' -Recurse -Force).Count | Should -Be 0
        $show = @(Get-ChildItem -LiteralPath $script:CliOutputRoot -Directory)[0]
        $state = Read-PodcastHistory -Root $show.FullName
        $verified = @($state.episodes | Where-Object { $_.status -eq 'transfer_verified' })
        $verified.Count | Should -Be 1
        $verified[0].bytes | Should -Be $cancelMediaFixture.Length
        $verified[0].local_sha256 | Should -BeExactly ((Get-FileHash -LiteralPath $media[0].FullName -Algorithm SHA256).Hash.ToLowerInvariant())
        $serialized = ($result | ConvertTo-Json -Depth 10) + ([IO.File]::ReadAllText((Join-Path $show.FullName '.upd/state.json')))
        $serialized | Should -Not -Match 'privateMediaCancellationCanary|Exception|StackTrace|https://'
        Should -Invoke Invoke-PodcastMediaRequest -Times 2 -Exactly
        Should -Invoke Read-Host -Times 0 -Exactly
    }

    It 'does not relabel a recorded download when cancellation occurs during the completed summary' {
        Mock Write-Log { throw [OperationCanceledException]::new('privateSummaryCancellationCanary') } -ParameterFilter { $Message -eq 'Run completed.' }
        $result = Invoke-PodcastRun -FeedUrl 'https://feeds.example.invalid/cli.xml' -NonInteractive -OutputPath $script:CliOutputRoot
        $result.ExitCode | Should -Be 130
        $result.Status | Should -Be 'cancelled'
        $result.Planned | Should -Be 1
        $result.Downloaded | Should -Be 1
        $result.Cancelled | Should -Be 0
        $result.Episodes.Count | Should -Be 1
        $result.Episodes[0].Outcome | Should -BeExactly 'downloaded'
        $media = @(Get-ChildItem -LiteralPath $script:CliOutputRoot -Filter '*.mp3' -Recurse)
        $media.Count | Should -Be 1
        [BitConverter]::ToString([IO.File]::ReadAllBytes($media[0].FullName)) | Should -BeExactly ([BitConverter]::ToString($script:CliMediaFixture))
        $show = @(Get-ChildItem -LiteralPath $script:CliOutputRoot -Directory)[0]
        $state = Read-PodcastHistory -Root $show.FullName
        $state.episodes.Count | Should -Be 1
        $state.episodes[0].status | Should -Be 'transfer_verified'
        $state.episodes[0].local_sha256 | Should -BeExactly ((Get-FileHash -LiteralPath $media[0].FullName -Algorithm SHA256).Hash.ToLowerInvariant())
        ($result | ConvertTo-Json -Depth 8) | Should -Not -Match 'privateSummaryCancellationCanary|Exception|StackTrace'
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly
        Should -Invoke Read-Host -Times 0 -Exactly
    }
}

Describe 'A039: binding-aware pure CLI options' -Tag 'Unit', 'A039' {
    It 'infers Custom only from an explicitly bound count and preserves supplied bindings' {
        $bound = @{ CustomCount = 2; FeedUrl = 'https://feeds.example.invalid/feed' }
        $before = $bound | ConvertTo-Json -Compress
        $result = Resolve-PodcastCliOptions -BoundParameters $bound -CustomCount 2 -FeedUrl $bound.FeedUrl
        $result.Mode | Should -Be 'Custom'
        $result.CustomCount | Should -Be 2
        $result.NeedFeed | Should -BeFalse
        $result.NeedCount | Should -BeFalse
        ($bound | ConvertTo-Json -Compress) | Should -BeExactly $before
    }

    It 'requests guided inputs according to bindings <Kind>' -ForEach @(
        @{ Kind = 'unbound'; Bound = @{}; NeedFeed = $true; NeedCount = $true; NonInteractive = $false; Legacy = $false }
        @{ Kind = 'feed only'; Bound = @{ FeedUrl = 'https://feeds.example.invalid/feed' }; NeedFeed = $false; NeedCount = $true; NonInteractive = $false; Legacy = $false }
        @{ Kind = 'explicit mode'; Bound = @{ Mode = 'Latest' }; NeedFeed = $true; NeedCount = $false; NonInteractive = $false; Legacy = $false }
        @{ Kind = 'automation'; Bound = @{ FeedUrl = 'https://feeds.example.invalid/feed' }; NeedFeed = $false; NeedCount = $false; NonInteractive = $true; Legacy = $false }
        @{ Kind = 'legacy review'; Bound = @{}; NeedFeed = $true; NeedCount = $false; NonInteractive = $false; Legacy = $true }
    ) {
        $result = Resolve-PodcastCliOptions -BoundParameters $Bound -FeedUrl $Bound.FeedUrl -NonInteractive:$NonInteractive -LegacyRequested:$Legacy
        $result.Mode | Should -Be 'Latest'
        $result.NeedFeed | Should -Be $NeedFeed
        $result.NeedCount | Should -Be $NeedCount
    }

    It 'rejects explicitly bound invalid count <Count> in <Mode>' -ForEach @(
        foreach ($mode in 'Latest', 'All', 'Custom') {
            foreach ($count in 0, -1) { @{ Mode = $mode; Count = $count } }
        }
    ) {
        { Resolve-PodcastCliOptions -BoundParameters @{ Mode = $Mode; CustomCount = $Count } -Mode $Mode -CustomCount $Count } |
            Should -Throw '-CustomCount must be a positive integer.'
    }

    It 'rejects Custom without its required explicit count' {
        { Resolve-PodcastCliOptions -BoundParameters @{ Mode = 'Custom' } -Mode Custom } |
            Should -Throw "Mode 'Custom' requires -CustomCount with a value >= 1."
    }

    It 'accepts explicit noninteractive Confirm false' {
        $result = Resolve-PodcastCliOptions -BoundParameters @{ FeedUrl = 'https://feeds.example.invalid/feed'; Confirm = $false } -FeedUrl 'https://feeds.example.invalid/feed' -NonInteractive
        $result.NeedCount | Should -BeFalse
        $result.NeedFeed | Should -BeFalse
    }
}

Describe 'A040: private episode and run result contract' -Tag 'Unit', 'A040' {
    It 'represents outcome <Outcome> with only safe explicit result fields' -ForEach @(
        foreach ($outcome in 'downloaded', 'verified_skip', 'legacy_unverified', 'conflict', 'deferred', 'failed', 'cancelled', 'planned') { @{ Outcome = $outcome } }
    ) {
        $result = New-PodcastEpisodeResult -EpisodeId ('a' * 64) -Outcome $Outcome -Bytes 123L -Verification 'sha256' -Attempts 2 -Message 'Fixed episode result.'
        $result.Outcome | Should -BeExactly $Outcome
        $result.EpisodeId | Should -BeExactly ('a' * 64)
        $result.Bytes | Should -Be 123L
        $result.Attempts | Should -Be 2
        $result.Verification | Should -Be 'sha256'
        $result.Message | Should -BeExactly 'Fixed episode result.'
        @($result.PSObject.Properties.Name).Count | Should -Be 6
        $result.PSObject.Properties.Name | Should -Not -Contain 'Url'
        $result.PSObject.Properties.Name | Should -Not -Contain 'Episode'
        $result.PSObject.Properties.Name | Should -Not -Contain 'Error'
    }

    It 'keeps unknown bytes nullable rather than describing missing data as zero' {
        $result = New-PodcastEpisodeResult -Outcome failed
        $result.Bytes | Should -BeNullOrEmpty
        $result.Attempts | Should -Be 0
    }

    It 'retains stable result and plan arrays for <Kind>' -ForEach @(
        @{ Kind = 'default'; Arguments = @{}; Count = 0 }
        @{ Kind = 'explicit null'; Arguments = @{ EpisodeResults = $null; Plan = $null }; Count = 0 }
        @{ Kind = 'singleton'; Arguments = @{ EpisodeResults = @([pscustomobject]@{ Outcome = 'downloaded'; EpisodeId = 'one' }); Plan = @([pscustomobject]@{ EpisodeId = 'one'; Source = 'opaque' }) }; Count = 1 }
        @{ Kind = 'multiple'; Arguments = @{ EpisodeResults = @([pscustomobject]@{ Outcome = 'downloaded'; EpisodeId = 'one' }, [pscustomobject]@{ Outcome = 'verified_skip'; EpisodeId = 'two' }); Plan = @([pscustomobject]@{ EpisodeId = 'one' }, [pscustomobject]@{ EpisodeId = 'two' }) }; Count = 2 }
    ) {
        $result = New-PodcastRunResult @Arguments
        ($result.Episodes -is [array]) | Should -BeTrue
        ($result.Plan -is [array]) | Should -BeTrue
        $result.Episodes.Count | Should -Be $Count
        $result.Plan.Count | Should -Be $Count
        if ($Count -gt 0) { $result.Episodes[0].EpisodeId | Should -BeExactly 'one' }
        if ($Count -gt 1) { $result.Episodes[1].EpisodeId | Should -BeExactly 'two' }
    }

    It 'defaults a missing catalogue to complete without inventing fetched pages' {
        $result = New-PodcastRunResult
        $result.Type | Should -BeExactly 'Podcast.RunResult'
        $result.SchemaVersion | Should -Be 1
        $result.CatalogueComplete | Should -BeTrue
        $result.CatalogueStopReason | Should -BeNullOrEmpty
        $result.PagesFetched | Should -Be 0
        $result.ExitCode | Should -Be 0
        $result.Status | Should -Be 'success'
        $result.Complete | Should -BeTrue
        $result.PSObject.Properties.Name | Should -Not -Contain 'LegacyResult'
    }

    It 'counts every completed and unresolved outcome separately' {
        $episodes = @(foreach ($outcome in 'downloaded', 'downloaded', 'verified_skip', 'legacy_unverified', 'conflict', 'deferred', 'failed', 'planned') {
            New-PodcastEpisodeResult -Outcome $outcome
        })
        $result = New-PodcastRunResult -Mode All -EpisodeResults $episodes -Planned 8
        $result.Planned | Should -Be 8
        $result.Downloaded | Should -Be 2
        $result.VerifiedSkipped | Should -Be 1
        $result.LegacyUnverified | Should -Be 1
        $result.Conflicts | Should -Be 1
        $result.Deferred | Should -Be 1
        $result.Failed | Should -Be 1
        $result.Cancelled | Should -Be 0
        $result.ExitCode | Should -Be 2
        $result.Status | Should -Be 'incomplete'
        $result.Complete | Should -BeFalse
    }

    It 'maps isolated unresolved outcome <Outcome> to incomplete' -ForEach @(
        @{ Outcome = 'legacy_unverified' }, @{ Outcome = 'conflict' }, @{ Outcome = 'deferred' }, @{ Outcome = 'failed' }
    ) {
        $result = New-PodcastRunResult -EpisodeResults @(New-PodcastEpisodeResult -Outcome $Outcome)
        $result.ExitCode | Should -Be 2
        $result.Status | Should -Be 'incomplete'
        $result.Complete | Should -BeFalse
    }

    It 'preserves incomplete catalogue status even when fetched episodes succeeded <Preview>' -ForEach @(
        @{ Preview = $false }, @{ Preview = $true }
    ) {
        $result = New-PodcastRunResult -Preview:$Preview -EpisodeResults @(New-PodcastEpisodeResult -Outcome downloaded) -Catalogue ([pscustomobject]@{ Complete = $false; StopReason = 'page_failed'; PagesFetched = 2 })
        $result.ExitCode | Should -Be 2
        $result.Status | Should -Be 'incomplete'
        $result.Preview | Should -Be $Preview
        $result.CatalogueComplete | Should -BeFalse
        $result.CatalogueStopReason | Should -BeExactly 'page_failed'
        $result.PagesFetched | Should -Be 2
    }

    It 'returns a clean preview with planned items and no claimed downloads' {
        $result = New-PodcastRunResult -Preview -Planned 1 -EpisodeResults @(New-PodcastEpisodeResult -Outcome planned)
        $result.ExitCode | Should -Be 0
        $result.Status | Should -Be 'preview'
        $result.Downloaded | Should -Be 0
        $result.Complete | Should -BeTrue
    }

    It 'gives fatal errors precedence over unfinished episodes and catalogue gaps' {
        $result = New-PodcastRunResult -Fatal -EpisodeResults @(New-PodcastEpisodeResult -Outcome failed) -Catalogue ([pscustomobject]@{ Complete = $false; StopReason = 'page_failed'; PagesFetched = 1 })
        $result.ExitCode | Should -Be 1
        $result.Status | Should -Be 'fatal'
    }

    It 'gives <Kind> cancellation precedence over a fatal incomplete run' -ForEach @(
        @{ Kind = 'flag'; Arguments = @{ Cancelled = $true }; Count = 0 }
        @{ Kind = 'episode'; Arguments = @{ EpisodeResults = @([pscustomobject]@{ Outcome = 'cancelled' }) }; Count = 1 }
    ) {
        $result = New-PodcastRunResult -Fatal -Catalogue ([pscustomobject]@{ Complete = $false; StopReason = 'page_failed'; PagesFetched = 1 }) @Arguments
        $result.ExitCode | Should -Be 130
        $result.Status | Should -Be 'cancelled'
        $result.Cancelled | Should -Be $Count
        $result.Complete | Should -BeFalse
    }

    It 'includes local legacy action data only by explicit opt-in without counting it as a download' {
        $legacy = [pscustomobject]@{ Status = 'reviewed'; Rows = @() }
        $result = New-PodcastRunResult -LegacyResult $legacy
        $result.LegacyResult | Should -Be $legacy
        $result.Downloaded | Should -Be 0
        $result.LegacyUnverified | Should -Be 0
        $result.ExitCode | Should -Be 0
    }
}

Describe 'A040: cancellation classification uses exception types' -Tag 'Unit', 'A040' {
    It 'recognizes <Kind> cancellation without changing caller exit state' -ForEach @(
        @{ Kind = 'operation'; Factory = { [OperationCanceledException]::new('private cancellation') } }
        @{ Kind = 'task'; Factory = { [Threading.Tasks.TaskCanceledException]::new('private cancellation') } }
        @{ Kind = 'pipeline'; Factory = { [Management.Automation.PipelineStoppedException]::new() } }
        @{ Kind = 'nested'; Factory = { [Exception]::new('outer', [OperationCanceledException]::new('inner')) } }
        @{ Kind = 'error record'; Factory = { [Management.Automation.ErrorRecord]::new([OperationCanceledException]::new(), 'cancel', [Management.Automation.ErrorCategory]::OperationStopped, $null) } }
    ) {
        $errorValue = & $Factory
        $LASTEXITCODE = 19
        Test-PodcastCancellation -ErrorObject $errorValue | Should -BeTrue
        $LASTEXITCODE | Should -Be 19
    }

    It 'does not infer cancellation from <Kind> text' -ForEach @(
        @{ Kind = 'ordinary exception'; Value = [IO.IOException]::new('cancelled private publisher message') }
        @{ Kind = 'string'; Value = 'OperationCanceledException' }
        @{ Kind = 'null'; Value = $null }
    ) {
        Test-PodcastCancellation -ErrorObject $Value | Should -BeFalse
    }
}
