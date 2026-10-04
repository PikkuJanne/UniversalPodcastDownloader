BeforeAll {
    $script:CleanupRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:CleanupRepository 'UniversalPodcastDownloader.ps1')
    Mock Get-PodcastHttpClient { throw 'Cleanup units must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Cleanup units must not request metadata.' }
    Mock Read-Host { throw 'Cleanup units must not prompt.' }
    Mock Write-Host {}
    Mock Write-PodcastDiagnostic {}

    function Get-LegacyCleanupHandle {
        param([string]$Failure)
        $handle = [pscustomobject]@{ Failure = $Failure; DisposeAttempts = 0 }
        $handle | Add-Member ScriptMethod Dispose {
            $this.DisposeAttempts++
            if ($this.Failure -eq 'cancel') { throw [OperationCanceledException]::new('private cleanup cancellation') }
            if ($this.Failure -eq 'io') { throw [IO.IOException]::new('private cleanup failure') }
        }
        return $handle
    }
}

Describe 'A043 legacy cleanup attempts every owned resource' -Tag 'Unit', 'A043' {
    BeforeEach {
        $script:CleanupRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:CleanupShow = Join-Path $script:CleanupRoot 'Show'
        $null = [IO.Directory]::CreateDirectory($script:CleanupShow)
        $script:CleanupFile = 'Original.mp3'
        $script:CleanupMedia = Join-Path $script:CleanupShow $script:CleanupFile
        [IO.File]::WriteAllBytes($script:CleanupMedia, [IO.File]::ReadAllBytes((Join-Path $script:CleanupRepository 'tools/codex-handoff/fixtures/silence.mp3')))
        $script:CleanupDigest = (Get-FileHash -LiteralPath $script:CleanupMedia -Algorithm SHA256).Hash.ToLowerInvariant()
        $script:CleanupFeed = 'https://feed.example.invalid/cleanup'
        $script:CleanupFeedId = Get-PodcastNameHash -IdentityKey ('feed:' + $script:CleanupFeed)
        $script:CleanupEpisode = [pscustomobject]@{
            Title = 'Original'; Guid = 'cleanup-episode'; AtomId = ''; PubDate = $null
            Url = 'https://media.example.invalid/original.mp3'; EnclosureLength = $null
        }
        $script:CleanupId = (Get-PodcastEpisodeIdentity -Episode $script:CleanupEpisode -FeedId $script:CleanupFeedId).Id
        $script:CleanupArguments = @{
            Root = $script:CleanupRoot; LegacyRoot = $script:CleanupShow; FeedUrl = $script:CleanupFeed
            Episodes = @($script:CleanupEpisode); Action = 'Adopt'; EpisodeId = $script:CleanupId
            FileName = $script:CleanupFile; Sha256 = $script:CleanupDigest; Confirm = $false
        }
        $script:CleanupArchiveHandle = Get-LegacyCleanupHandle
        $script:CleanupHistoryHandle = Get-LegacyCleanupHandle
        Mock Invoke-PodcastDestinationPreflight {}
        Mock Enter-PodcastArchiveLock { $script:CleanupArchiveHandle }
        Mock Enter-PodcastHistoryLock { [pscustomobject]@{ Root = $script:CleanupShow; Stream = $script:CleanupHistoryHandle } }
        Mock ConvertTo-PodcastHistoryV2 { $State.generation = 1; return $State }
        Mock Save-PodcastLegacyCheckpoint { 'legacy-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa.json' }
        Mock Save-PodcastEpisodeRecord {}
        Mock Invoke-PodcastRecordedTransfer { throw 'Adoption must not transfer media.' }
    }

    AfterEach {
        (Get-FileHash -LiteralPath $script:CleanupMedia -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $script:CleanupDigest
        $guardProbe = [IO.File]::Open($script:CleanupMedia, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $guardProbe.Dispose()
        Should -Invoke Invoke-PodcastRecordedTransfer -Times 0 -Exactly
    }

    It 'preserves operation cancellation while both writer cleanup attempts fail and closes the real media guard' {
        $script:CleanupArchiveHandle.Failure = 'io'
        $script:CleanupHistoryHandle.Failure = 'io'
        Mock ConvertTo-PodcastHistoryV2 {
            { [IO.File]::Open($script:CleanupMedia, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) } | Should -Throw
            throw [OperationCanceledException]::new('private primary cancellation')
        }
        $caught = $null
        try { $null = Invoke-PodcastLegacyMigration @script:CleanupArguments }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $script:CleanupHistoryHandle.DisposeAttempts | Should -Be 1
        $script:CleanupArchiveHandle.DisposeAttempts | Should -Be 1
        Should -Invoke Save-PodcastEpisodeRecord -Times 0 -Exactly
        Should -Invoke Write-PodcastDiagnostic -Times 1 -Exactly -ParameterFilter { $Level -eq 'WARN' -and $Message -eq 'Legacy action cleanup failed; retained media, history and the primary failure are preserved.' }
    }

    It 'preserves an ordinary primary failure when secondary cleanup is cancellation' {
        $script:CleanupArchiveHandle.Failure = 'cancel'
        $script:CleanupHistoryHandle.Failure = 'cancel'
        Mock ConvertTo-PodcastHistoryV2 { throw [IO.InvalidDataException]::new('primary metadata failure') }
        $caught = $null
        try { $null = Invoke-PodcastLegacyMigration @script:CleanupArguments }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeFalse
        $caught.Exception.Message | Should -BeExactly 'primary metadata failure'
        $script:CleanupHistoryHandle.DisposeAttempts | Should -Be 1
        $script:CleanupArchiveHandle.DisposeAttempts | Should -Be 1
        Should -Invoke Save-PodcastEpisodeRecord -Times 0 -Exactly
    }

    It 'releases the archive writer when acquiring the show writer is cancelled' {
        $script:CleanupArchiveHandle.Failure = 'io'
        Mock Enter-PodcastHistoryLock { throw [OperationCanceledException]::new('private writer cancellation') }
        $caught = $null
        try { $null = Invoke-PodcastLegacyMigration @script:CleanupArguments }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $script:CleanupArchiveHandle.DisposeAttempts | Should -Be 1
        $script:CleanupHistoryHandle.DisposeAttempts | Should -Be 0
        Should -Invoke ConvertTo-PodcastHistoryV2 -Times 0 -Exactly
    }

    It 'reports an ordinary closing failure after the action without returning success' {
        $script:CleanupHistoryHandle.Failure = 'io'
        { Invoke-PodcastLegacyMigration @script:CleanupArguments } |
            Should -Throw 'Legacy action resources could not be closed safely; retained media and history require review.'
        $script:CleanupHistoryHandle.DisposeAttempts | Should -Be 1
        $script:CleanupArchiveHandle.DisposeAttempts | Should -Be 1
        Should -Invoke Save-PodcastEpisodeRecord -Times 1 -Exactly
    }

    It 'preserves typed cancellation raised while closing after the action' {
        $script:CleanupHistoryHandle.Failure = 'cancel'
        $caught = $null
        try { $null = Invoke-PodcastLegacyMigration @script:CleanupArguments }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $script:CleanupHistoryHandle.DisposeAttempts | Should -Be 1
        $script:CleanupArchiveHandle.DisposeAttempts | Should -Be 1
        Should -Invoke Save-PodcastEpisodeRecord -Times 1 -Exactly
    }

    It 'preserves typed inventory cancellation in direct observation' {
        Mock Get-PodcastFileEvidence {
            { [IO.File]::Open($script:CleanupMedia, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) } | Should -Throw
            throw [OperationCanceledException]::new('private inventory hashing cancellation')
        }
        $caught = $null
        try { $null = Get-PodcastLegacyFileObservation -Root $script:CleanupShow -RelativePath $script:CleanupFile }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:CleanupShow '.upd') | Should -BeFalse
        Should -Invoke Enter-PodcastArchiveLock -Times 0 -Exactly
        Should -Invoke Enter-PodcastHistoryLock -Times 0 -Exactly
        Should -Invoke Save-PodcastEpisodeRecord -Times 0 -Exactly
    }

    It 'preserves typed inventory cancellation in complete legacy preview' {
        Mock Test-PodcastMediaFile { throw [OperationCanceledException]::new('private inventory sniff cancellation') }
        $script:CleanupArguments.Action = 'Preview'
        $caught = $null
        try { $null = Invoke-PodcastLegacyMigration @script:CleanupArguments }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:CleanupShow '.upd') | Should -BeFalse
        Should -Invoke Invoke-PodcastDestinationPreflight -Times 0 -Exactly
        Should -Invoke Enter-PodcastArchiveLock -Times 0 -Exactly
        Should -Invoke Enter-PodcastHistoryLock -Times 0 -Exactly
        Should -Invoke Save-PodcastEpisodeRecord -Times 0 -Exactly
    }
}
