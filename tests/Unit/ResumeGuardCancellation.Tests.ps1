BeforeAll {
    $script:ResumeGuardRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:ResumeGuardRepository 'UniversalPodcastDownloader.ps1')
    Mock Get-PodcastHttpClient { throw 'Resume guard units must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Resume guard units must not request metadata.' }
    Mock Read-Host { throw 'Resume guard units must not prompt.' }
    Mock Write-Host {}
    Mock Write-PodcastDiagnostic {}
}

Describe 'A043 resume prefix guard cancellation preserves owned evidence' -Tag 'Unit', 'A043' {
    BeforeEach {
        $script:ResumeGuardRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($script:ResumeGuardRoot)
        $script:ResumeGuardLock = Enter-PodcastHistoryLock -Root $script:ResumeGuardRoot
        $script:ResumeGuardPartialName = '.upd-' + [guid]::NewGuid().ToString('N') + '.tmp'
        $script:ResumeGuardPartial = Join-Path $script:ResumeGuardRoot $script:ResumeGuardPartialName
        $script:ResumeGuardBytes = [byte[]]@(255, 251, 144, 0, 1, 2, 3, 4, 5, 6, 7, 8)
        [IO.File]::WriteAllBytes($script:ResumeGuardPartial, $script:ResumeGuardBytes)
        $script:ResumeGuardState = [pscustomobject]@{
            schema_version = 1; feed_id = 'a' * 64; episode_id = 'b' * 64
            relative_path = 'episode.mp3'; partial_name = $script:ResumeGuardPartialName
            request_fingerprint = 'c' * 64; final_uri_fingerprint = 'd' * 64
            etag = '"guard-fixture"'; total_length = [long]1048576; content_type = 'audio/mpeg'
            content_encoding = 'identity'; offset = [long]$script:ResumeGuardBytes.Length
            prefix_sha256 = (Get-FileHash -LiteralPath $script:ResumeGuardPartial -Algorithm SHA256).Hash.ToLowerInvariant()
        }
        $script:ResumeGuardState = Write-PodcastResumeState -Lock $script:ResumeGuardLock -State $script:ResumeGuardState
        $script:ResumeGuardSidecar = Get-PodcastResumeStatePath -Root $script:ResumeGuardRoot -EpisodeId $script:ResumeGuardState.episode_id
        $script:ResumeGuardSidecarDigest = (Get-FileHash -LiteralPath $script:ResumeGuardSidecar -Algorithm SHA256).Hash
        $script:ResumeGuardSidecarBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:ResumeGuardSidecar))
        $script:ResumeGuardObservedStream = $null
    }

    AfterEach {
        try {
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:ResumeGuardPartial)) |
                Should -BeExactly ([Convert]::ToBase64String($script:ResumeGuardBytes))
            (Get-FileHash -LiteralPath $script:ResumeGuardSidecar -Algorithm SHA256).Hash | Should -BeExactly $script:ResumeGuardSidecarDigest
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:ResumeGuardSidecar)) | Should -BeExactly $script:ResumeGuardSidecarBytes
            if ($null -ne $script:ResumeGuardObservedStream) { $script:ResumeGuardObservedStream.CanRead | Should -BeFalse }
            $proof = [IO.File]::Open($script:ResumeGuardPartial, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            try { $proof.Length | Should -Be $script:ResumeGuardBytes.Length }
            finally { $proof.Dispose() }
            Test-Path -LiteralPath (Join-Path $script:ResumeGuardRoot 'episode.mp3') | Should -BeFalse
        }
        finally { $script:ResumeGuardLock.Stream.Dispose() }
    }

    It 'preserves catchable <Kind> during prefix hashing and closes the actual exclusive partial' -ForEach @(
        @{ Kind = 'operation cancellation'; Factory = { [OperationCanceledException]::new('private prefix cancellation') } }
        @{ Kind = 'task cancellation'; Factory = { [Threading.Tasks.TaskCanceledException]::new('private prefix cancellation') } }
        # A direct PipelineStopped throw stops the PowerShell test pipeline. A
        # wrapper exercises bounded typed detection as a catchable error instead.
        @{ Kind = 'wrapped pipeline cancellation'; Factory = { [Exception]::new('private prefix wrapper', [Management.Automation.PipelineStoppedException]::new()) } }
    ) {
        $script:ResumeGuardCause = & $Factory
        Mock Get-PodcastResumeStreamHash {
            $script:ResumeGuardObservedStream = $Stream
            $Stream | Should -BeOfType ([IO.FileStream])
            $Stream.CanRead | Should -BeTrue
            $Stream.CanWrite | Should -BeTrue
            { [IO.File]::Open($script:ResumeGuardPartial, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite) } | Should -Throw
            throw $script:ResumeGuardCause
        }
        $failure = $null
        try {
            $null = Open-PodcastResumePartial -Lock $script:ResumeGuardLock -State $script:ResumeGuardState `
                -FeedId $script:ResumeGuardState.feed_id -EpisodeId $script:ResumeGuardState.episode_id `
                -RelativePath $script:ResumeGuardState.relative_path -RequestFingerprint $script:ResumeGuardState.request_fingerprint
        }
        catch { $failure = $_ }
        $failure | Should -Not -BeNullOrEmpty
        Test-PodcastCancellation -ErrorObject $failure | Should -BeTrue
        Should -Invoke Get-PodcastResumeStreamHash -Times 1 -Exactly
    }
}
