BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'UniversalPodcastDownloader.ps1')
    Mock Get-PodcastHttpClient { throw 'Cancellation units must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Unexpected unit metadata request.' }
    Mock Write-Host {}
    Mock Write-Progress {}
    Mock Read-Host { throw 'Unexpected unit prompt.' }
}

Describe 'A043 graceful cancellation retains closed owned partials' -Tag 'Unit', 'A043' {
    BeforeEach {
        $script:CancellationRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($script:CancellationRoot)
        $script:CancellationLock = $null
        $script:CancellationPartial = $null
        $script:CancellationBytes = [byte[]]@(255, 251, 144, 0, 1, 2, 3, 4, 5, 6, 7, 8)
        $script:UnknownPartial = Join-Path $script:CancellationRoot 'unknown.part'
        [IO.File]::WriteAllText($script:UnknownPartial, 'unclaimed original bytes')
    }
    AfterEach {
        if ($null -ne $script:CancellationLock) { $script:CancellationLock.Stream.Dispose() }
    }

    It 'preserves a newly reserved partial without resume evidence and releases its handle' {
        Mock Invoke-PodcastMediaRequest {
            $script:CancellationPartial = $DestinationStream.Name
            $DestinationStream.Write($script:CancellationBytes, 0, $script:CancellationBytes.Length)
            throw [OperationCanceledException]::new('private-cancellation-canary')
        }
        $caught = $null
        try {
            $null = Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/cancel.mp3' `
                -Root $script:CancellationRoot -RelativePath 'episode.mp3'
        }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        Test-Path -LiteralPath $script:CancellationPartial | Should -BeTrue
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:CancellationPartial)) |
            Should -BeExactly ([Convert]::ToBase64String($script:CancellationBytes))
        $reopened = [IO.File]::Open($script:CancellationPartial, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $reopened.Dispose()
        [IO.File]::ReadAllText($script:UnknownPartial) | Should -BeExactly 'unclaimed original bytes'
        Test-Path -LiteralPath (Join-Path $script:CancellationRoot 'episode.mp3') | Should -BeFalse
        @(Get-ChildItem -LiteralPath $script:CancellationRoot -Recurse -Filter 'resume-*.json').Count | Should -Be 0
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly
    }

    It 'checkpoints bytes written after a periodic checkpoint before releasing cancellation handles' {
        $script:CancellationLock = Enter-PodcastHistoryLock -Root $script:CancellationRoot
        $resumeContext = [pscustomobject]@{ Lock = $script:CancellationLock; FeedId = 'a' * 64; EpisodeId = 'b' * 64 }
        Mock Invoke-PodcastMediaRequest {
            $script:CancellationPartial = $DestinationStream.Name
            $info = [pscustomobject]@{
                ResumeSupported = $true; TotalLength = [long]1048576; Offset = [long]0; ResponseLength = [long]1048576
                ETag = '"cancel-fixture"'; ContentType = 'audio/mpeg'
                FinalUriFingerprint = Get-PodcastResumeUriFingerprint -Uri ([Uri]'https://media.example.invalid/cancel.mp3')
            }
            $null = & $OnResponse $info
            $DestinationStream.Write($script:CancellationBytes, 0, 8)
            $null = & $OnProgress ([long]8)
            $DestinationStream.Write($script:CancellationBytes, 8, 4)
            $null = & $OnProgress ([long]12)
            throw [OperationCanceledException]::new('private-cancellation-canary')
        }
        $caught = $null
        try {
            $null = Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/cancel.mp3' `
                -Root $script:CancellationRoot -RelativePath 'episode.mp3' -ResumeContext $resumeContext
        }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $state = Read-PodcastResumeState -Root $script:CancellationRoot -EpisodeId $resumeContext.EpisodeId
        $state.offset | Should -Be 12
        $state.prefix_sha256 | Should -BeExactly ((Get-FileHash -LiteralPath $script:CancellationPartial -Algorithm SHA256).Hash.ToLowerInvariant())
        $stream = Open-PodcastResumePartial -Lock $script:CancellationLock -State $state `
            -FeedId $resumeContext.FeedId -EpisodeId $resumeContext.EpisodeId -RelativePath 'episode.mp3' `
            -RequestFingerprint (Get-PodcastResumeUriFingerprint -Uri ([Uri]'https://media.example.invalid/cancel.mp3'))
        try { $stream.Position | Should -Be 12; $stream.Length | Should -Be 12 }
        finally { if ($null -ne $stream) { $stream.Dispose() } }
        [IO.File]::ReadAllText($script:UnknownPartial) | Should -BeExactly 'unclaimed original bytes'
        Test-Path -LiteralPath (Join-Path $script:CancellationRoot 'episode.mp3') | Should -BeFalse
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly
    }

    It 'retains prior evidence and cancellation when its final checkpoint write fails' {
        $script:CancellationLock = Enter-PodcastHistoryLock -Root $script:CancellationRoot
        $resumeContext = [pscustomobject]@{ Lock = $script:CancellationLock; FeedId = 'a' * 64; EpisodeId = 'b' * 64 }
        Mock Write-PodcastResumeState { throw [IO.IOException]::new('private-checkpoint-canary') } -ParameterFilter { $State.offset -eq 12 }
        Mock Invoke-PodcastMediaRequest {
            $script:CancellationPartial = $DestinationStream.Name
            $info = [pscustomobject]@{
                ResumeSupported = $true; TotalLength = [long]1048576; Offset = [long]0; ResponseLength = [long]1048576
                ETag = '"cancel-fixture"'; ContentType = 'audio/mpeg'
                FinalUriFingerprint = Get-PodcastResumeUriFingerprint -Uri ([Uri]'https://media.example.invalid/cancel.mp3')
            }
            $null = & $OnResponse $info
            $DestinationStream.Write($script:CancellationBytes, 0, 8)
            $null = & $OnProgress ([long]8)
            $DestinationStream.Write($script:CancellationBytes, 8, 4)
            throw [OperationCanceledException]::new('private-cancellation-canary')
        }
        $caught = $null
        try {
            $null = Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/cancel.mp3' `
                -Root $script:CancellationRoot -RelativePath 'episode.mp3' -ResumeContext $resumeContext
        }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $state = Read-PodcastResumeState -Root $script:CancellationRoot -EpisodeId $resumeContext.EpisodeId
        $state.offset | Should -Be 8
        [IO.File]::ReadAllBytes($script:CancellationPartial).Length | Should -Be 12
        $reopened = [IO.File]::Open($script:CancellationPartial, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $reopened.Dispose()
        { Open-PodcastResumePartial -Lock $script:CancellationLock -State $state `
            -FeedId $resumeContext.FeedId -EpisodeId $resumeContext.EpisodeId -RelativePath 'episode.mp3' `
            -RequestFingerprint (Get-PodcastResumeUriFingerprint -Uri ([Uri]'https://media.example.invalid/cancel.mp3')) } |
            Should -Throw '*inconsistent*preserved*'
        [IO.File]::ReadAllText($script:UnknownPartial) | Should -BeExactly 'unclaimed original bytes'
        Test-Path -LiteralPath (Join-Path $script:CancellationRoot 'episode.mp3') | Should -BeFalse
        Should -Invoke Write-PodcastResumeState -ParameterFilter { $State.offset -eq 12 } -Times 1 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly
    }
}
