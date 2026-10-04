BeforeAll {
    $script:CleanupRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:CleanupRepository 'UniversalPodcastDownloader.ps1')
    Mock Get-PodcastHttpClient { throw 'Cleanup units must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Cleanup units must not request metadata.' }
    Mock Read-Host { throw 'Cleanup units must not prompt.' }
    Mock Write-Host {}
    Mock Write-PodcastDiagnostic {}
    if (-not ('UpdCancellationDisposeFileStream' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.IO;
public sealed class UpdCancellationDisposeFileStream : FileStream {
    public int DisposeAttempts;
    public string CleanupFailure;
    public UpdCancellationDisposeFileStream(string path, string failure)
        : base(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None) { CleanupFailure = failure; }
    protected override void Dispose(bool disposing) {
        base.Dispose(disposing);
        if (!disposing) return;
        DisposeAttempts++;
        if (CleanupFailure == "cancel") throw new OperationCanceledException("private stream cleanup cancellation");
        if (CleanupFailure == "io") throw new IOException("private stream cleanup failure");
    }
}
'@
    }
}

Describe 'A043 media cleanup disposal cancellation precedence' -Tag 'Unit', 'A043' {
    BeforeEach {
        $script:CleanupTransferRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($script:CleanupTransferRoot)
        $script:CleanupPartialName = '.upd-' + [guid]::NewGuid().ToString('N') + '.tmp'
        $script:CleanupPartial = Join-Path $script:CleanupTransferRoot $script:CleanupPartialName
        $script:CleanupTransferStream = [UpdCancellationDisposeFileStream]::new($script:CleanupPartial, 'cancel')
        $script:CleanupTransferContext = [pscustomobject]@{
            Lock = [pscustomobject]@{ Root = $script:CleanupTransferRoot }; FeedId = 'a' * 64; EpisodeId = 'b' * 64
        }
        Mock Assert-PodcastResumeLock {}
        Mock Read-PodcastResumeState {
            [pscustomobject]@{ partial_name = $script:CleanupPartialName; offset = [long]0; total_length = [long]4; etag = '"fixture"'; content_type = 'audio/mpeg'; final_uri_fingerprint = 'c' * 64 }
        }
        Mock Open-PodcastResumePartial { $script:CleanupTransferStream }
        Mock Update-PodcastResumeCheckpoint {}
        Mock Invoke-PodcastMediaRequest {
            $DestinationStream.Write([byte[]]@(255, 251, 144, 0), 0, 4)
            [pscustomobject]@{ Completed = $true; Bytes = [long]4; ContentLength = [long]4; ContentType = 'audio/mpeg' }
        }
        $script:CleanupTransferArguments = @{
            Uri = 'https://media.example.invalid/cleanup.mp3'; Root = $script:CleanupTransferRoot
            RelativePath = 'episode.mp3'; ResumeContext = $script:CleanupTransferContext
        }
    }

    AfterEach {
        $script:CleanupTransferStream.CanWrite | Should -BeFalse
        $script:CleanupTransferStream.DisposeAttempts | Should -Be 1
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:CleanupPartial)) | Should -BeExactly '//uQAA=='
        $guardProbe = [IO.File]::Open($script:CleanupPartial, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $guardProbe.Dispose()
        Test-Path -LiteralPath (Join-Path $script:CleanupTransferRoot 'episode.mp3') | Should -BeFalse
    }

    It 'preserves typed cancellation raised while closing after a completed response' {
        $caught = $null
        try { $null = Invoke-PodcastMediaTransfer @script:CleanupTransferArguments }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
    }

    It 'retains an ordinary request failure when secondary stream closing is cancellation' {
        Mock Invoke-PodcastMediaRequest {
            $DestinationStream.Write([byte[]]@(255, 251, 144, 0), 0, 4)
            throw [IO.InvalidDataException]::new('primary transfer failure')
        }
        $caught = $null
        try { $null = Invoke-PodcastMediaTransfer @script:CleanupTransferArguments }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeFalse
        $caught.Exception.Message | Should -BeExactly 'primary transfer failure'
    }

    It 'preserves request cancellation when secondary closing is <Secondary>' -ForEach @(
        @{ Secondary = 'io' }, @{ Secondary = 'cancel' }
    ) {
        $script:CleanupTransferStream.CleanupFailure = $Secondary
        Mock Invoke-PodcastMediaRequest {
            $DestinationStream.Write([byte[]]@(255, 251, 144, 0), 0, 4)
            throw [OperationCanceledException]::new('primary transfer cancellation')
        }
        $caught = $null
        try { $null = Invoke-PodcastMediaTransfer @script:CleanupTransferArguments }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $caught.Exception.Message | Should -BeExactly 'primary transfer cancellation'
    }
}
