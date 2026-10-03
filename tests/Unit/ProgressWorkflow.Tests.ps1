BeforeAll {
    $script:progressWorkflowRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:progressWorkflowRoot 'UniversalPodcastDownloader.ps1')
    $script:progressWorkflowMedia = [IO.File]::ReadAllBytes((Join-Path $script:progressWorkflowRoot 'tools/codex-handoff/fixtures/silence.mp3'))
    $script:progressWorkflowHash = (Get-FileHash -LiteralPath (Join-Path $script:progressWorkflowRoot 'tools/codex-handoff/fixtures/silence.mp3') -Algorithm SHA256).Hash
    $script:progressWorkflowClose = (Get-Command Close-PodcastProgress).ScriptBlock
}

Describe 'A044/A045 cleanup boundary after committed or failed work' -Tag 'Unit', 'A044', 'A045' {
    It 'preserves the correct result for <Kind>' -ForEach @(
        @{ Kind='typed-after-success'; Exit=130; Status='cancelled'; Downloaded=1; Failed=0 },
        @{ Kind='typed-after-failure'; Exit=2; Status='incomplete'; Downloaded=0; Failed=1 },
        @{ Kind='ordinary-after-success'; Exit=0; Status='success'; Downloaded=1; Failed=0 }
    ) {
        Mock Read-Host { throw 'Unit checks must not prompt.' }
        Mock Get-PodcastHttpClient { throw 'Unit checks must not contact a network.' }
        Mock Test-PodcastProgressInteractive { $true }
        Mock Write-Host {}
        Mock Write-Progress {}
        Mock Write-PodcastDiagnosticFallback {}
        Mock Invoke-PodcastMetadataRequest {
            [pscustomobject]@{ Content='<rss><channel><title>Owned unit show</title><item><guid>owned-cleanup</guid><title>Owned unit episode</title><enclosure url="https://media.example.invalid/owned.mp3" type="audio/mpeg"/></item></channel></rss>' }
        }
        Mock Invoke-PodcastMediaRequest {
            $bytes = if ($Kind -eq 'typed-after-failure') { [Text.Encoding]::UTF8.GetBytes('<html>owned invalid body</html>') } else { $script:progressWorkflowMedia }
            $DestinationStream.Write($bytes, 0, $bytes.Length)
            [pscustomobject]@{ Completed=$true; Bytes=$bytes.Length; ContentLength=$bytes.Length; ContentType='audio/mpeg' }
        }
        Mock Start-PodcastKeepAwake {
            [pscustomobject]@{ Requested=$true; Active=$true; Lease='owned unit lease'; Message='Temporary keep-awake is active for this confirmed operation.' }
        }
        Mock Stop-PodcastKeepAwake {
            [pscustomobject]@{ Requested=$true; Active=$false; Restored=$true; Message='Temporary keep-awake was released and its prior thread state restored.' }
        }
        Mock Close-PodcastProgress {
            & $script:progressWorkflowClose -Context $Context
            if ($Kind -like 'typed-*') { throw [OperationCanceledException]::new('Owned cleanup cancellation sentinel.') }
            throw [IO.IOException]::new('Owned display cleanup failure sentinel.')
        }

        $outputRoot = Join-Path $TestDrive $Kind
        $result = Invoke-PodcastRun -FeedUrl 'https://feed.example.invalid/owned.xml' -OutputPath $outputRoot -Mode All -KeepAwake
        $result.Type | Should -Be 'Podcast.RunResult'
        $result.ExitCode | Should -Be $Exit
        $result.Status | Should -Be $Status
        $result.Complete | Should -Be ($Exit -eq 0)
        $result.Downloaded | Should -Be $Downloaded
        $result.Failed | Should -Be $Failed
        # Cleanup cancellation does not relabel already committed media as an
        # unfinished episode or fabricate another transfer outcome.
        $result.Cancelled | Should -Be 0
        Should -Invoke Start-PodcastKeepAwake -Times 1 -Exactly -ParameterFilter { $Enabled }
        Should -Invoke Stop-PodcastKeepAwake -Times 1 -Exactly -ParameterFilter { $Context.Requested }
        Should -Invoke Write-Progress -Times 1 -Exactly -ParameterFilter { $Completed }
        @(Get-ChildItem -LiteralPath $outputRoot -Recurse -Force -File -Filter '.upd-*.tmp').Count | Should -Be 0
        foreach ($path in @(Get-ChildItem -LiteralPath $outputRoot -Recurse -Force -File -Filter 'writer.lock')) {
            $guard = [IO.File]::Open($path.FullName, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            $guard.Dispose()
        }
        $statePath = @(Get-ChildItem -LiteralPath $outputRoot -Recurse -Force -File -Filter 'state.json')[0]
        $state = [IO.File]::ReadAllText($statePath.FullName) | ConvertFrom-Json
        @($state.episodes).Count | Should -Be 1
        if ($Downloaded -eq 1) {
            $state.episodes[0].status | Should -Be 'transfer_verified'
            $state.episodes[0].local_sha256 | Should -Be $script:progressWorkflowHash.ToLowerInvariant()
            $media = @(Get-ChildItem -LiteralPath $outputRoot -Recurse -File -Filter '*.mp3')
            $media.Count | Should -Be 1
            (Get-FileHash -LiteralPath $media[0].FullName -Algorithm SHA256).Hash | Should -Be $script:progressWorkflowHash
        }
        else {
            $state.episodes[0].status | Should -Be 'failed'
            @(Get-ChildItem -LiteralPath $outputRoot -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        }
    }
}
