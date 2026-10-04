BeforeAll {
    $script:ProgressRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:ProgressRepository 'UniversalPodcastDownloader.ps1')
    . (Join-Path $script:ProgressRepository 'src/Progress.ps1')
    if (Get-Command Test-PodcastProgressInteractive -ErrorAction SilentlyContinue) {
        Mock Test-PodcastProgressInteractive { $true }
    }
    Mock Get-PodcastHttpClient { throw 'Progress units must not create a network client.' }
    Mock Read-Host { throw 'Progress units must not prompt.' }
    Mock Write-Host {}
}

Describe 'A044 desired main progress boundaries' -Tag 'Unit', 'A044' {
    BeforeEach {
        $script:ProgressOutput = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:ProgressRows = [Collections.Generic.List[object]]::new()
        $script:ProgressRowsBeforeTransfer = @()
        $script:ProgressPreferenceDuringTransfer = $null
        $script:ProgressPreferenceBeforeTest = $global:ProgressPreference
        $script:ProgressFixture = [IO.File]::ReadAllBytes((Join-Path $script:ProgressRepository 'tools/codex-handoff/fixtures/silence.mp3'))
        Mock Write-Progress {
            $script:ProgressRows.Add([pscustomobject]@{ Id = $Id; ParentId = $ParentId; Status = $Status; PercentComplete = $PercentComplete; Completed = [bool]$Completed })
        }
        Mock Invoke-PodcastMetadataRequest {
            [pscustomobject]@{
                Content = '<rss><channel><title>Progress fixture</title><item><title>private publisher title</title><guid>progress-unit</guid><enclosure url="https://media.example.invalid/unit.mp3" type="audio/mpeg"/></item></channel></rss>'
                FinalUri = [uri]'https://feed.example.invalid/progress'; ContentType = 'application/rss+xml'
            }
        }
        Mock Invoke-PodcastRecordedTransfer {
            $script:ProgressRowsBeforeTransfer = @($script:ProgressRows.ToArray())
            $script:ProgressPreferenceDuringTransfer = $global:ProgressPreference
            $path = Join-Path $Context.Lock.Root $Planned.FileName
            [IO.File]::WriteAllBytes($path, $script:ProgressFixture)
            [pscustomobject]@{
                File = $path; RelativePath = $Planned.FileName; Bytes = [long]$script:ProgressFixture.Length
                Verification = 'length-and-signature'; Attempts = 1; Warnings = @()
            }
        }
    }

    AfterEach { $global:ProgressPreference = $script:ProgressPreferenceBeforeTest }

    It 'keeps preparing progress below complete before any episode transfer returns' {
        $global:ProgressPreference = 'Continue'
        $result = Invoke-PodcastRun -FeedUrl 'https://feed.example.invalid/progress' -Mode All -OutputPath $script:ProgressOutput -Confirm:$false
        $result.ExitCode | Should -Be 0
        @($script:ProgressRowsBeforeTransfer | Where-Object { $_.Status -like 'Preparing: episode*' }).Count | Should -BeGreaterThan 0
        @($script:ProgressRowsBeforeTransfer | Where-Object { $_.PercentComplete -ge 100 -and -not $_.Completed }).Count | Should -Be 0
        Should -Invoke Invoke-PodcastRecordedTransfer -Times 1 -Exactly
    }

    It 'suppresses progress records in noninteractive mode while preserving the result' {
        $global:ProgressPreference = 'Continue'
        $result = Invoke-PodcastRun -FeedUrl 'https://feed.example.invalid/progress' -Mode All -OutputPath $script:ProgressOutput -NonInteractive -Confirm:$false
        $result.ExitCode | Should -Be 0
        $result.Downloaded | Should -Be 1
        $script:ProgressRows.Count | Should -Be 0
    }

    It 'respects the caller progress preference throughout execution' {
        $global:ProgressPreference = 'SilentlyContinue'
        $result = Invoke-PodcastRun -FeedUrl 'https://feed.example.invalid/progress' -Mode All -OutputPath $script:ProgressOutput -Confirm:$false
        $result.ExitCode | Should -Be 0
        $script:ProgressPreferenceDuringTransfer | Should -Be 'SilentlyContinue'
        $global:ProgressPreference | Should -Be 'SilentlyContinue'
        $script:ProgressRows.Count | Should -Be 0
    }
}

Describe 'A044 honest private progress state' -Tag 'Unit', 'A044' {
    BeforeEach {
        $script:ProgressPreferenceBeforeTest = $global:ProgressPreference
        $global:ProgressPreference = 'Continue'
        $script:ProgressTick = 0.0
        $script:RenderedProgress = [Collections.Generic.List[object]]::new()
        Mock Write-Progress {
            $script:RenderedProgress.Add([pscustomobject]@{
                Id = $Id; ParentId = $ParentId; Activity = $Activity; Status = $Status
                CurrentOperation = $CurrentOperation; PercentComplete = $PercentComplete; Completed = [bool]$Completed
            })
            'private sink pipeline output'
        }
    }
    AfterEach { $global:ProgressPreference = $script:ProgressPreferenceBeforeTest }

    It 'creates independent owned IDs without starting progress or writing to the pipeline twice' {
        $contexts = @(New-PodcastProgressContext -TotalEpisodes 2; New-PodcastProgressContext -TotalEpisodes 3)
        $contexts.Count | Should -Be 2
        @($contexts.RootId + $contexts.TransferId | Sort-Object -Unique).Count | Should -Be 4
        @($contexts.RootId + $contexts.TransferId | Where-Object { $_ -le 1 }).Count | Should -Be 0
        $script:RenderedProgress.Count | Should -Be 0
    }

    It 'keeps <Kind> contexts quiet through byte completion and cleanup' -ForEach @(
        @{ Kind = 'noninteractive'; Preference = 'Continue'; NonInteractive = $true; Total = 1 }
        @{ Kind = 'silent preference'; Preference = 'SilentlyContinue'; NonInteractive = $false; Total = 1 }
        @{ Kind = 'stop preference'; Preference = 'Stop'; NonInteractive = $false; Total = 1 }
        @{ Kind = 'inquire preference'; Preference = 'Inquire'; NonInteractive = $false; Total = 1 }
        @{ Kind = 'redirected or unsupported host'; Preference = 'Continue'; NonInteractive = $false; Total = 1 }
        @{ Kind = 'empty selection'; Preference = 'Continue'; NonInteractive = $false; Total = 0 }
    ) {
        $global:ProgressPreference = $Preference
        if ($Kind -eq 'redirected or unsupported host') { Mock Test-PodcastProgressInteractive { $false } }
        $context = New-PodcastProgressContext -TotalEpisodes $Total -NonInteractive:$NonInteractive
        @(
            Start-PodcastEpisodeProgress -Context $context -Index 1
            Start-PodcastTransferProgress -Context $context -Offset 0 -TotalBytes 10 -Attempt 1
            Update-PodcastTransferProgress -Context $context -Bytes 10 -Stage verifying
            Complete-PodcastEpisodeProgress -Context $context -Outcome downloaded
            Complete-PodcastRunProgress -Context $context -Verified:$true
            Close-PodcastProgress -Context $context
        ).Count | Should -Be 0
        $context.Enabled | Should -BeFalse
        $script:RenderedProgress.Count | Should -Be 0
        $global:ProgressPreference | Should -Be $Preference
    }

    It 'renders known bytes <Bytes> of <Total> without reaching complete before verification' -ForEach @(
        @{ Bytes = 0L; Total = 100L; Percent = 0 }
        @{ Bytes = 50L; Total = 100L; Percent = 50 }
        @{ Bytes = 99L; Total = 100L; Percent = 99 }
        @{ Bytes = 100L; Total = 100L; Percent = 99 }
        @{ Bytes = 150L; Total = 100L; Percent = 99 }
        @{ Bytes = [long]::MaxValue; Total = 1L; Percent = 99 }
        @{ Bytes = [long]::MaxValue; Total = [long]::MaxValue; Percent = 99 }
        @{ Bytes = 0L; Total = 0L; Percent = 0 }
    ) {
        $context = New-PodcastProgressContext -TotalEpisodes 1 -Clock { $script:ProgressTick }
        Start-PodcastEpisodeProgress -Context $context -Index 1
        Start-PodcastTransferProgress -Context $context -TotalBytes $Total
        Update-PodcastTransferProgress -Context $context -Bytes $Bytes
        $context.Bytes | Should -Be $Bytes
        $context.TransferPercent | Should -Be $Percent
        @($script:RenderedProgress | Where-Object { $_.PercentComplete -ge 100 }).Count | Should -Be 0
    }

    It 'keeps an unknown total indeterminate and displays the exact actual byte count' {
        $context = New-PodcastProgressContext -TotalEpisodes 1 -Clock { $script:ProgressTick }
        Start-PodcastEpisodeProgress -Context $context -Index 1
        Start-PodcastTransferProgress -Context $context -TotalBytes $null -Attempt 2
        Update-PodcastTransferProgress -Context $context -Bytes ([long]::MaxValue) -Stage verifying
        $context.TotalBytes | Should -BeNullOrEmpty
        $context.TransferPercent | Should -Be -1
        $child = @($script:RenderedProgress | Where-Object { $_.Id -eq $context.TransferId })[-1]
        $child.Status | Should -BeExactly 'Validating and recording; attempt 2; 9223372036854775807 bytes (total unknown)'
        $child.PercentComplete | Should -Be -1
    }

    It 'resets retries and resumes from actual offsets without adding a received prefix twice' {
        $context = New-PodcastProgressContext -TotalEpisodes 1 -Clock { $script:ProgressTick }
        Start-PodcastEpisodeProgress -Context $context -Index 1
        Start-PodcastTransferProgress -Context $context -Offset 600 -TotalBytes 1000 -Attempt 1
        Update-PodcastTransferProgress -Context $context -Bytes 700
        $context.Bytes | Should -Be 700
        $context.TransferPercent | Should -Be 70
        Start-PodcastTransferProgress -Context $context -Offset 0 -TotalBytes 2000 -Attempt 2
        $context.Offset | Should -Be 0
        $context.Bytes | Should -Be 0
        $context.TransferPercent | Should -Be 0
        Start-PodcastTransferProgress -Context $context -Offset 1000 -TotalBytes 2000 -Attempt 3
        $context.Bytes | Should -Be 1000
        $context.TransferPercent | Should -Be 50
        $context.Attempt | Should -Be 3
    }

    It 'throttles repeated byte records while retaining current observations and forcing phase changes' {
        $context = New-PodcastProgressContext -TotalEpisodes 1 -Clock { $script:ProgressTick }
        Start-PodcastEpisodeProgress -Context $context -Index 1
        Start-PodcastTransferProgress -Context $context -TotalBytes 100
        Update-PodcastTransferProgress -Context $context -Bytes 1
        $before = $script:RenderedProgress.Count
        $script:ProgressTick = 249.0
        Update-PodcastTransferProgress -Context $context -Bytes 50
        $context.Bytes | Should -Be 50
        $script:RenderedProgress.Count | Should -Be $before
        $script:ProgressTick = 250.0
        Update-PodcastTransferProgress -Context $context -Bytes 75
        $script:RenderedProgress.Count | Should -Be ($before + 2)
        $script:ProgressTick = 251.0
        Update-PodcastTransferProgress -Context $context -Bytes 100 -Stage verifying
        $script:RenderedProgress.Count | Should -Be ($before + 4)
        $context.TransferPercent | Should -Be 99
    }

    It 'waits for episode verification and complete run coverage before rendering 100' {
        $context = New-PodcastProgressContext -TotalEpisodes 1
        Start-PodcastEpisodeProgress -Context $context -Index 1
        Start-PodcastTransferProgress -Context $context -TotalBytes 100
        Update-PodcastTransferProgress -Context $context -Bytes 100 -Stage verifying
        $context.TransferPercent | Should -Be 99
        Complete-PodcastEpisodeProgress -Context $context -Outcome downloaded
        $context.TransferPercent | Should -Be 100
        $context.EpisodePercent | Should -Be 99
        Complete-PodcastRunProgress -Context $context -Verified:$false
        $context.EpisodePercent | Should -Be 99
        Complete-PodcastRunProgress -Context $context -Verified:$true
        $context.EpisodePercent | Should -Be 100
        $context.VerifiedEpisodes | Should -Be 1
    }

    It 'does not claim full verification for <Outcome> even if all episodes were processed' -ForEach @(
        @{ Outcome = 'legacy_unverified' }, @{ Outcome = 'conflict' }, @{ Outcome = 'deferred' }
        @{ Outcome = 'failed' }, @{ Outcome = 'cancelled' }
    ) {
        $context = New-PodcastProgressContext -TotalEpisodes 1
        Start-PodcastEpisodeProgress -Context $context -Index 1
        Start-PodcastTransferProgress -Context $context -TotalBytes 100
        Update-PodcastTransferProgress -Context $context -Bytes 100 -Stage verifying
        Complete-PodcastEpisodeProgress -Context $context -Outcome $Outcome
        Complete-PodcastRunProgress -Context $context -Verified:$true
        $context.ProcessedEpisodes | Should -Be 1
        $context.VerifiedEpisodes | Should -Be 0
        $context.RunVerified | Should -BeFalse
        @($script:RenderedProgress | Where-Object { $_.PercentComplete -ge 100 }).Count | Should -Be 0
    }

    It 'retires the previous byte bar before preparing the next episode and counts completion once' {
        $context = New-PodcastProgressContext -TotalEpisodes 2
        Start-PodcastEpisodeProgress -Context $context -Index 1
        Start-PodcastTransferProgress -Context $context -TotalBytes 100
        Update-PodcastTransferProgress -Context $context -Bytes 100
        Complete-PodcastEpisodeProgress -Context $context -Outcome downloaded
        Complete-PodcastEpisodeProgress -Context $context -Outcome downloaded
        $context.ProcessedEpisodes | Should -Be 1
        $context.VerifiedEpisodes | Should -Be 1
        $before = $script:RenderedProgress.Count
        Start-PodcastEpisodeProgress -Context $context -Index 2
        $script:RenderedProgress[$before].Id | Should -Be $context.TransferId
        $script:RenderedProgress[$before].Completed | Should -BeTrue
        $script:RenderedProgress[$before + 1].CurrentOperation | Should -Be 'Episode 2 of 2'
        $context.Bytes | Should -Be 0
        $context.TotalBytes | Should -BeNullOrEmpty
        $context.TransferStarted | Should -BeFalse
    }

    It 'emits no pipeline values or publisher text from any progress mutation' {
        $context = New-PodcastProgressContext -TotalEpisodes 1
        $context | Add-Member NoteProperty Title 'privatePublisherTitle'
        $context | Add-Member NoteProperty Url 'https://user:secret@example.invalid/private?token=secret'
        $context | Add-Member NoteProperty Path 'C:\private\archive'
        @(
            Start-PodcastEpisodeProgress -Context $context -Index 1
            Start-PodcastTransferProgress -Context $context -TotalBytes 100
            Update-PodcastTransferProgress -Context $context -Bytes 100 -Stage verifying
            Complete-PodcastEpisodeProgress -Context $context -Outcome downloaded
            Complete-PodcastRunProgress -Context $context -Verified:$true
            Close-PodcastProgress -Context $context
        ).Count | Should -Be 0
        ($script:RenderedProgress | ConvertTo-Json -Depth 4) | Should -Not -Match 'privatePublisherTitle|private|secret|example\.invalid'
    }

    It 'accepts null contexts and leaves closed contexts unchanged' {
        @(
            Start-PodcastEpisodeProgress -Context $null -Index 1
            Start-PodcastTransferProgress -Context $null
            Update-PodcastTransferProgress -Context $null -Bytes 1
            Complete-PodcastEpisodeProgress -Context $null -Outcome failed
            Complete-PodcastRunProgress -Context $null
            Close-PodcastProgress -Context $null
        ).Count | Should -Be 0
        $context = New-PodcastProgressContext -TotalEpisodes 1
        Close-PodcastProgress -Context $context
        Start-PodcastEpisodeProgress -Context $context -Index 1
        Update-PodcastTransferProgress -Context $context -Bytes 100
        $context.Bytes | Should -Be 0
        $script:RenderedProgress.Count | Should -Be 0
    }
}

Describe 'A044 owned progress cleanup and host errors' -Tag 'Unit', 'A044' {
    BeforeEach {
        $script:ProgressPreferenceBeforeTest = $global:ProgressPreference
        $global:ProgressPreference = 'Continue'
        $script:ClosedProgressIds = [Collections.Generic.List[int]]::new()
        Mock Write-Progress { if ($Completed) { $script:ClosedProgressIds.Add($Id) } }
        $script:CleanupProgressContext = New-PodcastProgressContext -TotalEpisodes 1
        Start-PodcastEpisodeProgress -Context $script:CleanupProgressContext -Index 1
        Start-PodcastTransferProgress -Context $script:CleanupProgressContext -TotalBytes 10
    }
    AfterEach { $global:ProgressPreference = $script:ProgressPreferenceBeforeTest }

    It 'clears only owned IDs after the caller changes preference to <Preference>' -ForEach @(
        @{ Preference = 'SilentlyContinue' }, @{ Preference = 'Stop' }
    ) {
        $global:ProgressPreference = $Preference
        Close-PodcastProgress -Context $script:CleanupProgressContext
        @($script:ClosedProgressIds.ToArray()).Count | Should -Be 2
        $script:ClosedProgressIds[0] | Should -Be $script:CleanupProgressContext.TransferId
        $script:ClosedProgressIds[1] | Should -Be $script:CleanupProgressContext.RootId
        $global:ProgressPreference | Should -Be $Preference
        Close-PodcastProgress -Context $script:CleanupProgressContext
        $script:ClosedProgressIds.Count | Should -Be 2
    }

    It 'attempts every owned completion after a <Failure> child cleanup failure' -ForEach @(
        @{ Failure = 'ordinary' }, @{ Failure = 'cancellation' }
    ) {
        $script:ProgressCleanupFailure = $Failure
        Mock Write-Progress {
            if ($Completed) {
                $script:ClosedProgressIds.Add($Id)
                if ($Id -eq $script:CleanupProgressContext.TransferId) {
                    if ($script:ProgressCleanupFailure -eq 'cancellation') { throw [OperationCanceledException]::new('private cleanup cancellation') }
                    throw [IO.IOException]::new('private progress cleanup failure')
                }
            }
        }
        $caught = $null
        try { Close-PodcastProgress -Context $script:CleanupProgressContext }
        catch { $caught = $_ }
        if ($Failure -eq 'cancellation') { Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue }
        else { $caught | Should -BeNullOrEmpty }
        $script:ClosedProgressIds.Count | Should -Be 2
        $script:CleanupProgressContext.Closed | Should -BeTrue
    }

    It 'disables rendering after an ordinary host failure without changing transfer observations' {
        Mock Write-Progress { throw [IO.IOException]::new('private host failure') }
        $context = New-PodcastProgressContext -TotalEpisodes 1
        @(Start-PodcastEpisodeProgress -Context $context -Index 1).Count | Should -Be 0
        $context.Enabled | Should -BeFalse
        Start-PodcastTransferProgress -Context $context -TotalBytes 10
        Update-PodcastTransferProgress -Context $context -Bytes 7
        $context.Bytes | Should -Be 7
        $context.TransferPercent | Should -Be 70
        Close-PodcastProgress -Context $context
        $context.Closed | Should -BeTrue
    }

    It 'preserves typed renderer cancellation while retiring its attempted owned record' {
        Mock Write-Progress {
            if (-not $Completed) { throw [OperationCanceledException]::new('private render cancellation') }
            $script:ClosedProgressIds.Add($Id)
        }
        $context = New-PodcastProgressContext -TotalEpisodes 1
        $caught = $null
        try { Start-PodcastEpisodeProgress -Context $context -Index 1 }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        Close-PodcastProgress -Context $context
        $script:ClosedProgressIds.ToArray() | Should -Contain $context.RootId
        $script:ClosedProgressIds.ToArray() | Should -Not -Contain $context.TransferId
    }
}
