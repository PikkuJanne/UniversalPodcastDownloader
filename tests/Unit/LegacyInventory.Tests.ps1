BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Mock Read-Host { throw 'Unit tests must not prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request in unit test.' }
    . (Join-Path $script:RepositoryRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    . (Join-Path $script:RepositoryRoot 'src/LegacyInventory.ps1')
    $script:LegacyMedia = [IO.File]::ReadAllBytes((Join-Path $script:RepositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
    $script:LegacyFeedId = Get-PodcastNameHash -IdentityKey 'feed:https://feed.example.invalid/private?secret=feed-secret'
    function New-LegacyEpisode {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This fixture factory returns an in-memory object only.')]
        [CmdletBinding()]
        param([string]$Title = 'Original episode', [string]$Guid = 'publisher-secret', [AllowNull()]$PubDate = [datetime]'2026-09-01', [string]$Url = 'https://media.example.invalid/audio.mp3?token=media-secret')
        [pscustomobject]@{ Title = $Title; Guid = $Guid; AtomId = ''; PubDate = $PubDate; Url = $Url; EnclosureLength = $null }
    }
    function Get-LegacyOriginalHash {
        [CmdletBinding()]
        param([string]$Root)
        @([IO.Directory]::EnumerateFiles($Root) | Sort-Object | ForEach-Object {
            $evidence = Get-PodcastFileEvidence -Root $Root -RelativePath ([IO.Path]::GetFileName($_))
            [pscustomobject]@{ RelativePath = [IO.Path]::GetFileName($_); Bytes = $evidence.Bytes; Sha256 = $evidence.Sha256 }
        }) | ConvertTo-Json -Compress
    }
}

Describe 'A017 read-only legacy inventory and conservative matching' -Tag 'Unit' {
    BeforeEach {
        $script:LegacyRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($script:LegacyRoot)
        $script:LegacyEpisode = New-LegacyEpisode
        $script:LegacyState = [pscustomobject]@{ schema_version = 1; feed_id = $script:LegacyFeedId; feed_alias_fingerprints = @($script:LegacyFeedId); generation = 0; episodes = @() }
        $script:HistoricalName = '2026-09-01 - Original episode.mp3'
    }
    It 'inventories all immediate ordinary files while preserving every original byte and path' {
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $script:HistoricalName), $script:LegacyMedia)
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot 'unmatched.bin'), $script:LegacyMedia)
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot 'unfinished.part'), $script:LegacyMedia)
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot 'empty.mp3'), [byte[]]@())
        [IO.File]::WriteAllText((Join-Path $script:LegacyRoot 'error.mp3'), '<html>Unavailable</html>')
        $before = Get-LegacyOriginalHash -Root $script:LegacyRoot
        $stateBefore = $script:LegacyState | ConvertTo-Json -Depth 8 -Compress
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Files.Count | Should -Be 5
        $plan.Files.Sha256 | Should -HaveCount 5
        $plan.Episodes[0].Classification | Should -Be 'unverified'
        $plan.Episodes[0].SuggestedPath | Should -BeExactly $script:HistoricalName
        $plan.NeedsReview | Should -BeTrue
        (Get-LegacyOriginalHash -Root $script:LegacyRoot) | Should -BeExactly $before
        ($script:LegacyState | ConvertTo-Json -Depth 8 -Compress) | Should -BeExactly $stateBefore
        @(Get-ChildItem -LiteralPath $script:LegacyRoot -Force).Count | Should -Be 5
        Test-Path -LiteralPath (Join-Path $script:LegacyRoot '.upd') | Should -BeFalse
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
        Should -Invoke Read-Host -Times 0 -Exactly
    }
    It 'uses the historical short SHA-1 suffix when the date is absent' {
        $episode = New-LegacyEpisode -Guid 'abc' -PubDate $null
        $name = 'Original episode-a9993e36.mp3'
        (Get-PodcastHistoricalFileName -Episode $episode) | Should -BeExactly $name
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $name), $script:LegacyMedia)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($episode) -State $script:LegacyState
        $plan.Episodes[0].Candidates | Should -Contain $name
        $plan.Episodes[0].Classification | Should -Be 'unverified'
    }
    It 'uses the historical short suffix for default titles with a date' {
        $episode = New-LegacyEpisode -Title '' -Guid 'abc'
        Get-PodcastHistoricalFileName -Episode $episode | Should -BeExactly '2026-09-01 - Episode-a9993e36.mp3'
    }
    It 'reproduces the historical URL fallback and sanitization without retaining the URL' {
        $episode = New-LegacyEpisode -Title 'Title: /name?' -Guid '' -PubDate $null -Url 'abc'
        Get-PodcastHistoricalFileName -Episode $episode | Should -BeExactly 'Title_ _name_-a9993e36.mp3'
        Get-PodcastLegacyFolderName -FeedTitle 'Show: /name?' | Should -BeExactly 'Show_ _name_'
    }
    It 'recognizes the UPD-0102 full-hash filename as an unverified hint' {
        $name = New-EpisodeFileName -Episode $script:LegacyEpisode
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $name), $script:LegacyMedia)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Episodes[0].SuggestedPath | Should -BeExactly $name
        $plan.Episodes[0].Classification | Should -Be 'unverified'
    }
    It 'recognizes an unknown current filename without marking it complete' {
        $historyPlan = New-PodcastHistoryPlan -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $historyPlan[0].FileName), $script:LegacyMedia)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Episodes[0].SuggestedPath | Should -BeExactly $historyPlan[0].FileName
        $plan.Episodes[0].Classification | Should -Be 'unverified'
    }
    It 'recognizes a full old hash suffix after title date and extension changes' {
        $name = New-EpisodeFileName -Episode $script:LegacyEpisode -MaxLength 95
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $name), $script:LegacyMedia)
        $script:LegacyEpisode.Title = 'Different title'; $script:LegacyEpisode.PubDate = [datetime]'2026-09-20'
        $script:LegacyEpisode.Url = 'https://media.example.invalid/new.m4a?token=changed'
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Episodes[0].SuggestedPath | Should -BeExactly $name
        $plan.Episodes[0].Classification | Should -Be 'unverified'
    }
    It 'refuses a suggestion when repeated title and date names match two identities' {
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $script:HistoricalName), $script:LegacyMedia)
        $second = New-LegacyEpisode -Guid 'another-publisher-id'
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode, $second) -State $script:LegacyState
        $plan.Episodes.Count | Should -Be 2
        foreach ($row in $plan.Episodes) {
            $row.Classification | Should -Be 'conflict'
            $row.SuggestedPath | Should -BeNullOrEmpty
            $row.Reason | Should -Be 'multiple_episode_candidates'
        }
        $plan.Files[0].Classification | Should -Be 'conflict'
        $plan.Files[0].CandidateEpisodeIds.Count | Should -Be 2
    }
    It 'refuses a suggestion when one episode has two candidate files' {
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $script:HistoricalName), $script:LegacyMedia)
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot (New-EpisodeFileName -Episode $script:LegacyEpisode)), $script:LegacyMedia)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Episodes[0].Classification | Should -Be 'conflict'
        $plan.Episodes[0].Reason | Should -Be 'multiple_file_candidates'
        $plan.Episodes[0].Candidates.Count | Should -Be 2
        $plan.Episodes[0].SuggestedPath | Should -BeNullOrEmpty
        @($plan.Files | Where-Object { $_.Classification -eq 'conflict' }).Count | Should -Be 2
    }
    It 'keeps an unrelated playable file unverified without inventing an episode match' {
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot 'My renamed archive copy.mp3'), $script:LegacyMedia)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Files[0].Classification | Should -Be 'unverified'
        $plan.Files[0].CandidateEpisodeIds.Count | Should -Be 0
        $plan.Episodes[0].Classification | Should -Be 'missing'
        $plan.Episodes[0].Candidates.Count | Should -Be 0
    }
    It 'distinguishes recorded unchanged bytes from candidate names without mutating history' {
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $script:HistoricalName), $script:LegacyMedia)
        $planned = (New-PodcastHistoryPlan -Episodes @($script:LegacyEpisode) -State $script:LegacyState)[0]
        $record = New-PodcastEpisodeRecord -Planned $planned
        $record.relative_path = $script:HistoricalName; $record.status = 'transfer_verified'
        $evidence = Get-PodcastFileEvidence -Root $script:LegacyRoot -RelativePath $script:HistoricalName
        $record.bytes = $evidence.Bytes; $record.local_sha256 = $evidence.Sha256
        $script:LegacyState.episodes = @($record)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Files[0].Classification | Should -Be 'recorded'
        $plan.Files[0].RecordedStatus | Should -Be 'transfer_verified'
        $plan.Episodes[0].Classification | Should -Be 'recorded'
        $plan.NeedsReview | Should -BeFalse
        $record.status | Should -Be 'transfer_verified'
        [IO.File]::AppendAllText((Join-Path $script:LegacyRoot $script:HistoricalName), 'changed')
        $changed = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $changed.Files[0].Classification | Should -Be 'conflict'
        $record.status | Should -Be 'transfer_verified'
    }
    It 'never suggests a file already owned by another recorded identity' {
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $script:HistoricalName), $script:LegacyMedia)
        $old = New-LegacyEpisode -Guid 'old-id'
        $record = New-PodcastEpisodeRecord -Planned (New-PodcastHistoryPlan -Episodes @($old) -State $script:LegacyState)[0]
        $record.relative_path = $script:HistoricalName; $record.status = 'adopted'
        $evidence = Get-PodcastFileEvidence -Root $script:LegacyRoot -RelativePath $script:HistoricalName
        $record.bytes = $evidence.Bytes; $record.local_sha256 = $evidence.Sha256
        $script:LegacyState.episodes = @($record)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Files[0].RecordedEpisodeId | Should -BeExactly $record.episode_id
        $plan.Files[0].RecordedStatus | Should -Be 'adopted'
        $plan.Episodes[0].Classification | Should -Be 'conflict'
        $plan.Episodes[0].Reason | Should -Be 'path_owned_by_another_episode'
        $plan.Episodes[0].SuggestedPath | Should -BeNullOrEmpty
    }
    It 'reports recorded absent files and preserves empty array shapes' {
        $record = New-PodcastEpisodeRecord -Planned (New-PodcastHistoryPlan -Episodes @($script:LegacyEpisode) -State $script:LegacyState)[0]
        $script:LegacyState.episodes = @($record)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $plan.Episodes[0].Classification | Should -Be 'missing'
        $plan.Episodes[0].Reason | Should -Be 'recorded_path_missing'
        $plan.Files -is [array] | Should -BeTrue
        $plan.Files.Count | Should -Be 0
        $empty = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @() -State $script:LegacyState
        $empty.Episodes -is [array] | Should -BeTrue
        $empty.Episodes.Count | Should -Be 0
    }
    It 'does not expose exact feed media or publisher identity strings in the plan' {
        [IO.File]::WriteAllBytes((Join-Path $script:LegacyRoot $script:HistoricalName), $script:LegacyMedia)
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @($script:LegacyEpisode) -State $script:LegacyState
        $serialized = $plan | ConvertTo-Json -Depth 8
        $serialized | Should -Not -Match 'https://|feed-secret|media-secret|publisher-secret|Verification'
        $plan.Episodes[0].EpisodeId | Should -Match '^[0-9a-f]{64}$'
    }
    It 'requires an existing directory and never creates one for preview' {
        $missing = Join-Path $script:LegacyRoot 'missing'
        { New-PodcastLegacyPlan -Root $missing -Episodes @($script:LegacyEpisode) -State $script:LegacyState } | Should -Throw '*existing ordinary show directory*'
        Test-Path -LiteralPath $missing | Should -BeFalse
    }
    It 'does not traverse ordinary child directories or treat owned metadata as media' {
        $child = Join-Path $script:LegacyRoot '.upd'
        $null = [IO.Directory]::CreateDirectory($child)
        [IO.File]::WriteAllText((Join-Path $child 'state.json'), 'test metadata')
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @() -State $script:LegacyState
        $plan.Files.Count | Should -Be 0
        [IO.File]::ReadAllText((Join-Path $child 'state.json')) | Should -BeExactly 'test metadata'
    }
    It 'reports a reparse entry without opening or following it' {
        $script:ReparseFixture = Join-Path $script:LegacyRoot 'unsafe.mp3'
        [IO.File]::WriteAllBytes($script:ReparseFixture, $script:LegacyMedia)
        Mock Get-PodcastPathAttribute {
            if ($LiteralPath -eq $script:ReparseFixture) { return [IO.FileAttributes]::ReparsePoint }
            return [IO.File]::GetAttributes($LiteralPath)
        }
        Mock Get-PodcastLegacyFileObservation { throw 'A reparse entry must not be opened.' }
        $plan = New-PodcastLegacyPlan -Root $script:LegacyRoot -Episodes @() -State $script:LegacyState
        $plan.Files[0].Classification | Should -Be 'conflict'
        $plan.Files[0].Reason | Should -Be 'unsafe_path'
        $plan.Files[0].Sha256 | Should -BeNullOrEmpty
        Should -Invoke Get-PodcastLegacyFileObservation -Times 0 -Exactly
    }
}

Describe 'A018 local signature observations never claim a verified legacy transfer' -Tag 'Unit' {
    BeforeEach {
        $script:ObservationRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($script:ObservationRoot)
    }
    It 'exposes bounded local plausibility and digest without transfer evidence' {
        [IO.File]::WriteAllBytes((Join-Path $script:ObservationRoot 'different name.mp3'), $script:LegacyMedia)
        $observed = Get-PodcastLegacyFileObservation -Root $script:ObservationRoot -RelativePath 'different name.mp3'
        $observed.Plausible | Should -BeTrue
        $observed.Classification | Should -Be 'unverified'
        $observed.Reason | Should -Be 'local_signature_only'
        $observed.InspectedBytes | Should -BeGreaterThan 0
        $observed.InspectedBytes | Should -BeLessOrEqual 65536
        $observed.Bytes | Should -Be $script:LegacyMedia.Length
        $observed.Sha256 | Should -Match '^[0-9a-f]{64}$'
        $observed.PSObject.Properties.Name | Should -Not -Contain 'Verification'
        $observed.PSObject.Properties.Name | Should -Not -Contain 'TransferCompleted'
    }
    It 'hashes but rejects zero-byte legacy files' {
        [IO.File]::WriteAllBytes((Join-Path $script:ObservationRoot 'empty.mp3'), [byte[]]@())
        $observed = Get-PodcastLegacyFileObservation -Root $script:ObservationRoot -RelativePath 'empty.mp3'
        $observed.Bytes | Should -Be 0
        $observed.Sha256 | Should -BeExactly 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
        $observed.Plausible | Should -BeFalse
        $observed.Classification | Should -Be 'conflict'
        $observed.Reason | Should -Be 'empty_body'
    }
    It 'rejects obvious text despite an audio extension and reports its hash' {
        [IO.File]::WriteAllText((Join-Path $script:ObservationRoot 'text.mp3'), 'The download failed.')
        $observed = Get-PodcastLegacyFileObservation -Root $script:ObservationRoot -RelativePath 'text.mp3'
        $observed.Plausible | Should -BeFalse
        $observed.Reason | Should -Be 'non_audio_text'
        $observed.Sha256 | Should -Match '^[0-9a-f]{64}$'
    }
    It 'refuses partial names even when their prefix contains plausible media' {
        [IO.File]::WriteAllBytes((Join-Path $script:ObservationRoot 'unfinished.part.mp3'), $script:LegacyMedia)
        $observed = Get-PodcastLegacyFileObservation -Root $script:ObservationRoot -RelativePath 'unfinished.part.mp3'
        $observed.Plausible | Should -BeFalse
        $observed.Reason | Should -Be 'partial_name'
        $observed.Sha256 | Should -Match '^[0-9a-f]{64}$'
    }
    It 'reports an inaccessible original without changing or adopting it' {
        $path = Join-Path $script:ObservationRoot 'locked.mp3'
        [IO.File]::WriteAllBytes($path, $script:LegacyMedia)
        $held = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        try {
            $observed = Get-PodcastLegacyFileObservation -Root $script:ObservationRoot -RelativePath 'locked.mp3'
            $observed.Plausible | Should -BeFalse
            $observed.Reason | Should -Be 'file_unreadable'
            $observed.Sha256 | Should -BeNullOrEmpty
        }
        finally { $held.Dispose() }
        [IO.File]::ReadAllBytes($path).Length | Should -Be $script:LegacyMedia.Length
    }
    It 'rejects traversal and nested paths before opening any file' {
        { Get-PodcastLegacyFileObservation -Root $script:ObservationRoot -RelativePath '..\outside.mp3' } | Should -Throw '*path component*'
        { Get-PodcastLegacyFileObservation -Root $script:ObservationRoot -RelativePath 'nested\inside.mp3' } | Should -Throw '*path component*'
    }
}
