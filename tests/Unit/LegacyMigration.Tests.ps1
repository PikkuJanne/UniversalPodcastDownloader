BeforeAll {
    $script:MigrationRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:MigrationRepository 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    $script:MigrationMedia = [IO.File]::ReadAllBytes((Join-Path $script:MigrationRepository 'tools/codex-handoff/fixtures/silence.mp3'))
    $script:OriginalLegacyChoice = (Get-Command Get-PodcastLegacyChoice).ScriptBlock
    $script:OriginalRecordedTransfer = (Get-Command Invoke-PodcastRecordedTransfer).ScriptBlock
    Mock Invoke-WebRequest { throw 'Unit migration tests must not request a network feed.' }
    Mock Read-Host { throw 'Unit migration tests must not prompt.' }

    function Get-MigrationSnapshot {
        param([Parameter(Mandatory)][string]$Root)
        return @(Get-ChildItem -LiteralPath $Root -Force -Recurse | Sort-Object FullName | ForEach-Object {
            $relative = $_.FullName.Substring($Root.Length)
            if ($_.PSIsContainer) { 'directory|' + $relative }
            else { 'file|' + $relative + '|' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
        }) -join "`n"
    }
}

Describe 'A017 A018: legacy migration safety boundaries' -Tag 'Unit', 'A017', 'A018' {
    BeforeEach {
        $script:MigrationRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:MigrationShow = Join-Path $script:MigrationRoot 'Show'
        $null = [IO.Directory]::CreateDirectory($script:MigrationShow)
        $script:MigrationFile = 'Owner original.mp3'
        [IO.File]::WriteAllBytes((Join-Path $script:MigrationShow $script:MigrationFile), $script:MigrationMedia)
        $script:MigrationFeed = 'https://feed.example.invalid/feed'
        $script:MigrationFeedId = Get-PodcastNameHash -IdentityKey ('feed:' + $script:MigrationFeed)
        $script:MigrationEpisodes = @(
            [pscustomobject]@{ Title = 'First'; Guid = 'first'; AtomId = ''; PubDate = [datetime]'2026-09-01'; Url = 'https://media.example.invalid/first.mp3'; EnclosureLength = $null },
            [pscustomobject]@{ Title = 'Second'; Guid = 'second'; AtomId = ''; PubDate = [datetime]'2026-09-02'; Url = 'https://media.example.invalid/second.mp3'; EnclosureLength = $null }
        )
        $script:MigrationFirstId = (Get-PodcastEpisodeIdentity -Episode $script:MigrationEpisodes[0] -FeedId $script:MigrationFeedId).Id
        $script:MigrationSecondId = (Get-PodcastEpisodeIdentity -Episode $script:MigrationEpisodes[1] -FeedId $script:MigrationFeedId).Id
        $script:MigrationDigest = (Get-PodcastFileEvidence -Root $script:MigrationShow -RelativePath $script:MigrationFile).Sha256
        $script:MigrationArguments = @{
            Root = $script:MigrationRoot; LegacyRoot = $script:MigrationShow; FeedUrl = $script:MigrationFeed
            Episodes = $script:MigrationEpisodes; Action = 'Adopt'; EpisodeId = $script:MigrationSecondId
            FileName = $script:MigrationFile; Sha256 = $script:MigrationDigest; Confirm = $false
        }
        Mock Invoke-PodcastMediaRequest { throw 'Unexpected media request.' }
    }

    It 'adopts only the explicitly selected second identity in a feed with multiple episodes' {
        $result = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $result.Outcome | Should -BeExactly 'adopted'
        $state = Read-PodcastHistory -Root $script:MigrationShow
        $state.episodes.Count | Should -Be 1
        $state.episodes[0].episode_id | Should -BeExactly $script:MigrationSecondId
        $state.episodes[0].relative_path | Should -BeExactly $script:MigrationFile
        $state.episodes[0].status | Should -BeExactly 'adopted'
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }

    It 'accepts uppercase hexadecimal digests from Windows Get-FileHash' {
        $script:MigrationArguments.Sha256 = $script:MigrationDigest.ToUpperInvariant()
        $result = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $result.Outcome | Should -BeExactly 'adopted'
        (Read-PodcastHistory -Root $script:MigrationShow).episodes[0].local_sha256 | Should -BeExactly $script:MigrationDigest
    }

    It 'rejects a changed reviewed digest before creating locks or metadata' {
        $script:MigrationArguments.Sha256 = 'e' * 64
        $before = Get-MigrationSnapshot -Root $script:MigrationRoot
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw '*digest*'
        Get-MigrationSnapshot -Root $script:MigrationRoot | Should -BeExactly $before
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }

    It 'ignores the invalid historical folder hint <Title> while preserving modern destination handling' -ForEach @(
        @{ Title = '.' },
        @{ Title = '..' },
        @{ Title = 'NUL' },
        @{ Title = 'Trailing.' }
    ) {
        $state = New-PodcastHistory -FeedId $script:MigrationFeedId -FeedAliasFingerprint $script:MigrationFeedId
        $modernRoot = Join-Path $script:MigrationRoot 'Safe modern folder'
        $before = Get-MigrationSnapshot -Root $script:MigrationRoot
        $plan = Find-PodcastLegacyReview -Root $script:MigrationRoot -ArchiveRoot $modernRoot -FeedTitle $Title `
            -Episodes $script:MigrationEpisodes -SelectedEpisodes $script:MigrationEpisodes -State $state
        $plan | Should -BeNullOrEmpty
        Get-MigrationSnapshot -Root $script:MigrationRoot | Should -BeExactly $before
        Test-Path -LiteralPath $modernRoot | Should -BeFalse
    }

    It 'rejects binding an already claimed path to another episode and preserves its current history' {
        $null = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $before = Get-MigrationSnapshot -Root $script:MigrationRoot
        $script:MigrationArguments.EpisodeId = $script:MigrationFirstId
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw '*owned by another episode*'
        Get-MigrationSnapshot -Root $script:MigrationRoot | Should -BeExactly $before
    }

    It 'rejects a selected archive belonging to another feed without changing its metadata or media' {
        $lock = Enter-PodcastHistoryLock -Root $script:MigrationShow
        try {
            $state = New-PodcastHistory -FeedId ('e' * 64) -FeedAliasFingerprint ('e' * 64)
            $state.generation = 1
            $null = Write-PodcastHistory -Lock $lock -State $state
        }
        finally { $lock.Stream.Dispose() }
        $before = Get-MigrationSnapshot -Root $script:MigrationRoot
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw '*another feed*'
        Get-MigrationSnapshot -Root $script:MigrationRoot | Should -BeExactly $before
    }

    It 'rejects a second archive claiming the same feed alias before writing the selected archive' {
        $other = Join-Path $script:MigrationRoot 'Established'
        $lock = Enter-PodcastHistoryLock -Root $other
        try {
            $state = New-PodcastHistory -FeedId $script:MigrationFeedId -FeedAliasFingerprint $script:MigrationFeedId
            $state.generation = 1
            $null = Write-PodcastHistory -Lock $lock -State $state
        }
        finally { $lock.Stream.Dispose() }
        $before = Get-MigrationSnapshot -Root $script:MigrationRoot
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw '*another archive*'
        Get-MigrationSnapshot -Root $script:MigrationRoot | Should -BeExactly $before
    }

    It 'fails when another writer owns the archive and releases its output-root lock' {
        $lock = Enter-PodcastHistoryLock -Root $script:MigrationShow
        try {
            { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw '*writer lock*'
            Test-Path -LiteralPath (Join-Path $script:MigrationShow '.upd/state.json') | Should -BeFalse
            @(Get-ChildItem -LiteralPath (Join-Path $script:MigrationShow '.upd') -Filter 'legacy-*.json').Count | Should -Be 0
            $probe = Enter-PodcastArchiveLock -Root $script:MigrationRoot
            $probe.Dispose()
        }
        finally { $lock.Stream.Dispose() }
        (Get-PodcastFileEvidence -Root $script:MigrationShow -RelativePath $script:MigrationFile).Sha256 | Should -BeExactly $script:MigrationDigest
    }

    It 'checks bytes again after the locked plan and before recording owner approval' {
        $script:LegacyChoiceCalls = 0
        Mock Get-PodcastLegacyChoice {
            param($Plan, $State, $Episodes, $FileBudget, $Action, $EpisodeId, $FileName, $Sha256, $Checkpoint)
            $selected = & $script:OriginalLegacyChoice -Plan $Plan -State $State -Episodes $Episodes -FileBudget $FileBudget -Action $Action -EpisodeId $EpisodeId -FileName $FileName -Sha256 $Sha256 -Checkpoint $Checkpoint
            $script:LegacyChoiceCalls++
            if ($script:LegacyChoiceCalls -eq 2) {
                $changed = [byte[]]$script:MigrationMedia.Clone()
                $changed[$changed.Length - 1] = $changed[$changed.Length - 1] -bxor 1
                [IO.File]::WriteAllBytes((Join-Path $script:MigrationShow $script:MigrationFile), $changed)
            }
            return $selected
        }
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw '*changed since review*'
        $script:LegacyChoiceCalls | Should -Be 2
        Test-Path -LiteralPath (Join-Path $script:MigrationShow '.upd/state.json') | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Join-Path $script:MigrationShow '.upd') -Filter 'legacy-*.json').Count | Should -Be 0
        (Get-PodcastFileEvidence -Root $script:MigrationShow -RelativePath $script:MigrationFile).Sha256 | Should -Not -Be $script:MigrationDigest
    }

    It 'retains adopted evidence and media when an explicitly requested new transfer fails' {
        $null = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $before = [IO.File]::ReadAllText((Join-Path $script:MigrationShow '.upd/state.json'))
        $script:MigrationArguments.Action = 'Redownload'
        Mock Invoke-PodcastMediaRequest { throw 'Synthetic media request failed.' }
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw '*Synthetic media request failed*'
        [IO.File]::ReadAllText((Join-Path $script:MigrationShow '.upd/state.json')) | Should -BeExactly $before
        (Get-PodcastFileEvidence -Root $script:MigrationShow -RelativePath $script:MigrationFile).Sha256 | Should -BeExactly $script:MigrationDigest
        @(Get-ChildItem -LiteralPath (Join-Path $script:MigrationShow '.upd') -Filter 'legacy-*.json').Count | Should -Be 2
        @(Get-ChildItem -LiteralPath $script:MigrationShow -Filter '.upd-*.tmp').Count | Should -Be 0
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly
    }

    It 'preserves a competing redownload destination created after planning and never reports it verified' {
        $null = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $script:MigrationArguments.Action = 'Redownload'
        $script:CompetingPath = $null
        Mock Invoke-PodcastRecordedTransfer {
            param($Context, $Planned)
            $script:CompetingPath = Join-Path $Context.Lock.Root $Planned.FileName
            & $script:OriginalRecordedTransfer -Context $Context -Planned $Planned
        }
        Mock Invoke-PodcastMediaRequest {
            param($Uri, $DestinationStream)
            $Uri | Should -BeExactly $script:MigrationEpisodes[1].Url
            $DestinationStream.Write($script:MigrationMedia, 0, $script:MigrationMedia.Length)
            [IO.File]::WriteAllText($script:CompetingPath, 'Competing owner file')
            [pscustomobject]@{ Completed = $true; Bytes = [long]$script:MigrationMedia.Length; ContentLength = [long]$script:MigrationMedia.Length; ContentType = 'audio/mpeg' }
        }
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw 'Media destination already exists; preserving it.'
        [IO.File]::ReadAllText($script:CompetingPath) | Should -BeExactly 'Competing owner file'
        (Get-PodcastFileEvidence -Root $script:MigrationShow -RelativePath $script:MigrationFile).Sha256 | Should -BeExactly $script:MigrationDigest
        (Read-PodcastHistory -Root $script:MigrationShow).episodes[0].status | Should -BeExactly 'adopted'
        @(Get-ChildItem -LiteralPath $script:MigrationShow -Filter '.upd-*.tmp').Count | Should -Be 0
    }

    It 'rejects an unsafe rollback checkpoint basename <Name> before writing' -ForEach @(
        @{ Name = '../state.json' },
        @{ Name = 'state.json.bak' },
        @{ Name = 'nested/legacy-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa.json' }
    ) {
        $null = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $script:MigrationArguments.Action = 'Rollback'
        $script:MigrationArguments.Checkpoint = $Name
        $before = Get-MigrationSnapshot -Root $script:MigrationRoot
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw '*checkpoint basename*'
        Get-MigrationSnapshot -Root $script:MigrationRoot | Should -BeExactly $before
    }

    It 'rejects a rollback checkpoint with <Case> while preserving every file' -ForEach @(
        @{ Case = 'another feed'; Change = { param($s) $s.feed_id = 'e' * 64; $s.feed_alias_fingerprints = @($s.feed_id); $s.episodes = @() } },
        @{ Case = 'a future generation'; Change = { param($s) $s.generation = 1000 } },
        @{ Case = 'different aliases'; Change = { param($s) $s.feed_alias_fingerprints += 'e' * 64 } },
        @{ Case = 'unsupported schema'; Change = { param($s) $s.schema_version = 3 } }
    ) {
        $null = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $checkpoint = Read-PodcastHistory -Root $script:MigrationShow
        & $Change $checkpoint
        $name = 'legacy-' + ('a' * 32) + '.json'
        $json = ConvertTo-Json -InputObject $checkpoint -Depth 8
        [IO.File]::WriteAllText((Join-Path $script:MigrationShow ('.upd/' + $name)), $json, (New-Object Text.UTF8Encoding($false)))
        $script:MigrationArguments.Action = 'Rollback'
        $script:MigrationArguments.Checkpoint = $name
        $before = Get-MigrationSnapshot -Root $script:MigrationRoot
        { Invoke-PodcastLegacyMigration @script:MigrationArguments } | Should -Throw
        Get-MigrationSnapshot -Root $script:MigrationRoot | Should -BeExactly $before
    }

    It 'restores version 1 checkpoint metadata into a new version 2 generation without deleting media' {
        $null = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $current = Read-PodcastHistory -Root $script:MigrationShow
        $checkpoint = New-PodcastHistory -FeedId $script:MigrationFeedId -FeedAliasFingerprint $script:MigrationFeedId
        $checkpoint.generation = 1
        $name = 'legacy-' + ('a' * 32) + '.json'
        $json = ConvertTo-Json -InputObject $checkpoint -Depth 8
        [IO.File]::WriteAllText((Join-Path $script:MigrationShow ('.upd/' + $name)), $json, (New-Object Text.UTF8Encoding($false)))
        $script:MigrationArguments.Action = 'Rollback'
        $script:MigrationArguments.Checkpoint = $name
        $result = Invoke-PodcastLegacyMigration @script:MigrationArguments
        $result.Outcome | Should -BeExactly 'metadata_restored'
        $restored = Read-PodcastHistory -Root $script:MigrationShow
        $restored.schema_version | Should -Be 2
        $restored.generation | Should -Be ($current.generation + 1)
        $restored.episodes.Count | Should -Be 0
        (Read-PodcastHistoryFile -Root $script:MigrationShow -RelativePath '.upd/state.json.bak').episodes[0].status | Should -BeExactly 'adopted'
        [IO.File]::ReadAllText((Join-Path $script:MigrationShow ('.upd/' + $name))) | Should -BeExactly $json
        (Get-PodcastFileEvidence -Root $script:MigrationShow -RelativePath $script:MigrationFile).Sha256 | Should -BeExactly $script:MigrationDigest
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }
}
