BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Mock Read-Host { throw 'Unit tests must not prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request in unit test.' }
    . (Join-Path $script:RepositoryRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Unit tests must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Unexpected metadata request in unit test.' }
    . (Join-Path $script:RepositoryRoot 'src/HistoryStore.ps1')
    . (Join-Path $script:RepositoryRoot 'src/HistoryIdentity.ps1')
    $script:FeedIdentity = Get-PodcastNameHash -IdentityKey 'feed:https://feed.example.invalid/one'
    function New-IdentityEpisode {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This fixture factory only returns an in-memory object.')]
        [CmdletBinding()]
        param([string]$Title = 'A title', [string]$Guid = 'publisher-guid', [string]$AtomId = '', [string]$Url = 'https://media.example.invalid/audio.mp3?signature=exact')
        [PSCustomObject]@{ Title = $Title; Guid = $Guid; AtomId = $AtomId; Url = $Url; PubDate = [datetime]'2026-09-01' }
    }
    function New-IdentityHistory {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This fixture factory only returns an in-memory object.')]
        [CmdletBinding()]
        param([object[]]$Records = @(), [string]$FeedId = $script:FeedIdentity)
        [PSCustomObject]@{ feed_id = $FeedId; feed_alias_fingerprints = @($FeedId); episodes = @($Records) }
    }
    function New-IdentityRecord {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This fixture factory only returns an in-memory object.')]
        [CmdletBinding()]
        param($Episode, [string]$RelativePath = 'established.mp3')
        $identity = Get-PodcastEpisodeIdentity -Episode $Episode -FeedId $script:FeedIdentity
        [PSCustomObject]@{ episode_id = $identity.Id; identity_source = $identity.Source; identity_fingerprint = $identity.Fingerprint; relative_path = $RelativePath }
    }
}

Describe 'A014 exact publisher identity and feed scoping' -Tag 'Unit' {
    It 'keeps a GUID stable across title date and signed URL changes' {
        $episode = New-IdentityEpisode
        $original = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity
        $episode.Title = 'Renamed'
        $episode.PubDate = [datetime]'2026-09-15'
        $episode.Url = 'https://different.example.invalid/new.m4a?signature=changed'
        $changed = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity
        $changed.Id | Should -BeExactly $original.Id
        $changed.Fingerprint | Should -BeExactly $original.Fingerprint
        $changed.Source | Should -Be 'rss-guid'
        $changed.Id | Should -Match '^[0-9a-f]{64}$'
    }
    It 'prefers RSS GUID then Atom ID then exact media URL' {
        $episode = New-IdentityEpisode -AtomId 'urn:atom:id'
        (Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity).Source | Should -Be 'rss-guid'
        $episode.Guid = ''
        $atom = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity
        $atom.Source | Should -Be 'atom-id'
        $episode.Url = 'https://media.example.invalid/changed.mp3?token=new'
        (Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity).Id | Should -BeExactly $atom.Id
        $episode.AtomId = ''
        (Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity).Source | Should -Be 'media-url'
    }
    It 'keeps publisher IDs exact including case and surrounding spaces' {
        $ids = @(foreach ($guid in @('ID', 'id', ' ID ')) {
            (Get-PodcastEpisodeIdentity -Episode (New-IdentityEpisode -Guid $guid) -FeedId $script:FeedIdentity).Id
        })
        @($ids | Select-Object -Unique).Count | Should -Be 3
    }
    It 'scopes an identical publisher identity to its established local feed' {
        $episode = New-IdentityEpisode
        $one = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity
        $two = Get-PodcastEpisodeIdentity -Episode $episode -FeedId ('b' * 64)
        $one.Id | Should -Not -Be $two.Id
        $one.Fingerprint | Should -BeExactly $two.Fingerprint
    }
    It 'preserves full exact request URLs and distinguishes query and encoding changes' {
        $urls = @(
            'https://media.example.invalid/Case/%2f.mp3?a=1&sig=x%2By',
            'https://media.example.invalid/Case/%2f.mp3?a=2&sig=x%2By',
            'https://media.example.invalid/case/%2f.mp3?a=1&sig=x%2By',
            'https://media.example.invalid/Case/%2F.mp3?a=1&sig=x%2By',
            'https://media.example.invalid/Case/%2f.mp3?sig=x%2By&a=1'
        )
        $identifiers = @(foreach ($url in $urls) {
            $episode = New-IdentityEpisode -Guid '' -Url $url
            $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity
            $episode.Url | Should -BeExactly $url
            $identity.Fingerprint | Should -BeExactly (Get-PodcastNameHash -IdentityKey ('media-url:' + $url))
            $identity.Id
        })
        @($identifiers | Select-Object -Unique).Count | Should -Be $urls.Count
    }
    It 'rejects absent durable identity instead of treating titles as identity' {
        $episode = New-IdentityEpisode -Guid '' -Url ''
        { Get-PodcastEpisodeIdentity -Episode $episode -FeedId $script:FeedIdentity } | Should -Throw '*requires an RSS GUID*'
    }
}

Describe 'A014 durable history destination planning' -Tag 'Unit' {
    It 'reuses the recorded path across metadata and media extension changes' {
        $episode = New-IdentityEpisode
        $record = New-IdentityRecord -Episode $episode -RelativePath 'Original title.mp3'
        $state = New-IdentityHistory -Records @($record)
        $episode.Title = 'New title'
        $episode.PubDate = [datetime]'2026-09-20'
        $episode.Url = 'https://media.example.invalid/reissued.m4a?new=signature'
        $plan = New-PodcastHistoryPlan -Episodes @($episode) -State $state
        $plan[0].FileName | Should -BeExactly 'Original title.mp3'
        $plan[0].StateRecord | Should -Be $record
        $plan[0].Episode | Should -Be $episode
        $state.episodes[0].relative_path | Should -BeExactly 'Original title.mp3'
    }
    It 'uses the full feed-scoped episode identity in new filenames' {
        $episode = New-IdentityEpisode
        $state = New-IdentityHistory
        $plan = New-PodcastHistoryPlan -Episodes @($episode) -State $state
        $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $state.feed_id
        $plan[0].FileName | Should -BeLike ('*-' + $identity.Id + '.mp3')
        $plan[0].EpisodeId | Should -BeExactly $identity.Id
        $plan[0].IdentitySource | Should -Be 'rss-guid'
        $plan[0].IdentityFingerprint | Should -BeExactly $identity.Fingerprint
        $plan[0].StateRecord | Should -BeNullOrEmpty
        $state.episodes.Count | Should -Be 0
    }
    It 'keeps array shapes and collapses only identical entries in a snapshot' {
        $episode = New-IdentityEpisode
        $state = New-IdentityHistory
        $plan = New-PodcastHistoryPlan -Episodes @($episode, $episode) -State $state
        $plan -is [array] | Should -BeTrue
        $plan.Count | Should -Be 1
        $empty = New-PodcastHistoryPlan -Episodes @() -State $state
        $empty -is [array] | Should -BeTrue
        $empty.Count | Should -Be 0
    }
    It 'rejects contradictory GUID reuse within the same snapshot' {
        $first = New-IdentityEpisode
        $second = New-IdentityEpisode -Url 'https://media.example.invalid/another.mp3'
        $state = New-IdentityHistory
        { New-PodcastHistoryPlan -Episodes @($first, $second) -State $state } | Should -Throw '*Conflicting episode metadata*'
    }
    It 'detects full-hash collisions against distinct exact publisher identities' {
        $first = New-IdentityEpisode -Guid 'one'
        $second = New-IdentityEpisode -Guid 'two'
        $state = New-IdentityHistory
        Mock Get-PodcastNameHash { 'c' * 64 }
        { New-PodcastHistoryPlan -Episodes @($first, $second) -State $state } | Should -Throw '*hash collision*'
    }
    It 'detects episode ID collisions with a different stored fingerprint' {
        $episode = New-IdentityEpisode
        $record = New-IdentityRecord -Episode $episode
        $record.identity_fingerprint = 'a' * 64
        $state = New-IdentityHistory -Records @($record)
        { New-PodcastHistoryPlan -Episodes @($episode) -State $state } | Should -Throw '*recorded fingerprint*'
    }
    It 'reserves names of every historical episode even when absent from selection' {
        $record = New-IdentityRecord -Episode (New-IdentityEpisode -Guid 'old') -RelativePath 'Known.MP3'
        $state = New-IdentityHistory -Records @($record)
        $episode = New-IdentityEpisode -Guid 'new'
        Mock New-EpisodeFileName { 'known.mp3' }
        { New-PodcastHistoryPlan -Episodes @($episode) -State $state } | Should -Throw '*case-insensitive*'
    }
    It 'rejects a case-insensitive collision between newly planned episodes' {
        $first = New-IdentityEpisode -Guid 'one'
        $second = New-IdentityEpisode -Guid 'two'
        $state = New-IdentityHistory
        Mock New-EpisodeFileName { param($Episode) if ($Episode.Guid -eq 'one') { 'name.mp3' } else { 'NAME.mp3' } }
        { New-PodcastHistoryPlan -Episodes @($first, $second) -State $state } | Should -Throw '*case-insensitive*'
    }
    It 'rejects unsafe persisted relative paths before returning a plan' {
        $episode = New-IdentityEpisode
        $state = New-IdentityHistory -Records @((New-IdentityRecord -Episode $episode -RelativePath '..\outside.mp3'))
        { New-PodcastHistoryPlan -Episodes @($episode) -State $state } | Should -Throw '*path component*'
    }
}

Describe 'A014 archive resolution uses explicit exact feed aliases' -Tag 'Unit' {
    BeforeEach {
        $script:ArchiveRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $script:ArchiveRoot
        $script:ArchiveFeedUrl = 'https://feed.example.invalid/one'
        Mock Read-PodcastHistory { $null }
    }
    It 'reuses an established directory when the feed title changes' {
        $null = New-Item -ItemType Directory -Path (Join-Path $script:ArchiveRoot 'Original folder')
        Mock Read-PodcastHistory { New-IdentityHistory }
        $archive = Resolve-PodcastArchive -Root $script:ArchiveRoot -FeedUrl $script:ArchiveFeedUrl -FeedTitle 'Different title'
        $archive.FolderName | Should -BeExactly 'Original folder'
        $archive.FeedId | Should -BeExactly $script:FeedIdentity
        $archive.State | Should -Not -BeNullOrEmpty
        @(Get-ChildItem -LiteralPath $script:ArchiveRoot).Count | Should -Be 1
    }
    It 'only accepts changed feed URLs through explicit aliases in the same state' {
        $null = New-Item -ItemType Directory -Path (Join-Path $script:ArchiveRoot 'Original folder')
        $script:MovedFeedUrl = 'https://feed.example.invalid/moved?token=exact'
        Mock Read-PodcastHistory {
            $state = New-IdentityHistory
            $state.feed_alias_fingerprints += Get-PodcastNameHash -IdentityKey ('feed:' + $script:MovedFeedUrl)
            $state
        }
        $archive = Resolve-PodcastArchive -Root $script:ArchiveRoot -FeedUrl $script:MovedFeedUrl -FeedTitle 'New title'
        $archive.FolderName | Should -BeExactly 'Original folder'
        $archive.FeedId | Should -BeExactly $script:FeedIdentity
    }
    It 'does not infer feed aliases from equal titles or similar URLs' {
        $folder = New-PodcastFolderName -FeedTitle 'Same title' -FeedUrl $script:ArchiveFeedUrl
        $null = New-Item -ItemType Directory -Path (Join-Path $script:ArchiveRoot $folder)
        Mock Read-PodcastHistory { New-IdentityHistory }
        $archive = Resolve-PodcastArchive -Root $script:ArchiveRoot -FeedUrl ($script:ArchiveFeedUrl + '?variant=2') -FeedTitle 'Same title'
        $archive.FolderName | Should -Not -Be $folder
        $archive.FeedId | Should -Not -Be $script:FeedIdentity
        $archive.State | Should -BeNullOrEmpty
        @(Get-ChildItem -LiteralPath $script:ArchiveRoot).Count | Should -Be 1
    }
    It 'rejects duplicate feed alias ownership instead of choosing a directory' {
        $null = New-Item -ItemType Directory -Path (Join-Path $script:ArchiveRoot 'one')
        $null = New-Item -ItemType Directory -Path (Join-Path $script:ArchiveRoot 'two')
        Mock Read-PodcastHistory { New-IdentityHistory }
        { Resolve-PodcastArchive -Root $script:ArchiveRoot -FeedUrl $script:ArchiveFeedUrl -FeedTitle 'title' } | Should -Throw '*multiple local archives*'
    }
    It 'rejects a proposed folder that is already bound to a different feed' {
        $null = New-Item -ItemType Directory -Path (Join-Path $script:ArchiveRoot 'Established')
        Mock Read-PodcastHistory { New-IdentityHistory -FeedId ('d' * 64) }
        Mock New-PodcastFolderName { 'Established' }
        { Resolve-PodcastArchive -Root $script:ArchiveRoot -FeedUrl $script:ArchiveFeedUrl -FeedTitle 'title' } | Should -Throw '*another feed identity*'
    }
    It 'keeps a known safe folder when new publisher title text is unsafe' {
        $null = New-Item -ItemType Directory -Path (Join-Path $script:ArchiveRoot 'Established')
        Mock Read-PodcastHistory { New-IdentityHistory }
        $archive = Resolve-PodcastArchive -Root $script:ArchiveRoot -FeedUrl $script:ArchiveFeedUrl -FeedTitle '..'
        $archive.FolderName | Should -BeExactly 'Established'
    }
    It 'preserves and reports unreadable history instead of assuming an unrelated feed' {
        $null = New-Item -ItemType Directory -Path (Join-Path $script:ArchiveRoot 'unknown')
        Mock Read-PodcastHistory { throw 'Corrupt state is preserved.' }
        { Resolve-PodcastArchive -Root $script:ArchiveRoot -FeedUrl $script:ArchiveFeedUrl -FeedTitle 'title' } | Should -Throw '*Corrupt state*'
        @(Get-ChildItem -LiteralPath $script:ArchiveRoot).Count | Should -Be 1
    }
    It 'plans a new absent archive without creating an output directory' {
        $root = Join-Path $script:ArchiveRoot 'not-created'
        $archive = Resolve-PodcastArchive -Root $root -FeedUrl $script:ArchiveFeedUrl -FeedTitle 'New feed'
        $archive.FolderName | Should -BeLike 'New feed-*'
        $archive.State | Should -BeNullOrEmpty
        Test-Path -LiteralPath $root | Should -BeFalse
    }
}

Describe 'A014 entrypoint identity boundaries' -Tag 'Unit' {
    BeforeAll {
        $script:IdentityMedia = [IO.File]::ReadAllBytes((Join-Path $script:RepositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
    }
    BeforeEach {
        Mock Write-Host {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        $identityFeedResponse = [pscustomobject]@{ Content = '' }
        $identityMedia = $script:IdentityMedia
        Mock Invoke-PodcastMetadataRequest { $identityFeedResponse }
        Mock Invoke-PodcastMediaRequest {
            $DestinationStream.Write($identityMedia, 0, $identityMedia.Length)
            [pscustomobject]@{ Completed = $true; Bytes = $identityMedia.Length; ContentLength = $identityMedia.Length; ContentType = 'audio/mpeg' }
        }
    }
    It 'rejects contradictory GUIDs even when Latest would hide the older item' {
        $identityFeedResponse.Content = '<rss><channel><title>Identity show</title><item><guid>reused</guid><title>New</title><pubDate>2026-09-02T12:00:00Z</pubDate><enclosure url="https://media.example.invalid/new.mp3"/></item><item><guid>reused</guid><title>Old</title><pubDate>2026-09-01T12:00:00Z</pubDate><enclosure url="https://media.example.invalid/old.mp3"/></item></channel></rss>'
        $output = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $result = Invoke-PodcastRun -FeedUrl 'https://feed.example.invalid/rss' -OutputPath $output -Mode Latest
        $result.ExitCode | Should -Be 1
        $result.Message | Should -BeExactly 'Conflicting episode metadata reuses one identity in this feed snapshot; no media destinations were created.'
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
        Test-Path -LiteralPath $output | Should -BeFalse
    }
    It 'binds initial history to the complete exact feed URL fingerprint' {
        $identityFeedResponse.Content = '<rss><channel><title>Identity show</title><item><guid>first</guid><title>Episode</title><enclosure url="https://media.example.invalid/audio.mp3"/></item></channel></rss>'
        $output = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $exactUrl = 'https://feed.example.invalid/Case/%2f.xml?feedId=one&amp=two&signature=A%2bB'
        $result = Invoke-PodcastRun -FeedUrl $exactUrl -OutputPath $output -Mode Latest
        $result.ExitCode | Should -Be 0
        $show = @(Get-ChildItem -LiteralPath $output -Directory)[0]
        $state = Read-PodcastHistory -Root $show.FullName
        $expected = Get-PodcastNameHash -IdentityKey ('feed:' + $exactUrl)
        $state.feed_id | Should -BeExactly $expected
        $state.feed_alias_fingerprints.Count | Should -Be 1
        $state.feed_alias_fingerprints[0] | Should -BeExactly $expected
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { $Uri -ceq $exactUrl }
        $saved = Get-Content -LiteralPath (Join-Path $show.FullName '.upd/state.json') -Raw
        $saved | Should -Not -Match 'signature=|https://|feedId='
    }
}
