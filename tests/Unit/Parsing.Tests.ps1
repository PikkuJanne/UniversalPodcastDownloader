BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:FixtureRoot = Join-Path $script:RepositoryRoot 'tools/codex-handoff/fixtures'
    # Fail closed if import ever regresses; no unit case can contact an external host.
    Mock Read-Host { throw 'Unit tests must not prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request in unit test.' }
    . (Join-Path $script:RepositoryRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Unit tests must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Unexpected metadata request in unit test.' }

    function Read-SyntheticFixture {
        param([string]$Name)
        (Get-Content -LiteralPath (Join-Path $script:FixtureRoot $Name) -Raw -Encoding UTF8).
            Replace('{{BASE_URL}}', 'https://media.example.invalid').Replace('{{AUDIO_BYTES}}', '100')
    }
}

Describe 'Feed parsing baseline' -Tag 'Unit' {
    It 'resolves a synthetic RSS feed and extracts the enclosure, title and GUID' {
        Mock Invoke-PodcastMetadataRequest { [PSCustomObject]@{ Content = Read-SyntheticFixture 'rss-single.xml' } }
        $resolved = Resolve-PodcastItems -Feeds 'https://feed.example.invalid/rss'
        $resolved.Url | Should -Be 'https://feed.example.invalid/rss'
        @($resolved.Items).Count | Should -Be 1
        $episode = Get-EpisodeData -XmlItem $resolved.Items[0]
        $episode.Title | Should -Be 'One synthetic episode'
        $episode.Guid | Should -Be 'fixture-001'
        $episode.Url | Should -Be 'https://media.example.invalid/media/ok.mp3'
        $episode.PubDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') | Should -Be '2026-09-01 12:00:00'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://feed.example.invalid/rss' }
    }

    It 'tries the next supplied feed after malformed XML' {
        Mock Invoke-PodcastMetadataRequest {
            if ($Uri -eq 'https://feed.example.invalid/bad') {
                return [PSCustomObject]@{ Content = Read-SyntheticFixture 'xml-malformed.xml' }
            }
            [PSCustomObject]@{ Content = Read-SyntheticFixture 'rss-single.xml' }
        }
        $resolved = Resolve-PodcastItems -Feeds @('https://feed.example.invalid/bad', 'https://feed.example.invalid/good')
        $resolved.Url | Should -Be 'https://feed.example.invalid/good'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 2 -Exactly
    }

    It 'reports the current no-episodes error for an empty RSS feed' {
        Mock Invoke-PodcastMetadataRequest { [PSCustomObject]@{ Content = Read-SyntheticFixture 'rss-empty.xml' } }
        { Resolve-PodcastItems -Feeds 'https://feed.example.invalid/empty' } |
            Should -Throw '*No episodes found in the feed*'
    }

    It 'extracts Atom enclosure links through the namespace-aware fixture' {
        Mock Invoke-PodcastMetadataRequest { [PSCustomObject]@{ Content = Read-SyntheticFixture 'atom-dates.xml' } }
        $resolved = Resolve-PodcastItems -Feeds 'https://feed.example.invalid/atom'
        @($resolved.Items).Count | Should -Be 2
        $episode = Get-EpisodeData -XmlItem $resolved.Items[1]
        $episode.Title | Should -Be 'Actual newest publication'
        $episode.Url | Should -Be 'https://media.example.invalid/media/ok.mp3?id=new'
    }

    It 'leaves malformed and missing dates empty without losing enclosures' {
        [xml]$xml = Read-SyntheticFixture 'rss-date-cases.xml'
        foreach ($index in 2, 3) {
            $episode = Get-EpisodeData -XmlItem $xml.rss.channel.item[$index]
            $episode.PubDate | Should -BeNullOrEmpty
            $episode.Url | Should -Match '^https://media\.example\.invalid/media/ok\.mp3'
        }
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'resolves a relative HTML feed link against the supplied page URL' {
        $html = '<html><head><link rel="alternate" type="application/rss+xml" href="../feed.xml"></head></html>'
        Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/podcast/page' |
            Should -Be 'https://show.example.invalid/feed.xml'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'returns no candidate when the synthetic page contains no feed link' {
        Find-RssInHtml -Html (Read-SyntheticFixture 'show-not-feed.html') -BaseUrl 'https://show.example.invalid/' |
            Should -BeNullOrEmpty
    }
}

Describe 'Known parser defects: characterization only, not future acceptance' -Tag 'Unit', 'BaselineCharacterization' {
    It 'currently uses Atom updated before published (UPD-0204)' {
        [xml]$xml = Read-SyntheticFixture 'atom-dates.xml'
        $entry = $xml.SelectSingleNode('//*[local-name()="entry"][1]')
        $episode = Get-EpisodeData -XmlItem $entry
        $episode.PubDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') | Should -Be '2026-09-30 12:00:00'
        # Desired publication ordering is a future regression, not a passing assertion here.
        $episode.PubDate.Year | Should -Not -Be 2020
    }

    It 'currently chooses the first enclosure even when it is video (UPD-0205)' {
        [xml]$xml = Read-SyntheticFixture 'rss-media.xml'
        $episode = Get-EpisodeData -XmlItem $xml.rss.channel.item[0]
        $episode.Url | Should -Be 'https://video.example.invalid/trailer.mp4'
    }
}
