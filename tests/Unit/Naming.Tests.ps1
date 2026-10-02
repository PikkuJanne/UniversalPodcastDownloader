BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Mock Read-Host { throw 'Unit tests must not prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request in unit test.' }
    . (Join-Path $script:RepositoryRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
}

Describe 'Episode naming baseline' -Tag 'Unit' {
    It 'replaces forbidden punctuation and trims surrounding spaces' {
        Sanitize-ForWindowsName '  One:Two/Three?  ' | Should -Be 'One_Two_Three_'
    }

    It 'returns no name for blank input' {
        Sanitize-ForWindowsName '   ' | Should -BeNullOrEmpty
    }

    It 'keeps the current date-title format and URL M4A extension' {
        $episode = [PSCustomObject]@{
            Title = 'A: title'; PubDate = [datetime]'2026-09-01'
            Url = 'https://media.example.invalid/episode.m4a?synthetic=1'; Guid = 'fixture-naming'
        }
        New-EpisodeFileName -Episode $episode -Index 1 | Should -Be '2026-09-01 - A_ title.m4a'
    }

    It 'uses a stable identity suffix for undated episodes despite index changes' {
        $episode = [PSCustomObject]@{
            Title = 'Undated'; PubDate = $null
            Url = 'https://media.example.invalid/episode.mp3'; Guid = 'fixture-naming'
        }
        $name = New-EpisodeFileName -Episode $episode -Index 1
        $name | Should -Match '^Undated-[0-9a-f]{8}\.mp3$'
        New-EpisodeFileName -Episode $episode -Index 99 | Should -Be $name
    }

    It 'gives untitled episodes with distinct GUIDs distinct fallback names' {
        $first = [PSCustomObject]@{ Title = ''; PubDate = $null; Url = $null; Guid = 'first' }
        $second = [PSCustomObject]@{ Title = ''; PubDate = $null; Url = $null; Guid = 'second' }
        $firstName = New-EpisodeFileName -Episode $first -Index 1
        $firstName | Should -Match '^Episode-[0-9a-f]{8}\.mp3$'
        New-EpisodeFileName -Episode $second -Index 2 | Should -Not -Be $firstName
    }
}

Describe 'Known destination defects: characterization only, not future acceptance' -Tag 'Unit', 'BaselineCharacterization' {
    BeforeAll {
        [xml]$script:CollisionFixture = Get-Content -LiteralPath (Join-Path $script:RepositoryRoot 'tools/codex-handoff/fixtures/rss-collisions.xml') -Raw -Encoding UTF8
    }

    It 'currently gives distinct same-date same-title episodes the same filename (UPD-0102)' {
        $first = Get-EpisodeData -XmlItem $script:CollisionFixture.rss.channel.item[0]
        $second = Get-EpisodeData -XmlItem $script:CollisionFixture.rss.channel.item[1]
        $first.Guid | Should -Not -Be $second.Guid
        New-EpisodeFileName -Episode $first -Index 1 |
            Should -Be (New-EpisodeFileName -Episode $second -Index 2)
    }

    It 'currently collides when punctuation sanitizes to the same title (UPD-0102)' {
        $first = Get-EpisodeData -XmlItem $script:CollisionFixture.rss.channel.item[2]
        $second = Get-EpisodeData -XmlItem $script:CollisionFixture.rss.channel.item[3]
        $first.Title | Should -Not -Be $second.Title
        New-EpisodeFileName -Episode $first -Index 3 |
            Should -Be (New-EpisodeFileName -Episode $second -Index 4)
    }

    It 'currently accepts Windows reserved and dot path names unchanged (UPD-0102)' {
        Sanitize-ForWindowsName 'CON' | Should -Be 'CON'
        Sanitize-ForWindowsName '..' | Should -Be '..'
        Sanitize-ForWindowsName 'trailing.' | Should -Be 'trailing.'
    }
}
