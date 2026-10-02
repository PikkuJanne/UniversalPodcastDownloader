BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Mock Read-Host { throw 'Unit tests must not prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request in unit test.' }
    . (Join-Path $script:RepositoryRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    . (Join-Path $script:RepositoryRoot 'src/Naming.ps1')
    function New-NamingEpisode {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This fixture factory only returns an in-memory object.')]
        [CmdletBinding()]
        param([string]$Title = 'A title', [string]$Guid = 'fixture-naming', [string]$Url = 'https://media.example.invalid/episode.mp3')
        [PSCustomObject]@{ Title = $Title; Guid = $Guid; Url = $Url; PubDate = [datetime]'2026-09-01' }
    }
}

Describe 'A008 safe deterministic Windows components' -Tag 'Unit' {
    It 'replaces forbidden punctuation and trims surrounding spaces' {
        Sanitize-ForWindowsName '  One:Two/Three?  ' | Should -Be 'One_Two_Three_'
    }
    It 'uses a valid fallback for a blank or empty component' -TestCases @(
        @{ Name = $null }, @{ Name = '' }, @{ Name = '   ' }, @{ Name = '...' }
    ) {
        param($Name)
        Sanitize-ForWindowsName $Name | Should -Be 'Untitled'
    }
    It 'removes trailing periods and spaces' {
        Sanitize-ForWindowsName 'trailing. .  ' | Should -Be 'trailing'
    }
    It 'escapes reserved devices even with extensions and mixed case' -TestCases @(
        @{ Name = 'CON' }, @{ Name = 'prn.txt' }, @{ Name = 'aUx' }, @{ Name = 'NUL.mp3' },
        @{ Name = 'COM1' }, @{ Name = 'com9.anything' }, @{ Name = 'LPT1' }, @{ Name = 'lpt9.mp3' },
        @{ Name = 'CONIN$' }, @{ Name = 'CONOUT$' }, @{ Name = 'CON .txt' }
    ) {
        param($Name)
        Sanitize-ForWindowsName $Name | Should -Be ('_' + $Name)
    }
    It 'escapes the superscript COM and LPT device names' {
        foreach ($digit in @([char]0x00B9, [char]0x00B2, [char]0x00B3)) {
            foreach ($prefix in @('COM', 'LPT')) {
                Sanitize-ForWindowsName ($prefix + $digit + '.mp3') | Should -Be ('_' + $prefix + $digit + '.mp3')
            }
        }
    }
    It 'retains similar but nonreserved names' {
        Sanitize-ForWindowsName 'COM10.mp3' | Should -Be 'COM10.mp3'
        Sanitize-ForWindowsName 'Console' | Should -Be 'Console'
    }
    It 'rejects dot and dot-dot components' -TestCases @(@{ Name = '.' }, @{ Name = '..' }, @{ Name = ' .. ' }) {
        param($Name)
        $null = $Name # Referenced by the deferred assertion below.
        { Sanitize-ForWindowsName $Name } | Should -Throw '*dot*'
    }
    It 'rejects metadata-controlled rooted or drive paths' -TestCases @(
        @{ Name = 'C:\archive' }, @{ Name = 'c:relative' }, @{ Name = '\\server\share' },
        @{ Name = '\rooted' }, @{ Name = '/rooted' }, @{ Name = '\\?\C:\archive' }
    ) {
        param($Name)
        $null = $Name # Referenced by the deferred assertion below.
        { Sanitize-ForWindowsName $Name } | Should -Throw '*Metadata cannot supply*'
    }
    It 'replaces C0, C1 and bidirectional format controls' {
        Sanitize-ForWindowsName ('one' + [char]0 + [char]31 + [char]127 + [char]159 + [char]0x202E + 'two') | Should -Be 'one_____two'
    }
    It 'normalizes Unicode canonically and preserves supplementary characters' {
        Sanitize-ForWindowsName ('cafe' + [char]0x0301) | Should -Be ('caf' + [char]0x00E9)
        $emoji = [char]::ConvertFromUtf32(0x1F3A7)
        Sanitize-ForWindowsName ($emoji + ' music') | Should -Be ($emoji + ' music')
    }
    It 'replaces malformed Unicode surrogate code units' {
        Sanitize-ForWindowsName ('a' + [char]0xD800 + 'b' + [char]0xDC00) | Should -Be 'a_b_'
    }
    It 'enforces the 255 UTF-16 unit component limit and preserves Unicode boundaries' {
        (Sanitize-ForWindowsName ('x' * 400)).Length | Should -Be 255
        $emoji = [char]::ConvertFromUtf32(0x1F3A7)
        Sanitize-ForWindowsName ('ab' + $emoji) -MaxLength 3 | Should -Be 'ab'
        Sanitize-ForWindowsName $emoji -MaxLength 1 | Should -Be '_'
    }
    It 'does not create reserved or trailing-dot names by truncation' {
        Sanitize-ForWindowsName 'CONtemporary' -MaxLength 3 | Should -Be '_CO'
        Sanitize-ForWindowsName 'abc. more' -MaxLength 4 | Should -Be 'abc'
    }
}

Describe 'A008 stable episode and feed filenames' -Tag 'Unit' {
    It 'retains date, safe title, full identity and lower-case M4A extension' {
        $episode = New-NamingEpisode -Title 'The: title' -Url 'https://media.example.invalid/episode.M4A?synthetic=1'
        New-EpisodeFileName -Episode $episode -Index 1 | Should -Match '^2026-09-01 - The_ title-[0-9a-f]{64}\.m4a$'
    }
    It 'keeps naming independent of episode order for dated and undated records' {
        $episode = New-NamingEpisode
        $first = New-EpisodeFileName -Episode $episode -Index 1
        New-EpisodeFileName -Episode $episode -Index 99 | Should -Be $first
        $episode.PubDate = $null
        $first = New-EpisodeFileName -Episode $episode -Index 1
        $first | Should -Match '^A title-[0-9a-f]{64}\.mp3$'
        New-EpisodeFileName -Episode $episode -Index 99 | Should -Be $first
    }
    It 'prefers GUID identity and retains URL query values for fallback identity' {
        $first = New-NamingEpisode -Url 'https://media.example.invalid/audio.mp3?episode=1'
        $second = New-NamingEpisode -Url 'https://media.example.invalid/audio.mp3?episode=2'
        Get-EpisodeIdentityKey $first | Should -Be (Get-EpisodeIdentityKey $second)
        $first.Guid = ''; $second.Guid = ''
        Get-EpisodeIdentityKey $first | Should -Not -Be (Get-EpisodeIdentityKey $second)
        New-EpisodeFileName $first | Should -Not -Be (New-EpisodeFileName $second)
    }
    It 'uses stable metadata instead of index when GUID and URL are absent' {
        $episode = New-NamingEpisode -Title '' -Guid '' -Url ''
        $episode.PubDate = $null
        $name = New-EpisodeFileName -Episode $episode -Index 1
        $name | Should -Match '^Episode-[0-9a-f]{64}\.mp3$'
        New-EpisodeFileName -Episode $episode -Index 9 | Should -Be $name
    }
    It 'retains the entire identity suffix after long titles are shortened' {
        $episode = New-NamingEpisode -Title ('long title ' * 100)
        $full = New-EpisodeFileName -Episode $episode -MaxLength 255
        $short = New-EpisodeFileName -Episode $episode -MaxLength 90
        $short.Length | Should -Be 90
        $short.Substring($short.Length - 69) | Should -Be $full.Substring($full.Length - 69)
    }
    It 'drops the optional date for the minimum filename budget and rejects smaller budgets' {
        $episode = New-NamingEpisode
        (New-EpisodeFileName -Episode $episode -MaxLength 70).Length | Should -Be 70
        { New-EpisodeFileName -Episode $episode -MaxLength 69 } | Should -Throw '*full episode identifier*'
    }
    It 'keeps complete Unicode characters when the episode title is shortened' {
        $episode = New-NamingEpisode -Title ([char]::ConvertFromUtf32(0x1F3A7))
        $episode.PubDate = $null
        New-EpisodeFileName -Episode $episode -MaxLength 70 | Should -Match '^_-[0-9a-f]{64}\.mp3$'
    }
    It 'rejects hostile metadata in episode and feed titles' {
        $episode = New-NamingEpisode -Title '..'
        { New-EpisodeFileName $episode } | Should -Throw '*dot*'
        { New-PodcastFolderName -FeedTitle 'C:\outside' -FeedUrl 'https://feed.example.invalid/rss' } | Should -Throw '*Metadata cannot supply*'
    }
    It 'uses distinct deterministic folders for equal titles from different feed URLs' {
        $first = New-PodcastFolderName -FeedTitle 'Same title' -FeedUrl 'https://feed.example.invalid/one'
        $second = New-PodcastFolderName -FeedTitle 'Same title' -FeedUrl 'https://feed.example.invalid/two'
        $first | Should -Match '^Same title-[0-9a-f]{64}$'
        $first | Should -Not -Be $second
        New-PodcastFolderName -FeedTitle 'Same title' -FeedUrl 'https://feed.example.invalid/one' | Should -Be $first
    }
    It 'retains the full feed suffix under folder length budgeting' {
        $feed = 'https://feed.example.invalid/rss'
        $long = New-PodcastFolderName -FeedTitle ('long ' * 100) -FeedUrl $feed
        $short = New-PodcastFolderName -FeedTitle ('long ' * 100) -FeedUrl $feed -MaxLength 66
        $long.Length | Should -BeLessOrEqual 120
        $short.Length | Should -Be 66
        $short.Substring(1) | Should -Be $long.Substring($long.Length - 65)
        { New-PodcastFolderName -FeedTitle 'title' -FeedUrl $feed -MaxLength 65 } | Should -Throw '*full feed identifier*'
    }
}

Describe 'A010 case-insensitive destination collision handling' -Tag 'Unit' {
    BeforeAll {
        [xml]$script:CollisionFixture = Get-Content -LiteralPath (Join-Path $script:RepositoryRoot 'tools/codex-handoff/fixtures/rss-collisions.xml') -Raw -Encoding UTF8
    }
    It 'separates distinct same-date same-title fixture episodes' {
        $first = Get-EpisodeData -XmlItem $script:CollisionFixture.rss.channel.item[0]
        $second = Get-EpisodeData -XmlItem $script:CollisionFixture.rss.channel.item[1]
        New-EpisodeFileName $first | Should -Not -Be (New-EpisodeFileName $second)
        (New-PodcastDestinationPlan -Episodes @($first, $second)).Count | Should -Be 2
    }
    It 'separates titles whose punctuation sanitizes to the same text' {
        $first = Get-EpisodeData -XmlItem $script:CollisionFixture.rss.channel.item[2]
        $second = Get-EpisodeData -XmlItem $script:CollisionFixture.rss.channel.item[3]
        New-EpisodeFileName $first | Should -Not -Be (New-EpisodeFileName $second)
    }
    It 'separates identities and titles that differ only in letter case' {
        $first = New-NamingEpisode -Title 'TITLE' -Guid 'GUID'
        $second = New-NamingEpisode -Title 'title' -Guid 'guid'
        $plan = New-PodcastDestinationPlan -Episodes @($first, $second)
        $plan.Count | Should -Be 2
        [StringComparer]::OrdinalIgnoreCase.Equals($plan[0].FileName, $plan[1].FileName) | Should -BeFalse
    }
    It 'deduplicates identical records while preserving array shapes' {
        $episode = New-NamingEpisode
        $plan = New-PodcastDestinationPlan -Episodes @($episode, $episode)
        $plan -is [array] | Should -BeTrue
        $plan.Count | Should -Be 1
        $plan[0].Episode | Should -Be $episode
        $empty = New-PodcastDestinationPlan -Episodes @()
        $empty -is [array] | Should -BeTrue
        $empty.Count | Should -Be 0
    }
    It 'rejects contradictory reuse of one full identity before returning a plan' {
        $first = New-NamingEpisode
        $second = New-NamingEpisode -Title 'A different title'
        { New-PodcastDestinationPlan -Episodes @($first, $second) } | Should -Throw '*Conflicting episode metadata*'
    }
    It 'detects a simulated SHA256 collision using the full identity keys' {
        Mock Get-PodcastNameHash { 'a' * 64 }
        $first = New-NamingEpisode -Guid 'first'
        $second = New-NamingEpisode -Guid 'second'
        { New-PodcastDestinationPlan -Episodes @($first, $second) } | Should -Throw '*hash collision*'
    }
    It 'rejects a simulated case-insensitive filename collision independently of hashes' {
        Mock New-EpisodeFileName {
            param($Episode)
            if ($Episode.Guid -eq 'first') { 'episode.mp3' } else { 'EPISODE.MP3' }
        }
        $first = New-NamingEpisode -Guid 'first'
        $second = New-NamingEpisode -Guid 'second'
        { New-PodcastDestinationPlan -Episodes @($first, $second) } | Should -Throw '*case-insensitive*'
    }
}
