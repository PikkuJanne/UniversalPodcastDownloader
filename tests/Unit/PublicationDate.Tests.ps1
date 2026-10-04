BeforeAll {
    $script:PublicationRepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Mock Read-Host { throw 'Publication units must not prompt.' }
    Mock Invoke-WebRequest { throw 'Publication units must not use external network.' }
    . (Join-Path $script:PublicationRepositoryRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Publication units must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Publication units must not fetch metadata.' }
    Mock Invoke-PodcastMediaRequest { throw 'Publication units must not request media.' }

    function Invoke-PublicationCulture {
        param([string]$Culture, [scriptblock]$Body)
        $previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        $previousUICulture = [Threading.Thread]::CurrentThread.CurrentUICulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            [Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            & $Body
        }
        finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
            [Threading.Thread]::CurrentThread.CurrentUICulture = $previousUICulture
        }
    }

    function Get-PublicationTestEpisode {
        param([string]$Title, [AllowNull()]$Date)
        [pscustomobject]@{ Title = $Title; Guid = $Title; Url = 'https://media.example.invalid/' + $Title + '.mp3'; PubDate = $Date }
    }
}

Describe 'A033: existing parsing and naming desired regressions' -Tag 'Unit', 'A033' {
    It 'rejects ambiguous numeric publication dates under <Culture>' -ForEach @(
        @{ Culture = 'en-US' }, @{ Culture = 'de-DE' }, @{ Culture = 'fi-FI' }
    ) {
        $previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        $previousUICulture = [Threading.Thread]::CurrentThread.CurrentUICulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            [Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            [xml]$document = '<item><title>Synthetic</title><pubDate>01/02/2026</pubDate></item>'
            (Get-EpisodeData -XmlItem $document.DocumentElement).PubDate | Should -BeNullOrEmpty
        }
        finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
            [Threading.Thread]::CurrentThread.CurrentUICulture = $previousUICulture
        }
    }

    It 'returns a UTC DateTimeOffset publication value under <Culture>' -ForEach @(
        @{ Culture = 'en-US' }, @{ Culture = 'de-DE' }, @{ Culture = 'fi-FI' }
    ) {
        $previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        $previousUICulture = [Threading.Thread]::CurrentThread.CurrentUICulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            [Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            [xml]$document = '<item><title>Synthetic</title><pubDate>2026-09-02T00:30:00+02:00</pubDate></item>'
            $episode = Get-EpisodeData -XmlItem $document.DocumentElement
            ($episode.PubDate -is [DateTimeOffset]) | Should -BeTrue
            $episode.PubDate.Offset | Should -Be ([TimeSpan]::Zero)
            $episode.PubDate.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) | Should -Be '2026-09-01 22:30:00'
        }
        finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
            [Threading.Thread]::CurrentThread.CurrentUICulture = $previousUICulture
        }
    }

    It 'uses the UTC calendar date when an offset crosses midnight in a new filename' {
        $episode = [pscustomobject]@{
            Title = 'Synthetic'
            PubDate = [DateTimeOffset]::ParseExact('2026-09-02T00:30:00+02:00', 'yyyy-MM-ddTHH:mm:sszzz', [Globalization.CultureInfo]::InvariantCulture)
            Guid = 'synthetic-one'
            Url = 'https://media.example.invalid/one.mp3'
        }
        New-EpisodeFileName -Episode $episode -Index 1 | Should -Match '^2026-09-01 - Synthetic-[a-f0-9]{64}\.mp3$'
    }
}

Describe 'A033: invariant numeric and English publication date grammar' -Tag 'Unit', 'A033' {
    It 'parses <Case> with explicit UTC semantics under <Culture>' -ForEach @(
        foreach ($culture in @('en-US', 'de-DE', 'fi-FI')) {
            foreach ($example in @(
                @{ Case = 'ISO UTC'; Value = '2026-09-01T12:34:56Z'; Expected = '2026-09-01 12:34:56.0000000' },
                @{ Case = 'lowercase ISO markers'; Value = '2026-09-01t12:34:56z'; Expected = '2026-09-01 12:34:56.0000000' },
                @{ Case = 'date only'; Value = '2026-09-01'; Expected = '2026-09-01 00:00:00.0000000' },
                @{ Case = 'timezone-less ISO with space'; Value = '2026-09-01 12:34:56'; Expected = '2026-09-01 12:34:56.0000000' },
                @{ Case = 'fractional timezone-less ISO'; Value = '2026-09-01T12:34:56.1'; Expected = '2026-09-01 12:34:56.1000000' },
                @{ Case = 'seven fractional digits'; Value = '2026-09-01T12:34:56.1234567+02:00'; Expected = '2026-09-01 10:34:56.1234567' },
                @{ Case = 'compact positive ISO offset'; Value = '2026-09-02T00:30:00+0200'; Expected = '2026-09-01 22:30:00.0000000' },
                @{ Case = 'negative ISO offset'; Value = '2026-09-01T23:30:00-03:30'; Expected = '2026-09-02 03:00:00.0000000' },
                @{ Case = 'RSS compact numeric zone'; Value = 'Tue, 1 Sep 2026 12:34:56 +0200'; Expected = '2026-09-01 10:34:56.0000000' },
                @{ Case = 'RSS colon numeric zone without weekday or seconds'; Value = '1 Sep 2026 12:34 +02:00'; Expected = '2026-09-01 10:34:00.0000000' },
                @{ Case = 'RSS two-digit year'; Value = '1 Sep 26 12:34:56 +0200'; Expected = '2026-09-01 10:34:56.0000000' }
            )) {
                @{ Culture = $culture; Case = $example.Case; Value = $example.Value; Expected = $example.Expected }
            }
        }
    ) {
        Invoke-PublicationCulture -Culture $Culture -Body {
            $parsed = ConvertTo-PodcastPublicationDate -Value $Value
            ($parsed -is [DateTimeOffset]) | Should -BeTrue
            $parsed.Offset | Should -Be ([TimeSpan]::Zero)
            $parsed.ToString('yyyy-MM-dd HH:mm:ss.fffffff', [Globalization.CultureInfo]::InvariantCulture) | Should -BeExactly $Expected
        }
    }

    It 'parses named English zone <Zone> under <Culture>' -ForEach @(
        foreach ($culture in @('en-US', 'de-DE', 'fi-FI')) {
            foreach ($zone in @(
                @{ Zone = 'UT'; Hour = 12 }, @{ Zone = 'GMT'; Hour = 12 }, @{ Zone = 'UTC'; Hour = 12 }, @{ Zone = 'Z'; Hour = 12 },
                @{ Zone = 'EST'; Hour = 17 }, @{ Zone = 'EDT'; Hour = 16 }, @{ Zone = 'CST'; Hour = 18 }, @{ Zone = 'CDT'; Hour = 17 },
                @{ Zone = 'MST'; Hour = 19 }, @{ Zone = 'MDT'; Hour = 18 }, @{ Zone = 'PST'; Hour = 20 }, @{ Zone = 'PDT'; Hour = 19 }
            )) {
                @{ Culture = $culture; Zone = $zone.Zone; Hour = $zone.Hour }
            }
        }
    ) {
        Invoke-PublicationCulture -Culture $Culture -Body {
            $parsed = ConvertTo-PodcastPublicationDate -Value ('1 Sep 2026 12:00:00 ' + $Zone)
            ($parsed -is [DateTimeOffset]) | Should -BeTrue
            $parsed.Offset | Should -Be ([TimeSpan]::Zero)
            $parsed.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) |
                Should -BeExactly ('2026-09-01 {0:00}:00:00' -f $Hour)
        }
    }

    It 'applies the documented two-digit year boundary <Value>' -ForEach @(
        @{ Value = '1 Sep 49 12:00 GMT'; Year = 2049 },
        @{ Value = '1 Sep 50 12:00 GMT'; Year = 1950 }
    ) {
        (ConvertTo-PodcastPublicationDate -Value $Value).Year | Should -Be $Year
    }

    It 'accepts the extreme valid numeric offset <Value>' -ForEach @(
        @{ Value = '2026-09-02T00:30:00+14:00'; Expected = '2026-09-01 10:30:00' },
        @{ Value = '2026-09-01T23:30:00-1400'; Expected = '2026-09-02 13:30:00' }
    ) {
        $parsed = ConvertTo-PodcastPublicationDate -Value $Value
        $parsed.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) | Should -BeExactly $Expected
    }

    It 'returns null for <Case> under <Culture>' -ForEach @(
        foreach ($culture in @('en-US', 'de-DE', 'fi-FI')) {
            foreach ($example in @(
                @{ Case = 'null'; Value = $null }, @{ Case = 'empty'; Value = '' }, @{ Case = 'blank'; Value = '   ' },
                @{ Case = 'ambiguous slash date'; Value = '01/02/2026' },
                @{ Case = 'localized dotted date'; Value = '02.01.2026' },
                @{ Case = 'French month'; Value = '2 janvier 2026 12:00 UTC' },
                @{ Case = 'Finnish month'; Value = '2 tammikuuta 2026 12:00 UTC' },
                @{ Case = 'invalid ISO day'; Value = '2026-02-30T12:00:00Z' },
                @{ Case = 'oversized UTC offset'; Value = '2026-09-01T12:00:00+14:01' },
                @{ Case = 'eight fractional digits'; Value = '2026-09-01T12:00:00.12345678Z' },
                @{ Case = 'missing RSS zone'; Value = '1 Sep 2026 12:00:00' },
                @{ Case = 'ambiguous named zone'; Value = '1 Sep 2026 12:00 CET' },
                @{ Case = 'private malformed text'; Value = 'privatePublicationCanary 2026-09-01' }
            )) {
                @{ Culture = $culture; Case = $example.Case; Value = $example.Value }
            }
        }
    ) {
        Invoke-PublicationCulture -Culture $Culture -Body {
            ConvertTo-PodcastPublicationDate -Value $Value | Should -BeNullOrEmpty
        }
    }

    It 'bounds direct date input without echoing private content or throwing' {
        $oversizedDate = 'privatePublicationCanary' + ('x' * 257)
        { $script:BoundedDateResult = ConvertTo-PodcastPublicationDate -Value $oversizedDate } | Should -Not -Throw
        $script:BoundedDateResult | Should -BeNullOrEmpty
    }
}

Describe 'A034: Atom and RSS date selection and local provenance' -Tag 'Unit', 'A034' {
    It 'prefers Atom published to updated for a <NamespaceCase> entry' -ForEach @(
        @{ NamespaceCase = 'default namespace'; Xml = '<entry xmlns="http://www.w3.org/2005/Atom"><id>urn:synthetic:one</id><published>2020-01-01T12:00:00Z</published><updated>2026-09-30T12:00:00Z</updated></entry>' },
        @{ NamespaceCase = 'prefixed namespace'; Xml = '<a:entry xmlns:a="http://www.w3.org/2005/Atom"><a:id>urn:synthetic:one</a:id><a:published>2020-01-01T12:00:00Z</a:published><a:updated>2026-09-30T12:00:00Z</a:updated></a:entry>' }
    ) {
        [xml]$document = $Xml
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.PubDate.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) | Should -BeExactly '2020-01-01 12:00:00'
        $episode.AtomId | Should -BeExactly 'urn:synthetic:one'
        $episode.PubDateSource | Should -BeExactly 'published'
        $episode.PubDateOriginal | Should -BeExactly '2020-01-01T12:00:00Z'
    }

    It 'uses updated only when published is <Case>' -ForEach @(
        @{ Case = 'missing'; Published = '' },
        @{ Case = 'empty'; Published = '<published />' },
        @{ Case = 'whitespace'; Published = '<published>   </published>' }
    ) {
        [xml]$document = '<entry xmlns="http://www.w3.org/2005/Atom"><id>urn:synthetic:fallback</id>' + $Published + '<updated>2026-09-02T00:30:00+02:00</updated></entry>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.PubDate.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) | Should -BeExactly '2026-09-01 22:30:00'
        $episode.PubDateSource | Should -BeExactly 'updated'
        $episode.PubDateOriginal | Should -BeExactly '2026-09-02T00:30:00+02:00'
        $episode.AtomId | Should -BeExactly 'urn:synthetic:fallback'
    }

    It 'keeps an invalid present published date undated instead of substituting updated' {
        [xml]$document = '<entry xmlns="http://www.w3.org/2005/Atom"><published>privatePublicationCanary</published><updated>2026-09-01T12:00:00Z</updated></entry>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.PubDate | Should -BeNullOrEmpty
        $episode.PubDateSource | Should -BeExactly 'published'
        $episode.PubDateOriginal | Should -BeExactly 'privatePublicationCanary'
    }

    It 'ignores foreign publication dates while retaining previously collected IDs for archive compatibility' {
        [xml]$document = '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:x="urn:extension"><x:id>foreign-id</x:id><x:published>2099-01-01T00:00:00Z</x:published><updated>2026-09-01T12:00:00Z</updated></entry>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        # Changing ID collection in a date-only change would rebind existing
        # archives from their established atom-id key to a media-url key.
        $episode.AtomId | Should -BeExactly 'foreign-id'
        (Get-PodcastEpisodeIdentity -Episode $episode -FeedId ('a' * 64)).Key | Should -BeExactly 'atom-id:foreign-id'
        $episode.PubDateSource | Should -BeExactly 'updated'
        $episode.PubDate.Year | Should -Be 2026
    }

    It 'retains an RSS item ID already collected by earlier versions without rebinding identity' {
        [xml]$document = '<item><id>rss-local-id</id><pubDate>2026-09-01T12:00:00Z</pubDate><enclosure url="https://media.example.invalid/one.mp3" /></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.AtomId | Should -BeExactly 'rss-local-id'
        $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId ('a' * 64)
        $identity.Source | Should -BeExactly 'atom-id'
        $identity.Key | Should -BeExactly 'atom-id:rss-local-id'
    }

    It 'uses only the unqualified RSS pubDate and preserves its exact original string' {
        [xml]$document = '<item xmlns:x="urn:extension"><x:pubDate>2099-01-01T00:00:00Z</x:pubDate><pubDate>  Tue, 1 Sep 2026 12:00:00 +0200  </pubDate><updated>2099-01-01T00:00:00Z</updated></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.PubDateSource | Should -BeExactly 'pubDate'
        $episode.PubDateOriginal | Should -BeExactly '  Tue, 1 Sep 2026 12:00:00 +0200  '
        $episode.PubDate.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) | Should -BeExactly '2026-09-01 10:00:00'
    }

    It 'leaves an RSS item without pubDate undated despite an updated extension' {
        [xml]$document = '<item><updated>2026-09-01T12:00:00Z</updated></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.PubDate | Should -BeNullOrEmpty
        $episode.PubDateSource | Should -BeNullOrEmpty
        $episode.PubDateOriginal | Should -BeNullOrEmpty
    }

    It 'retains the old Atom updated-first date solely as a legacy filename hint' {
        [xml]$document = '<entry xmlns="http://www.w3.org/2005/Atom"><title>Synthetic</title><published>2020-01-01T12:00:00Z</published><updated>2026-09-30T12:00:00Z</updated></entry>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.PubDate.Year | Should -Be 2020
        ($episode.LegacyPubDate -is [datetime]) | Should -BeTrue
        $episode.LegacyPubDate.Year | Should -Be 2026
        Get-PodcastHistoricalFileName -Episode $episode | Should -Match '^2026-09-30 - Synthetic\.mp3$'
    }

    It 'retains the original wall calendar day in a <Case> legacy filename hint' -ForEach @(
        @{ Case = 'timezone-less timestamp'; Value = '2026-09-02T00:30:00' },
        @{ Case = 'date-only value'; Value = '2026-09-02' }
    ) {
        [xml]$document = '<item><title>Synthetic</title><pubDate>' + $Value + '</pubDate></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.LegacyPubDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) | Should -BeExactly '2026-09-02'
        Get-PodcastHistoricalFileName -Episode $episode | Should -Match '^2026-09-02 - Synthetic\.mp3$'
    }

    It 'treats an old helper object with <Case> as undated for legacy filename lookup' -ForEach @(
        @{ Case = 'empty PubDate'; HasLegacyDate = $false },
        @{ Case = 'empty PubDate and LegacyPubDate'; HasLegacyDate = $true }
    ) {
        $episode = Get-PublicationTestEpisode -Title 'Synthetic' -Date ''
        if ($HasLegacyDate) { $episode | Add-Member NoteProperty LegacyPubDate '' }
        Get-PodcastHistoricalFileName -Episode $episode | Should -Match '^Synthetic-[a-f0-9]{8}\.mp3$'
        New-EpisodeFileName -Episode $episode -Index 1 | Should -Match '^Synthetic-[a-f0-9]{64}\.mp3$'
        $episode.PubDate | Should -BeExactly ''
    }
}

Describe 'A033: stable chronological selection with UTC filename dates' -Tag 'Unit', 'A033' {
    It 'orders dated entries before undated entries with stable ties in <Mode> under <Culture>' -ForEach @(
        foreach ($culture in @('en-US', 'de-DE', 'fi-FI')) {
            foreach ($mode in @('Latest', 'Custom', 'All')) {
                @{ Culture = $culture; Mode = $mode }
            }
        }
    ) {
        Invoke-PublicationCulture -Culture $Culture -Body {
            $episodes = @(
                Get-PublicationTestEpisode -Title 'undated-first' -Date $null
                Get-PublicationTestEpisode -Title 'tie-first' -Date ([DateTimeOffset]'2026-09-01T12:00:00Z')
                Get-PublicationTestEpisode -Title 'undated-second' -Date $null
                Get-PublicationTestEpisode -Title 'newest' -Date ([DateTimeOffset]'2026-09-01T07:30:00-05:00')
                Get-PublicationTestEpisode -Title 'tie-second' -Date ([DateTimeOffset]'2026-09-01T14:00:00+02:00')
                Get-PublicationTestEpisode -Title 'oldest' -Date ([DateTimeOffset]'2026-09-01T06:00:00Z')
            )
            $originalTitles = @($episodes | ForEach-Object { $_.Title })
            $originalDates = @($episodes | ForEach-Object { $_.PubDate })
            $selected = Select-PodcastEpisode -Episodes $episodes -Mode $Mode -CustomCount 3
            ($selected -is [array]) | Should -BeTrue
            $expectedTitles = @('newest', 'tie-first', 'tie-second', 'oldest', 'undated-first', 'undated-second')
            if ($Mode -eq 'Latest') { $expectedTitles = @('newest') }
            if ($Mode -eq 'Custom') { $expectedTitles = @('newest', 'tie-first', 'tie-second') }
            @($selected | ForEach-Object { $_.Title }) -join ',' | Should -BeExactly ($expectedTitles -join ',')
            @($episodes | ForEach-Object { $_.Title }) -join ',' | Should -BeExactly ($originalTitles -join ',')
            for ($index = 0; $index -lt $episodes.Count; $index++) {
                $episodes[$index].PubDate | Should -Be $originalDates[$index]
            }
            [object]::ReferenceEquals($selected[0], $episodes[3]) | Should -BeTrue
        }
    }

    It 'preserves supplied order when every item is undated in <Mode>' -ForEach @(
        @{ Mode = 'Latest'; Expected = 'first' },
        @{ Mode = 'Custom'; Expected = 'first,second' },
        @{ Mode = 'All'; Expected = 'first,second,third' }
    ) {
        $episodes = @('first', 'second', 'third' | ForEach-Object { Get-PublicationTestEpisode -Title $_ -Date $null })
        $selected = Select-PodcastEpisode -Episodes $episodes -Mode $Mode -CustomCount 2
        @($selected | ForEach-Object { $_.Title }) -join ',' | Should -BeExactly $Expected
    }

    It 'returns a stable empty array for <Case>' -ForEach @(
        @{ Case = 'null input'; Episodes = $null },
        @{ Case = 'empty input'; Episodes = @() }
    ) {
        $selected = Select-PodcastEpisode -Episodes $Episodes -Mode All
        ($selected -is [array]) | Should -BeTrue
        $selected.Count | Should -Be 0
    }

    It 'gives equal instants the same UTC new filename and metadata date key' {
        $first = Get-PublicationTestEpisode -Title 'same' -Date ([DateTimeOffset]'2026-09-02T00:30:00+02:00')
        $second = Get-PublicationTestEpisode -Title 'same' -Date ([DateTimeOffset]'2026-09-01T22:30:00Z')
        $beforeOffset = $first.PubDate.Offset
        New-EpisodeFileName -Episode $first -Index 1 | Should -BeExactly (New-EpisodeFileName -Episode $second -Index 99)
        New-EpisodeFileName -Episode $first -Index 1 | Should -Match '^2026-09-01 - same-'
        Get-EpisodeMetadataKey -Episode $first | Should -BeExactly (Get-EpisodeMetadataKey -Episode $second)
        $first.PubDate.Offset | Should -Be $beforeOffset
    }
}
