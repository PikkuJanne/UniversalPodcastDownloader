BeforeAll {
    $script:PaginationRepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Mock Read-Host { throw 'Pagination units must not prompt.' }
    Mock Invoke-WebRequest { throw 'Pagination units must not use external network.' }
    . (Join-Path $script:PaginationRepositoryRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Pagination units must not create a network client.' }
    Mock Invoke-PodcastMediaRequest { throw 'Pagination units must not request media.' }

    function Get-PaginationItem {
        param([string]$Guid = 'one', [string]$Title, [string]$Date = '2026-10-01', [string]$Url)
        if (-not $Title) { $Title = 'Episode ' + $Guid }
        if (-not $Url) { $Url = 'https://media.example.invalid/' + $Guid + '.mp3' }
        $identity = if ($Guid) { '<guid>' + [Security.SecurityElement]::Escape($Guid) + '</guid>' } else { '' }
        $publication = if ($Date) { '<pubDate>' + [Security.SecurityElement]::Escape($Date) + '</pubDate>' } else { '' }
        return ('<item><title>' + [Security.SecurityElement]::Escape($Title) + '</title>' + $identity + $publication +
            '<enclosure type="audio/mpeg" url="' + [Security.SecurityElement]::Escape($Url) + '" /></item>')
    }

    function Get-PaginationRss {
        param([string[]]$Items, [string[]]$Links = @())
        return ('<rss xmlns:atom="http://www.w3.org/2005/Atom"><channel><title>Pagination show</title>' +
            ($Links -join '') + ($Items -join '') + '</channel></rss>')
    }

    function Get-PaginationEntry {
        param([string]$Id = 'one', [string]$Title, [string]$Date = '2026-10-01', [string]$Url)
        if (-not $Title) { $Title = 'Episode ' + $Id }
        if (-not $Url) { $Url = 'https://media.example.invalid/' + $Id + '.mp3' }
        $identity = if ($Id) { '<id>' + [Security.SecurityElement]::Escape($Id) + '</id>' } else { '' }
        $publication = if ($Date) { '<published>' + [Security.SecurityElement]::Escape($Date) + '</published>' } else { '' }
        return ('<entry><title>' + [Security.SecurityElement]::Escape($Title) + '</title>' + $identity + $publication +
            '<link rel="enclosure" type="audio/mpeg" href="' + [Security.SecurityElement]::Escape($Url) + '" /></entry>')
    }

    function Get-PaginationAtom {
        param([string[]]$Entries, [string[]]$Links = @())
        return ('<feed xmlns="http://www.w3.org/2005/Atom"><title>Pagination show</title>' +
            ($Links -join '') + ($Entries -join '') + '</feed>')
    }

    function Get-PaginationSource {
        param([string]$Content, [string]$Url = 'https://feeds.example.invalid/current.xml', [string]$FinalUri)
        if (-not $FinalUri) { $FinalUri = $Url }
        return Resolve-PodcastSource -Uri $Url -Response (Get-PaginationResponse -Content $Content -FinalUri $FinalUri)
    }

    function Get-PaginationResponse {
        param([string]$Content, [string]$FinalUri)
        return [pscustomobject]@{ Content = $Content; FinalUri = [Uri]$FinalUri; ContentType = 'application/xml' }
    }
}

Describe 'A037/A038: advertised feed pagination' -Tag 'Unit', 'A037', 'A038' {
    BeforeEach {
        $script:PaginationResponses = @{}
        Mock Invoke-PodcastMetadataRequest {
            param($Uri)
            $requestKey = [string]$Uri
            if (-not $script:PaginationResponses.ContainsKey($requestKey)) { throw 'Unexpected synthetic pagination request.' }
            if ($script:PaginationResponses[$requestKey] -is [Exception]) { throw $script:PaginationResponses[$requestKey] }
            return $script:PaginationResponses[$requestKey]
        }
    }

    It 'follows an explicitly advertised second feed page and collapses exact repeated episodes' {
        $firstUrl = 'https://feeds.example.invalid/current.xml'
        $nextUrl = 'https://feeds.example.invalid/page2.xml'
        $firstItem = Get-PaginationItem -Guid 'one'
        $script:PaginationResponses[$firstUrl] = Get-PaginationResponse -FinalUri $firstUrl -Content (
            Get-PaginationRss -Items @($firstItem) -Links @('<atom:link rel="next" href="page2.xml" type="application/rss+xml" />'))
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (
            Get-PaginationRss -Items @($firstItem, (Get-PaginationItem -Guid 'two')))

        $resolved = Resolve-PodcastItems -Feeds @($firstUrl)
        $episodes = @($resolved.Items | ForEach-Object { Get-EpisodeData -XmlItem $_ })
        $episodes.Count | Should -Be 2
        ($episodes.Guid -join ',') | Should -BeExactly 'one,two'
        $resolved.Url | Should -BeExactly $firstUrl
        $resolved.Catalogue.Complete | Should -BeTrue
        $resolved.Catalogue.StopReason | Should -BeExactly 'end'
        $resolved.Catalogue.PagesFetched | Should -Be 2
        $resolved.Catalogue.ItemsSeen | Should -Be 3
        $resolved.Catalogue.DuplicateCount | Should -Be 1
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { [string]$Uri -ceq $nextUrl }
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }

    It 'supports exact advertised relation <Relation> at the <Kind> feed level' -ForEach @(
        @{ Relation = 'next'; Kind = 'RSS' }, @{ Relation = 'prev-archive'; Kind = 'RSS' },
        @{ Relation = 'http://www.iana.org/assignments/relation/next'; Kind = 'RSS' },
        @{ Relation = 'http://www.iana.org/assignments/relation/prev-archive'; Kind = 'RSS' },
        @{ Relation = 'next'; Kind = 'Atom' }, @{ Relation = 'prev-archive'; Kind = 'Atom' },
        @{ Relation = 'http://www.iana.org/assignments/relation/next'; Kind = 'Atom' },
        @{ Relation = 'http://www.iana.org/assignments/relation/prev-archive'; Kind = 'Atom' }
    ) {
        $link = '<atom:link rel="' + $Relation + '" href="older.xml" xmlns:atom="http://www.w3.org/2005/Atom" />'
        $content = if ($Kind -eq 'RSS') { Get-PaginationRss -Items @() -Links @($link) } else { Get-PaginationAtom -Entries @() -Links @($link) }
        $document = ConvertFrom-PodcastFeedXml -Content $content
        @(Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/show/current.xml')) |
            Should -BeExactly 'https://feeds.example.invalid/show/older.xml'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'ignores unsupported or inexact advertised relation <Relation>' -ForEach @(
        @{ Relation = 'previous' }, @{ Relation = 'prev' }, @{ Relation = 'next-archive' }, @{ Relation = 'last' },
        @{ Relation = 'next alternate' }, @{ Relation = 'Next' }, @{ Relation = '' },
        @{ Relation = 'https://www.iana.org/assignments/relation/next' }
    ) {
        $document = ConvertFrom-PodcastFeedXml -Content (Get-PaginationRss -Items @() -Links @(
            '<atom:link rel="' + $Relation + '" href="file:///privateUrlCanary" type="text/html" />'))
        @(Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/current.xml')).Count | Should -Be 0
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'requires the Atom link namespace and feed-level location rather than episode links' {
        $rss = '<rss xmlns:atom="http://www.w3.org/2005/Atom" xmlns:foreign="urn:foreign"><channel>' +
            '<link rel="next" href="wrong.xml" /><foreign:link rel="next" href="wrong.xml" />' +
            '<item><atom:link rel="next" href="wrong.xml" /></item></channel></rss>'
        $atom = '<feed xmlns="http://www.w3.org/2005/Atom"><entry><link rel="next" href="wrong.xml" /></entry></feed>'
        foreach ($content in @($rss, $atom)) {
            $document = ConvertFrom-PodcastFeedXml -Content $content
            @(Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/current.xml')).Count | Should -Be 0
        }
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'accepts missing or supported continuation MIME <Type>' -ForEach @(
        @{ Type = '' }, @{ Type = 'application/rss+xml' }, @{ Type = 'APPLICATION/ATOM+XML; charset=utf-8' }
    ) {
        $document = ConvertFrom-PodcastFeedXml -Content (Get-PaginationRss -Items @() -Links @(
            '<atom:link rel="next" href="older.xml" type="' + $Type + '" />'))
        @(Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/current.xml')).Count | Should -Be 1
    }

    It 'resolves inherited XML bases in ancestor order against the effective response URI' {
        $document = ConvertFrom-PodcastFeedXml -Content (
            '<rss xmlns:atom="http://www.w3.org/2005/Atom" xml:base="../archive/"><channel xml:base="season/">' +
            '<atom:link rel="next" xml:base="../" href="older.xml?part=1&amp;token=synthetic" /></channel></rss>')
        @(Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://cdn.example.invalid/show/current.xml')) |
            Should -BeExactly 'https://cdn.example.invalid/archive/older.xml?part=1&token=synthetic'
    }

    It 'deduplicates resolved continuation targets while retaining document order and distinct queries' {
        $document = ConvertFrom-PodcastFeedXml -Content (Get-PaginationRss -Items @() -Links @(
            '<atom:link rel="next" href="older.xml?part=1" />',
            '<atom:link rel="prev-archive" href="https://FEEDS.example.invalid:443/older.xml?part=1" />',
            '<atom:link rel="next" href="older.xml?part=2" />'))
        $links = @(Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/current.xml'))
        $links.Count | Should -Be 2
        ($links -join '|') | Should -BeExactly 'https://feeds.example.invalid/older.xml?part=1|https://feeds.example.invalid/older.xml?part=2'
    }

    It 'collapses fragment variants of one HTTP request target to a single continuation' {
        $document = ConvertFrom-PodcastFeedXml -Content (Get-PaginationRss -Items @() -Links @(
            '<atom:link rel="next" href="older.xml#one" />',
            '<atom:link rel="prev-archive" href="older.xml#two" />'))
        $links = @(Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/current.xml'))
        $links.Count | Should -Be 1
        ([Uri]$links[0]).GetComponents([UriComponents]::HttpRequestUrl, [UriFormat]::UriEscaped) | Should -BeExactly 'https://feeds.example.invalid/older.xml'
    }

    It 'stops resolving links once two distinct HTTP targets establish ambiguity' {
        $document = ConvertFrom-PodcastFeedXml -Content (Get-PaginationRss -Items @() -Links @(
            '<atom:link rel="next" href="one.xml" />',
            '<atom:link rel="prev-archive" href="two.xml" />',
            '<atom:link rel="next" href="file:///privateThirdTargetCanary" />'))
        $links = @(Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/current.xml'))
        $links.Count | Should -Be 2
        ($links -join '|') | Should -BeExactly 'https://feeds.example.invalid/one.xml|https://feeds.example.invalid/two.xml'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'rejects an unsafe supported continuation target <Case> without a request' -ForEach @(
        @{ Case = 'file scheme'; Target = 'file:///privateUrlCanary' },
        @{ Case = 'FTP scheme'; Target = 'ftp://feeds.example.invalid/older.xml' },
        @{ Case = 'userinfo'; Target = 'https://user:privatePasswordCanary@feeds.example.invalid/older.xml' },
        @{ Case = 'HTTPS downgrade'; Target = 'http://feeds.example.invalid/older.xml' },
        @{ Case = 'blank target'; Target = ' ' }
    ) {
        $document = ConvertFrom-PodcastFeedXml -Content (Get-PaginationRss -Items @() -Links @(
            '<atom:link rel="next" href="' + $Target + '" />'))
        { Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/current.xml') } | Should -Throw
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'rejects unsafe inherited XML bases before resolving the relative continuation' {
        $document = ConvertFrom-PodcastFeedXml -Content (
            '<rss xmlns:atom="http://www.w3.org/2005/Atom" xml:base="file:///privatePathCanary/"><channel>' +
            '<atom:link rel="next" href="older.xml" /></channel></rss>')
        { Get-PodcastFeedPageLinks -Xml $document -BaseUri ([Uri]'https://feeds.example.invalid/current.xml') } | Should -Throw
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'makes an unsupported continuation MIME an explicit incomplete catalogue' {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @(
            '<atom:link rel="next" href="privatePathCanary?token=privateQueryCanary" type="text/html" />'))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Items.Count | Should -Be 1
        $resolved.Catalogue.Complete | Should -BeFalse
        $resolved.Catalogue.StopReason | Should -BeExactly 'invalid_link'
        (Get-PodcastCatalogueMessage -Catalogue $resolved.Catalogue) | Should -Not -Match 'privatePathCanary|privateQueryCanary|feeds\.example'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'stops before requesting either of two distinct advertised continuation targets' {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @(
            '<atom:link rel="next" href="next.xml" />', '<atom:link rel="prev-archive" href="archive.xml" />'))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Catalogue.StopReason | Should -BeExactly 'ambiguous_link'
        $resolved.Catalogue.Complete | Should -BeFalse
        $resolved.Items.Count | Should -Be 1
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'follows identical advertised next and archive targets only once' {
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @(
            '<atom:link rel="next" href="older.xml" />', '<atom:link rel="prev-archive" href="older.xml" />'))
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (Get-PaginationRss -Items @((Get-PaginationItem -Guid 'two')))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Catalogue.Complete | Should -BeTrue
        $resolved.Items.Count | Should -Be 2
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
    }

    It 'uses the initial redirected feed directory while keeping the original feed identity' {
        $initial = Get-PaginationSource -Url 'https://origin.example.invalid/privateSubscription' -FinalUri 'https://cdn.example.invalid/show/current.xml' -Content (
            Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="older.xml" />'))
        $nextUrl = 'https://cdn.example.invalid/show/older.xml'
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (Get-PaginationRss -Items @((Get-PaginationItem -Guid 'two')))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Url | Should -BeExactly 'https://origin.example.invalid/privateSubscription'
        $resolved.FinalUri.AbsoluteUri | Should -BeExactly 'https://cdn.example.invalid/show/current.xml'
        $resolved.Catalogue.Complete | Should -BeTrue
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { [string]$Uri -ceq $nextUrl }
    }

    It 'traverses Atom pages and collapses repeated Atom identities in encounter order' {
        $firstEntry = Get-PaginationEntry -Id 'one'
        $initial = Get-PaginationSource -Content (Get-PaginationAtom -Entries @($firstEntry) -Links @('<link rel="prev-archive" href="older.xml" />'))
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (Get-PaginationAtom -Entries @($firstEntry, (Get-PaginationEntry -Id 'two')))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $episodes = @($resolved.Items | ForEach-Object { Get-EpisodeData -XmlItem $_ })
        ($episodes.AtomId -join ',') | Should -BeExactly 'one,two'
        $resolved.Catalogue.DuplicateCount | Should -Be 1
        $resolved.Catalogue.Complete | Should -BeTrue
    }

    It 'follows an empty initial page to its explicitly advertised accessible entries' {
        $firstUrl = 'https://feeds.example.invalid/current.xml'
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $script:PaginationResponses[$firstUrl] = Get-PaginationResponse -FinalUri $firstUrl -Content (Get-PaginationRss -Items @() -Links @('<atom:link rel="next" href="older.xml" />'))
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (Get-PaginationRss -Items @((Get-PaginationItem)))
        $resolved = Resolve-PodcastItems -Feeds @($firstUrl)
        $resolved.Items.Count | Should -Be 1
        $resolved.Catalogue.PagesFetched | Should -Be 2
        $resolved.Catalogue.Complete | Should -BeTrue
    }

    It 'detects a cycle through <Alias> before requesting it again' -ForEach @(
        @{ Alias = 'original requested URI'; Target = 'https://origin.example.invalid/request.xml' },
        @{ Alias = 'effective redirected URI'; Target = 'https://cdn.example.invalid/show/current.xml' },
        @{ Alias = 'fragment variant of effective URI'; Target = 'current.xml#another-section' }
    ) {
        $initial = Get-PaginationSource -Url 'https://origin.example.invalid/request.xml' -FinalUri 'https://cdn.example.invalid/show/current.xml' -Content (
            Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="' + $Target + '" />'))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Catalogue.StopReason | Should -BeExactly 'cycle'
        $resolved.Catalogue.Complete | Should -BeFalse
        $resolved.Catalogue.PagesFetched | Should -Be 1
        $resolved.Items.Count | Should -Be 1
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'stops a two-page cycle after retaining the distinct accessible entries' {
        $firstUrl = 'https://feeds.example.invalid/current.xml'
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="older.xml" />'))
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (
            Get-PaginationRss -Items @((Get-PaginationItem -Guid 'two')) -Links @('<atom:link rel="next" href="current.xml" />'))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Catalogue.PagesFetched | Should -Be 2
        $resolved.Catalogue.StopReason | Should -BeExactly 'cycle'
        $resolved.Items.Count | Should -Be 2
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly -ParameterFilter { [string]$Uri -ceq $firstUrl }
    }

    It 'rejects a continuation redirect to an already visited page without accepting its entries' {
        $firstUrl = 'https://feeds.example.invalid/current.xml'
        $nextUrl = 'https://feeds.example.invalid/alias.xml'
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="alias.xml" />'))
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $firstUrl -Content (Get-PaginationRss -Items @((Get-PaginationItem -Guid 'unaccepted')))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Catalogue.StopReason | Should -BeExactly 'cycle'
        $resolved.Catalogue.PagesFetched | Should -Be 2
        $resolved.Items.Count | Should -Be 1
        $resolved.Catalogue.ItemsSeen | Should -Be 1
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
    }

    It 'applies configured page limit <Limit> before another advertised request' -ForEach @(
        @{ Limit = 1; ExpectedItems = 1; ExpectedRequests = 0 },
        @{ Limit = 2; ExpectedItems = 2; ExpectedRequests = 1 }
    ) {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="older.xml" />'))
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (
            Get-PaginationRss -Items @((Get-PaginationItem -Guid 'two')) -Links @('<atom:link rel="next" href="third.xml" />'))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial -MaxPages $Limit
        $resolved.Catalogue.Complete | Should -BeFalse
        $resolved.Catalogue.StopReason | Should -BeExactly 'page_limit'
        $resolved.Catalogue.PagesFetched | Should -Be $Limit
        $resolved.Catalogue.MaxPages | Should -Be $Limit
        $resolved.Items.Count | Should -Be $ExpectedItems
        Should -Invoke Invoke-PodcastMetadataRequest -Times $ExpectedRequests -Exactly
    }

    It 'bounds default traversal to twenty pages without fetching page twenty-one' {
        foreach ($number in 1..21) {
            $url = 'https://feeds.example.invalid/page' + $number + '.xml'
            $links = if ($number -lt 21) { @('<atom:link rel="next" href="page' + ($number + 1) + '.xml" />') } else { @() }
            $script:PaginationResponses[$url] = Get-PaginationResponse -FinalUri $url -Content (
                Get-PaginationRss -Items @((Get-PaginationItem -Guid ('episode-' + $number))) -Links $links)
        }
        $resolved = Resolve-PodcastItems -Feeds @('https://feeds.example.invalid/page1.xml')
        $resolved.Catalogue.PagesFetched | Should -Be 20
        $resolved.Items.Count | Should -Be 20
        $resolved.Catalogue.StopReason | Should -BeExactly 'page_limit'
        $resolved.Catalogue.Complete | Should -BeFalse
        Should -Invoke Invoke-PodcastMetadataRequest -Times 20 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly -ParameterFilter { [string]$Uri -ceq 'https://feeds.example.invalid/page21.xml' }
    }

    It 'retains only the bounded entry prefix of an overflowing page' {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @(
            (Get-PaginationItem -Guid 'one'), (Get-PaginationItem -Guid 'two'), (Get-PaginationItem -Guid 'three')))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial -MaxItems 2
        $resolved.Items.Count | Should -Be 2
        (@($resolved.Items | ForEach-Object { (Get-EpisodeData -XmlItem $_).Guid }) -join ',') | Should -BeExactly 'one,two'
        $resolved.Catalogue.ItemsSeen | Should -Be 2
        $resolved.Catalogue.StopReason | Should -BeExactly 'item_limit'
        $resolved.Catalogue.Complete | Should -BeFalse
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'counts repeated entries against the raw-entry budget instead of permitting unbounded duplicates' {
        $item = Get-PaginationItem
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @($item, $item, (Get-PaginationItem -Guid 'two')))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial -MaxItems 2
        $resolved.Items.Count | Should -Be 1
        $resolved.Catalogue.ItemsSeen | Should -Be 2
        $resolved.Catalogue.DuplicateCount | Should -Be 1
        $resolved.Catalogue.StopReason | Should -BeExactly 'item_limit'
    }

    It 'distinguishes exact entry budget exhaustion with <Case>' -ForEach @(
        @{ Case = 'no continuation'; HasLink = $false; Complete = $true; Reason = 'end' },
        @{ Case = 'an advertised continuation'; HasLink = $true; Complete = $false; Reason = 'item_limit' }
    ) {
        $links = if ($HasLink) { @('<atom:link rel="next" href="older.xml" />') } else { @() }
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links $links)
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial -MaxItems 1
        $resolved.Catalogue.Complete | Should -Be $Complete
        $resolved.Catalogue.StopReason | Should -BeExactly $Reason
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'rejects an entire over-budget metadata page while retaining earlier accepted entries' {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="older.xml" />'))
        $laterContent = Get-PaginationRss -Items @((Get-PaginationItem -Guid 'two'), (Get-PaginationItem -Guid 'three'))
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content $laterContent
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial -MaxCharacters ($initial.Content.Length + $laterContent.Length - 1)
        $resolved.Catalogue.PagesFetched | Should -Be 2
        $resolved.Catalogue.ItemsSeen | Should -Be 1
        $resolved.Items.Count | Should -Be 1
        $resolved.Catalogue.StopReason | Should -BeExactly 'character_limit'
        $resolved.Catalogue.Complete | Should -BeFalse
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
    }

    It 'distinguishes exact character budget exhaustion with <Case>' -ForEach @(
        @{ Case = 'no continuation'; HasLink = $false; Complete = $true; Reason = 'end' },
        @{ Case = 'an advertised continuation'; HasLink = $true; Complete = $false; Reason = 'character_limit' }
    ) {
        $links = if ($HasLink) { @('<atom:link rel="next" href="older.xml" />') } else { @() }
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links $links)
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial -MaxCharacters $initial.Content.Length
        $resolved.Catalogue.Complete | Should -Be $Complete
        $resolved.Catalogue.StopReason | Should -BeExactly $Reason
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'deduplicates exact media-URL identities while keeping distinct query identities' {
        $firstItem = Get-PaginationItem -Guid '' -Title 'URL identity' -Url 'https://media.example.invalid/audio.mp3?token=one'
        $secondItem = Get-PaginationItem -Guid '' -Title 'URL identity' -Url 'https://media.example.invalid/audio.mp3?token=two'
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @($firstItem, $firstItem, $secondItem))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Items.Count | Should -Be 2
        $resolved.Catalogue.DuplicateCount | Should -Be 1
        $resolved.Catalogue.ItemsSeen | Should -Be 3
        $resolved.Catalogue.Complete | Should -BeTrue
    }

    It 'retains distinct publisher identities even when their media URL is identical' {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @(
            (Get-PaginationItem -Guid 'one' -Url 'https://media.example.invalid/shared.mp3'),
            (Get-PaginationItem -Guid 'two' -Url 'https://media.example.invalid/shared.mp3')))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Items.Count | Should -Be 2
        $resolved.Catalogue.DuplicateCount | Should -Be 0
    }

    It 'preserves the existing contradiction guard for repeated identity with changed <Field>' -ForEach @(
        @{ Field = 'title'; Changed = @{ Title = 'Changed title' } },
        @{ Field = 'media URL'; Changed = @{ Url = 'https://media.example.invalid/changed.mp3' } },
        @{ Field = 'publication instant'; Changed = @{ Date = '2026-10-02' } }
    ) {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="older.xml" />'))
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (Get-PaginationRss -Items @((Get-PaginationItem @Changed)))
        { Resolve-PodcastCatalogue -InitialResolution $initial } |
            Should -Throw 'Conflicting episode metadata reuses one identity in this feed snapshot; no media destinations were created.'
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }

    It 'detects contradictory repeated identities already present on the first page' {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem), (Get-PaginationItem -Title 'Changed title')))
        { Resolve-PodcastCatalogue -InitialResolution $initial } |
            Should -Throw 'Conflicting episode metadata reuses one identity in this feed snapshot; no media destinations were created.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }

    It 'retains earlier entries and reports a tagged later-page transport failure' {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="older.xml" />'))
        $script:PaginationResponses['https://feeds.example.invalid/older.xml'] = New-PodcastTransportException -Kind Connection -Message 'privateExceptionCanary'
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Items.Count | Should -Be 1
        $resolved.Catalogue.PagesFetched | Should -Be 1
        $resolved.Catalogue.StopReason | Should -BeExactly 'page_failed'
        $resolved.Catalogue.Complete | Should -BeFalse
        (Get-PodcastCatalogueMessage -Catalogue $resolved.Catalogue) | Should -Match 'incomplete'
        (Get-PodcastCatalogueMessage -Catalogue $resolved.Catalogue) | Should -Not -Match 'privateExceptionCanary'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
    }

    It 'retains earlier entries without HTML recursion for invalid continuation <Case>' -ForEach @(
        @{ Case = 'malformed XML'; Content = '<rss><channel></rss>' },
        @{ Case = 'unsupported XML root'; Content = '<document><item /></document>' },
        @{ Case = 'HTML with a feed-shaped link'; Content = '<html><link type="application/rss+xml" href="recursive.xml" /></html>' },
        @{ Case = 'different supported feed kind'; Content = '<feed xmlns="http://www.w3.org/2005/Atom"><entry><id>foreign</id></entry></feed>' }
    ) {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links @('<atom:link rel="next" href="older.xml" />'))
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content $Content
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Items.Count | Should -Be 1
        $resolved.Catalogue.StopReason | Should -BeExactly 'invalid_page'
        $resolved.Catalogue.Complete | Should -BeFalse
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly -ParameterFilter { [string]$Uri -like '*recursive.xml*' }
    }

    It 'fetches all advertised pages before <Mode> chooses dates without assuming page order' -ForEach @(
        @{ Mode = 'Latest'; Expected = 'newest' },
        @{ Mode = 'Custom'; Expected = 'newest,tied-one' },
        @{ Mode = 'All'; Expected = 'newest,tied-one,tied-two,undated' }
    ) {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @(
            (Get-PaginationItem -Guid 'tied-one' -Date '2026-10-01'),
            (Get-PaginationItem -Guid 'undated' -Date '')) -Links @('<atom:link rel="next" href="older.xml" />'))
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $script:PaginationResponses[$nextUrl] = Get-PaginationResponse -FinalUri $nextUrl -Content (Get-PaginationRss -Items @(
            (Get-PaginationItem -Guid 'tied-two' -Date '2026-10-01'),
            (Get-PaginationItem -Guid 'newest' -Date '2026-10-02')))
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $episodes = @($resolved.Items | ForEach-Object { Get-EpisodeData -XmlItem $_ })
        (@(Select-PodcastEpisode -Episodes $episodes -Mode $Mode -CustomCount 2).Guid -join ',') | Should -BeExactly $Expected
        $resolved.Catalogue.Complete | Should -BeTrue
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
    }

    It 'reuses a complete or partial guided resolution without fetching it again for <Case>' -ForEach @(
        @{ Case = 'complete chain'; HasLink = $false; Complete = $true },
        @{ Case = 'page-bounded chain'; HasLink = $true; Complete = $false }
    ) {
        $links = if ($HasLink) { @('<atom:link rel="next" href="older.xml" />') } else { @() }
        $firstUrl = 'https://feeds.example.invalid/current.xml'
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem)) -Links $links)
        $resolved = Resolve-PodcastItems -Feeds @($firstUrl) -InitialResolution $initial -MaxPages 1
        $reused = Resolve-PodcastItems -Feeds @($firstUrl) -InitialResolution $resolved -MaxPages 20
        [object]::ReferenceEquals($resolved, $reused) | Should -BeTrue
        $reused.Catalogue.Complete | Should -Be $Complete
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'returns stable item arrays and does not mutate the initial resolution' {
        $initial = Get-PaginationSource -Content (Get-PaginationRss -Items @((Get-PaginationItem), (Get-PaginationItem)))
        $beforeXml = $initial.Xml.OuterXml
        $initialItems = $initial.Items
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        ($resolved.Items -is [object[]]) | Should -BeTrue
        $resolved.Items.Count | Should -Be 1
        $initial.Items.Count | Should -Be 2
        [object]::ReferenceEquals($initial.Items, $initialItems) | Should -BeTrue
        $initial.Xml.OuterXml | Should -BeExactly $beforeXml
        $initial.PSObject.Properties.Name | Should -Not -Contain 'Catalogue'
        $empty = Get-PaginationSource -Content (Get-PaginationRss -Items @())
        $emptyResolved = Resolve-PodcastCatalogue -InitialResolution $empty
        ($emptyResolved.Items -is [object[]]) | Should -BeTrue
        $emptyResolved.Items.Count | Should -Be 0
        $emptyResolved.Catalogue.Complete | Should -BeTrue
    }

    It 'retains unsupported entries with whitespace-only <Field> without inventing an identity' -ForEach @(
        @{ Field = 'RSS GUID'; Kind = 'RSS' }, @{ Field = 'Atom ID'; Kind = 'Atom' }
    ) {
        $content = if ($Kind -eq 'RSS') {
            Get-PaginationRss -Items @('<item><title>Unsupported entry</title><guid>   </guid></item>', (Get-PaginationItem))
        } else {
            Get-PaginationAtom -Entries @('<entry><title>Unsupported entry</title><id>   </id></entry>', (Get-PaginationEntry))
        }
        $initial = Get-PaginationSource -Content $content
        $resolved = Resolve-PodcastCatalogue -InitialResolution $initial
        $resolved.Catalogue.Complete | Should -BeTrue
        $resolved.Items.Count | Should -Be 2
        $resolved.Catalogue.DuplicateCount | Should -Be 0
        $episodes = @($resolved.Items | ForEach-Object { Get-EpisodeData -XmlItem $_ })
        $episodes[0].Url | Should -BeNullOrEmpty
        $episodes[1].Url | Should -Not -BeNullOrEmpty
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'reports an empty initial catalogue with a failed continuation as incomplete rather than an empty complete feed' {
        $firstUrl = 'https://feeds.example.invalid/current.xml'
        $nextUrl = 'https://feeds.example.invalid/older.xml'
        $script:PaginationResponses[$firstUrl] = Get-PaginationResponse -FinalUri $firstUrl -Content (
            Get-PaginationRss -Items @() -Links @('<atom:link rel="next" href="older.xml" />'))
        $script:PaginationResponses[$nextUrl] = New-PodcastTransportException -Kind Connection -Message 'privateFailureCanary'
        { Resolve-PodcastItems -Feeds @($firstUrl) } | Should -Throw 'Feed catalogue incomplete; no accessible episodes were found.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 2 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }
}
