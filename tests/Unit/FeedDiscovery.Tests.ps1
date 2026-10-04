BeforeAll {
    $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    Mock Read-Host { throw 'Discovery units must not prompt unless a test supplies choices.' }
    Mock Invoke-WebRequest { throw 'Discovery units must not use the legacy network client.' }
    . $script:DownloaderPath -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Discovery units must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Unexpected metadata request in discovery unit.' }
    Mock Invoke-PodcastMediaRequest { throw 'Discovery units must not request media.' }

    $script:DiscoveryRss = '<rss><channel><title>Synthetic show</title><item><title>One</title><guid>synthetic-one</guid><enclosure url="https://media.example.invalid/one.mp3" /></item></channel></rss>'
    $script:DiscoveryAtom = '<feed xmlns="http://www.w3.org/2005/Atom"><title>Synthetic show</title><entry><title>One</title><id>synthetic-one</id><link rel="enclosure" href="https://media.example.invalid/one.mp3" /></entry></feed>'

    function Get-DiscoveryResponse {
        param([AllowEmptyString()][string]$Content, [string]$FinalUri)
        $response = [pscustomobject]@{ Content = $Content }
        if ($FinalUri) { $response | Add-Member NoteProperty FinalUri ([Uri]$FinalUri) }
        return $response
    }
}

Describe 'A031: HTML feed discovery and effective base URI' -Tag 'Unit', 'A031' {
    It 'finds a declared feed with <Case>' -ForEach @(
        @{ Case = 'href before type'; Link = '<link href="feed.xml" type="application/rss+xml" rel="alternate">' },
        @{ Case = 'single quotes and reordered attributes'; Link = "<link href='feed.xml' rel='alternate' type='application/atom+xml'>" },
        @{ Case = 'case-insensitive HTML attributes and MIME'; Link = '<LINK HREF="feed.xml" TYPE="APPLICATION/RSS+XML" REL="ALTERNATE">' },
        @{ Case = 'whitespace around equals'; Link = "<link`n href = 'feed.xml'`t type = 'application/rss+xml' rel = 'alternate' >" },
        @{ Case = 'alternate among relation tokens'; Link = '<link rel="stylesheet alternate" href="feed.xml" type="application/rss+xml">' },
        @{ Case = 'missing relation for existing compatibility'; Link = '<link href="feed.xml" type="application/rss+xml">' },
        @{ Case = 'unquoted attribute values'; Link = '<link href=feed.xml type=application/rss+xml rel=alternate>' }
    ) {
        $candidates = @(Find-RssInHtml -Html ('<html><head>' + $Link + '</head></html>') -BaseUrl 'https://show.example.invalid/shows/index.html')
        $candidates.Count | Should -Be 1
        $candidates[0] | Should -BeExactly 'https://show.example.invalid/shows/feed.xml'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'returns RSS and Atom candidates in document order and deduplicates resolved URLs' {
        $html = @'
<html><head>
<link type="application/rss+xml" href="../rss.xml">
<link type="application/atom+xml" href="https://show.example.invalid/atom.xml">
<link type="application/rss+xml" href="https://SHOW.example.invalid:443/rss.xml">
<link type="application/atom+xml" href="../atom.xml">
<link type="application/rss+xml" href="../RSS.xml">
</head></html>
'@
        $candidates = @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/shows/index.html')
        $candidates.Count | Should -Be 3
        $candidates[0] | Should -BeExactly 'https://show.example.invalid/rss.xml'
        $candidates[1] | Should -BeExactly 'https://show.example.invalid/atom.xml'
        $candidates[2] | Should -BeExactly 'https://show.example.invalid/RSS.xml'
    }

    It 'decodes named and numeric character references before resolution and deduplication' {
        $html = '<link href="feed.xml?show=1&amp;token=synthetic&#x2D;value" type="application/rss+xml">' +
            '<link type="application/atom+xml" href="feed.xml?show=1&#38;token=synthetic-value">'
        $candidates = @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/')
        $candidates.Count | Should -Be 1
        $candidates[0] | Should -BeExactly 'https://show.example.invalid/feed.xml?show=1&token=synthetic-value'
    }

    It 'retains distinct query strings and fragments rather than merging private subscriptions' {
        $html = '<link type="application/rss+xml" href="feed.xml?token=one">' +
            '<link type="application/rss+xml" href="feed.xml?token=two">' +
            '<link type="application/rss+xml" href="feed.xml?token=one#section">'
        $candidates = @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/')
        $candidates.Count | Should -Be 3
        $candidates[0] | Should -BeExactly 'https://show.example.invalid/feed.xml?token=one'
        $candidates[1] | Should -BeExactly 'https://show.example.invalid/feed.xml?token=two'
        $candidates[2] | Should -BeExactly 'https://show.example.invalid/feed.xml?token=one#section'
    }

    It 'resolves a <Case> base href against the final response URI' -ForEach @(
        @{ Case = 'relative directory'; Base = '../feeds/'; Expected = 'https://destination.example.invalid/feeds/rss.xml' },
        @{ Case = 'absolute directory'; Base = 'https://feeds.example.invalid/show/'; Expected = 'https://feeds.example.invalid/show/rss.xml' },
        @{ Case = 'protocol-relative directory'; Base = '//feeds.example.invalid/show/'; Expected = 'https://feeds.example.invalid/show/rss.xml' }
    ) {
        $response = Get-DiscoveryResponse -Content ('<html><head><base href="' + $Base + '"><link type="application/rss+xml" href="rss.xml"></head></html>') `
            -FinalUri 'https://destination.example.invalid/shows/index.html'
        $resolved = Resolve-PodcastSource -Uri 'https://origin.example.invalid/podcast' -Response $response
        $resolved.Kind | Should -BeExactly 'Html'
        $resolved.Url | Should -BeExactly 'https://origin.example.invalid/podcast'
        $resolved.FinalUri.AbsoluteUri | Should -BeExactly 'https://destination.example.invalid/shows/index.html'
        $resolved.Candidates.Count | Should -Be 1
        $resolved.Candidates[0] | Should -BeExactly $Expected
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'uses only the first base with href and applies it to links appearing before it' {
        $html = '<html><head><base target="_blank"><link type="application/rss+xml" href="rss.xml">' +
            '<base href="../feeds/"><base href="https://other.example.invalid/ignored/"></head></html>'
        @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/shows/index.html') |
            Should -BeExactly 'https://show.example.invalid/feeds/rss.xml'
    }

    It 'uses the redirected page directory when no base is present' {
        $response = Get-DiscoveryResponse -Content '<html><link type="application/rss+xml" href="feed.xml"></html>' `
            -FinalUri 'https://destination.example.invalid/shows/index.html'
        $resolved = Resolve-PodcastSource -Uri 'https://origin.example.invalid/' -Response $response
        $resolved.Candidates[0] | Should -BeExactly 'https://destination.example.invalid/shows/feed.xml'
    }

    It 'falls back to the requested URI when a mocked response has no final URI' {
        $response = Get-DiscoveryResponse -Content '<html><link type="application/rss+xml" href="feed.xml"></html>'
        $resolved = Resolve-PodcastSource -Uri 'https://show.example.invalid/shows/index.html' -Response $response
        $resolved.FinalUri.AbsoluteUri | Should -BeExactly 'https://show.example.invalid/shows/index.html'
        $resolved.Candidates[0] | Should -BeExactly 'https://show.example.invalid/shows/feed.xml'
    }

    It 'ignores feed-shaped text in <Case>' -ForEach @(
        @{ Case = 'comments'; Prefix = '<!-- <base href="https://wrong.example.invalid/"><link type="application/rss+xml" href="false.xml"> -->' },
        @{ Case = 'scripts'; Prefix = '<script>var text = ''<base href="https://wrong.example.invalid/"><link type="application/rss+xml" href="false.xml">'';</script>' },
        @{ Case = 'styles'; Prefix = '<style>.example::after { content: ''<link type="application/rss+xml" href="false.xml">''; }</style>' },
        @{ Case = 'textarea'; Prefix = '<textarea><link type="application/rss+xml" href="false.xml"></textarea>' }
    ) {
        $html = '<html><head>' + $Prefix + '<link type="application/rss+xml" href="true.xml"></head></html>'
        $candidates = @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/')
        $candidates.Count | Should -Be 1
        $candidates[0] | Should -BeExactly 'https://show.example.invalid/true.xml'
    }

    It 'does not mistake quoted markup inside another tag for a link' {
        $html = '<html><div data-example=''<link type="application/rss+xml" href="false.xml">''></div>' +
            '<link type="application/rss+xml" href="true.xml"></html>'
        $candidates = @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/')
        $candidates.Count | Should -Be 1
        $candidates[0] | Should -BeExactly 'https://show.example.invalid/true.xml'
    }

    It 'does not discover markup swallowed by a malformed outer tag <Case>' -ForEach @(
        @{ Case = 'unterminated quoted attribute'; Html = '<html><div data-example="<link type=application/rss+xml href=false.xml>' },
        @{ Case = 'unquoted nested opening tag'; Html = '<html><div data-example=<link type=application/rss+xml href=false.xml>' }
    ) {
        @(Find-RssInHtml -Html $Html -BaseUrl 'https://show.example.invalid/').Count | Should -Be 0
    }

    It 'continues after a malformed closed outer tag without promoting its inner markup' {
        $html = '<html><div data-example=<link type=application/rss+xml href=false.xml>' +
            '<link type="application/rss+xml" href="true.xml"></html>'
        $candidates = @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/')
        $candidates.Count | Should -Be 1
        $candidates[0] | Should -BeExactly 'https://show.example.invalid/true.xml'
    }

    It 'ignores base and feed markup in nested inert templates and disabled scripting content' {
        $html = '<html><template><template><base href="https://wrong.example.invalid/">' +
            '<link type="application/rss+xml" href="false.xml"></template></template>' +
            '<noscript><link type="application/rss+xml" href="other-false.xml"></noscript>' +
            '<link type="application/rss+xml" href="true.xml"></html>'
        $candidates = @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/')
        $candidates.Count | Should -Be 1
        $candidates[0] | Should -BeExactly 'https://show.example.invalid/true.xml'
    }

    It 'ignores all later markup after the plaintext element' {
        $html = '<html><plaintext><link type="application/rss+xml" href="false.xml"></plaintext>' +
            '<link type="application/rss+xml" href="also-false.xml"></html>'
        @(Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/').Count | Should -Be 0
    }

    It 'ignores links with <Case>' -ForEach @(
        @{ Case = 'an unsupported MIME type'; Link = '<link rel="alternate" type="text/html" href="page.html">' },
        @{ Case = 'a relation without alternate'; Link = '<link rel="stylesheet" type="application/rss+xml" href="feed.xml">' },
        @{ Case = 'an attribute name that merely ends in type'; Link = '<link data-type="application/rss+xml" href="feed.xml">' },
        @{ Case = 'an attribute name that merely ends in href'; Link = '<link type="application/rss+xml" data-href="feed.xml">' },
        @{ Case = 'no href'; Link = '<link rel="alternate" type="application/rss+xml">' }
    ) {
        @(Find-RssInHtml -Html ('<html>' + $Link + '</html>') -BaseUrl 'https://show.example.invalid/').Count | Should -Be 0
    }

    It 'rejects an unsafe discovered <Case> with a fixed private diagnostic' -ForEach @(
        @{ Case = 'file scheme'; Target = 'file:///C:/privateDiscoveryCanary.xml' },
        @{ Case = 'javascript scheme'; Target = 'javascript:privateDiscoveryCanary' },
        @{ Case = 'credentials'; Target = 'https://privateUser:privateDiscoveryCanary@show.example.invalid/feed' },
        @{ Case = 'backslash'; Target = '\\show.example.invalid\privateDiscoveryCanary' },
        @{ Case = 'space'; Target = 'feed privateDiscoveryCanary.xml' },
        @{ Case = 'entity-encoded newline'; Target = 'feed&#10;privateDiscoveryCanary.xml' }
    ) {
        $html = '<html><link type="application/rss+xml" href="' + $Target + '"></html>'
        { Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/' } |
            Should -Throw 'Discovered feed URL is not allowed by the network policy.'
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
    }

    It 'rejects an unsafe first base instead of trusting a later base' {
        $html = '<html><base href="file:///C:/privateBaseCanary/"><base href="https://show.example.invalid/">' +
            '<link type="application/rss+xml" href="feed.xml"></html>'
        { Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/' } |
            Should -Throw 'HTML base URL is not allowed by the network policy.'
    }

    It 'enforces the HTML input limit before discovery or requests' {
        { Find-RssInHtml -Html ('x' * 8388609) -BaseUrl 'https://show.example.invalid/' } |
            Should -Throw 'HTML metadata exceeds the safe character limit.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }
}

Describe 'A032: actual feed root classification and useful diagnostics' -Tag 'Unit', 'A032' {
    It 'classifies a supported <Kind> root and keeps a stable item array' -ForEach @(
        @{ Kind = 'Rss'; Content = '<rss><channel><item><title>One</title></item></channel></rss>'; Count = 1 },
        @{ Kind = 'Rss'; Content = '<?xml version="1.0"?><!-- preface --><rss><channel><item /><item /></channel></rss>'; Count = 2 },
        @{ Kind = 'Atom'; Content = '<feed xmlns="http://www.w3.org/2005/Atom"><entry><title>One</title></entry></feed>'; Count = 1 },
        @{ Kind = 'Atom'; Content = '<a:feed xmlns:a="http://www.w3.org/2005/Atom"><a:entry><a:title>One</a:title></a:entry></a:feed>'; Count = 1 }
    ) {
        $response = Get-DiscoveryResponse -Content $Content -FinalUri 'https://cdn.example.invalid/feed.xml'
        $resolved = Resolve-PodcastSource -Uri 'https://show.example.invalid/feed?private=synthetic' -Response $response
        $resolved.Kind | Should -BeExactly $Kind
        $resolved.Url | Should -BeExactly 'https://show.example.invalid/feed?private=synthetic'
        $resolved.Content | Should -BeExactly $Content
        ($resolved.Xml -is [Xml.XmlDocument]) | Should -BeTrue
        ($resolved.Items -is [array]) | Should -BeTrue
        $resolved.Items.Count | Should -Be $Count
        ($resolved.Candidates -is [string[]]) | Should -BeTrue
        $resolved.Candidates.Count | Should -Be 1
        $resolved.Candidates[0] | Should -BeExactly $resolved.Url
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'classifies an empty supported <Kind> root without treating it as invalid XML' -ForEach @(
        @{ Kind = 'Rss'; Content = '<rss><channel><title>Empty</title></channel></rss>' },
        @{ Kind = 'Atom'; Content = '<feed xmlns="http://www.w3.org/2005/Atom"><title>Empty</title></feed>' }
    ) {
        $resolved = Resolve-PodcastSource -Uri 'https://show.example.invalid/empty' -Response (Get-DiscoveryResponse -Content $Content)
        $resolved.Kind | Should -BeExactly $Kind
        ($resolved.Items -is [array]) | Should -BeTrue
        $resolved.Items.Count | Should -Be 0
    }

    It 'requires a supported root for <Case>' -ForEach @(
        @{ Case = 'nested RSS'; Content = '<wrapper><rss><channel><item /></channel></rss></wrapper>' },
        @{ Case = 'nested Atom'; Content = '<wrapper><feed xmlns="http://www.w3.org/2005/Atom"><entry /></feed></wrapper>' },
        @{ Case = 'RSS in an unrelated namespace'; Content = '<rss xmlns="urn:unrelated"><channel><item /></channel></rss>' },
        @{ Case = 'Atom without the Atom namespace'; Content = '<feed><entry /></feed>' },
        @{ Case = 'Atom in an unrelated namespace'; Content = '<feed xmlns="urn:unrelated"><entry /></feed>' },
        @{ Case = 'root name case mismatch'; Content = '<RSS><channel><item /></channel></RSS>' },
        @{ Case = 'RSS-shaped comment'; Content = '<document><!-- <rss><channel><item /></channel></rss> --></document>' }
    ) {
        { Resolve-PodcastSource -Uri 'https://show.example.invalid/privateRootCanary' -Response (Get-DiscoveryResponse -Content $Content) } |
            Should -Throw 'The source XML root is not a supported RSS or Atom feed.'
    }

    It 'does not count nested or unrelated namespace entries as episodes' -ForEach @(
        @{ Content = '<rss><channel><description><item /></description><x:item xmlns:x="urn:other" /></channel></rss>' },
        @{ Content = '<feed xmlns="http://www.w3.org/2005/Atom"><summary><entry /></summary><entry xmlns="urn:other" /></feed>' }
    ) {
        $resolved = Resolve-PodcastSource -Uri 'https://show.example.invalid/feed' -Response (Get-DiscoveryResponse -Content $Content)
        $resolved.Items.Count | Should -Be 0
    }

    It 'classifies HTML containing feed-shaped text as a page' {
        $content = '<!doctype html><html><body><!-- <rss><channel><item /></channel></rss> -->' +
            '<script>var feed = "<feed><entry /></feed>";</script><p>Unclosed HTML paragraph</body></html>'
        $resolved = Resolve-PodcastSource -Uri 'https://show.example.invalid/page' -Response (Get-DiscoveryResponse -Content $content)
        $resolved.Kind | Should -BeExactly 'Html'
        $resolved.Xml | Should -BeNullOrEmpty
        ($resolved.Items -is [array]) | Should -BeTrue
        $resolved.Items.Count | Should -Be 0
        ($resolved.Candidates -is [string[]]) | Should -BeTrue
        $resolved.Candidates.Count | Should -Be 0
    }

    It 'rejects malformed and forbidden XML with the same bounded private parser diagnostic for <Case>' -ForEach @(
        @{ Case = 'malformed root'; Content = '<rss><channel><privateXmlCanary></rss>' },
        @{ Case = 'a DTD'; Content = '<!DOCTYPE rss [<!ENTITY x "privateXmlCanary">]><rss><channel><item>&x;</item></channel></rss>' },
        @{ Case = 'an external DTD'; Content = '<!DOCTYPE rss SYSTEM "https://entity.example.invalid/privateXmlCanary"><rss />' }
    ) {
        { Resolve-PodcastSource -Uri 'https://show.example.invalid/privatePathCanary?token=privateQueryCanary' -Response (Get-DiscoveryResponse -Content $Content) } |
            Should -Throw 'Source XML is invalid or exceeds safe parser limits.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'keeps the exact caller feed identity while allowing the response URI to differ' {
        $requested = 'https://show.example.invalid/feed?token=one%2ftwo#private-section'
        $resolved = Resolve-PodcastSource -Uri $requested -Response (Get-DiscoveryResponse -Content $script:DiscoveryRss -FinalUri 'https://cdn.example.invalid/redirected.xml')
        $resolved.Url | Should -BeExactly $requested
        $resolved.FinalUri.AbsoluteUri | Should -BeExactly 'https://cdn.example.invalid/redirected.xml'
    }

    It 'rejects malformed prefixed Atom even when its response is mislabeled as HTML' {
        $response = Get-DiscoveryResponse -Content '<a:feed xmlns:a="http://www.w3.org/2005/Atom"><a:entry></a:feed>'
        $response | Add-Member NoteProperty ContentType 'text/html'
        { Resolve-PodcastSource -Uri 'https://show.example.invalid/feed' -Response $response } |
            Should -Throw 'Source XML is invalid or exceeds safe parser limits.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }
}

Describe 'A031/A032: shared candidate selection and response reuse' -Tag 'Unit', 'A031', 'A032' {
    BeforeEach { Mock Write-Host {} }

    It 'fetches a direct <Kind> feed once and returns its stable array' -ForEach @(
        @{ Kind = 'RSS' }, @{ Kind = 'Atom' }
    ) {
        $content = $script:DiscoveryRss
        if ($Kind -eq 'Atom') { $content = $script:DiscoveryAtom }
        Mock Invoke-PodcastMetadataRequest { Get-DiscoveryResponse -Content $content -FinalUri 'https://cdn.example.invalid/feed.xml' }
        $resolved = Resolve-PodcastItems -Feeds 'https://show.example.invalid/feed'
        $resolved.Url | Should -BeExactly 'https://show.example.invalid/feed'
        ($resolved.Items -is [array]) | Should -BeTrue
        $resolved.Items.Count | Should -Be 1
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
        Should -Invoke Read-Host -Times 0 -Exactly
    }

    It 'reuses an already resolved direct feed without a second request' {
        $source = Resolve-PodcastSource -Uri 'https://show.example.invalid/feed' -Response (Get-DiscoveryResponse -Content $script:DiscoveryRss)
        $resolved = Resolve-PodcastItems -Feeds 'https://show.example.invalid/feed' -InitialResolution $source
        $resolved.Url | Should -BeExactly 'https://show.example.invalid/feed'
        [object]::ReferenceEquals($resolved.Xml, $source.Xml) | Should -BeTrue
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'consumes an empty initial response once and fetches only the next explicit source' {
        $source = Resolve-PodcastSource -Uri 'https://show.example.invalid/empty' -Response (Get-DiscoveryResponse -Content '<rss><channel /></rss>')
        Mock Invoke-PodcastMetadataRequest { Get-DiscoveryResponse -Content $script:DiscoveryRss }
        $resolved = Resolve-PodcastItems -Feeds @('https://show.example.invalid/empty', 'https://show.example.invalid/good') -InitialResolution $source
        $resolved.Url | Should -BeExactly 'https://show.example.invalid/good'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://show.example.invalid/good' }
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly -ParameterFilter { $Uri -eq 'https://show.example.invalid/empty' }
    }

    It 'fetches only the declared single feed after resolving its page' {
        Mock Invoke-PodcastMetadataRequest {
            if ($Uri -eq 'https://show.example.invalid/page') {
                return Get-DiscoveryResponse -Content '<html><link type="application/rss+xml" href="feed.xml"></html>'
            }
            Get-DiscoveryResponse -Content $script:DiscoveryRss
        }
        $resolved = Resolve-PodcastItems -Feeds 'https://show.example.invalid/page'
        $resolved.Url | Should -BeExactly 'https://show.example.invalid/feed.xml'
        $resolved.Items.Count | Should -Be 1
        Should -Invoke Invoke-PodcastMetadataRequest -Times 2 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://show.example.invalid/page' }
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://show.example.invalid/feed.xml' }
        Should -Invoke Read-Host -Times 0 -Exactly
    }

    It 'reuses an already fetched page and fetches only its selected feed' {
        $response = Get-DiscoveryResponse -Content '<html><link type="application/atom+xml" href="feed.xml"></html>' `
            -FinalUri 'https://destination.example.invalid/shows/index.html'
        $source = Resolve-PodcastSource -Uri 'https://show.example.invalid/page' -Response $response
        Mock Invoke-PodcastMetadataRequest { Get-DiscoveryResponse -Content $script:DiscoveryAtom }
        $resolved = Resolve-PodcastItems -Feeds 'https://show.example.invalid/page' -InitialResolution $source
        $resolved.Url | Should -BeExactly 'https://destination.example.invalid/shows/feed.xml'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://destination.example.invalid/shows/feed.xml' }
    }

    It 'rejects noninteractive ambiguity before fetching any candidate' {
        Mock Invoke-PodcastMetadataRequest {
            Get-DiscoveryResponse -Content '<html><link type="application/rss+xml" href="one.xml"><link type="application/atom+xml" href="two.xml"></html>'
        }
        { Resolve-PodcastItems -Feeds 'https://show.example.invalid/privatePageCanary?token=privateQueryCanary' } |
            Should -Throw 'Multiple feed links were found. Supply a direct feed URL with -FeedUrl.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
        Should -Invoke Read-Host -Times 0 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }

    It 'requires a valid explicit interactive choice and fetches only that candidate' {
        $script:DiscoveryChoices = [Collections.Generic.Queue[string]]::new()
        foreach ($value in @('-1', '0', '3', 'not-a-number', '1.0', '', '2')) { $script:DiscoveryChoices.Enqueue($value) }
        Mock Read-Host { $script:DiscoveryChoices.Dequeue() }
        Mock Invoke-PodcastMetadataRequest {
            if ($Uri -eq 'https://show.example.invalid/page') {
                return Get-DiscoveryResponse -Content '<html><link type="application/rss+xml" href="one.xml?token=privateOneCanary"><link type="application/atom+xml" href="two.xml?token=privateTwoCanary"></html>'
            }
            Get-DiscoveryResponse -Content $script:DiscoveryAtom -FinalUri 'https://cdn.example.invalid/selected.xml'
        }
        $resolved = Resolve-PodcastItems -Feeds 'https://show.example.invalid/page' -Interactive
        $resolved.Url | Should -BeExactly 'https://show.example.invalid/two.xml?token=privateTwoCanary'
        $resolved.Items.Count | Should -Be 1
        Should -Invoke Read-Host -Times 7 -Exactly -ParameterFilter { $Prompt -eq 'Choose feed number (1-2)' }
        Should -Invoke Invoke-PodcastMetadataRequest -Times 2 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly -ParameterFilter { $Uri -like '*one.xml*' }
        Should -Invoke Write-Host -Times 0 -Exactly -ParameterFilter { ($Object | Out-String) -match 'privateOneCanary|privateTwoCanary|two\.xml|one\.xml' }
    }

    It 'does not recursively discover another HTML page from a declared feed' {
        Mock Invoke-PodcastMetadataRequest {
            if ($Uri -eq 'https://show.example.invalid/page') {
                return Get-DiscoveryResponse -Content '<html><link type="application/rss+xml" href="candidate.xml"></html>'
            }
            Get-DiscoveryResponse -Content '<html><link type="application/rss+xml" href="recursive.xml"></html>'
        }
        { Resolve-PodcastItems -Feeds 'https://show.example.invalid/page' } |
            Should -Throw 'The discovered URL did not return an RSS or Atom feed.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 2 -Exactly
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly -ParameterFilter { $Uri -like '*recursive.xml*' }
    }

    It 'reports an empty supported <Kind> feed distinctly' -ForEach @(
        @{ Kind = 'RSS'; Content = '<rss><channel /></rss>' },
        @{ Kind = 'Atom'; Content = '<feed xmlns="http://www.w3.org/2005/Atom" />' }
    ) {
        Mock Invoke-PodcastMetadataRequest { Get-DiscoveryResponse -Content $Content }
        { Resolve-PodcastItems -Feeds 'https://show.example.invalid/empty' } |
            Should -Throw 'The RSS or Atom feed is valid but contains no episodes.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
    }

    It 'reports a page without declared feeds distinctly' {
        Mock Invoke-PodcastMetadataRequest { Get-DiscoveryResponse -Content '<html><body>No feeds here</body></html>' }
        { Resolve-PodcastItems -Feeds 'https://show.example.invalid/page' } |
            Should -Throw 'No RSS or Atom feed links were found on the page.'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
    }

    It 'preserves the last useful source diagnostic for <Case>' -ForEach @(
        @{ Case = 'malformed XML'; Content = '<rss><privateXmlCanary></rss>'; Message = 'Source XML is invalid or exceeds safe parser limits.' },
        @{ Case = 'unsupported root'; Content = '<document><rss><channel><item /></channel></rss></document>'; Message = 'The source XML root is not a supported RSS or Atom feed.' }
    ) {
        Mock Invoke-PodcastMetadataRequest { Get-DiscoveryResponse -Content $Content }
        { Resolve-PodcastItems -Feeds 'https://show.example.invalid/privateUrlCanary' } | Should -Throw $Message
    }

    It 'continues to the next explicit source after <Case>' -ForEach @(
        @{ Case = 'an empty valid feed'; Content = '<rss><channel /></rss>' },
        @{ Case = 'a page without candidates'; Content = '<html><body>No feeds</body></html>' },
        @{ Case = 'malformed feed XML'; Content = '<rss><channel></rss>' }
    ) {
        Mock Invoke-PodcastMetadataRequest {
            if ($Uri -eq 'https://show.example.invalid/first') { return Get-DiscoveryResponse -Content $Content }
            Get-DiscoveryResponse -Content $script:DiscoveryRss
        }
        $resolved = Resolve-PodcastItems -Feeds @('https://show.example.invalid/first', 'https://show.example.invalid/second')
        $resolved.Url | Should -BeExactly 'https://show.example.invalid/second'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 2 -Exactly
    }

    It 'uses the shared resolver for a direct TUI feed and returns the fetched object' {
        Mock Read-Host { 'https://show.example.invalid/feed' }
        Mock Invoke-PodcastMetadataRequest { Get-DiscoveryResponse -Content $script:DiscoveryRss }
        $resolved = Get-FeedUrlInteractive
        $resolved.Url | Should -BeExactly 'https://show.example.invalid/feed'
        $resolved.Items.Count | Should -Be 1
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
        Should -Invoke Read-Host -Times 1 -Exactly
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }
}
