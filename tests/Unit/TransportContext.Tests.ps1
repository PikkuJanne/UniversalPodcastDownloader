BeforeAll {
    $script:TransportContextRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Mock Read-Host { throw 'Transport context units must not prompt.' }
    Mock Invoke-WebRequest { throw 'Transport context units must not use the legacy network client.' }
    . (Join-Path $script:TransportContextRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Transport context units must not create a network client.' }
    Mock Invoke-PodcastMediaRequest { throw 'Transport context units must not request media.' }

    function Get-TransportContextResponse {
        param([string]$Content, [string]$Uri)
        return [pscustomobject]@{ Content = $Content; FinalUri = [Uri]$Uri; ContentType = 'application/xml' }
    }

    $script:TransportContextCurrentUrl = 'https://feeds.example.invalid/current.xml'
    $script:TransportContextNextUrl = 'https://feeds.example.invalid/older.xml'
    $script:TransportContextItem = '<item><title>Synthetic</title><guid>one</guid><enclosure type="audio/mpeg" url="https://media.example.invalid/one.mp3" /></item>'
    $script:TransportContextRss = '<rss><channel><title>Synthetic show</title>' + $script:TransportContextItem + '</channel></rss>'
    $script:TransportContextPagedRss = '<rss xmlns:atom="http://www.w3.org/2005/Atom"><channel><title>Synthetic show</title><atom:link rel="next" href="older.xml" />' + $script:TransportContextItem + '</channel></rss>'
}

Describe 'A048 transport context is explicit and run scoped' -Tag 'Unit', 'A048' {
    BeforeEach {
        $script:TransportContextResponses = @{}
        $script:TransportContextRequests = [Collections.Generic.List[object]]::new()
        Mock Invoke-PodcastMetadataRequest {
            param($Uri, $Policy)
            $script:TransportContextRequests.Add([pscustomobject]@{ Uri = [string]$Uri; Policy = $Policy })
            if (-not $script:TransportContextResponses.ContainsKey([string]$Uri)) { throw 'Unexpected synthetic transport context request.' }
            return $script:TransportContextResponses[[string]$Uri]
        }
    }

    It 'forwards an explicit policy from source classification to the existing metadata adapter' {
        $policy = New-PodcastTransportPolicy -MaxAttempts 1 -HeaderTimeoutSeconds 2 -IdleTimeoutSeconds 3
        $script:TransportContextResponses[$script:TransportContextCurrentUrl] = Get-TransportContextResponse -Content $script:TransportContextRss -Uri $script:TransportContextCurrentUrl

        $source = Resolve-PodcastSource -Uri $script:TransportContextCurrentUrl -Policy $policy

        $source.Kind | Should -BeExactly 'Rss'
        $script:TransportContextRequests.Count | Should -Be 1
        [object]::ReferenceEquals($script:TransportContextRequests[0].Policy, $policy) | Should -BeTrue
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }

    It 'forwards an explicit catalogue policy only to the advertised next-page fetch' {
        $policy = New-PodcastTransportPolicy -MaxAttempts 1 -RetryBudgetSeconds 7
        $initial = Resolve-PodcastSource -Uri $script:TransportContextCurrentUrl -Response (Get-TransportContextResponse -Content $script:TransportContextPagedRss -Uri $script:TransportContextCurrentUrl)
        $script:TransportContextResponses[$script:TransportContextNextUrl] = Get-TransportContextResponse -Content $script:TransportContextRss -Uri $script:TransportContextNextUrl

        $catalogue = Resolve-PodcastCatalogue -InitialResolution $initial -Policy $policy

        $catalogue.Catalogue.Complete | Should -BeTrue
        $catalogue.Catalogue.PagesFetched | Should -Be 2
        $catalogue.Catalogue.DuplicateCount | Should -Be 1
        $script:TransportContextRequests.Count | Should -Be 1
        $script:TransportContextRequests[0].Uri | Should -BeExactly $script:TransportContextNextUrl
        [object]::ReferenceEquals($script:TransportContextRequests[0].Policy, $policy) | Should -BeTrue
    }

    It 'carries one explicit policy through page discovery feed selection and pagination without extra requests' {
        $policy = New-PodcastTransportPolicy -MaxAttempts 2 -HeaderTimeoutSeconds 4 -IdleTimeoutSeconds 5 -RetryBudgetSeconds 6
        $pageUrl = 'https://feeds.example.invalid/page'
        $script:TransportContextResponses[$pageUrl] = Get-TransportContextResponse -Content '<html><link rel="alternate" type="application/rss+xml" href="current.xml"></html>' -Uri $pageUrl
        $script:TransportContextResponses[$script:TransportContextCurrentUrl] = Get-TransportContextResponse -Content $script:TransportContextPagedRss -Uri $script:TransportContextCurrentUrl
        $script:TransportContextResponses[$script:TransportContextNextUrl] = Get-TransportContextResponse -Content $script:TransportContextRss -Uri $script:TransportContextNextUrl

        $resolved = Resolve-PodcastItems -Feeds @($pageUrl) -Policy $policy -MaxPages 2

        $resolved.SourceUrl | Should -BeExactly $pageUrl
        $resolved.Url | Should -BeExactly $script:TransportContextCurrentUrl
        $resolved.Catalogue.Complete | Should -BeTrue
        $resolved.Catalogue.PagesFetched | Should -Be 2
        $resolved.Catalogue.MaxPages | Should -Be 2
        $resolved.Items.Count | Should -Be 1
        ($script:TransportContextRequests.Uri -join ',') | Should -BeExactly ($pageUrl + ',' + $script:TransportContextCurrentUrl + ',' + $script:TransportContextNextUrl)
        foreach ($request in $script:TransportContextRequests) {
            [object]::ReferenceEquals($request.Policy, $policy) | Should -BeTrue
        }
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
        Should -Invoke Read-Host -Times 0 -Exactly
    }

    It 'keeps a supplied source response fetch free with an explicit policy' {
        $policy = New-PodcastTransportPolicy -MaxAttempts 1
        $response = Get-TransportContextResponse -Content $script:TransportContextRss -Uri 'https://cdn.example.invalid/redirected.xml'

        $source = Resolve-PodcastSource -Uri $script:TransportContextCurrentUrl -Response $response -Policy $policy

        $source.Kind | Should -BeExactly 'Rss'
        $source.Url | Should -BeExactly $script:TransportContextCurrentUrl
        $source.FinalUri.AbsoluteUri | Should -BeExactly 'https://cdn.example.invalid/redirected.xml'
        $source.Items.Count | Should -Be 1
        $script:TransportContextRequests.Count | Should -Be 0
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'reuses an initial discovered page and applies the explicit policy only to feed and continuation requests' {
        $policy = New-PodcastTransportPolicy -MaxAttempts 1
        $pageUrl = 'https://feeds.example.invalid/page'
        $initial = Resolve-PodcastSource -Uri $pageUrl -Response (Get-TransportContextResponse -Content '<html><link type="application/rss+xml" href="current.xml"></html>' -Uri $pageUrl)
        $script:TransportContextResponses[$script:TransportContextCurrentUrl] = Get-TransportContextResponse -Content $script:TransportContextPagedRss -Uri $script:TransportContextCurrentUrl
        $script:TransportContextResponses[$script:TransportContextNextUrl] = Get-TransportContextResponse -Content $script:TransportContextRss -Uri $script:TransportContextNextUrl

        $resolved = Resolve-PodcastItems -Feeds @($pageUrl) -InitialResolution $initial -Policy $policy

        $resolved.Catalogue.Complete | Should -BeTrue
        ($script:TransportContextRequests.Uri -join ',') | Should -BeExactly ($script:TransportContextCurrentUrl + ',' + $script:TransportContextNextUrl)
        foreach ($request in $script:TransportContextRequests) {
            [object]::ReferenceEquals($request.Policy, $policy) | Should -BeTrue
        }
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly -ParameterFilter { $Uri -ceq $pageUrl }
    }

    It 'reuses an initial completed catalogue without fetching it under a later policy' {
        $source = Resolve-PodcastSource -Uri $script:TransportContextCurrentUrl -Response (Get-TransportContextResponse -Content $script:TransportContextRss -Uri $script:TransportContextCurrentUrl)
        $initial = Resolve-PodcastCatalogue -InitialResolution $source
        $policy = New-PodcastTransportPolicy -MaxAttempts 1

        $resolved = Resolve-PodcastItems -Feeds @($script:TransportContextCurrentUrl) -InitialResolution $initial -Policy $policy -MaxPages 1

        [object]::ReferenceEquals($resolved, $initial) | Should -BeTrue
        $resolved.Catalogue.Complete | Should -BeTrue
        $script:TransportContextRequests.Count | Should -Be 0
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'creates independent default source policies rather than consulting a previous script policy' {
        $previous = Get-Variable PodcastTransportPolicy -Scope Script -ErrorAction SilentlyContinue
        $previousValue = if ($null -ne $previous) { $previous.Value } else { $null }
        $poison = New-PodcastTransportPolicy -MaxAttempts 10 -HeaderTimeoutSeconds 9
        $script:TransportContextResponses[$script:TransportContextCurrentUrl] = Get-TransportContextResponse -Content $script:TransportContextRss -Uri $script:TransportContextCurrentUrl
        try {
            $script:PodcastTransportPolicy = $poison
            $null = Resolve-PodcastSource -Uri $script:TransportContextCurrentUrl
            $null = Resolve-PodcastSource -Uri $script:TransportContextCurrentUrl

            $script:TransportContextRequests.Count | Should -Be 2
            foreach ($request in $script:TransportContextRequests) {
                [object]::ReferenceEquals($request.Policy, $poison) | Should -BeFalse
                $request.Policy.MaxAttempts | Should -Be 3
                $request.Policy.HeaderTimeoutSeconds | Should -Be 30
            }
            [object]::ReferenceEquals($script:TransportContextRequests[0].Policy, $script:TransportContextRequests[1].Policy) | Should -BeFalse
        }
        finally {
            if ($null -ne $previous) { $script:PodcastTransportPolicy = $previousValue }
            else { Remove-Variable PodcastTransportPolicy -Scope Script -ErrorAction SilentlyContinue }
        }
    }

    It 'uses one fresh default pipeline policy and the documented page limit despite prior script values' {
        $previousPolicy = Get-Variable PodcastTransportPolicy -Scope Script -ErrorAction SilentlyContinue
        $previousPages = Get-Variable PodcastMaxFeedPages -Scope Script -ErrorAction SilentlyContinue
        $previousPolicyValue = if ($null -ne $previousPolicy) { $previousPolicy.Value } else { $null }
        $previousPagesValue = if ($null -ne $previousPages) { $previousPages.Value } else { $null }
        $poison = New-PodcastTransportPolicy -MaxAttempts 10
        $script:TransportContextResponses[$script:TransportContextCurrentUrl] = Get-TransportContextResponse -Content $script:TransportContextPagedRss -Uri $script:TransportContextCurrentUrl
        $script:TransportContextResponses[$script:TransportContextNextUrl] = Get-TransportContextResponse -Content $script:TransportContextRss -Uri $script:TransportContextNextUrl
        try {
            $script:PodcastTransportPolicy = $poison
            $script:PodcastMaxFeedPages = 1

            $resolved = Resolve-PodcastItems -Feeds @($script:TransportContextCurrentUrl)

            $resolved.Catalogue.Complete | Should -BeTrue
            $resolved.Catalogue.MaxPages | Should -Be 20
            $resolved.Catalogue.PagesFetched | Should -Be 2
            $script:TransportContextRequests.Count | Should -Be 2
            $firstPolicy = $script:TransportContextRequests[0].Policy
            [object]::ReferenceEquals($firstPolicy, $poison) | Should -BeFalse
            $firstPolicy.MaxAttempts | Should -Be 3
            [object]::ReferenceEquals($script:TransportContextRequests[1].Policy, $firstPolicy) | Should -BeTrue
        }
        finally {
            if ($null -ne $previousPolicy) { $script:PodcastTransportPolicy = $previousPolicyValue }
            else { Remove-Variable PodcastTransportPolicy -Scope Script -ErrorAction SilentlyContinue }
            if ($null -ne $previousPages) { $script:PodcastMaxFeedPages = $previousPagesValue }
            else { Remove-Variable PodcastMaxFeedPages -Scope Script -ErrorAction SilentlyContinue }
        }
    }
}
