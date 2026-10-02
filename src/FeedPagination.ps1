#requires -Version 5.1

function Get-PodcastFeedPageLinks {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Returns the bounded continuation link collection.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][Xml.XmlDocument]$Xml, [Parameter(Mandatory)][Uri]$BaseUri)

    $namespaces = [Xml.XmlNamespaceManager]::new($Xml.NameTable)
    $namespaces.AddNamespace('atom', 'http://www.w3.org/2005/Atom')
    $links = $Xml.SelectNodes('/rss/channel/atom:link | /atom:feed/atom:link', $namespaces)
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($link in $links) {
        $relation = $link.GetAttribute('rel')
        if (@('next', 'prev-archive', 'http://www.iana.org/assignments/relation/next',
                'http://www.iana.org/assignments/relation/prev-archive') -cnotcontains $relation) { continue }
        $type = $link.GetAttribute('type').Split(';')[0].Trim()
        if ($type.Length -gt 0 -and $type -ine 'application/rss+xml' -and $type -ine 'application/atom+xml') {
            throw 'The advertised continuation is not a supported feed link.'
        }
        $href = $link.GetAttribute('href')
        if ([string]::IsNullOrWhiteSpace($href)) { throw 'The advertised continuation has no usable target.' }
        $ancestors = [Collections.Generic.List[Xml.XmlElement]]::new()
        $node = $link
        while ($node -is [Xml.XmlElement]) { $ancestors.Add($node); $node = $node.ParentNode }
        $effectiveBase = $BaseUri
        for ($i = $ancestors.Count - 1; $i -ge 0; $i--) {
            if ($ancestors[$i].HasAttribute('base', 'http://www.w3.org/XML/1998/namespace')) {
                $reference = $ancestors[$i].GetAttribute('base', 'http://www.w3.org/XML/1998/namespace')
                $candidate = Get-PodcastDiscoveredUri -Reference $reference -BaseUri $effectiveBase
                $effectiveBase = Get-PodcastRequestUri -Uri $candidate.AbsoluteUri -PreviousUri $effectiveBase
            }
        }
        $candidate = Get-PodcastDiscoveredUri -Reference $href -BaseUri $effectiveBase
        $target = Get-PodcastRequestUri -Uri $candidate.AbsoluteUri -PreviousUri $effectiveBase
        $key = $target.GetComponents([UriComponents]::HttpRequestUrl, [UriFormat]::UriEscaped)
        if ($seen.Add($key)) {
            $target.AbsoluteUri
            # Only one chain is supported. Two distinct targets already prove
            # ambiguity; do not expand and retain the rest of a large link set.
            if ($seen.Count -ge 2) { return }
        }
    }
}

function Resolve-PodcastCatalogue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$InitialResolution,
        [ValidateRange(1, 100)][int]$MaxPages = 20,
        [ValidateRange(1, 10000)][int]$MaxItems = 10000,
        [ValidateRange(1, 33554432)][int]$MaxCharacters = 33554432
    )

    $first = $InitialResolution
    $page = $first
    $items = [Collections.Generic.List[object]]::new()
    $identities = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $feedId = Get-PodcastNameHash -IdentityKey ('feed:' + $first.Url)
    $pagesFetched = 1
    $itemsSeen = 0
    $duplicates = 0
    $characters = 0
    $reason = 'end'
    while ($true) {
        # Fragment identifiers are not transmitted in HTTP requests. Both the
        # requested and effective URI are aliases for cycle detection only;
        # the original selected feed URL remains the archive identity.
        foreach ($alias in @($page.Url, $page.FinalUri.AbsoluteUri)) {
            $uri = Get-PodcastRequestUri -Uri $alias
            $key = $uri.GetComponents([UriComponents]::HttpRequestUrl, [UriFormat]::UriEscaped)
            $null = $visited.Add($key)
        }
        if ($characters + $page.Content.Length -gt $MaxCharacters) { $reason = 'character_limit'; break }
        $characters += $page.Content.Length
        foreach ($item in $page.Items) {
            if ($itemsSeen -ge $MaxItems) { $reason = 'item_limit'; break }
            $itemsSeen++
            $episode = Get-EpisodeData -XmlItem $item
            # Unsupported entries are retained for the ordinary planning warning.
            # Entries without any durable identity cannot contribute a duplicate.
            if (-not [string]::IsNullOrWhiteSpace([string]$episode.Guid) -or
                -not [string]::IsNullOrWhiteSpace([string]$episode.AtomId) -or
                -not [string]::IsNullOrWhiteSpace([string]$episode.Url)) {
                $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $feedId
                $metadata = Get-EpisodeMetadataKey -Episode $episode
                if ($identities.ContainsKey($identity.Key)) {
                    if (-not [StringComparer]::Ordinal.Equals($identities[$identity.Key], $metadata)) {
                        throw 'Conflicting episode metadata reuses one identity in this feed snapshot; no media destinations were created.'
                    }
                    $duplicates++
                    continue
                }
                $identities.Add($identity.Key, $metadata)
            }
            $items.Add($item)
        }
        if ($reason -ne 'end') { break }
        try { $links = @(Get-PodcastFeedPageLinks -Xml $page.Xml -BaseUri $page.FinalUri) }
        catch { $reason = 'invalid_link'; break }
        if ($links.Count -eq 0) { break }
        if ($links.Count -gt 1) { $reason = 'ambiguous_link'; break }
        $target = Get-PodcastRequestUri -Uri $links[0]
        $key = $target.GetComponents([UriComponents]::HttpRequestUrl, [UriFormat]::UriEscaped)
        if ($visited.Contains($key)) { $reason = 'cycle'; break }
        if ($pagesFetched -ge $MaxPages) { $reason = 'page_limit'; break }
        if ($itemsSeen -ge $MaxItems) { $reason = 'item_limit'; break }
        if ($characters -ge $MaxCharacters) { $reason = 'character_limit'; break }
        try { $nextPage = Resolve-PodcastSource -Uri $target.AbsoluteUri }
        catch {
            $reason = if ($null -ne (Get-PodcastTransportFailure -ErrorObject $_)) { 'page_failed' } else { 'invalid_page' }
            break
        }
        if ($nextPage.Kind -cne $first.Kind) { $reason = 'invalid_page'; break }
        $pagesFetched++
        $finalKey = $nextPage.FinalUri.GetComponents([UriComponents]::HttpRequestUrl, [UriFormat]::UriEscaped)
        if ($visited.Contains($finalKey)) { $reason = 'cycle'; break }
        $page = $nextPage
    }
    return [pscustomobject]@{
        Url = $first.Url
        Xml = $first.Xml
        Items = [object[]]$items.ToArray()
        Kind = $first.Kind
        FinalUri = $first.FinalUri
        Content = $first.Content
        Candidates = $first.Candidates
        Catalogue = [pscustomobject]@{
            Complete = ($reason -eq 'end')
            StopReason = $reason
            PagesFetched = $pagesFetched
            ItemsSeen = $itemsSeen
            DuplicateCount = $duplicates
            MaxPages = $MaxPages
        }
    }
}

function Get-PodcastCatalogueMessage {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Catalogue)

    $detail = switch ([string]$Catalogue.StopReason) {
        'cycle' { 'an advertised page cycle was detected.' }
        'page_limit' { 'the configured page limit was reached.' }
        'item_limit' { 'the entry limit was reached.' }
        'character_limit' { 'the metadata character limit was reached.' }
        'invalid_link' { 'an advertised continuation link was invalid or unsupported.' }
        'ambiguous_link' { 'multiple distinct continuation targets were advertised.' }
        'page_failed' { 'an advertised page could not be retrieved.' }
        'invalid_page' { 'an advertised page did not return a valid feed of the same kind.' }
        default { 'advertised pages remain unresolved.' }
    }
    return ('Feed catalogue incomplete: ' + $detail)
}
