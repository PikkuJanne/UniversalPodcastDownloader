#requires -Version 5.1

function Get-PodcastHtmlAttribute {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    # This deliberately small HTML lexer accepts quoted and unquoted attribute
    # values. First occurrences win, as in HTML; attribute names are case blind.
    $attributes = @{}
    $pattern = '(?<name>[^\t\n\f\r "''<>/=]+)(?:[\t\n\f\r ]*=[\t\n\f\r ]*(?:"(?<double>[^"]*)"|''(?<single>[^'']*)''|(?<plain>[^\t\n\f\r "''=<>`]+)))?'
    foreach ($match in [Text.RegularExpressions.Regex]::Matches($Text, $pattern,
        [Text.RegularExpressions.RegexOptions]::CultureInvariant, [TimeSpan]::FromMilliseconds(250))) {
        $name = $match.Groups['name'].Value.ToLowerInvariant()
        if ($attributes.ContainsKey($name)) { continue }
        $value = ''
        foreach ($groupName in @('double', 'single', 'plain')) {
            if ($match.Groups[$groupName].Success) { $value = $match.Groups[$groupName].Value; break }
        }
        $attributes[$name] = [Net.WebUtility]::HtmlDecode($value)
    }
    return $attributes
}

function Get-PodcastHtmlToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Html)

    # Consume comments and raw-text elements as complete tokens so their text
    # cannot contribute apparent base/link elements. Quoted > signs are legal.
    # A stray < inside a tag still belongs to that tag; it cannot start a link.
    # Unterminated tags/quotes consume the remaining input without discoveries.
    $attributeText = '(?:[^"''>]|"[^"]*"|''[^'']*'')*'
    $pattern = '<!--[\s\S]*?(?:-->|$)|<plaintext(?=[\t\n\f\r />])' + $attributeText + '>[\s\S]*$|<(?<raw>script|style|textarea|title|xmp|iframe|noembed|noframes|noscript)(?=[\t\n\f\r />])' +
        $attributeText + '>[\s\S]*?(?:</\k<raw>[\t\n\f\r ]*>|$)|<![^>]*>|<\?[\s\S]*?\?>|<(?<closing>/)?(?<name>[A-Za-z][A-Za-z0-9:-]*)(?=[\t\n\f\r />])(?<attributes>' + $attributeText + ')>|' +
        '</?[A-Za-z][A-Za-z0-9:-]*(?=[\t\n\f\r />])' + $attributeText + '(?:"[^"]*$|''[^'']*$|$)'
    $options = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant
    $count = 0
    foreach ($token in [Text.RegularExpressions.Regex]::Matches($Html, $pattern, $options, [TimeSpan]::FromMilliseconds(250))) {
        $count++
        if ($count -gt 100000) { throw 'HTML feed discovery exceeds the safe token limit.' }
        if ($token.Groups['name'].Success) { $token }
    }
}

function Get-PodcastDiscoveredUri {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Reference, [Parameter(Mandatory)][Uri]$BaseUri)

    # HTML permits surrounding ASCII whitespace. Check the remaining reference
    # before URI construction can escape or normalize unsafe input for us.
    $value = $Reference.Trim([char[]]@([char]9, [char]10, [char]12, [char]13, [char]32))
    if ($value.Length -gt 32768 -or $value -match '[\\\x00-\x20\x7f]') { throw 'Invalid discovery target.' }
    $resolved = $null
    if (-not [Uri]::TryCreate($BaseUri, $value, [ref]$resolved)) { throw 'Invalid discovery target.' }
    return (Get-PodcastRequestUri -Uri $resolved.AbsoluteUri)
}

function Find-RssInHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Html,
        [Parameter(Mandatory)][string]$BaseUrl
    )

    if ($Html.Length -gt 8388608) { throw 'HTML metadata exceeds the safe character limit.' }
    $pageBase = Get-PodcastRequestUri -Uri $BaseUrl
    $documentBase = $pageBase
    $hasBase = $false
    $templateDepth = 0
    $references = [Collections.Generic.List[string]]::new()
    try {
        foreach ($token in Get-PodcastHtmlToken -Html $Html) {
            $name = $token.Groups['name'].Value
            if ($name -ieq 'template') {
                if ($token.Groups['closing'].Success) { $templateDepth = [Math]::Max(0, $templateDepth - 1) }
                else { $templateDepth++ }
                continue
            }
            if ($templateDepth -gt 0 -or $token.Groups['closing'].Success) { continue }
            if ($name -ine 'base' -and $name -ine 'link') { continue }
            $attributes = Get-PodcastHtmlAttribute -Text $token.Groups['attributes'].Value
            if ($name -ieq 'base') {
                if (-not $hasBase -and $attributes.ContainsKey('href')) {
                    $hasBase = $true
                    try { $documentBase = Get-PodcastDiscoveredUri -Reference $attributes['href'] -BaseUri $pageBase }
                    catch { throw 'HTML base URL is not allowed by the network policy.' }
                }
                continue
            }
            if (-not $attributes.ContainsKey('href') -or -not $attributes.ContainsKey('type')) { continue }
            $type = $attributes['type'].Trim()
            if ($type -ine 'application/rss+xml' -and $type -ine 'application/atom+xml') { continue }
            # Missing rel remains compatible with the old discovery helper.
            if ($attributes.ContainsKey('rel') -and
                -not [Text.RegularExpressions.Regex]::IsMatch($attributes['rel'], '(?:^|[\t\n\f\r ])alternate(?:$|[\t\n\f\r ])',
                    ([Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant),
                    [TimeSpan]::FromMilliseconds(250))) { continue }
            if (-not [string]::IsNullOrWhiteSpace($attributes['href'])) { $references.Add($attributes['href']) }
        }
    }
    catch [Text.RegularExpressions.RegexMatchTimeoutException] { throw 'HTML feed discovery exceeded its safe parser timeout.' }

    $candidates = [Collections.Generic.List[string]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($reference in $references) {
        try { $candidate = (Get-PodcastDiscoveredUri -Reference $reference -BaseUri $documentBase).AbsoluteUri }
        catch { throw 'Discovered feed URL is not allowed by the network policy.' }
        if ($seen.Add($candidate)) { $candidates.Add($candidate) }
    }
    # Emit candidate strings, including no output for an empty collection.
    return $candidates.ToArray()
}

function Test-PodcastHtmlSource {
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Content, [AllowNull()][string]$ContentType)

    if ($Content.Length -gt 8388608) { throw 'Source content exceeds the safe character limit.' }
    # Feed proof comes only from the XML DOM below. This prefix check merely
    # recognizes HTML declarations/root markup when ordinary HTML is not XML.
    $prefix = '\A[\s\uFEFF]*(?:<!--[\s\S]*?-->[\s\uFEFF]*|<\?[\s\S]*?\?>[\s\uFEFF]*)*'
    $pattern = $prefix + '(?:<!DOCTYPE[\t\n\f\r ]+html(?=[\t\n\f\r >])|<html(?=[\t\n\f\r />]))'
    $options = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant
    try {
        if ([Text.RegularExpressions.Regex]::IsMatch($Content, $pattern, $options,
            [TimeSpan]::FromMilliseconds(250))) { return $true }
        # A mislabeled malformed feed or forbidden non-HTML DTD remains an XML
        # error rather than silently becoming a page with no discovered links.
        if ([Text.RegularExpressions.Regex]::IsMatch($Content, $prefix + '(?:<!DOCTYPE\b|<(?:[A-Za-z_][A-Za-z0-9._-]*:)?(?:rss|feed)(?=[\t\n\f\r />]))',
            $options, [TimeSpan]::FromMilliseconds(250))) { return $false }
    }
    catch [Text.RegularExpressions.RegexMatchTimeoutException] { throw 'HTML feed discovery exceeded its safe parser timeout.' }
    # Omitted HTML root tags are supported for an explicit HTML response type.
    return $ContentType -match '^(?i:text/html|application/xhtml\+xml)(?:\s*;|\s*$)'
}

function Resolve-PodcastSource {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri, $Response)

    $original = Get-PodcastRequestUri -Uri $Uri
    if ($null -eq $Response) { $Response = Invoke-PodcastWebRequest -Uri $Uri }
    if ($null -eq $Response -or $null -eq $Response.PSObject.Properties['Content'] -or
        $null -eq $Response.Content -or $Response.Content -isnot [string]) {
        throw 'The metadata response does not contain supported source text.'
    }
    $content = [string]$Response.Content
    $finalUri = $original
    if ($null -ne $Response.PSObject.Properties['FinalUri'] -and $null -ne $Response.FinalUri) {
        $finalUri = Get-PodcastRequestUri -Uri ([string]$Response.FinalUri) -PreviousUri $original
    }
    $contentType = ''
    if ($null -ne $Response.PSObject.Properties['ContentType']) { $contentType = [string]$Response.ContentType }
    $xml = $null
    try { $xml = ConvertFrom-PodcastFeedXml -Content $content }
    catch {
        if (-not (Test-PodcastHtmlSource -Content $content -ContentType $contentType)) {
            throw 'Source XML is invalid or exceeds safe parser limits.'
        }
    }

    $kind = 'Html'
    $items = [object[]]@()
    if ($null -ne $xml) {
        $root = $xml.DocumentElement
        if ($root.LocalName -ceq 'rss' -and $root.NamespaceURI -ceq '') {
            $kind = 'Rss'
            $items = [object[]]@($xml.SelectNodes('/rss/channel/item'))
        }
        elseif ($root.LocalName -ceq 'feed' -and $root.NamespaceURI -ceq 'http://www.w3.org/2005/Atom') {
            $kind = 'Atom'
            $namespaces = [Xml.XmlNamespaceManager]::new($xml.NameTable)
            $namespaces.AddNamespace('atom', 'http://www.w3.org/2005/Atom')
            $items = [object[]]@($xml.SelectNodes('/atom:feed/atom:entry', $namespaces))
        }
        elseif ($root.LocalName -ine 'html' -or ($root.NamespaceURI -cne '' -and $root.NamespaceURI -cne 'http://www.w3.org/1999/xhtml')) {
            throw 'The source XML root is not a supported RSS or Atom feed.'
        }
    }

    $candidates = [string[]]@($Uri)
    if ($kind -eq 'Html') {
        $xml = $null
        $candidates = [string[]]@(Find-RssInHtml -Html $content -BaseUrl $finalUri.AbsoluteUri)
    }
    return [pscustomobject]@{
        Kind = $kind
        Url = $Uri
        FinalUri = $finalUri
        Content = $content
        Xml = $xml
        Items = $items
        Candidates = $candidates
    }
}
