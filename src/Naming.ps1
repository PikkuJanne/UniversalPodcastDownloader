# Pure Windows component naming and destination planning. No disk or network access.

function Get-ShortenedNameText {
    param([string]$Text, [int]$MaxLength)

    if ($MaxLength -lt 1) { return '' }
    if ($Text.Length -le $MaxLength) { return $Text }
    $result = $Text.Substring(0, $MaxLength)
    # NTFS limits count UTF-16 code units. Keep supplementary Unicode characters whole.
    if ([char]::IsHighSurrogate($result[$result.Length - 1])) {
        $result = $result.Substring(0, $result.Length - 1)
    }
    return $result
}

function Sanitize-ForWindowsName {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Keep the existing importable helper name.')]
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$Name,
        [ValidateRange(1, 255)][int]$MaxLength = 255,
        [string]$FallbackName = 'Untitled'
    )

    $raw = if ($null -eq $Name) { '' } else { $Name.Trim() }
    if ($raw -eq '.' -or $raw -eq '..') {
        throw 'A Windows name cannot be a dot or dot-dot path component.'
    }
    if ($raw -match '^[\\/]' -or $raw -match '^[A-Za-z]:') {
        throw 'Metadata cannot supply an absolute, rooted or drive-relative Windows path.'
    }

    # Normalize valid Unicode but replace malformed surrogate code units first.
    $raw = [regex]::Replace($raw, '[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]', '_')
    $safe = $raw.Normalize([Text.NormalizationForm]::FormC)
    $safe = [regex]::Replace($safe, '[\p{Cc}\p{Cf}<>:"/\\|?*]', '_').Trim().TrimEnd([char[]]' .')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        if ([string]::IsNullOrWhiteSpace($FallbackName) -or $FallbackName -eq $Name) {
            throw 'A nonempty safe fallback name is required.'
        }
        return Sanitize-ForWindowsName -Name $FallbackName -MaxLength $MaxLength -FallbackName '_'
    }

    # Windows also reserves the superscript forms of 1, 2 and 3 for COM/LPT.
    $reserved = '^(CON|PRN|AUX|NUL|CONIN\$|CONOUT\$|COM[1-9\u00B9\u00B2\u00B3]|LPT[1-9\u00B9\u00B2\u00B3])(?:[ .]|$)'
    if ($safe -match $reserved) { $safe = '_' + $safe }
    $safe = (Get-ShortenedNameText -Text $safe -MaxLength $MaxLength).TrimEnd([char[]]' .')
    if ([string]::IsNullOrEmpty($safe)) { $safe = '_' }
    if ($safe -match $reserved) {
        $safe = '_' + (Get-ShortenedNameText -Text $safe -MaxLength ($MaxLength - 1))
    }
    return $safe
}

function Get-PodcastNameHash {
    param([Parameter(Mandatory)][string]$IdentityKey)

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($IdentityKey)))).Replace('-', '').ToLowerInvariant()
    }
    finally { $sha.Dispose() }
}

function Get-EpisodeMetadataKey {
    param([Parameter(Mandatory)]$Episode)

    $title = [string]$Episode.Title
    $url = [string]$Episode.Url
    $date = if ($Episode.PubDate) { ([datetime]$Episode.PubDate).ToString('o', [Globalization.CultureInfo]::InvariantCulture) } else { '' }
    # Length prefixes keep embedded delimiters unambiguous without exposing them in names.
    return ('{0}:{1}{2}:{3}{4}:{5}' -f $title.Length, $title, $date.Length, $date, $url.Length, $url)
}

function Get-EpisodeIdentityKey {
    param([Parameter(Mandatory)]$Episode)

    if (-not [string]::IsNullOrWhiteSpace([string]$Episode.Guid)) { return 'guid:' + [string]$Episode.Guid }
    if (-not [string]::IsNullOrWhiteSpace([string]$Episode.Url)) { return 'url:' + [string]$Episode.Url }
    return 'metadata:' + (Get-EpisodeMetadataKey -Episode $Episode)
}

function New-EpisodeFileName {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This helper returns a string and has no side effects.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Episode,
        [int]$Index,
        [ValidateRange(1, 255)][int]$MaxLength = 180
    )

    # Index is accepted for compatibility, but never participates in stable identity.
    $null = $Index
    $extension = 'mp3'
    if ([string]$Episode.Url -match '\.(mp3|m4a)(?:$|[?#])') { $extension = $Matches[1].ToLowerInvariant() }
    $hash = Get-PodcastNameHash -IdentityKey (Get-EpisodeIdentityKey -Episode $Episode)
    $suffix = '-' + $hash + '.' + $extension
    $titleBudget = $MaxLength - $suffix.Length
    if ($titleBudget -lt 1) {
        throw 'The filename length budget cannot retain a title, the full episode identifier and extension. Choose a shorter output root.'
    }

    $title = Sanitize-ForWindowsName -Name ([string]$Episode.Title) -FallbackName 'Episode'
    $prefix = if ($Episode.PubDate) { ([datetime]$Episode.PubDate).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) + ' - ' } else { '' }
    if ($titleBudget -le $prefix.Length) { $prefix = '' }
    $title = Get-ShortenedNameText -Text $title -MaxLength ($titleBudget - $prefix.Length)
    if ([string]::IsNullOrEmpty($title)) { $title = '_' }
    return $prefix + $title + $suffix
}

function New-PodcastFolderName {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This helper returns a string and has no side effects.')]
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$FeedTitle,
        [Parameter(Mandatory)][string]$FeedUrl,
        [ValidateRange(1, 255)][int]$MaxLength = 120
    )

    $suffix = '-' + (Get-PodcastNameHash -IdentityKey ('feed:' + $FeedUrl))
    if ($MaxLength -le $suffix.Length) {
        throw 'The folder length budget cannot retain a name and the full feed identifier. Choose a shorter output root.'
    }
    $title = Sanitize-ForWindowsName -Name $FeedTitle -FallbackName 'Podcast' -MaxLength ($MaxLength - $suffix.Length)
    return $title + $suffix
}

function New-PodcastDestinationPlan {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This helper returns an in-memory plan and has no side effects.')]
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Episodes,
        [ValidateRange(1, 255)][int]$MaxFileNameLength = 180
    )

    $identities = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $hashes = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $names = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
    $plan = [Collections.Generic.List[object]]::new()
    foreach ($episode in $Episodes) {
        if ($null -eq $episode) { throw 'A destination cannot be planned for an empty episode.' }
        $key = Get-EpisodeIdentityKey -Episode $episode
        $metadata = Get-EpisodeMetadataKey -Episode $episode
        if ($identities.ContainsKey($key)) {
            if (-not [StringComparer]::Ordinal.Equals($identities[$key], $metadata)) {
                throw 'Conflicting episode metadata reuses one identity; no media destinations were created.'
            }
            continue
        }
        $hash = Get-PodcastNameHash -IdentityKey $key
        if ($hashes.ContainsKey($hash) -and -not [StringComparer]::Ordinal.Equals($hashes[$hash], $key)) {
            throw 'Distinct episode identities have a hash collision; no media destinations were created.'
        }
        $fileName = New-EpisodeFileName -Episode $episode -MaxLength $MaxFileNameLength
        if ($names.ContainsKey($fileName)) {
            throw 'Episode destinations collide under Windows case-insensitive comparison; no media destinations were created.'
        }
        $identities.Add($key, $metadata)
        $hashes.Add($hash, $key)
        $names.Add($fileName, $key)
        $plan.Add([PSCustomObject]@{ Episode = $episode; FileName = $fileName; IdentityKey = $key; IdentityHash = $hash })
    }
    return ,($plan.ToArray())
}
