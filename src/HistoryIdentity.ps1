#requires -Version 5.1

# Publisher identifiers and request URLs remain exact in memory. State stores
# only full SHA-256 fingerprints, scoped to the established local feed identity.
function Get-PodcastEpisodeIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Episode,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{64}$')][string]$FeedId
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$Episode.Guid)) {
        $source = 'rss-guid'
        $value = [string]$Episode.Guid
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$Episode.AtomId)) {
        $source = 'atom-id'
        $value = [string]$Episode.AtomId
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$Episode.Url)) {
        $source = 'media-url'
        $value = [string]$Episode.Url
    }
    else { throw 'An episode requires an RSS GUID, Atom ID or exact media URL for durable identity.' }

    $key = $source + ':' + $value
    $fingerprint = Get-PodcastNameHash -IdentityKey $key
    $identifier = Get-PodcastNameHash -IdentityKey ('episode:' + $FeedId + ':' + $fingerprint)
    [PSCustomObject]@{ Id = $identifier; Source = $source; Fingerprint = $fingerprint; Key = $key }
}

function Resolve-PodcastArchive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$FeedUrl,
        [AllowNull()][AllowEmptyString()][string]$FeedTitle,
        [ValidateRange(1, 255)][int]$MaxFolderLength = 120
    )

    $canonicalRoot = Assert-PodcastDestination -Root $Root -Directory
    $feedFingerprint = Get-PodcastNameHash -IdentityKey ('feed:' + $FeedUrl)
    $archive = $null
    $establishedFolders = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($null -ne (Get-PodcastPathAttribute -LiteralPath $canonicalRoot)) {
        foreach ($directory in [IO.Directory]::EnumerateDirectories($canonicalRoot)) {
            $folderName = [IO.Path]::GetFileName($directory)
            $attributes = Get-PodcastPathAttribute -LiteralPath $directory
            if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw 'Archive discovery cannot safely inspect an output subdirectory that is a reparse point.'
            }
            # Read every local archive before choosing one: an unreadable state
            # could contain this exact alias, so creating a new archive would be
            # an unsafe guess. The reader preserves and reports corrupt state.
            $state = Read-PodcastHistory -Root $directory
            if ($null -eq $state) { continue }
            $null = $establishedFolders.Add($folderName)
            $hasAlias = $false
            foreach ($aliasFingerprint in $state.feed_alias_fingerprints) {
                if ([StringComparer]::Ordinal.Equals([string]$aliasFingerprint, $feedFingerprint)) { $hasAlias = $true }
            }
            if (-not $hasAlias) { continue }
            if ($null -ne $archive) {
                throw 'The exact feed alias belongs to multiple local archives; resolve the conflicting histories before downloading.'
            }
            $null = Assert-PodcastDestination -Root $canonicalRoot -RelativePath $folderName -Directory
            $archive = [PSCustomObject]@{ FolderName = $folderName; FeedId = [string]$state.feed_id; State = $state }
        }
    }
    if ($null -ne $archive) { return $archive }

    $newFolder = New-PodcastFolderName -FeedTitle $FeedTitle -FeedUrl $FeedUrl -MaxLength $MaxFolderLength
    if ($establishedFolders.Contains($newFolder)) {
        throw 'The proposed archive folder already belongs to another feed identity; no new history was created.'
    }
    $null = Assert-PodcastDestination -Root $canonicalRoot -RelativePath $newFolder -Directory
    [PSCustomObject]@{ FolderName = $newFolder; FeedId = $feedFingerprint; State = $null }
}

function New-PodcastHistoryPlan {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This helper returns an in-memory plan and has no side effects.')]
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Episodes,
        [Parameter(Mandatory)]$State,
        [ValidateRange(1, 255)][int]$MaxFileNameLength = 180
    )

    if ([string]$State.feed_id -cnotmatch '^[0-9a-f]{64}$') { throw 'History needs a valid local feed identity before destinations can be planned.' }
    $records = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    $names = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $State.episodes) {
        if ($records.ContainsKey([string]$record.episode_id)) { throw 'History contains duplicate episode identities.' }
        $records.Add([string]$record.episode_id, $record)
        $relative = ([string]$record.relative_path).Replace('/', '\')
        foreach ($component in $relative.Split([char]92)) { Assert-PodcastPathComponent -Component $component }
        if ($names.ContainsKey($relative)) { throw 'History destinations collide under Windows case-insensitive comparison.' }
        $names.Add($relative, [string]$record.episode_id)
    }

    $identities = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $hashes = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $fingerprints = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $plan = [Collections.Generic.List[object]]::new()
    foreach ($episode in $Episodes) {
        if ($null -eq $episode) { throw 'A destination cannot be planned for an empty episode.' }
        $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $State.feed_id
        $metadata = Get-EpisodeMetadataKey -Episode $episode
        if ($identities.ContainsKey($identity.Key)) {
            if (-not [StringComparer]::Ordinal.Equals($identities[$identity.Key], $metadata)) {
                throw 'Conflicting episode metadata reuses one identity in this feed snapshot; no media destinations were created.'
            }
            continue
        }
        if (($hashes.ContainsKey($identity.Id) -and -not [StringComparer]::Ordinal.Equals($hashes[$identity.Id], $identity.Key)) -or
            ($fingerprints.ContainsKey($identity.Fingerprint) -and -not [StringComparer]::Ordinal.Equals($fingerprints[$identity.Fingerprint], $identity.Key))) {
            throw 'Distinct episode identities have a hash collision; no media destinations were created.'
        }
        $record = if ($records.ContainsKey($identity.Id)) { $records[$identity.Id] } else { $null }
        if ($null -ne $record) {
            if (-not [StringComparer]::Ordinal.Equals([string]$record.identity_source, $identity.Source) -or
                -not [StringComparer]::Ordinal.Equals([string]$record.identity_fingerprint, $identity.Fingerprint)) {
                throw 'An episode identity conflicts with its recorded fingerprint; no media destinations were created.'
            }
            $fileName = [string]$record.relative_path
        }
        else {
            $fileName = New-EpisodeFileName -Episode $episode -MaxLength $MaxFileNameLength -IdentityHash $identity.Id
            $nameKey = $fileName.Replace('/', '\')
            if ($names.ContainsKey($nameKey)) {
                throw 'Episode destinations collide under Windows case-insensitive comparison; no media destinations were created.'
            }
            $names.Add($nameKey, $identity.Id)
        }
        $identities.Add($identity.Key, $metadata)
        $hashes.Add($identity.Id, $identity.Key)
        $fingerprints.Add($identity.Fingerprint, $identity.Key)
        $plan.Add([PSCustomObject]@{
            Episode = $episode
            FileName = $fileName
            EpisodeId = $identity.Id
            IdentitySource = $identity.Source
            IdentityFingerprint = $identity.Fingerprint
            StateRecord = $record
            MaxFileNameLength = $MaxFileNameLength
        })
    }
    return ,($plan.ToArray())
}
