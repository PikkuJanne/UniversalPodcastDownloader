#requires -Version 5.1

function Assert-PodcastHistoryProperty {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Value, [Parameter(Mandatory)][string[]]$Names)

    if ($Value -isnot [pscustomobject]) { throw 'History contains an invalid object.' }
    $actual = @($Value.PSObject.Properties.Name)
    if ($actual.Count -ne $Names.Count) { throw 'History contains missing or unsupported fields.' }
    foreach ($name in $actual) {
        if ($Names -cnotcontains $name) { throw 'History contains missing or unsupported fields.' }
    }
}

function Test-PodcastHistoryInteger {
    [CmdletBinding()]
    param($Value)

    return ($Value -is [int] -or $Value -is [long])
}

function Assert-PodcastHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$State, [Parameter(Mandatory)][string]$Root, [switch]$AllowInitial)

    Assert-PodcastHistoryProperty -Value $State -Names @('schema_version', 'feed_id', 'feed_alias_fingerprints', 'generation', 'episodes')
    if (-not (Test-PodcastHistoryInteger $State.schema_version) -or @(1, 2) -notcontains $State.schema_version) {
        throw 'Unsupported history schema version. Preserve the history and use a compatible downloader.'
    }
    if ($State.feed_id -isnot [string] -or $State.feed_id -cnotmatch '^[a-f0-9]{64}$') { throw 'History has an invalid feed identity.' }
    $minimum = if ($AllowInitial) { 0 } else { 1 }
    if (-not (Test-PodcastHistoryInteger $State.generation) -or $State.generation -lt $minimum -or $State.generation -eq [long]::MaxValue) {
        throw 'History has an invalid generation.'
    }
    if ($State.feed_alias_fingerprints -isnot [array] -or $State.feed_alias_fingerprints.Count -lt 1 -or $State.feed_alias_fingerprints.Count -gt 128) {
        throw 'History has an invalid feed alias list.'
    }
    $aliases = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    foreach ($alias in $State.feed_alias_fingerprints) {
        if ($alias -isnot [string] -or $alias -cnotmatch '^[a-f0-9]{64}$' -or -not $aliases.Add($alias)) { throw 'History contains an invalid or repeated feed alias.' }
    }
    if (-not $aliases.Contains($State.feed_id)) { throw 'History is missing its original feed alias.' }
    if ($State.episodes -isnot [array] -or $State.episodes.Count -gt 100000) { throw 'History has an invalid episode list.' }
    $identities = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    $destinations = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($episode in $State.episodes) {
        Assert-PodcastHistoryProperty -Value $episode -Names @('episode_id', 'identity_source', 'identity_fingerprint', 'relative_path', 'status', 'bytes', 'local_sha256', 'completed_utc', 'verification')
        if ($episode.episode_id -isnot [string] -or $episode.episode_id -cnotmatch '^[a-f0-9]{64}$' -or
            $episode.identity_fingerprint -isnot [string] -or $episode.identity_fingerprint -cnotmatch '^[a-f0-9]{64}$' -or
            $episode.identity_source -isnot [string] -or @('rss-guid', 'atom-id', 'media-url') -cnotcontains $episode.identity_source -or -not $identities.Add($episode.episode_id)) {
            throw 'History contains an invalid or repeated episode identity.'
        }
        $expectedIdentity = Get-PodcastNameHash -IdentityKey ('episode:' + $State.feed_id + ':' + $episode.identity_fingerprint)
        if ($episode.episode_id -cne $expectedIdentity) { throw 'History episode identity does not match its feed and fingerprint.' }
        if ($episode.relative_path -isnot [string]) { throw 'History contains an invalid relative destination.' }
        Assert-PodcastPathComponent -Component $episode.relative_path
        if ($episode.relative_path -ieq '.upd' -or -not $destinations.Add($episode.relative_path)) { throw 'History contains colliding relative destinations.' }
        $null = Get-PodcastDestination -Root $Root -RelativePath $episode.relative_path
        $allowedStatuses = @('prepared', 'transfer_verified', 'missing', 'conflict', 'failed')
        if ($State.schema_version -eq 2) { $allowedStatuses += @('adopted', 'unverified') }
        if ($episode.status -isnot [string] -or $allowedStatuses -cnotcontains $episode.status) { throw 'History contains an unsupported episode status.' }
        if ($null -ne $episode.bytes -and (-not (Test-PodcastHistoryInteger $episode.bytes) -or $episode.bytes -lt 0)) { throw 'History contains an invalid byte count.' }
        if ($null -ne $episode.local_sha256 -and ($episode.local_sha256 -isnot [string] -or $episode.local_sha256 -cnotmatch '^[a-f0-9]{64}$')) { throw 'History contains an invalid local digest.' }
        if ($null -ne $episode.completed_utc) {
            $parsedUtc = [datetime]::MinValue
            if ($episode.completed_utc -isnot [string] -or $episode.completed_utc -cnotmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?Z$' -or
                -not [datetime]::TryParse($episode.completed_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsedUtc)) {
                throw 'History contains an invalid UTC completion timestamp.'
            }
        }
        if (@('prepared', 'transfer_verified', 'adopted') -ccontains $episode.status) {
            if ($null -eq $episode.bytes -or $episode.bytes -le 0 -or $null -eq $episode.local_sha256 -or $null -eq $episode.completed_utc) {
                throw 'History completion evidence is incomplete.'
            }
        }
        Assert-PodcastHistoryProperty -Value $episode.verification -Names @('method', 'media_kind', 'notes')
        if ($episode.verification.method -isnot [string] -or $episode.verification.method -cnotmatch '^[a-z0-9-]{1,64}$' -or
            ($null -ne $episode.verification.media_kind -and ($episode.verification.media_kind -isnot [string] -or $episode.verification.media_kind -cnotmatch '^[a-z0-9-]{1,32}$'))) {
            throw 'History contains invalid verification evidence.'
        }
        if (@('prepared', 'transfer_verified') -ccontains $episode.status) {
            if (@('completed-http-length-and-signature', 'completed-eof-and-signature') -cnotcontains $episode.verification.method -or
                @('mpeg-audio', 'wave', 'flac', 'ogg-container', 'mp4-container') -cnotcontains $episode.verification.media_kind) {
                throw 'History completion evidence uses an unsupported verification method or media kind.'
            }
        }
        if ($episode.verification.notes -isnot [array] -or $episode.verification.notes.Count -gt 16) { throw 'History contains invalid verification notes.' }
        foreach ($note in $episode.verification.notes) {
            if ($note -isnot [string] -or $note.Length -gt 256 -or $note -match '://|@|[\x00-\x1f]') { throw 'History contains unsafe or overlong verification notes.' }
        }
        if ($episode.status -ceq 'adopted') {
            if ($episode.verification.method -cne 'owner-approved-local-signature' -or
                @('mpeg-audio', 'wave', 'flac', 'ogg-container', 'mp4-container') -cnotcontains $episode.verification.media_kind -or
                $episode.verification.notes -cnotcontains 'local-signature-only; transfer-completeness-unverified') {
                throw 'History adoption evidence must record owner approval and unverified transfer completeness.'
            }
        }
    }
}

function Assert-PodcastHistoryJson {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Json)

    # Bound nesting before ConvertFrom-Json. Reject duplicate/escaped property
    # names so runtimes cannot silently interpret the same document differently.
    if (-not $Json.TrimStart().StartsWith('{', [StringComparison]::Ordinal)) { throw 'History JSON must contain one object.' }
    $stack = New-Object 'System.Collections.Generic.Stack[object]'
    foreach ($token in [regex]::Matches($Json, '"(?:\\.|[^"\\])*"\s*:?|[{}\[\]]')) {
        $value = $token.Value
        if ($value -eq '{' -or $value -eq '[') {
            if ($stack.Count -ge 12) { throw 'History JSON exceeds its nesting limit.' }
            $stack.Push([pscustomobject]@{ Kind = $value; Names = (New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)) })
        }
        elseif ($value -eq '}' -or $value -eq ']') {
            if ($stack.Count -eq 0) { throw 'History JSON has invalid nesting.' }
            $expected = if ($value -eq '}') { '{' } else { '[' }
            if ($stack.Pop().Kind -ne $expected) { throw 'History JSON has invalid nesting.' }
        }
        elseif ($value.EndsWith(':', [StringComparison]::Ordinal)) {
            if ($stack.Count -eq 0 -or $stack.Peek().Kind -ne '{' -or $value -cnotmatch '^"([a-z_][a-z0-9_]*)"\s*:$') {
                throw 'History JSON has an invalid property name.'
            }
            if (-not $stack.Peek().Names.Add($Matches[1])) { throw 'History JSON has a repeated property name.' }
            if ($Matches[1] -ceq 'completed_utc') {
                $offset = $token.Index + $token.Length
                $timestampJson = $Json.Substring($offset, [Math]::Min(80, $Json.Length - $offset))
                if ($timestampJson -cnotmatch '^\s*(?:null(?=\s*[,}])|"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?Z")') {
                    throw 'History contains an invalid UTC completion timestamp.'
                }
            }
        }
    }
    if ($stack.Count -ne 0) { throw 'History JSON has invalid nesting.' }
}

function Read-PodcastHistoryFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$RelativePath)

    $path = Assert-PodcastDestination -Root $Root -RelativePath $RelativePath
    if ($null -eq (Get-PodcastPathAttribute -LiteralPath $path)) { return $null }
    $stream = $null
    $reader = $null
    try {
        $stream = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        if ($stream.Length -eq 0 -or $stream.Length -gt 16777216) { throw 'History exceeds the supported 16 MiB size limit or is empty.' }
        $encoding = New-Object Text.UTF8Encoding($false, $true)
        $reader = New-Object IO.StreamReader($stream, $encoding, $false)
        $json = $reader.ReadToEnd()
        Assert-PodcastHistoryJson -Json $json
        $parseParameters = @{ InputObject = $json; ErrorAction = 'Stop' }
        if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $parseParameters.DateKind = 'String' }
        $state = ConvertFrom-Json @parseParameters
        # Earlier PS7 releases eagerly deserialize ISO timestamps. Raw JSON
        # timestamps were checked above; retain this schema's string contract.
        foreach ($episode in $state.episodes) {
            if ($episode.completed_utc -is [datetime]) {
                $episode.completed_utc = $episode.completed_utc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ', [Globalization.CultureInfo]::InvariantCulture)
            }
        }
        Assert-PodcastHistory -State $state -Root $Root
        return $state
    }
    catch { throw 'History is corrupt, unsupported, or inaccessible. Preserved the state files; inspect them before retrying.' }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $stream) { $stream.Dispose() }
    }
}

function Read-PodcastHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $state = Read-PodcastHistoryFile -Root $Root -RelativePath '.upd\state.json'
    $backup = Read-PodcastHistoryFile -Root $Root -RelativePath '.upd\state.json.bak'
    if ($null -eq $state -and $null -ne $backup) { throw 'History primary is missing while a backup exists. Preserve both paths and inspect before retrying.' }
    if ($null -ne $state -and $null -ne $backup -and ($state.feed_id -cne $backup.feed_id -or $backup.generation -ge $state.generation -or $backup.schema_version -gt $state.schema_version)) {
        throw 'History backup contradicts the current state. Preserve both files and inspect before retrying.'
    }
    return $state
}

function New-PodcastHistory {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Constructs an in-memory value; does not write state.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FeedId,
        [Parameter(Mandatory)][string]$FeedAliasFingerprint,
        [ValidateSet(1, 2)][int]$SchemaVersion = 1
    )

    return [pscustomobject]@{
        schema_version = $SchemaVersion
        feed_id = $FeedId
        feed_alias_fingerprints = @($FeedId, $FeedAliasFingerprint | Select-Object -Unique)
        generation = 0
        episodes = @()
    }
}

function ConvertTo-PodcastHistoryV2 {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$State, [Parameter(Mandatory)][string]$Root)

    # Migration is an explicit in-memory operation. Normal history writes retain
    # their version, and this copy leaves the source and its generation intact.
    Assert-PodcastHistory -State $State -Root $Root -AllowInitial
    $episodes = @(
        foreach ($episode in $State.episodes) {
            [pscustomobject]@{
                episode_id = $episode.episode_id
                identity_source = $episode.identity_source
                identity_fingerprint = $episode.identity_fingerprint
                relative_path = $episode.relative_path
                status = $episode.status
                bytes = $episode.bytes
                local_sha256 = $episode.local_sha256
                completed_utc = $episode.completed_utc
                verification = [pscustomobject]@{
                    method = $episode.verification.method
                    media_kind = $episode.verification.media_kind
                    notes = @($episode.verification.notes)
                }
            }
        }
    )
    $promoted = [pscustomobject]@{
        schema_version = 2
        feed_id = $State.feed_id
        feed_alias_fingerprints = @($State.feed_alias_fingerprints)
        generation = $State.generation
        episodes = $episodes
    }
    Assert-PodcastHistory -State $promoted -Root $Root -AllowInitial
    return $promoted
}

function Enter-PodcastHistoryLock {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $canonicalRoot = Assert-PodcastDestination -Root $Root -Directory
    $metadataRoot = Assert-PodcastDestination -Root $canonicalRoot -RelativePath '.upd' -Directory
    $null = [IO.Directory]::CreateDirectory($metadataRoot)
    $path = Assert-PodcastDestination -Root $canonicalRoot -RelativePath '.upd\writer.lock'
    $stream = $null
    try {
        $stream = New-Object IO.FileStream($path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    }
    catch { throw 'Cannot acquire the archive writer lock. Another downloader may be active or the folder may be inaccessible.' }
    try {
        $state = Read-PodcastHistory -Root $canonicalRoot
        $generation = if ($null -eq $state) { 0 } else { $state.generation }
        $lock = [pscustomobject]@{ Root = $canonicalRoot; Stream = $stream; Generation = $generation }
        $lock.PSObject.TypeNames.Insert(0, 'UPD.PodcastHistoryLock')
        return $lock
    }
    catch { $stream.Dispose(); throw }
}

function Write-PodcastHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Lock, [Parameter(Mandatory)]$State)

    if ($Lock.PSObject.TypeNames -notcontains 'UPD.PodcastHistoryLock' -or $Lock.Stream -isnot [IO.FileStream] -or -not $Lock.Stream.CanWrite) {
        throw 'A held archive writer lock is required to update history.'
    }
    $lockPath = Assert-PodcastDestination -Root $Lock.Root -RelativePath '.upd\writer.lock'
    if (-not [string]::Equals($Lock.Stream.Name, $lockPath, [StringComparison]::OrdinalIgnoreCase)) { throw 'The writer lock does not belong to this archive.' }
    Assert-PodcastHistory -State $State -Root $Lock.Root
    $current = Read-PodcastHistory -Root $Lock.Root
    $generation = if ($null -eq $current) { 0 } else { $current.generation }
    if ($Lock.Generation -ne $generation -or $State.generation -ne ($generation + 1)) { throw 'History generation changed; refusing a stale update.' }
    if ($null -ne $current -and $State.feed_id -cne $current.feed_id) { throw 'History cannot change the archive feed identity.' }
    if ($null -ne $current -and $State.schema_version -lt $current.schema_version) { throw 'History cannot downgrade its schema version.' }
    $json = ConvertTo-Json -InputObject $State -Depth 8 -Compress
    $encoding = New-Object Text.UTF8Encoding($false, $true)
    $bytes = $encoding.GetBytes($json)
    if ($bytes.Length -gt 16777216) { throw 'History exceeds the supported 16 MiB size limit.' }
    $relativeTemporary = '.upd\state.' + [guid]::NewGuid().ToString('N') + '.tmp'
    $temporary = Assert-PodcastDestination -Root $Lock.Root -RelativePath $relativeTemporary
    $writer = $null
    $ownedTemporary = $false
    try {
        $writer = New-Object IO.FileStream($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $ownedTemporary = $true
        $writer.Write($bytes, 0, $bytes.Length)
        $writer.Flush($true)
        $writer.Dispose()
        $writer = $null
        $destination = Assert-PodcastDestination -Root $Lock.Root -RelativePath '.upd\state.json'
        $backupPath = Assert-PodcastDestination -Root $Lock.Root -RelativePath '.upd\state.json.bak'
        $null = Assert-PodcastDestination -Root $Lock.Root -RelativePath $relativeTemporary
        if ($null -eq $current) {
            [IO.File]::Move($temporary, $destination)
        }
        else {
            # Same-volume replacement keeps the preceding validated generation.
            # Filesystems that cannot replace atomically fail without fallback.
            [IO.File]::Replace($temporary, $destination, $backupPath, $true)
        }
        $ownedTemporary = $false
        $Lock.Generation = $State.generation
        return $State
    }
    finally {
        if ($null -ne $writer) { $writer.Dispose() }
        if ($ownedTemporary) {
            $safeTemporary = Assert-PodcastDestination -Root $Lock.Root -RelativePath $relativeTemporary
            [IO.File]::Delete($safeTemporary)
        }
    }
}
