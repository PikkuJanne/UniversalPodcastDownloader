#requires -Version 5.1

function Assert-PodcastResumeLock {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Lock)
    if ($Lock.PSObject.TypeNames -notcontains 'UPD.PodcastHistoryLock' -or
        $Lock.Stream -isnot [IO.FileStream] -or -not $Lock.Stream.CanWrite) {
        throw 'A held archive writer lock is required to update resume evidence.'
    }
    $lockPath = Assert-PodcastDestination -Root $Lock.Root -RelativePath '.upd\writer.lock'
    if (-not [string]::Equals($Lock.Stream.Name, $lockPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The resume writer lock does not belong to this archive.'
    }
}

function Get-PodcastResumeStatePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{64}$')][string]$EpisodeId)
    return (Assert-PodcastDestination -Root $Root -RelativePath ('.upd\resume-' + $EpisodeId.Substring(0, 32) + '.json'))
}

function Assert-PodcastResumeState {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$State, [Parameter(Mandatory)][string]$Root)
    Assert-PodcastHistoryProperty -Value $State -Names @('schema_version', 'feed_id', 'episode_id', 'relative_path', 'partial_name',
        'request_fingerprint', 'final_uri_fingerprint', 'etag', 'total_length', 'content_type', 'content_encoding', 'offset', 'prefix_sha256')
    if (-not (Test-PodcastHistoryInteger $State.schema_version) -or $State.schema_version -ne 1) { throw 'Unsupported resume schema.' }
    foreach ($field in @('feed_id', 'episode_id', 'request_fingerprint', 'final_uri_fingerprint', 'prefix_sha256')) {
        if ($State.$field -isnot [string] -or $State.$field -cnotmatch '^[a-f0-9]{64}$') { throw 'Invalid resume identity or digest.' }
    }
    if ($State.relative_path -isnot [string] -or $State.partial_name -isnot [string] -or
        $State.partial_name -cnotmatch '^\.upd-[a-f0-9]{32}\.tmp$') { throw 'Invalid resume destination.' }
    Assert-PodcastPathComponent -Component $State.relative_path
    if ($State.relative_path -ieq '.upd' -or $State.relative_path -imatch '^\.upd-') { throw 'Invalid resume destination.' }
    $null = Assert-PodcastDestination -Root $Root -RelativePath $State.relative_path
    $null = Assert-PodcastDestination -Root $Root -RelativePath $State.partial_name
    if ($State.etag -isnot [string] -or -not (Test-PodcastStrongETag $State.etag) -or
        $State.content_type -isnot [string] -or $State.content_type.Length -gt 1024 -or
        [string]::IsNullOrEmpty($State.content_type) -or
        (Get-PodcastResumeContentType -ContentType $State.content_type) -cne $State.content_type -or
        $State.content_type -cmatch '^multipart/' -or $State.content_encoding -isnot [string] -or
        $State.content_encoding -cne 'identity') { throw 'Invalid resume representation.' }
    if (-not (Test-PodcastHistoryInteger $State.total_length) -or $State.total_length -le 0 -or
        -not (Test-PodcastHistoryInteger $State.offset) -or $State.offset -lt 0 -or $State.offset -gt $State.total_length) {
        throw 'Invalid resume byte checkpoint.'
    }
}

function Read-PodcastResumeState {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$EpisodeId)
    $path = Get-PodcastResumeStatePath -Root $Root -EpisodeId $EpisodeId
    if ($null -eq (Get-PodcastPathAttribute -LiteralPath $path)) { return $null }
    $reader = $null
    $stream = $null
    try {
        $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        if ($stream.Length -le 0 -or $stream.Length -gt 16384) { throw 'Invalid resume file size.' }
        $reader = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false, $true), $false)
        $json = $reader.ReadToEnd()
        Assert-PodcastHistoryJson -Json $json
        $state = ConvertFrom-Json -InputObject $json -ErrorAction Stop
        Assert-PodcastResumeState -State $state -Root $Root
        if ($state.episode_id -cne $EpisodeId) { throw 'Resume identity mismatch.' }
        return $state
    }
    catch { throw 'Resume evidence is corrupt, unsupported, or inaccessible; partial and sidecar were preserved for review.' }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $stream) { $stream.Dispose() }
    }
}

function Write-PodcastResumeState {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Lock, [Parameter(Mandatory)]$State, [AllowNull()]$Expected)
    Assert-PodcastResumeLock -Lock $Lock
    Assert-PodcastResumeState -State $State -Root $Lock.Root
    $current = Read-PodcastResumeState -Root $Lock.Root -EpisodeId $State.episode_id
    $currentJson = if ($null -ne $current) { ConvertTo-Json -InputObject $current -Compress } else { '' }
    $expectedJson = if ($null -ne $Expected) { ConvertTo-Json -InputObject $Expected -Compress } else { '' }
    if ($currentJson -cne $expectedJson) { throw 'Resume evidence changed; preserved the files instead of replacing it.' }
    if ($null -ne $current -and ($current.feed_id -cne $State.feed_id -or $current.relative_path -cne $State.relative_path)) {
        throw 'Resume evidence cannot change the episode archive or destination.'
    }
    $json = ConvertTo-Json -InputObject $State -Compress
    $bytes = [Text.UTF8Encoding]::new($false, $true).GetBytes($json)
    if ($bytes.Length -gt 16384) { throw 'Resume evidence exceeds its size limit.' }
    $temporaryRelative = '.upd\resume-' + [guid]::NewGuid().ToString('N') + '.tmp'
    $temporary = Assert-PodcastDestination -Root $Lock.Root -RelativePath $temporaryRelative
    $writer = $null
    $owned = $false
    try {
        $writer = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $owned = $true
        $writer.Write($bytes, 0, $bytes.Length)
        $writer.Flush($true)
        $writer.Dispose()
        $writer = $null
        $destination = Get-PodcastResumeStatePath -Root $Lock.Root -EpisodeId $State.episode_id
        $null = Assert-PodcastDestination -Root $Lock.Root -RelativePath $temporaryRelative
        if ($null -eq $current) { [IO.File]::Move($temporary, $destination) }
        else {
            $backup = [Management.Automation.Language.NullString]::Value
            if ($current.partial_name -cne $State.partial_name) {
                # Replacing the active transfer retains the exact old sidecar
                # together with its old partial, including changed entities.
                $backup = Assert-PodcastDestination -Root $Lock.Root -RelativePath ('.upd\resume-' + [guid]::NewGuid().ToString('N') + '.old.json')
                if ($null -ne (Get-PodcastPathAttribute -LiteralPath $backup)) { throw 'Resume snapshot destination already exists.' }
            }
            [IO.File]::Replace($temporary, $destination, $backup, $true)
        }
        $owned = $false
        return $State
    }
    finally {
        if ($null -ne $writer) { $writer.Dispose() }
        if ($owned) {
            $null = Assert-PodcastDestination -Root $Lock.Root -RelativePath $temporaryRelative
            [IO.File]::Delete($temporary)
        }
    }
}

function Get-PodcastResumeStreamHash {
    [CmdletBinding()]
    param([Parameter(Mandatory)][IO.Stream]$Stream)
    $position = $Stream.Position
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $Stream.Position = 0
        return ([BitConverter]::ToString($sha.ComputeHash($Stream))).Replace('-', '').ToLowerInvariant()
    }
    finally { $Stream.Position = $position; $sha.Dispose() }
}

function Open-PodcastResumePartial {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Lock, [Parameter(Mandatory)]$State,
        [Parameter(Mandatory)][string]$FeedId, [Parameter(Mandatory)][string]$EpisodeId,
        [Parameter(Mandatory)][string]$RelativePath, [Parameter(Mandatory)][string]$RequestFingerprint)
    Assert-PodcastResumeLock -Lock $Lock
    Assert-PodcastResumeState -State $State -Root $Lock.Root
    if ($State.feed_id -cne $FeedId -or $State.episode_id -cne $EpisodeId -or $State.relative_path -cne $RelativePath) {
        throw 'Resume identity does not match this episode; partial and sidecar were preserved for review.'
    }
    $path = Assert-PodcastDestination -Root $Lock.Root -RelativePath $State.partial_name
    $stream = $null
    try {
        $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        if ($stream.Length -ne $State.offset -or (Get-PodcastResumeStreamHash -Stream $stream) -cne $State.prefix_sha256) {
            throw 'Resume checkpoint does not match the partial.'
        }
        if ($State.request_fingerprint -cne $RequestFingerprint -or $State.offset -eq 0 -or $State.offset -eq $State.total_length) {
            $stream.Dispose()
            return $null
        }
        $stream.Position = $State.offset
        $result = $stream
        $stream = $null
        return $result
    }
    catch { throw 'Resume checkpoint is inconsistent; partial and sidecar were preserved for review.' }
    finally { if ($null -ne $stream) { $stream.Dispose() } }
}

function Update-PodcastResumeCheckpoint {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Writes owned transactional metadata within the caller-confirmed download and held archive lock.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Session, [switch]$Force)
    if ($null -eq $Session.State -or $Session.State.partial_name -cne $Session.PartialName) { return }
    [long]$length = $Session.Stream.Length
    if (-not $Force -and $Session.State.offset -gt 0 -and
        [double]$length -lt [Math]::Max(8388608.0, 2.0 * [double]$Session.State.offset)) { return }
    $null = Assert-PodcastDestination -Root $Session.Lock.Root -RelativePath $Session.PartialName
    $Session.Stream.Flush($true)
    $next = $Session.State.PSObject.Copy()
    $next.offset = $length
    $next.prefix_sha256 = Get-PodcastResumeStreamHash -Stream $Session.Stream
    $Session.State = Write-PodcastResumeState -Lock $Session.Lock -State $next -Expected $Session.State
}

function Remove-PodcastResumeState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Removes only verified metadata for the completed owned transfer after the caller confirmed the operation.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Lock, [Parameter(Mandatory)]$State)
    Assert-PodcastResumeLock -Lock $Lock
    $current = Read-PodcastResumeState -Root $Lock.Root -EpisodeId $State.episode_id
    if ($null -eq $current) { return }
    if ((ConvertTo-Json -InputObject $current -Compress) -cne (ConvertTo-Json -InputObject $State -Compress)) {
        throw 'Resume evidence changed; preserved it for review.'
    }
    [IO.File]::Delete((Get-PodcastResumeStatePath -Root $Lock.Root -EpisodeId $State.episode_id))
}
