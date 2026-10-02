# File evidence and the download/history commit protocol. No import-time effects.
function Enter-PodcastArchiveLock {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $path = Assert-PodcastDestination -Root $Root -RelativePath '.upd-archive.lock'
    try { return [IO.File]::Open($path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
    catch { throw 'Another archive selection writer is active, or the output lock is inaccessible. Retry after that run finishes.' }
}

function Get-PodcastFileEvidence {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$RelativePath)

    $path = Assert-PodcastDestination -Root $Root -RelativePath $RelativePath
    if ($null -eq (Get-PodcastPathAttribute -LiteralPath $path)) { return $null }
    # Permit concurrent readers, but deny modification/deletion while hashing.
    $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $null = Assert-PodcastDestination -Root $Root -RelativePath $RelativePath
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $digest = ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() }
        finally { $sha.Dispose() }
        return [pscustomobject]@{ Bytes = $stream.Length; Sha256 = $digest }
    }
    finally { $stream.Dispose() }
}

function New-PodcastEpisodeRecord {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory history record only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Planned)

    return [pscustomobject]@{
        episode_id = $Planned.EpisodeId
        identity_source = $Planned.IdentitySource
        identity_fingerprint = $Planned.IdentityFingerprint
        relative_path = $Planned.FileName
        status = 'failed'
        bytes = $null
        local_sha256 = $null
        completed_utc = $null
        verification = [pscustomobject]@{ method = 'none'; media_kind = $null; notes = @() }
    }
}

function Save-PodcastEpisodeRecord {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Required transactional metadata update within the caller-held archive lock.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Record)

    $next = [pscustomobject]@{
        schema_version = $Context.State.schema_version
        feed_id = $Context.State.feed_id
        feed_alias_fingerprints = @($Context.State.feed_alias_fingerprints)
        generation = $Context.State.generation + 1
        episodes = @(@($Context.State.episodes | Where-Object { $_.episode_id -cne $Record.episode_id }) + @($Record))
    }
    $Context.State = Write-PodcastHistory -Lock $Context.Lock -State $next
}

function Resolve-PodcastHistoryItem {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Planned)

    $record = $Planned.StateRecord
    $evidence = Get-PodcastFileEvidence -Root $Context.Lock.Root -RelativePath $Planned.FileName
    if ($null -eq $record) {
        if ($null -eq $evidence) { return 'download' }
        $record = New-PodcastEpisodeRecord -Planned $Planned
        $record.status = 'conflict'
        Save-PodcastEpisodeRecord -Context $Context -Record $record
        return 'conflict'
    }
    # Clone the record so a failed commit cannot change the last read state.
    $record = $record.PSObject.Copy()
    if ($record.status -eq 'adopted') {
        if ($null -ne $evidence -and $record.bytes -eq $evidence.Bytes -and $record.local_sha256 -ceq $evidence.Sha256) {
            return 'adopted_skip'
        }
        # Adoption never becomes transfer evidence or permission to redownload.
        return 'conflict'
    }
    if ($record.status -eq 'unverified') { return 'conflict' }
    if ($null -eq $evidence) {
        if ($record.status -ne 'missing') {
            $record.status = 'missing'
            Save-PodcastEpisodeRecord -Context $Context -Record $record
        }
        return 'download'
    }
    if ($record.status -in @('prepared', 'transfer_verified', 'missing') -and
        $null -ne $record.local_sha256 -and $record.bytes -eq $evidence.Bytes -and
        $record.local_sha256 -ceq $evidence.Sha256) {
        if ($record.status -ne 'transfer_verified') {
            $record.status = 'transfer_verified'
            Save-PodcastEpisodeRecord -Context $Context -Record $record
        }
        return 'verified_skip'
    }
    if ($record.status -ne 'conflict') {
        $record.status = 'conflict'
        Save-PodcastEpisodeRecord -Context $Context -Record $record
    }
    return 'conflict'
}

function Invoke-PodcastRecordedTransfer {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Planned)

    $beforeFinalize = {
        param($Evidence)
        $prepared = New-PodcastEpisodeRecord -Planned $Planned
        $prepared.status = 'prepared'
        $prepared.bytes = $Evidence.Bytes
        $prepared.local_sha256 = $Evidence.Sha256
        $prepared.completed_utc = [datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        $prepared.verification = [pscustomobject]@{
            method = $Evidence.Verification.Replace('_', '-')
            media_kind = $Evidence.DetectedFormat.ToLowerInvariant().Replace('_', '-')
            notes = @($Evidence.Warnings)
        }
        Save-PodcastEpisodeRecord -Context $Context -Record $prepared
    }
    $result = Invoke-PodcastMediaTransfer -Uri $Planned.Episode.Url -Root $Context.Lock.Root `
        -RelativePath $Planned.FileName -EnclosureLength $Planned.Episode.EnclosureLength -BeforeFinalize $beforeFinalize
    $completed = @($Context.State.episodes | Where-Object { $_.episode_id -ceq $Planned.EpisodeId })[0].PSObject.Copy()
    $completed.status = 'transfer_verified'
    Save-PodcastEpisodeRecord -Context $Context -Record $completed
    return $result
}
