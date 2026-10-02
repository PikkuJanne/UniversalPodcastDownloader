#requires -Version 5.1
# Legacy names are lookup hints only. Inventory never changes files or history.

function Get-PodcastLegacyFolderName {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$FeedTitle)

    if ([string]::IsNullOrWhiteSpace($FeedTitle)) { return 'Podcast' }
    return ($FeedTitle -replace '[\\/:*?"<>|]', '_').Trim()
}

function Get-PodcastHistoricalFileName {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingBrokenHashAlgorithms', '', Justification = 'SHA-1 reproduces an old filename for lookup only; it is never integrity or identity evidence.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Episode)

    $extension = 'mp3'
    if ([string]$Episode.Url -match '\.(mp3|m4a)($|\?)') { $extension = $Matches[1] }
    $title = if ([string]::IsNullOrWhiteSpace([string]$Episode.Title)) { 'Episode' } else { [string]$Episode.Title }
    $title = ($title -replace '[\\/:*?"<>|]', '_').Trim()
    if ([string]::IsNullOrWhiteSpace($title)) { $title = 'Episode' }
    $prefix = if ($Episode.PubDate) { ([datetime]$Episode.PubDate).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) + ' - ' } else { '' }
    $suffix = ''
    if ($title -eq 'Episode' -or $prefix.Length -eq 0) {
        $oldIdentity = if ($Episode.Guid) { [string]$Episode.Guid } else { [string]$Episode.Url }
        if ($oldIdentity.Length -eq 0) { return $null }
        $sha = [Security.Cryptography.SHA1]::Create()
        try { $suffix = '-' + (([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($oldIdentity)))).Replace('-', '').ToLowerInvariant()).Substring(0, 8) }
        finally { $sha.Dispose() }
    }
    return $prefix + $title + $suffix + '.' + $extension
}

function Get-PodcastLegacyFileObservation {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$RelativePath)

    Assert-PodcastPathComponent -Component $RelativePath
    $path = Assert-PodcastDestination -Root $Root -RelativePath $RelativePath
    $result = [pscustomobject]@{
        RelativePath = $RelativePath; Bytes = $null; Sha256 = $null
        Plausible = $false; MediaKind = $null; InspectedBytes = 0
        Classification = 'conflict'; Reason = 'file_unreadable'
        CandidateEpisodeIds = @(); RecordedEpisodeId = $null; RecordedStatus = $null
    }
    $guard = $null
    try {
        # Keep an ordinary writer from replacing or changing the bytes between
        # hashing and sniffing. Both helpers open additional read-only handles.
        $guard = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $evidence = Get-PodcastFileEvidence -Root $Root -RelativePath $RelativePath
        $result.Bytes = $evidence.Bytes
        $result.Sha256 = $evidence.Sha256
        # The old validator's switch enables its bounded local sniff. There was
        # no observed transfer: discard its Verification field and never return
        # it as transfer evidence or a claim of complete decoding.
        $sniff = Test-PodcastMediaFile -LiteralPath $path -TransferCompleted $true
        $result.InspectedBytes = $sniff.InspectedBytes
        $result.MediaKind = $sniff.DetectedFormat
        if ($RelativePath -match '(?i)\.(part|partial|tmp)(?:\.|$)') {
            $result.Reason = 'partial_name'
        }
        elseif ($sniff.Valid) {
            $result.Plausible = $true
            $result.Classification = 'unverified'
            $result.Reason = 'local_signature_only'
        }
        else { $result.Reason = $sniff.Category }
    }
    catch { $result.Reason = 'file_unreadable' }
    finally { if ($null -ne $guard) { $guard.Dispose() } }
    return $result
}

function New-PodcastLegacyPlan {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This read-only inventory returns an in-memory plan and does not write files or state.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [AllowNull()][AllowEmptyCollection()][object[]]$Episodes,
        [Parameter(Mandatory)]$State,
        [ValidateRange(1, 255)][int]$MaxFileNameLength = 180
    )

    $canonicalRoot = Assert-PodcastDestination -Root $Root -Directory
    if ($null -eq (Get-PodcastPathAttribute -LiteralPath $canonicalRoot)) { throw 'Legacy inventory requires an existing ordinary show directory.' }
    $historyPlan = New-PodcastHistoryPlan -Episodes $Episodes -State $State -MaxFileNameLength $MaxFileNameLength
    $nameOwners = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $hashOwners = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $recordOwners = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $State.episodes) { $recordOwners.Add([string]$record.relative_path, $record) }
    foreach ($planned in $historyPlan) {
        $hints = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $null = $hints.Add($planned.FileName)
        $historical = Get-PodcastHistoricalFileName -Episode $planned.Episode
        if ($historical) { $null = $hints.Add($historical) }
        # Both full-hash filename generations can retain old titles, dates or
        # extensions. Their suffix is a lookup hint, still not identity proof.
        $priorHash = Get-PodcastNameHash -IdentityKey (Get-EpisodeIdentityKey -Episode $planned.Episode)
        foreach ($hash in @($priorHash, $planned.EpisodeId)) {
            if (-not $hashOwners.ContainsKey($hash)) { $hashOwners.Add($hash, [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)) }
            $null = $hashOwners[$hash].Add($planned.EpisodeId)
        }
        foreach ($hint in $hints) {
            if (-not $nameOwners.ContainsKey($hint)) { $nameOwners.Add($hint, [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)) }
            $null = $nameOwners[$hint].Add($planned.EpisodeId)
        }
    }

    $files = [Collections.Generic.List[object]]::new()
    $filesByName = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in @([IO.Directory]::EnumerateFileSystemEntries($canonicalRoot) | Sort-Object)) {
        $name = [IO.Path]::GetFileName($entry)
        $attributes = Get-PodcastPathAttribute -LiteralPath $entry
        $isReparse = ($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
        if (($attributes -band [IO.FileAttributes]::Directory) -ne 0 -and -not $isReparse) { continue }
        try {
            if ($isReparse) { throw 'Inventory does not follow reparse entries.' }
            $file = Get-PodcastLegacyFileObservation -Root $canonicalRoot -RelativePath $name
        }
        catch {
            $file = [pscustomobject]@{
                RelativePath = $name; Bytes = $null; Sha256 = $null; Plausible = $false
                MediaKind = $null; InspectedBytes = 0; Classification = 'conflict'
                Reason = 'unsafe_path'; CandidateEpisodeIds = @(); RecordedEpisodeId = $null; RecordedStatus = $null
            }
        }
        $candidateIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        if ($nameOwners.ContainsKey($name)) {
            foreach ($owner in $nameOwners[$name]) { $null = $candidateIds.Add($owner) }
        }
        if ($name -match '(?i)-([0-9a-f]{64})\.(mp3|m4a)$' -and $hashOwners.ContainsKey($Matches[1])) {
            foreach ($owner in $hashOwners[$Matches[1]]) { $null = $candidateIds.Add($owner) }
        }
        $file.CandidateEpisodeIds = @($candidateIds)
        if ($recordOwners.ContainsKey($name)) {
            $record = $recordOwners[$name]
            $file.RecordedEpisodeId = [string]$record.episode_id
            $file.RecordedStatus = [string]$record.status
            if ($record.status -in @('prepared', 'transfer_verified', 'adopted', 'missing') -and
                $null -ne $file.Sha256 -and $file.Sha256 -ceq $record.local_sha256 -and $file.Bytes -eq $record.bytes) {
                $file.Classification = 'recorded'; $file.Reason = 'matches_recorded_bytes'
            }
            else { $file.Classification = 'conflict'; $file.Reason = 'recorded_bytes_unconfirmed' }
        }
        elseif ($file.CandidateEpisodeIds.Count -gt 1) {
            $file.Classification = 'conflict'; $file.Reason = 'multiple_episode_candidates'
        }
        if ($filesByName.ContainsKey($name)) {
            $file.Classification = 'conflict'; $file.Reason = 'case_insensitive_collision'
            $filesByName[$name].Classification = 'conflict'; $filesByName[$name].Reason = 'case_insensitive_collision'
        }
        else { $filesByName.Add($name, $file) }
        $files.Add($file)
    }

    $rows = [Collections.Generic.List[object]]::new()
    foreach ($planned in $historyPlan) {
        $candidates = @($files | Where-Object { $_.CandidateEpisodeIds -ccontains $planned.EpisodeId })
        $classification = 'missing'; $reason = 'no_filename_candidate'; $suggested = $null
        if ($null -ne $planned.StateRecord) {
            if ($filesByName.ContainsKey($planned.FileName)) {
                $bound = $filesByName[$planned.FileName]
                $classification = $bound.Classification; $reason = $bound.Reason
            }
            else { $reason = 'recorded_path_missing' }
        }
        elseif ($candidates.Count -gt 1) {
            $classification = 'conflict'; $reason = 'multiple_file_candidates'
            foreach ($candidate in $candidates) {
                if ($candidate.Classification -ne 'recorded') { $candidate.Classification = 'conflict'; $candidate.Reason = 'multiple_file_candidates' }
            }
        }
        elseif ($candidates.Count -eq 1) {
            $candidate = $candidates[0]
            if ($candidate.RecordedEpisodeId) { $classification = 'conflict'; $reason = 'path_owned_by_another_episode' }
            elseif ($candidate.CandidateEpisodeIds.Count -ne 1) { $classification = 'conflict'; $reason = 'multiple_episode_candidates' }
            elseif ($candidate.Classification -eq 'conflict') { $classification = 'conflict'; $reason = $candidate.Reason }
            else { $classification = 'unverified'; $reason = 'unique_filename_hint_only'; $suggested = $candidate.RelativePath }
        }
        $rows.Add([pscustomobject]@{
            EpisodeId = $planned.EpisodeId; Title = [string]$planned.Episode.Title
            IdentitySource = $planned.IdentitySource; IdentityFingerprint = $planned.IdentityFingerprint
            ProposedPath = $planned.FileName; Candidates = @($candidates | ForEach-Object { $_.RelativePath })
            SuggestedPath = $suggested; Classification = $classification; Reason = $reason
        })
    }
    return [pscustomobject]@{
        Root = $canonicalRoot; FeedId = [string]$State.feed_id
        Files = @($files.ToArray()); Episodes = @($rows.ToArray())
        NeedsReview = @($files | Where-Object { $_.Classification -ne 'recorded' }).Count -gt 0
    }
}
