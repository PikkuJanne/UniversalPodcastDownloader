#requires -Version 5.1
# Explicit legacy decisions. Preview is read-only; rollback restores metadata only.

function Find-PodcastLegacyReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$ArchiveRoot,
        [AllowNull()][AllowEmptyString()][string]$FeedTitle,
        [AllowEmptyCollection()][object[]]$Episodes,
        [AllowEmptyCollection()][object[]]$SelectedEpisodes,
        [Parameter(Mandatory)]$State,
        [int]$MaxFileNameLength = 180
    )

    $selected = New-PodcastHistoryPlan -Episodes $SelectedEpisodes -State $State -MaxFileNameLength $MaxFileNameLength
    if (@($selected | Where-Object { $null -eq $_.StateRecord }).Count -eq 0) { return $null }
    $folders = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $null = $folders.Add($ArchiveRoot)
    if ($State.generation -eq 0) {
        # Title matching is a reason to request review, never proof of feed identity.
        $oldName = Get-PodcastLegacyFolderName -FeedTitle $FeedTitle
        $oldPath = $null
        try { $oldPath = Get-PodcastDestination -Root $Root -RelativePath $oldName -Directory }
        catch { Write-Verbose 'The historical title cannot name a supported legacy directory; checking the safe modern directory.' }
        if ($oldPath) {
            $null = Assert-PodcastDestination -Root $Root -RelativePath $oldName -Directory
            $null = $folders.Add($oldPath)
        }
    }
    foreach ($folder in $folders) {
        if ($null -eq (Get-PodcastPathAttribute -LiteralPath $folder)) { continue }
        $existing = Read-PodcastHistory -Root $folder
        if ($null -ne $existing -and $existing.feed_id -cne $State.feed_id) { continue }
        $folderFileBudget = [Math]::Min(180, 259 - $folder.Length - 1)
        $plan = New-PodcastLegacyPlan -Root $folder -Episodes $Episodes -State $State -MaxFileNameLength $folderFileBudget
        if (@($plan.Files | Where-Object { $_.Reason -eq 'unsafe_path' }).Count -gt 0) {
            throw 'Legacy inventory found an unsafe path or reparse/junction; preserved the archive for review.'
        }
        $unclaimed = @($plan.Files | Where-Object {
            # A previous interrupted transfer in an established archive can
            # leave this temporary name. Preserve it; a fresh transfer never
            # reuses it or treats it as completed media.
            -not ($State.generation -gt 0 -and $_.RelativePath -cmatch '^\.upd-[a-f0-9]{32}\.tmp$') -and
            -not $_.RecordedEpisodeId -and ($_.Plausible -or $_.CandidateEpisodeIds.Count -gt 0 -or
                $_.RelativePath -match '(?i)\.(mp3|m4a|mp4|aac|wav|flac|ogg|opus|part|partial|tmp)(?:\.|$)')
        })
        if ($unclaimed.Count -gt 0) { return $plan }
    }
    return $null
}

function Get-PodcastLegacyState {
    [CmdletBinding()]
    param([string]$Root, [string]$LegacyRoot, [string]$FeedUrl)

    # Discovery refuses corrupt state, reparse directories and duplicate aliases.
    $archive = Resolve-PodcastArchive -Root $Root -FeedUrl $FeedUrl -FeedTitle 'Legacy'
    $state = Read-PodcastHistory -Root $LegacyRoot
    $alias = Get-PodcastNameHash -IdentityKey ('feed:' + $FeedUrl)
    if ($null -ne $archive.State -and
        -not [StringComparer]::OrdinalIgnoreCase.Equals($archive.FolderName, [IO.Path]::GetFileName($LegacyRoot))) {
        throw 'The feed already belongs to another archive; review the existing history before migration.'
    }
    if ($null -eq $state) { return New-PodcastHistory -FeedId $alias -FeedAliasFingerprint $alias }
    if ($alias -cnotin $state.feed_alias_fingerprints) { throw 'The selected legacy folder belongs to another feed; history was preserved.' }
    return $state
}

function Get-PodcastLegacyChoice {
    [CmdletBinding()]
    param($Plan, $State, [object[]]$Episodes, [int]$FileBudget, [string]$Action,
        [string]$EpisodeId, [string]$FileName, [string]$Sha256, [string]$Checkpoint)

    if ($Action -eq 'Rollback') {
        if ($Checkpoint -cnotmatch '^legacy-[a-f0-9]{32}\.json$') { throw 'Rollback requires an explicit legacy checkpoint basename.' }
        $snapshot = Read-PodcastHistoryFile -Root $Plan.Root -RelativePath ('.upd/' + $Checkpoint)
        if ($null -eq $snapshot -or $snapshot.feed_id -cne $State.feed_id -or $snapshot.generation -gt $State.generation) {
            throw 'The checkpoint does not match this archive or its current generation.'
        }
        if (($snapshot.feed_alias_fingerprints -join ',') -cne ($State.feed_alias_fingerprints -join ',')) {
            throw 'The checkpoint feed aliases differ; review history before rollback.'
        }
        return $snapshot
    }
    if ($EpisodeId -cnotmatch '^[a-f0-9]{64}$') { throw 'Select one explicit LegacyEpisodeId from the preview.' }
    $historyPlan = New-PodcastHistoryPlan -Episodes $Episodes -State $State -MaxFileNameLength $FileBudget
    $planned = @($historyPlan | Where-Object { $_.EpisodeId -ceq $EpisodeId })
    if ($planned.Count -ne 1) { throw 'The selected episode is not uniquely present in the current feed; review a new preview.' }
    if ($Action -eq 'Adopt') {
        if (-not $FileName -or $Sha256 -cnotmatch '^[a-f0-9]{64}$') {
            throw 'Adoption requires an exact LegacyFile and reviewed LegacySha256 from the preview.'
        }
        Assert-PodcastPathComponent -Component $FileName
        $file = @($Plan.Files | Where-Object { [StringComparer]::OrdinalIgnoreCase.Equals($_.RelativePath, $FileName) })
        if ($file.Count -ne 1 -or -not $file[0].Plausible) { throw 'Selected legacy file is not plausible local audio; preserve it for review.' }
        if ($file[0].Sha256 -cne $Sha256) { throw 'The legacy file changed since review: its SHA-256 differs from the reviewed digest.' }
        if ($file[0].RecordedEpisodeId -and $file[0].RecordedEpisodeId -cne $EpisodeId) { throw 'Selected file is already owned by another episode.' }
        if ($null -ne $planned[0].StateRecord -and $planned[0].StateRecord.status -notin @('unverified', 'conflict', 'failed')) {
            throw 'The episode already has recorded evidence. Use explicit Redownload to preserve its file and obtain a separate transfer.'
        }
        # Use the actual on-disk spelling selected by the inventory.
        $planned[0].FileName = $file[0].RelativePath
    }
    return $planned[0]
}

function Save-PodcastLegacyCheckpoint {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Called only after explicit migration ShouldProcess and while holding the history writer lock.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    $name = 'legacy-' + [guid]::NewGuid().ToString('N') + '.json'
    $path = Assert-PodcastDestination -Root $Context.Lock.Root -RelativePath ('.upd/' + $name)
    $bytes = [Text.UTF8Encoding]::new($false, $true).GetBytes(($Context.State | ConvertTo-Json -Depth 8 -Compress))
    if ($bytes.Length -gt 16777216) { throw 'Legacy checkpoint exceeds the supported history size limit.' }
    $stream = [IO.File]::Open($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) }
    finally { $stream.Dispose() }
    $null = Read-PodcastHistoryFile -Root $Context.Lock.Root -RelativePath ('.upd/' + $name)
    return $name
}

function Invoke-PodcastLegacyMigration {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$LegacyRoot,
        [Parameter(Mandatory)][string]$FeedUrl,
        [Parameter(Mandatory)][object[]]$Episodes,
        [ValidateSet('Preview', 'Adopt', 'Redownload', 'Rollback')][string]$Action = 'Preview',
        [string]$EpisodeId, [string]$FileName, [string]$Sha256, [string]$Checkpoint,
        $Policy
    )

    if ($Sha256) { $Sha256 = $Sha256.ToLowerInvariant() }
    $Root = Assert-PodcastDestination -Root $Root -Directory
    if (-not [IO.Path]::IsPathRooted($LegacyRoot)) { $LegacyRoot = [IO.Path]::Combine($Root, $LegacyRoot) }
    $LegacyRoot = Assert-PodcastDestination -Root $LegacyRoot -Directory
    if (-not [StringComparer]::OrdinalIgnoreCase.Equals([IO.Path]::GetDirectoryName($LegacyRoot), $Root)) {
        throw 'LegacyPath must be an existing immediate show directory inside OutputPath.'
    }
    $null = Assert-PodcastDestination -Root $Root -RelativePath ([IO.Path]::GetFileName($LegacyRoot)) -Directory
    if ($null -eq (Get-PodcastPathAttribute -LiteralPath $LegacyRoot)) { throw 'LegacyPath must already exist.' }
    $fileBudget = [Math]::Min(180, 259 - $LegacyRoot.Length - 1)
    $state = Get-PodcastLegacyState -Root $Root -LegacyRoot $LegacyRoot -FeedUrl $FeedUrl
    $plan = New-PodcastLegacyPlan -Root $LegacyRoot -Episodes $Episodes -State $state -MaxFileNameLength $fileBudget
    if ($Action -eq 'Preview') { return $plan }
    $choice = Get-PodcastLegacyChoice -Plan $plan -State $state -Episodes $Episodes -FileBudget $fileBudget `
        -Action $Action -EpisodeId $EpisodeId -FileName $FileName -Sha256 $Sha256 -Checkpoint $Checkpoint
    $restoreRecords = @()
    if ($Action -eq 'Rollback') { $restoreRecords = @($choice.episodes) }
    $plan | Add-Member -NotePropertyName Operation -NotePropertyValue ([pscustomobject]@{
        Action = $Action; EpisodeId = $EpisodeId; File = $FileName; RestoreFrom = $Checkpoint; RestoreRecords = $restoreRecords
    })
    if (-not $PSCmdlet.ShouldProcess('Selected legacy archive', ($Action + ' selected legacy episode/history; preserve all media'))) { return $plan }
    if ($script:PodcastDiagnostics -and $script:PodcastDiagnostics.Preview) { $null = Initialize-PodcastDiagnostics }

    $archiveLock = $null; $historyLock = $null; $fileGuard = $null
    $operationFailed = $false
    try {
        $null = Invoke-PodcastDestinationPreflight -Root $Root
        $archiveLock = Enter-PodcastArchiveLock -Root $Root
        $historyLock = Enter-PodcastHistoryLock -Root $LegacyRoot
        $null = Invoke-PodcastDestinationPreflight -Root $LegacyRoot
        $state = Get-PodcastLegacyState -Root $Root -LegacyRoot $LegacyRoot -FeedUrl $FeedUrl
        $plan = New-PodcastLegacyPlan -Root $LegacyRoot -Episodes $Episodes -State $state -MaxFileNameLength $fileBudget
        $choice = Get-PodcastLegacyChoice -Plan $plan -State $state -Episodes $Episodes -FileBudget $fileBudget `
            -Action $Action -EpisodeId $EpisodeId -FileName $FileName -Sha256 $Sha256 -Checkpoint $Checkpoint
        if ($Action -eq 'Adopt') {
            $path = Assert-PodcastDestination -Root $LegacyRoot -RelativePath $choice.FileName
            $fileGuard = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            $observation = Get-PodcastLegacyFileObservation -Root $LegacyRoot -RelativePath $choice.FileName
            if (-not $observation.Plausible -or $observation.Sha256 -cne $Sha256) { throw 'The legacy file changed since review; no adoption was recorded.' }
        }
        $state = ConvertTo-PodcastHistoryV2 -State $state -Root $LegacyRoot
        if ($state.generation -eq 0) {
            $state.generation = 1
            $state = Write-PodcastHistory -Lock $historyLock -State $state
        }
        $context = @{ Lock = $historyLock; State = $state }
        $savedCheckpoint = Save-PodcastLegacyCheckpoint -Context $context
        Write-Host ('Migration rollback checkpoint: ' + $savedCheckpoint)
        $archiveLock.Dispose(); $archiveLock = $null
        if ($Action -eq 'Rollback') {
            $restored = ConvertTo-PodcastHistoryV2 -State $choice -Root $LegacyRoot
            $restored.generation = $context.State.generation + 1
            $context.State = Write-PodcastHistory -Lock $historyLock -State $restored
            return [pscustomobject]@{ Outcome = 'metadata_restored'; Checkpoint = $savedCheckpoint; RestoredFrom = $Checkpoint }
        }
        if ($Action -eq 'Adopt') {
            $record = New-PodcastEpisodeRecord -Planned $choice
            $record.status = 'adopted'; $record.bytes = $observation.Bytes; $record.local_sha256 = $observation.Sha256
            $record.completed_utc = [datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
            $record.verification = [pscustomobject]@{
                method = 'owner-approved-local-signature'
                media_kind = $observation.MediaKind.ToLowerInvariant().Replace('_', '-')
                notes = @('local-signature-only; transfer-completeness-unverified')
            }
            Save-PodcastEpisodeRecord -Context $context -Record $record
            return [pscustomobject]@{ Outcome = 'adopted'; EpisodeId = $EpisodeId; File = $choice.FileName; Checkpoint = $savedCheckpoint }
        }
        # A new bounded name retains every old file, even if the usual modern
        # destination is already occupied. The transfer itself uses no-overwrite Move.
        $found = $false
        foreach ($attempt in 1..8) {
            $identity = Get-PodcastNameHash -IdentityKey ('redownload:' + $EpisodeId + ':' + [guid]::NewGuid().ToString('N'))
            $choice.FileName = New-EpisodeFileName -Episode $choice.Episode -IdentityHash $identity -MaxLength $fileBudget
            $path = Assert-PodcastDestination -Root $LegacyRoot -RelativePath $choice.FileName
            if ($null -eq (Get-PodcastPathAttribute -LiteralPath $path) -and
                @($state.episodes | Where-Object { $_.relative_path -ieq $choice.FileName }).Count -eq 0) { $found = $true; break }
        }
        if (-not $found) { throw 'Cannot allocate a separate redownload path; original files were preserved.' }
        $choice | Add-Member NoteProperty NewFileIdentityHash $identity
        Write-Verbose ('Migration rollback checkpoint: ' + $savedCheckpoint)
        $transfer = Invoke-PodcastRecordedTransfer -Context $context -Planned $choice -Policy $Policy
        return [pscustomobject]@{ Outcome = 'transfer_verified'; EpisodeId = $EpisodeId; File = $transfer.RelativePath; Checkpoint = $savedCheckpoint }
    }
    catch { $operationFailed = $true; throw }
    finally {
        $historyStream = if ($null -ne $historyLock) { $historyLock.Stream } else { $null }
        $cleanupFailed = $false
        $cleanupCancellation = $null
        foreach ($resource in @($fileGuard, $historyStream, $archiveLock)) {
            if ($null -ne $resource) {
                try { $resource.Dispose() }
                catch {
                    $cleanupFailed = $true
                    if ($null -eq $cleanupCancellation -and (Test-PodcastCancellation -ErrorObject $_)) { $cleanupCancellation = $_ }
                }
            }
        }
        if ($cleanupFailed) {
            if (-not $operationFailed) {
                if ($null -ne $cleanupCancellation) { throw $cleanupCancellation }
                throw 'Legacy action resources could not be closed safely; retained media and history require review.'
            }
            Write-PodcastDiagnostic -Level WARN -Message 'Legacy action cleanup failed; retained media, history and the primary failure are preserved.'
        }
    }
}
