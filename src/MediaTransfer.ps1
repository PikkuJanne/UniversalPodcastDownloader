# Resume only a checkpoint owned by this episode, under the archive writer lock.
function Invoke-PodcastMediaTransfer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$RelativePath,
        [Nullable[long]]$EnclosureLength,
        [string]$EnclosureContentType,
        [scriptblock]$ResolveFinalPath,
        [scriptblock]$BeforeFinalize,
        $Policy,
        $ResumeContext
    )

    $destination = Assert-PodcastDestination -Root $Root -RelativePath $RelativePath
    if (Test-Path -LiteralPath $destination) { throw 'Media destination already exists; preserving it.' }
    $relativeDirectory = [IO.Path]::GetDirectoryName($RelativePath)
    $temporaryRelative = $null
    $temporary = $null
    $destinationStream = $null
    $owned = $false
    $resume = $null
    $session = [pscustomobject]@{ State = $null; Lock = $null; Stream = $null; PartialName = ''; Preserve = $false }
    $requestFingerprint = ''
    if ($null -ne $ResumeContext) {
        Assert-PodcastResumeLock -Lock $ResumeContext.Lock
        if (-not [string]::Equals([IO.Path]::GetFullPath($Root), [IO.Path]::GetFullPath($ResumeContext.Lock.Root), [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Resume context does not belong to this archive.'
        }
        $session.Lock = $ResumeContext.Lock
        $requestFingerprint = Get-PodcastResumeUriFingerprint -Uri (Get-PodcastRequestUri -Uri $Uri)
        $session.State = Read-PodcastResumeState -Root $Root -EpisodeId $ResumeContext.EpisodeId
        if ($null -ne $session.State) {
            $destinationStream = Open-PodcastResumePartial -Lock $session.Lock -State $session.State -FeedId $ResumeContext.FeedId `
                -EpisodeId $ResumeContext.EpisodeId -RelativePath $RelativePath -RequestFingerprint $requestFingerprint
            if ($null -ne $destinationStream) {
                $temporaryRelative = $session.State.partial_name
                $resume = [pscustomobject]@{ Offset = $session.State.offset; TotalLength = $session.State.total_length;
                    ETag = $session.State.etag; ContentType = $session.State.content_type; FinalUriFingerprint = $session.State.final_uri_fingerprint }
                $session.Preserve = $true
            }
            else { Write-PodcastDiagnostic -Level WARN -Message 'Resume checkpoint requires a fresh transfer; previous partial preserved.' }
        }
    }
    try {
        # At most one safe fresh GET follows a rejected range in this attempt.
        # The existing outer transport policy alone owns retries/backoff.
        for ($requestNumber = 0; $requestNumber -lt 2; $requestNumber++) {
            if ($null -eq $destinationStream) {
                $temporaryRelative = [IO.Path]::Combine($relativeDirectory, ('.upd-' + [guid]::NewGuid().ToString('N') + '.tmp'))
                $temporary = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative
                $destinationStream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
                $owned = $true
                $session.Preserve = $false
            }
            else { $temporary = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative }
            $session.Stream = $destinationStream
            $session.PartialName = $temporaryRelative
            $requestArguments = @{ Uri = $Uri; DestinationStream = $destinationStream; Policy = $Policy }
            if ($null -ne $ResumeContext) {
                $requestArguments.Resume = $resume
                $requestArguments.OnResponse = {
                    param($Info)
                    if (-not $Info.ResumeSupported) { return }
                    # Preserve the reserved file even if checkpoint replacement
                    # fails: its durable ownership outcome may be uncertain.
                    $session.Preserve = $true
                    $next = [pscustomobject]@{
                        schema_version = 1; feed_id = $ResumeContext.FeedId; episode_id = $ResumeContext.EpisodeId
                        relative_path = $RelativePath; partial_name = $session.PartialName
                        request_fingerprint = $requestFingerprint; final_uri_fingerprint = $Info.FinalUriFingerprint
                        etag = $Info.ETag; total_length = [long]$Info.TotalLength; content_type = $Info.ContentType
                        content_encoding = 'identity'; offset = [long]$session.Stream.Length
                        prefix_sha256 = Get-PodcastResumeStreamHash -Stream $session.Stream
                    }
                    $session.State = Write-PodcastResumeState -Lock $session.Lock -State $next -Expected $session.State
                }
                $requestArguments.OnProgress = {
                    param([long]$Bytes)
                    if ($Bytes -ne $session.Stream.Length) { throw 'Resume progress does not match the owned stream.' }
                    Update-PodcastResumeCheckpoint -Session $session
                }
            }
            try {
                $null = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative
                $transfer = Invoke-PodcastMediaRequest @requestArguments
                $destinationStream.Flush($true)
                if ($transfer.Completed) { Update-PodcastResumeCheckpoint -Session $session -Force }
            }
            catch {
                # A caught network failure can checkpoint its written prefix.
                # Local errors remain non-retryable and preserve existing evidence.
                if ($null -ne (Get-PodcastTransportFailure -ErrorObject $_)) {
                    Update-PodcastResumeCheckpoint -Session $session -Force
                }
                throw
            }
            finally { $destinationStream.Dispose(); $destinationStream = $null }
            if ($transfer.Completed) { break }
            if ($null -eq $resume -or -not $transfer.RestartRequired) { throw 'Media transfer did not complete.' }
            Write-PodcastDiagnostic -Level WARN -Message 'Resume response requires a fresh transfer; previous partial preserved.'
            $resume = $null
            $owned = $false
        }

        $null = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative
        $validation = Test-PodcastMediaFile -LiteralPath $temporary -TransferCompleted:$transfer.Completed `
            -HttpContentLength $transfer.ContentLength -ContentType $transfer.ContentType -EnclosureLength $EnclosureLength `
            -EnclosureContentType $EnclosureContentType -MediaUrl $Uri
        if (-not $validation.Valid) {
            throw [IO.InvalidDataException]::new(('Media validation failed: {0}.' -f $validation.Category))
        }
        if ($validation.Bytes -ne $transfer.Bytes) {
            throw [IO.InvalidDataException]::new('The temporary file length changed after transfer.')
        }
        $evidence = Get-PodcastFileEvidence -Root $Root -RelativePath $temporaryRelative
        if ($evidence.Bytes -ne $validation.Bytes) { throw 'Media changed before recording transfer evidence.' }
        # A checkpoint continues to own the provisional path. Only a completed,
        # validated transfer may resolve the final path, before prepared history.
        $finalRelativePath = $RelativePath
        if ($ResolveFinalPath) { $finalRelativePath = & $ResolveFinalPath $validation }
        if ([string]::IsNullOrWhiteSpace($finalRelativePath) -or
            [IO.Path]::GetDirectoryName($finalRelativePath) -cne $relativeDirectory) {
            throw 'Resolved media destination must remain in the planned directory.'
        }
        $destination = Assert-PodcastDestination -Root $Root -RelativePath $finalRelativePath
        if (Test-Path -LiteralPath $destination) { throw 'Media destination already exists; preserving it.' }
        if ([IO.Path]::GetExtension($finalRelativePath) -ine ('.' + $validation.Extension)) {
            $validation.Warnings += 'Recorded destination extension retained; it differs from the recognized audio format.'
        }
        $result = [pscustomobject]@{
            Outcome = 'downloaded'
            File = $destination
            RelativePath = $finalRelativePath
            Bytes = $validation.Bytes
            Sha256 = $evidence.Sha256
            Verification = $validation.Verification
            DetectedFormat = $validation.DetectedFormat
            Warnings = @($validation.Warnings)
        }
        # A durable prepared record bridges final placement and history commit.
        if ($BeforeFinalize) { $null = & $BeforeFinalize $result }
        if ($null -ne $session.State -and $session.State.partial_name -ceq $temporaryRelative) {
            # Retire ownership before rename, so a crash cannot leave an active
            # checkpoint referring to a partial already placed as final media.
            Remove-PodcastResumeState -Lock $session.Lock -State $session.State
            # Once its sidecar is retired, a new file is again this attempt's
            # cleanup responsibility. Previously resumed partials stay preserved.
            if ($owned) { $session.Preserve = $false }
        }
        $null = Assert-PodcastDestination -Root $Root -RelativePath $finalRelativePath
        # Close and validate before atomic no-overwrite placement on the same volume.
        [IO.File]::Move($temporary, $destination)
        $owned = $false
        return $result
    }
    finally {
        if ($null -ne $destinationStream) { $destinationStream.Dispose() }
        if ($owned -and -not $session.Preserve) {
            # Clean only an uncheckpointed file reserved by this attempt.
            # Previously owned or unclaimed partials always remain untouched.
            $null = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative
            [IO.File]::Delete($temporary)
        }
    }
}
