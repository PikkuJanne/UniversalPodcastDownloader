# One attempt owns one newly created sibling; previous partial files are never reused.
function Invoke-PodcastMediaTransfer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$RelativePath,
        [Nullable[long]]$EnclosureLength,
        [scriptblock]$BeforeFinalize
    )

    $destination = Assert-PodcastDestination -Root $Root -RelativePath $RelativePath
    if (Test-Path -LiteralPath $destination) { throw 'Media destination already exists; preserving it.' }
    $relativeDirectory = [IO.Path]::GetDirectoryName($RelativePath)
    $temporaryRelative = [IO.Path]::Combine($relativeDirectory, ('.upd-' + [guid]::NewGuid().ToString('N') + '.tmp'))
    $temporary = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative
    $owned = $false
    try {
        # Keep the exclusively reserved handle open throughout the transfer.
        $destinationStream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $owned = $true
        try {
            $null = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative
            $transfer = Invoke-PodcastMediaRequest -Uri $Uri -DestinationStream $destinationStream
            $destinationStream.Flush($true)
        }
        finally { $destinationStream.Dispose() }

        $null = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative
        $validation = Test-PodcastMediaFile -LiteralPath $temporary -TransferCompleted:$transfer.Completed `
            -HttpContentLength $transfer.ContentLength -ContentType $transfer.ContentType -EnclosureLength $EnclosureLength
        if (-not $validation.Valid) {
            throw [IO.InvalidDataException]::new(('Media validation failed: {0}.' -f $validation.Category))
        }
        if ($validation.Bytes -ne $transfer.Bytes) {
            throw [IO.InvalidDataException]::new('The temporary file length changed after transfer.')
        }
        $evidence = Get-PodcastFileEvidence -Root $Root -RelativePath $temporaryRelative
        if ($evidence.Bytes -ne $validation.Bytes) { throw 'Media changed before recording transfer evidence.' }
        $result = [pscustomobject]@{
            Outcome = 'downloaded'
            File = $destination
            Bytes = $validation.Bytes
            Sha256 = $evidence.Sha256
            Verification = $validation.Verification
            DetectedFormat = $validation.DetectedFormat
            Warnings = @($validation.Warnings)
        }
        # A durable prepared record bridges final placement and history commit.
        if ($BeforeFinalize) { $null = & $BeforeFinalize $result }
        $null = Assert-PodcastDestination -Root $Root -RelativePath $RelativePath
        # Close and validate before atomic no-overwrite placement on the same volume.
        [IO.File]::Move($temporary, $destination)
        $owned = $false
        return $result
    }
    finally {
        if ($owned) {
            # Clean only the file reserved by this attempt. A crash can leave it;
            # later runs start with a new name and preserve all unclaimed files.
            $null = Assert-PodcastDestination -Root $Root -RelativePath $temporaryRelative
            [IO.File]::Delete($temporary)
        }
    }
}
