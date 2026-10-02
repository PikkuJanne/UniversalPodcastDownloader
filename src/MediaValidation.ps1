#requires -Version 5.1
# Bounded signature checks are evidence about a completed transfer, not decoding.

function Test-PodcastMpegHeader {
    param([byte[]]$Buffer, [int]$Count, [long]$AvailableBytes)

    if ($Count -lt 4 -or $Buffer[0] -ne 255 -or ($Buffer[1] -band 224) -ne 224) { return $false }
    $version = ($Buffer[1] -shr 3) -band 3
    $layer = ($Buffer[1] -shr 1) -band 3
    $bitrateIndex = $Buffer[2] -shr 4
    $sampleIndex = ($Buffer[2] -shr 2) -band 3
    if ($version -eq 1 -or $layer -eq 0 -or $bitrateIndex -eq 0 -or $bitrateIndex -eq 15 -or $sampleIndex -eq 3) { return $false }
    $sampleRates = @(44100, 48000, 32000)
    $sampleRate = $sampleRates[$sampleIndex]
    if ($version -eq 2) { $sampleRate /= 2 }
    if ($version -eq 0) { $sampleRate /= 4 }
    if ($version -eq 3) {
        $rates = switch ($layer) {
            3 { @(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448) }
            2 { @(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384) }
            1 { @(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320) }
        }
    }
    elseif ($layer -eq 3) { $rates = @(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256) }
    else { $rates = @(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160) }
    $bitrate = $rates[$bitrateIndex] * 1000
    $padding = ($Buffer[2] -shr 1) -band 1
    if ($layer -eq 3) { $frameLength = ([Math]::Floor(12 * $bitrate / $sampleRate) + $padding) * 4 }
    else {
        $coefficient = if ($layer -eq 1 -and $version -ne 3) { 72 } else { 144 }
        $frameLength = [Math]::Floor($coefficient * $bitrate / $sampleRate) + $padding
    }
    # Reject even a plausible four-byte header when its first frame cannot fit.
    return $AvailableBytes -ge $frameLength
}

function Get-PodcastBinarySignature {
    param([byte[]]$Buffer, [int]$Count, [long]$AvailableBytes)

    if (Test-PodcastMpegHeader -Buffer $Buffer -Count $Count -AvailableBytes $AvailableBytes) { return 'mpeg_audio' }
    if ($Count -lt 12) { return $null }
    $first = [Text.Encoding]::ASCII.GetString($Buffer, 0, 4)
    if ($first -eq 'RIFF' -and $AvailableBytes -ge 44 -and
        [Text.Encoding]::ASCII.GetString($Buffer, 8, 4) -eq 'WAVE') {
        $riffSize = [long]$Buffer[4] + [long]$Buffer[5] * 256 + [long]$Buffer[6] * 65536 + [long]$Buffer[7] * 16777216
        # The RIFF size excludes its eight-byte header. A declared container
        # must fit the stored body; bytes after that container remain untouched.
        if ($riffSize -ge 36 -and ($riffSize + 8) -le $AvailableBytes) { return 'wave' }
    }
    if ($first -eq 'fLaC' -and $AvailableBytes -gt 42 -and ($Buffer[4] -band 127) -eq 0 -and
        $Buffer[5] -eq 0 -and $Buffer[6] -eq 0 -and $Buffer[7] -eq 34) { return 'flac' }
    if ($first -eq 'OggS' -and $Count -ge 28 -and $Buffer[4] -eq 0) {
        $segments = [int]$Buffer[26]
        $headerSize = 27 + $segments
        if ($segments -gt 0 -and $Count -ge $headerSize) {
            $payloadSize = 0
            for ($index = 27; $index -lt $headerSize; $index++) { $payloadSize += $Buffer[$index] }
            if ($payloadSize -gt 0 -and $AvailableBytes -ge ($headerSize + $payloadSize)) { return 'ogg_container' }
        }
    }
    if ([Text.Encoding]::ASCII.GetString($Buffer, 4, 4) -eq 'ftyp') {
        $boxSize = [long]$Buffer[0] * 16777216 + [long]$Buffer[1] * 65536 + [long]$Buffer[2] * 256 + $Buffer[3]
        # A bounded ISO base media container check; codec/track parsing is separate.
        if ($boxSize -ge 16 -and $boxSize % 4 -eq 0 -and $AvailableBytes -gt ($boxSize + 8)) { return 'mp4_container' }
    }
    return $null
}

function Test-PodcastTextBody {
    param([byte[]]$Buffer, [int]$Count)

    if ($Count -eq 0) { return $false }
    $encoding = [Text.UTF8Encoding]::new($false, $true)
    $offset = 0
    if ($Count -ge 4 -and $Buffer[0] -eq 255 -and $Buffer[1] -eq 254 -and $Buffer[2] -eq 0 -and $Buffer[3] -eq 0) {
        $encoding = [Text.UTF32Encoding]::new($false, $false, $true); $offset = 4
    }
    elseif ($Count -ge 4 -and $Buffer[0] -eq 0 -and $Buffer[1] -eq 0 -and $Buffer[2] -eq 254 -and $Buffer[3] -eq 255) {
        $encoding = [Text.UTF32Encoding]::new($true, $false, $true); $offset = 4
    }
    elseif ($Count -ge 2 -and $Buffer[0] -eq 255 -and $Buffer[1] -eq 254) {
        $encoding = [Text.UnicodeEncoding]::new($false, $false, $true); $offset = 2
    }
    elseif ($Count -ge 2 -and $Buffer[0] -eq 254 -and $Buffer[1] -eq 255) {
        $encoding = [Text.UnicodeEncoding]::new($true, $false, $true); $offset = 2
    }
    elseif ($Count -ge 3 -and $Buffer[0] -eq 239 -and $Buffer[1] -eq 187 -and $Buffer[2] -eq 191) { $offset = 3 }
    try { $text = $encoding.GetString($Buffer, $offset, $Count - $offset) }
    catch [Text.DecoderFallbackException] { return $false }
    # HTML/XML/JSON and plain-text errors all fail, even with an audio MIME type.
    foreach ($character in $text.ToCharArray()) {
        if ([char]::IsControl($character) -and $character -notin @([char]9, [char]10, [char]13)) { return $false }
    }
    return $true
}

function Test-PodcastMediaFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LiteralPath,
        [bool]$TransferCompleted = $false,
        [Nullable[long]]$HttpContentLength,
        [bool]$ContentLengthAppliesToStoredBytes = $true,
        [string]$ContentType,
        [Nullable[long]]$EnclosureLength
    )

    $result = [pscustomobject]@{
        Valid = $false; Category = 'incomplete_transfer'; Bytes = [long]0
        Verification = 'none'; Warnings = @(); DetectedFormat = $null; InspectedBytes = 0
    }
    if (-not $TransferCompleted) { return $result }
    $stream = $null
    try {
        $stream = [IO.File]::Open($LiteralPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $result.Bytes = $stream.Length
        if ($result.Bytes -eq 0) { $result.Category = 'empty_body'; return $result }
        if ($null -ne $HttpContentLength -and $HttpContentLength -lt 0) { $result.Category = 'invalid_http_length'; return $result }
        $framingKnown = $null -ne $HttpContentLength -and $ContentLengthAppliesToStoredBytes
        if ($framingKnown -and $result.Bytes -ne $HttpContentLength) { $result.Category = 'http_length_mismatch'; return $result }
        if ($null -ne $HttpContentLength -and -not $ContentLengthAppliesToStoredBytes) {
            $result.Warnings += 'HTTP Content-Length describes a different representation and was not compared to stored bytes.'
        }
        if ($null -ne $EnclosureLength -and $EnclosureLength -ge 0 -and $result.Bytes -ne $EnclosureLength) {
            $result.Warnings += 'Publisher enclosure length differs from stored bytes; the estimate is advisory.'
        }

        # Reserve 4KiB of the 64KiB total budget to inspect past a large ID3 tag.
        $sampleLength = [int][Math]::Min($result.Bytes, 61440)
        $buffer = [byte[]]::new($sampleLength)
        $count = 0
        while ($count -lt $sampleLength) {
            $read = $stream.Read($buffer, $count, $sampleLength - $count)
            if ($read -eq 0) { break }
            $count += $read
        }
        $result.InspectedBytes = $count
        $format = Get-PodcastBinarySignature -Buffer $buffer -Count $count -AvailableBytes $result.Bytes
        if ($count -ge 10 -and [Text.Encoding]::ASCII.GetString($buffer, 0, 3) -eq 'ID3' -and
            $buffer[3] -ge 2 -and $buffer[3] -le 4 -and $buffer[4] -lt 255 -and
            $buffer[6] -lt 128 -and $buffer[7] -lt 128 -and $buffer[8] -lt 128 -and $buffer[9] -lt 128) {
            $audioOffset = 10 + [long]$buffer[6] * 2097152 + [long]$buffer[7] * 16384 + [long]$buffer[8] * 128 + $buffer[9]
            if ($buffer[3] -eq 4 -and ($buffer[5] -band 16) -ne 0) { $audioOffset += 10 }
            if ($audioOffset -lt $result.Bytes) {
                $afterTagCount = [int][Math]::Min(4096, $result.Bytes - $audioOffset)
                $afterTag = [byte[]]::new($afterTagCount)
                $null = $stream.Seek($audioOffset, [IO.SeekOrigin]::Begin)
                $read = 0
                while ($read -lt $afterTagCount) {
                    $current = $stream.Read($afterTag, $read, $afterTagCount - $read)
                    if ($current -eq 0) { break }
                    $read += $current
                }
                $result.InspectedBytes += $read
                if (Test-PodcastMpegHeader -Buffer $afterTag -Count $read -AvailableBytes ($result.Bytes - $audioOffset)) { $format = 'mpeg_audio' }
            }
        }
        if (-not $format) {
            $result.Category = if (Test-PodcastTextBody -Buffer $buffer -Count $count) { 'non_audio_text' } else { 'unrecognized_media' }
            return $result
        }
        if ($ContentType -match '^(?i:text/|application/(?:json|xml|xhtml\+xml)(?:;|$))') {
            $result.Warnings += 'Response Content-Type disagrees with the recognized binary signature; MIME is advisory.'
        }
        if ($format -in @('mp4_container', 'ogg_container')) {
            $result.Warnings += 'The container signature was recognized; its audio tracks were not decoded or verified.'
        }
        $result.Valid = $true
        $result.Category = 'accepted'
        $result.DetectedFormat = $format
        $result.Verification = if ($framingKnown) { 'completed_http_length_and_signature' } else { 'completed_eof_and_signature' }
        return $result
    }
    catch {
        $result.Category = 'file_unreadable'
        return $result
    }
    finally { if ($null -ne $stream) { $stream.Dispose() } }
}
