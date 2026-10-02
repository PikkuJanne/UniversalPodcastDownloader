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

function Get-PodcastMediaUInt32 {
    param([byte[]]$Buffer, [int]$Offset, [switch]$LittleEndian)

    if ($LittleEndian) {
        return [long]$Buffer[$Offset] + [long]$Buffer[$Offset + 1] * 256 +
            [long]$Buffer[$Offset + 2] * 65536 + [long]$Buffer[$Offset + 3] * 16777216
    }
    return [long]$Buffer[$Offset] * 16777216 + [long]$Buffer[$Offset + 1] * 65536 +
        [long]$Buffer[$Offset + 2] * 256 + [long]$Buffer[$Offset + 3]
}

function Get-PodcastMp4BoxContent {
    param([byte[]]$Buffer, [int]$Count, [long]$Start, [long]$End, [string]$ContainerPath, [int]$Depth, $Context)

    if ($Depth -gt 4) { return $false }
    $cursor = $Start
    while ($cursor -lt $End) {
        $Context.Nodes++
        if ($Context.Nodes -gt 256) { return $false }
        if ($End - $cursor -lt 8) { $Context.Invalid = $true; return $false }
        if ($cursor -gt $Count - 8) { return $false }
        $position = [int]$cursor
        $boxSize = Get-PodcastMediaUInt32 -Buffer $Buffer -Offset $position
        $type = [Text.Encoding]::ASCII.GetString($Buffer, $position + 4, 4)
        $headerLength = 8
        if ($boxSize -eq 1) {
            if ($End - $cursor -lt 16) { $Context.Invalid = $true; return $false }
            if ($cursor -gt $Count - 16) { return $false }
            $high = Get-PodcastMediaUInt32 -Buffer $Buffer -Offset ($position + 8)
            $low = Get-PodcastMediaUInt32 -Buffer $Buffer -Offset ($position + 12)
            if ($high -gt 2147483647) { $Context.Invalid = $true; return $false }
            $boxSize = $high * 4294967296L + $low
            $headerLength = 16
        }
        elseif ($boxSize -eq 0) { $boxSize = $End - $cursor }
        if ($boxSize -lt $headerLength -or $boxSize -gt $End - $cursor) {
            $Context.Invalid = $true
            return $false
        }
        $boxEnd = $cursor + $boxSize
        $childPath = $null
        if ($ContainerPath -ceq '' -and $type -ceq 'moov') {
            $Context.HasMovie = $true
            $childPath = 'moov'
        }
        elseif ($ContainerPath -ceq 'moov' -and $type -ceq 'trak') { $childPath = 'moov/trak' }
        elseif ($ContainerPath -ceq 'moov/trak' -and $type -ceq 'mdia') { $childPath = 'moov/trak/mdia' }
        if ($null -ne $childPath) {
            $complete = Get-PodcastMp4BoxContent -Buffer $Buffer -Count $Count -Start ($cursor + $headerLength) `
                -End $boxEnd -ContainerPath $childPath -Depth ($Depth + 1) -Context $Context
            if (-not $complete) { return $false }
        }
        elseif ($ContainerPath -ceq 'moov/trak/mdia' -and $type -ceq 'hdlr') {
            # FullBox flags, predefined/component type, handler type, reserved.
            if ($boxSize -lt $headerLength + 24) { $Context.Invalid = $true; return $false }
            if ($cursor + $headerLength + 12 -gt $Count) { return $false }
            $handler = [Text.Encoding]::ASCII.GetString($Buffer, $position + $headerLength + 8, 4)
            if ($handler -ceq 'soun') { $Context.HasAudio = $true }
            if ($handler -ceq 'vide') { $Context.HasVideo = $true }
        }
        # Opaque boxes, including mdat, are skipped by their checked size. Their
        # payload cannot supply handler evidence or consume additional file reads.
        $cursor = $boxEnd
    }
    return $true
}

function Get-PodcastMp4AudioEvidence {
    param([byte[]]$Buffer, [int]$Count, [long]$AvailableBytes)

    $result = [pscustomobject]@{ Valid = $false; Category = 'ambiguous_media'; Warning = $null }
    if ($Count -lt 16 -or [Text.Encoding]::ASCII.GetString($Buffer, 4, 4) -cne 'ftyp') { return $result }
    $fileTypeSize = Get-PodcastMediaUInt32 -Buffer $Buffer -Offset 0
    # Extended/EOF-sized ftyp boxes are outside this modest signature subset.
    if ($fileTypeSize -lt 16 -or $fileTypeSize % 4 -ne 0 -or $fileTypeSize -gt $Count -or
        $AvailableBytes -le $fileTypeSize + 8) { return $result }
    $audioBrand = [Text.Encoding]::ASCII.GetString($Buffer, 8, 4) -ceq 'M4A '
    for ($brandOffset = 16; $brandOffset -lt $fileTypeSize; $brandOffset += 4) {
        if ([Text.Encoding]::ASCII.GetString($Buffer, $brandOffset, 4) -ceq 'M4A ') { $audioBrand = $true }
    }
    $context = [pscustomobject]@{ Nodes = 0; Invalid = $false; HasMovie = $false; HasAudio = $false; HasVideo = $false }
    $complete = Get-PodcastMp4BoxContent -Buffer $Buffer -Count $Count -Start $fileTypeSize `
        -End $AvailableBytes -ContainerPath '' -Depth 0 -Context $context
    if ($context.HasVideo) { $result.Category = 'unsupported_media'; return $result }
    if ($context.Invalid) { return $result }
    if ($audioBrand) {
        $result.Valid = $true
        $result.Warning = 'The M4A brand suggests audio; audio-only tracks and codecs were not fully verified.'
    }
    elseif ($complete -and $context.HasMovie -and $context.HasAudio) {
        $result.Valid = $true
        $result.Warning = 'Audio handler evidence was found within the bounded MP4 inspection; codecs were not decoded.'
    }
    if ($result.Valid) { $result.Category = 'accepted' }
    return $result
}

function Get-PodcastOggAudioEvidence {
    param([byte[]]$Buffer, [int]$Count, [long]$AvailableBytes)

    $result = [pscustomobject]@{ Valid = $false; Category = 'ambiguous_media' }
    if ($Count -lt 27 -or $Buffer[4] -ne 0 -or ($Buffer[5] -band 2) -eq 0 -or
        ($Buffer[5] -band 1) -ne 0 -or (Get-PodcastMediaUInt32 -Buffer $Buffer -Offset 18 -LittleEndian) -ne 0) { return $result }
    $segments = [int]$Buffer[26]
    $headerLength = 27 + $segments
    if ($segments -eq 0 -or $Count -lt $headerLength) { return $result }
    $pagePayload = 0
    $packetLength = 0
    $packetComplete = $false
    for ($segment = 27; $segment -lt $headerLength; $segment++) {
        $pagePayload += $Buffer[$segment]
        if (-not $packetComplete) {
            $packetLength += $Buffer[$segment]
            if ($Buffer[$segment] -lt 255) { $packetComplete = $true }
        }
    }
    if (-not $packetComplete -or $packetLength -eq 0 -or $AvailableBytes -lt $headerLength + $pagePayload -or
        $Count -lt $headerLength + $packetLength) { return $result }
    $packet = $headerLength
    if ($packetLength -ge 7 -and $Buffer[$packet] -eq 128 -and
        [Text.Encoding]::ASCII.GetString($Buffer, $packet + 1, 6) -ceq 'theora') {
        $result.Category = 'unsupported_media'
        return $result
    }
    if ($packetLength -ge 19 -and [Text.Encoding]::ASCII.GetString($Buffer, $packet, 8) -ceq 'OpusHead' -and
        $Buffer[$packet + 8] -ge 1 -and $Buffer[$packet + 8] -le 15 -and $Buffer[$packet + 9] -gt 0) { $result.Valid = $true }
    elseif ($packetLength -ge 30 -and $Buffer[$packet] -eq 1 -and
        [Text.Encoding]::ASCII.GetString($Buffer, $packet + 1, 6) -ceq 'vorbis' -and
        (Get-PodcastMediaUInt32 -Buffer $Buffer -Offset ($packet + 7) -LittleEndian) -eq 0 -and
        $Buffer[$packet + 11] -gt 0 -and
        (Get-PodcastMediaUInt32 -Buffer $Buffer -Offset ($packet + 12) -LittleEndian) -gt 0 -and
        ($Buffer[$packet + 29] -band 1) -ne 0) { $result.Valid = $true }
    elseif ($packetLength -ge 80 -and [Text.Encoding]::ASCII.GetString($Buffer, $packet, 8) -ceq 'Speex   ') {
        $speexHeaderLength = Get-PodcastMediaUInt32 -Buffer $Buffer -Offset ($packet + 32) -LittleEndian
        if ($speexHeaderLength -ge 80 -and $speexHeaderLength -le $packetLength -and
            (Get-PodcastMediaUInt32 -Buffer $Buffer -Offset ($packet + 36) -LittleEndian) -gt 0 -and
            (Get-PodcastMediaUInt32 -Buffer $Buffer -Offset ($packet + 48) -LittleEndian) -gt 0) { $result.Valid = $true }
    }
    if ($result.Valid) { $result.Category = 'accepted' }
    return $result
}

function Get-PodcastMediaTypeExtension {
    param([AllowNull()][AllowEmptyString()][string]$ContentType)

    if ([string]::IsNullOrWhiteSpace($ContentType)) { return 'generic' }
    if ($ContentType.Length -gt 1024) { return 'unsupported' }
    $mediaType = $ContentType.Split(';')[0].Trim().ToLowerInvariant()
    switch ($mediaType) {
        { $_ -in @('application/octet-stream', 'binary/octet-stream') } { return 'generic' }
        { $_ -in @('audio/mpeg', 'audio/mp3', 'audio/x-mp3', 'audio/x-mpeg', 'audio/x-mpegmp3') } { return 'mp3' }
        { $_ -in @('audio/mp4', 'audio/m4a', 'audio/x-m4a', 'application/mp4') } { return 'm4a' }
        { $_ -in @('audio/ogg', 'audio/vorbis', 'audio/opus', 'audio/speex', 'application/ogg') } { return 'ogg' }
        { $_ -in @('audio/wav', 'audio/wave', 'audio/x-wav', 'audio/vnd.wave') } { return 'wav' }
        { $_ -in @('audio/flac', 'audio/x-flac') } { return 'flac' }
        default { return 'unsupported' }
    }
}

function Get-PodcastMediaUrlExtension {
    param([AllowNull()][AllowEmptyString()][string]$Url)

    $target = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$target)) { return $null }
    # URL hints are advisory even when a caller supplies malformed path text.
    try { $extension = [IO.Path]::GetExtension($target.AbsolutePath).TrimStart('.').ToLowerInvariant() }
    catch [ArgumentException] { return $null }
    switch ($extension) {
        { $_ -in @('mp3', 'm4a', 'wav', 'flac') } { return $extension }
        { $_ -in @('ogg', 'oga', 'opus') } { return 'ogg' }
        'mp4' { return 'm4a' }
        { $_ -in @('m4v', 'webm', 'mkv', 'avi', 'mov', 'jpg', 'png', 'html', 'json') } { return 'unsupported' }
        default { return $null }
    }
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
        [string]$EnclosureContentType,
        [string]$MediaUrl,
        [Nullable[long]]$EnclosureLength
    )

    $result = [pscustomobject]@{
        Valid = $false; Category = 'incomplete_transfer'; Bytes = [long]0
        Verification = 'none'; Warnings = @(); DetectedFormat = $null; Extension = $null; InspectedBytes = 0
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
        $mpegLayer = if ($format -eq 'mpeg_audio') { ($buffer[1] -shr 1) -band 3 } else { 0 }
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
                if (Test-PodcastMpegHeader -Buffer $afterTag -Count $read -AvailableBytes ($result.Bytes - $audioOffset)) {
                    $format = 'mpeg_audio'
                    $mpegLayer = ($afterTag[1] -shr 1) -band 3
                }
            }
        }
        # A recognizable container must supply structured audio evidence; MIME
        # and URL hints cannot turn an ambiguous container into validated audio.
        if ($count -ge 8 -and [Text.Encoding]::ASCII.GetString($buffer, 4, 4) -ceq 'ftyp') {
            $container = Get-PodcastMp4AudioEvidence -Buffer $buffer -Count $count -AvailableBytes $result.Bytes
            if (-not $container.Valid) { $result.Category = $container.Category; return $result }
            $format = 'mp4_container'
            $result.Warnings += $container.Warning
        }
        elseif ($count -ge 4 -and [Text.Encoding]::ASCII.GetString($buffer, 0, 4) -ceq 'OggS') {
            $container = Get-PodcastOggAudioEvidence -Buffer $buffer -Count $count -AvailableBytes $result.Bytes
            if (-not $container.Valid) { $result.Category = $container.Category; return $result }
            $format = 'ogg_container'
        }
        if ($format -eq 'mpeg_audio' -and $mpegLayer -ne 1) { $result.Category = 'unsupported_media'; return $result }
        if (-not $format) {
            $result.Category = if (Test-PodcastTextBody -Buffer $buffer -Count $count) { 'non_audio_text' } else { 'unrecognized_media' }
            return $result
        }
        $extension = switch ($format) {
            'mpeg_audio' { 'mp3' }; 'mp4_container' { 'm4a' }; 'ogg_container' { 'ogg' }; 'wave' { 'wav' }; 'flac' { 'flac' }
        }
        $responseHint = Get-PodcastMediaTypeExtension -ContentType $ContentType
        if ($responseHint -ne 'generic' -and $responseHint -ne $extension) {
            $result.Warnings += 'Response Content-Type disagrees with the recognized binary signature; MIME is advisory.'
        }
        $publisherHint = Get-PodcastMediaTypeExtension -ContentType $EnclosureContentType
        if ($publisherHint -ne 'generic' -and $publisherHint -ne $extension) {
            $result.Warnings += 'Publisher enclosure type disagrees with the detected audio format; byte evidence was used.'
        }
        $urlHint = Get-PodcastMediaUrlExtension -Url $MediaUrl
        if ($urlHint -and $urlHint -ne $extension) {
            $result.Warnings += 'Media URL extension disagrees with the detected audio format; byte evidence was used.'
        }
        if ($format -in @('mp4_container', 'ogg_container')) {
            $result.Warnings += 'The container signature was recognized; its audio tracks were not decoded or verified.'
        }
        $result.Valid = $true
        $result.Category = 'accepted'
        $result.DetectedFormat = $format
        $result.Extension = $extension
        $result.Verification = if ($framingKnown) { 'completed_http_length_and_signature' } else { 'completed_eof_and_signature' }
        return $result
    }
    catch {
        $result.Category = 'file_unreadable'
        return $result
    }
    finally { if ($null -ne $stream) { $stream.Dispose() } }
}
