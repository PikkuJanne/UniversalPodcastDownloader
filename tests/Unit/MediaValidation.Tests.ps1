BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'src/MediaSelection.ps1')
    $helper = Join-Path $repositoryRoot 'src/MediaValidation.ps1'
    if (Test-Path -LiteralPath $helper) { . $helper }
    Mock Invoke-WebRequest { throw 'Validation must not make network requests.' }
    Mock Read-Host { throw 'Validation must not prompt.' }
    $script:AudioBytes = [IO.File]::ReadAllBytes((Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
}

Describe 'A012 bounded conservative media validation' -Tag 'Unit', 'A012' {
    BeforeEach { $script:MediaPath = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.tmp') }

    It 'rejects an empty successful response' {
        [IO.File]::WriteAllBytes($script:MediaPath, [byte[]]@())
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -HttpContentLength 0
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'empty_body'
        $result.Bytes | Should -Be 0
        $result.Verification | Should -Be 'none'
    }

    It 'requires explicit completed-transfer evidence even for valid audio bytes' {
        [IO.File]::WriteAllBytes($script:MediaPath, $script:AudioBytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'incomplete_transfer'
    }

    It 'accepts known audio with matching HTTP framing and generic MIME' {
        [IO.File]::WriteAllBytes($script:MediaPath, $script:AudioBytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -HttpContentLength $script:AudioBytes.Length -ContentType 'application/octet-stream'
        $result.Valid | Should -BeTrue
        $result.Category | Should -Be 'accepted'
        $result.Bytes | Should -Be $script:AudioBytes.Length
        $result.Verification | Should -Be 'completed_http_length_and_signature'
        $result.DetectedFormat | Should -Be 'mpeg_audio'
        $result.Warnings -is [array] | Should -BeTrue
        $result.Warnings.Count | Should -Be 0
    }

    It 'accepts a completed EOF-terminated audio response without a length' {
        [IO.File]::WriteAllBytes($script:MediaPath, $script:AudioBytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true
        $result.Valid | Should -BeTrue
        $result.Verification | Should -Be 'completed_eof_and_signature'
    }

    It 'rejects either direction of HTTP Content-Length mismatch <Delta>' -ForEach @(
        @{ Delta = 1 }, @{ Delta = -1 }
    ) {
        [IO.File]::WriteAllBytes($script:MediaPath, $script:AudioBytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -HttpContentLength ($script:AudioBytes.Length + $Delta)
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'http_length_mismatch'
    }

    It 'treats an enclosure length mismatch as advisory' {
        [IO.File]::WriteAllBytes($script:MediaPath, $script:AudioBytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -HttpContentLength $script:AudioBytes.Length -EnclosureLength ($script:AudioBytes.Length + 100)
        $result.Valid | Should -BeTrue
        $result.Warnings.Count | Should -Be 1
        $result.Warnings[0] | Should -Match 'enclosure.*advisory'
    }

    It 'does not compare an encoded HTTP length to decoded stored bytes' {
        [IO.File]::WriteAllBytes($script:MediaPath, $script:AudioBytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -HttpContentLength 1 -ContentLengthAppliesToStoredBytes $false
        $result.Valid | Should -BeTrue
        $result.Verification | Should -Be 'completed_eof_and_signature'
        ($result.Warnings -join ' ') | Should -Match 'representation'
    }

    It 'rejects negative HTTP length metadata' {
        [IO.File]::WriteAllBytes($script:MediaPath, $script:AudioBytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -HttpContentLength -1
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'invalid_http_length'
    }

    It 'rejects <Label> despite an audio MIME type' -ForEach @(
        @{ Label = 'HTML'; Text = '  <!DOCTYPE html><html><body>Access denied</body></html>' },
        @{ Label = 'XML'; Text = '<?xml version="1.0"?><Error>Access denied</Error>' },
        @{ Label = 'JSON object'; Text = '{"error":"Access denied"}' },
        @{ Label = 'JSON array'; Text = '[{"error":"Access denied"}]' },
        @{ Label = 'plain text'; Text = 'Access denied. Try signing in again.' },
        @{ Label = 'whitespace'; Text = " `r`n`t " }
    ) {
        [IO.File]::WriteAllText($script:MediaPath, $Text, [Text.UTF8Encoding]::new($false))
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -ContentType 'audio/mpeg'
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'non_audio_text'
        $result.Verification | Should -Be 'none'
    }

    It 'rejects text with a <EncodingName> byte-order mark' -ForEach @(
        @{ EncodingName = 'utf8'; Encoding = [Text.UTF8Encoding]::new($true) },
        @{ EncodingName = 'utf16le'; Encoding = [Text.Encoding]::Unicode },
        @{ EncodingName = 'utf16be'; Encoding = [Text.Encoding]::BigEndianUnicode },
        @{ EncodingName = 'utf32le'; Encoding = [Text.Encoding]::UTF32 }
    ) {
        [IO.File]::WriteAllText($script:MediaPath, '<html>Error</html>', $Encoding)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -ContentType 'audio/mpeg'
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'non_audio_text'
    }

    It 'rejects unknown binary rather than claiming audio validation from MIME' {
        [IO.File]::WriteAllBytes($script:MediaPath, [byte[]](0, 255, 128, 0, 42, 0, 255, 1))
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -ContentType 'audio/mpeg'
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'unrecognized_media'
    }

    It 'accepts the bounded <Format> signature without relying on MIME' -ForEach @(
        @{ Format = 'wave'; Bytes = [byte[]](82,73,70,70,40,0,0,0,87,65,86,69,102,109,116,32,16,0,0,0,1,0,1,0,68,172,0,0,136,88,1,0,2,0,16,0,100,97,116,97,4,0,0,0,0,0,0,0) },
        @{ Format = 'flac'; Bytes = [byte[]](102,76,97,67,128,0,0,34) + [byte[]]::new(34) + [byte[]](255,248,0,0) },
        @{ Format = 'ogg_container'; Bytes = [byte[]](79,103,103,83,0,2) + [byte[]]::new(20) + [byte[]](1,19,79,112,117,115,72,101,97,100,1,1) + [byte[]]::new(9) },
        @{ Format = 'mp4_container'; Bytes = [byte[]](0,0,0,20,102,116,121,112,77,52,65,32,0,0,0,0,77,52,65,32,0,0,0,12,109,100,97,116,0,0,0,0) }
    ) {
        [IO.File]::WriteAllBytes($script:MediaPath, $Bytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -ContentType 'application/octet-stream'
        $result.Valid | Should -BeTrue
        $result.DetectedFormat | Should -Be $Format
        $result.Verification | Should -Not -Match 'fully|decoded|complete_audio'
    }

    It 'does not accept an ID3 tag header as a complete audio file' {
        [IO.File]::WriteAllBytes($script:MediaPath, [byte[]](73,68,51,4,0,0,0,0,0,0))
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'unrecognized_media'
    }

    It 'rejects an impossible RIFF declared length <DeclaredSize> even with matching HTTP framing' -ForEach @(
        @{ DeclaredSize = 4096 }, @{ DeclaredSize = 35 }, @{ DeclaredSize = 4294967295L }
    ) {
        $wave = [byte[]](82,73,70,70,40,0,0,0,87,65,86,69,102,109,116,32,16,0,0,0,1,0,1,0,68,172,0,0,136,88,1,0,2,0,16,0,100,97,116,97,4,0,0,0,0,0,0,0)
        [BitConverter]::GetBytes([uint32]$DeclaredSize).CopyTo($wave, 4)
        [IO.File]::WriteAllBytes($script:MediaPath, $wave)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -HttpContentLength $wave.Length -ContentType 'audio/wav'
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'unrecognized_media'
        $result.Verification | Should -Be 'none'
    }

    It 'accepts a sane ID3 header followed by MPEG audio and rejects an impossible tag size' {
        $header = [byte[]](73,68,51,4,0,0,0,0,0,0)
        # The checked-in silence fixture has a 44-byte leading ID3 tag.
        [IO.File]::WriteAllBytes($script:MediaPath, ($header + $script:AudioBytes[44..($script:AudioBytes.Length - 1)]))
        (Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true).Valid | Should -BeTrue
        $header[9] = 127
        [IO.File]::WriteAllBytes($script:MediaPath, ($header + [byte[]](255,251,144,0)))
        (Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true).Valid | Should -BeFalse
    }

    It 'rejects obviously truncated <Format> signatures' -ForEach @(
        @{ Format = 'mpeg'; Bytes = [byte[]](255,251,144,0) },
        @{ Format = 'flac'; Bytes = [byte[]](102,76,97,67) },
        @{ Format = 'ogg'; Bytes = [byte[]](79,103,103,83,0) + [byte[]]::new(22) },
        @{ Format = 'wave'; Bytes = [byte[]](82,73,70,70,4,0,0,0,87,65,86,69) },
        @{ Format = 'mp4'; Bytes = [byte[]](0,0,0,20,102,116,121,112,77,52,65,32,0,0,0,0,77,52,65,32) }
    ) {
        [IO.File]::WriteAllBytes($script:MediaPath, $Bytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -HttpContentLength $Bytes.Length
        $result.Valid | Should -BeFalse
        $result.Verification | Should -Be 'none'
    }

    It 'handles large ID3 artwork within the total inspection budget' {
        $header = [byte[]](73,68,51,4,0,0,0,4,0,0)
        $rawAudio = $script:AudioBytes[44..($script:AudioBytes.Length - 1)]
        [IO.File]::WriteAllBytes($script:MediaPath, ($header + [byte[]]::new(65536) + $rawAudio))
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true
        $result.Valid | Should -BeTrue
        $result.DetectedFormat | Should -Be 'mpeg_audio'
        $result.InspectedBytes | Should -BeGreaterThan 61440
        $result.InspectedBytes | Should -BeLessOrEqual 65536
    }

    It 'treats a text MIME type as advisory when a real binary audio signature matches' {
        [IO.File]::WriteAllBytes($script:MediaPath, $script:AudioBytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true -ContentType 'text/plain; charset=utf-8'
        $result.Valid | Should -BeTrue
        ($result.Warnings -join ' ') | Should -Match 'MIME is advisory'
    }

    It 'caps inspection at 64KiB and closes its read handle without modifying bytes' {
        $large = [byte[]]::new(1024 * 1024)
        [Array]::Copy($script:AudioBytes, $large, $script:AudioBytes.Length)
        [IO.File]::WriteAllBytes($script:MediaPath, $large)
        $hash = (Get-FileHash -LiteralPath $script:MediaPath -Algorithm SHA256).Hash
        $stamp = (Get-Item -LiteralPath $script:MediaPath).LastWriteTimeUtc
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true
        $result.Valid | Should -BeTrue
        $result.Bytes | Should -Be $large.Length
        $result.InspectedBytes | Should -BeLessOrEqual 65536
        $result.InspectedBytes | Should -BeGreaterThan 0
        (Get-FileHash -LiteralPath $script:MediaPath -Algorithm SHA256).Hash | Should -Be $hash
        (Get-Item -LiteralPath $script:MediaPath).LastWriteTimeUtc | Should -Be $stamp
        $exclusive = [IO.File]::Open($script:MediaPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $exclusive.Dispose()
    }

    It 'closes the file after a validation failure too' {
        [IO.File]::WriteAllText($script:MediaPath, '<html>Not media</html>')
        (Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true).Valid | Should -BeFalse
        $exclusive = [IO.File]::Open($script:MediaPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $exclusive.Dispose()
    }

    It 'returns a safe failure for an unreadable or missing file' {
        $result = Test-PodcastMediaFile -LiteralPath $script:MediaPath -TransferCompleted $true
        $result.Valid | Should -BeFalse
        $result.Category | Should -Be 'file_unreadable'
        ($result | ConvertTo-Json -Depth 3) | Should -Not -Match ([regex]::Escape($script:MediaPath))
    }
}
