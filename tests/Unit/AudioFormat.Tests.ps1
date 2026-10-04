BeforeAll {
    $script:AudioRepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Mock Read-Host { throw 'Audio units must not prompt.' }
    Mock Invoke-WebRequest { throw 'Audio units must not contact external network.' }
    . (Join-Path $script:AudioRepositoryRoot 'UniversalPodcastDownloader.ps1') -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Audio units must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Audio units must not fetch metadata.' }
    Mock Invoke-PodcastMediaRequest { throw 'Audio units must not request media.' }
    $script:FormatMp3 = [IO.File]::ReadAllBytes((Join-Path $script:AudioRepositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
    $script:FormatWave = [byte[]](82,73,70,70,40,0,0,0,87,65,86,69,102,109,116,32,16,0,0,0,1,0,1,0,68,172,0,0,136,88,1,0,2,0,16,0,100,97,116,97,4,0,0,0,0,0,0,0)
    $script:FormatFlac = [byte[]](102,76,97,67,128,0,0,34) + [byte[]]::new(34) + [byte[]](255,248,0,0)

    function Get-AudioTestBox {
        param([string]$Type, [byte[]]$Payload)
        $size = $Payload.Length + 8
        $header = [byte[]]@([byte](($size -shr 24) -band 255), [byte](($size -shr 16) -band 255), [byte](($size -shr 8) -band 255), [byte]($size -band 255))
        return ,([byte[]]($header + [Text.Encoding]::ASCII.GetBytes($Type) + $Payload))
    }

    function Get-AudioTestMp4 {
        param([string]$Brand = 'M4A ', [string[]]$Handlers = @(), [byte[]]$Payload = ([byte[]]::new(32)))
        $ftyp = Get-AudioTestBox -Type 'ftyp' -Payload ([Text.Encoding]::ASCII.GetBytes($Brand) + [byte[]]::new(4) + [Text.Encoding]::ASCII.GetBytes($Brand))
        $tracks = [byte[]]@()
        foreach ($handler in $Handlers) {
            $hdlr = Get-AudioTestBox -Type 'hdlr' -Payload ([byte[]]::new(8) + [Text.Encoding]::ASCII.GetBytes($handler) + [byte[]]::new(12))
            $mdia = Get-AudioTestBox -Type 'mdia' -Payload $hdlr
            $tracks += Get-AudioTestBox -Type 'trak' -Payload $mdia
        }
        $moov = [byte[]]@()
        if ($tracks.Length -gt 0) { $moov = Get-AudioTestBox -Type 'moov' -Payload $tracks }
        return ,([byte[]]($ftyp + $moov + (Get-AudioTestBox -Type 'mdat' -Payload $Payload)))
    }

    function Get-AudioTestOgg {
        param([string]$Codec = 'opus', [byte]$Flags = 2)
        $packet = switch ($Codec) {
            'opus' { [Text.Encoding]::ASCII.GetBytes('OpusHead') + [byte[]](1,1) + [byte[]]::new(9) }
            'vorbis' {
                $vorbis = [byte[]]([byte[]](1) + [Text.Encoding]::ASCII.GetBytes('vorbis') + [byte[]]::new(23))
                $vorbis[11] = 1
                [BitConverter]::GetBytes([int]44100).CopyTo($vorbis, 12)
                $vorbis[29] = 1
                $vorbis
            }
            'speex' {
                $speex = [byte[]]([Text.Encoding]::ASCII.GetBytes('Speex   ') + [byte[]]::new(72))
                [BitConverter]::GetBytes([int]80).CopyTo($speex, 32)
                [BitConverter]::GetBytes([int]16000).CopyTo($speex, 36)
                [BitConverter]::GetBytes([int]1).CopyTo($speex, 48)
                $speex
            }
            'theora' { [byte[]](128) + [Text.Encoding]::ASCII.GetBytes('theora') + [byte[]]::new(16) }
            default { [Text.Encoding]::ASCII.GetBytes('unknown-packet') }
        }
        $header = [byte[]]::new(27)
        [Text.Encoding]::ASCII.GetBytes('OggS').CopyTo($header, 0)
        $header[5] = $Flags
        $header[26] = 1
        return ,([byte[]]($header + [byte[]]@($packet.Length) + $packet))
    }
}

Describe 'A035: ranked audio hints and deterministic enclosure selection' -Tag 'Unit', 'A035' {
    It 'recognizes supported audio MIME <Mime> with canonical extension <Extension>' -ForEach @(
        @{ Mime = 'audio/mpeg'; Extension = 'mp3' },
        @{ Mime = 'Audio/MPEG; charset=binary'; Extension = 'mp3' },
        @{ Mime = 'audio/mp4'; Extension = 'm4a' },
        @{ Mime = 'audio/ogg'; Extension = 'ogg' },
        @{ Mime = 'audio/opus'; Extension = 'ogg' },
        @{ Mime = 'audio/wav'; Extension = 'wav' },
        @{ Mime = 'audio/flac'; Extension = 'flac' }
    ) {
        $hint = Get-PodcastAudioHint -ContentType $Mime -Url 'https://media.example.invalid/download?signature=privateHintCanary'
        $hint.Eligible | Should -BeTrue
        $hint.Rank | Should -Be 2
        $hint.Extension | Should -BeExactly $Extension
        $hint.Reason | Should -Not -Match 'privateHintCanary|media\.example\.invalid|signature'
    }

    It 'uses the path suffix with <MimeCase> MIME for <Extension>' -ForEach @(
        foreach ($mime in @('', 'application/octet-stream')) {
            foreach ($extension in @('mp3', 'm4a', 'ogg', 'opus', 'wav', 'flac')) {
                @{ Mime = $mime; MimeCase = $(if ($mime) { 'generic' } else { 'missing' }); Extension = $extension }
            }
        }
    ) {
        $hint = Get-PodcastAudioHint -ContentType $Mime -Url ('https://media.example.invalid/audio.' + $Extension.ToUpperInvariant() + '?signature=privateHintCanary#section')
        $hint.Eligible | Should -BeTrue
        $hint.Rank | Should -Be 1
        $expected = if ($Extension -eq 'opus') { 'ogg' } else { $Extension }
        $hint.Extension | Should -BeExactly $expected
    }

    It 'keeps an extensionless generic candidate eligible for later byte inspection' {
        $hint = Get-PodcastAudioHint -ContentType 'application/octet-stream' -Url 'https://media.example.invalid/download?id=one'
        $hint.Eligible | Should -BeTrue
        $hint.Rank | Should -Be 0
        $hint.Extension | Should -BeNullOrEmpty
    }

    It 'does not treat an audio-looking query value as a path suffix' {
        $hint = Get-PodcastAudioHint -ContentType '' -Url 'https://media.example.invalid/download?filename=episode.mp3'
        $hint.Rank | Should -Be 0
        $hint.Extension | Should -BeNullOrEmpty
    }

    It 'rejects explicit unsupported type <Mime> despite an MP3-looking path' -ForEach @(
        @{ Mime = 'video/mp4' }, @{ Mime = 'text/html' }, @{ Mime = 'audio/aac' }, @{ Mime = 'audio/midi' }
    ) {
        $hint = Get-PodcastAudioHint -ContentType $Mime -Url 'https://media.example.invalid/episode.mp3'
        $hint.Eligible | Should -BeFalse
    }

    It 'keeps the first eligible known audio URL when later alternatives have stronger MIME hints' {
        [xml]$document = '<item><enclosure url="https://media.example.invalid/weak.mp3" type="application/octet-stream" length="10" />' +
            '<enclosure url="https://media.example.invalid/download?id=strong" type="audio/mp4" length="20" />' +
            '<enclosure url="https://media.example.invalid/later.flac" type="audio/flac" length="30" /></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.Url | Should -BeExactly 'https://media.example.invalid/weak.mp3'
        $episode.EnclosureLength | Should -Be 10L
        $episode.MediaContentType | Should -BeExactly 'application/octet-stream'
        $episode.MediaExtension | Should -BeExactly 'mp3'
        ($episode.Candidates -is [array]) | Should -BeTrue
        $episode.Candidates.Count | Should -Be 3
    }

    It 'chooses known audio ahead of an earlier generic unknown candidate' {
        [xml]$document = '<item><enclosure url="https://media.example.invalid/unknown" type="application/octet-stream" />' +
            '<enclosure url="https://media.example.invalid/audio" type="audio/mp4" length="20" />' +
            '<enclosure url="https://media.example.invalid/later.mp3" /></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.Url | Should -BeExactly 'https://media.example.invalid/audio'
        $episode.EnclosureLength | Should -Be 20L
        $episode.MediaExtension | Should -BeExactly 'm4a'
    }

    It 'preserves the established URL-only identity when a publisher adds a later typed alternative' {
        [xml]$document = '<item><enclosure url="https://media.example.invalid/established.mp3" />' +
            '<enclosure url="https://media.example.invalid/new.m4a" type="audio/mp4" /></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.Url | Should -BeExactly 'https://media.example.invalid/established.mp3'
        (Get-PodcastEpisodeIdentity -Episode $episode -FeedId ('a' * 64)).Key |
            Should -BeExactly 'media-url:https://media.example.invalid/established.mp3'
    }

    It 'breaks equal supported MIME ranks in source order for Atom links' {
        [xml]$document = '<entry xmlns="http://www.w3.org/2005/Atom"><link rel="alternate" href="https://show.example.invalid/entry" />' +
            '<link rel="enclosure" href="https://media.example.invalid/first" type="audio/mpeg" length="100" />' +
            '<link rel="enclosure" href="https://media.example.invalid/second.m4a" type="audio/mp4" length="200" /></entry>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.Url | Should -BeExactly 'https://media.example.invalid/first'
        $episode.EnclosureLength | Should -Be 100L
        $episode.MediaExtension | Should -BeExactly 'mp3'
        $episode.Candidates.Count | Should -Be 2
    }

    It 'prefers supported audio MIME to a contradictory weak filename hint' {
        [xml]$document = '<item><enclosure url="https://media.example.invalid/misnamed.m4a" type="audio/mpeg" /></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.MediaExtension | Should -BeExactly 'mp3'
        New-EpisodeFileName -Episode $episode -Index 1 | Should -Match '\.mp3$'
    }

    It 'keeps a generic candidate ahead of later generic candidates when neither has a format hint' {
        [xml]$document = '<item><enclosure url="https://media.example.invalid/one" type="application/octet-stream" />' +
            '<enclosure url="https://media.example.invalid/two" /></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.Url | Should -BeExactly 'https://media.example.invalid/one'
        $episode.MediaExtension | Should -BeNullOrEmpty
    }

    It 'returns no selected media URL when all declared candidates are unsupported' {
        [xml]$document = '<item><guid>unsupported-one</guid><enclosure url="https://media.example.invalid/video.mp3" type="video/mp4" />' +
            '<enclosure url="https://media.example.invalid/audio.aac" type="audio/aac" /></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        $episode.Url | Should -BeNullOrEmpty
        $episode.Candidates.Count | Should -Be 2
        $episode.MediaSelectionReason | Should -Not -BeNullOrEmpty
    }
}

Describe 'A035: canonical audio filenames and manual episode compatibility' -Tag 'Unit', 'A035' {
    It 'uses a declared supported MIME for an extensionless <Extension> URL' -ForEach @(
        @{ Mime = 'audio/mpeg'; Extension = 'mp3' }, @{ Mime = 'audio/mp4'; Extension = 'm4a' },
        @{ Mime = 'audio/opus'; Extension = 'ogg' }, @{ Mime = 'audio/wav'; Extension = 'wav' }, @{ Mime = 'audio/flac'; Extension = 'flac' }
    ) {
        [xml]$document = '<item><title>Synthetic</title><guid>one</guid><enclosure url="https://media.example.invalid/download?id=one" type="' + $Mime + '" /></item>'
        $episode = Get-EpisodeData -XmlItem $document.DocumentElement
        New-EpisodeFileName -Episode $episode -Index 1 | Should -Match ('-[a-f0-9]{64}\.' + $Extension + '$')
    }

    It 'preserves manual episode filename behavior for <Case>' -ForEach @(
        @{ Case = 'MP3 URL'; Url = 'https://media.example.invalid/one.mp3'; Expected = 'mp3' },
        @{ Case = 'M4A URL'; Url = 'https://media.example.invalid/one.M4A?signature=one'; Expected = 'm4a' },
        @{ Case = 'extensionless URL'; Url = 'https://media.example.invalid/download'; Expected = 'mp3' }
    ) {
        $episode = [pscustomobject]@{ Title = 'Synthetic'; Guid = 'one'; Url = $Url; PubDate = $null }
        New-EpisodeFileName -Episode $episode -Index 1 | Should -Match ('\.' + $Expected + '$')
        $episode.PSObject.Properties.Name | Should -Not -Contain 'MediaExtension'
    }

    It 'allows a detected extension override without changing publisher identity or episode metadata' {
        $episode = [pscustomobject]@{ Title = 'Synthetic'; Guid = 'one'; Url = 'https://media.example.invalid/download'; PubDate = $null; MediaExtension = 'mp3' }
        $beforeIdentity = Get-EpisodeIdentityKey -Episode $episode
        New-EpisodeFileName -Episode $episode -Index 1 -Extension 'flac' | Should -Match '-[a-f0-9]{64}\.flac$'
        Get-EpisodeIdentityKey -Episode $episode | Should -BeExactly $beforeIdentity
        $episode.MediaExtension | Should -BeExactly 'mp3'
    }

    It 'retains the full identifier and filename budget when a detected extension is longer' {
        $episode = [pscustomobject]@{ Title = ('long title ' * 50); Guid = 'one'; Url = 'https://media.example.invalid/download'; PubDate = $null }
        $name = New-EpisodeFileName -Episode $episode -Index 1 -MaxLength 80 -Extension 'flac'
        $name.Length | Should -BeLessOrEqual 80
        $name | Should -Match '-[a-f0-9]{64}\.flac$'
    }
}

Describe 'A035: bounded body evidence and canonical audio extensions' -Tag 'Unit', 'A035' {
    BeforeEach { $script:FormatPath = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.tmp') }

    It 'recognizes decisive <Format> bytes with generic headers and preserves them exactly' -ForEach @(
        @{ Format = 'MP3'; Extension = 'mp3' }, @{ Format = 'WAVE'; Extension = 'wav' }, @{ Format = 'FLAC'; Extension = 'flac' }
    ) {
        $bytes = switch ($Format) { 'MP3' { $script:FormatMp3 }; 'WAVE' { $script:FormatWave }; 'FLAC' { $script:FormatFlac } }
        [IO.File]::WriteAllBytes($script:FormatPath, $bytes)
        $before = [BitConverter]::ToString([IO.File]::ReadAllBytes($script:FormatPath))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -HttpContentLength $bytes.Length -ContentType 'application/octet-stream'
        $result.Valid | Should -BeTrue
        $result.Extension | Should -BeExactly $Extension
        $result.InspectedBytes | Should -BeLessOrEqual 65536
        $result.Verification | Should -Not -Match 'decoded|full_audio|complete_audio'
        [BitConverter]::ToString([IO.File]::ReadAllBytes($script:FormatPath)) | Should -BeExactly $before
    }

    It 'lets decisive MP3 bytes override conflicting feed HTTP and URL hints with fixed private warnings' {
        [IO.File]::WriteAllBytes($script:FormatPath, $script:FormatMp3)
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/mp4; privateHeaderCanary=one' `
            -EnclosureContentType 'audio/flac; privatePublisherCanary=two' -MediaUrl 'https://media.example.invalid/privateUrlCanary.wav?signature=privateQueryCanary'
        $result.Valid | Should -BeTrue
        $result.Extension | Should -BeExactly 'mp3'
        $result.Warnings.Count | Should -BeGreaterThan 0
        ($result.Warnings -join ' ') | Should -Not -Match 'private(?:Header|Publisher|Url|Query)Canary|media\.example\.invalid|signature='
    }

    It 'rejects HTML despite audio declarations and an MP3-looking URL' {
        [IO.File]::WriteAllText($script:FormatPath, '<html><body>privateBodyCanary</body></html>', [Text.UTF8Encoding]::new($false))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/mpeg' `
            -EnclosureContentType 'audio/mpeg' -MediaUrl 'https://media.example.invalid/episode.mp3'
        $result.Valid | Should -BeFalse
        $result.Category | Should -BeExactly 'non_audio_text'
        $result.Extension | Should -BeNullOrEmpty
        ($result | ConvertTo-Json -Depth 3) | Should -Not -Match 'privateBodyCanary'
    }

    It 'rejects unsupported MPEG <Layer> rather than giving it an MP3 extension' -ForEach @(
        @{ Layer = 'Layer I'; HeaderByte = 255 }, @{ Layer = 'Layer II'; HeaderByte = 253 }
    ) {
        $bytes = [byte[]]([byte[]]@(255, $HeaderByte, 144, 0) + [byte[]]::new(2048))
        [IO.File]::WriteAllBytes($script:FormatPath, $bytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/mpeg'
        $result.Valid | Should -BeFalse
        $result.Category | Should -BeExactly 'unsupported_media'
        $result.Extension | Should -BeNullOrEmpty
    }

    It 'recognizes first Ogg BOS packet <Codec> as canonical Ogg audio' -ForEach @(
        @{ Codec = 'opus' }, @{ Codec = 'vorbis' }, @{ Codec = 'speex' }
    ) {
        [IO.File]::WriteAllBytes($script:FormatPath, (Get-AudioTestOgg -Codec $Codec))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'application/octet-stream'
        $result.Valid | Should -BeTrue
        $result.DetectedFormat | Should -BeExactly 'ogg_container'
        $result.Extension | Should -BeExactly 'ogg'
        $result.InspectedBytes | Should -BeLessOrEqual 65536
    }

    It 'rejects an unknown Ogg packet as ambiguous despite audio MIME' {
        [IO.File]::WriteAllBytes($script:FormatPath, (Get-AudioTestOgg -Codec 'unknown'))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/ogg'
        $result.Valid | Should -BeFalse
        $result.Category | Should -BeExactly 'ambiguous_media'
        $result.Extension | Should -BeNullOrEmpty
    }

    It 'rejects a Theora Ogg packet as unsupported video despite an audio MIME' {
        [IO.File]::WriteAllBytes($script:FormatPath, (Get-AudioTestOgg -Codec 'theora'))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/ogg'
        $result.Valid | Should -BeFalse
        $result.Category | Should -BeExactly 'unsupported_media'
    }

    It 'rejects malformed Ogg evidence <Case>' -ForEach @(
        @{ Case = 'no BOS'; Mutation = 'flags' }, @{ Case = 'truncated lacing table'; Mutation = 'table' }, @{ Case = 'truncated packet'; Mutation = 'payload' }
    ) {
        $bytes = Get-AudioTestOgg -Codec 'opus'
        switch ($Mutation) {
            'flags' { $bytes[5] = 0 }
            'table' { $bytes[26] = 255 }
            'payload' { $bytes = [byte[]]$bytes[0..($bytes.Length - 2)] }
        }
        [IO.File]::WriteAllBytes($script:FormatPath, $bytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/ogg'
        $result.Valid | Should -BeFalse
        $result.Extension | Should -BeNullOrEmpty
        $result.Verification | Should -BeExactly 'none'
    }

    It 'recognizes bounded MP4 audio evidence <Case> as M4A' -ForEach @(
        @{ Case = 'M4A brand'; Brand = 'M4A '; Handlers = @() },
        @{ Case = 'generic brand with sound handler'; Brand = 'isom'; Handlers = @('soun') }
    ) {
        [IO.File]::WriteAllBytes($script:FormatPath, (Get-AudioTestMp4 -Brand $Brand -Handlers $Handlers))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'application/octet-stream'
        $result.Valid | Should -BeTrue
        $result.DetectedFormat | Should -BeExactly 'mp4_container'
        $result.Extension | Should -BeExactly 'm4a'
        $result.InspectedBytes | Should -BeLessOrEqual 65536
    }

    It 'rejects a generic MP4 without structural audio proof despite an audio MIME' {
        [IO.File]::WriteAllBytes($script:FormatPath, (Get-AudioTestMp4 -Brand 'isom' -Payload ([Text.Encoding]::ASCII.GetBytes('fake soun hdlr data and no actual audio track'))))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/mp4'
        $result.Valid | Should -BeFalse
        $result.Category | Should -BeExactly 'ambiguous_media'
    }

    It 'rejects structural MP4 video evidence for <Case>' -ForEach @(
        @{ Case = 'video track'; Handlers = @('vide') },
        @{ Case = 'mixed sound and video tracks'; Handlers = @('soun', 'vide') }
    ) {
        [IO.File]::WriteAllBytes($script:FormatPath, (Get-AudioTestMp4 -Brand 'M4A ' -Handlers $Handlers))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/mp4'
        $result.Valid | Should -BeFalse
        $result.Category | Should -BeExactly 'unsupported_media'
    }

    It 'does not mistake strings inside media payload for structural video handlers' {
        [IO.File]::WriteAllBytes($script:FormatPath, (Get-AudioTestMp4 -Payload ([Text.Encoding]::ASCII.GetBytes('fake vide hdlr text within mdat audio payload'))))
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/mp4'
        $result.Valid | Should -BeTrue
        $result.Extension | Should -BeExactly 'm4a'
    }

    It 'does not infer audio from a nested box whose declared size exceeds the stored container' {
        $bytes = Get-AudioTestMp4 -Brand 'isom' -Handlers @('soun')
        # The first box is a 20-byte ftyp; corrupt the following moov size.
        $bytes[20] = 0; $bytes[21] = 0; $bytes[22] = 127; $bytes[23] = 255
        [IO.File]::WriteAllBytes($script:FormatPath, $bytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/mp4'
        $result.Valid | Should -BeFalse
        $result.Category | Should -BeExactly 'ambiguous_media'
    }

    It 'does not claim audio proof from a handler outside the bounded inspection budget' {
        $ftyp = Get-AudioTestBox -Type 'ftyp' -Payload ([Text.Encoding]::ASCII.GetBytes('isom') + [byte[]]::new(4) + [Text.Encoding]::ASCII.GetBytes('isom'))
        $hdlr = Get-AudioTestBox -Type 'hdlr' -Payload ([byte[]]::new(8) + [Text.Encoding]::ASCII.GetBytes('soun') + [byte[]]::new(12))
        $mdia = Get-AudioTestBox -Type 'mdia' -Payload $hdlr
        $trak = Get-AudioTestBox -Type 'trak' -Payload $mdia
        $moov = Get-AudioTestBox -Type 'moov' -Payload $trak
        $bytes = [byte[]]($ftyp + (Get-AudioTestBox -Type 'mdat' -Payload ([byte[]]::new(65536))) + $moov)
        [IO.File]::WriteAllBytes($script:FormatPath, $bytes)
        $result = Test-PodcastMediaFile -LiteralPath $script:FormatPath -TransferCompleted $true -ContentType 'audio/mp4'
        $result.Valid | Should -BeFalse
        $result.Category | Should -BeExactly 'ambiguous_media'
        $result.InspectedBytes | Should -BeLessOrEqual 65536
    }
}

Describe 'A035: detected final paths respect archive publication boundaries' -Tag 'Unit', 'A035' {
    BeforeEach {
        $script:FormatTransferRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($script:FormatTransferRoot)
        $waveTransferFixture = $script:FormatWave
        Mock Invoke-PodcastMediaRequest {
            $DestinationStream.Write($waveTransferFixture, 0, $waveTransferFixture.Length)
            [pscustomobject]@{ Completed = $true; Bytes = $waveTransferFixture.Length; ContentLength = $waveTransferFixture.Length; ContentType = 'application/octet-stream' }
        }
    }

    It 'records the detected final filename before publishing the untouched audio bytes' {
        $preparedEvidence = [pscustomobject]@{ Result = $null; FinalExisted = $true }
        $resolve = { param($Validation) 'episode.' + $Validation.Extension }
        $prepare = {
            param($Result)
            $preparedEvidence.Result = $Result
            $preparedEvidence.FinalExisted = Test-Path -LiteralPath $Result.File
        }
        $result = Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/download' -Root $script:FormatTransferRoot `
            -RelativePath 'episode.mp3' -ResolveFinalPath $resolve -BeforeFinalize $prepare
        $result.RelativePath | Should -BeExactly 'episode.wav'
        $preparedEvidence.Result.RelativePath | Should -BeExactly 'episode.wav'
        $preparedEvidence.FinalExisted | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:FormatTransferRoot 'episode.mp3') | Should -BeFalse
        [BitConverter]::ToString([IO.File]::ReadAllBytes($result.File)) | Should -BeExactly ([BitConverter]::ToString($script:FormatWave))
        @(Get-ChildItem -LiteralPath $script:FormatTransferRoot -Force).Count | Should -Be 1
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly
    }

    It 'preserves an existing detected-format destination before recording prepared history' {
        $existingPath = Join-Path $script:FormatTransferRoot 'episode.wav'
        [IO.File]::WriteAllText($existingPath, 'unknown existing bytes')
        $preparedEvidence = [pscustomobject]@{ Calls = 0 }
        $resolve = { param($Validation) 'episode.' + $Validation.Extension }
        $prepare = { param($Result) $null = $Result; $preparedEvidence.Calls++ }
        { Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/download' -Root $script:FormatTransferRoot `
                -RelativePath 'episode.mp3' -ResolveFinalPath $resolve -BeforeFinalize $prepare } | Should -Throw '*already exists*'
        [IO.File]::ReadAllText($existingPath) | Should -BeExactly 'unknown existing bytes'
        $preparedEvidence.Calls | Should -Be 0
        @(Get-ChildItem -LiteralPath $script:FormatTransferRoot -Force).Count | Should -Be 1
    }

    It 'retains the separate legacy redownload allocation when WAVE evidence corrects its extension' {
        $feedId = Get-PodcastNameHash -IdentityKey 'feed:synthetic-legacy-allocation'
        $episode = [pscustomobject]@{
            Title = 'Legacy replacement'; Guid = 'legacy-replacement'; AtomId = ''
            PubDate = $null; Url = 'https://media.example.invalid/download'; EnclosureLength = $null
        }
        $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $feedId
        $allocationHash = Get-PodcastNameHash -IdentityKey ([guid]::NewGuid().ToString('N'))
        $planned = [pscustomobject]@{
            Episode = $episode; EpisodeId = $identity.Id; IdentitySource = $identity.Source
            IdentityFingerprint = $identity.Fingerprint; FileName = 'Owner original.mp3'
            StateRecord = $null; NewFileIdentityHash = $allocationHash; MaxFileNameLength = 100
        }
        $originalPath = Join-Path $script:FormatTransferRoot $planned.FileName
        [IO.File]::WriteAllBytes($originalPath, $script:FormatMp3)
        $originalEvidence = Get-PodcastFileEvidence -Root $script:FormatTransferRoot -RelativePath $planned.FileName
        $oldRecord = New-PodcastEpisodeRecord -Planned $planned
        $oldRecord.status = 'adopted'
        $oldRecord.bytes = $originalEvidence.Bytes
        $oldRecord.local_sha256 = $originalEvidence.Sha256
        $oldRecord.completed_utc = [datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        $oldRecord.verification = [pscustomobject]@{
            method = 'owner-approved-local-signature'; media_kind = 'mpeg-audio'
            notes = @('local-signature-only; transfer-completeness-unverified')
        }
        $planned.StateRecord = $oldRecord
        $planned.FileName = New-EpisodeFileName -Episode $episode -IdentityHash $allocationHash -MaxLength $planned.MaxFileNameLength
        $historyLock = Enter-PodcastHistoryLock -Root $script:FormatTransferRoot
        try {
            $state = New-PodcastHistory -FeedId $feedId -FeedAliasFingerprint $feedId -SchemaVersion 2
            $state.generation = 1
            $state.episodes = @($oldRecord)
            $context = @{ Lock = $historyLock; State = Write-PodcastHistory -Lock $historyLock -State $state }
            $result = Invoke-PodcastRecordedTransfer -Context $context -Planned $planned

            $result.RelativePath.EndsWith(('-' + $allocationHash + '.wav'), [StringComparison]::Ordinal) | Should -BeTrue
            $result.RelativePath | Should -Not -Match ([regex]::Escape($identity.Id))
            $result.RelativePath.Length | Should -BeLessOrEqual $planned.MaxFileNameLength
            $completed = (Read-PodcastHistory -Root $script:FormatTransferRoot).episodes[0]
            $completed.relative_path | Should -BeExactly $result.RelativePath
            $completed.status | Should -BeExactly 'transfer_verified'
            $completed.verification.media_kind | Should -BeExactly 'wave'
            $prepared = (Read-PodcastHistoryFile -Root $script:FormatTransferRoot -RelativePath '.upd/state.json.bak').episodes[0]
            $prepared.relative_path | Should -BeExactly $result.RelativePath
            $prepared.status | Should -BeExactly 'prepared'
            [BitConverter]::ToString([IO.File]::ReadAllBytes($result.File)) | Should -BeExactly ([BitConverter]::ToString($script:FormatWave))
            (Get-PodcastFileEvidence -Root $script:FormatTransferRoot -RelativePath 'Owner original.mp3').Sha256 | Should -BeExactly $originalEvidence.Sha256
            Test-Path -LiteralPath (Join-Path $script:FormatTransferRoot $planned.FileName) | Should -BeFalse
            Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly
        }
        finally { $historyLock.Stream.Dispose() }
    }

    It 'refuses a resolved destination outside the planned directory before prepared history' {
        $preparedEvidence = [pscustomobject]@{ Calls = 0 }
        $resolve = { param($Validation) $null = $Validation; '..\outside.wav' }
        $prepare = { param($Result) $null = $Result; $preparedEvidence.Calls++ }
        { Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/download' -Root $script:FormatTransferRoot `
                -RelativePath 'episode.mp3' -ResolveFinalPath $resolve -BeforeFinalize $prepare } |
            Should -Throw 'Resolved media destination must remain in the planned directory.'
        $preparedEvidence.Calls | Should -Be 0
        @(Get-ChildItem -LiteralPath $script:FormatTransferRoot -Force).Count | Should -Be 0
    }

    It 'corrects an extensionless failed allocation on retry while resume retains the provisional path' {
        $feedId = Get-PodcastNameHash -IdentityKey 'feed:synthetic-failed-allocation'
        $episode = [pscustomobject]@{
            Title = 'Retry allocation'; Guid = 'retry-allocation'; AtomId = ''
            PubDate = $null; Url = 'https://media.example.invalid/download'; EnclosureLength = $null
        }
        $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $feedId
        $planned = [pscustomobject]@{
            Episode = $episode; EpisodeId = $identity.Id; IdentitySource = $identity.Source
            IdentityFingerprint = $identity.Fingerprint
            FileName = New-EpisodeFileName -Episode $episode -IdentityHash $identity.Id
            StateRecord = $null
        }
        $failedRecord = New-PodcastEpisodeRecord -Planned $planned
        $planned.StateRecord = $failedRecord
        $mp4TransferFixture = Get-AudioTestMp4
        $resumeObservation = [pscustomobject]@{ RelativePath = $null }
        Mock Invoke-PodcastMediaRequest {
            param($DestinationStream, $OnResponse)
            $null = & $OnResponse ([pscustomobject]@{
                ResumeSupported = $true; FinalUriFingerprint = 'b' * 64; ETag = '"fixture-retry"'
                TotalLength = $mp4TransferFixture.Length; ContentType = 'application/octet-stream'
            })
            $DestinationStream.Write($mp4TransferFixture, 0, $mp4TransferFixture.Length)
            $resumeObservation.RelativePath = (Read-PodcastResumeState -Root $script:FormatTransferRoot -EpisodeId $identity.Id).relative_path
            [pscustomobject]@{
                Completed = $true; Bytes = $mp4TransferFixture.Length
                ContentLength = $mp4TransferFixture.Length; ContentType = 'application/octet-stream'
            }
        }
        $historyLock = Enter-PodcastHistoryLock -Root $script:FormatTransferRoot
        try {
            $state = New-PodcastHistory -FeedId $feedId -FeedAliasFingerprint $feedId
            $state.generation = 1
            $state.episodes = @($failedRecord)
            $context = @{ Lock = $historyLock; State = Write-PodcastHistory -Lock $historyLock -State $state }
            $result = Invoke-PodcastRecordedTransfer -Context $context -Planned $planned

            $resumeObservation.RelativePath | Should -BeExactly $planned.FileName
            $resumeObservation.RelativePath | Should -Match '\.mp3$'
            $result.RelativePath | Should -Match '\.m4a$'
            $completed = (Read-PodcastHistory -Root $script:FormatTransferRoot).episodes[0]
            $completed.relative_path | Should -BeExactly $result.RelativePath
            $completed.status | Should -BeExactly 'transfer_verified'
            $prepared = (Read-PodcastHistoryFile -Root $script:FormatTransferRoot -RelativePath '.upd/state.json.bak').episodes[0]
            $prepared.relative_path | Should -BeExactly $result.RelativePath
            $prepared.status | Should -BeExactly 'prepared'
            [BitConverter]::ToString([IO.File]::ReadAllBytes($result.File)) | Should -BeExactly ([BitConverter]::ToString($mp4TransferFixture))
            Test-Path -LiteralPath (Join-Path $script:FormatTransferRoot $planned.FileName) | Should -BeFalse
            Read-PodcastResumeState -Root $script:FormatTransferRoot -EpisodeId $identity.Id | Should -BeNullOrEmpty
            Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly
        }
        finally { $historyLock.Stream.Dispose() }
    }
}
