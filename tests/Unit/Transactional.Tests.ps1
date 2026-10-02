BeforeAll {
    $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    . $script:DownloaderPath -OutputPath $TestDrive
    Mock Read-Host { throw 'Unexpected unit prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request.' }
    Mock Invoke-PodcastMediaRequest { throw 'Unexpected media request.' }
}

Describe 'A012: publisher enclosure length is advisory metadata' -Tag 'Unit', 'A012' {
    It 'extracts a selected RSS enclosure length <Value> as <Expected>' -ForEach @(
        @{ Value = '1689'; Expected = 1689L }
        @{ Value = '0'; Expected = 0L }
        @{ Value = ''; Expected = $null }
        @{ Value = '-1'; Expected = $null }
        @{ Value = 'wrong'; Expected = $null }
        @{ Value = '9223372036854775808'; Expected = $null }
    ) {
        [xml]$xml = '<item><title>Synthetic</title><enclosure url="https://media.example.invalid/audio.mp3" length="' + $Value + '" /></item>'
        (Get-EpisodeData -XmlItem $xml.DocumentElement).EnclosureLength | Should -Be $Expected
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
    }

    It 'takes the length from the selected Atom enclosure link' {
        [xml]$xml = '<entry xmlns="http://www.w3.org/2005/Atom"><link rel="alternate" length="1" href="https://feed.example.invalid/entry"/><link rel="enclosure" length="1689" href="https://media.example.invalid/audio.mp3"/></entry>'
        (Get-EpisodeData -XmlItem $xml.DocumentElement).EnclosureLength | Should -Be 1689L
    }
}

Describe 'A011: failed transfer attempts cannot become completed episodes' -Tag 'Unit', 'A011' {
    BeforeEach {
        Mock Write-Host {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        $script:OutputRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $temporaryNames = [Collections.Generic.List[string]]::new()
        Mock Invoke-PodcastMediaRequest {
            $temporaryNames.Add($DestinationStream.Name)
            throw 'Unexpected media request.'
        }
        $response = [pscustomobject]@{ Content = '<rss><channel><title>Failure fixture</title><item><title>Synthetic</title><guid>one</guid><enclosure url="https://media.example.invalid/audio.mp3"/></item></channel></rss>' }
        Mock Invoke-WebRequest { $response }
    }

    It 'retries invalid empty responses using new owned files and reports an incomplete run' {
        Mock Invoke-PodcastMediaRequest {
            $temporaryNames.Add($DestinationStream.Name)
            [pscustomobject]@{ Completed = $true; Bytes = 0L; ContentLength = 0L; ContentType = 'audio/mpeg' }
        }
        { & $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/rss' -OutputPath $script:OutputRoot -Mode All } |
            Should -Throw '*Download incomplete: 1 episode(s) failed*'
        Should -Invoke Invoke-PodcastMediaRequest -Times 3 -Exactly
        @($temporaryNames | Select-Object -Unique).Count | Should -Be 3
        @(Get-ChildItem -LiteralPath $script:OutputRoot -Recurse -Filter '*.mp3').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $script:OutputRoot -Recurse -Filter '*.tmp' -Force).Count | Should -Be 0
        Should -Invoke Write-Host -Times 0 -Exactly -ParameterFilter { $Object -like '[[]OK[]]*' }
        $log = Get-Content -LiteralPath (Get-ChildItem -LiteralPath $script:OutputRoot -Recurse -Filter '*.log').FullName -Raw
        $log | Should -Match 'Summary: Downloaded=0, Skipped=0, Failed=1'
        $log | Should -Not -Match 'Download succeeded|Run completed'
    }

    It 'closes the exclusive file handle after a request throws and preserves unknown files' {
        $null = [IO.Directory]::CreateDirectory($script:OutputRoot)
        $unknown = Join-Path $script:OutputRoot '.upd-unknown.tmp'
        [IO.File]::WriteAllText($unknown, 'unowned bytes')
        Mock Invoke-PodcastMediaRequest {
            $temporaryNames.Add($DestinationStream.Name)
            $DestinationStream.WriteByte(255)
            throw 'Synthetic transfer failure.'
        }
        { Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/audio.mp3' -Root $script:OutputRoot -RelativePath 'episode.mp3' } |
            Should -Throw '*Synthetic transfer failure*'
        $temporary = $temporaryNames[0]
        Test-Path -LiteralPath $temporary | Should -BeFalse
        # CreateNew proves the failed attempt no longer holds a handle or file.
        $stream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $stream.Dispose()
        [IO.File]::ReadAllText($unknown) | Should -Be 'unowned bytes'
        Test-Path -LiteralPath (Join-Path $script:OutputRoot 'episode.mp3') | Should -BeFalse
    }
}
