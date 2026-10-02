BeforeAll {
    $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    Mock Read-Host { throw 'Unexpected unit prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request.' }
    . $script:DownloaderPath -OutputPath $TestDrive
}

Describe 'A010: media destination write boundaries' -Tag 'Unit', 'A010' {
    BeforeEach {
        $script:OutputRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($script:OutputRoot)
        $script:FinalPath = Join-Path $script:OutputRoot 'episode.mp3'
    }

    It 'publishes the original bytes through a unique sibling file' {
        Mock Invoke-WebRequest { [IO.File]::WriteAllBytes($OutFile, [byte[]](1, 2, 3, 255)) }
        Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/episode.mp3' -Root $script:OutputRoot -RelativePath 'episode.mp3'
        [BitConverter]::ToString([IO.File]::ReadAllBytes($script:FinalPath)) | Should -Be '01-02-03-FF'
        @(Get-ChildItem -LiteralPath $script:OutputRoot -Force).Count | Should -Be 1
        Should -Invoke Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $UseBasicParsing -and $OutFile -like '*.tmp' }
    }

    It 'preserves an unknown existing destination without requesting media' {
        [IO.File]::WriteAllText($script:FinalPath, 'unknown existing bytes')
        $before = (Get-Item -LiteralPath $script:FinalPath).LastWriteTimeUtc
        { Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/episode.mp3' -Root $script:OutputRoot -RelativePath 'episode.mp3' } |
            Should -Throw '*already exists*'
        [IO.File]::ReadAllText($script:FinalPath) | Should -Be 'unknown existing bytes'
        (Get-Item -LiteralPath $script:FinalPath).LastWriteTimeUtc | Should -Be $before
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
    }

    It 'preserves a competing file created during the request' {
        Mock Invoke-WebRequest {
            [IO.File]::WriteAllText($OutFile, 'downloaded bytes')
            [IO.File]::WriteAllText($script:FinalPath, 'competing bytes')
        }
        { Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/episode.mp3' -Root $script:OutputRoot -RelativePath 'episode.mp3' } |
            Should -Throw
        [IO.File]::ReadAllText($script:FinalPath) | Should -Be 'competing bytes'
        @(Get-ChildItem -LiteralPath $script:OutputRoot -Force).Count | Should -Be 1
    }

    It 'cleans only the failed attempt temporary file and keeps unknown partials' {
        $unowned = Join-Path $script:OutputRoot '.upd-unknown.tmp'
        [IO.File]::WriteAllText($unowned, 'unowned bytes')
        Mock Invoke-WebRequest {
            [IO.File]::WriteAllText($OutFile, 'partial bytes')
            throw 'Synthetic interrupted request.'
        }
        { Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/episode.mp3' -Root $script:OutputRoot -RelativePath 'episode.mp3' } |
            Should -Throw '*Synthetic interrupted request*'
        Test-Path -LiteralPath $script:FinalPath | Should -BeFalse
        [IO.File]::ReadAllText($unowned) | Should -Be 'unowned bytes'
        @(Get-ChildItem -LiteralPath $script:OutputRoot -Force).Count | Should -Be 1
    }

    It 'refuses traversal before creating a temporary file or requesting media' {
        { Invoke-PodcastMediaTransfer -Uri 'https://media.example.invalid/episode.mp3' -Root $script:OutputRoot -RelativePath '..\escape.mp3' } |
            Should -Throw
        @(Get-ChildItem -LiteralPath $script:OutputRoot -Force).Count | Should -Be 0
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
    }

    It 'validates <Kind> before creating podcast folders or logs' -ForEach @(
        @{ Kind = 'hostile episode title'; Items = '<item><title>..</title><guid>one</guid><enclosure url="https://media.example.invalid/one.mp3" /></item>' }
        @{ Kind = 'conflicting episode identity'; Items = '<item><title>First</title><guid>same</guid><enclosure url="https://media.example.invalid/one.mp3" /></item><item><title>Second</title><guid>same</guid><enclosure url="https://media.example.invalid/two.mp3" /></item>' }
    ) {
        Mock Write-Host {}
        $response = [pscustomobject]@{ Content = '<rss><channel><title>Show</title>' + $Items + '</channel></rss>' }
        Mock Invoke-WebRequest { $response }
        { & $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/rss' -OutputPath $script:OutputRoot -Mode All } |
            Should -Throw
        @(Get-ChildItem -LiteralPath $script:OutputRoot -Force).Count | Should -Be 0
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly -ParameterFilter { $OutFile }
    }

    It 'fits both full identifiers under a tight absolute output-root budget' {
        Mock Write-Host {}
        Mock Write-Progress {}
        $response = [pscustomobject]@{ Content = '<rss><channel><title>Show</title><item><title>Episode</title><guid>one</guid><pubDate>2026-09-01</pubDate><enclosure url="https://media.example.invalid/one.mp3" /></item></channel></rss>' }
        Mock Invoke-WebRequest {
            if ($OutFile) { [IO.File]::WriteAllText($OutFile, 'synthetic media') }
            else { $response }
        }
        $tightRoot = Join-Path $TestDrive ('x' * (120 - $TestDrive.Length - 1))
        & $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/rss' -OutputPath $tightRoot -Mode All
        $media = @(Get-ChildItem -LiteralPath $tightRoot -File -Recurse -Filter '*.mp3')
        $media.Count | Should -Be 1
        $media[0].FullName.Length | Should -BeLessOrEqual 259
        $media[0].Name | Should -Match '^E-[a-f0-9]{64}\.mp3$'
        $media[0].Directory.Name | Should -Match '^Sh-[a-f0-9]{64}$'
    }
}
