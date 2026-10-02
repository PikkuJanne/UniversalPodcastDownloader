BeforeAll {
    $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    Mock Read-Host { throw 'Unexpected unit prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request in unit test.' }
    . $script:DownloaderPath -OutputPath $TestDrive
}

Describe 'A007: centralized safe web requests' -Tag 'Unit', 'A007' {
    It 'returns content with BasicParsing even when the caller default is false' {
        Mock Invoke-WebRequest { [pscustomobject]@{ Content = 'synthetic response' } }
        $previousDefaults = $PSDefaultParameterValues
        try {
            $PSDefaultParameterValues = @{ 'Invoke-WebRequest:UseBasicParsing' = $false }
            $result = Invoke-PodcastWebRequest -Uri 'https://feed.example.invalid/rss'
            $result.Content | Should -Be 'synthetic response'
            $PSDefaultParameterValues['Invoke-WebRequest:UseBasicParsing'] | Should -BeFalse
            Should -Invoke Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $UseBasicParsing -and -not $OutFile -and $Uri -eq 'https://feed.example.invalid/rss' }
        }
        finally { $PSDefaultParameterValues = $previousDefaults }
    }

    It 'forwards the media destination with BasicParsing' {
        Mock Invoke-WebRequest {}
        $destination = Join-Path $TestDrive 'synthetic.mp3'
        Invoke-PodcastWebRequest -Uri 'https://media.example.invalid/episode.mp3' -OutFile $destination
        Should -Invoke Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $UseBasicParsing -and $OutFile -eq $destination -and $Uri -eq 'https://media.example.invalid/episode.mp3' }
    }

    It 'preserves request errors' {
        Mock Invoke-WebRequest { throw 'synthetic request failure' }
        { Invoke-PodcastWebRequest -Uri 'https://feed.example.invalid/rss' } | Should -Throw '*synthetic request failure*'
    }

    It 'uses safe parsing for interactive HTML discovery' {
        Mock Write-Host {}
        Mock Read-Host {
            if ($Prompt -eq 'Use this feed? (Y/n)') { return 'y' }
            return 'https://show.example.invalid/podcast'
        }
        Mock Invoke-WebRequest {
            [pscustomobject]@{ Content = '<html><head><link type="application/rss+xml" href="/feed.xml"></head></html>' }
        }
        Get-FeedUrlInteractive | Should -Be 'https://show.example.invalid/feed.xml'
        Should -Invoke Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $UseBasicParsing -and $Uri -eq 'https://show.example.invalid/podcast' }
        Should -Invoke Read-Host -Times 2 -Exactly
    }

    It 'uses safe parsing when resolving a feed' {
        Mock Invoke-WebRequest { [pscustomobject]@{ Content = '<rss><channel><item><title>Synthetic episode</title></item></channel></rss>' } }
        $resolved = Resolve-PodcastItems -Feeds 'https://feed.example.invalid/rss'
        $resolved.Items.Count | Should -Be 1
        Should -Invoke Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $UseBasicParsing -and $Uri -eq 'https://feed.example.invalid/rss' }
    }
}
