BeforeAll {
    $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    Mock Read-Host { throw 'Unexpected unit prompt.' }
    Mock Invoke-WebRequest { throw 'The legacy web/DOM parser must not be called.' }
    . $script:DownloaderPath -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Unexpected network client in unit test.' }
}

Describe 'A007/A023: centralized metadata requests without legacy DOM parsing' -Tag 'Unit', 'A007', 'A023' {
    It 'delegates metadata to the bounded adapter regardless of web cmdlet defaults' {
        Mock Invoke-PodcastMetadataRequest { [pscustomobject]@{ Content = 'synthetic response' } }
        $previousDefaults = $PSDefaultParameterValues
        try {
            $PSDefaultParameterValues = @{ 'Invoke-WebRequest:UseBasicParsing' = $false }
            (Invoke-PodcastWebRequest -Uri 'https://feed.example.invalid/rss').Content | Should -Be 'synthetic response'
            $PSDefaultParameterValues['Invoke-WebRequest:UseBasicParsing'] | Should -BeFalse
            Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://feed.example.invalid/rss' }
            Should -Invoke Invoke-WebRequest -Times 0 -Exactly
        }
        finally { $PSDefaultParameterValues = $previousDefaults }
    }

    It 'has no file-output parameter on the metadata adapter' {
        $destination = Join-Path $TestDrive 'must-not-exist.mp3'
        { Invoke-PodcastWebRequest -Uri 'https://media.example.invalid/episode.mp3' -OutFile $destination } | Should -Throw
        Test-Path -LiteralPath $destination | Should -BeFalse
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
    }

    It 'preserves safe request errors' {
        Mock Invoke-PodcastMetadataRequest { throw 'synthetic request failure' }
        { Invoke-PodcastWebRequest -Uri 'https://feed.example.invalid/rss' } | Should -Throw '*synthetic request failure*'
    }

    It 'resolves interactive HTML links against the final redirected page' {
        Mock Write-Host {}
        Mock Read-Host {
            if ($Prompt -eq 'Use this feed? (Y/n)') { return 'y' }
            return 'https://show.example.invalid/podcast'
        }
        Mock Invoke-PodcastMetadataRequest {
            [pscustomobject]@{ Content = '<html><head><link type="application/rss+xml" href="feed.xml"></head></html>'; FinalUri = [Uri]'https://destination.example.invalid/shows/index.html' }
        }
        Get-FeedUrlInteractive | Should -Be 'https://destination.example.invalid/shows/feed.xml'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
        Should -Invoke Read-Host -Times 2 -Exactly
    }

    It 'retains the original feed identity across an allowed redirect' {
        Mock Invoke-PodcastMetadataRequest {
            [pscustomobject]@{ Content = '<rss><channel><item><title>Synthetic episode</title></item></channel></rss>'; FinalUri = [Uri]'https://cdn.example.invalid/feed.xml' }
        }
        $resolved = Resolve-PodcastItems -Feeds 'https://feed.example.invalid/rss'
        $resolved.Items.Count | Should -Be 1
        $resolved.Url | Should -Be 'https://feed.example.invalid/rss'
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
    }

    It 'rejects an unsafe discovered feed target <Target>' -TestCases @(
        @{ Target = 'file:///C:/secret.xml' }
        @{ Target = 'https://show.example.invalid/private file.xml' }
        @{ Target = '\\other.example.invalid\feed.xml' }
    ) {
        param($Target)
        $html = '<link type="application/rss+xml" href="' + $Target + '">'
        { Find-RssInHtml -Html $html -BaseUrl 'https://show.example.invalid/' } | Should -Throw '*network policy*'
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
    }

    It 'rejects an unsafe enclosure during preview before writes or media' {
        Mock Invoke-PodcastMetadataRequest {
            [pscustomobject]@{ Content = '<rss><channel><title>Unsafe</title><item><title>Synthetic</title><enclosure url="file:///C:/secret.mp3" /></item></channel></rss>' }
        }
        Mock Invoke-PodcastMediaRequest { throw 'Media must not be requested.' }
        $root = Join-Path $TestDrive 'absent'
        { & $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/rss' -Mode All -OutputPath $root -WhatIf } | Should -Throw
        Test-Path -LiteralPath $root | Should -BeFalse
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
    }
}
