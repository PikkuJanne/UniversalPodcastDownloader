BeforeAll {
    $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    Mock Read-Host { throw 'Unit tests must not prompt.' }
    Mock Invoke-WebRequest { throw 'Unexpected network request in unit test.' }
    . $script:DownloaderPath -OutputPath $TestDrive
    Mock Get-PodcastHttpClient { throw 'Unit tests must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Unexpected metadata request in unit test.' }
    $script:FixtureMedia = [IO.File]::ReadAllBytes((Join-Path (Split-Path $script:DownloaderPath -Parent) 'tools/codex-handoff/fixtures/silence.mp3'))
}

Describe 'A006: array selection for every supported mode' -Tag 'Unit', 'A006' {
    It 'returns an array for <InputCount> entries in <Mode> with count <CustomCount>' -ForEach @(
        foreach ($inputCount in 0, 1, 3) {
            foreach ($mode in 'Latest', 'All', 'Custom') {
                foreach ($count in $(if ($mode -eq 'Custom') { 1, 2, 5 } else { 0 })) {
                    @{
                        InputCount = $inputCount
                        Mode = $mode
                        CustomCount = $count
                        ExpectedCount = if ($mode -eq 'Latest') { [Math]::Min(1, $inputCount) }
                            elseif ($mode -eq 'All') { $inputCount }
                            else { [Math]::Min($count, $inputCount) }
                    }
                }
            }
        }
    ) {
        $episodes = @(for ($i = 0; $i -lt $InputCount; $i++) {
            [pscustomobject]@{ Title = "Episode $i"; PubDate = ([datetime]'2026-09-01').AddDays($i) }
        })

        $selected = Select-PodcastEpisode -Episodes $episodes -Mode $Mode -CustomCount $CustomCount

        ($selected -is [array]) | Should -BeTrue
        $selected.Count | Should -Be $ExpectedCount
        for ($i = 0; $i -lt $ExpectedCount; $i++) {
            $selected[$i].Title | Should -Be ("Episode " + ($InputCount - $i - 1))
        }
        $episodes.Count | Should -Be $InputCount
        if ($InputCount -gt 0) { $episodes[0].Title | Should -Be 'Episode 0' }
        Should -Invoke Invoke-PodcastMetadataRequest -Times 0 -Exactly
    }

    It 'treats null input as an empty array' {
        $selected = Select-PodcastEpisode -Episodes $null -Mode All
        ($selected -is [array]) | Should -BeTrue
        $selected.Count | Should -Be 0
    }

    It 'rejects invalid Custom count <Count> before selection' -ForEach @(@{ Count = 0 }, @{ Count = -1 }) {
        { Select-PodcastEpisode -Episodes @() -Mode Custom -CustomCount $Count } |
            Should -Throw "*Mode 'Custom' requires -CustomCount*"
    }
}

Describe 'A006: entrypoint counts and progress' -Tag 'Unit', 'A006' {
    BeforeEach {
        $fixtureMedia = $script:FixtureMedia
        Mock Write-Host {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        $feedResponse = [pscustomobject]@{ Content = '' }
        Mock Invoke-PodcastMetadataRequest { $feedResponse }
        Mock Invoke-PodcastMediaRequest {
            $DestinationStream.Write($fixtureMedia, 0, $fixtureMedia.Length)
            [pscustomobject]@{ Completed = $true; Bytes = $fixtureMedia.Length; ContentLength = $fixtureMedia.Length; ContentType = 'audio/mpeg' }
        }
    }

    It 'counts and progresses correctly for <InputCount> items in <Mode> with count <CustomCount>' -ForEach @(
        foreach ($inputCount in 1, 3) {
            foreach ($mode in 'Latest', 'All', 'Custom') {
                foreach ($count in $(if ($mode -eq 'Custom') { 1, 2, 5 } else { 0 })) {
                    @{
                        InputCount = $inputCount
                        Mode = $mode
                        CustomCount = $count
                        ExpectedCount = if ($mode -eq 'Latest') { 1 }
                            elseif ($mode -eq 'All') { $inputCount }
                            else { [Math]::Min($count, $inputCount) }
                    }
                }
            }
        }
    ) {
        $items = for ($i = 0; $i -lt $InputCount; $i++) {
            '<item><title>Episode {0}</title><pubDate>2026-09-{1:00}T12:00:00Z</pubDate><enclosure url="https://media.example.invalid/{0}.mp3" /></item>' -f $i, ($i + 1)
        }
        # A non-downloadable item also exercises filtering down to a singleton.
        $feedResponse.Content = '<rss><channel><title>Unit show</title><item><title>No media</title></item>' + ($items -join '') + '</channel></rss>'
        $output = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $arguments = @{ FeedUrl = 'https://feed.example.invalid/rss'; OutputPath = $output; Mode = $Mode; CustomCount = $CustomCount }

        & $script:DownloaderPath @arguments

        @(Get-ChildItem -LiteralPath $output -Filter '*.mp3' -Recurse).Count | Should -Be $ExpectedCount
        $log = Get-Content -LiteralPath (Get-ChildItem -LiteralPath $output -Filter '*.log' -Recurse).FullName -Raw
        $log | Should -Match "Feed items with valid URLs: $InputCount"
        $log | Should -Match "Episodes to download \(after mode/filter\): $ExpectedCount"
        $log | Should -Match "Summary: Downloaded=$ExpectedCount, Skipped=0, Failed=0"
        Should -Invoke Invoke-PodcastMediaRequest -Times $ExpectedCount -Exactly -ParameterFilter { $DestinationStream }
        Should -Invoke Write-Progress -Times $ExpectedCount -Exactly -ParameterFilter { -not $Completed }
        for ($i = 1; $i -le $ExpectedCount; $i++) {
            $operation = "Episode $i of $ExpectedCount"
            $percent = [int](($i / $ExpectedCount) * 100)
            Should -Invoke Write-Progress -Times 1 -Exactly -ParameterFilter {
                -not $Completed -and $CurrentOperation -eq $operation -and $PercentComplete -eq $percent
            }
        }

        & $script:DownloaderPath @arguments

        Should -Invoke Invoke-PodcastMediaRequest -Times $ExpectedCount -Exactly
        Should -Invoke Write-Progress -Times $ExpectedCount -Exactly -ParameterFilter { $Status -like 'Skipping (verified history):*' }
        Should -Invoke Write-Progress -Times 0 -Exactly -ParameterFilter { -not $Completed -and ($PercentComplete -lt 0 -or $PercentComplete -gt 100) }
    }

    It 'rejects <Kind> input explicitly in <Mode> without media requests or episode progress' -ForEach @(
        foreach ($mode in 'Latest', 'Custom', 'All') {
            @{ Mode = $mode; Kind = 'empty'; Items = ''; Message = '*No episodes found in the feed*' }
            @{ Mode = $mode; Kind = 'no enclosure'; Items = '<item><title>No media</title></item>'; Message = '*no downloadable enclosure URLs*' }
        }
    ) {
        $feedResponse.Content = '<rss><channel><title>Empty show</title>' + $Items + '</channel></rss>'
        $output = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        { & $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/rss' -OutputPath $output -Mode $Mode -CustomCount 1 } |
            Should -Throw $Message
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
        Should -Invoke Write-Progress -Times 0 -Exactly
        Test-Path -LiteralPath $output | Should -BeFalse
    }
}
