BeforeAll {
    $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    . $script:DownloaderPath
    Mock Get-PodcastHttpClient { throw 'Unit tests must not create a network client.' }
    Mock Invoke-PodcastMetadataRequest { throw 'Unexpected metadata request in unit test.' }
    $script:diagnosticMedia = [IO.File]::ReadAllBytes((Join-Path (Split-Path $script:DownloaderPath -Parent) 'tools/codex-handoff/fixtures/silence.mp3'))
}

Describe 'A020: entrypoint diagnostic privacy' -Tag 'Unit', 'A020' {
    It 'A025 renders only fixed transport categories despite a private error message' -ForEach @(
        @{ Kind = 'Deferred'; Expected = '*deferred*' },
        @{ Kind = 'HeaderTimeout'; Expected = '*connection/header timeout*' },
        @{ Kind = 'IdleTimeout'; Expected = '*idle transfer timeout*' },
        @{ Kind = 'Connection'; Expected = '*network connection failed*' },
        @{ Kind = 'HttpStatus'; Expected = '*unsuccessful HTTP status*' },
        @{ Kind = 'IncompleteBody'; Expected = '*body was incomplete*' },
        @{ Kind = 'Permanent'; Expected = '*network policy*' }
    ) {
        $errorValue = New-PodcastTransportException -Kind $Kind -Message 'privateTransportCanary https://feed.invalid/token'
        $publicMessage = Get-PodcastDiagnosticError -Error $errorValue
        $publicMessage | Should -BeLike $Expected
        $publicMessage | Should -Not -Match 'privateTransportCanary|feed.invalid|token'
    }

    BeforeEach {
        Mock Read-Host { throw 'Unit tests must not prompt.' }
        Mock Write-Progress {}
        Mock Start-Sleep {}
        $media = $script:diagnosticMedia
        Mock Invoke-PodcastMetadataRequest {
            [pscustomobject]@{ Content = '<rss><channel><title>privateTitleCanary</title><item><title>privateEpisodeCanary</title><guid>privateGuidCanary</guid><enclosure url="https://media.example.invalid/privatePathCanary?credential=privateQueryCanary" /></item></channel></rss>' }
        }
        Mock Invoke-PodcastMediaRequest {
            $DestinationStream.Write($media, 0, $media.Length)
            [pscustomobject]@{ Completed = $true; Bytes = $media.Length; ContentLength = $media.Length; ContentType = 'audio/mpeg' }
        }
    }

    It 'keeps secrets out of successful console verbose logs and the restricted export' {
        $root = Join-Path $TestDrive 'privateDirectoryCanary'
        $export = Join-Path $TestDrive 'diagnostics.json'
        $output = @(& $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/privateFeedCanary?token=privateFeedQueryCanary' -Mode All -OutputPath $root -DiagnosticExportPath $export -Verbose *>&1)
        $logs = @(Get-ChildItem -LiteralPath $root -Filter '*.log' -Recurse | ForEach-Object { [IO.File]::ReadAllText($_.FullName) })
        $publicText = ($output | Out-String) + ($logs -join "`n") + [IO.File]::ReadAllText($export)
        $publicText | Should -Not -Match 'private(?:Title|Episode|Guid|Path|Query|Directory|Feed|FeedQuery)Canary'
        $publicText | Should -Match 'feed.example.invalid'
        $publicText | Should -Match 'media.example.invalid'
        Should -Invoke Invoke-PodcastMetadataRequest -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://feed.example.invalid/privateFeedCanary?token=privateFeedQueryCanary'
        }
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://media.example.invalid/privatePathCanary?credential=privateQueryCanary' -or $DestinationStream
        }
        @(Get-ChildItem -LiteralPath $root -Filter '*.mp3' -Recurse).Count | Should -Be 1
    }

    It 'returns a WhatIf projection without internal episode or history objects and creates no export' {
        $root = Join-Path $TestDrive 'absent'
        $export = Join-Path $TestDrive 'must-not-exist.json'
        $rows = @(& $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/privateFeedCanary' -Mode All -OutputPath $root -DiagnosticExportPath $export -WhatIf -Verbose *>&1)
        $plan = @($rows | Where-Object { $_.PSObject.Properties['EpisodeId'] })
        $serialized = ($rows | Out-String) + ($plan | ConvertTo-Json -Depth 10)
        $serialized | Should -Not -Match 'private(?:Title|Episode|Guid|Path|Query|Feed)Canary'
        $plan.Count | Should -Be 1
        $plan[0].PSObject.Properties.Name | Should -Not -Contain 'Episode'
        $plan[0].PSObject.Properties.Name | Should -Not -Contain 'StateRecord'
        $plan[0].Source | Should -Match '^media.example.invalid \[url [a-f0-9]{32}\]$'
        Test-Path -LiteralPath $root | Should -BeFalse
        Test-Path -LiteralPath $export | Should -BeFalse
        Should -Invoke Invoke-PodcastMediaRequest -Times 0 -Exactly
    }

    It 'does not replace a successful download when optional export fails' {
        $root = Join-Path $TestDrive 'export-failure-output'
        $existing = Join-Path $TestDrive 'existing.json'
        [IO.File]::WriteAllText($existing, 'Preserved original export target')
        { & $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/rss' -Mode All -OutputPath $root -DiagnosticExportPath $existing } | Should -Not -Throw
        [IO.File]::ReadAllText($existing) | Should -Be 'Preserved original export target'
        @(Get-ChildItem -LiteralPath $root -Filter '*.mp3' -Recurse).Count | Should -Be 1
    }

    It 'preserves the primary exception when optional export fails' {
        Mock Invoke-PodcastMetadataRequest { throw [InvalidOperationException]::new('privateErrorCanary https://example.invalid/privateExceptionPath') }
        $root = Join-Path $TestDrive 'failure-output'
        $caught = $null
        try {
            & $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/privateFeedCanary' -Mode All -OutputPath $root -DiagnosticExportPath (Join-Path $TestDrive 'absent-parent/export.json')
        }
        catch { $caught = $_ }
        $caught | Should -Not -BeNullOrEmpty
        $caught.Exception.Message | Should -Match '^No episodes found in the feed\.'
        $caught.ErrorDetails.Message | Should -Not -Match 'privateErrorCanary|privateExceptionPath|privateFeedCanary'
        Test-Path -LiteralPath $root | Should -BeFalse
    }

    It 'keeps the original exception instance while formatting a private failure safely' {
        $primary = [InvalidOperationException]::new('privateErrorCanary https://example.invalid/privateExceptionPath Authorization: privateHeaderCanary')
        Mock Get-EpisodeData { throw $primary }
        $caught = $null
        try {
            & $script:DownloaderPath -FeedUrl 'https://feed.example.invalid/rss' -Mode All -OutputPath (Join-Path $TestDrive 'primary')
        }
        catch { $caught = $_ }
        $caught | Should -Not -BeNullOrEmpty
        [object]::ReferenceEquals($caught.Exception, $primary) | Should -BeTrue
        $caught.Exception.Message | Should -Match 'privateErrorCanary'
        ($caught | Out-String) | Should -Not -Match 'private(?:Error|ExceptionPath|Header)Canary'
        $caught.ErrorDetails.Message | Should -Match 'operation failed'
    }

    It 'retains an exact safe review instruction while omitting injected suffixes' {
        $message = 'Legacy archive requires review; use -LegacyPath with -LegacyAction Preview.'
        Get-PodcastDiagnosticError -Error ([InvalidOperationException]::new($message)) | Should -Be $message
        Get-PodcastDiagnosticError -Error ([InvalidOperationException]::new($message + ' privateErrorCanary')) |
            Should -Not -Match 'privateErrorCanary|LegacyPath'
    }
}
