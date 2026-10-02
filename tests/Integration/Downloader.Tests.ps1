BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    . (Join-Path $repositoryRoot 'UniversalPodcastDownloader.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python 3.9+ is required for loopback integration tests. It is a development dependency only.' }
    $script:pythonPath = $python.Source
    $expectedMedia = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
    $script:expectedHash = (Get-FileHash -LiteralPath $expectedMedia -Algorithm SHA256).Hash
    $script:expectedLength = (Get-Item -LiteralPath $expectedMedia).Length

    function Get-UpdSyntheticDestination {
        param($Context)
        $output = Join-Path $Context.Root 'output'
        $folderBudget = [Math]::Min(100, 259 - $output.TrimEnd('\').Length - 2 - 84)
        $folderName = New-PodcastFolderName -FeedTitle 'Fixture Podcast' -FeedUrl ($Context.BaseUrl + '/feeds/single.xml') -MaxLength $folderBudget
        $folder = Join-Path $output $folderName
        $episode = [pscustomobject]@{
            Title = 'One synthetic episode'
            Guid = 'fixture-001'
            Url = $Context.BaseUrl + '/media/ok.mp3'
            PubDate = [datetime]'2026-09-01T12:00:00'
        }
        $feedId = Get-PodcastNameHash -IdentityKey ('feed:' + $Context.BaseUrl + '/feeds/single.xml')
        $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $feedId
        $fileName = New-EpisodeFileName -Episode $episode -IdentityHash $identity.Id -MaxLength ([Math]::Min(180, 259 - $folder.Length - 1))
        [pscustomobject]@{ Output = $output; Folder = $folder; File = Join-Path $folder $fileName }
    }
}

Describe 'Real downloader against synthetic loopback fixtures' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:pythonPath
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'A007 discovers a feed from a real HTML response without a legacy parsing prompt' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Discover -FeedPath '/show'
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.ExitCode | Should -Be 0
        $run.Result.PromptCount | Should -Be 2
        # The current discovery parser selects this fixture's Atom candidate;
        # candidate selection and HTML parsing improvements belong to UPD-0203.
        $run.Result.ResolvedUrl | Should -Be ($context.BaseUrl + '/feeds/atom.xml')
        $run.Stdout | Should -Not -Match 'Security Warning|Script Execution Risk|UseBasicParsing'
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/show' | Should -Be 1
        $stats.'/feeds/atom.xml' | Should -BeNullOrEmpty
        $stats.'/media/ok.mp3' | Should -BeNullOrEmpty
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
    }

    It 'A007 resolves RSS and Atom over loopback without a harness parsing override' {
        $rss = Invoke-UpdIntegrationWorker -Context $context -Action Resolve -FeedPath '/feeds/single.xml'
        $rss.Result.Succeeded | Should -BeTrue -Because ($rss.Stdout + $rss.Stderr + $rss.Result.ErrorMessage)
        $rss.ExitCode | Should -Be 0
        $rss.Result.ItemCount | Should -Be 1
        $rss.Result.ResolvedUrl | Should -Be ($context.BaseUrl + '/feeds/single.xml')
        $rss.Result.EpisodeTitles[0] | Should -Be 'One synthetic episode'
        $rss.Result.EpisodeUrls[0] | Should -Be ($context.BaseUrl + '/media/ok.mp3')

        $atom = Invoke-UpdIntegrationWorker -Context $context -Action Resolve -FeedPath '/feeds/atom.xml'
        $atom.Result.Succeeded | Should -BeTrue -Because ($atom.Stdout + $atom.Stderr + $atom.Result.ErrorMessage)
        $atom.ExitCode | Should -Be 0
        $atom.Result.ItemCount | Should -Be 2
        @($atom.Result.EpisodeUrls).Count | Should -Be 2
        $atom.Result.EpisodeUrls[0] | Should -Be ($context.BaseUrl + '/media/ok.mp3?id=old')
        $atom.Result.EpisodeUrls[1] | Should -Be ($context.BaseUrl + '/media/ok.mp3?id=new')
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/single.xml' | Should -Be 1
        $stats.'/feeds/atom.xml' | Should -Be 1
        $stats.'/media/ok.mp3' | Should -BeNullOrEmpty
        Test-Path -LiteralPath $rss.OutputPath | Should -BeFalse
    }

    It 'A007 downloads a single episode and preserves it on repeat without a legacy parsing prompt' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/single.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $first.ExitCode | Should -Be 0
        $first.Stdout | Should -Match 'Downloaded\s+: 1'
        $first.Stdout | Should -Not -Match 'Security Warning|Script Execution Risk|UseBasicParsing|divide by zero'
        $files = @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -Filter '*.mp3')
        $files.Count | Should -Be 1
        $files[0].FullName | Should -Be (Get-UpdSyntheticDestination -Context $context).File
        $files[0].Length | Should -Be $script:expectedLength
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
        $stamp = $files[0].LastWriteTimeUtc
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/single.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.ExitCode | Should -Be 0
        $second.Stdout | Should -Match 'Skipping \(verified history\)'
        $second.Stdout | Should -Match 'Skipped\s+: 1'
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
        (Get-Item -LiteralPath $files[0].FullName).LastWriteTimeUtc | Should -Be $stamp
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
    }

    It 'A006 downloads <ExpectedCount> episode(s) from <FeedPath> in <Mode> mode with CustomCount=<CustomCount>' -TestCases @(
        @{ FeedPath = '/feeds/single.xml'; Mode = 'Latest'; CustomCount = 1; ExpectedCount = 1 }
        @{ FeedPath = '/feeds/single.xml'; Mode = 'Custom'; CustomCount = 1; ExpectedCount = 1 }
        @{ FeedPath = '/feeds/single.xml'; Mode = 'Custom'; CustomCount = 5; ExpectedCount = 1 }
        @{ FeedPath = '/feeds/atom.xml'; Mode = 'Latest'; CustomCount = 1; ExpectedCount = 1 }
        @{ FeedPath = '/feeds/atom.xml'; Mode = 'Custom'; CustomCount = 1; ExpectedCount = 1 }
        @{ FeedPath = '/feeds/atom.xml'; Mode = 'Custom'; CustomCount = 2; ExpectedCount = 2 }
        @{ FeedPath = '/feeds/atom.xml'; Mode = 'Custom'; CustomCount = 5; ExpectedCount = 2 }
    ) {
        param($FeedPath, $Mode, $CustomCount, $ExpectedCount)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath -Mode $Mode -CustomCount $CustomCount
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.ExitCode | Should -Be 0
        $run.Stdout | Should -Match ('Downloaded\s+: ' + $ExpectedCount)
        $run.Stdout | Should -Match 'Failed\s+: 0'
        $files = @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Filter '*.mp3')
        $files.Count | Should -Be $ExpectedCount
        foreach ($file in $files) {
            $file.Length | Should -Be $script:expectedLength
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
        }
        $stats = Get-UpdFixtureState -Context $context
        $stats.$FeedPath | Should -Be 1
        $stats.'/media/ok.mp3' | Should -Be $ExpectedCount
    }

    It 'A006 rejects an empty feed explicitly in <Mode> mode without requesting media' -TestCases @(
        @{ Mode = 'Latest' }
        @{ Mode = 'Custom' }
        @{ Mode = 'All' }
    ) {
        param($Mode)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/empty.xml' -Mode $Mode
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match '^No episodes found in the feed\.'
        $run.Result.ErrorMessage | Should -Not -Match 'divide by zero|null'
        $run.Stdout | Should -Not -Match 'Feed failed:'
        $run.Stdout | Should -Not -Match 'Security Warning|Script Execution Risk|UseBasicParsing|divide by zero'
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/empty.xml' | Should -Be 1
        $stats.'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'A006 downloads two synthetic media files in All mode and preserves them on repeat' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/atom.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $first.ExitCode | Should -Be 0
        $first.Stdout | Should -Match 'Downloaded\s+: 2'
        $files = @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -Filter '*.mp3' | Sort-Object Name)
        $files.Count | Should -Be 2
        $stamps = @{}
        foreach ($file in $files) {
            $file.Length | Should -Be $script:expectedLength
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
            $stamps[$file.FullName] = $file.LastWriteTimeUtc
        }
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/atom.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.ExitCode | Should -Be 0
        $second.Stdout | Should -Match 'Skipped\s+: 2'
        foreach ($file in $files) {
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
            (Get-Item -LiteralPath $file.FullName).LastWriteTimeUtc | Should -Be $stamps[$file.FullName]
        }
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/atom.xml' | Should -Be 2
        $stats.'/media/ok.mp3' | Should -Be 2
    }

    It 'A009 rejects hostile feed title <FeedPath> before log or media writes' -TestCases @(
        @{ FeedPath = '/feeds/hostile-dot.xml' }
        @{ FeedPath = '/feeds/hostile-dotdot.xml' }
        @{ FeedPath = '/feeds/hostile-drive.xml' }
        @{ FeedPath = '/feeds/hostile-unc.xml' }
    ) {
        param($FeedPath)
        $sentinel = Join-Path $context.Root 'original.mp3'
        [IO.File]::WriteAllText($sentinel, 'Synthetic original outside the selected output root.')
        $originalHash = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match 'path|component|dot|absolute|root'
        @(Get-ChildItem -LiteralPath $context.Root -Recurse -Filter '*.log').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $context.Root -Recurse -Filter '*.mp3').Count | Should -Be 1
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -Be $originalHash
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'A009 refuses an existing junction at <Location> and preserves the synthetic target' -TestCases @(
        @{ Location = 'output-root' }
        @{ Location = 'output-ancestor' }
        @{ Location = 'podcast-folder' }
        @{ Location = 'media-destination' }
    ) {
        param($Location)
        $archive = Join-Path $context.Root 'synthetic-originals'
        $null = New-Item -ItemType Directory -Path $archive
        $sentinel = Join-Path $archive 'original.mp3'
        [IO.File]::WriteAllText($sentinel, 'Synthetic archive bytes must survive a rejected junction.')
        $originalHash = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        $originalStamp = (Get-Item -LiteralPath $sentinel).LastWriteTimeUtc
        $paths = Get-UpdSyntheticDestination -Context $context
        $outputName = 'output'
        switch ($Location) {
            'output-root' { $junction = $paths.Output }
            'output-ancestor' {
                $junction = Join-Path $context.Root 'ancestor'
                $outputName = 'ancestor/output'
            }
            'podcast-folder' {
                $null = New-Item -ItemType Directory -Path $paths.Output
                $junction = $paths.Folder
            }
            'media-destination' {
                $null = New-Item -ItemType Directory -Path $paths.Folder -Force
                $junction = $paths.File
            }
        }
        New-UpdOwnedJunction -Context $context -Path $junction -Target $archive
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/single.xml' -OutputName $outputName
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match 'reparse|junction'
        @(Get-ChildItem -LiteralPath $archive -Force).Count | Should -Be 1
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -Be $originalHash
        (Get-Item -LiteralPath $sentinel).LastWriteTimeUtc | Should -Be $originalStamp
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'A009 rechecks a destination junction inserted at <Stage>' -TestCases @(
        @{ Stage = 'Preparing'; ExpectedRequests = 0 }
        @{ Stage = 'AfterTransfer'; ExpectedRequests = 1 }
    ) {
        param($Stage, $ExpectedRequests)
        $archive = Join-Path $context.Root 'synthetic-originals'
        $null = New-Item -ItemType Directory -Path $archive
        $sentinel = Join-Path $archive 'original.mp3'
        [IO.File]::WriteAllText($sentinel, 'Synthetic archive preserved at the later write boundary.')
        $originalHash = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        $paths = Get-UpdSyntheticDestination -Context $context
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/single.xml' `
            -BoundaryJunctionPath $paths.File -BoundaryJunctionTarget $archive -BoundaryStage $Stage
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Result.BoundaryInjectionCount | Should -Be 1
        $run.Result.ErrorMessage | Should -Match 'reparse|junction'
        @(Get-ChildItem -LiteralPath $archive -Force).Count | Should -Be 1
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -Be $originalHash
        @(Get-ChildItem -LiteralPath $paths.Folder -Filter '*.tmp' -Force).Count | Should -Be 0
        [int](Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be $ExpectedRequests
    }

    It 'A009 downloads and preserves <Count> distinct files from <FeedPath>' -TestCases @(
        @{ FeedPath = '/feeds/collisions.xml'; Count = 6 }
        @{ FeedPath = '/feeds/long-names.xml'; Count = 2 }
    ) {
        param($FeedPath, $Count)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $first.ExitCode | Should -Be 0
        $first.Stdout | Should -Match ('Downloaded\s+: ' + $Count)
        $files = @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be $Count
        @($files.Name | Select-Object -Unique).Count | Should -Be $Count
        $stamps = @{}
        foreach ($file in $files) {
            $file.Name | Should -Match '-[a-f0-9]{64}\.mp3$'
            $file.FullName.Length | Should -BeLessOrEqual 259
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
            $stamps[$file.FullName] = $file.LastWriteTimeUtc
        }
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.ExitCode | Should -Be 0
        $second.Stdout | Should -Match ('Skipped\s+: ' + $Count)
        foreach ($file in $files) {
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
            (Get-Item -LiteralPath $file.FullName).LastWriteTimeUtc | Should -Be $stamps[$file.FullName]
        }
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be $Count
    }

    It 'A009 A017 stops for review when a title matches legacy media and preserves that folder' {
        $legacy = Join-Path (Join-Path $context.Root 'output') 'Fixture Podcast'
        $null = New-Item -ItemType Directory -Path $legacy -Force
        $sentinel = Join-Path $legacy '2026-09-01 - One synthetic episode.mp3'
        [IO.File]::WriteAllText($sentinel, 'Synthetic legacy media remains untouched.')
        $originalHash = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        $stamp = (Get-Item -LiteralPath $sentinel).LastWriteTimeUtc
        foreach ($feed in '/feeds/single.xml', '/feeds/single-alias.xml') {
            $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $feed
            $run.Result.Succeeded | Should -BeFalse
            $run.Stdout | Should -Match 'review'
        }
        $folders = @(Get-ChildItem -LiteralPath (Join-Path $context.Root 'output') -Directory)
        $folders.Count | Should -Be 1
        @($folders | Where-Object { $_.Name -match '-[a-f0-9]{64}$' }).Count | Should -Be 0
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -Be $originalHash
        (Get-Item -LiteralPath $sentinel).LastWriteTimeUtc | Should -Be $stamp
        @((Get-UpdFixtureState -Context $context).PSObject.Properties | Where-Object { $_.Name -like '/media/*' }).Count | Should -Be 0
    }
}
