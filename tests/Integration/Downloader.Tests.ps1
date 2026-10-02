BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python 3.9+ is required for loopback integration tests. It is a development dependency only.' }
    $script:pythonPath = $python.Source
    $expectedMedia = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
    $script:expectedHash = (Get-FileHash -LiteralPath $expectedMedia -Algorithm SHA256).Hash
    $script:expectedLength = (Get-Item -LiteralPath $expectedMedia).Length
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
        $files[0].Name | Should -Be '2026-09-01 - One synthetic episode.mp3'
        $files[0].Length | Should -Be $script:expectedLength
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
        $stamp = $files[0].LastWriteTimeUtc
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/single.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.ExitCode | Should -Be 0
        $second.Stdout | Should -Match 'Skipping \(already exists\)'
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
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Filter '*.mp3').Count | Should -Be 0
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
}
