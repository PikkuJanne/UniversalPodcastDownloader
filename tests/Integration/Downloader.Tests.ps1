BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python 3.9+ is required for loopback integration tests. It is a development dependency only.' }
    $script:pythonPath = $python.Source
    $script:legacyEngine = $PSVersionTable.PSVersion.Major -eq 5
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

    It 'resolves RSS and Atom over loopback (PS5.1 uses an explicit harness BasicParsing default)' {
        $rss = Invoke-UpdIntegrationWorker -Context $context -Action Resolve -FeedPath '/feeds/single.xml' -BasicParsing:$script:legacyEngine
        $rss.Result.Succeeded | Should -BeTrue -Because ($rss.Stdout + $rss.Stderr + $rss.Result.ErrorMessage)
        $rss.ExitCode | Should -Be 0
        $rss.Result.ItemCount | Should -Be 1
        $rss.Result.ResolvedUrl | Should -Be ($context.BaseUrl + '/feeds/single.xml')
        $rss.Result.EpisodeTitles[0] | Should -Be 'One synthetic episode'
        $rss.Result.EpisodeUrls[0] | Should -Be ($context.BaseUrl + '/media/ok.mp3')

        $atom = Invoke-UpdIntegrationWorker -Context $context -Action Resolve -FeedPath '/feeds/atom.xml' -BasicParsing:$script:legacyEngine
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

    It 'records the untouched single-episode entrypoint baseline: PS7 download/repeat, PS5.1 failure' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/single.xml'
        if ($script:legacyEngine) {
            # Patched PS5.1 refuses its legacy web confirmation under NonInteractive.
            # Older PS5.1 can get farther and expose the scalar Count defect instead.
            $first.Result.Succeeded | Should -BeFalse
            $first.ExitCode | Should -Be 1
            $first.Result.ErrorMessage | Should -Match 'No episodes found|divide by zero'
            @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -Filter '*.mp3').Count | Should -Be 0
            (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
            Write-Host ('Observed unmodified PS5.1 baseline failure: ' + $first.Result.ErrorMessage)
            Write-Host $first.Stdout
        }
        else {
            $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
            $first.ExitCode | Should -Be 0
            $files = @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -Filter '*.mp3')
            $files.Count | Should -Be 1
            $files[0].Name | Should -Be '2026-09-01 - One synthetic episode.mp3'
            $files[0].Length | Should -Be $script:expectedLength
            (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
            $stamp = $files[0].LastWriteTimeUtc
            $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/single.xml'
            $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
            $second.Stdout | Should -Match 'Skipping \(already exists\)'
            (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
            (Get-Item -LiteralPath $files[0].FullName).LastWriteTimeUtc | Should -Be $stamp
            (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
        }
    }

    It 'isolates the PS5.1 singleton Count defect with an explicit harness BasicParsing default' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/single.xml' -BasicParsing
        if ($script:legacyEngine) {
            $run.Result.Succeeded | Should -BeFalse
            $run.ExitCode | Should -Be 1
            $run.Result.ErrorMessage | Should -Match 'divide by zero'
            @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Filter '*.mp3').Count | Should -Be 0
            (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
            Write-Host ('Observed PS5.1 singleton defect after harness-only BasicParsing: ' + $run.Result.ErrorMessage)
        }
        else {
            $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
            $run.ExitCode | Should -Be 0
            $file = @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Filter '*.mp3')
            $file.Count | Should -Be 1
            (Get-FileHash -LiteralPath $file[0].FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
            (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
        }
    }

    It 'downloads two real synthetic media files and preserves them on repeat (PS5.1 harness BasicParsing)' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/atom.xml' -BasicParsing:$script:legacyEngine
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $first.ExitCode | Should -Be 0
        $files = @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -Filter '*.mp3' | Sort-Object Name)
        $files.Count | Should -Be 2
        $stamps = @{}
        foreach ($file in $files) {
            $file.Length | Should -Be $script:expectedLength
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be $script:expectedHash
            $stamps[$file.FullName] = $file.LastWriteTimeUtc
        }
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/atom.xml' -BasicParsing:$script:legacyEngine
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
