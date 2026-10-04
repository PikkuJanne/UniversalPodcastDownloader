BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required only for the owned documentation loopback fixture.' }
    $script:documentationPython = $python.Source

    function Invoke-UpdDocumentationObservation {
        param($Context, [ValidateSet('Help', 'Readme', 'HelpExamples', 'Exit', 'Guided')][string]$Action,
            [string]$CommandName = 'UniversalPodcastDownloader.ps1', [int]$ExpectedExit = 0)
        $id = [guid]::NewGuid().ToString('N')
        $resultPath = Join-Path $Context.Root ($id + '-result.json')
        $configPath = Join-Path $Context.Root ($id + '-config.json')
        $config = @{ Root = $Context.Root; Token = $Context.Token; BaseUrl = $Context.BaseUrl; Action = $Action;
            ProductScript = Join-Path $package 'UniversalPodcastDownloader.ps1'; ResultPath = $resultPath;
            CommandName = $CommandName; ExpectedExit = $ExpectedExit;
            Readme = [IO.File]::ReadAllText((Join-Path $repositoryRoot 'README.md'));
            Sample = Join-Path $Context.Root 'original-sample.mp3' }
        [IO.File]::Copy((Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'), $config.Sample)
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $worker = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engine) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
            (Join-Path $repositoryRoot 'tests/support/Invoke-DocumentationWorker.ps1'), '-ConfigPath', $configPath)
        if (-not $worker.Process.WaitForExit(60000)) { throw 'Owned documentation worker exceeded its bounded wait.' }
        if (-not [IO.File]::Exists($resultPath)) { throw ('Documentation worker returned no result: ' + $worker.ErrorOutput.Result) }
        [pscustomobject]@{ ExitCode = $worker.Process.ExitCode; Stdout = $worker.Output.Result; Stderr = $worker.ErrorOutput.Result;
            Result = [IO.File]::ReadAllText($resultPath) | ConvertFrom-Json }
    }
}

# A049/A050 remain manual acceptance cases. These deterministic observations
# support the separately recorded manual review; they do not rename that layer.
Describe 'A049/A050 verified user documentation in a fresh copied Windows runtime' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        $package = Join-Path $context.Root ('Extracted application space ' + [char]0x00e4)
        $null = [IO.Directory]::CreateDirectory((Join-Path $package 'src'))
        foreach ($name in @('UniversalPodcastDownloader.ps1', 'UniversalPodcastDownloader.bat')) {
            [IO.File]::Copy((Join-Path $repositoryRoot $name), (Join-Path $package $name))
        }
        foreach ($source in Get-ChildItem -LiteralPath (Join-Path $repositoryRoot 'src') -File -Filter '*.ps1') {
            [IO.File]::Copy($source.FullName, (Join-Path $package ('src/' + $source.Name)))
        }
        Start-UpdFixtureServer -Context $context -PythonPath $script:documentationPython
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'renders authored standard help for <CommandName> and every actual public parameter' -TestCases @(
        @{ CommandName = 'UniversalPodcastDownloader.ps1' }
        @{ CommandName = 'Invoke-PodcastRun' }
    ) {
        param($CommandName)
        $run = Invoke-UpdDocumentationObservation -Context $context -Action Help -CommandName $CommandName
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.Help.Authored | Should -BeTrue
        @($run.Result.Help.MissingParameters).Count | Should -Be 0
        $run.Result.Help.ExampleCount | Should -BeGreaterThan 0
        $run.Result.PackageChanged | Should -BeFalse
        $run.Result.RuntimePythonAvailable | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/feeds/history.xml' | Should -BeNullOrEmpty
    }

    It 'executes every literal README command with synthetic feeds and preserves copied recovery originals' {
        $run = Invoke-UpdDocumentationObservation -Context $context -Action Readme
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.Blocks.Id | Should -Be @('help', 'setup', 'preview', 'launcher', 'custom', 'saved', 'api',
            'legacy-preview', 'legacy-adopt', 'legacy-redownload', 'legacy-rollback')
        @($run.Result.Blocks | Where-Object { -not $_.Succeeded }).Count | Should -Be 0
        $run.Result.OriginalsPreserved | Should -BeTrue
        $run.Result.ObservedVerifiedBytes | Should -BeTrue
        $run.Result.AdoptionIsUnverified | Should -BeTrue
        $run.Result.PackageChanged | Should -BeFalse
        $run.Result.DefaultConfigCreated | Should -BeFalse
        $run.Result.RuntimePythonAvailable | Should -BeFalse
        $run.Result.PreviewChangedTree | Should -BeFalse
        $run.Result.CallableResult.Type | Should -Be 'Podcast.RunResult'
        $run.Result.CallableResult.ExitCode | Should -Be 0
        $run.Result.RecoveryOutcomes | Should -Be @('adopted', 'transfer_verified', 'metadata_restored')
    }

    It 'executes all rendered script and callable help examples literally from the copied runtime' {
        $run = Invoke-UpdDocumentationObservation -Context $context -Action HelpExamples
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error + ($run.Result.Blocks | ConvertTo-Json))
        @($run.Result.Blocks).Count | Should -BeGreaterThan 3
        @($run.Result.Blocks | Where-Object { -not $_.Succeeded }).Count | Should -Be 0
        $run.Result.RenderedCodeMatchesAuthored | Should -BeTrue
        $run.Result.PassThruResult.Type | Should -Be 'Podcast.RunResult'
        $run.Result.CallableResult.Type | Should -Be 'Podcast.RunResult'
        $run.Result.PackageChanged | Should -BeFalse
        $run.Result.DefaultConfigCreated | Should -BeFalse
    }

    It 'returns documented process exit <ExpectedExit> from the actual copied entry script' -TestCases @(
        @{ ExpectedExit = 0 }
        @{ ExpectedExit = 1 }
        @{ ExpectedExit = 2 }
    ) {
        param($ExpectedExit)
        $run = Invoke-UpdDocumentationObservation -Context $context -Action Exit -ExpectedExit $ExpectedExit
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.ChildExitCode | Should -Be $ExpectedExit
        $run.Result.PackageChanged | Should -BeFalse
        $run.Result.DefaultConfigCreated | Should -BeFalse
    }

    It 'runs the actual argument-free batch guided launcher against only an owned synthetic profile' {
        $run = Invoke-UpdDocumentationObservation -Context $context -Action Guided
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.ChildExitCode | Should -Be 0
        $run.Result.GuidedPromptsObserved | Should -BeTrue
        $run.Result.ObservedVerifiedBytes | Should -BeTrue
        $run.Result.PackageChanged | Should -BeFalse
        $run.Result.DefaultConfigCreated | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }
}
