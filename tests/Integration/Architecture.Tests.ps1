BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required only for the owned architecture loopback fixture.' }
    $script:architecturePython = $python.Source
    function Invoke-UpdArchitectureObservation {
        param($Context, [ValidateSet('Import', 'Portable', 'Isolation')][string]$Action)
        $id = [guid]::NewGuid().ToString('N')
        $configPath = Join-Path $Context.Root ($id + '-config.json')
        $resultPath = Join-Path $Context.Root ($id + '-result.json')
        $config = @{ Root = $Context.Root; BaseUrl = $Context.BaseUrl; Action = $Action;
            ProductScript = Join-Path $package 'UniversalPodcastDownloader.ps1';
            OutputPath = Join-Path $Context.Root 'output'; ResultPath = $resultPath }
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $child = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engine) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
            (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-ArchitectureWorker.ps1'), '-ConfigPath', $configPath
        )
        if (-not $child.Process.WaitForExit(30000)) { throw 'Owned architecture worker exceeded its bounded wait.' }
        [pscustomobject]@{ ExitCode = $child.Process.ExitCode; Stdout = $child.Output.Result; Stderr = $child.ErrorOutput.Result;
            Result = ([IO.File]::ReadAllText($resultPath) | ConvertFrom-Json); OutputPath = $config.OutputPath }
    }
}

Describe 'A048 runtime-only architecture and explicit per-run state' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        $package = Join-Path $context.Root ('application space ' + [char]0x00e5)
        $null = [IO.Directory]::CreateDirectory((Join-Path $package 'src'))
        foreach ($entry in @('UniversalPodcastDownloader.ps1', 'UniversalPodcastDownloader.bat')) {
            [IO.File]::Copy((Join-Path $repositoryRoot $entry), (Join-Path $package $entry))
        }
        foreach ($source in Get-ChildItem -LiteralPath (Join-Path $repositoryRoot 'src') -File -Filter '*.ps1') {
            [IO.File]::Copy($source.FullName, (Join-Path $package ('src/' + $source.Name)))
        }
        Start-UpdFixtureServer -Context $context -PythonPath $script:architecturePython
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'imports from only the entry files and src without runtime initialization or caller-state changes' {
        @((Get-ChildItem -LiteralPath $package -Directory).Name) | Should -Be @('src')
        @(Get-ChildItem -LiteralPath $package -Recurse -File | Where-Object { $_.Extension -notin @('.ps1', '.bat') }).Count | Should -Be 0
        $run = Invoke-UpdArchitectureObservation -Context $context -Action Import
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.HostSurvived | Should -BeTrue
        $run.Result.ImportOutput | Should -Be 0
        $run.Result.ImportCallCount | Should -Be 0
        $run.Result.ImportChangedTree | Should -BeFalse
        $run.Result.ImportPowerType | Should -BeFalse
        $run.Result.ImportCreatedPolicy | Should -BeFalse
        $run.Result.ImportCreatedPageLimit | Should -BeFalse
        $run.Result.PreferencesPreserved | Should -BeTrue
        (Get-UpdFixtureState -Context $context).'/feeds/single.xml' | Should -BeNullOrEmpty
    }

    It 'uses the copied real pipeline for original bytes and verified repeat then preview with closed resources and URL correlation' {
        $run = Invoke-UpdArchitectureObservation -Context $context -Action Portable
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.HostSurvived | Should -BeTrue
        @($run.Result.Runs).Count | Should -Be 3
        $run.Result.Runs[0].Downloaded | Should -Be 1
        $run.Result.Runs[1].VerifiedSkipped | Should -Be 1
        $run.Result.Runs[2].Preview | Should -BeTrue
        @($run.Result.Runs | Where-Object ExitCode -ne 0).Count | Should -Be 0
        @($run.Result.CorrelationCounts | Where-Object { $_ -ne 0 }).Count | Should -Be 0
        $run.Result.HandlesClosed | Should -BeTrue
        $run.Result.PreviewChangedTree | Should -BeFalse
        $run.Result.PackageChanged | Should -BeFalse
        $run.Result.RuntimePythonAvailable | Should -BeFalse
        $run.Result.RunCreatedOptionState | Should -BeFalse
        $run.Result.RunPreferencesPreserved | Should -BeTrue
        $media = @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3')
        $media.Count | Should -Be 1
        (Get-FileHash -LiteralPath $media[0].FullName -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3') -Algorithm SHA256).Hash
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
        ($run.Result.Runs | ConvertTo-Json -Depth 20) | Should -Not -Match ([regex]::Escape($context.Root) + '|http://|One synthetic episode|Exception')
        @(Get-ChildItem -LiteralPath (Join-Path $context.Root 'local') -Recurse -File | Where-Object Extension -ne '.log').Count | Should -Be 0
    }

    It 'keeps two explicit invocation policies separate and standalone resolution at twenty pages and three attempts' {
        $run = Invoke-UpdArchitectureObservation -Context $context -Action Isolation
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.Runs[0].ExitCode | Should -Be 2
        $run.Result.Runs[0].PagesFetched | Should -Be 1
        $run.Result.Runs[1].ExitCode | Should -Be 0
        $run.Result.Runs[1].PagesFetched | Should -Be 2
        $run.Result.Standalone.MaxPages | Should -Be 20
        $run.Result.Standalone.PagesFetched | Should -Be 2
        @($run.Result.Requests).Count | Should -Be 5
        ($run.Result.Requests.MaxAttempts -join ',') | Should -Be '1,2,2,3,3'
        ($run.Result.Requests.HeaderTimeoutSeconds -join ',') | Should -Be '5,7,7,30,30'
        ($run.Result.Requests.IdleTimeoutSeconds -join ',') | Should -Be '6,8,8,30,30'
        $run.Result.RunCreatedOptionState | Should -BeFalse
        $run.Result.RunPreferencesPreserved | Should -BeTrue
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $context.Root 'local') | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/media/pagination.mp3' | Should -BeNullOrEmpty
    }
}
