BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned loopback CLI process integration fixtures.' }
    $script:cliPython = $python.Source
    function Assert-UpdPrivateRunResult {
        param($Result, [int]$ExitCode, [string]$Status)
        $Result.Type | Should -Be 'Podcast.RunResult'
        $Result.SchemaVersion | Should -Be 1
        $Result.ExitCode | Should -Be $ExitCode
        $Result.Status | Should -Be $Status
        $Result.Complete | Should -Be ($ExitCode -eq 0)
        $names = @('Type','SchemaVersion','Mode','Preview','Complete','ExitCode','Status','Planned','Downloaded','VerifiedSkipped','LegacyUnverified','Conflicts','Failed','Deferred','Cancelled','CatalogueComplete','CatalogueStopReason','PagesFetched','Episodes','Message','Plan')
        @($Result.PSObject.Properties.Name | Where-Object { $_ -notin $names }).Count | Should -Be 0
        ($Result | ConvertTo-Json -Depth 20) | Should -Not -Match 'https?://|callable-output|UPD-Integration-|Synthetic private cancellation sentinel|Original resume audio|One synthetic episode|StackTrace|Exception'
        foreach ($episode in $Result.Episodes) {
            @($episode.PSObject.Properties.Name | Where-Object { $_ -notin @('EpisodeId','Outcome','Bytes','Verification','Attempts','Message') }).Count | Should -Be 0
        }
    }
}

Describe 'A039/A040 actual CLI process semantics and private runtime results' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:cliPython
    }
    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'infers Custom when CustomCount alone is supplied at the process boundary' {
        $run = Invoke-UpdCliProcess -Context $context -CustomCount 2
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr)
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 2
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 2
    }

    It 'rejects an explicit Latest and count conflict before fetching a feed or writing output' {
        $run = Invoke-UpdCliProcess -Context $context -Mode Latest -CustomCount 2
        $run.ExitCode | Should -Be 1
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/feeds/atom.xml' | Should -BeNullOrEmpty
    }

    It 'maps an episode download failure to incomplete exit two without a completed banner' {
        $run = Invoke-UpdCliProcess -Context $context -FeedPath '/feeds/transaction-html.xml' -Mode All
        $run.ExitCode | Should -Be 2
        $run.Stdout | Should -Match 'Failed\s+: 1'
        $run.Stdout | Should -Not -Match 'Run completed|\[OK\]'
    }

    It 'maps an unresolved cyclic catalogue to incomplete exit two after accessible work' {
        $run = Invoke-UpdCliProcess -Context $context -FeedPath '/feeds/page-1.xml' -Mode All
        $run.ExitCode | Should -Be 2
        $run.Stdout | Should -Match 'Downloaded\s+: 2'
        $run.Stdout | Should -Not -Match 'Run completed|\[OK\]'
    }

    It 'accepts explicit Custom count at the process boundary' {
        $run = Invoke-UpdCliProcess -Context $context -Mode Custom -CustomCount 2 -NonInteractive
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr)
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 2
    }

    It 'rejects invalid explicit options <Label> before network or archive writes' -TestCases @(
        @{ Label = 'All plus count'; Parameters = @{ Mode = 'All'; CustomCount = 2 } }
        @{ Label = 'zero count'; Parameters = @{ CustomCount = 0 } }
        @{ Label = 'negative count'; Parameters = @{ CustomCount = -1 } }
        @{ Label = 'Custom without count'; Parameters = @{ Mode = 'Custom' } }
    ) {
        param($Label, $Parameters)
        $null = $Label
        $run = Invoke-UpdCliProcess -Context $context -NonInteractive @Parameters
        $run.ExitCode | Should -Be 1
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/feeds/atom.xml' | Should -BeNullOrEmpty
    }

    It 'fails noninteractive missing input or ambiguous discovery without calling Read-Host: <Label>' -TestCases @(
        @{ Label = 'missing feed'; Parameters = @{ WithoutFeed = $true }; Requests = 0 }
        @{ Label = 'multiple discovered feeds'; Parameters = @{ FeedPath = '/show' }; Requests = 1 }
    ) {
        param($Label, $Parameters, $Requests)
        $null = $Label
        $run = Invoke-UpdCliResultWorker -Context $context -Mode All @Parameters
        $run.ExitCode | Should -Be 1 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.HostSurvived | Should -BeTrue
        $run.Result.PromptCount | Should -Be 0
        Assert-UpdPrivateRunResult -Result $run.Result.RunResults[0] -ExitCode 1 -Status fatal
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
        [int](Get-UpdFixtureState -Context $context).'/show' | Should -Be $Requests
    }

    It 'maps <Label> preview to the correct process exit and leaves archive and media untouched' -TestCases @(
        @{ Label = 'clean'; FeedPath = '/feeds/single.xml'; ExitCode = 0 }
        @{ Label = 'partial catalogue'; FeedPath = '/feeds/page-1.xml'; ExitCode = 2 }
    ) {
        param($Label, $FeedPath, $ExitCode)
        $null = $Label
        $run = Invoke-UpdCliProcess -Context $context -Mode All -FeedPath $FeedPath -NonInteractive -Preview
        $run.ExitCode | Should -Be $ExitCode -Because ($run.Stdout + $run.Stderr)
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
        $run.Stdout | Should -Not -Match 'Run completed|\[OK\]'
    }

    It 'returns incomplete for an initially empty catalogue with an unresolved advertised page' {
        $run = Invoke-UpdCliProcess -Context $context -Mode All -FeedPath '/feeds/empty-pagination-gap.xml' -NonInteractive
        $run.ExitCode | Should -Be 2 -Because ($run.Stdout + $run.Stderr)
        $run.Stdout | Should -Match 'Feed catalogue incomplete'
        $run.Stdout | Should -Not -Match 'Run completed|\[OK\]'
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/pagination/missing.xml' | Should -Be 1
    }

    It 'imports silently and returns one private result per repeat without exiting the caller' {
        $run = Invoke-UpdCliResultWorker -Context $context -Mode All -Repeat 2
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.HostSurvived | Should -BeTrue
        $run.Result.ImportOutputCount | Should -Be 0
        $run.Result.PromptCount | Should -Be 0
        @($run.Result.RunResults).Count | Should -Be 2
        Assert-UpdPrivateRunResult -Result $run.Result.RunResults[0] -ExitCode 0 -Status success
        Assert-UpdPrivateRunResult -Result $run.Result.RunResults[1] -ExitCode 0 -Status success
        $run.Result.RunResults[0].Downloaded | Should -Be 1
        $run.Result.RunResults[1].VerifiedSkipped | Should -Be 1
        $run.Result.RunResults[1].Episodes[0].Outcome | Should -Be 'verified_skip'
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
        $file = Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3'
        (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3') -Algorithm SHA256).Hash
    }

    It 'keeps a private consistent result for <Label> and survives the callable run' -TestCases @(
        @{ Label = 'episode failure'; FeedPath = '/feeds/transaction-html.xml'; Preview = $false; Code = 2; Status = 'incomplete'; Failed = 1; Catalogue = $true }
        @{ Label = 'catalogue gap'; FeedPath = '/feeds/pagination-gap.xml'; Preview = $false; Code = 2; Status = 'incomplete'; Failed = 0; Catalogue = $false }
        @{ Label = 'fatal malformed feed'; FeedPath = '/feeds/malformed.xml'; Preview = $false; Code = 1; Status = 'fatal'; Failed = 0; Catalogue = $true }
        @{ Label = 'clean preview'; FeedPath = '/feeds/single.xml'; Preview = $true; Code = 0; Status = 'preview'; Failed = 0; Catalogue = $true }
        @{ Label = 'partial preview'; FeedPath = '/feeds/pagination-gap.xml'; Preview = $true; Code = 2; Status = 'incomplete'; Failed = 0; Catalogue = $false }
    ) {
        param($Label, $FeedPath, $Preview, $Code, $Status, $Failed, $Catalogue)
        $null = $Label
        $run = Invoke-UpdCliResultWorker -Context $context -Mode All -FeedPath $FeedPath -Preview:$Preview
        $run.ExitCode | Should -Be $Code -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.HostSurvived | Should -BeTrue
        @($run.Result.RunResults).Count | Should -Be 1
        $result = $run.Result.RunResults[0]
        Assert-UpdPrivateRunResult -Result $result -ExitCode $Code -Status $Status
        $result.Preview | Should -Be $Preview
        $result.Failed | Should -Be $Failed
        $result.CatalogueComplete | Should -Be $Catalogue
        if ($Preview) {
            Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
            (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
            $result.Planned | Should -Be 1
        }
    }

    It 'classifies catchable cancellation after a real media write without retry or host exit' {
        $run = Invoke-UpdCliResultWorker -Context $context -Mode All -Action Cancel -FeedPath '/resume/feed/valid'
        $run.ExitCode | Should -Be 130 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.HostSurvived | Should -BeTrue
        $run.Result.CancelBytes | Should -BeGreaterThan 0
        $run.Result.PartialClosed | Should -BeTrue
        $result = $run.Result.RunResults[0]
        Assert-UpdPrivateRunResult -Result $result -ExitCode 130 -Status cancelled
        $result.Cancelled | Should -Be 1
        $result.Downloaded | Should -Be 0
        $result.Episodes[0].Outcome | Should -Be 'cancelled'
        (Get-UpdFixtureState -Context $context).'/resume/media/valid' | Should -Be 1
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Force -File -Filter 'resume-*.json').Count | Should -Be 1
        $run.Stdout | Should -Not -Match 'Run completed|\[OK\]'
    }
}
