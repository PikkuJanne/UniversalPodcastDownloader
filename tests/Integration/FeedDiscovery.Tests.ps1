BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python 3.9+ is required for loopback integration tests. It is a development dependency only.' }
    $script:pythonPath = $python.Source

    function Reset-UpdDiscoveryRequest {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test helper resets only counters in its owned loopback fixture server; each assertion requires a fresh request count.')]
        [CmdletBinding()]
        param($Context)
        $null = Invoke-WebRequest -Uri ($Context.BaseUrl + '/__reset') -Method Post -UseBasicParsing -TimeoutSec 10
    }

    function Assert-UpdDiscoveryNoWrite {
        param($Context, $Run)
        Test-Path -LiteralPath $Run.OutputPath | Should -BeFalse
        @(Get-ChildItem -LiteralPath $Context.Root -Recurse -File -Filter '*.log').Count | Should -Be 0
        @((Get-UpdFixtureState -Context $Context).PSObject.Properties | Where-Object { $_.Name -like '/media/*' }).Count | Should -Be 0
    }
}

Describe 'A030 shared feed discovery against synthetic loopback fixtures' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:pythonPath
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'returns all page candidates in document order from <FeedPath>' -TestCases @(
        @{ FeedPath = '/show'; FinalPath = '/show' }
        @{ FeedPath = '/redirect/show'; FinalPath = '/show' }
    ) {
        param($FeedPath, $FinalPath)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Source -FeedPath $FeedPath
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.Kind | Should -Be 'Html'
        $run.Result.ResolvedUrl | Should -Be ($context.BaseUrl + $FeedPath)
        $run.Result.FinalUri | Should -Be ($context.BaseUrl + $FinalPath)
        @($run.Result.Candidates).Count | Should -Be 2
        $run.Result.Candidates[0] | Should -Be ($context.BaseUrl + '/feeds/single.xml?fake_token=NOT_A_SECRET&x=1')
        $run.Result.Candidates[1] | Should -Be ($context.BaseUrl + '/feeds/atom.xml')
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/show' | Should -Be 1
        $stats.'/feeds/single.xml' | Should -BeNullOrEmpty
        $stats.'/feeds/atom.xml' | Should -BeNullOrEmpty
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'uses final URI, relative base and decoded duplicate links for <FeedPath>' -TestCases @(
        @{ FeedPath = '/discovery/redirect'; FinalPath = '/discovery/final/show.html'; Candidate = '/feeds/single.xml' }
        @{ FeedPath = '/discovery/base.html'; FinalPath = '/discovery/base.html'; Candidate = '/feeds/single.xml?fixture=base&part=1' }
        @{ FeedPath = '/discovery/single.html'; FinalPath = '/discovery/single.html'; Candidate = '/feeds/single.xml?fixture=single&part=1' }
    ) {
        param($FeedPath, $FinalPath, $Candidate)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Source -FeedPath $FeedPath
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.Kind | Should -Be 'Html'
        $run.Result.FinalUri | Should -Be ($context.BaseUrl + $FinalPath)
        @($run.Result.Candidates).Count | Should -Be 1
        $run.Result.Candidates[0] | Should -Be ($context.BaseUrl + $Candidate)
        (Get-UpdFixtureState -Context $context).'/feeds/single.xml' | Should -BeNullOrEmpty
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'reuses an already fetched <Kind> response without another request' -TestCases @(
        @{ FeedPath = '/feeds/single.xml'; Kind = 'Rss'; Count = 1 }
        @{ FeedPath = '/feeds/atom.xml'; Kind = 'Atom'; Count = 2 }
    ) {
        param($FeedPath, $Kind, $Count)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Source -FeedPath $FeedPath -ReuseResponse
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.Kind | Should -Be $Kind
        $run.Result.ItemCount | Should -Be $Count
        $run.Result.ContentLength | Should -BeGreaterThan 0
        (Get-UpdFixtureState -Context $context).PSObject.Properties[$FeedPath].Value | Should -Be 1
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'reads <FeedPath> once through either direct CLI or real TUI preview' -TestCases @(
        @{ FeedPath = '/feeds/single.xml'; RequestPath = '/feeds/single.xml' }
        @{ FeedPath = '/feeds/atom.xml'; RequestPath = '/feeds/atom.xml' }
        @{ FeedPath = '/redirect/feed'; RequestPath = '/feeds/single.xml' }
    ) {
        param($FeedPath, $RequestPath)
        foreach ($action in @('Preview', 'InteractivePreview')) {
            Reset-UpdDiscoveryRequest -Context $context
            $run = Invoke-UpdIntegrationWorker -Context $context -Action $action -FeedPath $FeedPath
            $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
            $run.ExitCode | Should -Be 0
            $run.Stdout | Should -Match 'Using feed:'
            (Get-UpdFixtureState -Context $context).PSObject.Properties[$RequestPath].Value | Should -Be 1
            if ($action -eq 'InteractivePreview') { $run.Result.PromptCount | Should -Be 1 }
            Assert-UpdDiscoveryNoWrite -Context $context -Run $run
        }
    }

    It 'resolves one discovered feed from <FeedPath> once through CLI and TUI' -TestCases @(
        @{ FeedPath = '/discovery/redirect'; FinalPath = '/discovery/final/show.html' }
        @{ FeedPath = '/discovery/base.html'; FinalPath = '/discovery/base.html' }
        @{ FeedPath = '/discovery/single.html'; FinalPath = '/discovery/single.html' }
    ) {
        param($FeedPath, $FinalPath)
        foreach ($action in @('Preview', 'InteractivePreview')) {
            Reset-UpdDiscoveryRequest -Context $context
            $run = Invoke-UpdIntegrationWorker -Context $context -Action $action -FeedPath $FeedPath
            $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
            $run.ExitCode | Should -Be 0
            $stats = Get-UpdFixtureState -Context $context
            $stats.PSObject.Properties[$FinalPath].Value | Should -Be 1
            $stats.'/feeds/single.xml' | Should -Be 1
            if ($action -eq 'InteractivePreview') { $run.Result.PromptCount | Should -Be 1 }
            Assert-UpdDiscoveryNoWrite -Context $context -Run $run
        }
    }

    It 'rejects multiple candidates before any candidate request in <Action>' -TestCases @(
        @{ Action = 'Resolve' }
        @{ Action = 'Preview' }
    ) {
        param($Action)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action $Action -FeedPath '/show'
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match 'Multiple feed links were found.*direct feed URL.*-FeedUrl'
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/show' | Should -Be 1
        $stats.'/feeds/single.xml' | Should -BeNullOrEmpty
        $stats.'/feeds/atom.xml' | Should -BeNullOrEmpty
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'reuses completed interactive resolution for numbered candidate <Selection>' -TestCases @(
        @{ Selection = '1'; RequestPath = '/feeds/single.xml'; OtherPath = '/feeds/atom.xml'; Count = 1; UrlSuffix = '/feeds/single.xml?fake_token=NOT_A_SECRET&x=1' }
        @{ Selection = '2'; RequestPath = '/feeds/atom.xml'; OtherPath = '/feeds/single.xml'; Count = 2; UrlSuffix = '/feeds/atom.xml' }
    ) {
        param($Selection, $RequestPath, $OtherPath, $Count, $UrlSuffix)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Discover -FeedPath '/show' -Selection $Selection
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.ResolvedUrl | Should -Be ($context.BaseUrl + $UrlSuffix)
        $run.Result.ItemCount | Should -Be $Count
        $run.Result.PromptCount | Should -Be 2
        @($run.Result.Candidates).Count | Should -Be 2
        $run.Result.Candidates[0] | Should -Be ($context.BaseUrl + '/feeds/single.xml?fake_token=NOT_A_SECRET&x=1')
        $run.Result.Candidates[1] | Should -Be ($context.BaseUrl + '/feeds/atom.xml')
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/show' | Should -Be 1
        $stats.PSObject.Properties[$RequestPath].Value | Should -Be 1
        $stats.PSObject.Properties[$OtherPath].Value | Should -BeNullOrEmpty
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'fetches only numbered candidate <Selection> once in the real TUI entry point' -TestCases @(
        @{ Selection = '1'; RequestPath = '/feeds/single.xml'; OtherPath = '/feeds/atom.xml' }
        @{ Selection = '2'; RequestPath = '/feeds/atom.xml'; OtherPath = '/feeds/single.xml' }
    ) {
        param($Selection, $RequestPath, $OtherPath)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action InteractivePreview -FeedPath '/show' -Selection $Selection
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.PromptCount | Should -Be 2
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/show' | Should -Be 1
        $stats.PSObject.Properties[$RequestPath].Value | Should -Be 1
        $stats.PSObject.Properties[$OtherPath].Value | Should -BeNullOrEmpty
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'repeats an invalid numbered choice without repeating metadata requests' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action InteractivePreview -FeedPath '/show' -Selection @('0', 'word', '2')
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.PromptCount | Should -Be 4
        $run.Stdout | Should -Match 'Enter one of the listed feed numbers'
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/show' | Should -Be 1
        $stats.'/feeds/atom.xml' | Should -Be 1
        $stats.'/feeds/single.xml' | Should -BeNullOrEmpty
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'keeps useful distinct resolution failures for <FeedPath>' -TestCases @(
        @{ FeedPath = '/feeds/empty.xml'; Message = 'valid but contains no episodes' }
        @{ FeedPath = '/feeds/empty-atom.xml'; Message = 'valid but contains no episodes' }
        @{ FeedPath = '/feeds/malformed.xml'; Message = 'Source XML is invalid or exceeds safe parser limits' }
        @{ FeedPath = '/feeds/wrong-root.xml'; Message = 'source XML root is not a supported RSS or Atom feed' }
        @{ FeedPath = '/feeds/wrong-atom-namespace.xml'; Message = 'source XML root is not a supported RSS or Atom feed' }
        @{ FeedPath = '/show/not-feed'; Message = 'No RSS or Atom feed links were found on the page' }
    ) {
        param($FeedPath, $Message)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Resolve -FeedPath $FeedPath
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match $Message
        (Get-UpdFixtureState -Context $context).PSObject.Properties[$FeedPath].Value | Should -Be 1
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'does not recursively crawl when an advertised feed target is another HTML page' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Resolve -FeedPath '/discovery/nonfeed-link.html'
        $run.Result.Succeeded | Should -BeFalse
        $run.Result.ErrorMessage | Should -Match 'discovered URL did not return an RSS or Atom feed'
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/discovery/nonfeed-link.html' | Should -Be 1
        $stats.'/show/not-feed' | Should -Be 1
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }

    It 'classifies a valid empty <Kind> feed before resolution reports no episodes' -TestCases @(
        @{ FeedPath = '/feeds/empty.xml'; Kind = 'Rss' }
        @{ FeedPath = '/feeds/empty-atom.xml'; Kind = 'Atom' }
    ) {
        param($FeedPath, $Kind)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Source -FeedPath $FeedPath
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.Kind | Should -Be $Kind
        $run.Result.ItemCount | Should -Be 0
        @($run.Result.Candidates).Count | Should -Be 1
        $run.Result.Candidates[0] | Should -Be ($context.BaseUrl + $FeedPath)
        Assert-UpdDiscoveryNoWrite -Context $context -Run $run
    }
}
