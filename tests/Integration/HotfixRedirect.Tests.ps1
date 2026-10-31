BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for the owned redirect regression fixture.' }
    $script:hotfixPython = $python.Source
    $script:hotfixAudio = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
}

Describe 'HOTFIX: measurement redirect chain through the public entry point' -Tag 'Integration', 'HOTFIX' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:hotfixPython
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'downloads seven-hop original audio and verifies an unchanged skip on repeat' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -Mode Latest -FeedPath '/feeds/hotfix-chain.xml'
        $first.ExitCode | Should -Be 0 -Because ($first.Stdout + $first.Stderr)
        $first.Result.Succeeded | Should -BeTrue
        $first.Result.RunResult.Downloaded | Should -Be 1
        $files = @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be 1
        $hash = (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash
        $hash | Should -Be (Get-FileHash -LiteralPath $script:hotfixAudio -Algorithm SHA256).Hash
        $first.Result.RemainingTemporaryCount | Should -Be 0
        $before = Get-UpdFixtureState -Context $context
        foreach ($hop in 0..7) { $hopPath = '/hotfix/chain/' + $hop; $before.$hopPath | Should -Be 1 }
        $before.'/credential-received' | Should -BeNullOrEmpty
        $again = Invoke-UpdIntegrationWorker -Context $context -Action Download -Mode Latest -FeedPath '/feeds/hotfix-chain.xml'
        $again.ExitCode | Should -Be 0
        $again.Result.RunResult.VerifiedSkipped | Should -Be 1
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be $hash
        $after = Get-UpdFixtureState -Context $context
        foreach ($hop in 0..7) { $hopPath = '/hotfix/chain/' + $hop; $after.$hopPath | Should -Be 1 }
    }

    It 'rejects a permanent <Scenario> once without successful history or final garbage' -ForEach @(
        @{ Scenario = 'cycle'; LastHop = 0; InitialRequests = 2; Reason = 'redirect cycle detected' },
        @{ Scenario = 'limit'; LastHop = 12; InitialRequests = 1; Reason = 'redirect limit reached' }
    ) {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -Mode Latest -FeedPath ('/feeds/hotfix-' + $Scenario + '.xml')
        $run.ExitCode | Should -Be 2 -Because ($run.Stdout + $run.Stderr)
        $run.Result.RunResult.Failed | Should -Be 1
        $run.Result.RunResult.Episodes[0].Attempts | Should -Be 1
        $run.Stdout | Should -Match $Reason
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        $run.Result.RemainingTemporaryCount | Should -Be 0
        $stats = Get-UpdFixtureState -Context $context
        $initialPath = '/hotfix/' + $Scenario + '/0'
        $stats.$initialPath | Should -Be $InitialRequests
        $unrequestedPath = '/hotfix/' + $Scenario + '/' + ($LastHop + 1)
        $stats.$unrequestedPath | Should -BeNullOrEmpty
        foreach ($file in @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Force -File -Filter state.json)) {
            $state = [IO.File]::ReadAllText($file.FullName) | ConvertFrom-Json
            foreach ($episode in @($state.episodes)) {
                $episode.status | Should -Be 'failed'
                $episode.completed_utc | Should -BeNullOrEmpty
            }
        }
    }

    It 'retries one transient refusal within the existing budget then verifies original bytes' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -Mode Latest -FeedPath '/feeds/hotfix-transient.xml'
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr)
        $run.Result.RunResult.Downloaded | Should -Be 1
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/hotfix/transient/0' | Should -Be 2
        foreach ($hop in 1..7) { $hopPath = '/hotfix/transient/' + $hop; $stats.$hopPath | Should -Be 1 }
        $files = @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be 1
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $script:hotfixAudio -Algorithm SHA256).Hash
    }
}
