BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for the owned transport loopback fixture server.' }
    $script:transportPython = $python.Source
    $script:transportAudio = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'

    function Invoke-UpdTransportWorker {
        param(
            [Parameter(Mandatory)]$Context,
            [Parameter(Mandatory)][ValidateSet('Metadata', 'Media')][string]$Kind,
            [Parameter(Mandatory)][ValidatePattern('^[a-z0-9-]+$')][string]$Scenario,
            [hashtable]$Settings = @{}
        )
        $defaults = @{
            MaxAttempts = 3; HeaderTimeoutSeconds = 3.0; IdleTimeoutSeconds = 3.0
            RetryBudgetSeconds = 10.0; BaseDelaySeconds = 0.01; MaxDelaySeconds = 0.02
        }
        foreach ($key in $Settings.Keys) { $defaults[$key] = $Settings[$key] }
        $token = [guid]::NewGuid().ToString('N')
        $prefix = if ($Kind -eq 'Media') { 'feed' } else { 'metadata' }
        $configPath = Join-Path $Context.Root ($token + '-input.json')
        $config = @{
            ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'
            Kind = $Kind; Url = $Context.BaseUrl + '/transport/' + $prefix + '/' + $Scenario
            OutputPath = Join-Path $Context.Root 'output'; Settings = $defaults
            ResultPath = Join-Path $Context.Root ($token + '-result.json')
        }
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $owned = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engine) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-TransportWorker.ps1'), '-ConfigPath', $configPath
        )
        if (-not $owned.Process.HasExited -and -not $owned.Process.WaitForExit(30000)) {
            $owned.Process.Kill()
            $owned.Process.WaitForExit()
            throw 'Owned transport test child exceeded 30 seconds.'
        }
        $stdout = $owned.Output.Result
        $stderr = $owned.ErrorOutput.Result
        if (-not [IO.File]::Exists($config.ResultPath)) {
            throw "Transport test child produced no result: stdout=$stdout; stderr=$stderr"
        }
        [pscustomobject]@{
            Result = ([IO.File]::ReadAllText($config.ResultPath) | ConvertFrom-Json)
            ExitCode = $owned.Process.ExitCode; Stdout = $stdout; Stderr = $stderr
            OutputPath = $config.OutputPath; Route = '/transport/' + $Kind.ToLowerInvariant() + '/' + $Scenario
        }
    }

    function Assert-UpdTransportSuccess {
        param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)][string]$Kind)
        $Run.ExitCode | Should -Be 0 -Because ($Run.Stdout + $Run.Stderr + $Run.Result.ErrorMessage)
        $Run.Result.Succeeded | Should -BeTrue
        if ($Kind -eq 'Metadata') {
            $Run.Result.Bytes | Should -BeGreaterThan 0
            $Run.Result.Content | Should -Match '<title>Transport fixture</title>'
            Test-Path -LiteralPath $Run.OutputPath | Should -BeFalse
            return
        }
        $Run.Stdout | Should -Match 'Downloaded\s+: 1'
        $Run.Stdout | Should -Match 'Failed\s+: 0'
        $files = @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be 1
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $script:transportAudio -Algorithm SHA256).Hash
        $states = @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -Force -File -Filter state.json)
        $states.Count | Should -Be 1
        $state = [IO.File]::ReadAllText($states[0].FullName) | ConvertFrom-Json
        @($state.episodes).Count | Should -Be 1
        $state.episodes[0].status | Should -Be 'transfer_verified'
        @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -Force -File -Filter '.upd-*.tmp').Count | Should -Be 0
    }

    function Assert-UpdTransportFailure {
        param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)][string]$Kind)
        $Run.ExitCode | Should -Be 1 -Because ($Run.Stdout + $Run.Stderr + $Run.Result.ErrorMessage)
        $Run.Result.Succeeded | Should -BeFalse
        if ($Kind -eq 'Metadata') {
            $Run.Result.Bytes | Should -Be 0
            $Run.Result.Content | Should -BeNullOrEmpty
            Test-Path -LiteralPath $Run.OutputPath | Should -BeFalse
            return
        }
        $Run.Stdout | Should -Match 'Downloaded\s+: 0'
        $Run.Stdout | Should -Match 'Failed\s+: 1'
        @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -Force -File -Filter '.upd-*.tmp').Count | Should -Be 0
        foreach ($file in @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -Force -File -Filter state.json)) {
            $state = [IO.File]::ReadAllText($file.FullName) | ConvertFrom-Json
            foreach ($episode in @($state.episodes)) {
                $episode.status | Should -Be 'failed'
                $episode.completed_utc | Should -BeNullOrEmpty
                $episode.local_sha256 | Should -BeNullOrEmpty
                $episode.bytes | Should -BeNullOrEmpty
            }
        }
    }
}

Describe 'A025/A026: bounded transport against real loopback responses' -Tag 'Integration', 'A025', 'A026' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:transportPython
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'A025 retries transient <Code> for <Kind> and preserves complete content' -ForEach @(
        foreach ($kind in @('Metadata', 'Media')) {
            foreach ($code in @(408, 429, 500, 502, 503, 504)) { @{ Kind = $kind; Code = $code } }
        }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario ('once-' + $Code)
        Assert-UpdTransportSuccess -Run $run -Kind $Kind
        (Get-UpdFixtureState -Context $context).($run.Route) | Should -Be 2
    }

    It 'A025 does not retry permanent <Code> for <Kind> or publish failed content' -ForEach @(
        foreach ($kind in @('Metadata', 'Media')) {
            foreach ($code in @(400, 403, 404, 501)) { @{ Kind = $kind; Code = $code } }
        }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario ('permanent-' + $Code)
        Assert-UpdTransportFailure -Run $run -Kind $Kind
        (Get-UpdFixtureState -Context $context).($run.Route) | Should -Be 1
    }

    It 'A025 limits persistent transient <Kind> failures to exactly the configured attempts' -ForEach @(
        @{ Kind = 'Metadata' }; @{ Kind = 'Media' }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario 'always-503' -Settings @{ MaxAttempts = 2 }
        Assert-UpdTransportFailure -Run $run -Kind $Kind
        (Get-UpdFixtureState -Context $context).($run.Route) | Should -Be 2
    }

    It 'A025 honors Retry-After <Form> for <Kind> without an early retry' -ForEach @(
        foreach ($kind in @('Metadata', 'Media')) {
            foreach ($form in @('delta', 'date')) { @{ Kind = $kind; Form = $form } }
        }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario ('retry-' + $Form)
        Assert-UpdTransportSuccess -Run $run -Kind $Kind
        $allEvents = Invoke-RestMethod -Uri ($context.BaseUrl + '/__transport')
        $events = @($allEvents | Where-Object { $_.path -eq $run.Route })
        $events.Count | Should -Be 2
        # Server timestamps avoid conflating engine startup with retry waiting.
        $events[1].received | Should -BeGreaterOrEqual ($events[0].not_before - 0.02)
    }

    It 'A025 defers Retry-After <Form> beyond the retry budget for <Kind>' -ForEach @(
        foreach ($kind in @('Metadata', 'Media')) {
            foreach ($form in @('delta', 'date')) { @{ Kind = $kind; Form = $form } }
        }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario ('defer-' + $Form) -Settings @{ RetryBudgetSeconds = 0.25 }
        Assert-UpdTransportFailure -Run $run -Kind $Kind
        (Get-UpdFixtureState -Context $context).($run.Route) | Should -Be 1
        $run.Result.ElapsedSeconds | Should -BeLessThan 4
    }

    It 'A025 restarts truncated <Kind> content and never appends a partial response' -ForEach @(
        @{ Kind = 'Metadata' }; @{ Kind = 'Media' }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario 'truncated-once'
        Assert-UpdTransportSuccess -Run $run -Kind $Kind
        (Get-UpdFixtureState -Context $context).($run.Route) | Should -Be 2
    }

    It 'A025 honors a redirect server delay before requesting the <Kind> target' -ForEach @(
        @{ Kind = 'Metadata' }; @{ Kind = 'Media' }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario 'redirect-wait'
        Assert-UpdTransportSuccess -Run $run -Kind $Kind
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__transport')
        $initial = @($events | Where-Object { $_.path -eq $run.Route })
        $target = @($events | Where-Object { $_.path -eq ('/transport/' + $Kind.ToLowerInvariant() + '/redirect-target') })
        $initial.Count | Should -Be 1
        $target.Count | Should -Be 1
        $target[0].received | Should -BeGreaterOrEqual ($initial[0].not_before - 0.02)
    }

    It 'A025 defers an over-budget redirect without contacting the <Kind> target' -ForEach @(
        @{ Kind = 'Metadata' }; @{ Kind = 'Media' }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario 'redirect-defer' -Settings @{ RetryBudgetSeconds = 0.25 }
        Assert-UpdTransportFailure -Run $run -Kind $Kind
        $stats = Get-UpdFixtureState -Context $context
        $stats.($run.Route) | Should -Be 1
        $targetRoute = '/transport/' + $Kind.ToLowerInvariant() + '/redirect-target'
        $stats.$targetRoute | Should -BeNullOrEmpty
        $run.Result.ElapsedSeconds | Should -BeLessThan 4
    }

    It 'A026 bounds <Scenario> in <Kind> and leaves no completed result' -ForEach @(
        foreach ($kind in @('Metadata', 'Media')) {
            foreach ($scenario in @('delay-headers', 'stall-body')) { @{ Kind = $kind; Scenario = $scenario } }
        }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario $Scenario -Settings @{
            HeaderTimeoutSeconds = 0.15; IdleTimeoutSeconds = 0.15; MaxAttempts = 2
        }
        Assert-UpdTransportFailure -Run $run -Kind $Kind
        (Get-UpdFixtureState -Context $context).($run.Route) | Should -Be 2
        $run.Result.ElapsedSeconds | Should -BeLessThan 4
    }

    It 'A026 accepts actively progressing <Kind> content beyond header timeout and retry budget' -ForEach @(
        @{ Kind = 'Metadata' }; @{ Kind = 'Media' }
    ) {
        $run = Invoke-UpdTransportWorker -Context $context -Kind $Kind -Scenario 'progress' -Settings @{
            HeaderTimeoutSeconds = 0.3; IdleTimeoutSeconds = 0.5; RetryBudgetSeconds = 0.2
        }
        Assert-UpdTransportSuccess -Run $run -Kind $Kind
        (Get-UpdFixtureState -Context $context).($run.Route) | Should -Be 1
        $run.Result.ElapsedSeconds | Should -BeGreaterThan 1.0
    }
}
