BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned loopback diagnostic integration fixtures.' }
    $script:diagnosticPython = $python.Source

    function Start-UpdDiagnosticWorker {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates configuration and a tracked child only inside the caller-owned integration directory.')]
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]$Context,
            [Parameter(Mandatory)][string]$Action,
            [string]$FeedPath = '/feeds/malformed.xml',
            [string]$LogRoot,
            [string]$LegacyPath,
            [switch]$BlockPrimary,
            [switch]$BlockBoth
        )

        $token = [guid]::NewGuid().ToString('N')
        $configPath = Join-Path $Context.Root ($token + '-diagnostic-config.json')
        $config = @{
            Action = $Action
            ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'
            Root = Join-Path $Context.Root ('r-' + $token.Substring(0, 8))
            ResultPath = Join-Path $Context.Root ($token + '-diagnostic-result.json')
            FeedUrl = $Context.BaseUrl + $FeedPath
            LogRoot = $LogRoot
            LegacyPath = $LegacyPath
            BlockPrimary = [bool]$BlockPrimary
            BlockBoth = [bool]$BlockBoth
        }
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $engineName = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $owned = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engineName) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-DiagnosticWorker.ps1'), '-ConfigPath', $configPath
        )
        return [pscustomobject]@{ Owned = $owned; Config = $config }
    }

    function Receive-UpdDiagnosticWorker {
        param([Parameter(Mandatory)]$Run)

        if (-not $Run.Owned.Process.HasExited -and -not $Run.Owned.Process.WaitForExit(30000)) {
            $Run.Owned.Process.Kill()
            $Run.Owned.Process.WaitForExit()
            throw 'Diagnostic integration child exceeded 30 seconds; only its owned process was stopped.'
        }
        $result = if (Test-Path -LiteralPath $Run.Config.ResultPath) {
            [IO.File]::ReadAllText($Run.Config.ResultPath) | ConvertFrom-Json
        } else { $null }
        $stdout = $Run.Owned.Output.Result
        $stderr = $Run.Owned.ErrorOutput.Result
        if ($null -eq $result -and $Run.Config.Action -ne 'RawFailure') {
            throw "Diagnostic child produced no result: stdout=$stdout; stderr=$stderr"
        }
        return [pscustomobject]@{
            Result = $result; ExitCode = $Run.Owned.Process.ExitCode
            Stdout = $stdout; Stderr = $stderr; Root = $Run.Config.Root
        }
    }

    function Read-UpdDiagnosticLog {
        param([Parameter(Mandatory)][string]$Root)

        $strictUtf8 = [Text.UTF8Encoding]::new($false, $true)
        return @(Get-ChildItem -LiteralPath $Root -Filter '*.log' -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            [pscustomobject]@{ Path = $_.FullName; Text = $strictUtf8.GetString([IO.File]::ReadAllBytes($_.FullName)) }
        })
    }
}

Describe 'A019/A021: startup diagnostics in isolated real processes' -Tag 'Integration', 'A019', 'A021' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:diagnosticPython
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'A019 records discovery failure before a podcast title or archive exists' {
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action FeedFailure)
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match '^No episodes found in the feed\.'
        $logs = @(Read-UpdDiagnosticLog -Root $run.Root)
        $logs.Count | Should -Be 1
        $logs[0].Path | Should -Match '[\\/]local[\\/]UniversalPodcastDownloader[\\/]Logs[\\/]'
        $logs[0].Text | Should -Match 'No episodes found|Fatal|ERROR'
        Test-Path -LiteralPath (Join-Path $run.Root 'output') | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/feeds/malformed.xml' | Should -Be 1
    }

    It 'A019 falls back when the preferred log directory cannot be created' {
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action FeedFailure -BlockPrimary)
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match '^No episodes found in the feed\.'
        $logs = @(Read-UpdDiagnosticLog -Root $run.Root)
        $logs.Count | Should -Be 1
        $logs[0].Path | Should -Match '[\\/]temp[\\/]UniversalPodcastDownloader[\\/]Logs[\\/]'
        [IO.File]::ReadAllText((Join-Path $run.Root 'local')) | Should -Be 'Owned synthetic obstruction.'
    }

    It 'A019 reports unavailable logging on stderr and retains the discovery failure' {
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action FeedFailure -BlockBoth)
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match '^No episodes found in the feed\.'
        $run.Stderr | Should -Match '(?i)diagnostic|log'
        $run.Stderr | Should -Match '(?i)unavailable|could not|cannot|failed'
        @(Read-UpdDiagnosticLog -Root $run.Root).Count | Should -Be 0
    }

    It 'A019 records an invalid output path without reaching discovery' {
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action InvalidOutput)
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Not -BeNullOrEmpty
        @(Read-UpdDiagnosticLog -Root $run.Root).Count | Should -Be 1
        [IO.File]::ReadAllText((Join-Path $run.Root 'output')) | Should -Be 'Preserve this synthetic existing file.'
        (Get-UpdFixtureState -Context $context).'/feeds/malformed.xml' | Should -BeNullOrEmpty
    }

    It 'A019 retains startup evidence when a successful download opens its archive log' {
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action Download -FeedPath '/feeds/single.xml')
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $logs = @(Read-UpdDiagnosticLog -Root $run.Root)
        $logs.Count | Should -Be 2
        @($logs | Where-Object { $_.Path.StartsWith((Join-Path $run.Root 'local') + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) }).Count | Should -Be 1
        $archiveLog = @($logs | Where-Object Path -Match '[\\/]output[\\/]')
        $archiveLog.Count | Should -Be 1
        $archiveLog[0].Text | Should -Match 'Run completed'
        $names = @($logs | ForEach-Object { [IO.Path]::GetFileName($_.Path) } | Select-Object -Unique)
        $names.Count | Should -Be 1 -Because 'startup and archive logs retain one correlated run identity'
        @(Get-ChildItem -LiteralPath (Join-Path $run.Root 'output') -Filter '*.mp3' -File -Recurse).Count | Should -Be 1
    }

    It 'A021 writes strict UTF-8 and unique run names for simultaneous startup' {
        $sharedLogs = Join-Path $context.Root 'shared-logs'
        # Isolate concurrent run-file allocation from a separate directory-
        # creation race. Unavailable roots have their own fallback assertions.
        $null = [IO.Directory]::CreateDirectory($sharedLogs)
        $started = @(1..3 | ForEach-Object { Start-UpdDiagnosticWorker -Context $context -Action ApiWrite -LogRoot $sharedLogs })
        $runs = @($started | ForEach-Object { Receive-UpdDiagnosticWorker -Run $_ })
        foreach ($run in $runs) { $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage) }
        @($runs.Result.RunId | Select-Object -Unique).Count | Should -Be 3
        $logs = @(Read-UpdDiagnosticLog -Root $sharedLogs)
        $logs.Count | Should -Be 3
        @($logs.Path | Select-Object -Unique).Count | Should -Be 3
        $unicode = 'Synthetic Unicode ' + [char]0x00e4 + [char]0x65e5 + [char]::ConvertFromUtf32(0x1f3a7)
        foreach ($log in $logs) { $log.Text | Should -Match ([regex]::Escape($unicode)) }
    }

    It 'A021 does not replace an active exception when a previously working writer fails' {
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action AppendFailure -LogRoot (Join-Path $context.Root 'write-failure'))
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Be 'Original synthetic operation failure.'
        $run.Result.ErrorType | Should -Be 'System.InvalidOperationException'
        $run.Result.OriginalExceptionPreserved | Should -BeTrue
        $run.Stderr | Should -Match '(?i)diagnostic|log'
        $run.Stderr | Should -Match '(?i)unavailable|could not|cannot|failed'
    }

    It 'A019/A021 leaves diagnostic and archive roots absent during normal WhatIf' {
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action Preview -FeedPath '/feeds/single.xml')
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        Test-Path -LiteralPath $run.Root | Should -BeFalse -Because ($run.Result.RootEntries -join ', ')
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'A019/A021 leaves the legacy preview tree unchanged without startup logs' {
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action LegacyPreview -FeedPath '/feeds/single.xml')
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        Test-Path -LiteralPath (Join-Path $run.Root 'local') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $run.Root 'temp') | Should -BeFalse -Because ($run.Result.RootEntries -join ', ')
        $entries = @(Get-ChildItem -LiteralPath (Join-Path $run.Root 'output') -Recurse -Force)
        $entries.Count | Should -Be 1
        $entries[0].Name | Should -Be 'Synthetic legacy show'
        $entries[0].PSIsContainer | Should -BeTrue
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'A020 keeps real verbose and engine-formatted error output free of URL secrets' {
        $secretPath = '/missing/FAKE_PATH_TOKEN?token=FAKE_QUERY_TOKEN&signature=FAKE_SIGNATURE#FAKE_FRAGMENT'
        $run = Receive-UpdDiagnosticWorker (Start-UpdDiagnosticWorker -Context $context -Action RawFailure -FeedPath $secretPath)
        $run.ExitCode | Should -Not -Be 0
        $shareable = $run.Stdout + $run.Stderr + ((Read-UpdDiagnosticLog -Root $run.Root).Text -join "`n")
        $shareable | Should -Not -Match 'FAKE_PATH_TOKEN|FAKE_QUERY_TOKEN|FAKE_SIGNATURE|FAKE_FRAGMENT'
        $shareable | Should -Not -Match ([regex]::Escape($context.BaseUrl + $secretPath))
        $shareable | Should -Match '127\.0\.0\.1'
    }
}
