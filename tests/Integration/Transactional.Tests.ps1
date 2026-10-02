BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    . (Join-Path $repositoryRoot 'UniversalPodcastDownloader.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for the owned loopback integration fixture server.' }
    $script:pythonPath = $python.Source
    $media = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
    $script:mediaHash = (Get-FileHash -LiteralPath $media -Algorithm SHA256).Hash
    $script:mediaLength = (Get-Item -LiteralPath $media).Length
}

Describe 'Transactional completion against real loopback responses' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:pythonPath
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'A012 rejects <Kind> without publishing a final media file' -TestCases @(
        @{ Kind = 'empty' }
        @{ Kind = 'html' }
        @{ Kind = 'json' }
        @{ Kind = 'xml' }
        @{ Kind = 'partial' }
    ) {
        param($Kind)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath ('/feeds/transaction-' + $Kind + '.xml')
        $run.Stdout | Should -Match 'Downloaded\s+: 0'
        $run.Stdout | Should -Match 'Failed\s+: 1'
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Result.RemainingTemporaryCount | Should -Be 0
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '.upd-*.tmp' -Force).Count | Should -Be 0
    }

    It 'A012 accepts valid <Kind> audio and preserves the exact original bytes' -TestCases @(
        @{ Kind = 'no-length'; Route = '/media/no-length.mp3' }
        @{ Kind = 'octet'; Route = '/media/octet-stream' }
        @{ Kind = 'enclosure-mismatch'; Route = '/media/ok.mp3' }
    ) {
        param($Kind, $Route)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath ('/feeds/transaction-' + $Kind + '.xml')
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.ExitCode | Should -Be 0
        $run.Stdout | Should -Match 'Downloaded\s+: 1'
        $run.Stdout | Should -Match 'Failed\s+: 0'
        $files = @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be 1
        $files[0].Length | Should -Be $script:mediaLength
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be $script:mediaHash
        $run.Result.RemainingTemporaryCount | Should -Be 0
        (Get-UpdFixtureState -Context $context).$Route | Should -Be 1
        if ($Kind -eq 'enclosure-mismatch') { $run.Stdout | Should -Match 'Publisher enclosure length differs.*advisory' }
    }

    It 'A011 A012 rejects a truncated HTTP body and downloads complete bytes on a safe restart' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/transaction-recover.xml'
        $first.Result.Succeeded | Should -BeFalse
        $first.ExitCode | Should -Be 1
        $first.Stdout | Should -Match 'Downloaded\s+: 0'
        $first.Stdout | Should -Match 'Failed\s+: 1'
        $first.Result.RemainingTemporaryCount | Should -Be 0
        @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        $requestsBefore = (Get-UpdFixtureState -Context $context).'/media/recover.mp3'
        $requestsBefore | Should -BeGreaterThan 0
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/transaction-recover.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.Stdout | Should -Match 'Downloaded\s+: 1'
        $second.Stdout | Should -Match 'Skipped\s+: 0'
        $files = @(Get-ChildItem -LiteralPath $second.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be 1
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be $script:mediaHash
        (Get-UpdFixtureState -Context $context).'/media/recover.mp3' | Should -Be ($requestsBefore + 1)
    }

    It 'A011 leaves only an owned sibling partial when the actual transfer process is interrupted' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/transaction-interrupt.xml' -InterruptOnPartial
        $first.Interrupted | Should -BeTrue -Because ($first.Stdout + $first.Stderr)
        $first.Result | Should -BeNullOrEmpty
        $partials = @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -File -Filter '.upd-*.tmp' -Force)
        $partials.Count | Should -Be 1
        $partials[0].Name | Should -Match '^\.upd-[a-f0-9]{32}\.tmp$'
        $partials[0].Length | Should -BeGreaterThan 0
        $partials[0].Length | Should -BeLessThan ($script:mediaLength * 64)
        @(Get-ChildItem -LiteralPath $first.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        $partialHash = (Get-FileHash -LiteralPath $partials[0].FullName -Algorithm SHA256).Hash
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/transaction-interrupt.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.Stdout | Should -Match 'Downloaded\s+: 1'
        $second.Stdout | Should -Match 'Skipped\s+: 0'
        $files = @(Get-ChildItem -LiteralPath $second.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be 1
        $files[0].DirectoryName | Should -Be $partials[0].DirectoryName
        $files[0].Length | Should -Be ($script:mediaLength * 64)
        $bytes = [IO.File]::ReadAllBytes($files[0].FullName)
        $original = [IO.File]::ReadAllBytes((Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
        for ($offset = 0; $offset -lt $bytes.Length; $offset += $original.Length) {
            [Convert]::ToBase64String($bytes, $offset, $original.Length) | Should -Be ([Convert]::ToBase64String($original))
        }
        (Get-FileHash -LiteralPath $partials[0].FullName -Algorithm SHA256).Hash | Should -Be $partialHash
        (Get-UpdFixtureState -Context $context).'/media/interrupt.mp3' | Should -Be 2
    }

    It 'A011 A013 recovers from <Stage> without counting a partial as complete' -TestCases @(
        @{ Stage = 'BeforeFinalizeCrash'; ExpectedDownloads = 1; ExpectedSkipped = 0; ExpectedRequests = 2 }
        @{ Stage = 'AfterFinalizeCrash'; ExpectedDownloads = 0; ExpectedSkipped = 1; ExpectedRequests = 1 }
    ) {
        param($Stage, $ExpectedDownloads, $ExpectedSkipped, $ExpectedRequests)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/transaction-crash.xml' -TransactionHook $Stage
        $first.Result | Should -BeNullOrEmpty
        $first.HookMarker.Hook | Should -Be $Stage
        [IO.Path]::GetDirectoryName($first.HookMarker.Temporary) | Should -Be ([IO.Path]::GetDirectoryName($first.HookMarker.Destination))
        [IO.Path]::GetPathRoot($first.HookMarker.Temporary) | Should -Be ([IO.Path]::GetPathRoot($first.HookMarker.Destination))
        if ($Stage -eq 'BeforeFinalizeCrash') {
            $first.HookMarker.TemporaryExclusiveOpen | Should -BeTrue
            Test-Path -LiteralPath $first.HookMarker.Temporary | Should -BeTrue
            Test-Path -LiteralPath $first.HookMarker.Destination | Should -BeFalse
            (Get-FileHash -LiteralPath $first.HookMarker.Temporary -Algorithm SHA256).Hash | Should -Be $script:mediaHash
        }
        else {
            Test-Path -LiteralPath $first.HookMarker.Temporary | Should -BeFalse
            Test-Path -LiteralPath $first.HookMarker.Destination | Should -BeTrue
        }
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/transaction-crash.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.Stdout | Should -Match ('Downloaded\s+: ' + $ExpectedDownloads)
        $second.Stdout | Should -Match ('Skipped\s+: ' + $ExpectedSkipped)
        (Get-FileHash -LiteralPath $first.HookMarker.Destination -Algorithm SHA256).Hash | Should -Be $script:mediaHash
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be $ExpectedRequests
        if ($Stage -eq 'BeforeFinalizeCrash') {
            (Get-FileHash -LiteralPath $first.HookMarker.Temporary -Algorithm SHA256).Hash | Should -Be $script:mediaHash
        }
    }

    It 'A013 preserves a competing final file appearing immediately before the no-overwrite move' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/transaction-crash.xml' -TransactionHook FinalRace
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 1
        $run.Stdout | Should -Match 'Downloaded\s+: 0'
        $run.Stdout | Should -Match 'Failed\s+: 1'
        $run.HookMarker.TemporaryExclusiveOpen | Should -BeTrue
        [IO.File]::ReadAllText($run.HookMarker.Destination) | Should -Be 'Synthetic concurrent final file. Preserve these exact bytes.'
        Test-Path -LiteralPath $run.HookMarker.Temporary | Should -BeFalse
        $run.Result.RemainingTemporaryCount | Should -Be 0
        $probe = [IO.File]::Open($run.HookMarker.Destination, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $probe.Dispose()
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
    }
}
