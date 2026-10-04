BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned input-boundary loopback fixtures.' }
    $script:boundaryPython = $python.Source
    $script:boundaryAudio = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'

    function Invoke-UpdBoundaryWorker {
        param(
            [Parameter(Mandatory)]$Context,
            [ValidateSet('Preview', 'Decline', 'Download', 'Resolve', 'Metadata')][string]$Action,
            [string]$FeedPath = '/feeds/single.xml',
            [string]$SentinelPath
        )

        $token = [guid]::NewGuid().ToString('N')
        $configPath = Join-Path $Context.Root ($token + '-input.json')
        $exportRoot = Join-Path $Context.Root 'export'
        $null = [IO.Directory]::CreateDirectory($exportRoot)
        $config = @{
            Action = $Action; ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'
            Root = Join-Path $Context.Root 'p'; ExportRoot = $exportRoot
            ResultPath = Join-Path $Context.Root ($token + '-result.json')
            FeedUrl = $Context.BaseUrl + $FeedPath; SentinelPath = $SentinelPath
        }
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $owned = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engine) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-InputBoundaryWorker.ps1'), '-ConfigPath', $configPath
        )
        if (-not $owned.Process.HasExited -and -not $owned.Process.WaitForExit(30000)) {
            $owned.Process.Kill(); $owned.Process.WaitForExit()
            throw 'Owned input-boundary child exceeded 30 seconds.'
        }
        $stdout = $owned.Output.Result; $stderr = $owned.ErrorOutput.Result
        if (-not [IO.File]::Exists($config.ResultPath)) {
            throw "Input-boundary child produced no result: stdout=$stdout; stderr=$stderr"
        }
        [pscustomobject]@{
            Result = ([IO.File]::ReadAllText($config.ResultPath) | ConvertFrom-Json)
            ExitCode = $owned.Process.ExitCode; Stdout = $stdout; Stderr = $stderr
            Root = $config.Root; Export = Join-Path $exportRoot 'diagnostics.json'
        }
    }

    function Get-UpdBoundarySnapshot {
        param([Parameter(Mandatory)][string]$Root)
        return @(Get-ChildItem -LiteralPath $Root -Recurse -Force | Sort-Object FullName | ForEach-Object {
            $relative = $_.FullName.Substring($Root.Length)
            if ($_.PSIsContainer) { 'directory|' + $relative }
            else { 'file|' + $relative + '|' + $_.Length + '|' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
        })
    }
}

Describe 'A022/A024: real preview and untrusted-input boundaries' -Tag 'Integration', 'A022', 'A024' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:boundaryPython
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'A022 <Action> allows only feed reads and a later real run downloads verified bytes' -ForEach @(
        @{ Action = 'Preview' }
        @{ Action = 'Decline' }
    ) {
        $preview = Invoke-UpdBoundaryWorker -Context $context -Action $Action
        $preview.ExitCode | Should -Be 0 -Because ($preview.Stdout + $preview.Stderr + $preview.Result.ErrorMessage)
        if ($Action -eq 'Decline') { $preview.Result.ConfirmationCount | Should -Be 1 }
        Test-Path -LiteralPath $preview.Root | Should -BeFalse -Because ($preview.Result.RootEntries -join ', ')
        Test-Path -LiteralPath $preview.Export | Should -BeFalse
        $before = Get-UpdFixtureState -Context $context
        $before.'/feeds/single.xml' | Should -Be 1
        $before.'/media/ok.mp3' | Should -BeNullOrEmpty
        $download = Invoke-UpdBoundaryWorker -Context $context -Action Download
        $download.ExitCode | Should -Be 0 -Because ($download.Stdout + $download.Stderr + $download.Result.ErrorMessage)
        $files = @(Get-ChildItem -LiteralPath (Join-Path $download.Root 'out') -Recurse -Filter '*.mp3' -File)
        $files.Count | Should -Be 1
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $script:boundaryAudio -Algorithm SHA256).Hash
        $statePath = @(Get-ChildItem -LiteralPath (Join-Path $download.Root 'out') -Recurse -Force -Filter state.json -File)
        $statePath.Count | Should -Be 1
        ([IO.File]::ReadAllText($statePath[0].FullName) | ConvertFrom-Json).episodes[0].status | Should -Be 'transfer_verified'
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
        $repeat = Invoke-UpdBoundaryWorker -Context $context -Action Download
        $repeat.ExitCode | Should -Be 0
        $repeat.Stdout | Should -Match 'Skipping \(verified history\)'
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
    }

    It 'A022 preview preserves every existing archive and diagnostic byte' {
        $first = Invoke-UpdBoundaryWorker -Context $context -Action Download
        $first.ExitCode | Should -Be 0 -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $snapshot = Get-UpdBoundarySnapshot -Root $first.Root
        $exportHash = (Get-FileHash -LiteralPath $first.Export -Algorithm SHA256).Hash
        $preview = Invoke-UpdBoundaryWorker -Context $context -Action Preview
        $preview.ExitCode | Should -Be 0
        (Get-UpdBoundarySnapshot -Root $preview.Root) | Should -Be $snapshot
        (Get-FileHash -LiteralPath $preview.Export -Algorithm SHA256).Hash | Should -Be $exportHash
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
    }

    It 'A024 rejects bounded hostile XML <FeedPath> without entity or media requests' -ForEach @(
        @{ FeedPath = '/feeds/external-http.xml' }
        @{ FeedPath = '/feeds/internal-dtd.xml' }
        @{ FeedPath = '/feeds/malformed.xml' }
        @{ FeedPath = '/feeds/deep.xml' }
        @{ FeedPath = '/feeds/many-nodes.xml' }
    ) {
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Resolve -FeedPath $FeedPath
        $run.ExitCode | Should -Be 1
        $run.Result.ItemCount | Should -Be 0
        Test-Path -LiteralPath $run.Root | Should -BeFalse -Because ($run.Result.RootEntries -join ', ')
        $stats = Get-UpdFixtureState -Context $context
        $stats.$FeedPath | Should -Be 1
        $stats.'/entity/never' | Should -BeNullOrEmpty
        $stats.'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'A024 rejects a file entity while the owned sentinel cannot be read' {
        $sentinel = Join-Path $context.Root 'entity-sentinel.txt'
        [IO.File]::WriteAllText($sentinel, 'LOCAL_ENTITY_MUST_NOT_APPEAR')
        $uri = [uri]$sentinel
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Resolve -SentinelPath $sentinel `
            -FeedPath ('/feeds/external-file.xml?file=' + [uri]::EscapeDataString($uri.AbsoluteUri))
        $run.ExitCode | Should -Be 1
        $run.Result.ItemCount | Should -Be 0
        ($run.Stdout + $run.Stderr + ($run.Result | ConvertTo-Json -Depth 12)) | Should -Not -Match 'LOCAL_ENTITY_MUST_NOT_APPEAR'
        [IO.File]::ReadAllText($sentinel) | Should -Be 'LOCAL_ENTITY_MUST_NOT_APPEAR'
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
        Test-Path -LiteralPath $run.Root | Should -BeFalse
    }

    It 'A024 rejects oversized metadata <FeedPath> with a finite byte limit' -ForEach @(
        @{ FeedPath = '/feeds/oversized.xml' }
        @{ FeedPath = '/feeds/oversized-no-length.xml' }
    ) {
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Metadata -FeedPath $FeedPath
        $run.ExitCode | Should -Be 1
        $run.Result.ErrorMessage | Should -Match '(?i)limit|exceed|large'
        $run.Result.Bytes | Should -Be 0
        Test-Path -LiteralPath $run.Root | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'A024 preserves valid feed parsing for <FeedPath>' -ForEach @(
        @{ FeedPath = '/feeds/single.xml'; Count = 1 }
        @{ FeedPath = '/feeds/atom.xml'; Count = 2 }
    ) {
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Resolve -FeedPath $FeedPath
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.ItemCount | Should -Be $Count
        Test-Path -LiteralPath $run.Root | Should -BeFalse
    }

    It 'A023 permits a relative redirect and private loopback feed reads' {
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Resolve -FeedPath '/redirect/feed'
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.ItemCount | Should -Be 1
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/redirect/feed' | Should -Be 1
        $stats.'/feeds/single.xml' | Should -Be 1
        $stats.'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'A023 rejects redirect targets at <FeedPath> before an unsafe follow-up request' -ForEach @(
        @{ FeedPath = '/redirect/file' }
        @{ FeedPath = '/redirect/userinfo' }
        @{ FeedPath = '/redirect/loop' }
    ) {
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Resolve -FeedPath $FeedPath
        $run.ExitCode | Should -Be 1
        $stats = Get-UpdFixtureState -Context $context
        $stats.$FeedPath | Should -BeLessOrEqual 6
        $stats.'/feeds/single.xml' | Should -BeNullOrEmpty
        $stats.'/media/ok.mp3' | Should -BeNullOrEmpty
        Test-Path -LiteralPath $run.Root | Should -BeFalse
    }

    It 'A023 does not carry response cookies across a hostname redirect' {
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Resolve -FeedPath '/redirect/cookie'
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/no-cookie.xml' | Should -Be 1
        $stats.'/credential-received' | Should -BeNullOrEmpty
    }

    It 'A023 applies the same redirect policy to media and preserves exact bytes' {
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Download -FeedPath '/feeds/redirect-media.xml'
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $files = @(Get-ChildItem -LiteralPath (Join-Path $run.Root 'out') -Recurse -Filter '*.mp3' -File)
        $files.Count | Should -Be 1
        (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $script:boundaryAudio -Algorithm SHA256).Hash
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/media/redirect.mp3' | Should -Be 1
        $stats.'/media/ok.mp3' | Should -Be 1
    }

    It 'A023 refuses a forbidden media redirect without completing an episode' {
        $run = Invoke-UpdBoundaryWorker -Context $context -Action Download -FeedPath '/feeds/redirect-file-media.xml'
        $run.ExitCode | Should -Be 2
        @(Get-ChildItem -LiteralPath (Join-Path $run.Root 'out') -Recurse -Filter '*.mp3' -File).Count | Should -Be 0
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/media/redirect-file.mp3' | Should -Be 1
        $stats.'/media/ok.mp3' | Should -BeNullOrEmpty
    }
}
