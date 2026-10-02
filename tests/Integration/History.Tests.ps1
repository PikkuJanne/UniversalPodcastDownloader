BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    . (Join-Path $repositoryRoot 'UniversalPodcastDownloader.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for the owned loopback history integration fixtures.' }
    $script:historyPython = $python.Source
    $sample = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
    $script:historyMediaHash = (Get-FileHash -LiteralPath $sample -Algorithm SHA256).Hash
    $script:historyMediaLength = (Get-Item -LiteralPath $sample).Length

    function Get-UpdRecordedArchive {
        param([Parameter(Mandatory)][string]$OutputPath)
        $states = @(Get-ChildItem -LiteralPath $OutputPath -Recurse -Force -File -Filter 'state.json')
        $states.Count | Should -Be 1
        $history = Get-Content -LiteralPath $states[0].FullName -Raw | ConvertFrom-Json
        $archiveRoot = Split-Path $states[0].DirectoryName -Parent
        [pscustomobject]@{
            Root = $archiveRoot
            Path = $states[0].FullName
            BackupPath = $states[0].FullName + '.bak'
            State = $history
            MediaPath = if (@($history.episodes).Count -eq 1) { Join-Path $archiveRoot $history.episodes[0].relative_path } else { $null }
        }
    }
}

Describe 'Durable history against actual processes and loopback transfers' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:historyPython
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'A015 records versioned completion evidence and retains the previous valid generation' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $run.OutputPath
        $archive.State.schema_version | Should -Be 1
        $archive.State.feed_id | Should -Match '^[a-f0-9]{64}$'
        @($archive.State.feed_alias_fingerprints).Count | Should -BeGreaterThan 0
        @($archive.State.episodes).Count | Should -Be 1
        $record = $archive.State.episodes[0]
        $record.episode_id | Should -Match '^[a-f0-9]{64}$'
        $record.identity_source | Should -Be 'rss-guid'
        $record.identity_fingerprint | Should -Match '^[a-f0-9]{64}$'
        [IO.Path]::IsPathRooted($record.relative_path) | Should -BeFalse
        $record.status | Should -Be 'transfer_verified'
        $record.bytes | Should -Be $script:historyMediaLength
        $record.local_sha256 | Should -Be $script:historyMediaHash
        $record.completed_utc | Should -Not -BeNullOrEmpty
        $record.verification.method | Should -Not -BeNullOrEmpty
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        Test-Path -LiteralPath $archive.BackupPath | Should -BeTrue
        $backup = Get-Content -LiteralPath $archive.BackupPath -Raw | ConvertFrom-Json
        $backup.schema_version | Should -Be 1
        $backup.feed_id | Should -Be $archive.State.feed_id
        $backup.generation | Should -Be ($archive.State.generation - 1)
        $serialized = Get-Content -LiteralPath $archive.Path -Raw
        $serialized | Should -Not -Match 'http://|signature=|history-stable-001'
    }

    It 'A016 verifies recorded bytes on repeat without a second media request or rewrite' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        $stamp = (Get-Item -LiteralPath $archive.MediaPath).LastWriteTimeUtc
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.Stdout | Should -Match 'Skipping \(verified history\)'
        $second.Stdout | Should -Match 'Downloaded\s+: 0'
        $second.Stdout | Should -Match 'Skipped\s+: 1'
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
        (Get-Item -LiteralPath $archive.MediaPath).LastWriteTimeUtc | Should -Be $stamp
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
    }

    It 'A015 A016 keeps the recorded folder and path after titles and signed media URL change' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.Stdout | Should -Match 'Renamed history show'
        $second.Stdout | Should -Match 'Skipped\s+: 1'
        $recorded = Get-UpdRecordedArchive -OutputPath $second.OutputPath
        $recorded.Path | Should -Be $archive.Path
        $recorded.MediaPath | Should -Be $archive.MediaPath
        $recorded.State.episodes[0].episode_id | Should -Be $archive.State.episodes[0].episode_id
        @(Get-ChildItem -LiteralPath $second.OutputPath -Directory).Count | Should -Be 1
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
        # A fresh download must use the exact renewed URL, including both query
        # values. The fixture returns 403 for any changed or stale signature.
        [IO.File]::Delete($archive.MediaPath)
        $third = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $third.Result.Succeeded | Should -BeTrue -Because ($third.Stdout + $third.Stderr + $third.Result.ErrorMessage)
        $third.Stdout | Should -Match 'Downloaded\s+: 1'
        (Get-UpdRecordedArchive -OutputPath $third.OutputPath).MediaPath | Should -Be $archive.MediaPath
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 2
    }

    It 'A016 redownloads missing completed media to its original recorded destination' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        [IO.File]::Delete($archive.MediaPath)
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.Stdout | Should -Match 'Downloaded\s+: 1'
        $second.Stdout | Should -Match 'Skipped\s+: 0'
        (Get-UpdRecordedArchive -OutputPath $second.OutputPath).MediaPath | Should -Be $archive.MediaPath
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 2
    }

    It 'A016 preserves changed completed media with <Change> and reports an incomplete run' -TestCases @(
        @{ Change = 'different-length' }
        @{ Change = 'same-length-different-hash' }
    ) {
        param($Change)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        $bytes = [IO.File]::ReadAllBytes($archive.MediaPath)
        if ($Change -eq 'different-length') { $bytes = $bytes[0..($bytes.Length - 2)] }
        else { $bytes[$bytes.Length - 1] = $bytes[$bytes.Length - 1] -bxor 1 }
        [IO.File]::WriteAllBytes($archive.MediaPath, $bytes)
        $changedHash = (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeFalse
        $second.ExitCode | Should -Be 1
        $second.Stdout | Should -Match 'Downloaded\s+: 0'
        $second.Stdout | Should -Match 'Skipped\s+: 0'
        $second.Stdout | Should -Match 'Failed\s+: 1'
        ($second.Stdout + $second.Result.ErrorMessage) | Should -Match 'conflict|changed|mismatch'
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $changedHash
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }

    It 'A016 preserves an unknown destination and unrelated media without adopting either' {
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        [IO.File]::Delete($archive.Path)
        [IO.File]::Delete($archive.BackupPath)
        $unrelated = Join-Path $archive.Root 'unrelated-user-recording.mp3'
        [IO.File]::WriteAllText($unrelated, 'Synthetic user-owned recording. Preserve these bytes.')
        $unknownHash = (Get-FileHash -LiteralPath $unrelated -Algorithm SHA256).Hash
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeFalse
        $second.ExitCode | Should -Be 1
        $second.Stdout | Should -Match 'Downloaded\s+: 0'
        $second.Stdout | Should -Match 'Skipped\s+: 0'
        $second.Stdout | Should -Match 'Failed\s+: 1'
        ($second.Stdout + $second.Result.ErrorMessage) | Should -Match 'unverified|unknown|conflict'
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        (Get-FileHash -LiteralPath $unrelated -Algorithm SHA256).Hash | Should -Be $unknownHash
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
        $newState = Get-UpdRecordedArchive -OutputPath $second.OutputPath
        @($newState.State.episodes | Where-Object { $_.status -eq 'transfer_verified' }).Count | Should -Be 0
    }

    It 'A015 preserves <Problem> history and its valid backup instead of resetting it' -TestCases @(
        @{ Problem = 'corrupt-json' }
        @{ Problem = 'newer-schema' }
    ) {
        param($Problem)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        $backupHash = (Get-FileHash -LiteralPath $archive.BackupPath -Algorithm SHA256).Hash
        if ($Problem -eq 'corrupt-json') { [IO.File]::WriteAllText($archive.Path, '{ synthetic broken history') }
        else {
            $archive.State.schema_version = 999
            [IO.File]::WriteAllText($archive.Path, ($archive.State | ConvertTo-Json -Depth 10))
        }
        $damagedHash = (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeFalse
        $second.ExitCode | Should -Be 1
        ($second.Stdout + $second.Result.ErrorMessage) | Should -Match 'state|history|schema'
        (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash | Should -Be $damagedHash
        (Get-FileHash -LiteralPath $archive.BackupPath -Algorithm SHA256).Hash | Should -Be $backupHash
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        @(Get-ChildItem -LiteralPath $second.OutputPath -Directory).Count | Should -Be 1
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }

    It 'A015 preserves an inconsistent backup pair with <Problem> instead of rebuilding history' -TestCases @(
        @{ Problem = 'missing-primary' }
        @{ Problem = 'corrupt-backup' }
    ) {
        param($Problem)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        $primaryHash = (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash
        if ($Problem -eq 'missing-primary') { [IO.File]::Delete($archive.Path) }
        else { [IO.File]::WriteAllText($archive.BackupPath, '{ synthetic broken backup') }
        $backupHash = (Get-FileHash -LiteralPath $archive.BackupPath -Algorithm SHA256).Hash
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeFalse
        $second.ExitCode | Should -Be 1
        ($second.Stdout + $second.Result.ErrorMessage) | Should -Match 'state|history|backup'
        if ($Problem -eq 'missing-primary') { Test-Path -LiteralPath $archive.Path | Should -BeFalse }
        else { (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash | Should -Be $primaryHash }
        (Get-FileHash -LiteralPath $archive.BackupPath -Algorithm SHA256).Hash | Should -Be $backupHash
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }

    It 'A015 retains valid state when the writer process dies <Stage>' -TestCases @(
        @{ Stage = 'BeforeStateReplaceCrash' }
        @{ Stage = 'AfterStateReplaceCrash' }
    ) {
        param($Stage)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        $stateHash = (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash
        $backupHash = (Get-FileHash -LiteralPath $archive.BackupPath -Algorithm SHA256).Hash
        [IO.File]::Delete($archive.MediaPath)
        $crashed = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml' -TransactionHook $Stage
        $crashed.Result | Should -BeNullOrEmpty
        $crashed.HookMarker.Hook | Should -Be $Stage
        $surviving = Get-UpdRecordedArchive -OutputPath $crashed.OutputPath
        if ($Stage -eq 'BeforeStateReplaceCrash') {
            (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash | Should -Be $stateHash
            (Get-FileHash -LiteralPath $archive.BackupPath -Algorithm SHA256).Hash | Should -Be $backupHash
        }
        else {
            $surviving.State.generation | Should -Be ($archive.State.generation + 1)
            (Get-FileHash -LiteralPath $archive.BackupPath -Algorithm SHA256).Hash | Should -Be $stateHash
        }
        $restarted = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $restarted.Result.Succeeded | Should -BeTrue -Because ($restarted.Stdout + $restarted.Stderr + $restarted.Result.ErrorMessage)
        (Get-UpdRecordedArchive -OutputPath $restarted.OutputPath).State.episodes[0].status | Should -Be 'transfer_verified'
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
    }

    It 'A015 reconciles the persisted prepared digest after <Stage>' -TestCases @(
        @{ Stage = 'BeforeFinalizeCrash'; Downloads = 1; Skips = 0; Requests = 2 }
        @{ Stage = 'AfterFinalizeCrash'; Downloads = 0; Skips = 1; Requests = 1 }
    ) {
        param($Stage, $Downloads, $Skips, $Requests)
        $crashed = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml' -TransactionHook $Stage
        $crashed.Result | Should -BeNullOrEmpty
        $crashed.HookMarker.Hook | Should -Be $Stage
        $prepared = Get-UpdRecordedArchive -OutputPath $crashed.OutputPath
        $prepared.State.episodes[0].status | Should -Be 'prepared'
        $prepared.State.episodes[0].bytes | Should -Be $script:historyMediaLength
        $prepared.State.episodes[0].local_sha256 | Should -Be $script:historyMediaHash
        if ($Stage -eq 'BeforeFinalizeCrash') {
            Test-Path -LiteralPath $prepared.MediaPath | Should -BeFalse
            (Get-FileHash -LiteralPath $crashed.HookMarker.Temporary -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        }
        else { (Get-FileHash -LiteralPath $prepared.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash }
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.Stdout | Should -Match ('Downloaded\s+: ' + $Downloads)
        $second.Stdout | Should -Match ('Skipped\s+: ' + $Skips)
        (Get-UpdRecordedArchive -OutputPath $second.OutputPath).State.episodes[0].status | Should -Be 'transfer_verified'
        (Get-FileHash -LiteralPath $prepared.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be $Requests
        if ($Stage -eq 'BeforeFinalizeCrash') {
            (Get-FileHash -LiteralPath $crashed.HookMarker.Temporary -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        }
    }

    It 'A015 refuses to reconcile a changed final file after the finalize-history crash gap' {
        $crashed = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml' -TransactionHook AfterFinalizeCrash
        $crashed.HookMarker.Hook | Should -Be 'AfterFinalizeCrash'
        $prepared = Get-UpdRecordedArchive -OutputPath $crashed.OutputPath
        $bytes = [IO.File]::ReadAllBytes($prepared.MediaPath)
        $bytes[$bytes.Length - 1] = $bytes[$bytes.Length - 1] -bxor 1
        [IO.File]::WriteAllBytes($prepared.MediaPath, $bytes)
        $changedHash = (Get-FileHash -LiteralPath $prepared.MediaPath -Algorithm SHA256).Hash
        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $second.Result.Succeeded | Should -BeFalse
        $second.Stdout | Should -Match 'Skipped\s+: 0'
        $second.Stdout | Should -Match 'Failed\s+: 1'
        (Get-FileHash -LiteralPath $prepared.MediaPath -Algorithm SHA256).Hash | Should -Be $changedHash
        (Get-UpdRecordedArchive -OutputPath $second.OutputPath).State.episodes[0].status | Should -Not -Be 'transfer_verified'
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }

    It 'A015 excludes a second <Scope> writer and releases the OS lock after the owner process dies' -TestCases @(
        @{ Scope = 'Show' }
        @{ Scope = 'Archive' }
    ) {
        param($Scope)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdRecordedArchive -OutputPath $first.OutputPath
        $stateHash = (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash
        $readyPath = Join-Path $context.Root 'history-lock-ready.txt'
        $engineName = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $lockRoot = if ($Scope -eq 'Archive') { $first.OutputPath } else { $archive.Root }
        $holder = Start-UpdOwnedProcess -Context $context -FilePath (Join-Path $PSHOME $engineName) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', (Join-Path $repositoryRoot 'tests/support/Invoke-HistoryLockWorker.ps1'),
            '-ProductScript', (Join-Path $repositoryRoot 'UniversalPodcastDownloader.ps1'),
            '-ArchiveRoot', $lockRoot, '-ReadyPath', $readyPath, '-Scope', $Scope
        )
        $watch = [Diagnostics.Stopwatch]::StartNew()
        while (-not (Test-Path -LiteralPath $readyPath)) {
            if ($holder.Process.HasExited) { throw ('History lock holder exited: ' + $holder.ErrorOutput.Result) }
            if ($watch.Elapsed.TotalSeconds -gt 10) { throw 'History lock holder readiness timed out.' }
            Start-Sleep -Milliseconds 50
        }
        $blocked = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $blocked.Result.Succeeded | Should -BeFalse
        ($blocked.Stdout + $blocked.Result.ErrorMessage) | Should -Match 'lock|writer|another|use'
        (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash | Should -Be $stateHash
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
        $holder.Process.Kill()
        $holder.Process.WaitForExit()
        $restarted = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/history.xml'
        $restarted.Result.Succeeded | Should -BeTrue -Because ($restarted.Stdout + $restarted.Stderr + $restarted.Result.ErrorMessage)
        $restarted.Stdout | Should -Match 'Skipped\s+: 1'
        (Get-FileHash -LiteralPath $archive.MediaPath -Algorithm SHA256).Hash | Should -Be $script:historyMediaHash
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }
}
