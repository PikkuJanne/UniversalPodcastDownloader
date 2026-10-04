BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned resume loopback fixtures.' }
    $script:resumePython = $python.Source
    $script:resumeAudio = [IO.File]::ReadAllBytes((Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
    $script:resumeLength = $script:resumeAudio.Length * 64

    function Invoke-UpdResumeWorker {
        param([Parameter(Mandatory)]$Context, [string]$Scenario = 'valid', [switch]$Interrupt,
            [ValidateRange(1, 10)][int]$MaxAttempts = 1)
        if ($Scenario -notmatch '^[a-z0-9-]+$') { throw 'Only named resume fixtures are allowed.' }
        $token = [guid]::NewGuid().ToString('N')
        $configPath = Join-Path $Context.Root ($token + '-input.json')
        $config = @{
            ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'
            FeedUrl = $Context.BaseUrl + '/resume/feed/' + $Scenario
            OutputPath = Join-Path $Context.Root 'output'; Interrupt = [bool]$Interrupt
            MaxAttempts = $MaxAttempts
            MarkerPath = Join-Path $Context.Root ($token + '-marker.json')
            ResultPath = Join-Path $Context.Root ($token + '-result.json')
        }
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $owned = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engine) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-ResumeWorker.ps1'), '-ConfigPath', $configPath
        )
        if (-not $owned.Process.HasExited -and -not $owned.Process.WaitForExit(30000)) {
            $owned.Process.Kill()
            $owned.Process.WaitForExit()
            throw 'Owned resume test worker exceeded 30 seconds.'
        }
        $stdout = $owned.Output.Result
        $stderr = $owned.ErrorOutput.Result
        $marker = if ([IO.File]::Exists($config.MarkerPath)) { [IO.File]::ReadAllText($config.MarkerPath) | ConvertFrom-Json } else { $null }
        $result = if ([IO.File]::Exists($config.ResultPath)) { [IO.File]::ReadAllText($config.ResultPath) | ConvertFrom-Json } else { $null }
        if (-not $result -and -not $marker) { throw "Resume child produced no result or crash marker: stdout=$stdout; stderr=$stderr" }
        [pscustomobject]@{
            Result = $result; Marker = $marker; ExitCode = $owned.Process.ExitCode
            OutputPath = $config.OutputPath; Stdout = $stdout; Stderr = $stderr
        }
    }

    function Get-UpdResumeSidecar {
        param([Parameter(Mandatory)]$Run)
        return @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -Force -File | Where-Object { $_.Name -match '^resume-[a-f0-9]{32}\.json$' })
    }

    function Assert-UpdResumeCompleted {
        param([Parameter(Mandatory)]$Run, [switch]$Changed)
        $Run.ExitCode | Should -Be 0 -Because ($Run.Stdout + $Run.Stderr + $Run.Result.ErrorMessage)
        $Run.Result.Succeeded | Should -BeTrue
        $Run.Stdout | Should -Match 'Downloaded\s+: 1'
        $files = @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be 1
        $bytes = [IO.File]::ReadAllBytes($files[0].FullName)
        $expected = New-Object byte[] ($script:resumeLength + $(if ($Changed) { 32 } else { 0 }))
        for ($offset = 0; $offset -lt $script:resumeLength; $offset += $script:resumeAudio.Length) {
            [Array]::Copy($script:resumeAudio, 0, $expected, $offset, $script:resumeAudio.Length)
        }
        $sha = [Security.Cryptography.SHA256]::Create()
        try { [Convert]::ToBase64String($sha.ComputeHash($bytes)) | Should -Be ([Convert]::ToBase64String($sha.ComputeHash($expected))) }
        finally { $sha.Dispose() }
        $stateFiles = @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -Force -File -Filter state.json)
        $stateFiles.Count | Should -Be 1
        $state = [IO.File]::ReadAllText($stateFiles[0].FullName) | ConvertFrom-Json
        @($state.episodes).Count | Should -Be 1
        $state.episodes[0].status | Should -Be 'transfer_verified'
        $state.episodes[0].bytes | Should -Be $expected.Length
    }

    function Start-UpdResumeInterruption {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test helper starts and interrupts only a tracked child inside its marked temporary context.')]
        [CmdletBinding()]
        param([Parameter(Mandatory)]$Context, [string]$Scenario = 'valid')
        $run = Invoke-UpdResumeWorker -Context $Context -Scenario $Scenario -Interrupt
        $run.Result | Should -BeNullOrEmpty -Because ($run.Stdout + $run.Stderr)
        $run.Marker | Should -Not -BeNullOrEmpty
        $run.Marker.Bytes | Should -BeGreaterThan 0
        $run.Marker.Bytes | Should -BeLessThan $script:resumeLength
        Test-Path -LiteralPath $run.Marker.PartialPath | Should -BeTrue
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        $null = Invoke-WebRequest -Uri ($Context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        return $run
    }

    function Assert-UpdResumePartialPreserved {
        param([Parameter(Mandatory)]$First, [Parameter(Mandatory)][string]$PartialHash)
        Test-Path -LiteralPath $First.Marker.PartialPath | Should -BeTrue
        (Get-FileHash -LiteralPath $First.Marker.PartialPath -Algorithm SHA256).Hash | Should -Be $PartialHash
    }

    function Assert-UpdResumeIncomplete {
        param([Parameter(Mandatory)]$Run)
        $Run.ExitCode | Should -Be 2 -Because ($Run.Stdout + $Run.Stderr + $Run.Result.ErrorMessage)
        $Run.Result.Succeeded | Should -BeFalse
        $Run.Stdout | Should -Match 'Downloaded\s+: 0'
        @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        foreach ($file in @(Get-ChildItem -LiteralPath $Run.OutputPath -Recurse -Force -File -Filter state.json)) {
            $state = [IO.File]::ReadAllText($file.FullName) | ConvertFrom-Json
            @($state.episodes | Where-Object { $_.status -eq 'transfer_verified' -or $_.completed_utc }).Count | Should -Be 0
        }
    }
}

Describe 'A027/A028/A029: safe resume against real loopback responses' -Tag 'Integration', 'A027', 'A028', 'A029' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:resumePython
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'A027 resumes a killed transfer using its durable identity and exact range' {
        $first = Start-UpdResumeInterruption -Context $context
        $sidecars = @(Get-UpdResumeSidecar -Run $first)
        $sidecars.Count | Should -Be 1
        $sidecar = [IO.File]::ReadAllText($sidecars[0].FullName) | ConvertFrom-Json
        $sidecar.schema_version | Should -Be 1
        $sidecar.offset | Should -Be $first.Marker.Bytes
        $sidecar.total_length | Should -Be $script:resumeLength
        $sidecar.etag | Should -Be '"resume-v1"'
        $sidecar.prefix_sha256 | Should -Be (Get-FileHash -LiteralPath $first.Marker.PartialPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $sidecar.partial_name | Should -Be ([IO.Path]::GetFileName($first.Marker.PartialPath))
        $second = Invoke-UpdResumeWorker -Context $context
        Assert-UpdResumeCompleted -Run $second
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 2
        $events[0].has_range | Should -BeFalse
        $events[1].range | Should -Be ('bytes=' + $first.Marker.Bytes + '-')
        $events[1].if_range | Should -Be '"resume-v1"'
        Test-Path -LiteralPath $first.Marker.PartialPath | Should -BeFalse
        @(Get-UpdResumeSidecar -Run $second).Count | Should -Be 0
    }

    It 'A028 starts fresh for <Scenario> and preserves the unowned crash partial' -ForEach @(
        @{ Scenario = 'weak' }; @{ Scenario = 'absent' }; @{ Scenario = 'last-modified' }; @{ Scenario = 'no-length' }
    ) {
        $first = Start-UpdResumeInterruption -Context $context -Scenario $Scenario
        @(Get-UpdResumeSidecar -Run $first).Count | Should -Be 0
        $partialHash = (Get-FileHash -LiteralPath $first.Marker.PartialPath -Algorithm SHA256).Hash
        $second = Invoke-UpdResumeWorker -Context $context -Scenario $Scenario
        Assert-UpdResumeCompleted -Run $second
        Assert-UpdResumePartialPreserved -First $first -PartialHash $partialHash
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 2
        $events[1].has_range | Should -BeFalse
        $events[1].has_if_range | Should -BeFalse
    }

    It 'A028 rejects <Scenario> before append and makes one fresh full request' -ForEach @(
        @{ Scenario = 'ignore-range' }; @{ Scenario = 'changed' }
        @{ Scenario = 'bad-start' }; @{ Scenario = 'bad-end' }; @{ Scenario = 'bad-total' }; @{ Scenario = 'bad-length' }
        @{ Scenario = 'bad-validator' }; @{ Scenario = 'missing-validator' }; @{ Scenario = 'weak-validator' }
        @{ Scenario = 'bad-type' }; @{ Scenario = 'bad-encoding' }; @{ Scenario = 'multipart' }
        @{ Scenario = 'range-416' }; @{ Scenario = 'range-416-local' }
    ) {
        $first = Start-UpdResumeInterruption -Context $context -Scenario $Scenario
        $sidecars = @(Get-UpdResumeSidecar -Run $first)
        $sidecars.Count | Should -Be 1
        $sidecarHash = (Get-FileHash -LiteralPath $sidecars[0].FullName -Algorithm SHA256).Hash
        $partialHash = (Get-FileHash -LiteralPath $first.Marker.PartialPath -Algorithm SHA256).Hash
        $second = Invoke-UpdResumeWorker -Context $context -Scenario $Scenario
        Assert-UpdResumeCompleted -Run $second -Changed:($Scenario -eq 'changed')
        Assert-UpdResumePartialPreserved -First $first -PartialHash $partialHash
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 3
        $events[1].range | Should -Be ('bytes=' + $first.Marker.Bytes + '-')
        $events[1].if_range | Should -Be '"resume-v1"'
        $events[2].has_range | Should -BeFalse
        $events[2].has_if_range | Should -BeFalse
        $saved = @(Get-ChildItem -LiteralPath $second.OutputPath -Recurse -Force -File -Filter 'resume-*.old.json')
        @($saved | Where-Object { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash -eq $sidecarHash }).Count | Should -Be 1
        $logs = @(Get-ChildItem -LiteralPath $second.OutputPath -Recurse -File -Filter 'download-*.log' | ForEach-Object { [IO.File]::ReadAllText($_.FullName) })
        ($logs -join "`n") | Should -Match 'Resume response requires a fresh transfer; previous partial preserved\.'
    }

    It 'A029 preserves a prior signed-URL partial and starts fresh for a renewed exact URL' {
        $first = Start-UpdResumeInterruption -Context $context -Scenario signed
        $partialHash = (Get-FileHash -LiteralPath $first.Marker.PartialPath -Algorithm SHA256).Hash
        $second = Invoke-UpdResumeWorker -Context $context -Scenario signed
        Assert-UpdResumeCompleted -Run $second
        Assert-UpdResumePartialPreserved -First $first -PartialHash $partialHash
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 2
        $events[0].query_variant | Should -Be 'original'
        $events[1].query_variant | Should -Be 'renewed'
        $events[1].has_range | Should -BeFalse
        $events[1].has_if_range | Should -BeFalse
    }

    It 'A029 never declares an always-416 response complete or deletes a prior partial' {
        $first = Start-UpdResumeInterruption -Context $context -Scenario 'always-416'
        $partialHash = (Get-FileHash -LiteralPath $first.Marker.PartialPath -Algorithm SHA256).Hash
        # An unclaimed protocol-shaped partial is ignored by legacy inventory;
        # a general .part file correctly requests legacy review before transfer.
        $unknown = Join-Path ([IO.Path]::GetDirectoryName($first.Marker.PartialPath)) '.upd-ffffffffffffffffffffffffffffffff.tmp'
        [IO.File]::WriteAllText($unknown, 'Synthetic unrelated partial; preserve these bytes.')
        $unknownHash = (Get-FileHash -LiteralPath $unknown -Algorithm SHA256).Hash
        $second = Invoke-UpdResumeWorker -Context $context -Scenario 'always-416'
        Assert-UpdResumeIncomplete -Run $second
        Assert-UpdResumePartialPreserved -First $first -PartialHash $partialHash
        (Get-FileHash -LiteralPath $unknown -Algorithm SHA256).Hash | Should -Be $unknownHash
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 3
        $events[1].has_range | Should -BeTrue
        $events[2].has_range | Should -BeFalse
    }

    It 'A029 rejects <Damage> recovery data before any media request and preserves it' -ForEach @(
        @{ Damage = 'schema' }; @{ Damage = 'offset' }; @{ Damage = 'prefix-hash' }
        @{ Damage = 'episode' }; @{ Damage = 'feed' }; @{ Damage = 'relative-path' }
        @{ Damage = 'same-size-change' }; @{ Damage = 'shorter-file' }; @{ Damage = 'uncheckpointed-tail' }; @{ Damage = 'malformed-json' }
    ) {
        $first = Start-UpdResumeInterruption -Context $context
        $sidecars = @(Get-UpdResumeSidecar -Run $first)
        $sidecars.Count | Should -Be 1
        $sidecar = [IO.File]::ReadAllText($sidecars[0].FullName) | ConvertFrom-Json
        switch ($Damage) {
            schema { $sidecar.schema_version = 99 }
            offset { $sidecar.offset = [long]$sidecar.offset + 1 }
            prefix-hash { $sidecar.prefix_sha256 = 'a' * 64 }
            episode { $sidecar.episode_id = 'b' * 64 }
            feed { $sidecar.feed_id = 'c' * 64 }
            relative-path { $sidecar.relative_path = 'another-episode.mp3' }
            same-size-change {
                $bytes = [IO.File]::ReadAllBytes($first.Marker.PartialPath)
                $bytes[0] = $bytes[0] -bxor 1
                [IO.File]::WriteAllBytes($first.Marker.PartialPath, $bytes)
            }
            shorter-file {
                $file = [IO.File]::Open($first.Marker.PartialPath, [IO.FileMode]::Open, [IO.FileAccess]::Write, [IO.FileShare]::None)
                try { $file.SetLength($file.Length - 1) } finally { $file.Dispose() }
            }
            uncheckpointed-tail {
                $file = [IO.File]::Open($first.Marker.PartialPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::None)
                try { $file.WriteByte(42) } finally { $file.Dispose() }
            }
        }
        $sidecarText = if ($Damage -eq 'malformed-json') { '{"schema_version":' } else { $sidecar | ConvertTo-Json -Depth 8 }
        [IO.File]::WriteAllText($sidecars[0].FullName, $sidecarText, [Text.UTF8Encoding]::new($false))
        $sidecarHash = (Get-FileHash -LiteralPath $sidecars[0].FullName -Algorithm SHA256).Hash
        $partialHash = (Get-FileHash -LiteralPath $first.Marker.PartialPath -Algorithm SHA256).Hash
        $second = Invoke-UpdResumeWorker -Context $context
        Assert-UpdResumeIncomplete -Run $second
        Assert-UpdResumePartialPreserved -First $first -PartialHash $partialHash
        (Get-FileHash -LiteralPath $sidecars[0].FullName -Algorithm SHA256).Hash | Should -Be $sidecarHash
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 1
    }

    It 'A029 restarts an already full partial without treating it as completed evidence' {
        $first = Start-UpdResumeInterruption -Context $context
        $sidecars = @(Get-UpdResumeSidecar -Run $first)
        $sidecars.Count | Should -Be 1
        $full = New-Object byte[] $script:resumeLength
        for ($offset = 0; $offset -lt $full.Length; $offset += $script:resumeAudio.Length) {
            [Array]::Copy($script:resumeAudio, 0, $full, $offset, $script:resumeAudio.Length)
        }
        [IO.File]::WriteAllBytes($first.Marker.PartialPath, $full)
        $sidecar = [IO.File]::ReadAllText($sidecars[0].FullName) | ConvertFrom-Json
        $partialHash = (Get-FileHash -LiteralPath $first.Marker.PartialPath -Algorithm SHA256).Hash
        $sidecar.offset = $full.Length
        $sidecar.prefix_sha256 = $partialHash.ToLowerInvariant()
        [IO.File]::WriteAllText($sidecars[0].FullName, ($sidecar | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
        $second = Invoke-UpdResumeWorker -Context $context
        Assert-UpdResumeCompleted -Run $second
        Assert-UpdResumePartialPreserved -First $first -PartialHash $partialHash
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 2
        $events[1].has_range | Should -BeFalse
    }

    It 'A027 advances the durable offset after a caught truncated 206 and retries within the configured limit' {
        $first = Start-UpdResumeInterruption -Context $context -Scenario 'range-truncated-once'
        $second = Invoke-UpdResumeWorker -Context $context -Scenario 'range-truncated-once' -MaxAttempts 2
        Assert-UpdResumeCompleted -Run $second
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 3
        $events[1].range | Should -Be ('bytes=' + $first.Marker.Bytes + '-')
        $events[2].range | Should -Match '^bytes=[0-9]+-$'
        [long]($events[2].range.Substring(6).TrimEnd('-')) | Should -BeGreaterThan $first.Marker.Bytes
        $events[1].if_range | Should -Be '"resume-v1"'
        $events[2].if_range | Should -Be '"resume-v1"'
        Test-Path -LiteralPath $first.Marker.PartialPath | Should -BeFalse
        @(Get-UpdResumeSidecar -Run $second).Count | Should -Be 0
    }

    It 'A027 recovers a caught truncated fresh 200 using a bounded resumed request in the same run' {
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $run = Invoke-UpdResumeWorker -Context $context -Scenario 'fresh-truncated-once' -MaxAttempts 2
        Assert-UpdResumeCompleted -Run $run
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        @($events).Count | Should -Be 2
        $events[0].has_range | Should -BeFalse
        $events[1].range | Should -Be ('bytes=' + [long]($script:resumeLength / 3) + '-')
        $events[1].if_range | Should -Be '"resume-v1"'
        @(Get-UpdResumeSidecar -Run $run).Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Force -File -Filter '.upd-*.tmp').Count | Should -Be 0
    }
}
