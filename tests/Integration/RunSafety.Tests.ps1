BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned loopback run safety fixtures.' }
    $script:safetyPython = $python.Source
    $script:safetyAudio = [IO.File]::ReadAllBytes((Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
    $script:safetyResumeBytes = New-Object byte[] ($script:safetyAudio.Length * 64)
    for ($offset = 0; $offset -lt $script:safetyResumeBytes.Length; $offset += $script:safetyAudio.Length) {
        [Array]::Copy($script:safetyAudio, 0, $script:safetyResumeBytes, $offset, $script:safetyAudio.Length)
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $script:safetyResumeHash = ([BitConverter]::ToString($sha.ComputeHash($script:safetyResumeBytes))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }

    function Start-UpdSafetyWorker {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Starts only a tracked hidden child and writes its configuration within this marked test root.')]
        [CmdletBinding()]
        param([Parameter(Mandatory)]$Context, [ValidateSet('Run', 'GateCancel', 'HoldArchive')][string]$Action = 'Run',
            [string]$FeedPath = '/resume/feed/valid', [string]$OutputName = 'output', [switch]$Preview,
            [int]$ChunkLimit = 0, [long]$CancelAfterBytes = 0, [AllowNull()][Nullable[long]]$AvailableSpace)
        if ($FeedPath -notmatch '^/(?:feeds/[a-z0-9-]+\.xml|resume/feed/[a-z0-9-]+)$' -or $OutputName -notmatch '^[a-z0-9-]+$') {
            throw 'Run safety workers require named loopback feeds and simple owned output directories.'
        }
        $identifier = [guid]::NewGuid().ToString('N')
        $configPath = Join-Path $Context.Root ($identifier + '-safety-config.json')
        $config = @{ Root = $Context.Root; Token = $Context.Token; ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1';
            Action = $Action; FeedUrl = $Context.BaseUrl + $FeedPath; OutputPath = Join-Path $Context.Root $OutputName; Preview = [bool]$Preview;
            ChunkLimit = $ChunkLimit; CancelAfterBytes = $CancelAfterBytes;
            ReadyPath = Join-Path $Context.Root ($identifier + '-ready.txt'); ReleasePath = Join-Path $Context.Root ($identifier + '-release.txt');
            ResultPath = Join-Path $Context.Root ($identifier + '-safety-result.json') }
        if ($PSBoundParameters.ContainsKey('AvailableSpace')) { $config.AvailableSpace = $AvailableSpace }
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $engineName = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $owned = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engineName) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
            (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-RunSafetyWorker.ps1'), '-ConfigPath', $configPath)
        [pscustomobject]@{ Owned = $owned; Config = $config }
    }

    function Wait-UpdSafetyReady {
        param([Parameter(Mandatory)]$Worker)
        $watch = [Diagnostics.Stopwatch]::StartNew()
        while (-not [IO.File]::Exists($Worker.Config.ReadyPath)) {
            if ($Worker.Owned.Process.HasExited) { throw ('Safety worker exited before its gate: ' + $Worker.Owned.ErrorOutput.Result + $Worker.Owned.Output.Result) }
            if ($watch.Elapsed.TotalSeconds -gt 15) { throw 'Owned safety worker readiness timed out.' }
            Start-Sleep -Milliseconds 25
        }
    }

    function Complete-UpdSafetyWorker {
        param([Parameter(Mandatory)]$Worker)
        if (-not $Worker.Owned.Process.WaitForExit(30000)) { throw 'Owned safety worker completion timed out.' }
        if (-not [IO.File]::Exists($Worker.Config.ResultPath)) { throw ('Safety worker produced no result: ' + $Worker.Owned.ErrorOutput.Result) }
        [pscustomobject]@{ ExitCode = $Worker.Owned.Process.ExitCode; Result = [IO.File]::ReadAllText($Worker.Config.ResultPath) | ConvertFrom-Json;
            Stdout = $Worker.Owned.Output.Result; Stderr = $Worker.Owned.ErrorOutput.Result; OutputPath = $Worker.Config.OutputPath }
    }

    function Get-UpdSafetyArchive {
        param([Parameter(Mandatory)][string]$OutputPath)
        $states = @(Get-ChildItem -LiteralPath $OutputPath -Recurse -Force -File -Filter 'state.json')
        $states.Count | Should -Be 1
        [pscustomobject]@{ Path = $states[0].FullName; Root = Split-Path $states[0].DirectoryName -Parent;
            State = [IO.File]::ReadAllText($states[0].FullName) | ConvertFrom-Json }
    }

    function Assert-UpdSafetyResumeFile {
        param([Parameter(Mandatory)]$Run)
        $Run.ExitCode | Should -Be 0 -Because ($Run.Stdout + $Run.Stderr + $Run.Result.Error)
        $archive = Get-UpdSafetyArchive -OutputPath $Run.OutputPath
        @($archive.State.episodes).Count | Should -Be 1
        $record = $archive.State.episodes[0]
        $record.status | Should -Be 'transfer_verified'
        $record.bytes | Should -Be $script:safetyResumeBytes.Length
        $record.local_sha256 | Should -Be $script:safetyResumeHash
        $media = Join-Path $archive.Root $record.relative_path
        (Get-Item -LiteralPath $media).Length | Should -Be $script:safetyResumeBytes.Length
        (Get-FileHash -LiteralPath $media -Algorithm SHA256).Hash.ToLowerInvariant() | Should -Be $script:safetyResumeHash
    }

    function Assert-UpdSafetyAclRestore {
        param([Parameter(Mandatory)]$OriginalAcl, [Parameter(Mandatory)]$RestoredAcl)
        $originalDescriptor = [Security.AccessControl.RawSecurityDescriptor]::new($OriginalAcl.GetSecurityDescriptorBinaryForm(), 0)
        $restoredDescriptor = [Security.AccessControl.RawSecurityDescriptor]::new($RestoredAcl.GetSecurityDescriptorBinaryForm(), 0)
        $restoredDescriptor.Owner.Value | Should -BeExactly $originalDescriptor.Owner.Value
        $restoredDescriptor.Group.Value | Should -BeExactly $originalDescriptor.Group.Value
        $originalFlags = [int]$originalDescriptor.ControlFlags
        $restoredFlags = [int]$restoredDescriptor.ControlFlags
        $autoInherited = [int][Security.AccessControl.ControlFlags]::DiscretionaryAclAutoInherited
        # Windows may add AI when Set-Acl restores inherited permissions. Permit
        # only that addition; protection, inheritance requirements and every
        # other control flag must remain unchanged, including any original AI.
        ($restoredFlags -eq $originalFlags -or $restoredFlags -eq ($originalFlags -bor $autoInherited)) |
            Should -BeTrue -Because 'Windows may only add the discretionary ACL automatic-inheritance bookkeeping flag'
        $restoredDescriptor.ResourceManagerControl | Should -Be $originalDescriptor.ResourceManagerControl
        foreach ($aclName in @('DiscretionaryAcl', 'SystemAcl')) {
            $originalEntries = $originalDescriptor.$aclName
            $restoredEntries = $restoredDescriptor.$aclName
            ($null -ne $restoredEntries) | Should -Be ($null -ne $originalEntries)
            if ($null -eq $originalEntries) { continue }
            $restoredEntries.Revision | Should -Be $originalEntries.Revision
            $restoredEntries.Count | Should -Be $originalEntries.Count
            $originalBytes = New-Object byte[] $originalEntries.BinaryLength
            $restoredBytes = New-Object byte[] $restoredEntries.BinaryLength
            $originalEntries.GetBinaryForm($originalBytes, 0)
            $restoredEntries.GetBinaryForm($restoredBytes, 0)
            # Complete ACL bytes preserve ACE order, identity, type, rights,
            # object flags, inheritance, propagation and inherited status.
            [Convert]::ToBase64String($restoredBytes) | Should -BeExactly ([Convert]::ToBase64String($originalBytes))
        }
    }
}

Describe 'A042/A043 run safety through actual owned processes and loopback media' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:safetyPython
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'reports a useful same-show lock conflict after a real write and releases ownership on catchable cancellation' {
        $first = Start-UpdSafetyWorker -Context $context -Action GateCancel
        Wait-UpdSafetyReady -Worker $first
        try {
            $stateFile = @(Get-ChildItem -LiteralPath $first.Config.OutputPath -Recurse -Force -File -Filter 'state.json')
            $stateFile.Count | Should -Be 1
            $stateHash = (Get-FileHash -LiteralPath $stateFile[0].FullName -Algorithm SHA256).Hash
            $contender = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context)
            $contender.ExitCode | Should -Be 1 -Because ($contender.Stdout + $contender.Stderr + $contender.Result.Error)
            $contender.Result.RunResult.Status | Should -Be 'fatal'
            (Get-FileHash -LiteralPath $stateFile[0].FullName -Algorithm SHA256).Hash | Should -Be $stateHash
            (Get-UpdFixtureState -Context $context).'/resume/media/valid' | Should -Be 1
        }
        finally { [IO.File]::WriteAllText($first.Config.ReleasePath, 'cancel') }
        $cancelled = Complete-UpdSafetyWorker -Worker $first
        $cancelled.ExitCode | Should -Be 130 -Because ($cancelled.Stdout + $cancelled.Stderr + $cancelled.Result.Error)
        $cancelled.Result.HostSurvived | Should -BeTrue
        $cancelled.Result.CancelBytes | Should -BeGreaterThan 0
        $cancelled.Result.PartialClosed | Should -BeTrue
        $contender.Result.RunResult.Message | Should -Be 'The podcast archive writer lock is in use. Wait for the current writer to finish, then retry.'
    }

    It 'checkpoints every caught cancellation byte between normal checkpoints before a safe range recovery' {
        $first = Start-UpdSafetyWorker -Context $context -Action GateCancel -ChunkLimit 4096 -CancelAfterBytes 8192
        Wait-UpdSafetyReady -Worker $first
        [IO.File]::WriteAllText($first.Config.ReleasePath, 'cancel')
        $cancelled = Complete-UpdSafetyWorker -Worker $first
        $cancelled.ExitCode | Should -Be 130 -Because ($cancelled.Stdout + $cancelled.Stderr + $cancelled.Result.Error)
        $cancelled.Result.HostSurvived | Should -BeTrue
        $cancelled.Result.PartialClosed | Should -BeTrue
        $sidecars = @(Get-ChildItem -LiteralPath $cancelled.OutputPath -Recurse -Force -File -Filter 'resume-*.json')
        $sidecars.Count | Should -Be 1
        $checkpoint = [IO.File]::ReadAllText($sidecars[0].FullName) | ConvertFrom-Json
        $showRoot = Split-Path $sidecars[0].DirectoryName -Parent
        $partialPath = Join-Path $showRoot $checkpoint.partial_name
        $partial = Get-Item -LiteralPath $partialPath
        $partial.Length | Should -BeGreaterThan 4096
        $partial.Length | Should -BeLessThan $checkpoint.total_length
        $checkpoint.offset | Should -Be $partial.Length
        $checkpoint.prefix_sha256 | Should -Be ((Get-FileHash -LiteralPath $partialPath -Algorithm SHA256).Hash.ToLowerInvariant())
        @(Get-ChildItem -LiteralPath $cancelled.OutputPath -Recurse -Force -File -Filter '*.mp3').Count | Should -Be 0
        $cancelled.Result.RunResult.Cancelled | Should -Be 1
        $cancelled.Result.RunResult.Downloaded | Should -Be 0
        $cancelled.Stdout | Should -Not -Match 'Run completed|\[OK\]'
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $remaining = [long]$checkpoint.total_length - [long]$checkpoint.offset
        $recovered = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -AvailableSpace $remaining)
        Assert-UpdSafetyResumeFile -Run $recovered
        $recovered.Result.ResponseLength | Should -Be $remaining
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        $events.Count | Should -Be 2
        $events[1].range | Should -Be ('bytes=' + $checkpoint.offset + '-')
        $events[1].if_range | Should -Be '"resume-v1"'
        Test-Path -LiteralPath $sidecars[0].FullName | Should -BeFalse
        Test-Path -LiteralPath $partialPath | Should -BeFalse
        $repeat = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context)
        $repeat.ExitCode | Should -Be 0
        $repeat.Result.RunResult.VerifiedSkipped | Should -Be 1
        (Get-UpdFixtureState -Context $context).'/resume/media/valid' | Should -Be 2
    }

    It 'rejects an actual owned destination ACL denial before media or history writes and preserves originals' {
        $output = Join-Path $context.Root 'output'
        $null = [IO.Directory]::CreateDirectory($output)
        $original = Join-Path $output 'original.mp3'
        [IO.File]::WriteAllBytes($original, [IO.File]::ReadAllBytes((Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3')))
        $originalHash = (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash
        $originalAcl = Get-Acl -LiteralPath $output
        $deniedAcl = Get-Acl -LiteralPath $output
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        try { $ownerSid = $identity.User }
        finally { $identity.Dispose() }
        $denial = [Security.AccessControl.FileSystemAccessRule]::new($ownerSid, [Security.AccessControl.FileSystemRights]::CreateFiles, [Security.AccessControl.AccessControlType]::Deny)
        $deniedAcl.AddAccessRule($denial)
        try {
            Set-Acl -LiteralPath $output -AclObject $deniedAcl
            $run = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/single.xml')
            $run.ExitCode | Should -Be 1 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
            $run.Result.RunResult.Status | Should -Be 'fatal'
            (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
            @(Get-ChildItem -LiteralPath $output -Recurse -Force -File -Filter 'state.json').Count | Should -Be 0
            @(Get-ChildItem -LiteralPath $output -Recurse -Force -File -Filter '*.tmp').Count | Should -Be 0
            (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash | Should -Be $originalHash
            $run.Result.RunResult.Message | Should -Be 'The destination is not writable. Check output-folder permissions and retry.'
        }
        finally { Set-Acl -LiteralPath $output -AclObject $originalAcl }
        Assert-UpdSafetyAclRestore -OriginalAcl $originalAcl -RestoredAcl (Get-Acl -LiteralPath $output)
    }

    It 'rejects reliable response length above available space before any body write or accepted episode evidence' {
        $run = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/single.xml' -AvailableSpace 0)
        $run.ExitCode | Should -Be 2 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.SpaceCalls | Should -BeGreaterThan 0
        $run.Result.ProgressCalls | Should -Be 0
        $run.Result.ResponseLength | Should -BeGreaterThan 0
        $run.Result.RunResult.Failed | Should -Be 1
        $run.Result.RunResult.Episodes[0].Message | Should -Be 'Insufficient available space for the validated media response. Free space or choose another output folder.'
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Force -File -Filter '*.mp3').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Force -File -Filter '.upd-*.tmp').Count | Should -Be 0
        foreach ($stateFile in @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -Force -File -Filter 'state.json')) {
            $state = [IO.File]::ReadAllText($stateFile.FullName) | ConvertFrom-Json
            @($state.episodes | Where-Object { $_.bytes -gt 0 -or $_.status -in @('prepared', 'transfer_verified') }).Count | Should -Be 0
        }
    }

    It 'reports denied <Stage> directory creation after file-probe access succeeds and preserves existing files' -TestCases @(
        @{ Stage = 'base output'; ExistingOutput = $false }
        @{ Stage = 'new show'; ExistingOutput = $true }
    ) {
        param($Stage, $ExistingOutput)
        $null = $Stage
        $output = Join-Path $context.Root 'output'
        if ($ExistingOutput) { $null = [IO.Directory]::CreateDirectory($output) }
        $parent = if ($ExistingOutput) { $output } else { $context.Root }
        $original = Join-Path $parent 'directory-denial-original.mp3'
        [IO.File]::WriteAllBytes($original, $script:safetyAudio)
        $originalHash = (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash
        $originalAcl = Get-Acl -LiteralPath $parent
        $deniedAcl = Get-Acl -LiteralPath $parent
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        try { $ownerSid = $identity.User }
        finally { $identity.Dispose() }
        $deniedAcl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($ownerSid, [Security.AccessControl.FileSystemRights]::CreateDirectories, [Security.AccessControl.AccessControlType]::Deny))
        try {
            Set-Acl -LiteralPath $parent -AclObject $deniedAcl
            # Denying directory creation must leave CreateNew file access intact;
            # otherwise this would repeat the earlier permission-probe boundary.
            $probePath = Join-Path $parent 'owned-create-file-proof.bin'
            $probe = [IO.File]::Open($probePath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try { $probe.WriteByte(0); $probe.Flush($true) }
            finally { $probe.Dispose() }
            [IO.File]::Delete($probePath)
            $run = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/single.xml')
            $run.ExitCode | Should -Be 1 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
            $run.Result.RunResult.Status | Should -Be 'fatal'
            $run.Result.HostSurvived | Should -BeTrue
            (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
            Test-Path -LiteralPath $output | Should -Be $ExistingOutput
            if ($ExistingOutput) { @(Get-ChildItem -LiteralPath $output -Recurse -Force -File -Filter 'state.json').Count | Should -Be 0 }
            (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash | Should -Be $originalHash
            $run.Result.RunResult.Message | Should -Be 'The destination is not writable. Check output-folder permissions and retry.'
        }
        finally { Set-Acl -LiteralPath $parent -AclObject $originalAcl }
        Assert-UpdSafetyAclRestore -OriginalAcl $originalAcl -RestoredAcl (Get-Acl -LiteralPath $parent)
    }

    It 'allows an independent show under the same base while another show holds its writer handle' {
        $first = Start-UpdSafetyWorker -Context $context -Action GateCancel
        Wait-UpdSafetyReady -Worker $first
        try {
            $firstArchive = Get-UpdSafetyArchive -OutputPath $first.Config.OutputPath
            $firstStateHash = (Get-FileHash -LiteralPath $firstArchive.Path -Algorithm SHA256).Hash
            $independent = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/history.xml')
            $independent.ExitCode | Should -Be 0 -Because ($independent.Stdout + $independent.Stderr + $independent.Result.Error)
            $independent.Result.RunResult.Downloaded | Should -Be 1
            $first.Owned.Process.HasExited | Should -BeFalse
            (Get-FileHash -LiteralPath $firstArchive.Path -Algorithm SHA256).Hash | Should -Be $firstStateHash
            $states = @(Get-ChildItem -LiteralPath $independent.OutputPath -Recurse -Force -File -Filter 'state.json')
            $states.Count | Should -Be 2
            $secondState = [IO.File]::ReadAllText(($states | Where-Object { $_.FullName -ne $firstArchive.Path }).FullName) | ConvertFrom-Json
            $secondState.episodes[0].status | Should -Be 'transfer_verified'
            (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
        }
        finally { [IO.File]::WriteAllText($first.Config.ReleasePath, 'cancel') }
        $cancelled = Complete-UpdSafetyWorker -Worker $first
        $cancelled.ExitCode | Should -Be 130
        $cancelled.Result.PartialClosed | Should -BeTrue
    }

    It 'reports brief archive-selection contention and permits another show after normal lock release' {
        $output = Join-Path $context.Root 'output'
        $null = [IO.Directory]::CreateDirectory($output)
        $holder = Start-UpdSafetyWorker -Context $context -Action HoldArchive
        Wait-UpdSafetyReady -Worker $holder
        try {
            $blocked = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/history.xml')
            $blocked.ExitCode | Should -Be 1
            $blocked.Result.RunResult.Message | Should -Be 'The archive selection lock is in use. Wait for the current writer to finish, then retry.'
            @(Get-ChildItem -LiteralPath $output -Recurse -Force -File -Filter 'state.json').Count | Should -Be 0
            (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -BeNullOrEmpty
        }
        finally { [IO.File]::WriteAllText($holder.Config.ReleasePath, 'release') }
        (Complete-UpdSafetyWorker -Worker $holder).ExitCode | Should -Be 0
        $released = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/history.xml')
        $released.ExitCode | Should -Be 0 -Because ($released.Stdout + $released.Stderr + $released.Result.Error)
        $released.Result.RunResult.Downloaded | Should -Be 1
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }

    It 'ignores stale lock contents without terminating a live owned fixture process or rewriting recorded media' {
        $first = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/history.xml')
        $first.ExitCode | Should -Be 0
        $archive = Get-UpdSafetyArchive -OutputPath $first.OutputPath
        $mediaPath = Join-Path $archive.Root $archive.State.episodes[0].relative_path
        $mediaHash = (Get-FileHash -LiteralPath $mediaPath -Algorithm SHA256).Hash
        $stateHash = (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash
        $server = $context.Processes[0].Process
        $staleContent = '{"pid":' + $server.Id + ',"stale":true,"private":"synthetic owned marker"}'
        $lockPaths = @((Join-Path $first.OutputPath '.upd-archive.lock'), (Join-Path $archive.Root '.upd/writer.lock'))
        foreach ($lockPath in $lockPaths) { [IO.File]::WriteAllText($lockPath, $staleContent) }
        $repeat = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/history.xml')
        $repeat.ExitCode | Should -Be 0 -Because ($repeat.Stdout + $repeat.Stderr + $repeat.Result.Error)
        $repeat.Result.RunResult.VerifiedSkipped | Should -Be 1
        $server.HasExited | Should -BeFalse
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
        (Get-FileHash -LiteralPath $mediaPath -Algorithm SHA256).Hash | Should -Be $mediaHash
        (Get-FileHash -LiteralPath $archive.Path -Algorithm SHA256).Hash | Should -Be $stateHash
        foreach ($lockPath in $lockPaths) { [IO.File]::ReadAllText($lockPath) | Should -Be $staleContent }
        $repeat.Stdout | Should -Not -Match 'synthetic owned marker'
    }

    It 'retains cancelled unknown-length bytes without a resume claim and safely starts a fresh transfer' {
        $first = Start-UpdSafetyWorker -Context $context -Action GateCancel -FeedPath '/resume/feed/no-length' -ChunkLimit 4096 -CancelAfterBytes 8192
        Wait-UpdSafetyReady -Worker $first
        [IO.File]::WriteAllText($first.Config.ReleasePath, 'cancel')
        $cancelled = Complete-UpdSafetyWorker -Worker $first
        $cancelled.ExitCode | Should -Be 130 -Because ($cancelled.Stdout + $cancelled.Stderr + $cancelled.Result.Error)
        $cancelled.Result.HostSurvived | Should -BeTrue
        $cancelled.Result.PartialClosed | Should -BeTrue
        $cancelled.Result.ResponseLength | Should -BeNullOrEmpty
        $cancelled.Stdout | Should -Match 'The media response size is unknown; available space cannot be compared\.'
        $partial = @(Get-ChildItem -LiteralPath $cancelled.OutputPath -Recurse -Force -File -Filter '.upd-*.tmp')
        $partial.Count | Should -Be 1
        $partial[0].Length | Should -BeGreaterThan 0
        $partialHash = (Get-FileHash -LiteralPath $partial[0].FullName -Algorithm SHA256).Hash
        @(Get-ChildItem -LiteralPath $cancelled.OutputPath -Recurse -Force -File -Filter 'resume-*.json').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $cancelled.OutputPath -Recurse -Force -File -Filter '*.mp3').Count | Should -Be 0
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $recovered = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/resume/feed/no-length')
        Assert-UpdSafetyResumeFile -Run $recovered
        $events = Invoke-RestMethod -Uri ($context.BaseUrl + '/__resume')
        $events.Count | Should -Be 2
        $events[1].has_range | Should -BeFalse
        $events[1].has_if_range | Should -BeFalse
        (Get-FileHash -LiteralPath $partial[0].FullName -Algorithm SHA256).Hash | Should -Be $partialHash
        @(Get-ChildItem -LiteralPath $recovered.OutputPath -Recurse -Force -File -Filter 'resume-*.json').Count | Should -Be 0
    }

    It 'continues a known-length transfer with an honest unknown-capacity message when the provider cannot estimate space' {
        $run = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/single.xml' -AvailableSpace $null)
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.SpaceCalls | Should -BeGreaterThan 0
        $run.Result.ResponseLength | Should -BeGreaterThan 0
        $run.Stdout | Should -Match 'Available destination space is unknown; the validated response size cannot be compared\.'
        $file = @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3')
        $file.Count | Should -Be 1
        (Get-FileHash -LiteralPath $file[0].FullName -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3') -Algorithm SHA256).Hash
    }

    It 'does not use an enclosure estimate as a hard space gate when the real response size is unknown' {
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing
        $run = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/resume/feed/no-length' -AvailableSpace 0)
        Assert-UpdSafetyResumeFile -Run $run
        $run.Result.SpaceCalls | Should -BeGreaterThan 0
        $run.Result.ResponseLength | Should -BeNullOrEmpty
        $run.Stdout | Should -Match 'The media response size is unknown; available space cannot be compared\.'
        (Get-UpdFixtureState -Context $context).'/resume/media/no-length' | Should -Be 1
    }

    It 'previews a denied destination without probing permissions or capacity or writing media and history' {
        $output = Join-Path $context.Root 'output'
        $null = [IO.Directory]::CreateDirectory($output)
        $original = Join-Path $output 'original.mp3'
        [IO.File]::WriteAllBytes($original, $script:safetyAudio)
        $originalHash = (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash
        $originalStamp = (Get-Item -LiteralPath $output).LastWriteTimeUtc
        $originalAcl = Get-Acl -LiteralPath $output
        $deniedAcl = Get-Acl -LiteralPath $output
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        try { $ownerSid = $identity.User }
        finally { $identity.Dispose() }
        $deniedAcl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($ownerSid, [Security.AccessControl.FileSystemRights]::CreateFiles, [Security.AccessControl.AccessControlType]::Deny))
        try {
            Set-Acl -LiteralPath $output -AclObject $deniedAcl
            $run = Complete-UpdSafetyWorker -Worker (Start-UpdSafetyWorker -Context $context -FeedPath '/feeds/single.xml' -AvailableSpace 0 -Preview)
            $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
            $run.Result.RunResult.Status | Should -Be 'preview'
            $run.Result.SpaceCalls | Should -Be 0
            $run.Result.ProgressCalls | Should -Be 0
            @(Get-ChildItem -LiteralPath $output -Recurse -Force).Count | Should -Be 1
            (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash | Should -Be $originalHash
            (Get-Item -LiteralPath $output).LastWriteTimeUtc | Should -Be $originalStamp
            (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
        }
        finally { Set-Acl -LiteralPath $output -AclObject $originalAcl }
        Assert-UpdSafetyAclRestore -OriginalAcl $originalAcl -RestoredAcl (Get-Acl -LiteralPath $output)
    }
}
