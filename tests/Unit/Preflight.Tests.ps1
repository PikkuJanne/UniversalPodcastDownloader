BeforeAll {
    $script:PreflightRepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:PreflightRepositoryRoot 'src/PathSafety.ps1')
    . (Join-Path $script:PreflightRepositoryRoot 'src/RunResult.ps1')
    . (Join-Path $script:PreflightRepositoryRoot 'src/Preflight.ps1')
    Mock Invoke-WebRequest { throw 'Preflight units must not use the network.' }
    $script:OriginalPreflightProbe = (Get-Command Open-PodcastPreflightProbe).ScriptBlock
}

Describe 'A043: honest validated transfer-space observations' -Tag 'Unit', 'A043' {
    BeforeEach {
        Mock Get-PodcastAvailableSpace { throw 'Explicit capacity should not query a drive.' }
    }

    It 'compares <Kind> without confusing unknown estimates with zero' -ForEach @(
        @{ Kind = 'unknown size and zero capacity'; Required = $null; Available = 0L; Check = 'unknown_size' }
        @{ Kind = 'unknown size and unknown capacity'; Required = $null; Available = $null; Check = 'unknown_size' }
        @{ Kind = 'known size and unknown capacity'; Required = 100L; Available = $null; Check = 'unknown_available' }
        @{ Kind = 'zero size and zero capacity'; Required = 0L; Available = 0L; Check = 'sufficient' }
        @{ Kind = 'exact capacity'; Required = 100L; Available = 100L; Check = 'sufficient' }
        @{ Kind = 'capacity above size'; Required = 100L; Available = 101L; Check = 'sufficient' }
        @{ Kind = 'Int64 maximum equality'; Required = [long]::MaxValue; Available = [long]::MaxValue; Check = 'sufficient' }
    ) {
        $result = Assert-PodcastTransferSpace -Root $TestDrive -RequiredBytes $Required -AvailableBytes $Available
        $result.RequiredBytes | Should -Be $Required
        $result.AvailableBytes | Should -Be $Available
        $result.SpaceCheck | Should -BeExactly $Check
        if ($Check -like 'unknown*') { $result.Message | Should -Match 'unknown|cannot be compared' }
        else { $result.Message | Should -Match 'capacity can change' }
        Should -Invoke Get-PodcastAvailableSpace -Times 0 -Exactly
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
    }

    It 'rejects known insufficient <Kind> capacity without overflowing' -ForEach @(
        @{ Kind = 'zero'; Required = 1L; Available = 0L }
        @{ Kind = 'one byte short'; Required = 100L; Available = 99L }
        @{ Kind = 'Int64 maximum'; Required = [long]::MaxValue; Available = [long]::MaxValue - 1L }
    ) {
        { Assert-PodcastTransferSpace -Root $TestDrive -RequiredBytes $Required -AvailableBytes $Available } |
            Should -Throw 'Insufficient available space for the validated media response. Free space or choose another output folder.'
        Should -Invoke Get-PodcastAvailableSpace -Times 0 -Exactly
    }

    It 'queries the destination provider when capacity was not explicitly supplied' {
        Mock Get-PodcastAvailableSpace { 100L }
        $result = Assert-PodcastTransferSpace -Root $TestDrive -RequiredBytes 100L
        $result.SpaceCheck | Should -Be 'sufficient'
        $result.AvailableBytes | Should -Be 100L
        Should -Invoke Get-PodcastAvailableSpace -Times 1 -Exactly -ParameterFilter { $Root -eq $TestDrive }
    }

    It 'retains unknown queried capacity without rejecting a known response size' {
        Mock Get-PodcastAvailableSpace { $null }
        $result = Assert-PodcastTransferSpace -Root $TestDrive -RequiredBytes 100L
        $result.SpaceCheck | Should -Be 'unknown_available'
        $result.AvailableBytes | Should -BeNullOrEmpty
    }

    It 'rejects negative <Kind> estimates before querying capacity' -ForEach @(
        @{ Kind = 'required'; Required = -1L; Available = $null }
        @{ Kind = 'available'; Required = 1L; Available = -1L }
    ) {
        { Assert-PodcastTransferSpace -Root $TestDrive -RequiredBytes $Required -AvailableBytes $Available } |
            Should -Throw 'Transfer space estimates must be nonnegative byte counts.'
        Should -Invoke Get-PodcastAvailableSpace -Times 0 -Exactly
    }

    It 'rejects an out-of-range <Parameter> value rather than wrapping into a false estimate' -ForEach @(
        @{ Parameter = 'RequiredBytes' }, @{ Parameter = 'AvailableBytes' }
    ) {
        $arguments = @{ Root = $TestDrive; RequiredBytes = 1L; AvailableBytes = 1L }
        $arguments[$Parameter] = '9223372036854775808'
        { Assert-PodcastTransferSpace @arguments } | Should -Throw
        Should -Invoke Get-PodcastAvailableSpace -Times 0 -Exactly
    }

    It 'rejects a provider returning an impossible negative capacity' {
        Mock Get-PodcastAvailableSpace { -1L }
        { Assert-PodcastTransferSpace -Root $TestDrive -RequiredBytes 1L } |
            Should -Throw 'Transfer space estimates must be nonnegative byte counts.'
    }
}

Describe 'A043: read-only quota-aware capacity lookup' -Tag 'Unit', 'A043' {
    It 'reports a local observation without creating output files' {
        $before = @(Get-ChildItem -LiteralPath $TestDrive -Force).Count
        $available = Get-PodcastAvailableSpace -Root $TestDrive
        if ($null -ne $available) {
            ($available -is [long]) | Should -BeTrue
            $available | Should -BeGreaterOrEqual 0
        }
        @(Get-ChildItem -LiteralPath $TestDrive -Force).Count | Should -Be $before
    }

    It 'reports UNC capacity as unknown without contacting the share' {
        Mock Get-PodcastPathAttribute { throw 'UNC capacity must not inspect a remote share.' }
        Get-PodcastAvailableSpace -Root '\\share.example.invalid\podcasts' | Should -BeNullOrEmpty
        Should -Invoke Get-PodcastPathAttribute -Times 0 -Exactly
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
    }
}

Describe 'A043: exclusive owned destination write probe' -Tag 'Unit', 'A043' {
    BeforeEach {
        $script:ProbeRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($script:ProbeRoot)
        [IO.File]::WriteAllText((Join-Path $script:ProbeRoot '.upd-test-owner'), $script:ProbeRoot)
        $script:ExistingProbeFile = Join-Path $script:ProbeRoot '.pf-original'
        [IO.File]::WriteAllText($script:ExistingProbeFile, 'existing private file bytes')
        $script:ExistingProbeHash = (Get-FileHash -LiteralPath $script:ExistingProbeFile -Algorithm SHA256).Hash
    }

    It 'proves create write and exclusive ownership then removes only its own probe' {
        $script:ObservedProbe = $null
        Mock Open-PodcastPreflightProbe {
            $script:ObservedProbe = & $script:OriginalPreflightProbe -LiteralPath $LiteralPath
            { [IO.File]::Open($LiteralPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite) } | Should -Throw
            return $script:ObservedProbe
        }
        $result = Invoke-PodcastDestinationPreflight -Root $script:ProbeRoot
        $result.Writable | Should -BeTrue
        $result.Probed | Should -BeTrue
        $result.Preview | Should -BeFalse
        $script:ObservedProbe.CanWrite | Should -BeFalse
        Test-Path -LiteralPath $script:ObservedProbe.Name | Should -BeFalse
        @(Get-ChildItem -LiteralPath $script:ProbeRoot -Force).Count | Should -Be 2
        (Get-FileHash -LiteralPath $script:ExistingProbeFile -Algorithm SHA256).Hash | Should -BeExactly $script:ExistingProbeHash
        Should -Invoke Open-PodcastPreflightProbe -Times 1 -Exactly
    }

    It 'tests the nearest existing directory without creating a missing output hierarchy' {
        $missing = Join-Path $script:ProbeRoot 'absent\nested'
        Mock Open-PodcastPreflightProbe { & $script:OriginalPreflightProbe -LiteralPath $LiteralPath }
        $result = Invoke-PodcastDestinationPreflight -Root $missing
        $result.Probed | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:ProbeRoot 'absent') | Should -BeFalse
        Should -Invoke Open-PodcastPreflightProbe -Times 1 -Exactly -ParameterFilter { [IO.Path]::GetDirectoryName($LiteralPath) -eq $script:ProbeRoot }
        @(Get-ChildItem -LiteralPath $script:ProbeRoot -Force).Count | Should -Be 2
    }

    It 'keeps <Kind> preview read-only and does not claim verified write access' -ForEach @(
        @{ Kind = 'existing'; Relative = '' }, @{ Kind = 'missing'; Relative = 'absent\nested' }
    ) {
        Mock Open-PodcastPreflightProbe { throw 'Preview must not create a write probe.' }
        $target = if ($Relative) { Join-Path $script:ProbeRoot $Relative } else { $script:ProbeRoot }
        $result = Invoke-PodcastDestinationPreflight -Root $target -Preview
        $result.Probed | Should -BeFalse
        $result.Writable | Should -BeNullOrEmpty
        $result.Preview | Should -BeTrue
        Should -Invoke Open-PodcastPreflightProbe -Times 0 -Exactly
        @(Get-ChildItem -LiteralPath $script:ProbeRoot -Force).Count | Should -Be 2
    }

    It 'preserves a colliding existing filename and retries with a fresh owned name' {
        $script:ProbeNames = [Collections.Generic.Queue[string]]::new()
        $script:ProbeNames.Enqueue('.pf-original')
        $script:ProbeNames.Enqueue('.pf-fresh')
        Mock Get-PodcastPreflightProbeName { $script:ProbeNames.Dequeue() }
        Invoke-PodcastDestinationPreflight -Root $script:ProbeRoot | Out-Null
        $script:ProbeNames.Count | Should -Be 0
        (Get-FileHash -LiteralPath $script:ExistingProbeFile -Algorithm SHA256).Hash | Should -BeExactly $script:ExistingProbeHash
        Test-Path -LiteralPath (Join-Path $script:ProbeRoot '.pf-fresh') | Should -BeFalse
        @(Get-ChildItem -LiteralPath $script:ProbeRoot -Force).Count | Should -Be 2
    }

    It 'preserves a file racing CreateNew and reserves another name' {
        $script:ProbeNames = [Collections.Generic.Queue[string]]::new()
        $script:ProbeNames.Enqueue('.pf-raced')
        $script:ProbeNames.Enqueue('.pf-fresh')
        Mock Get-PodcastPreflightProbeName { $script:ProbeNames.Dequeue() }
        Mock Open-PodcastPreflightProbe {
            if ([IO.Path]::GetFileName($LiteralPath) -eq '.pf-raced') {
                [IO.File]::WriteAllText($LiteralPath, 'competing bytes')
                throw [IO.IOException]::new('CreateNew collision')
            }
            & $script:OriginalPreflightProbe -LiteralPath $LiteralPath
        }
        Invoke-PodcastDestinationPreflight -Root $script:ProbeRoot | Out-Null
        [IO.File]::ReadAllText((Join-Path $script:ProbeRoot '.pf-raced')) | Should -BeExactly 'competing bytes'
        Test-Path -LiteralPath (Join-Path $script:ProbeRoot '.pf-fresh') | Should -BeFalse
        (Get-FileHash -LiteralPath $script:ExistingProbeFile -Algorithm SHA256).Hash | Should -BeExactly $script:ExistingProbeHash
    }

    It 'bounds name-collision retries without opening or deleting existing files' {
        Mock Get-PodcastPreflightProbeName { '.pf-original' }
        Mock Open-PodcastPreflightProbe { throw 'An existing file must never be opened.' }
        { Invoke-PodcastDestinationPreflight -Root $script:ProbeRoot } |
            Should -Throw 'The destination is not writable. Check output-folder permissions and retry.'
        Should -Invoke Get-PodcastPreflightProbeName -Times 8 -Exactly
        Should -Invoke Open-PodcastPreflightProbe -Times 0 -Exactly
        (Get-FileHash -LiteralPath $script:ExistingProbeFile -Algorithm SHA256).Hash | Should -BeExactly $script:ExistingProbeHash
    }

    It 'reports denied creation with a fixed useful message and leaves existing files unchanged' {
        Mock Open-PodcastPreflightProbe { throw [UnauthorizedAccessException]::new('privateAccessCanary C:\private\archive') }
        { Invoke-PodcastDestinationPreflight -Root $script:ProbeRoot } |
            Should -Throw 'The destination is not writable. Check output-folder permissions and retry.'
        (Get-FileHash -LiteralPath $script:ExistingProbeFile -Algorithm SHA256).Hash | Should -BeExactly $script:ExistingProbeHash
        @(Get-ChildItem -LiteralPath $script:ProbeRoot -Force).Count | Should -Be 2
    }

    It 'adapts the random filename to the maximum supported directory length' {
        $name = Get-PodcastPreflightProbeName -Directory ('C:\' + ('x' * 244))
        (247 + 1 + $name.Length) | Should -BeLessOrEqual 259
        $name | Should -Match '^\.pf-[a-f0-9]{4,32}$'
    }

    It 'rejects an unsafe ancestor before opening a probe' {
        Mock Get-PodcastPathAttribute { [IO.FileAttributes]::Directory -bor [IO.FileAttributes]::ReparsePoint } -ParameterFilter { $LiteralPath -eq $script:ProbeRoot }
        Mock Open-PodcastPreflightProbe { throw 'Unsafe paths must not be opened.' }
        { Invoke-PodcastDestinationPreflight -Root $script:ProbeRoot } | Should -Throw '*reparse point*'
        Should -Invoke Open-PodcastPreflightProbe -Times 0 -Exactly
    }
}

Describe 'A043: probe failures release handles and preserve cancellation' -Tag 'Unit', 'A043' {
    BeforeEach {
        $script:FailureProbe = [pscustomobject]@{ Disposed = $false; Written = $false; Flushed = $false; Failure = '' }
        $script:FailureProbe | Add-Member ScriptMethod WriteByte {
            param($value)
            $this.Written = $value -eq 0
            if ($this.Failure -eq 'write') { throw [OperationCanceledException]::new('privateWriteCancellation') }
            if ($this.Failure -eq 'ordinary') { throw [IO.IOException]::new('privateWriteCanary') }
        }
        $script:FailureProbe | Add-Member ScriptMethod Flush {
            param($flushToDisk)
            $this.Flushed = $flushToDisk
            if ($this.Failure -eq 'flush') { throw [OperationCanceledException]::new('privateFlushCancellation') }
        }
        $script:FailureProbe | Add-Member ScriptMethod Dispose {
            $this.Disposed = $true
            if ($this.Failure -in @('write', 'cleanup')) { throw [IO.IOException]::new('privateCleanupCanary') }
            if ($this.Failure -eq 'dispose') { throw [OperationCanceledException]::new('privateDisposeCancellation') }
        }
        Mock Open-PodcastPreflightProbe { $script:FailureProbe }
    }

    It 'preserves typed cancellation during <Stage> and disposes the reserved handle' -ForEach @(
        @{ Stage = 'write' }, @{ Stage = 'flush' }, @{ Stage = 'dispose' }
    ) {
        $script:FailureProbe.Failure = $Stage
        $caught = $null
        try { $null = Invoke-PodcastDestinationPreflight -Root $TestDrive }
        catch { $caught = $_ }
        $caught | Should -Not -BeNullOrEmpty
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $script:FailureProbe.Disposed | Should -BeTrue
    }

    It 'preserves typed cancellation during creation without claiming probe ownership' {
        Mock Open-PodcastPreflightProbe { throw [OperationCanceledException]::new('privateCreateCancellation') }
        $caught = $null
        try { $null = Invoke-PodcastDestinationPreflight -Root $TestDrive }
        catch { $caught = $_ }
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $script:FailureProbe.Disposed | Should -BeFalse
    }

    It 'reports ordinary <Stage> failures safely after attempting cleanup' -ForEach @(
        @{ Stage = 'ordinary' }, @{ Stage = 'cleanup' }
    ) {
        $script:FailureProbe.Failure = $Stage
        { Invoke-PodcastDestinationPreflight -Root $TestDrive } |
            Should -Throw 'The destination is not writable. Check output-folder permissions and retry.'
        $script:FailureProbe.Disposed | Should -BeTrue
    }
}
