BeforeAll {
    $script:keepAwakeRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:keepAwakeSource = Join-Path $script:keepAwakeRepository 'src/KeepAwake.ps1'
    if (Test-Path -LiteralPath $script:keepAwakeSource) { . $script:keepAwakeSource }
    . (Join-Path $script:keepAwakeRepository 'src/RunResult.ps1')
    Mock Invoke-WebRequest { throw 'Keep-awake units must not request the network.' }
    Initialize-PodcastKeepAwakeType
    if (-not ('UPD.TestKeepAwakeApi' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
namespace UPD
{
    public sealed class TestKeepAwakeApi
    {
        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern uint GetCurrentThreadId();
        public readonly List<uint> Flags = new List<uint>();
        public readonly List<uint> Threads = new List<uint>();
        public uint State;
        public bool FailActivation;
        public bool ThrowActivation;
        public bool FailRestoration;
        public bool ThrowRestoration;
        public bool MismatchRestorationThread;
        private int threadCalls;
        public uint CleanupCallerThread;
        public Func<uint, uint> StateCall { get { return SetState; } }
        public Func<uint> ThreadCall { get { return ThreadId; } }
        public TestKeepAwakeApi(uint state) { State = state; }
        public uint ThreadId()
        {
            uint thread = GetCurrentThreadId();
            threadCalls++;
            return MismatchRestorationThread && threadCalls > 1 ? thread + 1 : thread;
        }
        public uint SetState(uint flags)
        {
            Flags.Add(flags);
            Threads.Add(GetCurrentThreadId());
            if (Flags.Count == 1 && ThrowActivation) { throw new InvalidOperationException("private activation canary"); }
            if (Flags.Count > 1 && ThrowRestoration) { throw new InvalidOperationException("private restoration canary"); }
            if ((Flags.Count == 1 && FailActivation) || (Flags.Count > 1 && FailRestoration)) { return 0; }
            uint previous = State;
            State = flags;
            return previous;
        }
        public bool StopElsewhere(object lease)
        {
            bool result = false;
            Thread caller = new Thread(delegate()
            {
                CleanupCallerThread = GetCurrentThreadId();
                result = (bool)lease.GetType().GetMethod("Stop").Invoke(lease, null);
            });
            caller.Start();
            caller.Join();
            return result;
        }
    }
}
'@
    }

    function New-KeepAwakeTestLease {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Allocates only an inactive lease with a compiled fake state provider.')]
        [CmdletBinding()]
        param([uint32]$Previous = 2147483648)
        $provider = [UPD.TestKeepAwakeApi]::new($Previous)
        $lease = [UPD.KeepAwakeLease]::new($provider.StateCall, $provider.ThreadCall)
        return [pscustomobject]@{ Provider = $provider; Lease = $lease }
    }
}

Describe 'A045 temporary keep-awake defaults' {
    It 'keeps disabled by default without compiling or activating native code' {
        Mock Add-Type { throw 'Disabled keep-awake must not compile native code.' }
        $context = Start-PodcastKeepAwake
        $context.Requested | Should -BeFalse
        $context.Active | Should -BeFalse
        $context.Lease | Should -BeNullOrEmpty
        Should -Invoke Add-Type -Times 0 -Exactly -Scope It
    }

    It 'accepts null cleanup without initializing native code' {
        Mock Add-Type { throw 'Null cleanup must not compile native code.' }
        $status = Stop-PodcastKeepAwake -Context $null
        $status.Requested | Should -BeFalse
        $status.Active | Should -BeFalse
        $status.Restored | Should -BeTrue
        Should -Invoke Add-Type -Times 0 -Exactly -Scope It
    }

    It 'treats explicit false as disabled without checking the platform' {
        Mock Test-PodcastKeepAwakePlatform { throw 'Disabled keep-awake must not check the platform.' }
        $context = Start-PodcastKeepAwake -Enabled:$false
        $context.Requested | Should -BeFalse
        $context.Active | Should -BeFalse
        Should -Invoke Test-PodcastKeepAwakePlatform -Times 0 -Exactly -Scope It
    }

    It 'imports silently without initializing or allocating a lease' {
        Mock Initialize-PodcastKeepAwakeType { throw 'Import must not initialize native code.' }
        Mock New-PodcastKeepAwakeLease { throw 'Import must not allocate a lease.' }
        @(. $script:keepAwakeSource).Count | Should -Be 0
        Should -Invoke Initialize-PodcastKeepAwakeType -Times 0 -Exactly -Scope It
        Should -Invoke New-PodcastKeepAwakeLease -Times 0 -Exactly -Scope It
    }

    It 'reports an unsupported platform privately without compiling or allocating native code' {
        Mock Test-PodcastKeepAwakePlatform { return $false }
        Mock Initialize-PodcastKeepAwakeType { throw 'Unsupported platforms must not initialize native code.' }
        Mock New-PodcastKeepAwakeLease { throw 'Unsupported platforms must not allocate a lease.' }
        $context = Start-PodcastKeepAwake -Enabled
        $context.Requested | Should -BeTrue
        $context.Active | Should -BeFalse
        $context.Lease | Should -BeNullOrEmpty
        $context.Message | Should -BeExactly 'Temporary keep-awake is unavailable; the download will continue without it.'
        Should -Invoke Initialize-PodcastKeepAwakeType -Times 0 -Exactly -Scope It
        Should -Invoke New-PodcastKeepAwakeLease -Times 0 -Exactly -Scope It
    }

    It 'omits private <Failure> errors and continues without activation' -ForEach @(
        @{ Failure = 'platform' }, @{ Failure = 'compilation' }, @{ Failure = 'allocation' }
    ) {
        Mock Test-PodcastKeepAwakePlatform { return $true }
        if ($Failure -eq 'platform') { Mock Test-PodcastKeepAwakePlatform { throw 'private power platform canary' } }
        elseif ($Failure -eq 'compilation') { Mock Initialize-PodcastKeepAwakeType { throw 'private power compilation canary' } }
        else { Mock New-PodcastKeepAwakeLease { throw 'private power allocation canary' } }
        $context = Start-PodcastKeepAwake -Enabled
        $context.Requested | Should -BeTrue
        $context.Active | Should -BeFalse
        $context.Lease | Should -BeNullOrEmpty
        $context.Message | Should -BeExactly 'Temporary keep-awake is unavailable; the download will continue without it.'
        ($context | ConvertTo-Json -Depth 2) | Should -Not -Match 'private power|canary|Exception'
    }

    It 'preserves catchable cancellation during initialization without allocating native state' {
        Mock Test-PodcastKeepAwakePlatform { return $true }
        Mock Initialize-PodcastKeepAwakeType { throw [OperationCanceledException]::new('private cancellation canary') }
        Mock New-PodcastKeepAwakeLease { throw 'Cancelled initialization must not allocate a lease.' }
        $failure = $null
        try { $null = Start-PodcastKeepAwake -Enabled } catch { $failure = $_ }
        Test-PodcastCancellation -ErrorObject $failure | Should -BeTrue
        Should -Invoke New-PodcastKeepAwakeLease -Times 0 -Exactly -Scope It
    }

    It 'returns one active context and one fixed successful cleanup status' {
        $fixture = New-KeepAwakeTestLease -Previous 2147483650
        Mock Test-PodcastKeepAwakePlatform { return $true }
        Mock New-PodcastKeepAwakeLease { return $fixture.Lease }
        try {
            $contexts = @(Start-PodcastKeepAwake -Enabled)
            $contexts.Count | Should -Be 1
            $context = $contexts[0]
            $context.Requested | Should -BeTrue
            $context.Active | Should -BeTrue
            [object]::ReferenceEquals($context.Lease, $fixture.Lease) | Should -BeTrue
            $statuses = @(Stop-PodcastKeepAwake -Context $context)
            $statuses.Count | Should -Be 1
            $statuses[0].Restored | Should -BeTrue
            $statuses[0].Active | Should -BeFalse
            $context.Active | Should -BeFalse
            $statuses[0].Message | Should -BeExactly 'Temporary keep-awake was released and its prior thread state restored.'
            $fixture.Provider.State | Should -Be 2147483650
            $fixture.Lease.IsAlive | Should -BeFalse
        }
        finally { $null = $fixture.Lease.Stop() }
    }
}

Describe 'A045 dedicated native-thread state ownership' {
    It 'requests only continuous system availability and restores exact prior flags <Previous>' -ForEach @(
        @{ Previous = [uint32]2147483648 }, @{ Previous = [uint32]2147483649 },
        @{ Previous = [uint32]2147483650 }, @{ Previous = [uint32]2147483651 },
        @{ Previous = [uint32]2147483712 }
    ) {
        $fixture = New-KeepAwakeTestLease -Previous $Previous
        try {
            $fixture.Lease.Start() | Should -BeTrue
            $fixture.Provider.Flags.Count | Should -Be 1
            $fixture.Provider.Flags[0] | Should -Be 2147483649
            $fixture.Lease.PreviousState | Should -Be $Previous
            $fixture.Lease.Stop() | Should -BeTrue
            $fixture.Provider.Flags.Count | Should -Be 2
            $fixture.Provider.Flags[1] | Should -Be $Previous
            $fixture.Provider.State | Should -Be $Previous
            $fixture.Provider.Threads[0] | Should -Be $fixture.Provider.Threads[1]
            $fixture.Lease.ActivationThreadId | Should -BeGreaterThan 0
            $fixture.Lease.RestorationThreadId | Should -Be $fixture.Lease.ActivationThreadId
            $fixture.Lease.IsAlive | Should -BeFalse
        }
        finally { $null = $fixture.Lease.Stop() }
    }

    It 'restores on the activation thread when another managed thread requests cleanup' {
        $fixture = New-KeepAwakeTestLease
        try {
            $fixture.Lease.Start() | Should -BeTrue
            $fixture.Provider.StopElsewhere($fixture.Lease) | Should -BeTrue
            $fixture.Provider.CleanupCallerThread | Should -Not -Be $fixture.Lease.ActivationThreadId
            $fixture.Lease.RestorationThreadId | Should -Be $fixture.Lease.ActivationThreadId
            $fixture.Provider.State | Should -Be 2147483648
            $fixture.Lease.IsAlive | Should -BeFalse
        }
        finally { $null = $fixture.Lease.Stop() }
    }

    It 'keeps independent leases active until each caller releases its own request' {
        $first = New-KeepAwakeTestLease
        $second = New-KeepAwakeTestLease
        try {
            $first.Lease.Start() | Should -BeTrue
            $second.Lease.Start() | Should -BeTrue
            $first.Lease.ActivationThreadId | Should -Not -Be $second.Lease.ActivationThreadId
            $first.Lease.Stop() | Should -BeTrue
            $second.Lease.Active | Should -BeTrue
            $second.Provider.Flags.Count | Should -Be 1
            $second.Lease.Stop() | Should -BeTrue
            $second.Provider.Flags.Count | Should -Be 2
            $first.Lease.IsAlive | Should -BeFalse
            $second.Lease.IsAlive | Should -BeFalse
        }
        finally { $null = $first.Lease.Stop(); $null = $second.Lease.Stop() }
    }

    It 'restores prior flags from caller finally after <Outcome>' -ForEach @(
        @{ Outcome = 'normal completion' }, @{ Outcome = 'failure' }, @{ Outcome = 'catchable cancellation' }
    ) {
        $fixture = New-KeepAwakeTestLease -Previous 2147483651
        $failure = $null
        try {
            $fixture.Lease.Start() | Should -BeTrue
            if ($Outcome -eq 'failure') { throw [InvalidOperationException]::new('private operation failure') }
            if ($Outcome -eq 'catchable cancellation') { throw [OperationCanceledException]::new('private operation cancellation') }
        }
        catch { $failure = $_ }
        finally { $fixture.Lease.Stop() | Should -BeTrue }
        $fixture.Provider.State | Should -Be 2147483651
        $fixture.Lease.RestorationThreadId | Should -Be $fixture.Lease.ActivationThreadId
        $fixture.Lease.IsAlive | Should -BeFalse
        if ($Outcome -eq 'catchable cancellation') { Test-PodcastCancellation -ErrorObject $failure | Should -BeTrue }
        elseif ($Outcome -eq 'failure') { $failure.Exception | Should -BeOfType ([InvalidOperationException]) }
        else { $failure | Should -BeNullOrEmpty }
    }

    It 'makes repeated cleanup idempotent without extra state changes' {
        $fixture = New-KeepAwakeTestLease
        try {
            $fixture.Lease.Start() | Should -BeTrue
            $fixture.Lease.Stop() | Should -BeTrue
            $fixture.Lease.Stop() | Should -BeTrue
            $fixture.Lease.Dispose()
            $fixture.Provider.Flags.Count | Should -Be 2
            $fixture.Lease.IsAlive | Should -BeFalse
        }
        finally { $null = $fixture.Lease.Stop() }
    }

    It 'does not activate after cleanup before start or a duplicate start' {
        $fixture = New-KeepAwakeTestLease
        $stopped = New-KeepAwakeTestLease
        try {
            $stopped.Lease.Stop() | Should -BeTrue
            $stopped.Lease.Start() | Should -BeFalse
            $stopped.Provider.Flags.Count | Should -Be 0
            $fixture.Lease.Start() | Should -BeTrue
            $fixture.Lease.Start() | Should -BeFalse
            $fixture.Lease.Active | Should -BeTrue
            $fixture.Provider.Flags.Count | Should -Be 1
        }
        finally { $null = $fixture.Lease.Stop(); $null = $stopped.Lease.Stop() }
    }

    It 'does not attempt restoration after failed activation <Failure>' -ForEach @(
        @{ Failure = 'zero' }, @{ Failure = 'exception' }
    ) {
        $fixture = New-KeepAwakeTestLease
        if ($Failure -eq 'zero') { $fixture.Provider.FailActivation = $true }
        else { $fixture.Provider.ThrowActivation = $true }
        try {
            $fixture.Lease.Start() | Should -BeFalse
            $fixture.Lease.Stop() | Should -BeTrue
            $fixture.Lease.ActivationSucceeded | Should -BeFalse
            $fixture.Provider.Flags.Count | Should -Be 1
            $fixture.Provider.State | Should -Be 2147483648
            $fixture.Lease.IsAlive | Should -BeFalse
        }
        finally { $null = $fixture.Lease.Stop() }
    }

    It 'reports failed restoration privately and closes its worker after <Failure>' -ForEach @(
        @{ Failure = 'zero' }, @{ Failure = 'exception' }, @{ Failure = 'different thread' }
    ) {
        $fixture = New-KeepAwakeTestLease
        if ($Failure -eq 'zero') { $fixture.Provider.FailRestoration = $true }
        elseif ($Failure -eq 'exception') { $fixture.Provider.ThrowRestoration = $true }
        else { $fixture.Provider.MismatchRestorationThread = $true }
        $context = [pscustomobject]@{ Requested = $true; Active = $false; Lease = $fixture.Lease }
        try {
            $fixture.Lease.Start() | Should -BeTrue
            $context.Active = $true
            $status = Stop-PodcastKeepAwake -Context $context
            $status.Restored | Should -BeFalse
            $status.Active | Should -BeFalse
            $status.Message | Should -BeExactly 'Temporary keep-awake cleanup could not be confirmed.'
            ($status | ConvertTo-Json) | Should -Not -Match 'private|canary|Exception'
            $fixture.Lease.IsAlive | Should -BeFalse
            if ($Failure -eq 'different thread') { $fixture.Provider.Flags.Count | Should -Be 1 }
            else { $fixture.Provider.Flags.Count | Should -Be 2 }
        }
        finally { $null = $fixture.Lease.Stop() }
    }
}
