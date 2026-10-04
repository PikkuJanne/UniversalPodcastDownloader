#requires -Version 5.1

function Test-PodcastKeepAwakePlatform {
    [CmdletBinding()]
    param()

    return [Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT
}

function Initialize-PodcastKeepAwakeType {
    [CmdletBinding()]
    param()

    if ('UPD.KeepAwakeLease' -as [type]) { return }
    # Compile only after an explicit, confirmed opt-in. The dedicated worker
    # invokes managed methods, never a PowerShell callback on a raw thread.
    Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace UPD
{
    public interface IKeepAwakeApi
    {
        uint SetExecutionState(uint flags);
        uint GetThreadId();
    }

    internal sealed class NativeKeepAwakeApi : IKeepAwakeApi
    {
        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern uint SetThreadExecutionState(uint flags);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern uint GetCurrentThreadId();
        public uint SetExecutionState(uint flags) { return SetThreadExecutionState(flags); }
        public uint GetThreadId() { return GetCurrentThreadId(); }
    }

    internal sealed class CallbackKeepAwakeApi : IKeepAwakeApi
    {
        private readonly Func<uint, uint> setState;
        private readonly Func<uint> threadId;
        public CallbackKeepAwakeApi(Func<uint, uint> setState, Func<uint> threadId)
        {
            if (setState == null || threadId == null) { throw new ArgumentNullException(); }
            this.setState = setState;
            this.threadId = threadId;
        }
        public uint SetExecutionState(uint flags) { return setState(flags); }
        public uint GetThreadId() { return threadId(); }
    }

    public sealed class KeepAwakeLease : IDisposable
    {
        public const uint Continuous = 0x80000000;
        public const uint SystemRequired = 0x00000001;
        private readonly IKeepAwakeApi api;
        private readonly object gate = new object();
        private readonly ManualResetEvent ready = new ManualResetEvent(false);
        private readonly ManualResetEvent stop = new ManualResetEvent(false);
        private Thread worker;
        private bool stopped;
        private bool eventsDisposed;
        private volatile bool active;
        private volatile bool restored = true;

        public bool Active { get { return active; } }
        public bool Restored { get { return restored; } }
        public bool ActivationSucceeded { get; private set; }
        public uint PreviousState { get; private set; }
        public uint ActivationThreadId { get; private set; }
        public uint RestorationThreadId { get; private set; }
        public bool IsAlive { get { return worker != null && worker.IsAlive; } }

        public KeepAwakeLease() : this(new NativeKeepAwakeApi()) { }
        // The test seam accepts compiled managed delegates, never PS callbacks.
        public KeepAwakeLease(Func<uint, uint> setState, Func<uint> threadId)
            : this(new CallbackKeepAwakeApi(setState, threadId)) { }
        public KeepAwakeLease(IKeepAwakeApi api)
        {
            if (api == null) { throw new ArgumentNullException("api"); }
            this.api = api;
        }

        public bool Start()
        {
            lock (gate)
            {
                if (stopped || eventsDisposed || worker != null) { return false; }
                worker = new Thread(Run);
                worker.IsBackground = true;
                worker.Name = "UPD temporary keep-awake";
                worker.Start();
            }
            // The local native calls normally return immediately. A failed or
            // delayed initialization remains advisory and is signalled to stop.
            if (!ready.WaitOne(5000)) { Stop(); return false; }
            return active;
        }

        private void Run()
        {
            bool affinity = false;
            uint previous = 0;
            try
            {
                Thread.BeginThreadAffinity();
                affinity = true;
                if (stop.WaitOne(0)) { return; }
                ActivationThreadId = api.GetThreadId();
                previous = api.SetExecutionState(Continuous | SystemRequired);
                if (previous == 0) { return; }
                PreviousState = previous;
                ActivationSucceeded = true;
                restored = false;
                active = true;
                ready.Set();
                stop.WaitOne();
            }
            catch
            {
                // OS/host failures are exposed only through fixed status fields.
            }
            finally
            {
                if (previous != 0)
                {
                    try
                    {
                        RestorationThreadId = api.GetThreadId();
                        restored = RestorationThreadId == ActivationThreadId &&
                            api.SetExecutionState(previous) != 0;
                    }
                    catch { restored = false; }
                }
                active = false;
                ready.Set();
                if (affinity)
                {
                    try { Thread.EndThreadAffinity(); } catch { restored = false; }
                }
            }
        }

        public bool Stop()
        {
            Thread current;
            lock (gate)
            {
                stopped = true;
                if (!eventsDisposed) { stop.Set(); }
                current = worker;
            }
            if (current != null && current.IsAlive && !current.Join(5000)) { return false; }
            lock (gate)
            {
                if (!eventsDisposed)
                {
                    ready.Dispose();
                    stop.Dispose();
                    eventsDisposed = true;
                }
            }
            return restored;
        }

        public void Dispose() { Stop(); }
    }
}
'@
}

function New-PodcastKeepAwakeLease {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Allocates an inactive owned lease; only Start after confirmed opt-in activates native state.')]
    [CmdletBinding()]
    param()

    return [UPD.KeepAwakeLease]::new()
}

function Start-PodcastKeepAwake {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'The explicit opt-in helper is called only after its caller accepts the operation.')]
    [CmdletBinding()]
    param([switch]$Enabled)

    $context = [pscustomobject]@{ Requested = [bool]$Enabled; Active = $false; Lease = $null; Message = 'Keep-awake is disabled.' }
    if (-not $Enabled) { return $context }
    $context.Message = 'Temporary keep-awake is unavailable; the download will continue without it.'
    try {
        if (-not (Test-PodcastKeepAwakePlatform)) { return $context }
        Initialize-PodcastKeepAwakeType
        $context.Lease = New-PodcastKeepAwakeLease
        $context.Active = $context.Lease.Start()
        if ($context.Active) { $context.Message = 'Temporary keep-awake is active for this confirmed operation.' }
        else { $null = $context.Lease.Stop() }
    }
    catch {
        $activationError = $_
        if ($null -ne $context.Lease) {
            try { $null = $context.Lease.Stop() } catch { $context.Active = $false }
        }
        if (Get-Command Test-PodcastCancellation -ErrorAction SilentlyContinue) {
            if (Test-PodcastCancellation -ErrorObject $activationError) { throw }
        }
        $context.Active = $false
    }
    return $context
}

function Stop-PodcastKeepAwake {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Releases only the temporary state owned by the confirmed-operation helper.')]
    [CmdletBinding()]
    param([AllowNull()]$Context)

    $requested = $null -ne $Context -and [bool]$Context.Requested
    $restored = $true
    if ($null -ne $Context -and $null -ne $Context.Lease) {
        try { $restored = [bool]$Context.Lease.Stop() }
        catch { $restored = $false }
        $Context.Active = $false
    }
    $message = 'Temporary keep-awake was released and its prior thread state restored.'
    if (-not $requested) { $message = 'Keep-awake is disabled.' }
    elseif (-not $restored) { $message = 'Temporary keep-awake cleanup could not be confirmed.' }
    elseif ($null -eq $Context.Lease) { $message = 'Temporary keep-awake is not active.' }
    return [pscustomobject]@{ Requested = $requested; Active = $false; Restored = $restored; Message = $message }
}
