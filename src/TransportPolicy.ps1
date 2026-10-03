#requires -Version 5.1

. (Join-Path $PSScriptRoot 'RunResult.ps1')

function New-PodcastTransportPolicy {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only an in-memory policy value.')]
    [CmdletBinding()]
    param(
        [ValidateRange(1, 10)][int]$MaxAttempts = 3,
        [ValidateRange(0.001, 86400)][double]$HeaderTimeoutSeconds = 30,
        [ValidateRange(0.001, 86400)][double]$IdleTimeoutSeconds = 30,
        [ValidateRange(0, 86400)][double]$RetryBudgetSeconds = 120,
        [ValidateRange(0, 3600)][double]$BaseDelaySeconds = 1,
        [ValidateRange(0, 3600)][double]$MaxDelaySeconds = 30,
        [scriptblock]$Clock = { [DateTimeOffset]::UtcNow },
        [scriptblock]$Delay = { param([int]$Milliseconds) Start-Sleep -Milliseconds $Milliseconds }
    )

    return [pscustomobject]@{
        MaxAttempts = $MaxAttempts
        HeaderTimeoutSeconds = $HeaderTimeoutSeconds
        IdleTimeoutSeconds = $IdleTimeoutSeconds
        RetryBudgetSeconds = $RetryBudgetSeconds
        BaseDelaySeconds = $BaseDelaySeconds
        MaxDelaySeconds = $MaxDelaySeconds
        Clock = $Clock
        Delay = $Delay
    }
}

function New-PodcastTransportException {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only an in-memory safe exception.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('HeaderTimeout', 'IdleTimeout', 'Connection', 'HttpStatus', 'IncompleteBody', 'Deferred', 'Permanent')][string]$Kind,
        [Parameter(Mandatory)][string]$Message,
        [bool]$Retryable = $false,
        [Nullable[DateTimeOffset]]$RetryAfterUtc,
        [int]$StatusCode = 0
    )

    # Do not retain an inner exception: platform messages can contain a signed URL.
    $failure = [InvalidOperationException]::new($Message)
    $failure.Data['PodcastTransport'] = $true
    $failure.Data['Kind'] = $Kind
    $failure.Data['Retryable'] = $Retryable
    if ($null -ne $RetryAfterUtc) { $failure.Data['RetryAfterUtc'] = $RetryAfterUtc }
    if ($StatusCode -gt 0) { $failure.Data['StatusCode'] = $StatusCode }
    return $failure
}

function Get-PodcastTransportFailure {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorObject)

    $exception = if ($ErrorObject -is [Management.Automation.ErrorRecord]) { $ErrorObject.Exception } else { $ErrorObject }
    for ($depth = 0; $null -ne $exception -and $depth -lt 16; $depth++) {
        if ($exception -is [Exception] -and $exception.Data['PodcastTransport'] -eq $true) { return $exception }
        $exception = $exception.InnerException
    }
    return $null
}

function Test-PodcastTransientException {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorObject, [switch]$ResponseRead)

    $exception = if ($ErrorObject -is [Management.Automation.ErrorRecord]) { $ErrorObject.Exception } else { $ErrorObject }
    $transient = $false
    for ($depth = 0; $null -ne $exception -and $depth -lt 16; $depth++) {
        if ($exception -is [Security.Authentication.AuthenticationException]) { return $false }
        if ($exception -is [Net.WebException]) {
            if (@('TrustFailure', 'SecureChannelFailure') -contains [string]$exception.Status) { return $false }
            if (@('Timeout', 'ConnectFailure', 'NameResolutionFailure', 'ProxyNameResolutionFailure', 'ConnectionClosed', 'ReceiveFailure', 'SendFailure', 'KeepAliveFailure') -contains [string]$exception.Status) { $transient = $true }
        }
        if ($exception -is [Net.Sockets.SocketException]) {
            if (@('ConnectionReset', 'ConnectionAborted', 'ConnectionRefused', 'TimedOut', 'HostNotFound', 'TryAgain', 'NetworkDown', 'NetworkUnreachable', 'HostUnreachable') -contains [string]$exception.SocketErrorCode) { $transient = $true }
        }
        # HttpRequestError exists only on recent .NET; property inspection keeps
        # the adapter compatible with .NET Framework without parsing messages.
        if ($null -ne $exception.PSObject.Properties['HttpRequestError']) {
            if ([string]$exception.HttpRequestError -eq 'SecureConnectionError') { return $false }
            if (@('NameResolutionError', 'ConnectionError', 'ResponseEnded') -contains [string]$exception.HttpRequestError) { $transient = $true }
        }
        if ($exception -is [IO.EndOfStreamException]) { $transient = $true }
        # Framework response streams can report a closed connection as a bare
        # IOException. This applies only at the HTTP source-read boundary;
        # destination writes and other local I/O failures must not be retried.
        if ($ResponseRead -and $exception -is [IO.IOException]) { $transient = $true }
        $exception = $exception.InnerException
    }
    return $transient
}

function Get-PodcastRetryAfterUtc {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Response, [Parameter(Mandatory)]$Policy)

    # Check raw numeric overflow before the typed parser discards it. A server
    # asking for an enormous delay must never receive an early local-backoff retry.
    if ($null -ne $Response.Headers -and $null -ne $Response.Headers.PSObject.Methods['TryGetValues']) {
        $values = $null
        if ($Response.Headers.TryGetValues('Retry-After', [ref]$values)) {
            foreach ($value in $values) {
                if ($value -match '^\s*[0-9]+\s*$') {
                    [long]$seconds = 0
                    if (-not [long]::TryParse($value.Trim(), [ref]$seconds) -or $seconds -gt 86400) { return [DateTimeOffset]::MaxValue }
                }
            }
        }
    }
    # Invalid headers fall back to local backoff; no raw header enters diagnostics.
    $retryAfter = $null
    try { $retryAfter = $Response.Headers.RetryAfter }
    catch { Write-Verbose 'The response Retry-After header was invalid.' }
    if ($null -eq $retryAfter) { return $null }
    if ($null -ne $retryAfter.Delta) {
        $now = [DateTimeOffset](& $Policy.Clock)
        try { return $now.Add($retryAfter.Delta) }
        catch { return [DateTimeOffset]::MaxValue }
    }
    if ($null -ne $retryAfter.Date) { return [DateTimeOffset]$retryAfter.Date }
    return $null
}

function Wait-PodcastRetryDelay {
    [CmdletBinding()]
    param([Parameter(Mandatory)][DateTimeOffset]$NotBefore, [Parameter(Mandatory)][DateTimeOffset]$Deadline,
        [Parameter(Mandatory)]$Policy, [DateTimeOffset]$ObservedUtc, [ValidateRange(1, 10)][int]$Attempts)

    # A retry owner supplies its existing observation so consolidation does not
    # add a clock sample or alter zero-delay/deadline decisions.
    $now = if ($PSBoundParameters.ContainsKey('ObservedUtc')) { $ObservedUtc } else { [DateTimeOffset](& $Policy.Clock) }
    if ($NotBefore -le $now) { return }
    if ($NotBefore -gt $Deadline) {
        $deferred = New-PodcastTransportException -Kind Deferred -Message 'The request was deferred because the required retry delay exceeds the retry budget.'
        if ($PSBoundParameters.ContainsKey('Attempts')) { $deferred.Data['Attempts'] = $Attempts }
        throw $deferred
    }
    while ($now -lt $NotBefore) {
        $milliseconds = [int][Math]::Ceiling(($NotBefore - $now).TotalMilliseconds)
        $null = & $Policy.Delay $milliseconds
        $afterDelay = [DateTimeOffset](& $Policy.Clock)
        if ($afterDelay -le $now) {
            $deferred = New-PodcastTransportException -Kind Deferred -Message 'The request was deferred because the retry clock did not advance.'
            if ($PSBoundParameters.ContainsKey('Attempts')) { $deferred.Data['Attempts'] = $Attempts }
            throw $deferred
        }
        $now = $afterDelay
    }
    if ($now -gt $Deadline) {
        $deferred = New-PodcastTransportException -Kind Deferred -Message 'The request was deferred because the retry budget expired while waiting.'
        if ($PSBoundParameters.ContainsKey('Attempts')) { $deferred.Data['Attempts'] = $Attempts }
        throw $deferred
    }
}

function Invoke-PodcastTransportOperation {
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Operation, $Policy = (New-PodcastTransportPolicy))

    $started = [DateTimeOffset](& $Policy.Clock)
    $deadline = $started.AddSeconds($Policy.RetryBudgetSeconds)
    $attemptPolicy = $Policy.PSObject.Copy()
    $attemptPolicy | Add-Member NoteProperty RetryDeadlineUtc $deadline -Force
    for ($attempt = 1; $attempt -le $Policy.MaxAttempts; $attempt++) {
        try { return (& $Operation $attempt $attemptPolicy) }
        catch {
            $failure = Get-PodcastTransportFailure -ErrorObject $_
            if ($null -eq $failure) {
                $_.Exception.Data['PodcastAttempts'] = $attempt
                throw
            }
            $failure.Data['Attempts'] = $attempt
            if (-not $failure.Data['Retryable'] -or $attempt -ge $Policy.MaxAttempts) { throw $failure }

            $now = [DateTimeOffset](& $Policy.Clock)
            $seconds = [Math]::Min($Policy.MaxDelaySeconds, $Policy.BaseDelaySeconds * [Math]::Pow(2, $attempt - 1))
            $next = $now.AddSeconds($seconds)
            if ($failure.Data.Contains('RetryAfterUtc') -and [DateTimeOffset]$failure.Data['RetryAfterUtc'] -gt $next) {
                $next = [DateTimeOffset]$failure.Data['RetryAfterUtc']
            }
            if ($Policy.RetryBudgetSeconds -le 0 -or $now -ge $deadline -or $next -gt $deadline) {
                $deferred = New-PodcastTransportException -Kind Deferred -Message 'The request was deferred because the required retry delay exceeds the retry budget.'
                $deferred.Data['Attempts'] = $attempt
                throw $deferred
            }
            # Redirects and failed-attempt backoff share the same rounded waits,
            # short-wakeup checks and deadline enforcement. Only failures made
            # by that helper acquire attempt metadata; injected errors pass through.
            Wait-PodcastRetryDelay -NotBefore $next -Deadline $deadline -Policy $Policy -ObservedUtc $now -Attempts $attempt
        }
    }
}

function Read-PodcastResponseChunk {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Source,
        [Parameter(Mandatory)][byte[]]$Buffer,
        [Parameter(Mandatory)][ValidateRange(1, 2147483647)][int]$Count,
        $Policy = (New-PodcastTransportPolicy)
    )

    $cancellation = [Threading.CancellationTokenSource]::new()
    $timedOut = $false
    try {
        $milliseconds = [int][Math]::Ceiling($Policy.IdleTimeoutSeconds * 1000)
        $readTask = $Source.ReadAsync($Buffer, 0, $Count, $cancellation.Token)
        if (-not $readTask.Wait($milliseconds)) {
            $timedOut = $true
            $cancellation.Cancel()
            # Framework response streams may ignore cancellation tokens. Closing
            # this response stream aborts that pending read before any retry.
            $Source.Dispose()
            throw (New-PodcastTransportException -Kind IdleTimeout -Message 'The response body exceeded the idle transfer timeout.' -Retryable $true)
        }
        return $readTask.GetAwaiter().GetResult()
    }
    catch {
        $known = Get-PodcastTransportFailure -ErrorObject $_
        if ($null -ne $known) { throw $known }
        if ($timedOut) { throw (New-PodcastTransportException -Kind IdleTimeout -Message 'The response body exceeded the idle transfer timeout.' -Retryable $true) }
        if (Test-PodcastCancellation -ErrorObject $_) { throw }
        throw (New-PodcastTransportException -Kind Connection -Message 'The response body could not be read completely.' -Retryable (Test-PodcastTransientException -ErrorObject $_ -ResponseRead))
    }
    finally { $cancellation.Dispose() }
}
