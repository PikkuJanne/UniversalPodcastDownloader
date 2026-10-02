BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/NetworkPolicy.ps1')
    $script:ClientFactory = (Get-Command Get-PodcastHttpClient -CommandType Function).ScriptBlock
    Add-Type -AssemblyName System.Net.Http
    Mock Get-PodcastHttpClient { throw 'Transport units must not open a network client.' }

    function New-TransportTestResponse {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory response fixture.')]
        [CmdletBinding()]
        param([int]$Status = 200, [string]$RetryAfter, [string]$Location, [byte[]]$Body = @(60, 114, 47, 62))
        $response = [Net.Http.HttpResponseMessage]::new([Enum]::ToObject([Net.HttpStatusCode], $Status))
        $response.Content = [Net.Http.ByteArrayContent]::new($Body)
        if ($RetryAfter) { $null = $response.Headers.TryAddWithoutValidation('Retry-After', $RetryAfter) }
        if ($Location) { $response.Headers.Location = [Uri]::new($Location, [UriKind]::RelativeOrAbsolute) }
        return $response
    }

    function New-TransportTestClient {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory client fixture.')]
        [CmdletBinding()]
        param([object[]]$Responses)
        $client = [pscustomobject]@{ Responses = $Responses; Calls = 0; Disposed = $false; Token = $null }
        $client | Add-Member ScriptMethod SendAsync {
            param($request, $completion, $token)
            $null = $request; $null = $completion
            $this.Token = $token
            $value = $this.Responses[[Math]::Min($this.Calls, $this.Responses.Count - 1)]
            $this.Calls++
            $task = [Threading.Tasks.TaskCompletionSource[object]]::new()
            $task.SetResult($value)
            return $task.Task
        }
        $client | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        return $client
    }
}

Describe 'A025: bounded transport retry decisions' -Tag 'Unit', 'A025' {
    BeforeEach {
        $script:Now = [DateTimeOffset]'2026-10-02T10:00:00Z'
        $script:Waits = [Collections.Generic.List[int]]::new()
        $script:Calls = 0
        $script:Policy = New-PodcastTransportPolicy -Clock { $script:Now } -Delay {
            param($Milliseconds)
            $script:Waits.Add($Milliseconds)
            $script:Now = $script:Now.AddMilliseconds($Milliseconds)
        }
    }

    It 'uses finite configurable defaults with independent header and idle limits' {
        $policy = New-PodcastTransportPolicy
        $policy.MaxAttempts | Should -Be 3
        $policy.HeaderTimeoutSeconds | Should -Be 30
        $policy.IdleTimeoutSeconds | Should -Be 30
        $policy.RetryBudgetSeconds | Should -Be 120
        $policy.BaseDelaySeconds | Should -Be 1
        $policy.MaxDelaySeconds | Should -Be 30
        { New-PodcastTransportPolicy -MaxAttempts 0 } | Should -Throw
        { New-PodcastTransportPolicy -IdleTimeoutSeconds 0 } | Should -Throw
        { New-PodcastTransportPolicy -HeaderTimeoutSeconds 0 } | Should -Throw
        { New-PodcastTransportPolicy -RetryBudgetSeconds -1 } | Should -Throw
    }

    It 'bounds transient attempts and exponential waits' {
        $caught = $null
        try {
            Invoke-PodcastTransportOperation -Policy $script:Policy -Operation {
                $script:Calls++
                throw (New-PodcastTransportException -Kind Connection -Message 'Safe connection failure.' -Retryable $true)
            }
        }
        catch { $caught = Get-PodcastTransportFailure -ErrorObject $_ }
        $script:Calls | Should -Be 3
        $script:Waits.ToArray() | Should -Be @(1000, 2000)
        $caught.Data['Attempts'] | Should -Be 3
        $caught.InnerException | Should -BeNullOrEmpty
    }

    It 'caps local backoff but preserves the successful result shape' {
        $script:Policy.MaxAttempts = 4
        $script:Policy.MaxDelaySeconds = 1.5
        $result = Invoke-PodcastTransportOperation -Policy $script:Policy -Operation {
            param($Attempt)
            if ($Attempt -lt 4) { throw (New-PodcastTransportException -Kind Connection -Message 'Safe connection failure.' -Retryable $true) }
            return [pscustomobject]@{ Completed = $true; Bytes = 4 }
        }
        $script:Waits.ToArray() | Should -Be @(1000, 1500, 1500)
        @($result.PSObject.Properties.Name) | Should -Be @('Completed', 'Bytes')
        $result.Bytes | Should -Be 4
    }

    It 'never retries permanent transport or unclassified operation errors' -ForEach @(@{ Tagged = $true }, @{ Tagged = $false }) {
        {
            Invoke-PodcastTransportOperation -Policy $script:Policy -Operation {
                $script:Calls++
                if ($Tagged) { throw (New-PodcastTransportException -Kind Permanent -Message 'Safe permanent failure.') }
                throw 'Safe local write failure.'
            }
        } | Should -Throw
        $script:Calls | Should -Be 1
        $script:Waits.Count | Should -Be 0
    }

    It 'waits at least a fractional server deadline even after a short wakeup' {
        $script:Policy.Delay = {
            param($Milliseconds)
            $script:Waits.Add($Milliseconds)
            if ($script:Waits.Count -eq 1) { $script:Now = $script:Now.AddMilliseconds(250) }
            else { $script:Now = $script:Now.AddMilliseconds($Milliseconds) }
        }
        $script:Required = $script:Now.AddSeconds(1.001)
        $result = Invoke-PodcastTransportOperation -Policy $script:Policy -Operation {
            param($Attempt)
            if ($Attempt -eq 1) { throw (New-PodcastTransportException -Kind HttpStatus -Message 'Retry later.' -Retryable $true -RetryAfterUtc $script:Required) }
            return $script:Now
        }
        $result | Should -BeGreaterOrEqual $script:Required
        $script:Waits.ToArray() | Should -Be @(1001, 751)
    }

    It 'defers without waiting or retrying when server delay exceeds the budget' {
        $caught = $null
        try {
            Invoke-PodcastTransportOperation -Policy $script:Policy -Operation {
                $script:Calls++
                throw (New-PodcastTransportException -Kind HttpStatus -Message 'Retry later.' -Retryable $true -RetryAfterUtc $script:Now.AddSeconds(121))
            }
        }
        catch { $caught = Get-PodcastTransportFailure -ErrorObject $_ }
        $caught.Data['Kind'] | Should -Be 'Deferred'
        $caught.Data['Attempts'] | Should -Be 1
        $script:Calls | Should -Be 1
        $script:Waits.Count | Should -Be 0
    }

    It 'does not start an attempt after scheduler oversleep exhausts the budget' {
        $script:Policy.Delay = { param($Milliseconds) $null = $Milliseconds; $script:Now = $script:Now.AddSeconds(121) }
        {
            Invoke-PodcastTransportOperation -Policy $script:Policy -Operation {
                $script:Calls++
                throw (New-PodcastTransportException -Kind Connection -Message 'Retry later.' -Retryable $true)
            }
        } | Should -Throw '*budget expired*'
        $script:Calls | Should -Be 1
    }

    It 'fails safely if the injected clock does not advance instead of retrying early' {
        $script:Policy.Delay = { param($Milliseconds) $null = $Milliseconds }
        { Invoke-PodcastTransportOperation -Policy $script:Policy -Operation { throw (New-PodcastTransportException -Kind Connection -Message 'Retry later.' -Retryable $true) } } |
            Should -Throw '*clock did not advance*'
    }

    It 'allows a successful active operation to outlast the retry budget' {
        $result = Invoke-PodcastTransportOperation -Policy $script:Policy -Operation { $script:Now = $script:Now.AddMinutes(10); return 'complete' }
        $result | Should -Be 'complete'
        $script:Waits.Count | Should -Be 0
    }

    It 'uses elapsed operation time when deciding whether another attempt fits' {
        { Invoke-PodcastTransportOperation -Policy $script:Policy -Operation {
            $script:Calls++
            $script:Now = $script:Now.AddMinutes(10)
            throw (New-PodcastTransportException -Kind Connection -Message 'Retry later.' -Retryable $true)
        } } | Should -Throw '*retry budget*'
        $script:Calls | Should -Be 1
    }

    It 'classifies wrapped connection failures and keeps TLS failures permanent' {
        $socket = [Net.Sockets.SocketException]::new([int][Net.Sockets.SocketError]::ConnectionReset)
        (Test-PodcastTransientException -ErrorObject ([IO.IOException]::new('secret-token', $socket))) | Should -BeTrue
        (Test-PodcastTransientException -ErrorObject ([Net.WebException]::new('secret-token', [Net.WebExceptionStatus]::TrustFailure))) | Should -BeFalse
        (Test-PodcastTransientException -ErrorObject ([Security.Authentication.AuthenticationException]::new('secret-token', $socket))) | Should -BeFalse
        (Test-PodcastTransientException -ErrorObject ([IO.IOException]::new('secret-token disk full'))) | Should -BeFalse
    }

    It 'retries a Framework bare IOException only at the HTTP response-read boundary' {
        $networkRead = [AggregateException]::new([IO.IOException]::new('synthetic-secret connection closure'))
        (Test-PodcastTransientException -ErrorObject $networkRead -ResponseRead) | Should -BeTrue
        (Test-PodcastTransientException -ErrorObject $networkRead) | Should -BeFalse
        $tlsRead = [IO.IOException]::new('synthetic-secret', [Security.Authentication.AuthenticationException]::new('synthetic-secret'))
        (Test-PodcastTransientException -ErrorObject $tlsRead -ResponseRead) | Should -BeFalse
    }
}

Describe 'A025: HTTP retry status and Retry-After policy' -Tag 'Unit', 'A025' {
    BeforeEach {
        $script:Now = [DateTimeOffset]'2026-10-02T10:00:00Z'
        $script:Waits = [Collections.Generic.List[int]]::new()
        $script:Policy = New-PodcastTransportPolicy -Clock { $script:Now } -Delay {
            param($Milliseconds)
            $script:Waits.Add($Milliseconds)
            $script:Now = $script:Now.AddMilliseconds($Milliseconds)
        }
        $script:Responses = @()
    }
    AfterEach { foreach ($response in $script:Responses) { $response.Dispose() } }

    It 'classifies HTTP <Status> with Retryable=<Retryable>' -ForEach @(
        @{ Status = 408; Retryable = $true }, @{ Status = 429; Retryable = $true },
        @{ Status = 500; Retryable = $true }, @{ Status = 502; Retryable = $true },
        @{ Status = 503; Retryable = $true }, @{ Status = 504; Retryable = $true },
        @{ Status = 400; Retryable = $false }, @{ Status = 401; Retryable = $false },
        @{ Status = 403; Retryable = $false }, @{ Status = 404; Retryable = $false },
        @{ Status = 410; Retryable = $false }, @{ Status = 501; Retryable = $false },
        @{ Status = 505; Retryable = $false }
    ) {
        $script:Responses = @((New-TransportTestResponse -Status $Status))
        $client = New-TransportTestClient -Responses $script:Responses
        $caught = $null
        try { Invoke-PodcastHttpGet -Uri 'https://feed.invalid/?secret-token' -Client $client -Policy $script:Policy }
        catch { $caught = Get-PodcastTransportFailure -ErrorObject $_ }
        $caught.Data['Retryable'] | Should -Be $Retryable
        $caught.Data['StatusCode'] | Should -Be $Status
        $caught.Message | Should -Not -Match 'secret-token|feed.invalid'
        $client.Calls | Should -Be 1
    }

    It 'honors delta and all three HTTP date forms <RetryAfter>' -ForEach @(
        @{ RetryAfter = '3' },
        @{ RetryAfter = 'Fri, 02 Oct 2026 10:00:03 GMT' },
        @{ RetryAfter = 'Friday, 02-Oct-26 10:00:03 GMT' },
        @{ RetryAfter = 'Fri Oct  2 10:00:03 2026' }
    ) {
        $script:Responses = @((New-TransportTestResponse -Status 503 -RetryAfter $RetryAfter), (New-TransportTestResponse))
        $client = New-TransportTestClient -Responses $script:Responses
        $exchange = Invoke-PodcastTransportOperation -Policy $script:Policy -Operation { Invoke-PodcastHttpGet -Uri 'https://feed.invalid/' -Client $client -Policy $script:Policy }
        $client.Calls | Should -Be 2
        $script:Waits.ToArray() | Should -Be @(3000)
        $exchange.Request.Dispose()
    }

    It 'treats numeric overflow as defer instead of retrying with a short fallback' {
        $script:Responses = @((New-TransportTestResponse -Status 503 -RetryAfter '999999999999999999999999999999'))
        $client = New-TransportTestClient -Responses $script:Responses
        { Invoke-PodcastTransportOperation -Policy $script:Policy -Operation { Invoke-PodcastHttpGet -Uri 'https://feed.invalid/' -Client $client -Policy $script:Policy } } |
            Should -Throw '*deferred*'
        $client.Calls | Should -Be 1
        $script:Waits.Count | Should -Be 0
    }

    It 'uses local backoff for malformed or past Retry-After <RetryAfter>' -ForEach @(
        @{ RetryAfter = 'not-a-date secret-token' }, @{ RetryAfter = '-5' }, @{ RetryAfter = 'Thu, 01 Oct 2026 10:00:00 GMT' }
    ) {
        $script:Responses = @((New-TransportTestResponse -Status 503 -RetryAfter $RetryAfter), (New-TransportTestResponse))
        $client = New-TransportTestClient -Responses $script:Responses
        $exchange = Invoke-PodcastTransportOperation -Policy $script:Policy -Operation { Invoke-PodcastHttpGet -Uri 'https://feed.invalid/' -Client $client -Policy $script:Policy }
        $script:Waits.ToArray() | Should -Be @(1000)
        $exchange.Request.Dispose()
    }

    It 'obeys Retry-After before a redirect hop without counting a new attempt' {
        $script:Responses = @((New-TransportTestResponse -Status 302 -RetryAfter '3' -Location '/final'), (New-TransportTestResponse))
        $client = New-TransportTestClient -Responses $script:Responses
        $exchange = Invoke-PodcastHttpGet -Uri 'https://feed.invalid/' -Client $client -Policy $script:Policy
        $client.Calls | Should -Be 2
        $script:Waits.ToArray() | Should -Be @(3000)
        $exchange.RedirectCount | Should -Be 1
        $exchange.Request.Dispose()
    }

    It 'defers a redirect whose server delay exceeds the budget' {
        $script:Responses = @((New-TransportTestResponse -Status 302 -RetryAfter '121' -Location '/final'))
        $client = New-TransportTestClient -Responses $script:Responses
        { Invoke-PodcastHttpGet -Uri 'https://feed.invalid/' -Client $client -Policy $script:Policy } | Should -Throw '*deferred*'
        $client.Calls | Should -Be 1
        $script:Waits.Count | Should -Be 0
    }

    It 'shares one deadline across transient retries and later redirect waits' {
        $script:Policy.RetryBudgetSeconds = 10
        $script:Responses = @((New-TransportTestResponse -Status 503 -RetryAfter '5'),
            (New-TransportTestResponse -Status 302 -RetryAfter '6' -Location '/final'))
        $client = New-TransportTestClient -Responses $script:Responses
        {
            Invoke-PodcastTransportOperation -Policy $script:Policy -Operation {
                param($Attempt, $AttemptPolicy)
                $null = $Attempt
                Invoke-PodcastHttpGet -Uri 'https://feed.invalid/' -Client $client -Policy $AttemptPolicy
            }
        } | Should -Throw '*deferred*'
        $client.Calls | Should -Be 2
        $script:Waits.ToArray() | Should -Be @(5000)
    }

    It 'restarts metadata in a fresh buffer after an incomplete response' {
        $script:Responses = @((New-TransportTestResponse -Body @(60)), (New-TransportTestResponse))
        $script:Responses[0].Content.Headers.ContentLength = 4
        $script:Client = New-TransportTestClient -Responses $script:Responses
        Mock Get-PodcastHttpClient { return $script:Client }
        $result = Invoke-PodcastMetadataRequest -Uri 'https://feed.invalid/' -Policy $script:Policy
        $result.Content | Should -BeExactly '<r/>'
        $result.Bytes | Should -Be 4
        $script:Client.Calls | Should -Be 2
        $script:Waits.ToArray() | Should -Be @(1000)
    }
}

Describe 'A026: independent bounded header and idle reads' -Tag 'Unit', 'A026' {
    It 'disables the HttpClient total timeout in favor of the explicit limits' {
        $client = & $script:ClientFactory
        try { $client.Timeout | Should -Be ([Threading.Timeout]::InfiniteTimeSpan) }
        finally { $client.Dispose() }
    }

    It 'bounds headers and cancels the pending request without raw diagnostics' {
        $client = [pscustomobject]@{ Token = $null }
        $client | Add-Member ScriptMethod SendAsync {
            param($request, $completion, $token)
            $null = $request; $null = $completion
            $this.Token = $token
            return ([Threading.Tasks.TaskCompletionSource[object]]::new()).Task
        }
        $caught = $null
        $watch = [Diagnostics.Stopwatch]::StartNew()
        try { Invoke-PodcastHttpGet -Uri 'https://feed.invalid/?secret-token' -Client $client -Policy (New-PodcastTransportPolicy -HeaderTimeoutSeconds 0.02) }
        catch { $caught = Get-PodcastTransportFailure -ErrorObject $_ }
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 2
        $caught.Data['Kind'] | Should -Be 'HeaderTimeout'
        $caught.Data['Retryable'] | Should -BeTrue
        $client.Token.IsCancellationRequested | Should -BeTrue
        $caught.Message | Should -Not -Match 'secret-token|feed.invalid'
    }

    It 'bounds token-ignoring body reads by disposing their source' {
        $source = [pscustomobject]@{ Token = $null; Disposed = $false }
        $source | Add-Member ScriptMethod ReadAsync {
            param($buffer, $offset, $count, $token)
            $null = $buffer; $null = $offset; $null = $count
            $this.Token = $token
            return ([Threading.Tasks.TaskCompletionSource[int]]::new()).Task
        }
        $source | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        $caught = $null
        $watch = [Diagnostics.Stopwatch]::StartNew()
        try { Read-PodcastResponseChunk -Source $source -Buffer (New-Object byte[] 16) -Count 16 -Policy (New-PodcastTransportPolicy -IdleTimeoutSeconds 0.02) }
        catch { $caught = Get-PodcastTransportFailure -ErrorObject $_ }
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 2
        $caught.Data['Kind'] | Should -Be 'IdleTimeout'
        $caught.Data['Retryable'] | Should -BeTrue
        $source.Disposed | Should -BeTrue
        $source.Token.IsCancellationRequested | Should -BeTrue
    }

    It 'returns received chunks and EOF without closing a healthy stream' {
        $source = [IO.MemoryStream]::new([byte[]]@(1, 2, 3))
        try {
            $buffer = New-Object byte[] 2
            (Read-PodcastResponseChunk -Source $source -Buffer $buffer -Count 2) | Should -Be 2
            $buffer | Should -Be @(1, 2)
            (Read-PodcastResponseChunk -Source $source -Buffer $buffer -Count 2) | Should -Be 1
            $buffer[0] | Should -Be 3
            (Read-PodcastResponseChunk -Source $source -Buffer $buffer -Count 2) | Should -Be 0
            $source.CanRead | Should -BeTrue
        }
        finally { $source.Dispose() }
    }
}
