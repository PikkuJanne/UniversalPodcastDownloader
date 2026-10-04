BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/NetworkPolicy.ps1')
    . (Join-Path $script:RepositoryRoot 'src/MediaRequest.ps1')
    . (Join-Path $script:RepositoryRoot 'src/RunResult.ps1')
    . (Join-Path $script:RepositoryRoot 'src/Diagnostics.ps1')
    Add-Type -AssemblyName System.Net.Http
    Mock Invoke-WebRequest { throw 'Media units must not use the external network.' }

    function Get-MediaTestTask {
        param($Value)
        $task = [pscustomobject]@{ Value = $Value }
        $task | Add-Member ScriptMethod GetAwaiter { return $this }
        $task | Add-Member ScriptMethod GetResult { return $this.Value }
        $task | Add-Member ScriptMethod Wait { param($milliseconds) return ($milliseconds -ge 0) }
        return $task
    }

    function Get-MediaTestClient {
        param([byte[]]$Body = @(1, 2, 3, 4), $Length = 4, [int]$Status = 200,
            [string[]]$Encoding = @(), [string]$Type = 'audio/mpeg', $Range = $null)
        $headers = [pscustomobject]@{ ContentLength = $Length; ContentType = $Type; ContentEncoding = $Encoding; ContentRange = $Range; HasLength = $null -ne $Length; HasRange = $null -ne $Range }
        $headers | Add-Member ScriptMethod Contains {
            param($name)
            if ($name -eq 'Content-Length') { return $this.HasLength }
            if ($name -eq 'Content-Range') { return $this.HasRange }
            return $false
        }
        $source = [IO.MemoryStream]::new($Body)
        $content = [pscustomobject]@{ Headers = $headers; Source = $source }
        $content | Add-Member ScriptMethod ReadAsStreamAsync { return (Get-MediaTestTask -Value $this.Source) }
        $response = [pscustomobject]@{ StatusCode = $Status; Content = $content; Disposed = $false }
        $response | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        $client = [pscustomobject]@{ Response = $response; Request = $null; Completion = $null; Calls = 0; Disposed = $false }
        $client | Add-Member ScriptMethod SendAsync {
            param($request, $completion)
            $this.Request = $request
            $this.Completion = $completion
            $this.Calls++
            return (Get-MediaTestTask -Value $this.Response)
        }
        $client | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        return $client
    }
}

Describe 'A040: media cleanup preserves primary outcomes' -Tag 'Unit', 'A040' {
    BeforeEach {
        $script:Client = Get-MediaTestClient
        $script:Destination = [IO.MemoryStream]::new()
        $script:Client.Response | Add-Member ScriptMethod Dispose {
            $this.Disposed = $true
            throw 'privateCleanupCanary'
        } -Force
        Mock Get-PodcastHttpClient { return $script:Client }
        Mock Write-Verbose {}
    }
    AfterEach {
        $script:Destination.Dispose()
        $script:Client.Response.Content.Source.Dispose()
    }

    It 'preserves typed cancellation after streamed bytes when response cleanup also fails' {
        $caught = $null
        try {
            $null = Invoke-PodcastMediaRequest -Uri 'https://media.invalid/audio.mp3' -DestinationStream $script:Destination -OnProgress {
                throw [OperationCanceledException]::new('privateProgressCancellationCanary')
            }
        }
        catch { $caught = $_ }
        $caught | Should -Not -BeNullOrEmpty
        Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue
        $script:Destination.ToArray() | Should -Be @(1, 2, 3, 4)
        $script:Destination.CanWrite | Should -BeTrue
        $script:Client.Disposed | Should -BeTrue
        $script:Client.Response.Disposed | Should -BeTrue
        $script:Client.Response.Content.Source.CanRead | Should -BeFalse
        Should -Invoke Write-Verbose -Times 1 -Exactly -ParameterFilter { $Message -eq 'Media request cleanup failed; the primary operation failure is preserved.' }
    }

    It 'retains an ordinary primary callback failure and its safe category when cleanup also fails' {
        $caught = $null
        try {
            $null = Invoke-PodcastMediaRequest -Uri 'https://media.invalid/audio.mp3' -DestinationStream $script:Destination -OnProgress {
                throw [InvalidOperationException]::new('privatePrimaryCanary https://media.invalid/private-token')
            }
        }
        catch { $caught = $_ }
        $caught | Should -Not -BeNullOrEmpty
        $caught.Exception.Message | Should -Match 'privatePrimaryCanary'
        $caught.Exception.Message | Should -Not -Match 'privateCleanupCanary|could not be closed'
        Test-PodcastCancellation -ErrorObject $caught | Should -BeFalse
        $message = Get-PodcastDiagnosticError -Error $caught
        $message | Should -Match 'operation failed'
        $message | Should -Not -Match 'privatePrimaryCanary|privateCleanupCanary|media.invalid|private-token'
        $script:Destination.ToArray() | Should -Be @(1, 2, 3, 4)
        $script:Client.Disposed | Should -BeTrue
        Should -Invoke Write-Verbose -Times 1 -Exactly -ParameterFilter { $Message -eq 'Media request cleanup failed; the primary operation failure is preserved.' }
    }

    It 'still fails a completed body when response resources cannot be closed safely' {
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/audio.mp3' -DestinationStream $script:Destination } |
            Should -Throw 'Media request resources could not be closed safely.'
        $script:Destination.ToArray() | Should -Be @(1, 2, 3, 4)
        $script:Destination.CanWrite | Should -BeTrue
        $script:Client.Disposed | Should -BeTrue
        $script:Client.Response.Content.Source.CanRead | Should -BeFalse
        Should -Invoke Write-Verbose -Times 0 -Exactly
    }
}

Describe 'A023: media transport shares the HTTP policy' -Tag 'Unit', 'A023' {
    BeforeEach {
        $script:Client = Get-MediaTestClient
        $script:Destination = [IO.MemoryStream]::new()
        $script:RedirectResponse = $null
        Mock Get-PodcastHttpClient { return $script:Client }
    }
    AfterEach {
        $script:Destination.Dispose()
        $script:Client.Response.Content.Source.Dispose()
        if ($null -ne $script:RedirectResponse) { $script:RedirectResponse.Content.Source.Dispose() }
    }

    It 'rejects unsupported media targets before creating a client or writing bytes' -ForEach @(
        @{ Uri = 'file:///private-path' },
        @{ Uri = 'https://private-user:private-token@media.invalid/episode.mp3' }
    ) {
        { Invoke-PodcastMediaRequest -Uri $Uri -DestinationStream $script:Destination } | Should -Throw
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
        $script:Destination.Length | Should -Be 0
        $script:Destination.CanWrite | Should -BeTrue
    }

    It 'downloads only the final body after an allowed media redirect' {
        $script:RedirectResponse = (Get-MediaTestClient -Status 302).Response
        $script:RedirectResponse | Add-Member NoteProperty Headers ([pscustomobject]@{ Location = [Uri]::new('/final.mp3', [UriKind]::Relative) })
        $script:Client | Add-Member NoteProperty RedirectResponse $script:RedirectResponse
        $script:Client | Add-Member ScriptMethod SendAsync {
            param($request, $completion)
            $this.Request = $request
            $this.Completion = $completion
            $this.Calls++
            if ($this.Calls -eq 1) { return (Get-MediaTestTask $this.RedirectResponse) }
            return (Get-MediaTestTask $this.Response)
        } -Force
        $result = Invoke-PodcastMediaRequest -Uri 'https://media.invalid/start.mp3' -DestinationStream $script:Destination
        $result.Bytes | Should -Be 4
        $script:Destination.ToArray() | Should -Be @(1, 2, 3, 4)
        $script:Client.Calls | Should -Be 2
        $script:Client.Request.RequestUri.AbsoluteUri | Should -Be 'https://media.invalid/final.mp3'
        $script:RedirectResponse.Disposed | Should -BeTrue
    }

    It 'rejects an HTTPS media downgrade before reading or writing a response body' {
        $script:Client.Response.StatusCode = 302
        $script:Client.Response | Add-Member NoteProperty Headers ([pscustomobject]@{ Location = [Uri]'http://media.invalid/final.mp3' })
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/start.mp3' -DestinationStream $script:Destination } | Should -Throw '*failed before completion*'
        $script:Client.Calls | Should -Be 1
        $script:Destination.Length | Should -Be 0
        $script:Client.Response.Disposed | Should -BeTrue
        $script:Client.Disposed | Should -BeTrue
    }
}

Describe 'A012/A013: media HTTP response streaming' -Tag 'Unit', 'A012', 'A013' {
    BeforeEach {
        $script:Client = Get-MediaTestClient
        $script:Destination = [IO.MemoryStream]::new()
        Mock Get-PodcastHttpClient { return $script:Client }
    }

    AfterEach {
        $script:Destination.Dispose()
        $script:Client.Response.Content.Source.Dispose()
    }

    It 'streams the exact body with one GET and returns the same response metadata' {
        $result = Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination
        $result.Completed | Should -BeTrue
        $result.StatusCode | Should -Be 200
        $result.Bytes | Should -Be 4
        $result.ContentLength | Should -Be 4
        $result.ContentType | Should -Be 'audio/mpeg'
        $result.ContentEncoding | Should -BeNullOrEmpty
        $script:Destination.ToArray() | Should -Be @(1, 2, 3, 4)
        $script:Destination.CanWrite | Should -BeTrue
        $script:Client.Request.Method.Method | Should -Be 'GET'
        @($script:Client.Request.Headers.AcceptEncoding)[0].Value | Should -Be 'identity'
        $script:Client.Completion | Should -Be ([Net.Http.HttpCompletionOption]::ResponseHeadersRead)
        $script:Client.Calls | Should -Be 1
        $script:Client.Disposed | Should -BeTrue
        $script:Client.Response.Disposed | Should -BeTrue
        $script:Client.Response.Content.Source.CanRead | Should -BeFalse
    }

    It 'allows complete EOF-terminated bodies without an advertised length' {
        $script:Client = Get-MediaTestClient -Length $null
        $result = Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination
        $result.Bytes | Should -Be 4
        $result.ContentLength | Should -BeNullOrEmpty
    }

    It 'rejects truncated or oversized bodies even when the source silently returns EOF' -ForEach @(
        @{ Length = 5 }, @{ Length = 3 }
    ) {
        $script:Client = Get-MediaTestClient -Length $Length
        $failure = $null
        try { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination }
        catch { $failure = Get-PodcastTransportFailure -ErrorObject $_ }
        $failure.Message | Should -BeLike '*Content-Length*byte count*'
        $failure.Data['Retryable'] | Should -Be (4 -lt $Length)
        $script:Client.Disposed | Should -BeTrue
        $script:Client.Response.Disposed | Should -BeTrue
        $script:Client.Response.Content.Source.CanRead | Should -BeFalse
        $script:Destination.CanWrite | Should -BeTrue
    }

    It 'rejects unsolicited partial and other non-200 statuses before reading' -ForEach @(
        @{ Status = 206 }, @{ Status = 204 }, @{ Status = 201 }, @{ Status = 404 }, @{ Status = 500 }
    ) {
        $script:Client = Get-MediaTestClient -Status $Status
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } | Should -Throw '*complete HTTP 200*'
        $script:Destination.Length | Should -Be 0
        $script:Client.Disposed | Should -BeTrue
        $script:Client.Response.Disposed | Should -BeTrue
    }

    It 'rejects Content-Range even when the response status says 200' {
        $script:Client = Get-MediaTestClient -Range 'bytes 0-3/10'
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } | Should -Throw '*partial responses*'
        $script:Destination.Length | Should -Be 0
    }

    It 'rejects unparsed Content-Length instead of treating it as EOF framing' {
        $script:Client = Get-MediaTestClient -Length $null
        $script:Client.Response.Content.Headers.HasLength = $true
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } | Should -Throw '*invalid Content-Length*'
    }

    It 'rejects an unparsed Content-Range header on a 200 response' {
        $script:Client.Response.Content.Headers.HasRange = $true
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } | Should -Throw '*partial responses*'
        $script:Destination.Length | Should -Be 0
    }

    It 'rejects content encoding that would make original-audio byte comparisons ambiguous' -ForEach @(
        @{ Encoding = 'gzip' }, @{ Encoding = 'br' }, @{ Encoding = 'deflate' }
    ) {
        $script:Client = Get-MediaTestClient -Encoding @($Encoding)
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } | Should -Throw '*unsupported Content-Encoding*'
        $script:Destination.Length | Should -Be 0
    }

    It 'accepts explicit identity encoding and leaves content validation to its caller' {
        $script:Client = Get-MediaTestClient -Encoding @('identity') -Type 'text/html'
        $result = Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination
        $result.ContentEncoding | Should -Be 'identity'
        $result.ContentType | Should -Be 'text/html'
    }

    It 'reports a completed empty transfer for the caller to reject as invalid media' {
        $script:Client = Get-MediaTestClient -Body @() -Length 0
        $result = Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination
        $result.Completed | Should -BeTrue
        $result.Bytes | Should -Be 0
    }

    It 'does not expose a raw request exception or its synthetic secret' {
        $script:Client | Add-Member ScriptMethod SendAsync { throw 'secret-token from https://media.invalid/private?secret-token' } -Force
        $errorRecord = $null
        try { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/private?secret-token' -DestinationStream $script:Destination }
        catch { $errorRecord = $_ }
        $errorRecord.Exception.Message | Should -Be 'HTTP request failed before response headers were available.'
        $errorRecord.Exception.Message | Should -Not -Match 'secret-token|media.invalid'
        $script:Client.Disposed | Should -BeTrue
        $script:Destination.CanWrite | Should -BeTrue
    }

    It 'closes request resources after a destination write failure without closing caller ownership' {
        $script:Destination.Dispose()
        $script:Destination = [IO.MemoryStream]::new([byte[]]@(0), $true)
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } | Should -Throw '*failed before completion*'
        $script:Client.Disposed | Should -BeTrue
        $script:Client.Response.Disposed | Should -BeTrue
        $script:Client.Response.Content.Source.CanRead | Should -BeFalse
        $script:Destination.CanWrite | Should -BeTrue
    }

    It 'closes response resources when reading fails and hides the raw read exception' {
        $script:Client.Response.Content.Source.Dispose()
        $source = [pscustomobject]@{ Disposed = $false }
        $source | Add-Member ScriptMethod ReadAsync { throw 'synthetic-secret in response stream exception' }
        $source | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        $script:Client.Response.Content.Source = $source
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } |
            Should -Throw 'The response body could not be read completely.'
        $source.Disposed | Should -BeTrue
        $script:Client.Response.Disposed | Should -BeTrue
        $script:Client.Disposed | Should -BeTrue
        $script:Destination.CanWrite | Should -BeTrue
    }
}
