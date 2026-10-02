BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/MediaRequest.ps1')
    Add-Type -AssemblyName System.Net.Http
    Mock Invoke-WebRequest { throw 'Media units must not use the external network.' }

    function Get-MediaTestTask {
        param($Value)
        $task = [pscustomobject]@{ Value = $Value }
        $task | Add-Member ScriptMethod GetAwaiter { return $this }
        $task | Add-Member ScriptMethod GetResult { return $this.Value }
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
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } | Should -Throw '*Content-Length*byte count*'
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
        $errorRecord.Exception.Message | Should -Be 'Media request or stream failed before completion.'
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
        $source | Add-Member ScriptMethod Read { throw 'synthetic-secret in response stream exception' }
        $source | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        $script:Client.Response.Content.Source = $source
        { Invoke-PodcastMediaRequest -Uri 'https://media.invalid/episode.mp3' -DestinationStream $script:Destination } |
            Should -Throw 'Media request or stream failed before completion.'
        $source.Disposed | Should -BeTrue
        $script:Client.Response.Disposed | Should -BeTrue
        $script:Client.Disposed | Should -BeTrue
        $script:Destination.CanWrite | Should -BeTrue
    }
}
