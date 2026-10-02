BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/NetworkPolicy.ps1')
    Add-Type -AssemblyName System.Net.Http
    Mock Invoke-WebRequest { throw 'Network policy units must not use external network.' }

    function Get-PolicyTestTask {
        param($Value)
        $task = [pscustomobject]@{ Value = $Value }
        $task | Add-Member ScriptMethod GetAwaiter { return $this }
        $task | Add-Member ScriptMethod GetResult { return $this.Value }
        return $task
    }

    function Get-PolicyTestResponse {
        param([int]$Status = 200, [string]$Location, [byte[]]$Bytes = @(60, 114, 47, 62), $Length,
            [string[]]$Encoding = @(), [string]$Charset = 'utf-8')
        if (-not $PSBoundParameters.ContainsKey('Length')) { $Length = $Bytes.Length }
        $contentType = [Net.Http.Headers.MediaTypeHeaderValue]::new('application/xml')
        if ($Charset) { $contentType.CharSet = $Charset }
        $headers = [pscustomobject]@{ ContentLength = $Length; ContentEncoding = $Encoding; ContentType = $contentType; HasLength = $null -ne $Length; HasRange = $false }
        $headers | Add-Member ScriptMethod Contains {
            param($name)
            if ($name -eq 'Content-Length') { return $this.HasLength }
            if ($name -eq 'Content-Range') { return $this.HasRange }
            return $false
        }
        $content = [pscustomobject]@{ Headers = $headers; Source = [IO.MemoryStream]::new($Bytes); ReadCalled = $false }
        $content | Add-Member ScriptMethod ReadAsStreamAsync { $this.ReadCalled = $true; return (Get-PolicyTestTask $this.Source) }
        $response = [pscustomobject]@{ StatusCode = $Status; Headers = [pscustomobject]@{ Location = $null }; Content = $content; Disposed = $false }
        if ($Location) { $response.Headers.Location = [Uri]::new($Location, [UriKind]::RelativeOrAbsolute) }
        $response | Add-Member ScriptMethod Dispose { $this.Disposed = $true; $this.Content.Source.Dispose() }
        return $response
    }

    function Get-PolicyTestClient {
        param([object[]]$Responses)
        $client = [pscustomobject]@{ Responses = $Responses; Requests = (New-Object 'System.Collections.Generic.List[object]'); Completion = $null; Disposed = $false }
        $client | Add-Member ScriptMethod SendAsync {
            param($request, $completion)
            $this.Requests.Add($request)
            $this.Completion = $completion
            return (Get-PolicyTestTask $this.Responses[[Math]::Min($this.Requests.Count - 1, $this.Responses.Count - 1)])
        }
        $client | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        return $client
    }
}

Describe 'A023: HTTP target and handler policy' -Tag 'Unit', 'A023' {
    It 'accepts explicit HTTP(S) including local and private targets <Url>' -ForEach @(
        @{ Url = 'https://podcast.invalid/feed?opaque=token' },
        @{ Url = 'http://podcast.invalid/feed' },
        @{ Url = 'http://127.0.0.1:12345/feed' },
        @{ Url = 'http://[::1]:12345/feed' },
        @{ Url = 'https://10.0.0.1/feed' },
        @{ Url = 'https://192.168.1.1/feed' },
        @{ Url = 'https://podcast.invalid/escaped%20path' }
    ) {
        $target = Get-PodcastRequestUri -Uri $Url
        $target.IsAbsoluteUri | Should -BeTrue
        $target.OriginalString | Should -BeExactly $Url
    }

    It 'rejects an unsafe target <Case> without including it in the error' -ForEach @(
        @{ Case = 'file scheme'; Url = 'file:///private-path' },
        @{ Case = 'FTP scheme'; Url = 'ftp://podcast.invalid/private-path' },
        @{ Case = 'data scheme'; Url = 'data:text/plain,private-token' },
        @{ Case = 'relative URL'; Url = '/private-path' },
        @{ Case = 'protocol relative URL'; Url = '//podcast.invalid/private-path' },
        @{ Case = 'userinfo'; Url = 'https://private-user:private-token@podcast.invalid/' },
        @{ Case = 'username only'; Url = 'https://private-user@podcast.invalid/' },
        @{ Case = 'empty userinfo'; Url = 'https://@podcast.invalid/private-path' },
        @{ Case = 'backslash'; Url = 'https://podcast.invalid\private-path' },
        @{ Case = 'whitespace'; Url = 'https://podcast.invalid/private path' },
        @{ Case = 'control'; Url = "https://podcast.invalid/private`npath" },
        @{ Case = 'empty'; Url = '' },
        @{ Case = 'malformed host'; Url = 'https://[/private-path' }
    ) {
        $caught = $null
        try { $null = Get-PodcastRequestUri -Uri $Url }
        catch { $caught = $_ }
        $caught | Should -Not -BeNullOrEmpty
        $caught.Exception.Message | Should -Not -Match 'private-|podcast.invalid'
    }

    It 'allows an HTTP upgrade but rejects an HTTPS downgrade' {
        (Get-PodcastRequestUri -Uri 'https://podcast.invalid/feed' -PreviousUri 'http://podcast.invalid/feed').Scheme | Should -Be 'https'
        { Get-PodcastRequestUri -Uri 'http://podcast.invalid/feed' -PreviousUri 'https://podcast.invalid/feed' } | Should -Throw '*cannot redirect*'
    }

    It 'uses explicit redirect credential cookie and byte-representation settings while preserving TLS and proxy defaults' {
        $protocol = [Net.ServicePointManager]::SecurityProtocol
        $callback = [Net.ServicePointManager]::ServerCertificateValidationCallback
        $handler = Get-PodcastHttpHandler
        $defaults = [Net.Http.HttpClientHandler]::new()
        try {
            $handler.AllowAutoRedirect | Should -BeFalse
            $handler.UseCookies | Should -BeFalse
            $handler.UseDefaultCredentials | Should -BeFalse
            $handler.Credentials | Should -BeNullOrEmpty
            $handler.PreAuthenticate | Should -BeFalse
            $handler.AutomaticDecompression | Should -Be ([Net.DecompressionMethods]::None)
            $handler.UseProxy | Should -Be $defaults.UseProxy
            $handler.Proxy | Should -Be $defaults.Proxy
            if ($null -ne $handler.PSObject.Properties['ServerCertificateCustomValidationCallback']) {
                $handler.ServerCertificateCustomValidationCallback | Should -BeNullOrEmpty
            }
            [Net.ServicePointManager]::SecurityProtocol | Should -Be $protocol
            [object]::ReferenceEquals([Net.ServicePointManager]::ServerCertificateValidationCallback, $callback) | Should -BeTrue
        }
        finally { $handler.Dispose(); $defaults.Dispose() }
    }
}

Describe 'A023: explicit validated redirect requests' -Tag 'Unit', 'A023' {
    BeforeEach {
        $script:Responses = @()
        $script:Exchange = $null
    }
    AfterEach {
        if ($null -ne $script:Exchange) { $script:Exchange.Request.Dispose() }
        foreach ($response in $script:Responses) { $response.Dispose() }
    }

    It 'follows relative redirects with fresh safe GET requests and returns the final URI' -ForEach @(
        @{ Status = 301 }, @{ Status = 302 }, @{ Status = 303 }, @{ Status = 307 }, @{ Status = 308 }
    ) {
        $script:Responses = @((Get-PolicyTestResponse -Status $Status -Location '../feed.xml'), (Get-PolicyTestResponse))
        $client = Get-PolicyTestClient -Responses $script:Responses
        $script:Exchange = Invoke-PodcastHttpGet -Uri 'https://podcast.invalid/start/page' -Client $client
        $script:Exchange.FinalUri.AbsoluteUri | Should -Be 'https://podcast.invalid/feed.xml'
        $script:Exchange.RedirectCount | Should -Be 1
        $client.Requests.Count | Should -Be 2
        $script:Responses[0].Disposed | Should -BeTrue
        $script:Responses[0].Content.ReadCalled | Should -BeFalse
        $script:Responses[1].Disposed | Should -BeFalse
        foreach ($request in $client.Requests) {
            $request.Method.Method | Should -Be 'GET'
            @($request.Headers).Count | Should -Be 1
            @($request.Headers.AcceptEncoding)[0].Value | Should -Be 'identity'
        }
        $client.Completion | Should -Be ([Net.Http.HttpCompletionOption]::ResponseHeadersRead)
    }

    It 'does not forward cookies credentials referer or origin when following another origin' {
        $script:Responses = @((Get-PolicyTestResponse -Status 302 -Location 'https://other.invalid/feed'), (Get-PolicyTestResponse))
        $script:Responses[0].Headers | Add-Member NoteProperty Cookie 'private-cookie'
        $script:Responses[0].Headers | Add-Member NoteProperty Authorization 'private-authorization'
        $client = Get-PolicyTestClient -Responses $script:Responses
        $script:Exchange = Invoke-PodcastHttpGet -Uri 'https://podcast.invalid/private-path?private-token' -Client $client
        $redirected = $client.Requests[1]
        $redirected.RequestUri.AbsoluteUri | Should -Be 'https://other.invalid/feed'
        foreach ($name in @('Authorization', 'Cookie', 'Referer', 'Origin')) { $redirected.Headers.Contains($name) | Should -BeFalse }
    }

    It 'rejects an unsafe redirect before contacting its target <Case>' -ForEach @(
        @{ Case = 'file'; Location = 'file:///private-path' },
        @{ Case = 'userinfo'; Location = 'https://private-user:private-token@other.invalid/feed' },
        @{ Case = 'downgrade'; Location = 'http://other.invalid/feed' },
        @{ Case = 'missing'; Location = '' }
    ) {
        $script:Responses = @((Get-PolicyTestResponse -Status 302 -Location $Location))
        $client = Get-PolicyTestClient -Responses $script:Responses
        { Invoke-PodcastHttpGet -Uri 'https://podcast.invalid/feed' -Client $client } | Should -Throw
        $client.Requests.Count | Should -Be 1
        $script:Responses[0].Disposed | Should -BeTrue
        $script:Responses[0].Content.ReadCalled | Should -BeFalse
    }

    It 'allows five redirects and rejects a sixth before making a seventh request' {
        $script:Responses = @(foreach ($number in 1..6) { Get-PolicyTestResponse -Status 302 -Location ('/hop-' + $number) })
        $client = Get-PolicyTestClient -Responses $script:Responses
        { Invoke-PodcastHttpGet -Uri 'https://podcast.invalid/feed' -Client $client } | Should -Throw '*five hops*'
        $client.Requests.Count | Should -Be 6
        @($script:Responses | Where-Object { -not $_.Disposed }).Count | Should -Be 0
    }

    It 'does not interpret other 3xx statuses as redirects' {
        $script:Responses = @((Get-PolicyTestResponse -Status 304 -Location 'https://other.invalid/feed'))
        $client = Get-PolicyTestClient -Responses $script:Responses
        $script:Exchange = Invoke-PodcastHttpGet -Uri 'https://podcast.invalid/feed' -Client $client
        $client.Requests.Count | Should -Be 1
        $script:Exchange.Response.StatusCode | Should -Be 304
    }
}

Describe 'A023/A024: bounded metadata transfer and decoding' -Tag 'Unit', 'A023', 'A024' {
    BeforeEach {
        $script:Response = Get-PolicyTestResponse
        $script:Client = Get-PolicyTestClient -Responses @($script:Response)
        Mock Get-PodcastHttpClient { return $script:Client }
    }
    AfterEach { $script:Response.Dispose() }

    It 'validates an initial URI before opening a client' {
        { Invoke-PodcastMetadataRequest -Uri 'file:///private-path' } | Should -Throw '*HTTP or HTTPS*'
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
    }

    It 'returns decoded content final URI and a byte count without writing files' {
        $before = @(Get-ChildItem -LiteralPath $TestDrive -Recurse -Force).Count
        $result = Invoke-PodcastMetadataRequest -Uri 'https://podcast.invalid/feed'
        $result.Content | Should -Be '<r/>'
        $result.Bytes | Should -Be 4
        $result.FinalUri.AbsoluteUri | Should -Be 'https://podcast.invalid/feed'
        $script:Client.Disposed | Should -BeTrue
        $script:Response.Disposed | Should -BeTrue
        @(Get-ChildItem -LiteralPath $TestDrive -Recurse -Force).Count | Should -Be $before
    }

    It 'rejects declared oversized metadata before reading its body' {
        $script:Response.Content.Headers.ContentLength = 8388609
        { Invoke-PodcastMetadataRequest -Uri 'https://podcast.invalid/feed' } | Should -Throw '*byte limit*'
        $script:Response.Content.ReadCalled | Should -BeFalse
        $script:Response.Disposed | Should -BeTrue
    }

    It 'bounds unknown-length bodies while streaming and releases all resources' {
        $script:Response.Content.Headers.ContentLength = $null
        $script:Response.Content.Headers.HasLength = $false
        { Invoke-PodcastMetadataRequest -Uri 'https://podcast.invalid/feed' -MaximumBytes 3 } | Should -Throw '*byte limit*'
        $script:Response.Content.ReadCalled | Should -BeTrue
        $script:Client.Disposed | Should -BeTrue
        $script:Response.Content.Source.CanRead | Should -BeFalse
    }

    It 'rejects a misleading declared length after stream completion' {
        $script:Response.Content.Headers.ContentLength = 8
        { Invoke-PodcastMetadataRequest -Uri 'https://podcast.invalid/feed' } | Should -Throw '*Content-Length*byte count*'
    }

    It 'rejects content encodings before reading compressed bodies' {
        $script:Response.Content.Headers.ContentEncoding = @('gzip')
        { Invoke-PodcastMetadataRequest -Uri 'https://podcast.invalid/feed' } | Should -Throw '*Content-Encoding*'
        $script:Response.Content.ReadCalled | Should -BeFalse
    }

    It 'rejects partial metadata before reading its body' {
        $script:Response.Content.Headers.HasRange = $true
        { Invoke-PodcastMetadataRequest -Uri 'https://podcast.invalid/feed' } | Should -Throw '*complete HTTP 200*'
        $script:Response.Content.ReadCalled | Should -BeFalse
    }

    It 'does not permit raising the fixed maximum budget' {
        { Invoke-PodcastMetadataRequest -Uri 'https://podcast.invalid/feed' -MaximumBytes 8388609 } | Should -Throw
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
    }

    It 'honors a BOM and preserves valid non-ASCII content for <Encoding>' -ForEach @(
        @{ Encoding = 'utf-8' }, @{ Encoding = 'utf-16' }, @{ Encoding = 'utf-16BE' }, @{ Encoding = 'utf-32' }, @{ Encoding = 'utf-32BE' }
    ) {
        $text = '<rss>' + [char]0x00e4 + [char]0x65e5 + [char]::ConvertFromUtf32(0x1f3a7) + '</rss>'
        $codec = [Text.Encoding]::GetEncoding($Encoding)
        [byte[]]$bytes = @($codec.GetPreamble()) + @($codec.GetBytes($text))
        (ConvertFrom-PodcastMetadataBody -Bytes $bytes -Charset 'us-ascii') | Should -BeExactly $text
    }

    It 'uses a bounded XML declaration when no HTTP charset or BOM is present' {
        $text = '<?xml version="1.0" encoding="windows-1252"?><rss>' + [char]0x00e4 + '</rss>'
        $bytes = [Text.Encoding]::GetEncoding(1252).GetBytes($text)
        (ConvertFrom-PodcastMetadataBody -Bytes $bytes) | Should -BeExactly $text
    }

    It 'honors a supported HTTP charset without a BOM' {
        (ConvertFrom-PodcastMetadataBody -Bytes @([byte]0xe4) -Charset 'iso-8859-1') | Should -Be ([string][char]0x00e4)
    }

    It 'rejects invalid UTF-8 rather than replacing bytes silently' {
        { ConvertFrom-PodcastMetadataBody -Bytes @(0xc3, 0x28) -Charset 'utf-8' } | Should -Throw '*invalid*encoding*'
    }

    It 'rejects unsupported or injected charset names without printing them' -ForEach @(
        @{ Charset = 'utf-7' }, @{ Charset = 'private-unknown-encoding' }, @{ Charset = 'private-token:header' }
    ) {
        $caught = $null
        try { $null = ConvertFrom-PodcastMetadataBody -Bytes @(60, 114, 47, 62) -Charset $Charset }
        catch { $caught = $_ }
        $caught.Exception.Message | Should -Match 'unsupported character encoding'
        $caught.Exception.Message | Should -Not -Match 'private-|utf-7'
    }
}
