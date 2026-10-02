#requires -Version 5.1

function Get-PodcastRequestUri {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Uri, [Uri]$PreviousUri)

    $parsed = $null
    if ($Uri.Length -gt 32768 -or $Uri -match '[\\\x00-\x20\x7f]' -or
        $Uri -notmatch '^(?i:https?)://' -or
        -not [Uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed) -or
        -not $parsed.IsWellFormedOriginalString() -or [string]::IsNullOrEmpty($parsed.Host)) {
        throw 'The request target must be a well-formed absolute HTTP or HTTPS URL.'
    }
    $authority = [regex]::Match($Uri, '^[^:]+://([^/?#]*)').Groups[1].Value
    if ($parsed.UserInfo.Length -gt 0 -or $authority.Contains('@')) {
        throw 'Request URLs containing user information are unsupported.'
    }
    if ($null -ne $PreviousUri -and $PreviousUri.Scheme -eq 'https' -and $parsed.Scheme -ne 'https') {
        throw 'An HTTPS request cannot redirect to unencrypted HTTP.'
    }
    return $parsed
}

function Get-PodcastHttpHandler {
    [CmdletBinding()]
    param()

    # System.Net.Http is part of both supported engines; load only on request.
    Add-Type -AssemblyName System.Net.Http -ErrorAction Stop
    $handler = [Net.Http.HttpClientHandler]::new()
    try {
        $handler.AllowAutoRedirect = $false
        $handler.UseCookies = $false
        $handler.UseDefaultCredentials = $false
        $handler.Credentials = $null
        $handler.PreAuthenticate = $false
        $handler.AutomaticDecompression = [Net.DecompressionMethods]::None
        # Proxy selection and certificate/TLS validation retain platform defaults.
        # We neither copy credentials nor install a certificate callback.
        return $handler
    }
    catch {
        $handler.Dispose()
        throw 'Could not initialize the HTTP request policy.'
    }
}

function Get-PodcastHttpClient {
    [CmdletBinding()]
    param()

    $handler = Get-PodcastHttpHandler
    try { return [Net.Http.HttpClient]::new($handler, $true) }
    catch {
        $handler.Dispose()
        throw 'Could not initialize the HTTP request client.'
    }
}

function Invoke-PodcastHttpGet {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)]$Client)

    $current = Get-PodcastRequestUri -Uri $Uri
    $request = $null
    $response = $null
    $redirects = 0
    try {
        while ($true) {
            # Each hop gets a fresh, fixed header set. No response cookies,
            # Authorization, Referer or caller-supplied headers are propagated.
            $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Get, $current)
            $request.Headers.AcceptEncoding.ParseAdd('identity')
            try { $response = $Client.SendAsync($request, [Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult() }
            catch { throw 'HTTP request failed before response headers were available.' }
            if (@(301, 302, 303, 307, 308) -notcontains [int]$response.StatusCode) {
                $result = [pscustomobject]@{ Response = $response; Request = $request; FinalUri = $current; RedirectCount = $redirects }
                $response = $null
                $request = $null
                return $result
            }
            if ($redirects -ge 5) { throw 'The HTTP redirect limit of five hops was exceeded.' }
            $location = $null
            try { $location = $response.Headers.Location }
            catch { throw 'The HTTP redirect has an invalid Location header.' }
            if ($null -eq $location) { throw 'The HTTP redirect is missing a Location header.' }
            $locationText = $location.OriginalString
            if ($locationText -match '[\\\x00-\x20\x7f]') { throw 'The HTTP redirect has an invalid Location header.' }
            $next = $null
            if (-not [Uri]::TryCreate($current, $location, [ref]$next)) { throw 'The HTTP redirect has an invalid Location header.' }
            $current = Get-PodcastRequestUri -Uri $next.AbsoluteUri -PreviousUri $current
            $redirects++
            $response.Dispose()
            $response = $null
            $request.Dispose()
            $request = $null
        }
    }
    finally {
        foreach ($resource in @($response, $request)) {
            if ($null -ne $resource) {
                try { $resource.Dispose() }
                catch { Write-Verbose 'An HTTP request resource could not be closed.' }
            }
        }
    }
}

function ConvertFrom-PodcastMetadataBody {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes, [AllowNull()][string]$Charset)

    $encodingName = 'utf-8'
    $offset = 0
    # A BOM is authoritative. Without one use HTTP charset, then the bounded
    # XML declaration. Reject unsupported/invalid encodings instead of guessing.
    if ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0xff -and $Bytes[1] -eq 0xfe -and $Bytes[2] -eq 0 -and $Bytes[3] -eq 0) { $encodingName = 'utf-32'; $offset = 4 }
    elseif ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0 -and $Bytes[1] -eq 0 -and $Bytes[2] -eq 0xfe -and $Bytes[3] -eq 0xff) { $encodingName = 'utf-32BE'; $offset = 4 }
    elseif ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xef -and $Bytes[1] -eq 0xbb -and $Bytes[2] -eq 0xbf) { $encodingName = 'utf-8'; $offset = 3 }
    elseif ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xff -and $Bytes[1] -eq 0xfe) { $encodingName = 'utf-16'; $offset = 2 }
    elseif ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xfe -and $Bytes[1] -eq 0xff) { $encodingName = 'utf-16BE'; $offset = 2 }
    elseif (-not [string]::IsNullOrWhiteSpace($Charset)) { $encodingName = $Charset.Trim('"') }
    elseif ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0x3c -and $Bytes[1] -eq 0 -and $Bytes[2] -eq 0x3f -and $Bytes[3] -eq 0) { $encodingName = 'utf-16' }
    elseif ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0 -and $Bytes[1] -eq 0x3c -and $Bytes[2] -eq 0 -and $Bytes[3] -eq 0x3f) { $encodingName = 'utf-16BE' }
    else {
        $prefix = [Text.Encoding]::ASCII.GetString($Bytes, 0, [Math]::Min(512, $Bytes.Length))
        $declaration = [regex]::Match($prefix, '^<\?xml\s[^?]*\bencoding\s*=\s*["''](?<encoding>[A-Za-z0-9._-]{1,40})["'']', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($declaration.Success) { $encodingName = $declaration.Groups['encoding'].Value }
    }
    if ($encodingName -notmatch '^[A-Za-z0-9._-]{1,40}$') { throw 'Metadata declares an unsupported character encoding.' }
    try {
        $encoding = [Text.Encoding]::GetEncoding($encodingName, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
        if (@(65001, 1200, 1201, 12000, 12001, 20127, 28591, 1252) -notcontains $encoding.CodePage) {
            throw 'Unsupported metadata encoding.'
        }
        return $encoding.GetString($Bytes, $offset, $Bytes.Length - $offset)
    }
    catch { throw 'Metadata bytes are invalid or use an unsupported character encoding.' }
}

function Invoke-PodcastMetadataRequest {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri, [ValidateRange(1, 8388608)][long]$MaximumBytes = 8388608)

    # Reject unsafe input before creating a transport client or opening a body.
    $target = Get-PodcastRequestUri -Uri $Uri
    $client = $null
    $request = $null
    $response = $null
    $source = $null
    $memory = $null
    $failure = 'Metadata request or stream failed before completion.'
    try {
        $client = Get-PodcastHttpClient
        $exchange = Invoke-PodcastHttpGet -Uri $target.AbsoluteUri -Client $client
        $request = $exchange.Request
        $response = $exchange.Response
        if ([int]$response.StatusCode -ne 200 -or $response.Content.Headers.Contains('Content-Range')) {
            $failure = 'Metadata response must be a complete HTTP 200 body.'
            throw $failure
        }
        $contentLength = $response.Content.Headers.ContentLength
        if ($response.Content.Headers.Contains('Content-Length') -and $null -eq $contentLength) {
            $failure = 'Metadata response has an invalid Content-Length header.'
            throw $failure
        }
        if ($null -ne $contentLength -and $contentLength -gt $MaximumBytes) {
            $failure = 'Metadata response exceeds the configured byte limit (at most 8 MiB).'
            throw $failure
        }
        foreach ($encoding in @($response.Content.Headers.ContentEncoding)) {
            if (-not [string]::Equals($encoding, 'identity', [StringComparison]::OrdinalIgnoreCase)) {
                $failure = 'Metadata response uses an unsupported Content-Encoding.'
                throw $failure
            }
        }
        $source = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $memory = [IO.MemoryStream]::new()
        $buffer = New-Object byte[] 16384
        while ($true) {
            $capacity = [int][Math]::Min($buffer.Length, ($MaximumBytes - $memory.Length + 1))
            $read = $source.Read($buffer, 0, $capacity)
            if ($read -eq 0) { break }
            if ($memory.Length + $read -gt $MaximumBytes) {
                $failure = 'Metadata response exceeds the configured byte limit (at most 8 MiB).'
                throw $failure
            }
            $memory.Write($buffer, 0, $read)
        }
        if ($null -ne $contentLength -and $memory.Length -ne $contentLength) {
            $failure = 'Metadata response Content-Length does not match the received byte count.'
            throw $failure
        }
        $charset = $null
        if ($null -ne $response.Content.Headers.ContentType) { $charset = $response.Content.Headers.ContentType.CharSet }
        $failure = 'Metadata bytes are invalid or use an unsupported character encoding.'
        $text = ConvertFrom-PodcastMetadataBody -Bytes $memory.ToArray() -Charset $charset
        return [pscustomobject]@{
            Content = $text
            FinalUri = $exchange.FinalUri
            StatusCode = [int]$response.StatusCode
            Bytes = $memory.Length
            ContentType = [string]$response.Content.Headers.ContentType
        }
    }
    catch { throw $failure }
    finally {
        foreach ($resource in @($source, $memory, $response, $request, $client)) {
            if ($null -ne $resource) {
                try { $resource.Dispose() }
                catch { Write-Verbose 'A metadata request resource could not be closed.' }
            }
        }
    }
}
