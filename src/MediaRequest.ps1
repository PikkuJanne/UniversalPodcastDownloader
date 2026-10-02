#requires -Version 5.1

function Get-PodcastHttpClient {
    [CmdletBinding()]
    param()

    # System.Net.Http ships with both supported engines. Load it only when a
    # transfer starts; importing the downloader must remain free of side effects.
    Add-Type -AssemblyName System.Net.Http -ErrorAction Stop
    $handler = [Net.Http.HttpClientHandler]::new()
    try {
        # Preserve the representation's byte count. Proxy and TLS validation
        # retain the platform defaults; no process-wide options are changed.
        $handler.AutomaticDecompression = [Net.DecompressionMethods]::None
        return [Net.Http.HttpClient]::new($handler, $true)
    }
    catch {
        $handler.Dispose()
        throw 'Could not initialize the media request client.'
    }
}

function Invoke-PodcastMediaRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][IO.Stream]$DestinationStream
    )

    $client = $null
    $request = $null
    $response = $null
    $source = $null
    $failure = 'Media request or stream failed before completion.'
    try {
        if (-not $DestinationStream.CanWrite) {
            $failure = 'The owned media destination stream is not writable.'
            throw $failure
        }
        $client = Get-PodcastHttpClient
        $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Get, $Uri)
        $request.Headers.AcceptEncoding.ParseAdd('identity')
        # The default 100-second HttpClient timeout covers headers only here.
        # Reading the response does not impose a total-duration audio limit.
        $response = $client.SendAsync($request, [Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        $status = [int]$response.StatusCode
        if ($status -ne 200 -or $response.Content.Headers.Contains('Content-Range')) {
            $failure = 'Media response must be a complete HTTP 200 body; partial responses are unsupported.'
            throw $failure
        }
        $contentLength = $response.Content.Headers.ContentLength
        if ($response.Content.Headers.Contains('Content-Length') -and $null -eq $contentLength) {
            $failure = 'Media response has an invalid Content-Length header.'
            throw $failure
        }
        $contentEncoding = @($response.Content.Headers.ContentEncoding)
        foreach ($encoding in $contentEncoding) {
            if (-not [string]::Equals($encoding, 'identity', [StringComparison]::OrdinalIgnoreCase)) {
                $failure = 'Media response uses an unsupported Content-Encoding.'
                throw $failure
            }
        }
        $source = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $buffer = New-Object byte[] 65536
        [long]$bytes = 0
        while (($read = $source.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $DestinationStream.Write($buffer, 0, $read)
            $bytes += $read
        }
        # .NET Framework can return EOF without an exception when a response
        # ends before its advertised length. Count bytes explicitly in both engines.
        if ($null -ne $contentLength -and $bytes -ne [long]$contentLength) {
            $failure = 'Media response Content-Length does not match the received byte count.'
            throw $failure
        }
        return [pscustomobject]@{
            Completed = $true
            StatusCode = $status
            Bytes = $bytes
            ContentLength = $contentLength
            ContentType = [string]$response.Content.Headers.ContentType
            ContentEncoding = [string]::Join(',', [string[]]$contentEncoding)
        }
    }
    catch {
        # Transport exceptions can contain a signed URL, credentials, or body
        # text. Expose only messages defined here, without the original exception.
        throw $failure
    }
    finally {
        $cleanupFailed = $false
        foreach ($resource in @($source, $response, $request, $client)) {
            if ($null -ne $resource) {
                try { $resource.Dispose() }
                catch { $cleanupFailed = $true }
            }
        }
        if ($cleanupFailed) { throw 'Media request resources could not be closed safely.' }
        # The caller keeps exclusive ownership of DestinationStream and closes it.
    }
}
