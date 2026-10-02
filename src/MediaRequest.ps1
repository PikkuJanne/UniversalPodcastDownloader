#requires -Version 5.1

function Invoke-PodcastMediaRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][IO.Stream]$DestinationStream
    )

    $target = Get-PodcastRequestUri -Uri $Uri
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
        # The default 100-second HttpClient timeout covers headers only here.
        # Reading the response does not impose a total-duration audio limit.
        $exchange = Invoke-PodcastHttpGet -Uri $target.AbsoluteUri -Client $client
        $request = $exchange.Request
        $response = $exchange.Response
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
