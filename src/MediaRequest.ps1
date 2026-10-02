#requires -Version 5.1

function Invoke-PodcastMediaRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][IO.Stream]$DestinationStream,
        $Policy,
        $Resume,
        [scriptblock]$OnResponse,
        [scriptblock]$OnProgress
    )

    $target = Get-PodcastRequestUri -Uri $Uri
    if ($null -eq $Policy) { $Policy = New-PodcastTransportPolicy }
    $client = $null
    $request = $null
    $response = $null
    $source = $null
    $callbackFailure = $null
    $failure = 'Media request or stream failed before completion.'
    try {
        if (-not $DestinationStream.CanWrite) {
            $failure = 'The owned media destination stream is not writable.'
            throw $failure
        }
        if ($null -ne $Resume -and (-not (Test-PodcastResumeRequest -Resume $Resume) -or
            -not $DestinationStream.CanSeek -or $DestinationStream.Length -ne $Resume.Offset -or $DestinationStream.Position -ne $Resume.Offset)) {
            $failure = 'The owned partial stream does not match the resume evidence.'
            throw $failure
        }
        $client = Get-PodcastHttpClient
        $exchange = Invoke-PodcastHttpGet -Uri $target.AbsoluteUri -Client $client -Policy $Policy -Resume $Resume
        $request = $exchange.Request
        $response = $exchange.Response
        $status = [int]$response.StatusCode
        $decision = Get-PodcastResumeResponseDecision -Response $response -FinalUri $exchange.FinalUri -Resume $Resume
        if ($decision.RestartRequired) {
            return [pscustomobject]@{ Completed = $false; RestartRequired = $true; RestartReason = $decision.RestartReason; StatusCode = $status; Bytes = [long]$Resume.Offset }
        }
        if ($null -eq $Resume -and ($status -ne 200 -or $response.Content.Headers.Contains('Content-Range'))) {
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
        if ($OnResponse) {
            try { $null = & $OnResponse $decision }
            catch { $callbackFailure = $_; throw }
        }
        $source = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $buffer = New-Object byte[] 65536
        [long]$bytes = 0
        while (($read = Read-PodcastResponseChunk -Source $source -Buffer $buffer -Count $buffer.Length -Policy $Policy) -gt 0) {
            if ($null -ne $Resume -and $bytes + $read -gt $decision.ResponseLength) {
                $failure = 'The resumed body exceeds its validated response range.'
                throw $failure
            }
            $DestinationStream.Write($buffer, 0, $read)
            $bytes += $read
            if ($OnProgress) {
                try { $null = & $OnProgress ([long]$decision.Offset + $bytes) }
                catch { $callbackFailure = $_; throw }
            }
        }
        # .NET Framework can return EOF without an exception when a response
        # ends before its advertised length. Count bytes explicitly in both engines.
        if ($null -ne $decision.ResponseLength -and $bytes -ne [long]$decision.ResponseLength) {
            $failure = 'Media response Content-Length does not match the received byte count.'
            throw (New-PodcastTransportException -Kind IncompleteBody -Message $failure -Retryable ($bytes -lt $decision.ResponseLength))
        }
        return [pscustomobject]@{
            Completed = $true
            StatusCode = $status
            Bytes = [long]$decision.Offset + $bytes
            ContentLength = $decision.TotalLength
            ContentType = [string]$response.Content.Headers.ContentType
            ContentEncoding = [string]::Join(',', [string[]]$contentEncoding)
        }
    }
    catch {
        if ($null -ne $callbackFailure) { throw $callbackFailure }
        $transportFailure = Get-PodcastTransportFailure -ErrorObject $_
        if ($null -ne $transportFailure) { throw $transportFailure }
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
