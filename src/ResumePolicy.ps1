#requires -Version 5.1

function Test-PodcastStrongETag {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$ETag)

    # Strong opaque tags only. Bound persisted/header values without logging them.
    return $null -ne $ETag -and $ETag.Length -le 1024 -and $ETag -cmatch '^"[\x21\x23-\x7e]*"$'
}

function Get-PodcastResumeContentType {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$ContentType)

    if (-not ('Net.Http.Headers.MediaTypeHeaderValue' -as [type])) { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop }
    $parsed = $null
    if ($ContentType.Length -gt 1024 -or -not [Net.Http.Headers.MediaTypeHeaderValue]::TryParse($ContentType, [ref]$parsed)) { return '' }
    return $parsed.MediaType.ToLowerInvariant()
}

function Get-PodcastResumeUriFingerprint {
    [CmdletBinding()]
    param([Parameter(Mandatory)][Uri]$Uri)

    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Uri.AbsoluteUri)))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Get-PodcastResumeResponseETag {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Response)

    if ($null -eq $Response.Headers) { return '' }
    if ($null -ne $Response.Headers.PSObject.Methods['TryGetValues']) {
        $values = $null
        if (-not $Response.Headers.TryGetValues('ETag', [ref]$values) -or @($values).Count -ne 1) { return '' }
        $tag = ([string]@($values)[0]).Trim()
        if (Test-PodcastStrongETag -ETag $tag) { return $tag }
        return ''
    }
    return [string]$Response.Headers.ETag
}

function Test-PodcastResumeRequest {
    [CmdletBinding()]
    param([AllowNull()]$Resume)

    if ($null -eq $Resume) { return $false }
    if ($Resume.Offset -isnot [long] -and $Resume.Offset -isnot [int]) { return $false }
    if ($Resume.TotalLength -isnot [long] -and $Resume.TotalLength -isnot [int]) { return $false }
    return $Resume.Offset -gt 0 -and $Resume.TotalLength -gt $Resume.Offset -and
        (Test-PodcastStrongETag -ETag $Resume.ETag) -and
        $Resume.FinalUriFingerprint -cmatch '^[a-f0-9]{64}$' -and
        -not [string]::IsNullOrWhiteSpace($Resume.ContentType) -and
        [string]::Equals((Get-PodcastResumeContentType -ContentType $Resume.ContentType), $Resume.ContentType, [StringComparison]::Ordinal) -and
        $Resume.ContentType -notmatch '^multipart/'
}

function Get-PodcastResumeResponseDecision {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Response, [Parameter(Mandatory)][Uri]$FinalUri, [AllowNull()]$Resume)

    $status = [int]$Response.StatusCode
    $headers = $Response.Content.Headers
    $contentLength = $headers.ContentLength
    $contentType = Get-PodcastResumeContentType -ContentType ([string]$headers.ContentType)
    $etag = Get-PodcastResumeResponseETag -Response $Response
    $fingerprint = Get-PodcastResumeUriFingerprint -Uri $FinalUri
    $identity = $true
    foreach ($encoding in @($headers.ContentEncoding)) {
        if (-not [string]::Equals($encoding, 'identity', [StringComparison]::OrdinalIgnoreCase)) { $identity = $false }
    }
    $decision = [pscustomobject]@{
        Action = 'Fresh'; RestartRequired = $false; RestartReason = ''; Offset = [long]0
        TotalLength = $contentLength; ResponseLength = $contentLength; ETag = $etag
        ContentType = $contentType; ContentEncoding = 'identity'; FinalUriFingerprint = $fingerprint
        ResumeSupported = $false; StatusCode = $status
    }
    if ($null -eq $Resume) {
        $decision.ResumeSupported = $status -eq 200 -and -not $headers.Contains('Content-Range') -and
            $null -ne $contentLength -and $contentLength -gt 0 -and $identity -and
            (Test-PodcastStrongETag -ETag $etag) -and
            -not [string]::IsNullOrWhiteSpace($contentType) -and $contentType.Length -le 1024 -and
            $contentType -notmatch '[\x00-\x1f\x7f]' -and $contentType -notmatch '^(?i:multipart/)'
        return $decision
    }

    $reason = ''
    if (-not (Test-PodcastResumeRequest -Resume $Resume)) { $reason = 'invalid-resume-evidence' }
    elseif ($status -eq 200) { $reason = 'range-ignored-or-entity-changed' }
    elseif ($status -eq 416) { $reason = 'range-not-satisfiable' }
    elseif ($status -ne 206) { $reason = 'unexpected-resume-status' }
    elseif (-not [string]::Equals($fingerprint, $Resume.FinalUriFingerprint, [StringComparison]::Ordinal)) { $reason = 'response-target-changed' }
    elseif (-not (Test-PodcastStrongETag -ETag $etag) -or -not [string]::Equals($etag, $Resume.ETag, [StringComparison]::Ordinal)) { $reason = 'entity-validator-changed' }
    elseif (-not $identity -or -not [string]::Equals($contentType, $Resume.ContentType, [StringComparison]::Ordinal) -or $contentType -match '^(?i:multipart/)') { $reason = 'response-representation-changed' }
    else {
        $range = $null
        $singleRange = $true
        if ($null -ne $headers.PSObject.Methods['TryGetValues']) {
            $rangeValues = $null
            $singleRange = $headers.TryGetValues('Content-Range', [ref]$rangeValues) -and @($rangeValues).Count -eq 1
        }
        try { $range = $headers.ContentRange }
        catch { Write-Verbose 'The resumed response had an invalid Content-Range.' }
        # Only one complete remaining tail is supported. A shorter valid range
        # is restarted conservatively rather than combining multiple segments.
        if (-not $singleRange -or $null -eq $range -or -not $range.HasRange -or -not $range.HasLength -or
            -not [string]::Equals($range.Unit, 'bytes', [StringComparison]::OrdinalIgnoreCase) -or
            $range.From -ne $Resume.Offset -or $range.Length -ne $Resume.TotalLength -or
            $range.To -ne ($Resume.TotalLength - 1)) { $reason = 'invalid-response-range' }
        elseif (($headers.Contains('Content-Length') -and $null -eq $contentLength) -or
            ($null -ne $contentLength -and $contentLength -ne ($Resume.TotalLength - $Resume.Offset))) { $reason = 'invalid-range-content-length' }
    }
    if ($reason) {
        $decision.Action = 'Restart'
        $decision.RestartRequired = $true
        $decision.RestartReason = $reason
        return $decision
    }
    $decision.Action = 'Append'
    $decision.Offset = [long]$Resume.Offset
    $decision.TotalLength = [long]$Resume.TotalLength
    $decision.ResponseLength = [long]$Resume.TotalLength - [long]$Resume.Offset
    $decision.ResumeSupported = $true
    return $decision
}
