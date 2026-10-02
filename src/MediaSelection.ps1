# Pure enclosure hints. Transfer signatures, rather than these hints, select final extensions.
function Get-PodcastAudioHint {
    [CmdletBinding()]
    param([string]$ContentType, [string]$Url)

    $mime = ($ContentType.Split(';')[0]).Trim().ToLowerInvariant()
    $extension = switch ($mime) {
        { $_ -in @('audio/mpeg', 'audio/mp3', 'audio/x-mp3', 'audio/x-mpeg', 'audio/x-mpegmp3') } { 'mp3'; break }
        { $_ -in @('audio/mp4', 'audio/m4a', 'audio/x-m4a') } { 'm4a'; break }
        { $_ -in @('audio/ogg', 'audio/opus', 'audio/vorbis', 'audio/speex') } { 'ogg'; break }
        { $_ -in @('audio/wav', 'audio/wave', 'audio/x-wav', 'audio/vnd.wave') } { 'wav'; break }
        { $_ -in @('audio/flac', 'audio/x-flac') } { 'flac'; break }
        default { $null }
    }
    if ($extension) {
        return [pscustomobject]@{ Extension = $extension; Rank = 2; Eligible = $true; Reason = 'supported_audio_type' }
    }
    if ($mime -notin @('', 'application/octet-stream', 'binary/octet-stream', 'application/ogg', 'application/mp4')) {
        return [pscustomobject]@{ Extension = $null; Rank = -1; Eligible = $false; Reason = 'unsupported_enclosure_type' }
    }
    $urlExtension = $null
    $parsedUrl = $null
    $urlPath = if ([Uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$parsedUrl)) { $parsedUrl.AbsolutePath } else { '' }
    if ($urlPath -match '\.(mp3|m4a|ogg|opus|wav|flac)$') {
        $urlExtension = $Matches[1].ToLowerInvariant()
        if ($urlExtension -eq 'opus') { $urlExtension = 'ogg' }
    }
    return [pscustomobject]@{
        Extension = $urlExtension
        Rank = $(if ($urlExtension) { 1 } else { 0 })
        Eligible = $true
        Reason = $(if ($urlExtension) { 'audio_url_hint' } else { 'requires_audio_signature' })
    }
}

function Select-PodcastAudioCandidate {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyCollection()][object[]]$Candidates)

    $selected = $null
    foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace([string]$candidate.Url)) { continue }
        $hint = Get-PodcastAudioHint -ContentType $candidate.ContentType -Url $candidate.Url
        if ($hint.Eligible) {
            $eligible = [pscustomobject]@{
                Url = [string]$candidate.Url; ContentType = [string]$candidate.ContentType
                Length = $candidate.Length; Extension = $hint.Extension; Reason = $hint.Reason
            }
            # Keep the first known audio URL stable when a publisher adds a
            # later typed alternative. An unknown generic URL is only fallback.
            if ($hint.Rank -ge 1) { return $eligible }
            if ($null -eq $selected) { $selected = $eligible }
        }
    }
    return $selected
}
