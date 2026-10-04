# Pure, bounded publication-date parsing and ordering for both Windows engines.

function ConvertTo-PodcastPublicationDate {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value.Length -gt 256) { return $null }
    $text = $Value.Trim()
    $options = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant
    $timeout = [TimeSpan]::FromMilliseconds(250)
    $iso = [regex]::Match($text, '\A(?<year>[0-9]{4})-(?<month>[0-9]{2})-(?<day>[0-9]{2})(?:[T ](?<hour>[0-9]{2}):(?<minute>[0-9]{2}):(?<second>[0-9]{2})(?:\.(?<fraction>[0-9]{1,7}))?(?<zone>Z|[+-][0-9]{2}:?[0-9]{2})?)?\z', $options, $timeout)
    $rss = $null
    if (-not $iso.Success) {
        $rss = [regex]::Match($text, '\A(?:(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun),[ \t]+)?(?<day>[0-9]{1,2})[ \t]+(?<month>[A-Za-z]{3})[ \t]+(?<year>[0-9]{4}|[0-9]{2})[ \t]+(?<hour>[0-9]{2}):(?<minute>[0-9]{2})(?::(?<second>[0-9]{2}))?[ \t]+(?<zone>[+-][0-9]{2}:?[0-9]{2}|UT|UTC|GMT|Z|EST|EDT|CST|CDT|MST|MDT|PST|PDT)\z', $options, $timeout)
        if (-not $rss.Success) { return $null }
    }

    $match = if ($iso.Success) { $iso } else { $rss }
    $year = [int]$match.Groups['year'].Value
    if ($match.Groups['year'].Value.Length -eq 2) {
        $year += $(if ($year -le 49) { 2000 } else { 1900 })
    }
    $month = 0
    if ($iso.Success) { $month = [int]$match.Groups['month'].Value }
    else {
        $months = @{ JAN=1; FEB=2; MAR=3; APR=4; MAY=5; JUN=6; JUL=7; AUG=8; SEP=9; OCT=10; NOV=11; DEC=12 }
        $month = [int]$months[$match.Groups['month'].Value.ToUpperInvariant()]
        if ($month -eq 0) { return $null }
    }
    $hour = if ($match.Groups['hour'].Success) { [int]$match.Groups['hour'].Value } else { 0 }
    $minute = if ($match.Groups['minute'].Success) { [int]$match.Groups['minute'].Value } else { 0 }
    $second = if ($match.Groups['second'].Success) { [int]$match.Groups['second'].Value } else { 0 }
    $zone = $match.Groups['zone'].Value.ToUpperInvariant()
    $offsetMinutes = 0
    if ($zone.StartsWith('+') -or $zone.StartsWith('-')) {
        $digits = $zone.Substring(1).Replace(':', '')
        $zoneHour = [int]$digits.Substring(0, 2)
        $zoneMinute = [int]$digits.Substring(2, 2)
        if ($zoneMinute -gt 59 -or $zoneHour -gt 14 -or ($zoneHour -eq 14 -and $zoneMinute -ne 0)) { return $null }
        $offsetMinutes = $zoneHour * 60 + $zoneMinute
        if ($zone.StartsWith('-')) { $offsetMinutes = -$offsetMinutes }
    }
    elseif ($zone -notin @('', 'UT', 'UTC', 'GMT', 'Z')) {
        $zones = @{ EST=-300; EDT=-240; CST=-360; CDT=-300; MST=-420; MDT=-360; PST=-480; PDT=-420 }
        $offsetMinutes = [int]$zones[$zone]
    }

    try {
        $date = [DateTimeOffset]::new($year, $month, [int]$match.Groups['day'].Value, $hour, $minute, $second, [TimeSpan]::FromMinutes($offsetMinutes))
        if ($iso.Success -and $iso.Groups['fraction'].Success) {
            $ticks = [long]::Parse($iso.Groups['fraction'].Value.PadRight(7, '0'), [Globalization.CultureInfo]::InvariantCulture)
            $date = $date.AddTicks($ticks)
        }
        return $date.ToUniversalTime()
    }
    catch [ArgumentException] { return $null }
}

function Get-PodcastPublicationUtc {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [DateTimeOffset]) { return $Value.ToUniversalTime() }
    if ($Value -is [datetime]) {
        # Existing in-memory callers with no offset use the same UTC convention.
        $date = if ($Value.Kind -eq [DateTimeKind]::Unspecified) { [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) } else { $Value }
        return ([DateTimeOffset]$date).ToUniversalTime()
    }
    return ConvertTo-PodcastPublicationDate -Value ([string]$Value)
}

function Get-PodcastLegacyPublicationDate {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    $date = ConvertTo-PodcastPublicationDate -Value $Value
    if ($null -eq $date) { return $null }
    # Reproduce the old machine-local filename only as a review hint. ISO values
    # without a zone formerly kept their wall date; no history binding uses this.
    if ($Value.Trim() -match '\A[0-9]{4}-[0-9]{2}-[0-9]{2}(?:[Tt ][0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]{1,7})?)?\z') { return $date.DateTime }
    return $date.LocalDateTime
}

function Get-OrderedPodcastEpisode {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyCollection()][object[]]$Episodes)

    $records = [Collections.Generic.List[object]]::new()
    $position = 0
    foreach ($episode in $Episodes) {
        if ($null -ne $episode) {
            $date = Get-PodcastPublicationUtc -Value $episode.PubDate
            $records.Add([pscustomobject]@{
                Episode = $episode; HasDate = [int]($null -ne $date)
                UtcTicks = $(if ($null -ne $date) { $date.UtcDateTime.Ticks } else { 0L })
                Position = $position
            })
        }
        $position++
    }
    $ordered = @($records | Sort-Object -Property @{ Expression='HasDate'; Descending=$true }, @{ Expression='UtcTicks'; Descending=$true }, @{ Expression='Position'; Descending=$false })
    return ,@($ordered | ForEach-Object { $_.Episode })
}
