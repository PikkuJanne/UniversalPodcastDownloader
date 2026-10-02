#requires -Version 5.1
# Diagnostic messages are authored by the application. Unknown exception bodies,
# response headers and runtime objects must never be supplied as message text.

function Write-PodcastDiagnosticFallback {
    [CmdletBinding()]
    param([string]$Message = 'Diagnostic logging is unavailable: no writable safe destination; continuing without a file log.')

    try { [Console]::Error.WriteLine($Message) }
    catch { return } # A closed stderr must not replace the operation's error.
}

function Get-PodcastSafeUrl {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Url)

    if ($null -eq $script:PodcastDiagnosticRequestIds) {
        $script:PodcastDiagnosticRequestIds = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
    }
    $key = if ($null -eq $Url) { '' } else { $Url }
    if ($key.Length -gt 32768) { return ('unknown-host [url ' + [guid]::NewGuid().ToString('N') + ']') }
    if (-not $script:PodcastDiagnosticRequestIds.ContainsKey($key)) {
        # Bound secret-bearing in-memory correlation keys, too. These keys are
        # never copied to logs or diagnostic exports.
        if ($script:PodcastDiagnosticRequestIds.Count -ge 256) { $script:PodcastDiagnosticRequestIds.Clear() }
        $script:PodcastDiagnosticRequestIds.Add($key, [guid]::NewGuid().ToString('N'))
    }
    $opaqueId = $script:PodcastDiagnosticRequestIds[$key]
    $displayHost = 'unknown-host'
    $parsed = $null
    if ([Uri]::TryCreate($key, [UriKind]::Absolute, [ref]$parsed)) {
        $candidateHost = $parsed.DnsSafeHost
        if ($candidateHost.Length -gt 0 -and $candidateHost -cmatch '^[A-Za-z0-9.\[\]:-]+$') {
            $displayHost = $candidateHost.ToLowerInvariant()
            $address = $null
            if ([Net.IPAddress]::TryParse($displayHost, [ref]$address)) { $displayHost = $address.ToString() }
        }
    }
    return "$displayHost [url $opaqueId]"
}

function Protect-PodcastDiagnosticText {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrEmpty($Text)) { return '' }
    # This one application-authored summary has a complete numeric grammar;
    # no free-form field can be mistaken for a permitted header value.
    if ($Text -cmatch '^Summary: Downloaded=[0-9]{1,9}, Skipped=[0-9]{1,9}, Failed=[0-9]{1,9}, Adopted=[0-9]{1,9}$') { return $Text }
    # This is defense in depth for application-authored text, not a license to
    # publish arbitrary exception messages. Every header name is sensitive.
    $safe = [regex]::Replace($Text, '(?m)^[ \t]*[!#$%&''*+.^_`|~0-9A-Za-z-]+[ \t]*:[^\r\n]*(?:\r?\n[ \t]+[^\r\n]*)*', '[header omitted]')
    $safe = [regex]::Replace($safe, '(?i)(?:\b[a-z][a-z0-9+.-]*:(?://)?[^\s<>"'']+|(?<![\w:])//[^\s<>"'']+|\\\\[^\s<>"'']+)', {
        param($urlMatch)
        Get-PodcastSafeUrl -Url $urlMatch.Value
    })
    # Keep valid surrogate pairs (for example emoji), but remove unpaired
    # surrogates so strict UTF-8 encoders and readers see the same text.
    $safe = [regex]::Replace($safe, '[\p{Cc}\p{Cf}]|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])', ' ')
    if ($safe.Length -gt 2048) {
        $length = if ([char]::IsHighSurrogate($safe[2047])) { 2047 } else { 2048 }
        $safe = $safe.Substring(0, $length) + ' [truncated]'
    }
    return $safe
}

function Get-PodcastDiagnosticError {
    [CmdletBinding()]
    param([Alias('Error', 'ErrorRecord', 'Exception')][AllowNull()]$ErrorValue)

    $cause = if ($ErrorValue -is [Management.Automation.ErrorRecord]) { $ErrorValue.Exception } else { $ErrorValue }
    # Only exact application-authored messages may retain actionable wording.
    # Never accept a prefix match or append details from an external exception.
    $publicMessages = @(
        'Feed catalogue incomplete; accessible selected episodes were processed, but advertised pages remain unresolved.',
        'Feed catalogue incomplete; no accessible episodes were found.',
        'No episodes found in the feed. Double-check the RSS URL.',
        'The RSS or Atom feed is valid but contains no episodes.',
        'Source XML is invalid or exceeds safe parser limits.',
        'The source XML root is not a supported RSS or Atom feed.',
        'No RSS or Atom feed links were found on the page.',
        'Multiple feed links were found. Supply a direct feed URL with -FeedUrl.',
        'The discovered URL did not return an RSS or Atom feed.',
        'HTML metadata exceeds the safe character limit.',
        'HTML feed discovery exceeds the safe token limit.',
        'HTML feed discovery exceeded its safe parser timeout.',
        'Discovered feed URL is not allowed by the network policy.',
        'HTML base URL is not allowed by the network policy.',
        'Source content exceeds the safe character limit.',
        'The metadata response does not contain supported source text.',
        'Feed parsed, but no downloadable enclosure URLs were found.',
        'Feed parsed, but no downloadable enclosure URLs were found. No supported audio candidate was declared.',
        'Resume identity does not match this episode; partial and sidecar were preserved for review.',
        'Legacy archive requires review; use -LegacyPath with -LegacyAction Preview.',
        'Legacy review requires an existing -LegacyPath and an explicit -FeedUrl.',
        'Legacy options require -LegacyPath and an explicit -FeedUrl.',
        'Output root is too long to retain safe identifiers. Choose a shorter output path.',
        "Mode 'Custom' requires -CustomCount with a value >= 1."
    )
    if ($cause -is [Exception]) {
        foreach ($publicMessage in $publicMessages) {
            if ([StringComparer]::Ordinal.Equals($cause.Message, $publicMessage)) { return $publicMessage }
        }
    }
    $depth = 0
    while ($cause -is [Exception] -and $depth -lt 16) {
        # Exact local validation messages map to fixed explanations; no response
        # body, MIME value or URL is incorporated into a diagnostic.
        switch ($cause.Message) {
            'Media validation failed: ambiguous_media.' { return 'Media validation failed: ambiguous_media. Supported audio evidence was not found within the inspection limit.' }
            'Media validation failed: unsupported_media.' { return 'Media validation failed: unsupported_media. The recognized media type is not supported audio.' }
            'Media validation failed: non_audio_text.' { return 'Media validation failed: non_audio_text. The response contains text rather than recognized audio.' }
            'Media validation failed: unrecognized_media.' { return 'Media validation failed: unrecognized_media. The response has no supported audio signature.' }
        }
        if ($cause.Data['PodcastTransport'] -eq $true) {
            # Map fixed categories only; never print the tagged message, header,
            # retry date or any arbitrary data that an exception may carry.
            switch ([string]$cause.Data['Kind']) {
                'Deferred' { return 'The request was deferred because its retry budget could not allow another attempt.' }
                'HeaderTimeout' { return 'The request exceeded the connection/header timeout.' }
                'IdleTimeout' { return 'The response body exceeded the idle transfer timeout.' }
                'Connection' { return 'The network connection failed before the transfer completed.' }
                'HttpStatus' { return 'The server returned an unsuccessful HTTP status.' }
                'IncompleteBody' { return 'The response body was incomplete.' }
                'Permanent' { return 'The request was rejected by the network policy.' }
            }
        }
        if ($null -eq $cause.InnerException) { break }
        $cause = $cause.InnerException
        $depth++
    }
    $category = 'operation error'
    if ($cause -is [UnauthorizedAccessException]) { $category = 'access error' }
    elseif ($cause -is [IO.IOException]) { $category = 'file or transport error' }
    elseif ($cause -is [Net.WebException] -or ($cause -is [Exception] -and $cause.GetType().FullName -eq 'System.Net.Http.HttpRequestException')) { $category = 'network error' }
    elseif ($cause -is [Xml.XmlException]) { $category = 'feed format error' }
    elseif ($cause -is [ArgumentException]) { $category = 'input error' }
    return "The operation failed ($category). Private error details were omitted."
}

function Open-PodcastDiagnosticWriter {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [switch]$CreateDirectory)

    $null = Assert-PodcastDestination -Root $Root -Directory
    if ($CreateDirectory) { $null = [IO.Directory]::CreateDirectory($Root) }
    if (-not [IO.Directory]::Exists($Root)) { throw 'The diagnostic directory does not exist.' }
    $name = 'download-' + $script:PodcastDiagnostics.StartedUtc.ToString('yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture) + '-' + $script:PodcastDiagnostics.RunId + '.log'
    $path = Assert-PodcastDestination -Root $Root -RelativePath $name
    $stream = $null
    try {
        $stream = New-Object IO.FileStream($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
        $writer.AutoFlush = $true
        return [pscustomobject]@{ Path = $path; Writer = $writer }
    }
    catch {
        if ($null -ne $stream) { $stream.Dispose() }
        throw
    }
}

function Initialize-PodcastDiagnostics {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Diagnostics names the run-wide diagnostic context.')]
    [CmdletBinding()]
    param([switch]$Preview, [string[]]$Roots)

    Close-PodcastDiagnostics
    $script:PodcastDiagnosticRequestIds = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
    $script:PodcastDiagnosticPreview = [bool]$Preview
    $script:PodcastDiagnostics = [pscustomobject]@{
        RunId = [guid]::NewGuid().ToString('N')
        StartedUtc = [datetime]::UtcNow
        Preview = [bool]$Preview
        Path = $null
        StartupPath = $null
        Writer = $null
        Events = New-Object 'System.Collections.Generic.List[object]'
        FallbackNotified = $false
    }
    if (-not $Preview) {
        if (-not $PSBoundParameters.ContainsKey('Roots')) {
            $Roots = @()
            try {
                $localRoot = $env:LOCALAPPDATA
                if ([string]::IsNullOrWhiteSpace($localRoot)) { $localRoot = [Environment]::GetFolderPath('LocalApplicationData') }
                $Roots += [IO.Path]::Combine($localRoot, 'UniversalPodcastDownloader\Logs')
            }
            catch { Write-PodcastDiagnosticFallback -Message 'The primary diagnostic location is unavailable; trying temporary storage.' }
            try {
                $tempRoot = $env:TEMP
                if ([string]::IsNullOrWhiteSpace($tempRoot)) { $tempRoot = [IO.Path]::GetTempPath() }
                $Roots += [IO.Path]::Combine($tempRoot, 'UniversalPodcastDownloader\Logs')
            }
            catch { Write-PodcastDiagnosticFallback -Message 'The temporary diagnostic location is unavailable.' }
        }
        foreach ($root in $Roots) {
            try {
                $sink = Open-PodcastDiagnosticWriter -Root $root -CreateDirectory
                $script:PodcastDiagnostics.Writer = $sink.Writer
                $script:PodcastDiagnostics.Path = $sink.Path
                $script:PodcastDiagnostics.StartupPath = $sink.Path
                break
            }
            catch { continue } # Never print directory paths or the sink exception.
        }
        if ($null -eq $script:PodcastDiagnostics.Writer) {
            Write-PodcastDiagnosticFallback
            $script:PodcastDiagnostics.FallbackNotified = $true
        }
    }
    Write-PodcastDiagnostic -Message 'Diagnostic run initialized.' -Code 'run-started'
    return $script:PodcastDiagnostics
}

function Write-PodcastDiagnostic {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$Message,
        [string]$Level = 'INFO',
        [string]$Code = 'message'
    )

    try {
        if ($null -eq $script:PodcastDiagnostics) { return }
        $safeLevel = if (@('INFO', 'WARN', 'ERROR', 'DEBUG') -ccontains $Level) { $Level } else { 'INFO' }
        $safeCode = if (@('run-started', 'archive-log', 'message') -ccontains $Code) { $Code } else { 'message' }
        $safeText = Protect-PodcastDiagnosticText -Text $Message
        $diagnosticEvent = [pscustomobject]@{ TimestampUtc = [datetime]::UtcNow; Level = $safeLevel; Code = $safeCode; Message = $safeText }
        if ($script:PodcastDiagnostics.Events.Count -ge 256) { $script:PodcastDiagnostics.Events.RemoveAt(0) }
        $script:PodcastDiagnostics.Events.Add($diagnosticEvent)
        if ($script:PodcastDiagnostics.Preview) { return }
        if ($null -ne $script:PodcastDiagnostics.Writer) {
            $script:PodcastDiagnostics.Writer.WriteLine(('{0} [{1}] {2}' -f $diagnosticEvent.TimestampUtc.ToString('o', [Globalization.CultureInfo]::InvariantCulture), $safeLevel, $safeText))
        }
        else { Write-PodcastDiagnosticFallback -Message ("[$safeLevel] $safeText") }
    }
    catch {
        Write-PodcastDiagnosticFallback -Message 'Diagnostic logging failed; the operation continues and its original error is preserved.'
        # A failed writer is never retried, and cannot mask the primary error.
        try {
            if ($null -ne $script:PodcastDiagnostics.Writer) { $script:PodcastDiagnostics.Writer.Dispose() }
            $script:PodcastDiagnostics.Writer = $null
        }
        catch { return }
    }
}

function Set-PodcastDiagnosticArchive {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Called only after the caller has confirmed normal archive creation; preview additionally forbids writes here.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $newWriter = $null
    try {
        if ($null -eq $script:PodcastDiagnostics -or $script:PodcastDiagnostics.Preview) { return }
        $sink = Open-PodcastDiagnosticWriter -Root $Root
        $newWriter = $sink.Writer
        foreach ($diagnosticEvent in $script:PodcastDiagnostics.Events) {
            $safeText = Protect-PodcastDiagnosticText -Text $diagnosticEvent.Message
            $safeLevel = if (@('INFO', 'WARN', 'ERROR', 'DEBUG') -ccontains $diagnosticEvent.Level) { $diagnosticEvent.Level } else { 'INFO' }
            $timestamp = if ($diagnosticEvent.TimestampUtc -is [datetime]) { $diagnosticEvent.TimestampUtc.ToUniversalTime() } else { [datetime]::UtcNow }
            $newWriter.WriteLine(('{0} [{1}] {2}' -f $timestamp.ToString('o', [Globalization.CultureInfo]::InvariantCulture), $safeLevel, $safeText))
        }
        if ($null -ne $script:PodcastDiagnostics.Writer) { $script:PodcastDiagnostics.Writer.Dispose() }
        $script:PodcastDiagnostics.Writer = $newWriter
        $script:PodcastDiagnostics.Path = $sink.Path
        $newWriter = $null
        Write-PodcastDiagnostic -Message 'Archive diagnostic log initialized; startup log retained.' -Code 'archive-log'
    }
    catch {
        if ($null -ne $newWriter) {
            try { $newWriter.Dispose() }
            catch { Write-PodcastDiagnosticFallback -Message 'The archive diagnostic log could not be closed.' }
        }
        Write-PodcastDiagnostic -Message 'Archive diagnostic logging is unavailable; retaining the startup log.' -Level WARN
    }
}

function Close-PodcastDiagnostics {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Diagnostics names the run-wide diagnostic context.')]
    [CmdletBinding()]
    param()

    try {
        if ($null -ne $script:PodcastDiagnostics -and $null -ne $script:PodcastDiagnostics.Writer) {
            $script:PodcastDiagnostics.Writer.Dispose()
            $script:PodcastDiagnostics.Writer = $null
        }
    }
    catch { Write-PodcastDiagnosticFallback -Message 'Diagnostic logging could not be closed; the original operation result is preserved.' }
}

function Export-PodcastDiagnostics {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Diagnostics names the deliberately restricted diagnostic export.')]
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Path)

    # Do not enumerate archives, deserialize runtime state or read any log. In
    # particular, free-form messages and extra object properties are excluded.
    if ($null -eq $script:PodcastDiagnostics) { throw 'There is no diagnostic run to export.' }
    if ($script:PodcastDiagnosticPreview -or $script:PodcastDiagnostics.Preview) { return }
    $context = $script:PodcastDiagnostics
    if ($context.RunId -isnot [string] -or $context.RunId -cnotmatch '^[a-f0-9]{32}$' -or $context.StartedUtc -isnot [datetime]) {
        throw 'Diagnostic run metadata is invalid; export was not written.'
    }
    $exportEvents = New-Object 'System.Collections.Generic.List[object]'
    foreach ($diagnosticEvent in $context.Events) {
        if ($exportEvents.Count -ge 256) { break }
        if ($diagnosticEvent.TimestampUtc -isnot [datetime] -or $diagnosticEvent.Level -isnot [string] -or $diagnosticEvent.Code -isnot [string]) { continue }
        if (@('INFO', 'WARN', 'ERROR', 'DEBUG') -cnotcontains $diagnosticEvent.Level -or @('run-started', 'archive-log', 'message') -cnotcontains $diagnosticEvent.Code) { continue }
        $exportEvents.Add([ordered]@{
            timestamp_utc = $diagnosticEvent.TimestampUtc.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture)
            level = $diagnosticEvent.Level
            code = $diagnosticEvent.Code
        })
    }
    $payload = [ordered]@{
        schema_version = 1
        run_id = $context.RunId
        started_utc = $context.StartedUtc.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        engine_version = $PSVersionTable.PSVersion.ToString()
        events = @($exportEvents.ToArray())
    }
    if (-not $PSCmdlet.ShouldProcess('diagnostic export', 'Create a new privacy-filtered JSON file')) { return }
    $stream = $null
    $writer = $null
    try {
        $parent = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Path))
        if (-not [IO.Directory]::Exists($parent)) { throw 'The export directory does not exist.' }
        $target = Assert-PodcastDestination -Root $parent -RelativePath ([IO.Path]::GetFileName($Path))
        $stream = New-Object IO.FileStream($target, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
        $writer.WriteLine(($payload | ConvertTo-Json -Depth 5))
        $writer.Flush()
    }
    catch { throw 'The diagnostic export could not be written. Choose a new file in an accessible ordinary directory.' }
    finally {
        if ($null -ne $writer) { $writer.Dispose() }
        elseif ($null -ne $stream) { $stream.Dispose() }
    }
}
