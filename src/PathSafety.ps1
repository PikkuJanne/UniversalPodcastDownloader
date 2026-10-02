#requires -Version 5.1

function Assert-PodcastPathComponent {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Component)

    # Windows compares these names after normalization. Reject ambiguous names
    # before GetFullPath can hide a traversal or Win32 can trim a suffix.
    if ([string]::IsNullOrEmpty($Component) -or $Component -eq '.' -or $Component -eq '..' -or
        $Component -match '[<>:"/\\|?*\x00-\x1f]' -or $Component -match '[. ]$') {
        throw 'The destination contains an invalid or ambiguous Windows path component.'
    }
    if ($Component.Length -gt 255) {
        throw 'A destination component exceeds the Windows limit of 255 characters.'
    }
    if ($Component -match '^(?i:CON|PRN|AUX|NUL|CONIN\$|CONOUT\$|COM[1-9\u00b9\u00b2\u00b3]|LPT[1-9\u00b9\u00b2\u00b3])(?:[. ]|$)') {
        throw 'A destination component is a reserved Windows device name.'
    }
}

function Get-PodcastCanonicalRoot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'The output root must be an absolute Windows path.' }
    $normalized = $Path.Replace('/', '\')
    # Extended/device namespaces bypass normal Win32 validation and are not
    # supported by both downloader engines. Drive-relative paths are ambiguous.
    if ($normalized -match '^\\\\[?.]\\' -or
        ($normalized -notmatch '^[A-Za-z]:\\' -and $normalized -notmatch '^\\\\[^\\]+\\[^\\]+(?:\\|$)')) {
        throw 'The output root must be a normal absolute drive or UNC path; device and relative paths are unsupported.'
    }
    $pathRoot = [IO.Path]::GetPathRoot($normalized)
    $rootPrefixLength = $pathRoot.Length
    if ($normalized.StartsWith('\\', [StringComparison]::Ordinal)) {
        foreach ($part in $pathRoot.TrimEnd([char]92).Substring(2).Split([char]92)) {
            Assert-PodcastPathComponent -Component $part
        }
        # Framework and modern .NET disagree about including the UNC share's
        # final separator in GetPathRoot. Parse that separator explicitly.
        $rootPrefixLength = [regex]::Match($normalized, '^\\\\[^\\]+\\[^\\]+(?:\\|$)').Length
    }
    $remainder = $normalized.Substring($rootPrefixLength)
    if ($remainder.StartsWith('\', [StringComparison]::Ordinal)) {
        throw 'The output root contains an empty Windows path component.'
    }
    # A single final separator is harmless on a caller-controlled output root.
    if ($remainder.EndsWith('\', [StringComparison]::Ordinal)) {
        $remainder = $remainder.Substring(0, $remainder.Length - 1)
    }
    if ($remainder.Length -gt 0) {
        foreach ($part in $remainder.Split([char]92)) { Assert-PodcastPathComponent -Component $part }
    }
    $canonical = [IO.Path]::GetFullPath($normalized)
    if ($canonical.Length -gt $pathRoot.Length) { $canonical = $canonical.TrimEnd([char]92) }
    if ($canonical.Length -gt 247) {
        throw 'The output root exceeds the supported Windows directory path limit of 247 characters.'
    }
    return $canonical
}

function Get-PodcastDestination {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][AllowEmptyString()][string]$RelativePath,
        [switch]$Directory
    )

    $canonicalRoot = Get-PodcastCanonicalRoot -Path $Root
    $relative = $RelativePath.Replace('/', '\')
    if ([string]::IsNullOrEmpty($relative) -or [IO.Path]::IsPathRooted($relative)) {
        throw 'A metadata destination must be a nonempty relative path.'
    }
    foreach ($part in $relative.Split([char]92)) { Assert-PodcastPathComponent -Component $part }
    $prefix = $canonicalRoot.TrimEnd([char]92) + '\'
    $target = [IO.Path]::GetFullPath($prefix + $relative)
    if (-not $target.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The destination escapes the configured output root.'
    }
    # MAX_PATH counts its terminating NUL. Directory creation needs another
    # twelve characters of headroom on the legacy Windows API used by PS5.1.
    $limit = if ($Directory) { 247 } else { 259 }
    if ($target.Length -gt $limit) {
        throw "The destination exceeds the supported Windows path limit of $limit characters. Choose a shorter output root."
    }
    return $target
}

function Get-PodcastPathAttribute {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LiteralPath)

    try {
        # GetAttributes inspects reparse points themselves, including dangling
        # links. Test-Path alone can mistake a dangling link for an absent path.
        return [IO.File]::GetAttributes($LiteralPath)
    }
    catch {
        $cause = $_.Exception
        while ($null -ne $cause.InnerException) { $cause = $cause.InnerException }
        if ($cause -is [IO.FileNotFoundException] -or $cause -is [IO.DirectoryNotFoundException]) { return $null }
        throw 'Cannot inspect the destination safely. Check output-folder access.'
    }
}

function Assert-PodcastDestination {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [AllowEmptyString()][string]$RelativePath = '',
        [switch]$Directory
    )

    $canonicalRoot = Get-PodcastCanonicalRoot -Path $Root
    $target = if ($RelativePath.Length -gt 0) {
        Get-PodcastDestination -Root $canonicalRoot -RelativePath $RelativePath -Directory:$Directory
    }
    else { $canonicalRoot }
    $volumeRoot = [IO.Path]::GetPathRoot($target)
    $current = $volumeRoot
    $paths = New-Object 'System.Collections.Generic.List[string]'
    $paths.Add($current)
    $remainder = $target.Substring($volumeRoot.Length).TrimStart([char]92)
    if ($remainder.Length -gt 0) {
        foreach ($part in $remainder.Split([char]92)) {
            $current = [IO.Path]::Combine($current, $part)
            $paths.Add($current)
        }
    }
    $parent = $null
    $parentExists = $false
    foreach ($candidate in $paths) {
        $attributes = Get-PodcastPathAttribute -LiteralPath $candidate
        if ($parentExists) {
            # NTFS can enable case sensitivity per directory. Treat siblings
            # case-insensitively even there so a differently cased file or link
            # cannot turn an existing-file check into a new destination.
            $name = [IO.Path]::GetFileName($candidate)
            $matchCount = 0
            try {
                foreach ($entry in [IO.Directory]::EnumerateFileSystemEntries($parent)) {
                    if ([string]::Equals([IO.Path]::GetFileName($entry), $name, [StringComparison]::OrdinalIgnoreCase)) {
                        $matchCount++
                    }
                }
            }
            catch { throw 'Cannot inspect destination siblings safely. Check output-folder access.' }
            if ($matchCount -gt 1 -or ($matchCount -eq 1 -and $null -eq $attributes)) {
                throw 'The destination conflicts with an existing name when compared case-insensitively.'
            }
        }
        $parent = $candidate
        $parentExists = $null -ne $attributes
        if ($null -eq $attributes) { continue }
        if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'The destination traverses a reparse point (junction or symbolic link). Choose an ordinary output directory.'
        }
        if (($candidate -ne $target -or $Directory -or $RelativePath.Length -eq 0) -and
            ($attributes -band [IO.FileAttributes]::Directory) -eq 0) {
            throw 'A destination ancestor or output directory is an existing file.'
        }
        if ($candidate -eq $target -and -not $Directory -and $RelativePath.Length -gt 0 -and
            ($attributes -band [IO.FileAttributes]::Directory) -ne 0) {
            throw 'The media or log destination is an existing directory.'
        }
    }
    return $target
}
