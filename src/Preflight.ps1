#requires -Version 5.1

function Get-PodcastAvailableSpace {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $canonical = Get-PodcastCanonicalRoot -Path $Root
    # DriveInfo cannot reliably query an arbitrary UNC share on both engines.
    # Unknown capacity is distinct from an observed zero-byte capacity.
    if ($canonical.StartsWith('\\', [StringComparison]::Ordinal)) { return $null }
    try {
        $drive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($canonical))
        if (-not $drive.IsReady) { return $null }
        # AvailableFreeSpace accounts for the caller's quota, unlike TotalFreeSpace.
        [long]$available = $drive.AvailableFreeSpace
        if ($available -lt 0) { return $null }
        return $available
    }
    catch {
        if (Test-PodcastCancellation -ErrorObject $_) { throw }
        return $null
    }
}

function Assert-PodcastTransferSpace {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Nullable[long]]$RequiredBytes,
        [Nullable[long]]$AvailableBytes
    )

    $canonical = Assert-PodcastDestination -Root $Root -Directory
    if (($null -ne $RequiredBytes -and $RequiredBytes -lt 0) -or
        ($null -ne $AvailableBytes -and $AvailableBytes -lt 0)) {
        throw 'Transfer space estimates must be nonnegative byte counts.'
    }
    if (-not $PSBoundParameters.ContainsKey('AvailableBytes')) {
        $AvailableBytes = Get-PodcastAvailableSpace -Root $canonical
    }
    if ($null -ne $AvailableBytes -and $AvailableBytes -lt 0) {
        throw 'Transfer space estimates must be nonnegative byte counts.'
    }
    $check = 'sufficient'
    $message = 'Available space meets the validated response size; capacity can change during transfer.'
    if ($null -eq $RequiredBytes) {
        $check = 'unknown_size'
        $message = 'The media response size is unknown; available space cannot be compared.'
    }
    elseif ($null -eq $AvailableBytes) {
        $check = 'unknown_available'
        $message = 'Available destination space is unknown; the validated response size cannot be compared.'
    }
    elseif ($RequiredBytes -gt $AvailableBytes) {
        throw 'Insufficient available space for the validated media response. Free space or choose another output folder.'
    }
    return [pscustomobject]@{
        RequiredBytes = $RequiredBytes; AvailableBytes = $AvailableBytes
        SpaceCheck = $check; Message = $message
    }
}

function Get-PodcastPreflightProbeName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Directory)

    # Keep even a 247-character directory inside the existing legacy MAX_PATH.
    $remaining = 259 - $Directory.TrimEnd([char]92).Length - 1 - 4
    if ($remaining -lt 4) { throw 'The destination is not writable. Check output-folder permissions and retry.' }
    return '.pf-' + [guid]::NewGuid().ToString('N').Substring(0, [Math]::Min(32, $remaining))
}

function Open-PodcastPreflightProbe {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only an exclusive owned DeleteOnClose probe after the caller accepts archive execution.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LiteralPath)

    return [IO.FileStream]::new($LiteralPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite,
        [IO.FileShare]::None, 4096, [IO.FileOptions]::DeleteOnClose)
}

function Invoke-PodcastDestinationPreflight {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [switch]$Preview)

    $canonical = Assert-PodcastDestination -Root $Root -Directory
    if ($Preview) {
        return [pscustomobject]@{ Writable = $null; Probed = $false; Preview = $true }
    }
    $directory = $canonical
    while ($null -eq (Get-PodcastPathAttribute -LiteralPath $directory)) {
        $directory = [IO.Path]::GetDirectoryName($directory)
        if ([string]::IsNullOrEmpty($directory)) { throw 'The destination is not writable. Check output-folder permissions and retry.' }
    }
    $directory = Assert-PodcastDestination -Root $directory -Directory
    $probe = $null
    $primaryError = $null
    $cleanupError = $null
    try {
        for ($attempt = 0; $attempt -lt 8; $attempt++) {
            $name = Get-PodcastPreflightProbeName -Directory $directory
            $probePath = Assert-PodcastDestination -Root $directory -RelativePath $name
            # Existing names are never candidates for ownership or cleanup.
            if ($null -ne (Get-PodcastPathAttribute -LiteralPath $probePath)) { continue }
            try { $probe = Open-PodcastPreflightProbe -LiteralPath $probePath }
            catch {
                if (Test-PodcastCancellation -ErrorObject $_) { throw }
                # A CreateNew collision can arise after the preceding observation.
                if ($null -ne (Get-PodcastPathAttribute -LiteralPath $probePath)) { continue }
                throw
            }
            if ($null -eq $probe) { throw 'The destination is not writable. Check output-folder permissions and retry.' }
            break
        }
        if ($null -eq $probe) { throw 'The destination is not writable. Check output-folder permissions and retry.' }
        $probe.WriteByte(0)
        $probe.Flush($true)
    }
    catch { $primaryError = $_ }
    finally {
        if ($null -ne $probe) {
            try { $probe.Dispose() }
            catch { $cleanupError = $_ }
        }
    }
    if ($null -ne $primaryError) {
        if (Test-PodcastCancellation -ErrorObject $primaryError) { throw $primaryError }
        throw 'The destination is not writable. Check output-folder permissions and retry.'
    }
    if ($null -ne $cleanupError) {
        if (Test-PodcastCancellation -ErrorObject $cleanupError) { throw $cleanupError }
        throw 'The destination is not writable. Check output-folder permissions and retry.'
    }
    return [pscustomobject]@{ Writable = $true; Probed = $true; Preview = $false }
}
