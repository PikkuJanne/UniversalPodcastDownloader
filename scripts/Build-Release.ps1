#requires -Version 5.1
<#
.SYNOPSIS
Builds an unpublished portable candidate from a clean Git commit.
.DESCRIPTION
Exports the exact checked-in allowlist directly from Git blobs, creates a
versioned ZIP and matching manifest, and verifies SHA-256 checksums before
moving an owned staging directory into place. Requires Git and built-in .NET;
the resulting downloader has no dependency on Git or development modules.
Archive timestamps are fixed. Byte equality across compression implementations
is not guaranteed; inventory, source bytes and manifest are deterministic.
.PARAMETER OutputDirectory
Parent directory for a new candidate directory. Defaults to artifacts/releases
inside this checkout. Existing candidate directories are never overwritten.
.PARAMETER SourceCommit
Full 40-character commit ID in this repository. Defaults to HEAD. Version and
inventory are read from that commit, and the current checkout must be clean.
.EXAMPLE
pwsh -NoProfile -File ./scripts/Build-Release.ps1
.EXAMPLE
powershell -NoProfile -File ./scripts/Build-Release.ps1 -OutputDirectory 'D:\Release candidates'
#>
[CmdletBinding()]
param(
    [string]$OutputDirectory,
    [string]$SourceCommit
)

$ErrorActionPreference = 'Stop'

function Assert-UpdReleaseSafePath {
    param([Parameter(Mandatory)][string]$Path)
    $current = [IO.Path]::GetFullPath($Path)
    while ($current) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Reparse point is not permitted in a release path: $current"
            }
        }
        $parent = [IO.Directory]::GetParent($current)
        $current = if ($parent) { $parent.FullName } else { $null }
    }
}

function Invoke-UpdReleaseGit {
    param(
        [Parameter(Mandatory)][string]$Repository,
        [Parameter(Mandatory)][string[]]$Arguments,
        [string]$BinaryOutput
    )
    $start = New-Object System.Diagnostics.ProcessStartInfo
    $start.FileName = 'git'
    $values = @('--no-replace-objects', '-C', $Repository) + $Arguments
    # Arguments are internal fixed options, validated hashes or a repository
    # path; Windows directory paths cannot contain a double quote.
    foreach ($value in $values) {
        if ($value.Contains('"') -or $value.Contains("`r") -or $value.Contains("`n")) { throw 'Invalid Git argument.' }
    }
    $start.Arguments = ($values | ForEach-Object { '"' + $_.TrimEnd('\') + '"' }) -join ' '
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $start
    $destination = $null
    try {
        [void]$process.Start()
        $errors = $process.StandardError.ReadToEndAsync()
        if ($BinaryOutput) {
            $destination = New-Object IO.FileStream($BinaryOutput, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            $process.StandardOutput.BaseStream.CopyTo($destination)
            $destination.Dispose()
            $destination = $null
            $result = $null
        }
        else { $result = $process.StandardOutput.ReadToEnd() }
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw ('Git failed: ' + $errors.Result.Trim()) }
        return $result
    }
    finally {
        if ($destination) { $destination.Dispose() }
        $process.Dispose()
    }
}

function Get-UpdReleaseHash {
    param([Parameter(Mandatory)][IO.Stream]$Stream)
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hasher.ComputeHash($Stream))).Replace('-', '').ToLowerInvariant() }
    finally { $hasher.Dispose() }
}

function Get-UpdReleaseFileHash {
    param([Parameter(Mandatory)][string]$Path)
    $stream = [IO.File]::OpenRead($Path)
    try { return Get-UpdReleaseHash -Stream $stream }
    finally { $stream.Dispose() }
}

function Get-UpdReleaseOrderedPath {
    param([string[]]$Paths)
    $sorted = [string[]]$Paths.Clone()
    [Array]::Sort($sorted, [StringComparer]::Ordinal)
    return $sorted
}

function Invoke-UpdReleaseStageCleanup {
    param([string]$Path, [string]$Owner)
    if (-not $Path -or -not [IO.Directory]::Exists($Path)) { return }
    Assert-UpdReleaseSafePath -Path $Path
    $marker = Join-Path $Path '.upd-release-owner'
    if (-not [IO.File]::Exists($marker) -or [IO.File]::ReadAllText($marker) -cne $Owner) {
        throw "Release staging ownership changed; retained $Path"
    }
    $directories = New-Object 'Collections.Generic.Queue[string]'
    $directories.Enqueue($Path)
    while ($directories.Count) {
        foreach ($item in @(Get-ChildItem -LiteralPath $directories.Dequeue() -Force)) {
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Release staging changed; retained $Path" }
            if ($item.PSIsContainer) { $directories.Enqueue($item.FullName) }
        }
    }
    Remove-Item -LiteralPath $Path -Force -Recurse -ErrorAction Stop
}

$stage = $null
$owner = [guid]::NewGuid().ToString('N')
$publishedDirectory = $null
try {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $repo = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent)).TrimEnd('\', '/')
    Assert-UpdReleaseSafePath -Path $repo
    $actualRepo = (Invoke-UpdReleaseGit -Repository $repo -Arguments @('rev-parse', '--show-toplevel')).Trim().Replace('/', '\')
    if (-not $actualRepo.Equals($repo, [StringComparison]::OrdinalIgnoreCase)) { throw 'Build script must reside at the Git repository root scripts directory.' }
    $status = Invoke-UpdReleaseGit -Repository $repo -Arguments @('status', '--porcelain=v1', '--untracked-files=all')
    if ($status.Trim()) { throw 'Release source must be clean: commit tracked changes and remove or explicitly ignore untracked files.' }
    if (-not $SourceCommit) { $SourceCommit = (Invoke-UpdReleaseGit -Repository $repo -Arguments @('rev-parse', 'HEAD')).Trim() }
    if ($SourceCommit -cnotmatch '^[0-9a-fA-F]{40}$') { throw 'SourceCommit must be a full 40-character commit ID.' }
    $SourceCommit = $SourceCommit.ToLowerInvariant()
    $verifiedCommit = (Invoke-UpdReleaseGit -Repository $repo -Arguments @('rev-parse', '--verify', ($SourceCommit + '^{commit}'))).Trim()
    if ($verifiedCommit -cne $SourceCommit) { throw 'SourceCommit does not resolve to the specified commit.' }
    $sourceTree = (Invoke-UpdReleaseGit -Repository $repo -Arguments @('rev-parse', ($SourceCommit + '^{tree}'))).Trim()
    $treeText = Invoke-UpdReleaseGit -Repository $repo -Arguments @('ls-tree', '-r', '--full-tree', $SourceCommit)
    $tracked = @{}
    foreach ($line in @($treeText -split "`n")) {
        if (-not $line.Trim()) { continue }
        if ($line.TrimEnd("`r") -cnotmatch '^([0-9]{6}) ([a-z]+) ([0-9a-f]{40})\t(.+)$') { throw 'Unsupported Git tree entry.' }
        $tracked[$Matches[4]] = [pscustomobject]@{ Mode = $Matches[1]; Type = $Matches[2]; Blob = $Matches[3] }
    }
    $configPath = 'tools/release-package.json'
    if (-not $tracked.ContainsKey($configPath) -or $tracked[$configPath].Mode -notin @('100644', '100755')) { throw 'Missing committed release-package.json.' }
    $config = (Invoke-UpdReleaseGit -Repository $repo -Arguments @('cat-file', 'blob', $tracked[$configPath].Blob)) | ConvertFrom-Json
    $configKeys = @(Get-UpdReleaseOrderedPath -Paths $config.PSObject.Properties.Name)
    if (($configKeys -join ',') -cne 'files,packageName,releaseStatus,schemaVersion,sourceRepository,version' -or
        $config.schemaVersion -ne 1 -or $config.packageName -cne 'UniversalPodcastDownloader' -or
        $config.releaseStatus -cne 'UNRELEASED_CANDIDATE' -or
        $config.sourceRepository -cne 'https://github.com/PikkuJanne/UniversalPodcastDownloader') { throw 'Invalid committed release configuration.' }
    if ($config.version -cnotmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)-rc\.([1-9][0-9]*)$') { throw 'Release version must be an unpublished semantic-version rc candidate.' }
    $requiredRoot = @('CHANGELOG.md', 'LICENSE', 'README.md', 'RELEASE.md', 'UniversalPodcastDownloader.bat',
        'UniversalPodcastDownloader.ico', 'UniversalPodcastDownloader.ps1', 'UniversalPodcastDownloader_icon.png', 'UniversalPodcastDownloader_poster.png')
    $sourceFiles = @(Get-UpdReleaseOrderedPath -Paths @($tracked.Keys | Where-Object { $_.StartsWith('src/', [StringComparison]::Ordinal) }))
    if ($sourceFiles.Count -ne 27 -or @($sourceFiles | Where-Object { $_ -cnotmatch '^src/[A-Za-z]+\.ps1$' }).Count) { throw 'Committed src inventory must contain exactly 27 runtime PowerShell files.' }
    $expected = @(Get-UpdReleaseOrderedPath -Paths @($requiredRoot + $sourceFiles))
    $files = @($config.files)
    if ($files.Count -ne $expected.Count -or ($files -join "`n") -cne ($expected -join "`n")) { throw 'Release allowlist must be sorted and contain the exact root and complete src inventory.' }
    foreach ($path in $files) {
        if (-not $tracked.ContainsKey($path)) { throw "Missing committed payload file: $path" }
        if ($tracked[$path].Type -cne 'blob' -or $tracked[$path].Mode -notin @('100644', '100755')) { throw "Payload must be a regular Git blob: $path" }
        Assert-UpdReleaseSafePath -Path (Join-Path $repo $path)
    }
    if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo 'artifacts/releases' }
    $OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\', '/')
    $rootPath = [IO.Path]::GetPathRoot($OutputDirectory).TrimEnd('\', '/')
    if ($OutputDirectory -eq $rootPath -or $repo.Equals($OutputDirectory, [StringComparison]::OrdinalIgnoreCase) -or
        $repo.StartsWith($OutputDirectory + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'OutputDirectory cannot be a volume root, repository root or repository ancestor.' }
    if ($OutputDirectory.StartsWith($repo + '\', [StringComparison]::OrdinalIgnoreCase) -and
        -not $OutputDirectory.StartsWith($repo + '\artifacts\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Output inside this repository must be under ignored artifacts/.' }
    Assert-UpdReleaseSafePath -Path $OutputDirectory
    if ((Test-Path -LiteralPath $OutputDirectory) -and -not [IO.Directory]::Exists($OutputDirectory)) { throw 'OutputDirectory must be a directory.' }
    $candidateName = $config.packageName + '-' + $config.version + '-' + $SourceCommit
    $candidatePath = Join-Path $OutputDirectory $candidateName
    if (Test-Path -LiteralPath $candidatePath) { throw "Candidate output already exists; nothing overwritten: $candidatePath" }
    [void][IO.Directory]::CreateDirectory($OutputDirectory)
    Assert-UpdReleaseSafePath -Path $OutputDirectory
    $stage = Join-Path $OutputDirectory ('.upd-release-' + $owner)
    $null = New-Item -ItemType Directory -Path $stage -ErrorAction Stop
    [IO.File]::WriteAllText((Join-Path $stage '.upd-release-owner'), $owner)
    $payload = Join-Path $stage 'payload'
    $ready = Join-Path $stage 'ready'
    [void][IO.Directory]::CreateDirectory($payload)
    [void][IO.Directory]::CreateDirectory($ready)
    $records = @()
    foreach ($path in $files) {
        $destination = Join-Path $payload $path
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
        Invoke-UpdReleaseGit -Repository $repo -Arguments @('cat-file', 'blob', $tracked[$path].Blob) -BinaryOutput $destination
        $records += [ordered]@{ path = $path; sha256 = (Get-UpdReleaseFileHash -Path $destination); size = ([IO.FileInfo]$destination).Length }
    }
    $manifest = [ordered]@{ schemaVersion = 1; packageName = $config.packageName; version = $config.version;
        releaseStatus = $config.releaseStatus; sourceRepository = $config.sourceRepository;
        sourceCommit = $SourceCommit; sourceTree = $sourceTree;
        sourceCommitUrl = ($config.sourceRepository + '/commit/' + $SourceCommit); files = @($records) }
    $utf8 = New-Object Text.UTF8Encoding($false)
    $manifestBytes = $utf8.GetBytes(($manifest | ConvertTo-Json -Depth 6 -Compress) + "`n")
    $manifestPath = Join-Path $ready 'manifest.json'
    [IO.File]::WriteAllBytes($manifestPath, $manifestBytes)
    $zipName = $config.packageName + '-' + $config.version + '.zip'
    $zipPath = Join-Path $ready $zipName
    $zipStream = New-Object IO.FileStream($zipPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $zip = $null
    try {
        $zip = New-Object IO.Compression.ZipArchive($zipStream, [IO.Compression.ZipArchiveMode]::Create, $true)
        foreach ($path in @(Get-UpdReleaseOrderedPath -Paths @($files + 'manifest.json'))) {
            $entry = $zip.CreateEntry($path, [IO.Compression.CompressionLevel]::Optimal)
            $entry.LastWriteTime = [DateTimeOffset]::new(1980, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
            $entry.ExternalAttributes = 0
            $entryStream = $entry.Open()
            try {
                if ($path -ceq 'manifest.json') { $entryStream.Write($manifestBytes, 0, $manifestBytes.Length) }
                else {
                    $inputStream = [IO.File]::OpenRead((Join-Path $payload $path))
                    try { $inputStream.CopyTo($entryStream) } finally { $inputStream.Dispose() }
                }
            }
            finally { $entryStream.Dispose() }
        }
    }
    finally {
        if ($zip) { $zip.Dispose() }
        $zipStream.Dispose()
    }
    # Reopen the actual ZIP and validate every final entry before publishing.
    $archive = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $entryPaths = @(Get-UpdReleaseOrderedPath -Paths $archive.Entries.FullName)
        if (($entryPaths -join "`n") -cne (@(Get-UpdReleaseOrderedPath -Paths @($files + 'manifest.json')) -join "`n")) { throw 'ZIP inventory validation failed.' }
        foreach ($entry in $archive.Entries) {
            $stream = $entry.Open()
            try { $hash = Get-UpdReleaseHash -Stream $stream } finally { $stream.Dispose() }
            if ($entry.FullName -ceq 'manifest.json') {
                if ($hash -cne (Get-UpdReleaseFileHash -Path $manifestPath)) { throw 'ZIP manifest validation failed.' }
            }
            else {
                $record = @($records | Where-Object { $_.path -ceq $entry.FullName })[0]
                if ($hash -cne $record.sha256 -or $entry.Length -ne $record.size) { throw "ZIP payload validation failed: $($entry.FullName)" }
            }
        }
    }
    finally { $archive.Dispose() }
    $checksumPath = Join-Path $ready 'SHA256SUMS'
    [IO.File]::WriteAllText($checksumPath, ((Get-UpdReleaseFileHash -Path $zipPath) + '  ' + $zipName + "`n" +
        (Get-UpdReleaseFileHash -Path $manifestPath) + "  manifest.json`n"), $utf8)
    $finalStatus = Invoke-UpdReleaseGit -Repository $repo -Arguments @('status', '--porcelain=v1', '--untracked-files=all')
    if ($finalStatus.Trim()) { throw 'Release source changed during the build; no candidate finalized.' }
    Assert-UpdReleaseSafePath -Path $candidatePath
    [IO.Directory]::Move($ready, $candidatePath)
    $publishedDirectory = $candidatePath
    Invoke-UpdReleaseStageCleanup -Path $stage -Owner $owner
    $stage = $null
    [ordered]@{ version = $config.version; sourceCommit = $SourceCommit; sourceTree = $sourceTree;
        outputDirectory = $candidatePath; zipPath = (Join-Path $candidatePath $zipName);
        manifestPath = (Join-Path $candidatePath 'manifest.json'); checksumPath = (Join-Path $candidatePath 'SHA256SUMS') } | ConvertTo-Json -Compress
    exit 0
}
catch {
    $failure = $_.Exception.Message
    if ($publishedDirectory) { $failure += "; finalized candidate retained at $publishedDirectory" }
    if ($stage) {
        try { Invoke-UpdReleaseStageCleanup -Path $stage -Owner $owner }
        catch { $failure += '; staging cleanup failed: ' + $_.Exception.Message }
    }
    Write-Error -Message $failure -ErrorAction Continue
    exit 1
}
