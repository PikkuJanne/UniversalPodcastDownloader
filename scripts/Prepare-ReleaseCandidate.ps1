#requires -Version 5.1
<#
.SYNOPSIS
Builds and validates the three unpublished candidate files permitted in CI.
.DESCRIPTION
Requires the exact checked-out commit tested by the Windows engine matrix.
Runs the existing clean-source builder in a child process, then independently
checks the output allowlist, source identifiers, ZIP inventory and checksums.
No tag, GitHub release, credential or repository setting is written. Dot-source
this development-only script to load its validation functions without building.
.PARAMETER SourceCommit
Full lowercase 40-character commit ID. It must equal this checkout's HEAD.
.PARAMETER OutputDirectory
Parent directory for a new candidate, subject to the release builder's guards.
.PARAMETER GitHubOutput
Optional Actions output file. Safe single-line paths are appended only after
all candidate checks pass; no application logs or test results are exported.
#>
[CmdletBinding()]
param(
    [string]$SourceCommit,
    [string]$OutputDirectory,
    [string]$GitHubOutput
)

function Assert-UpdCandidateSafePath {
    param([Parameter(Mandatory)][string]$Path)
    $current = [IO.Path]::GetFullPath($Path)
    while ($current) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Candidate paths cannot contain reparse points.' }
        }
        $parent = [IO.Directory]::GetParent($current)
        $current = if ($parent) { $parent.FullName } else { $null }
    }
}

function Invoke-UpdCandidateGit {
    param([Parameter(Mandatory)][string]$Repository, [Parameter(Mandatory)][string[]]$Arguments)
    $result = & git --no-replace-objects -C $Repository @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) { throw 'Candidate source Git verification failed.' }
    return ($result -join "`n").Trim()
}

function Get-UpdCandidateSource {
    param([Parameter(Mandatory)][string]$Repository, [Parameter(Mandatory)][string]$SourceCommit)
    if ($SourceCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'SourceCommit must be a full lowercase 40-character commit ID.' }
    $repo = [IO.Path]::GetFullPath($Repository).TrimEnd('\', '/')
    Assert-UpdCandidateSafePath -Path $repo
    $actualRoot = (Invoke-UpdCandidateGit -Repository $repo -Arguments @('rev-parse', '--show-toplevel')).Replace('/', '\')
    if (-not $actualRoot.Equals($repo, [StringComparison]::OrdinalIgnoreCase)) { throw 'Candidate repository must be its Git root.' }
    $headCommit = Invoke-UpdCandidateGit -Repository $repo -Arguments @('rev-parse', 'HEAD')
    if ($headCommit -cne $SourceCommit) { throw 'Candidate source must equal the checked-out tested commit.' }
    $tree = Invoke-UpdCandidateGit -Repository $repo -Arguments @('rev-parse', ($SourceCommit + '^{tree}'))
    $config = (Invoke-UpdCandidateGit -Repository $repo -Arguments @('show', ($SourceCommit + ':tools/release-package.json'))) | ConvertFrom-Json
    if ($config.schemaVersion -ne 1 -or $config.packageName -cne 'UniversalPodcastDownloader' -or
        $config.releaseStatus -cne 'UNRELEASED_CANDIDATE' -or
        $config.sourceRepository -cne 'https://github.com/PikkuJanne/UniversalPodcastDownloader' -or
        $config.version -cnotmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)-rc\.([1-9][0-9]*)$') { throw 'Invalid committed candidate configuration.' }
    $rootPaths = @('CHANGELOG.md', 'LICENSE', 'README.md', 'RELEASE.md', 'UniversalPodcastDownloader.bat',
        'UniversalPodcastDownloader.ico', 'UniversalPodcastDownloader.ps1', 'UniversalPodcastDownloader_icon.png', 'UniversalPodcastDownloader_poster.png')
    $treePaths = @((Invoke-UpdCandidateGit -Repository $repo -Arguments @('ls-tree', '-r', '--name-only', $SourceCommit)) -split "`n")
    $sourcePaths = @($treePaths | Where-Object { $_.StartsWith('src/', [StringComparison]::Ordinal) })
    if ($sourcePaths.Count -ne 27 -or @($sourcePaths | Where-Object { $_ -cnotmatch '^src/[A-Za-z]+\.ps1$' }).Count) { throw 'Invalid committed candidate runtime inventory.' }
    $expected = [string[]]@($rootPaths + $sourcePaths)
    [Array]::Sort($expected, [StringComparer]::Ordinal)
    if (@($config.files).Count -ne 36 -or ($config.files -join "`n") -cne ($expected -join "`n")) { throw 'Invalid committed candidate allowlist.' }
    $blobs = @{}
    foreach ($line in @((Invoke-UpdCandidateGit -Repository $repo -Arguments @('ls-tree', '-r', '--full-tree', $SourceCommit)) -split "`n")) {
        if ($line -cmatch '^(100644|100755) blob ([0-9a-f]{40})\t(.+)$') { $blobs[$Matches[3]] = $Matches[2] }
    }
    foreach ($path in $config.files) {
        if (-not $blobs.ContainsKey($path)) { throw 'Candidate allowlist requires regular committed Git blobs.' }
    }
    return [pscustomobject]@{ Repository = $repo; Commit = $SourceCommit; Tree = $tree; Config = $config; Blobs = $blobs }
}

function Get-UpdCandidateStreamHash {
    param([Parameter(Mandatory)][IO.Stream]$Stream)
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hasher.ComputeHash($Stream))).Replace('-', '').ToLowerInvariant() }
    finally { $hasher.Dispose() }
}

function Get-UpdCandidateFileHash {
    param([Parameter(Mandatory)][string]$Path)
    $stream = [IO.File]::OpenRead($Path)
    try { return Get-UpdCandidateStreamHash -Stream $stream }
    finally { $stream.Dispose() }
}

function Get-UpdCandidateBlobHash {
    param([Parameter(Mandatory)][string]$Repository, [Parameter(Mandatory)][string]$Blob)
    if ($Blob -cnotmatch '^[0-9a-f]{40}$' -or $Repository -match '["\r\n]') { throw 'Invalid candidate blob verification argument.' }
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = 'git'
    $values = @('--no-replace-objects', '-C', $Repository, 'cat-file', 'blob', $Blob)
    $start.Arguments = ($values | ForEach-Object { '"' + $_.TrimEnd('\') + '"' }) -join ' '
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    try {
        [void]$process.Start()
        $errors = $process.StandardError.ReadToEndAsync()
        $hash = Get-UpdCandidateStreamHash -Stream $process.StandardOutput.BaseStream
        $process.WaitForExit()
        if ($process.ExitCode -ne 0 -or $errors.Result) { throw 'Candidate source blob verification failed.' }
        return $hash
    }
    finally { $process.Dispose() }
}

function Test-UpdReleaseCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Repository,
        [Parameter(Mandatory)][string]$SourceCommit,
        [Parameter(Mandatory)][string]$CandidateDirectory
    )
    $source = Get-UpdCandidateSource -Repository $Repository -SourceCommit $SourceCommit
    $candidate = [IO.Path]::GetFullPath($CandidateDirectory).TrimEnd('\', '/')
    Assert-UpdCandidateSafePath -Path $candidate
    # upload-artifact interprets paths as patterns even when they name files.
    if ($candidate -match '[\r\n\x00*?\[\]]') { throw 'Candidate upload paths cannot contain line breaks or glob characters.' }
    if (@($candidate.Replace('\', '/').Split('/') | Where-Object { $_.StartsWith('.', [StringComparison]::Ordinal) }).Count) {
        throw 'Candidate upload paths cannot contain hidden directory components.'
    }
    if (-not [IO.Directory]::Exists($candidate)) { throw 'Candidate output directory is missing.' }
    $zipName = $source.Config.packageName + '-' + $source.Config.version + '.zip'
    $expected = [string[]]@($zipName, 'manifest.json', 'SHA256SUMS')
    [Array]::Sort($expected, [StringComparer]::Ordinal)
    $items = @(Get-ChildItem -LiteralPath $candidate -Force -ErrorAction Stop)
    $observed = [string[]]@($items.Name)
    [Array]::Sort($observed, [StringComparer]::Ordinal)
    if ($items.Count -ne 3 -or ($observed -join "`n") -cne ($expected -join "`n") -or
        @($items | Where-Object { $_.PSIsContainer -or ($_.Attributes -band ([IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::Hidden)) -ne 0 }).Count) {
        throw 'Candidate output must contain exactly the three regular allowlisted files.'
    }
    $zipPath = Join-Path $candidate $zipName
    $manifestPath = Join-Path $candidate 'manifest.json'
    $checksumPath = Join-Path $candidate 'SHA256SUMS'
    $expectedChecksums = (Get-UpdCandidateFileHash -Path $zipPath) + '  ' + $zipName + "`n" +
        (Get-UpdCandidateFileHash -Path $manifestPath) + "  manifest.json`n"
    if ([IO.File]::ReadAllText($checksumPath) -cne $expectedChecksums) { throw 'Candidate checksum list does not match both exact output files.' }
    $manifest = [IO.File]::ReadAllText($manifestPath) | ConvertFrom-Json
    $manifestKeys = [string[]]@($manifest.PSObject.Properties.Name)
    [Array]::Sort($manifestKeys, [StringComparer]::Ordinal)
    if (($manifestKeys -join ',') -cne 'files,packageName,releaseStatus,schemaVersion,sourceCommit,sourceCommitUrl,sourceRepository,sourceTree,version') {
        throw 'Candidate manifest contains unsupported metadata.'
    }
    foreach ($field in @('packageName', 'version', 'releaseStatus', 'sourceRepository', 'sourceCommit', 'sourceTree', 'sourceCommitUrl')) {
        if ($manifest.$field -isnot [string]) { throw 'Candidate manifest identifiers must be scalar strings.' }
    }
    if ($manifest.schemaVersion -ne 1 -or $manifest.packageName -cne $source.Config.packageName -or
        $manifest.version -cne $source.Config.version -or $manifest.releaseStatus -cne 'UNRELEASED_CANDIDATE' -or
        $manifest.sourceRepository -cne $source.Config.sourceRepository -or $manifest.sourceCommit -cne $source.Commit -or
        $manifest.sourceTree -cne $source.Tree -or $manifest.sourceCommitUrl -cne ($source.Config.sourceRepository + '/commit/' + $source.Commit)) {
        throw 'Candidate manifest identifiers do not match the committed tested source.'
    }
    if (@($manifest.files).Count -ne 36 -or ($manifest.files.path -join "`n") -cne ($source.Config.files -join "`n")) {
        throw 'Candidate manifest inventory does not match the committed allowlist.'
    }
    foreach ($record in $manifest.files) {
        $recordKeys = [string[]]@($record.PSObject.Properties.Name)
        [Array]::Sort($recordKeys, [StringComparer]::Ordinal)
        if (($recordKeys -join ',') -cne 'path,sha256,size') { throw 'Candidate manifest file contains unsupported metadata.' }
        if ($record.path -isnot [string] -or $record.sha256 -isnot [string] -or $record.sha256 -cnotmatch '^[0-9a-f]{64}$' -or $null -eq $record.size -or
            $record.size -isnot [ValueType] -or $record.size -lt 0 -or [decimal]$record.size -ne [Math]::Truncate([decimal]$record.size)) {
            throw 'Candidate manifest contains an invalid file hash or length.'
        }
        $blob = $source.Blobs[$record.path]
        $blobSize = Invoke-UpdCandidateGit -Repository $source.Repository -Arguments @('cat-file', '-s', $blob)
        if ($record.sha256 -cne (Get-UpdCandidateBlobHash -Repository $source.Repository -Blob $blob) -or
            $blobSize -cne ([string]$record.size)) { throw 'Candidate manifest payload does not match the committed Git blobs.' }
    }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $expectedEntries = [string[]]@($source.Config.files + 'manifest.json')
        $observedEntries = [string[]]@($archive.Entries.FullName)
        [Array]::Sort($expectedEntries, [StringComparer]::Ordinal)
        [Array]::Sort($observedEntries, [StringComparer]::Ordinal)
        if ($archive.Entries.Count -ne 37 -or ($observedEntries -join "`n") -cne ($expectedEntries -join "`n")) {
            throw 'Candidate ZIP inventory contains missing, duplicate or unlisted entries.'
        }
        foreach ($entry in $archive.Entries) {
            $stream = $entry.Open()
            try { $hash = Get-UpdCandidateStreamHash -Stream $stream } finally { $stream.Dispose() }
            if ($entry.FullName -ceq 'manifest.json') {
                if ($hash -cne (Get-UpdCandidateFileHash -Path $manifestPath)) { throw 'Candidate ZIP and sidecar manifests differ.' }
            }
            else {
                $record = @($manifest.files | Where-Object { $_.path -ceq $entry.FullName })[0]
                if ($hash -cne $record.sha256 -or $entry.Length -ne $record.size) { throw 'Candidate ZIP payload differs from its manifest.' }
            }
        }
    }
    finally { $archive.Dispose() }
    return [pscustomobject][ordered]@{ version = $source.Config.version; sourceCommit = $source.Commit; sourceTree = $source.Tree;
        outputDirectory = $candidate; zipPath = $zipPath; manifestPath = $manifestPath; checksumPath = $checksumPath }
}

function Write-UpdCandidateGitHubOutput {
    param([Parameter(Mandatory)][object]$Candidate, [Parameter(Mandatory)][string]$Path)
    $lines = @()
    foreach ($name in @('zipPath', 'manifestPath', 'checksumPath')) {
        $value = [string]$Candidate.$name
        if (-not $value -or $value -match '[\r\n\x00*?\[\]]') { throw 'Unsafe candidate path cannot be exported to Actions.' }
        $lines += $name + '=' + $value
    }
    $outputPath = [IO.Path]::GetFullPath($Path)
    $candidatePath = [IO.Path]::GetFullPath($Candidate.outputDirectory).TrimEnd('\', '/')
    if ($outputPath.Equals($candidatePath, [StringComparison]::OrdinalIgnoreCase) -or
        $outputPath.StartsWith($candidatePath + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Actions output file cannot overlap the validated candidate directory.'
    }
    Assert-UpdCandidateSafePath -Path $outputPath
    if (-not [IO.File]::Exists($outputPath)) { throw 'Actions output file must already exist.' }
    [IO.File]::AppendAllText($outputPath, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))
}

if ($MyInvocation.InvocationName -eq '.') { return }

$ErrorActionPreference = 'Stop'
try {
    $repo = Split-Path $PSScriptRoot -Parent
    $source = Get-UpdCandidateSource -Repository $repo -SourceCommit $SourceCommit
    if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo 'artifacts/releases' }
    $engineName = if ($PSVersionTable.PSEdition -eq 'Desktop') { 'powershell.exe' } else { 'pwsh.exe' }
    $builderArguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
        (Join-Path $PSScriptRoot 'Build-Release.ps1'), '-SourceCommit', $SourceCommit, '-OutputDirectory', $OutputDirectory)
    $result = & (Join-Path $PSHOME $engineName) @builderArguments
    if ($LASTEXITCODE -ne 0) { throw 'Candidate builder failed; no upload outputs exported.' }
    $built = ($result -join "`n") | ConvertFrom-Json
    $expectedDirectory = [IO.Path]::GetFullPath((Join-Path $OutputDirectory ($source.Config.packageName + '-' + $source.Config.version + '-' + $SourceCommit)))
    if ($built.sourceCommit -cne $SourceCommit -or $built.sourceTree -cne $source.Tree -or
        $built.version -cne $source.Config.version -or $built.outputDirectory -cne $expectedDirectory) { throw 'Candidate builder result does not match its expected source and directory.' }
    $candidate = Test-UpdReleaseCandidate -Repository $repo -SourceCommit $SourceCommit -CandidateDirectory $expectedDirectory
    if ($GitHubOutput) { Write-UpdCandidateGitHubOutput -Candidate $candidate -Path $GitHubOutput }
    $candidate | ConvertTo-Json -Compress
    exit 0
}
catch {
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
    exit 1
}
