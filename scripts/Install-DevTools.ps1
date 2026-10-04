#requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DevTools.ps1')
$repo = Split-Path $PSScriptRoot -Parent
$cache = Join-Path $repo '.dev-tools'
$packages = Join-Path $cache 'packages'
$modules = Join-Path $cache 'Modules'

# Explicit developer setup only. Never called by the downloader or test runners.
# Refuse redirected ancestors before any write into the repository-local cache.
Assert-UpdPlainPath -Path $packages
Assert-UpdPlainPath -Path $modules
[void][IO.Directory]::CreateDirectory($packages)
[void][IO.Directory]::CreateDirectory($modules)
Add-Type -AssemblyName System.IO.Compression.FileSystem

foreach ($entry in (Get-UpdDevToolLock).modules) {
    $destination = Join-Path $modules (Join-Path $entry.name $entry.version)
    $manifest = Join-Path $destination ($entry.name + '.psd1')
    Assert-UpdPlainPath -Path $manifest
    if (Test-Path -LiteralPath $destination) {
        if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
            throw "Incomplete existing $($entry.name) directory; inspect it before retrying."
        }
        $metadata = Import-PowerShellDataFile -LiteralPath $manifest
        if ([version]$metadata.ModuleVersion -ne [version]$entry.version) { throw 'Installed module version differs from lock.' }
        Write-Host "$($entry.name) $($entry.version) already installed locally."
        continue
    }
    $package = Join-Path $packages ("{0}.{1}.nupkg" -f $entry.name, $entry.version)
    Assert-UpdPlainPath -Path $package
    if (-not (Test-Path -LiteralPath $package -PathType Leaf)) {
        $download = Join-Path $packages ([guid]::NewGuid().ToString('N') + '.download')
        Invoke-WebRequest -Uri $entry.uri -OutFile $download -UseBasicParsing -TimeoutSec 60
        if ((Get-FileHash -LiteralPath $download -Algorithm SHA256).Hash -ne $entry.sha256) {
            throw 'Downloaded package hash differs from the reviewed lock; file retained for inspection.'
        }
        [IO.File]::Move($download, $package)
    }
    if ((Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash -ne $entry.sha256) {
        throw 'Cached package hash differs from the reviewed lock.'
    }
    $staging = Join-Path $cache ('extract-' + [guid]::NewGuid().ToString('N'))
    $archive = [IO.Compression.ZipFile]::OpenRead($package)
    try {
        foreach ($item in $archive.Entries) {
            $relative = $item.FullName.Replace('\', '/')
            if ($relative.StartsWith('/') -or $relative.Contains(':') -or ($relative.Split('/') -contains '..')) {
                throw 'Unsafe path in development package.'
            }
        }
        [IO.Compression.ZipFileExtensions]::ExtractToDirectory($archive, $staging)
    }
    finally { $archive.Dispose() }
    $metadata = Import-PowerShellDataFile -LiteralPath (Join-Path $staging ($entry.name + '.psd1'))
    if ([version]$metadata.ModuleVersion -ne [version]$entry.version) { throw 'Package version differs from lock.' }
    [void][IO.Directory]::CreateDirectory((Split-Path $destination -Parent))
    [IO.Directory]::Move($staging, $destination)
    Write-Host "Installed $($entry.name) $($entry.version) in .dev-tools/Modules; SHA-256 verified."
}
