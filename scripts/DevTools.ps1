#requires -Version 5.1
# Shared development helpers; loading this file has no side effects.
function Assert-UpdPlainPath {
    param([Parameter(Mandatory)][string]$Path)
    $ancestor = [IO.Path]::GetFullPath($Path)
    while ($ancestor) {
        if ((Test-Path -LiteralPath $ancestor) -and
            ((Get-Item -LiteralPath $ancestor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Development cache must not use reparse points.'
        }
        $ancestor = [IO.Path]::GetDirectoryName($ancestor)
    }
}

function Get-UpdDiagnosticContext {
    param([string[]]$Lines, [int]$Line)
    if ($Line -lt 1) { return '<file>' }
    $first = [Math]::Max(0, $Line - 2)
    $last = [Math]::Min($Lines.Count - 1, $Line)
    return (($Lines[$first..$last] | ForEach-Object { $_.Trim() }) -join "`n")
}

function Get-UpdDevToolLock {
    $path = Join-Path (Split-Path $PSScriptRoot -Parent) 'tools/devtools.json'
    $lock = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($lock.schema_version -ne 1) { throw 'Unsupported development tool lock.' }
    return $lock
}

function Import-UpdDevModule {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$ToolsPath
    )
    $entry = @((Get-UpdDevToolLock).modules | Where-Object { $_.name -eq $Name })
    if ($entry.Count -ne 1) { throw "Missing lock entry for $Name." }
    $manifest = Join-Path $ToolsPath (Join-Path $Name (Join-Path $entry[0].version "$Name.psd1"))
    if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        throw "Missing pinned $Name $($entry[0].version). Run scripts/Install-DevTools.ps1 first."
    }
    Import-Module -Name $manifest -RequiredVersion $entry[0].version -Force -Global -ErrorAction Stop
    $module = Get-Module -Name $Name | Where-Object { $_.Version -eq [version]$entry[0].version } | Select-Object -First 1
    if (-not $module) { throw "Could not load pinned $Name." }
    Write-Host ("{0} {1}; PowerShell {2} ({3}); OS {4}" -f $Name, $module.Version, $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, [Environment]::OSVersion.Version)
}
