#requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$Path,
    [string]$ToolsPath
)

$ErrorActionPreference = 'Stop'
try {
    $repo = Split-Path $PSScriptRoot -Parent
    if (-not $ToolsPath) { $ToolsPath = Join-Path $repo '.dev-tools/Modules' }
    . (Join-Path $PSScriptRoot 'DevTools.ps1')
    Import-UpdDevModule -Name PSScriptAnalyzer -ToolsPath $ToolsPath
    if (-not $Path) { $Path = @((Join-Path $repo 'UniversalPodcastDownloader.ps1'), (Join-Path $repo 'src'), $PSScriptRoot, (Join-Path $repo 'tests')) }
    $files = @($Path | ForEach-Object {
        $item = Get-Item -LiteralPath $_ -ErrorAction Stop
        if ($item.PSIsContainer) { Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Filter '*.ps1' }
        else { $item }
    } | Sort-Object FullName -Unique)
    if ($files.Count -eq 0) { throw 'No PowerShell files found to analyze.' }
    $baseline = Get-Content -LiteralPath (Join-Path $repo 'tools/lint-baseline.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $settings = Join-Path $repo 'tools/PSScriptAnalyzerSettings.psd1'
    $observed = @{}
    $newFindings = 0
    $knownFindings = 0
    $parseErrors = 0
    foreach ($file in $files) {
        $lines = [IO.File]::ReadAllLines($file.FullName)
        $tokens = $null
        $errors = $null
        $null = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
        $parseErrors += @($errors).Count
        foreach ($parseError in $errors) { Write-Warning "$($file.Name): $($parseError.Message)" }
        $relative = $file.FullName
        if ($relative.StartsWith($repo + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            $relative = $relative.Substring($repo.Length + 1).Replace('\', '/')
        }
        $findings = @(Invoke-ScriptAnalyzer -Path $file.FullName -Settings $settings -Severity Warning,Error -ErrorAction Stop)
        foreach ($finding in $findings) {
            $context = Get-UpdDiagnosticContext -Lines $lines -Line $finding.Line
            $key = $relative + '|' + $finding.RuleName + '|' + $finding.Message + '|' + $context
            if (-not $observed.ContainsKey($key)) { $observed[$key] = 0 }
            $observed[$key]++
            $allowed = @($baseline.findings | Where-Object {
                $_.path -eq $relative -and $_.rule -eq $finding.RuleName -and $_.message -eq $finding.Message -and $_.context -eq $context
            })
            if ($finding.Severity.ToString() -eq 'Warning' -and $allowed.Count -eq 1 -and $observed[$key] -le $allowed[0].count) {
                $knownFindings++
            }
            else {
                $newFindings++
                Write-Warning ("{0}:{1} {2}: {3}" -f $relative, $finding.Line, $finding.RuleName, $finding.Message)
            }
        }
    }
    $summary = "ANALYSIS engine=$($PSVersionTable.PSVersion) files=$($files.Count) parse_errors=$parseErrors new_findings=$newFindings known_baseline_warnings=$knownFindings"
    Write-Host $summary
    if ($env:GITHUB_STEP_SUMMARY) { Add-Content -LiteralPath $env:GITHUB_STEP_SUMMARY -Value $summary -Encoding UTF8 }
    if ($parseErrors -gt 0 -or $newFindings -gt 0) { exit 1 }
    exit 0
}
catch {
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
    exit 1
}
