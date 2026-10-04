param([Parameter(Mandatory)][string]$ConfigPath)

# This worker and its reports are copied outside the checkout. The app is the
# actual extracted candidate; developer tools stay outside its isolated lookup.
$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
$ownedRoot = [IO.Path]::GetFullPath($config.Root).TrimEnd('\', '/')
$marker = Join-Path $ownedRoot '.upd-test-owner'
if (-not [IO.File]::Exists($marker) -or [IO.File]::ReadAllText($marker) -cne $config.Token) { throw 'Release worker requires its matching owned root.' }
foreach ($path in @($ConfigPath, $config.ProductScript, $config.OutputPath, $config.ResultPath)) {
    if (-not [IO.Path]::GetFullPath($path).StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Release worker paths must remain inside the owned root.' }
}
$feed = [uri]$config.BaseUrl
if ($feed.Scheme -ne 'http' -or $feed.Host -ne '127.0.0.1') { throw 'Release worker requires an owned loopback source.' }
$env:LOCALAPPDATA = Join-Path $ownedRoot 'local'
$env:USERPROFILE = Join-Path $ownedRoot 'home'
$env:TEMP = Join-Path $ownedRoot 'temporary'
$env:TMP = $env:TEMP
$env:PATH = $PSHOME + ';' + (Join-Path $env:WINDIR 'System32') + ';' + $env:WINDIR
$env:PSModulePath = Join-Path $PSHOME 'Modules'
$packageRoot = [IO.Path]::GetDirectoryName($config.ProductScript)
Set-Location -LiteralPath $packageRoot
$report = [ordered]@{ Engine = $PSVersionTable.PSVersion.ToString(); Error = $null; RuntimeGitAvailable = $false;
    RuntimePythonAvailable = $false; RuntimePesterAvailable = $false; RuntimeAnalyzerAvailable = $false;
    HelpAuthored = $false; HelpExampleCount = 0; Help = $null; Preview = $null; Fatal = $null;
    HostStartup = $null; HostStartupTreeDelta = @();
    PreviewChangedTree = $false; PreviewTreeDelta = @(); PreviewFullTreeDelta = @();
    NativeHostCachePath = (Join-Path $env:USERPROFILE 'AppData/Local/Microsoft/Windows/PowerShell/StartupProfileData-NonInteractive');
    NativeHostCacheChanged = $false; NativeHostCacheBefore = $null; NativeHostCacheAfter = $null; PackageChanged = $false }
$workerExit = 1

function Get-UpdReleaseWorkerTree {
    param([string]$Root)
    return @(Get-ChildItem -LiteralPath $Root -Force -Recurse | Sort-Object FullName | ForEach-Object {
        if ($_.PSIsContainer) { 'directory:' + $_.FullName }
        else { $_.FullName + ':' + $_.Length + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
    })
}

function Invoke-UpdReleaseLauncherChild {
    param([string]$Arguments, [switch]$HostOnly)
    $start = [Diagnostics.ProcessStartInfo]::new()
    if ($HostOnly) {
        $start.FileName = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
        $start.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "exit 0"'
    }
    else {
        $start.FileName = Join-Path $env:WINDIR 'System32/cmd.exe'
        $start.Arguments = '/d /s /v:off /c "".\UniversalPodcastDownloader.bat" ' + $Arguments + '"'
    }
    $start.WorkingDirectory = $packageRoot
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.EnvironmentVariables['PSModulePath'] = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules'
    $start.EnvironmentVariables['ERRORLEVEL'] = '77'
    $child = [Diagnostics.Process]::new()
    $child.StartInfo = $start
    try {
        if (-not $child.Start()) { throw 'Could not start the owned extracted batch child.' }
        $stdout = $child.StandardOutput.ReadToEndAsync()
        $stderr = $child.StandardError.ReadToEndAsync()
        $child.StandardInput.Close()
        if (-not $child.WaitForExit(15000)) { throw 'Owned extracted batch child exceeded its bounded wait.' }
        return [pscustomobject]@{ ExitCode = $child.ExitCode; Stdout = $stdout.Result; Stderr = $stderr.Result }
    }
    finally {
        if ($child.Id -and -not $child.HasExited) { $child.Kill(); $child.WaitForExit() }
        $child.Dispose()
    }
}

try {
    $packageBefore = Get-UpdReleaseWorkerTree -Root $packageRoot
    $report.RuntimeGitAvailable = $null -ne (Get-Command git -CommandType Application -ErrorAction SilentlyContinue)
    $report.RuntimePythonAvailable = $null -ne (Get-Command python -CommandType Application -ErrorAction SilentlyContinue)
    $report.RuntimePesterAvailable = @(Get-Module -ListAvailable Pester).Count -gt 0
    $report.RuntimeAnalyzerAvailable = @(Get-Module -ListAvailable PSScriptAnalyzer).Count -gt 0
    . $config.ProductScript
    $help = Get-Help '.\UniversalPodcastDownloader.ps1' -Full
    $report.Help = $help | Out-String -Width 160
    $report.HelpAuthored = -not [string]::IsNullOrWhiteSpace(($help.Description.Text -join ' '))
    $report.HelpExampleCount = @($help.Examples.Example | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Code) }).Count
    # Native PowerShell prepares profile/temp directories and its CLR startup
    # cache independently of the downloader. Prepare owned parents, then record
    # an exit-only native host with the exact same environment before baseline.
    foreach ($directory in @($env:USERPROFILE, (Join-Path $env:USERPROFILE 'AppData/Local'),
        (Join-Path $env:USERPROFILE 'AppData/Roaming'), $env:TEMP, $env:LOCALAPPDATA,
        [IO.Path]::GetDirectoryName($report.NativeHostCachePath))) {
        $null = [IO.Directory]::CreateDirectory($directory)
    }
    $coldBefore = Get-UpdReleaseWorkerTree -Root $ownedRoot
    $report.HostStartup = Invoke-UpdReleaseLauncherChild -HostOnly
    $report.HostStartupTreeDelta = @(Compare-Object $coldBefore (Get-UpdReleaseWorkerTree -Root $ownedRoot))
    $before = Get-UpdReleaseWorkerTree -Root $ownedRoot
    if ([IO.File]::Exists($report.NativeHostCachePath)) {
        $report.NativeHostCacheBefore = @{ Size = ([IO.FileInfo]$report.NativeHostCachePath).Length;
            Sha256 = (Get-FileHash -LiteralPath $report.NativeHostCachePath -Algorithm SHA256).Hash }
    }
    $report.Preview = Invoke-UpdReleaseLauncherChild -Arguments ('-FeedUrl "' + $config.BaseUrl + '/feeds/single.xml" -OutputPath "' + $config.OutputPath + '" -Mode All -NonInteractive -WhatIf')
    if ([IO.File]::Exists($report.NativeHostCachePath)) {
        $report.NativeHostCacheAfter = @{ Size = ([IO.FileInfo]$report.NativeHostCachePath).Length;
            Sha256 = (Get-FileHash -LiteralPath $report.NativeHostCachePath -Algorithm SHA256).Hash }
    }
    $report.PreviewFullTreeDelta = @(Compare-Object $before (Get-UpdReleaseWorkerTree -Root $ownedRoot))
    # The host's named startup optimization cache may change on every launch.
    # Retain its exact observation; compare every other owned directory/file,
    # including all downloader LOCALAPPDATA, archive, settings, logs and TEMP.
    $report.PreviewTreeDelta = @($report.PreviewFullTreeDelta | Where-Object {
        -not $_.InputObject.StartsWith($report.NativeHostCachePath + ':', [StringComparison]::OrdinalIgnoreCase)
    })
    $report.NativeHostCacheChanged = $report.PreviewFullTreeDelta.Count -ne $report.PreviewTreeDelta.Count
    $report.PreviewChangedTree = $report.PreviewTreeDelta.Count -gt 0
    # A genuine product planning failure reaches CMD through the shipped launcher.
    $report.Fatal = Invoke-UpdReleaseLauncherChild -Arguments ('-FeedUrl "' + $config.BaseUrl + '/feeds/single.xml" -OutputPath "' + $config.OutputPath + '" -Mode All -CustomCount 1 -NonInteractive')
    $report.PackageChanged = @(Compare-Object $packageBefore (Get-UpdReleaseWorkerTree -Root $packageRoot)).Count -gt 0
    $workerExit = 0
}
catch { $report.Error = $_.Exception.Message }
[IO.File]::WriteAllText($config.ResultPath, ($report | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
exit $workerExit
