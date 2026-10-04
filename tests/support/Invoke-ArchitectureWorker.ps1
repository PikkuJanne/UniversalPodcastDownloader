param([Parameter(Mandatory)][string]$ConfigPath)

# This child imports only a copied runtime. All observations remain in its
# marked synthetic root, separate from the application success stream.
$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
$ownedRoot = [IO.Path]::GetFullPath($config.Root).TrimEnd('\', '/')
if (-not [IO.File]::Exists((Join-Path $ownedRoot '.upd-test-owner'))) { throw 'Architecture worker requires an owned test root.' }
foreach ($path in @($ConfigPath, $config.ProductScript, $config.OutputPath, $config.ResultPath)) {
    if (-not [IO.Path]::GetFullPath($path).StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Architecture observation paths must remain inside the owned root.'
    }
}
$feed = [uri]$config.BaseUrl
if ($feed.Scheme -ne 'http' -or $feed.Host -ne '127.0.0.1') { throw 'Architecture worker requires an owned loopback source.' }
$env:LOCALAPPDATA = Join-Path $ownedRoot 'local'
$env:USERPROFILE = Join-Path $ownedRoot 'home'
$env:TEMP = Join-Path $ownedRoot 'temporary'
$env:TMP = $env:TEMP
$env:PATH = $PSHOME + ';' + (Join-Path $env:WINDIR 'System32') + ';' + $env:WINDIR
$env:PSModulePath = Join-Path $PSHOME 'Modules'
$report = [ordered]@{ HostSurvived = $false; ImportOutput = 0; ImportCallCount = 0; ImportChangedTree = $false;
    ImportCreatedPolicy = $false; ImportCreatedPageLimit = $false; ImportPowerType = $false; PreferencesPreserved = $false;
    Runs = @(); Requests = @(); Standalone = $null; CorrelationCounts = @(); HandlesClosed = $false;
    PreviewChangedTree = $false; PackageChanged = $false; RuntimePythonAvailable = $false;
    RunCreatedOptionState = $false; RunPreferencesPreserved = $false; Error = $null }
$workerExit = 1
$importCalls = @{ Count = 0 }
$breakpoint = $null
function Get-UpdArchitectureTree {
    param([string]$Root)
    return @((Get-ChildItem -LiteralPath $Root -Force -Recurse | Sort-Object FullName | ForEach-Object {
        if ($_.PSIsContainer) { 'directory:' + $_.FullName }
        else { $_.FullName + ':' + $_.Length + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
    }))
}
try {
    $before = Get-UpdArchitectureTree -Root $ownedRoot
    $packageRoot = [IO.Path]::GetDirectoryName($config.ProductScript)
    $packageBefore = Get-UpdArchitectureTree -Root $packageRoot
    $breakpoint = Set-PSBreakpoint -Command @('Read-Host', 'Add-Type', 'New-Item', 'Invoke-PodcastMetadataRequest',
        'Initialize-PodcastSavedShowProtection', 'Get-PodcastSavedShowConfigPath', 'Start-PodcastKeepAwake') -Action {
        $importCalls.Count++
        throw 'Import unexpectedly invoked a runtime side effect.'
    }
    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'
    $ConfirmPreference = 'High'
    $imported = @(. $config.ProductScript -MaxAttempts 2 -MaxFeedPages 1 -OutputPath $config.OutputPath)
    $report.ImportOutput = $imported.Count
    $report.ImportCallCount = $importCalls.Count
    $report.ImportCreatedPolicy = $null -ne (Get-Variable PodcastTransportPolicy -Scope Script -ErrorAction SilentlyContinue)
    $report.ImportCreatedPageLimit = $null -ne (Get-Variable PodcastMaxFeedPages -Scope Script -ErrorAction SilentlyContinue)
    $report.ImportPowerType = $null -ne ('UPD.KeepAwakeLease' -as [type])
    $report.PreferencesPreserved = $ErrorActionPreference -eq 'Continue' -and $ProgressPreference -eq 'SilentlyContinue' -and $ConfirmPreference -eq 'High'
    $report.ImportChangedTree = @(Compare-Object $before (Get-UpdArchitectureTree -Root $ownedRoot)).Count -gt 0
    Remove-PSBreakpoint -Breakpoint $breakpoint
    $breakpoint = $null
    $ErrorActionPreference = 'Stop'
    $report.RuntimePythonAvailable = $null -ne (Get-Command python -CommandType Application -ErrorAction SilentlyContinue)

    if ($config.Action -ne 'Import') {
        $originalMetadata = (Get-Command Invoke-PodcastMetadataRequest).ScriptBlock
        $observation = @{ Requests = [Collections.Generic.List[object]]::new() }
        function Invoke-PodcastMetadataRequest {
            [CmdletBinding()]
            param([string]$Uri, $Policy)
            $effective = if ($null -ne $Policy) { $Policy } else { New-PodcastTransportPolicy }
            $observation.Requests.Add([pscustomobject]@{ Path = ([uri]$Uri).AbsolutePath; MaxAttempts = $effective.MaxAttempts;
                HeaderTimeoutSeconds = $effective.HeaderTimeoutSeconds; IdleTimeoutSeconds = $effective.IdleTimeoutSeconds })
            & $originalMetadata @PSBoundParameters
        }
        $arguments = @{ FeedUrl = $config.BaseUrl + '/feeds/single.xml'; OutputPath = $config.OutputPath;
            Mode = 'All'; NonInteractive = $true; BaseDelaySeconds = 0; MaxDelaySeconds = 0 }
        if ($config.Action -eq 'Isolation') {
            $arguments.FeedUrl = $config.BaseUrl + '/feeds/pagination-rss.xml'
            foreach ($setting in @(
                @{ MaxFeedPages = 1; MaxAttempts = 1; HeaderTimeoutSeconds = 5; IdleTimeoutSeconds = 6 },
                @{ MaxFeedPages = 2; MaxAttempts = 2; HeaderTimeoutSeconds = 7; IdleTimeoutSeconds = 8 }
            )) {
                $output = @(Invoke-PodcastRun @arguments @setting -WhatIf)
                if ($output.Count -ne 1 -or $output[0].Type -cne 'Podcast.RunResult') { throw 'Callable run did not return one private result.' }
                $report.Runs += $output[0]
                $ids = Get-Variable PodcastDiagnosticRequestIds -Scope Script -ErrorAction SilentlyContinue
                $report.CorrelationCounts += $(if ($ids -and $null -ne $ids.Value) { $ids.Value.Count } else { 0 })
            }
            $standalone = Resolve-PodcastItems -Feeds @($arguments.FeedUrl)
            $report.Standalone = $standalone.Catalogue
        }
        else {
            foreach ($preview in @($false, $false, $true)) {
                if ($preview) { $previewBefore = Get-UpdArchitectureTree -Root $ownedRoot }
                $output = @(Invoke-PodcastCommand -Options ($arguments + @{ WhatIf = $preview; MaxAttempts = 1 }))
                if ($output.Count -ne 1 -or $output[0].Type -cne 'Podcast.RunResult') { throw 'Dispatcher did not return one private result.' }
                $report.Runs += $output[0]
                $ids = Get-Variable PodcastDiagnosticRequestIds -Scope Script -ErrorAction SilentlyContinue
                $report.CorrelationCounts += $(if ($ids -and $null -ne $ids.Value) { $ids.Value.Count } else { 0 })
                if ($preview) { $report.PreviewChangedTree = @(Compare-Object $previewBefore (Get-UpdArchitectureTree -Root $ownedRoot)).Count -gt 0 }
            }
            $report.HandlesClosed = $true
            foreach ($item in @(Get-ChildItem -LiteralPath $ownedRoot -Recurse -Force -File | Where-Object {
                $_.FullName.StartsWith($config.OutputPath + '\', [StringComparison]::OrdinalIgnoreCase) -or $_.Extension -eq '.log'
            })) {
                $handle = [IO.File]::Open($item.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
                $handle.Dispose()
            }
        }
        $report.Requests = @($observation.Requests.ToArray())
    }
    $report.PackageChanged = @(Compare-Object $packageBefore (Get-UpdArchitectureTree -Root $packageRoot)).Count -gt 0
    $report.RunCreatedOptionState = $null -ne (Get-Variable PodcastTransportPolicy -Scope Script -ErrorAction SilentlyContinue) -or
        $null -ne (Get-Variable PodcastMaxFeedPages -Scope Script -ErrorAction SilentlyContinue)
    $report.RunPreferencesPreserved = $ErrorActionPreference -eq 'Stop' -and $ProgressPreference -eq 'SilentlyContinue' -and $ConfirmPreference -eq 'High'
    $report.HostSurvived = $true
    $workerExit = 0
}
catch { $report.Error = $_.Exception.Message }
finally { if ($breakpoint) { Remove-PSBreakpoint -Breakpoint $breakpoint } }
[IO.File]::WriteAllText($config.ResultPath, ($report | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
exit $workerExit
