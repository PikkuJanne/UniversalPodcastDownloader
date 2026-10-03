# Development-only harness. Every child and temporary directory has one owner.
function New-UpdIntegrationContext {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test harness creates only a new GUID temporary directory; preview semantics would leave the test context unusable.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RepositoryRoot)

    $token = [guid]::NewGuid().ToString('N')
    $root = Join-Path ([IO.Path]::GetTempPath()) ('UPD-Integration-' + $token)
    $null = New-Item -ItemType Directory -Path $root -ErrorAction Stop
    [IO.File]::WriteAllText((Join-Path $root '.upd-test-owner'), $token)
    [pscustomobject]@{
        Root = [IO.Path]::GetFullPath($root)
        Token = $token
        RepositoryRoot = $RepositoryRoot
        Processes = New-Object 'System.Collections.Generic.List[object]'
        Junctions = New-Object 'System.Collections.Generic.List[string]'
        BaseUrl = $null
    }
}

function Start-UpdOwnedProcess {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test harness starts and tracks its own bounded child processes; tests require actual execution.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList
    )

    # These are executable/file paths or fixed harness arguments, never feed data.
    # Windows paths cannot contain double quotes. Reject them rather than allowing
    # an argument to escape the quoting used by .NET Framework's Arguments API.
    foreach ($argument in $ArgumentList) {
        if ($argument.Contains('"') -or $argument.EndsWith('\')) {
            throw 'Unsupported process argument in integration harness.'
        }
    }
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $FilePath
    $start.Arguments = ($ArgumentList | ForEach-Object { '"' + $_ + '"' }) -join ' '
    $start.WorkingDirectory = $Context.Root
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    if (-not $process.Start()) { throw "Could not start integration child: $FilePath" }
    $owned = [pscustomobject]@{
        Process = $process
        Output = $process.StandardOutput.ReadToEndAsync()
        ErrorOutput = $process.StandardError.ReadToEndAsync()
    }
    $Context.Processes.Add($owned)
    return $owned
}

function Start-UpdFixtureServer {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test harness starts only its owned loopback fixture server; tests require an actual listener.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][string]$PythonPath
    )

    $ready = Join-Path $Context.Root 'server-ready.json'
    $server = Start-UpdOwnedProcess -Context $Context -FilePath $PythonPath -ArgumentList @(
        '-B', (Join-Path $Context.RepositoryRoot 'tools/codex-handoff/fixture_server.py'),
        '--port', '0', '--ready-file', $ready
    )
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not (Test-Path -LiteralPath $ready)) {
        if ($server.Process.HasExited) {
            throw ('Fixture server exited before readiness: ' + $server.ErrorOutput.Result)
        }
        if ($watch.Elapsed.TotalSeconds -gt 15) { throw 'Fixture server readiness timed out after 15 seconds.' }
        Start-Sleep -Milliseconds 50
    }
    $payload = Get-Content -LiteralPath $ready -Raw | ConvertFrom-Json
    $uri = [uri]$payload.base_url
    if ($uri.Scheme -ne 'http' -or $uri.Host -ne '127.0.0.1' -or $uri.Port -lt 1) {
        throw 'Fixture server did not report its expected loopback address.'
    }
    $Context.BaseUrl = $uri.AbsoluteUri.TrimEnd('/')
}

function Invoke-UpdIntegrationWorker {
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][ValidateSet('Discover', 'Resolve', 'Source', 'Preview', 'InteractivePreview', 'Download')][string]$Action,
        [Parameter(Mandatory)][string]$FeedPath,
        [ValidateSet('Latest', 'Custom', 'All')][string]$Mode = 'All',
        [int]$CustomCount = 1,
        [string]$OutputName = 'output',
        [string]$BoundaryJunctionPath,
        [string]$BoundaryJunctionTarget,
        [ValidateSet('Preparing', 'AfterTransfer')][string]$BoundaryStage = 'Preparing',
        [ValidateSet('None', 'BeforeFinalizeCrash', 'AfterFinalizeCrash', 'AfterPrepareBeforeResumeRetireCrash', 'FinalRace', 'BeforeStateReplaceCrash', 'AfterStateReplaceCrash')][string]$TransactionHook = 'None',
        [switch]$InterruptOnPartial,
        [string[]]$Selection = @('1'),
        [switch]$ReuseResponse,
        [ValidateSet('en-US', 'de-DE', 'fi-FI')][string]$Culture,
        [ValidateRange(1, 100)][int]$MaxFeedPages
    )

    $discoveryPaths = @('/show', '/show/not-feed', '/redirect/show', '/redirect/feed',
        '/discovery/redirect', '/discovery/final/show.html', '/discovery/base.html',
        '/discovery/single.html', '/discovery/nonfeed-link.html')
    if ($FeedPath -notmatch '^/feeds/[a-z0-9-]+\.xml$' -and $FeedPath -notin $discoveryPaths) { throw 'Only named local feed or show fixtures are allowed.' }
    if ($OutputName -notmatch '^[a-z0-9-]+(?:[\\/][a-z0-9-]+)*$') { throw 'OutputName must contain only simple relative test directory names.' }
    $identifier = [guid]::NewGuid().ToString('N')
    $resultPath = Join-Path $Context.Root ($identifier + '-result.json')
    $configPath = Join-Path $Context.Root ($identifier + '-config.json')
    $config = @{
        Action = $Action
        ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'
        FeedUrl = $Context.BaseUrl + $FeedPath
        OutputPath = Join-Path $Context.Root $OutputName
        ResultPath = $resultPath
        Mode = $Mode
        CustomCount = $CustomCount
        Selection = @($Selection)
        ReuseResponse = [bool]$ReuseResponse
        Culture = $Culture
        TransactionHook = $TransactionHook
        HookMarkerPath = Join-Path $Context.Root ($identifier + '-hook.json')
    }
    if ($PSBoundParameters.ContainsKey('MaxFeedPages')) { $config.MaxFeedPages = $MaxFeedPages }
    if ($InterruptOnPartial) { $config.TransactionHook = 'DuringTransferCrash' }
    if ($BoundaryJunctionPath -or $BoundaryJunctionTarget) {
        $prefix = [IO.Path]::GetFullPath($Context.Root).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
        foreach ($candidate in @($BoundaryJunctionPath, $BoundaryJunctionTarget)) {
            if (-not $candidate -or -not [IO.Path]::GetFullPath($candidate).StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Boundary injection must use two paths within the owned test root.'
            }
        }
        $config.BoundaryJunctionPath = $BoundaryJunctionPath
        $config.BoundaryJunctionTarget = $BoundaryJunctionTarget
        $config.BoundaryStage = $BoundaryStage
        $Context.Junctions.Add([IO.Path]::GetFullPath($BoundaryJunctionPath))
    }
    $config | ConvertTo-Json | Set-Content -LiteralPath $configPath -Encoding UTF8
    $engineName = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    $worker = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engineName) -ArgumentList @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-IntegrationWorker.ps1'),
        '-ConfigPath', $configPath
    )
    $interrupted = $false
    $observedPartial = $null
    if (-not $worker.Process.HasExited -and -not $worker.Process.WaitForExit(30000)) {
        $worker.Process.Kill()
        $worker.Process.WaitForExit()
        throw 'Downloader integration child timed out after 30 seconds; only this owned child was stopped.'
    }
    $stdout = $worker.Output.Result
    $stderr = $worker.ErrorOutput.Result
    $hookMarker = if (Test-Path -LiteralPath $config.HookMarkerPath) { Get-Content -LiteralPath $config.HookMarkerPath -Raw | ConvertFrom-Json } else { $null }
    if ($InterruptOnPartial -and $hookMarker -and $hookMarker.Hook -eq 'DuringTransferCrash') {
        $interrupted = $true
        $observedPartial = Get-Item -LiteralPath $hookMarker.Temporary
    }
    if (-not (Test-Path -LiteralPath $resultPath)) {
        if (-not $interrupted -and -not ($TransactionHook -like '*Crash' -and $hookMarker)) {
            throw "Integration child produced no result. Exit=$($worker.Process.ExitCode); stdout=$stdout; stderr=$stderr"
        }
    }
    $workerResult = if (Test-Path -LiteralPath $resultPath) { Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json } else { $null }
    [pscustomobject]@{
        ExitCode = $worker.Process.ExitCode
        Result = $workerResult
        Stdout = $stdout
        Stderr = $stderr
        OutputPath = $config.OutputPath
        Interrupted = $interrupted
        ObservedPartial = $observedPartial
        HookMarker = $hookMarker
    }
}

function Invoke-UpdCliProcess {
    param(
        [Parameter(Mandatory)]$Context,
        [string]$FeedPath = '/feeds/atom.xml',
        [ValidateSet('Latest', 'Custom', 'All')][string]$Mode,
        [int]$CustomCount,
        [string]$OutputName = 'cli-output',
        [switch]$NonInteractive,
        [switch]$WithoutFeed,
        [switch]$Preview
    )

    if ($FeedPath -notmatch '^/feeds/[a-z0-9-]+\.xml$' -and $FeedPath -ne '/show') { throw 'CLI process checks require a named loopback feed fixture.' }
    if ($OutputName -notmatch '^[a-z0-9-]+$') { throw 'CLI process output requires a simple owned relative directory name.' }
    $output = Join-Path $Context.Root $OutputName
    $arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
        (Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'), '-OutputPath', $output)
    if (-not $WithoutFeed) { $arguments += @('-FeedUrl', ($Context.BaseUrl + $FeedPath)) }
    if ($PSBoundParameters.ContainsKey('Mode')) { $arguments += @('-Mode', $Mode) }
    if ($PSBoundParameters.ContainsKey('CustomCount')) { $arguments += @('-CustomCount', [string]$CustomCount) }
    if ($NonInteractive) { $arguments += '-NonInteractive' }
    if ($Preview) { $arguments += '-WhatIf' }
    $engineName = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    $process = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engineName) -ArgumentList $arguments
    if (-not $process.Process.WaitForExit(30000)) {
        $process.Process.Kill()
        $process.Process.WaitForExit()
        throw 'CLI process check timed out; only its owned child was stopped.'
    }
    [pscustomobject]@{
        ExitCode = $process.Process.ExitCode
        Stdout = $process.Output.Result
        Stderr = $process.ErrorOutput.Result
        OutputPath = $output
    }
}

function Invoke-UpdCliResultWorker {
    param(
        [Parameter(Mandatory)]$Context,
        [string]$FeedPath = '/feeds/single.xml',
        [ValidateSet('Run', 'Cancel')][string]$Action = 'Run',
        [ValidateSet('Latest', 'Custom', 'All')][string]$Mode,
        [int]$CustomCount,
        [ValidateRange(1, 2)][int]$Repeat = 1,
        [switch]$Preview,
        [switch]$WithoutFeed
    )
    if ($FeedPath -notmatch '^/feeds/[a-z0-9-]+\.xml$' -and $FeedPath -notmatch '^/resume/feed/[a-z0-9-]+$' -and $FeedPath -ne '/show') { throw 'Callable checks require a named loopback fixture.' }
    $identifier = [guid]::NewGuid().ToString('N')
    $resultPath = Join-Path $Context.Root ($identifier + '-result.json')
    $configPath = Join-Path $Context.Root ($identifier + '-config.json')
    $config = @{ Root = $Context.Root; ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'
        FeedUrl = $Context.BaseUrl + $FeedPath; OutputPath = Join-Path $Context.Root 'callable-output'
        ResultPath = $resultPath; Action = $Action; Repeat = $Repeat; Preview = [bool]$Preview; WithoutFeed = [bool]$WithoutFeed }
    foreach ($key in @('Mode', 'CustomCount')) { if ($PSBoundParameters.ContainsKey($key)) { $config[$key] = $PSBoundParameters[$key] } }
    [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $engineName = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    $process = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engineName) -ArgumentList @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
        (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-CliResultWorker.ps1'), '-ConfigPath', $configPath)
    if (-not $process.Process.WaitForExit(30000)) {
        $process.Process.Kill(); $process.Process.WaitForExit()
        throw 'Callable process check timed out; only its owned child was stopped.'
    }
    if (-not [IO.File]::Exists($resultPath)) { throw ('Callable worker did not survive: ' + $process.ErrorOutput.Result) }
    [pscustomobject]@{ ExitCode = $process.Process.ExitCode; Result = [IO.File]::ReadAllText($resultPath) | ConvertFrom-Json
        Stdout = $process.Output.Result; Stderr = $process.ErrorOutput.Result; OutputPath = $config.OutputPath }
}

function Get-UpdFixtureState {
    param([Parameter(Mandatory)]$Context)
    # BasicParsing here belongs to harness control traffic, not the downloader.
    $response = Invoke-WebRequest -Uri ($Context.BaseUrl + '/__stats') -UseBasicParsing -TimeoutSec 10
    return ($response.Content | ConvertFrom-Json)
}

function New-UpdOwnedJunction {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test harness creates a tracked junction between two canonical paths within its marked temporary root.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Target
    )

    $prefix = [IO.Path]::GetFullPath($Context.Root).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    foreach ($candidate in @($Path, $Target)) {
        if (-not [IO.Path]::GetFullPath($candidate).StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Test junction paths must remain inside the owned temporary root.'
        }
    }
    $null = New-Item -ItemType Junction -Path $Path -Target $Target -ErrorAction Stop
    $Context.Junctions.Add([IO.Path]::GetFullPath($Path))
}

function Remove-UpdIntegrationContext {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Mandatory test cleanup verifies canonical containment and a matching ownership marker, then stops only tracked child processes.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    foreach ($owned in $Context.Processes) {
        if (-not $owned.Process.HasExited) {
            $owned.Process.Kill()
            $owned.Process.WaitForExit()
        }
        $owned.Process.Dispose()
    }
    $root = [IO.Path]::GetFullPath($Context.Root)
    $temp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
    if ([IO.Path]::GetDirectoryName($root) -ne $temp -or
        [IO.Path]::GetFileName($root) -ne ('UPD-Integration-' + $Context.Token)) {
        throw "Refusing cleanup outside the owned temporary test root: $root"
    }
    $marker = Join-Path $root '.upd-test-owner'
    if (-not (Test-Path -LiteralPath $marker) -or [IO.File]::ReadAllText($marker) -ne $Context.Token) {
        throw "Refusing cleanup without the matching integration ownership marker: $root"
    }
    # Remove only explicitly tracked junction entries, without following their
    # targets. This happens before the recursive inventory/cleanup safety check.
    foreach ($junction in $Context.Junctions) {
        $canonical = [IO.Path]::GetFullPath($junction)
        if (-not $canonical.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing junction cleanup outside the owned test root: $canonical"
        }
        if (Test-Path -LiteralPath $canonical) {
            $entry = Get-Item -LiteralPath $canonical -Force
            if (-not ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                # A planned injection may not run if the downloader exits early.
                # Leave ordinary entries to the existing owned-root cleanup.
                continue
            }
            [IO.Directory]::Delete($canonical, $false)
        }
    }
    $entries = @(Get-Item -LiteralPath $root) + @(Get-ChildItem -LiteralPath $root -Force -Recurse)
    if ($entries | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }) {
        throw "Refusing cleanup of a test root containing reparse points: $root"
    }
    Remove-Item -LiteralPath $root -Recurse -Force
}
