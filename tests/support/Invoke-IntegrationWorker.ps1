param([Parameter(Mandatory)][string]$ConfigPath)

# Run under -NonInteractive in a bounded, owned child process. A legacy web
# confirmation becomes an observable failure instead of hanging a test runner.
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$result = [ordered]@{
    Succeeded = $false
    EngineVersion = $PSVersionTable.PSVersion.ToString()
    ErrorMessage = $null
    ErrorId = $null
}

try {
    if ($config.TransactionHook -eq 'DuringTransferCrash') {
        $requestSource = Join-Path (Split-Path $config.ProductScript -Parent) 'src/MediaRequest.ps1'
        $sourceLines = [IO.File]::ReadAllLines($requestSource)
        $writeLines = @(for ($line = 0; $line -lt $sourceLines.Length; $line++) {
            if ($sourceLines[$line] -match '^\s*\$DestinationStream\.Write\(') { $line + 2 }
        })
        if ($writeLines.Count -ne 1) { throw 'The integration interruption hook requires one explicit destination stream write.' }
        $null = Set-PSBreakpoint -Script $requestSource -Line $writeLines[0] -Action {
            $DestinationStream.Flush()
            $marker = [ordered]@{
                Hook = 'DuringTransferCrash'
                Temporary = $DestinationStream.Name
                BytesBeforeCrash = $DestinationStream.Length
            }
            [IO.File]::WriteAllText($config.HookMarkerPath, ($marker | ConvertTo-Json))
            [Diagnostics.Process]::GetCurrentProcess().Kill()
        }
    }
    elseif ($config.TransactionHook -and $config.TransactionHook -ne 'None' -or
        ($config.BoundaryJunctionPath -and $config.BoundaryStage -eq 'AfterTransfer')) {
        $transferSource = Join-Path (Split-Path $config.ProductScript -Parent) 'src/MediaTransfer.ps1'
        $sourceLines = [IO.File]::ReadAllLines($transferSource)
        $moveLines = @(for ($line = 0; $line -lt $sourceLines.Length; $line++) {
            if ($sourceLines[$line] -match '^\s*\[IO.File\]::Move\(\$temporary,\s*\$destination\)') { $line + 1 }
        })
        if ($moveLines.Count -ne 1) { throw 'The integration finalize hook requires one explicit no-overwrite move statement.' }
        $hookLine = $moveLines[0]
        if ($config.TransactionHook -eq 'AfterFinalizeCrash') { $hookLine++ }
        $transactionHookState = @{ Count = 0 }
        $boundaryInjectionState = @{ Count = 0 }
        $null = Set-PSBreakpoint -Script $transferSource -Line $hookLine -Action {
            if ($transactionHookState.Count -eq 0) {
                $transactionHookState.Count++
                $exclusive = $false
                if (Test-Path -LiteralPath $temporary) {
                    $probe = [IO.File]::Open($temporary, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
                    $probe.Dispose()
                    $exclusive = $true
                }
                $marker = [ordered]@{
                    Hook = $config.TransactionHook
                    Temporary = $temporary
                    Destination = $destination
                    TemporaryExclusiveOpen = $exclusive
                }
                [IO.File]::WriteAllText($config.HookMarkerPath, ($marker | ConvertTo-Json))
                if ($config.BoundaryJunctionPath) {
                    $null = New-Item -ItemType Junction -Path $config.BoundaryJunctionPath -Target $config.BoundaryJunctionTarget -ErrorAction Stop
                    $boundaryInjectionState.Count++
                }
                if ($config.TransactionHook -eq 'FinalRace') {
                    $raceBytes = [Text.Encoding]::UTF8.GetBytes('Synthetic concurrent final file. Preserve these exact bytes.')
                    $raceFile = [IO.File]::Open($destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
                    try { $raceFile.Write($raceBytes, 0, $raceBytes.Length) } finally { $raceFile.Dispose() }
                }
                if ($config.TransactionHook -like '*Crash') { [Diagnostics.Process]::GetCurrentProcess().Kill() }
            }
        }
    }
    switch ($config.Action) {
        'Discover' {
            . $config.ProductScript
            $script:discoveryPromptCount = 0
            function Read-Host {
                [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'This bounded test worker supplies only the two expected application UI responses; web cmdlet host prompts still fail under NonInteractive.')]
                [CmdletBinding()]
                param([string]$Prompt)

                $script:discoveryPromptCount++
                if ($script:discoveryPromptCount -eq 1 -and $Prompt -eq 'Paste RSS feed URL OR podcast page URL') {
                    return $config.FeedUrl
                }
                if ($script:discoveryPromptCount -eq 2 -and $Prompt -eq 'Use this feed? (Y/n)') {
                    return 'y'
                }
                throw 'Unexpected or repeated application prompt in discovery integration worker.'
            }
            $result.ResolvedUrl = Get-FeedUrlInteractive
            $result.PromptCount = $script:discoveryPromptCount
        }
        'Resolve' {
            . $config.ProductScript
            $resolved = Resolve-PodcastItems -Feeds @($config.FeedUrl)
            $result.ItemCount = $resolved.Items.Count
            $result.ResolvedUrl = $resolved.Url
            $episodes = @($resolved.Items | ForEach-Object { Get-EpisodeData $_ })
            $result.EpisodeTitles = @($episodes | ForEach-Object { $_.Title })
            $result.EpisodeUrls = @($episodes | ForEach-Object { $_.Url })
        }
        'Download' {
            if ($config.BoundaryJunctionPath -and $config.BoundaryStage -eq 'Preparing') {
                $boundaryInjectionState = @{ Count = 0 }
                function Write-Progress {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'This bounded integration hook inserts an owned junction after planning, exercising the real media write boundary with real loopback HTTP.')]
                    [CmdletBinding()]
                    param([string]$Activity, [string]$Status, [int]$PercentComplete, [string]$CurrentOperation, [switch]$Completed)

                    if ($boundaryInjectionState.Count -eq 0 -and $Status -like 'Preparing:*') {
                        $null = New-Item -ItemType Junction -Path $config.BoundaryJunctionPath -Target $config.BoundaryJunctionTarget -ErrorAction Stop
                        $boundaryInjectionState.Count++
                    }
                    Microsoft.PowerShell.Utility\Write-Progress @PSBoundParameters
                }
            }
            & $config.ProductScript -Mode $config.Mode -CustomCount $config.CustomCount -FeedUrl $config.FeedUrl -OutputPath $config.OutputPath -Verbose
        }
        default { throw 'Unknown integration worker action.' }
    }
    $result.Succeeded = $true
}
catch {
    $result.ErrorMessage = $_.Exception.Message
    $result.ErrorId = $_.FullyQualifiedErrorId
}
if ($config.BoundaryJunctionPath) { $result.BoundaryInjectionCount = $boundaryInjectionState.Count }
if ($config.Action -eq 'Download' -and (Test-Path -LiteralPath $config.OutputPath)) {
    # Count while the downloader process is still alive: a leaked exclusive
    # temporary handle would prevent the normal Windows cleanup from finishing.
    $result.RemainingTemporaryCount = @(Get-ChildItem -LiteralPath $config.OutputPath -Recurse -File -Filter '.upd-*.tmp' -Force).Count
}
$result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $config.ResultPath -Encoding UTF8
if ($result.Succeeded) { exit 0 }
exit 1
