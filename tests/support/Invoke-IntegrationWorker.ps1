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
    elseif ($config.TransactionHook -in @('BeforeStateReplaceCrash', 'AfterStateReplaceCrash')) {
        $historySource = Join-Path (Split-Path $config.ProductScript -Parent) 'src/HistoryStore.ps1'
        $sourceLines = [IO.File]::ReadAllLines($historySource)
        $replaceLines = @(for ($line = 0; $line -lt $sourceLines.Length; $line++) {
            if ($sourceLines[$line] -match '^\s*\[IO.File\]::Replace\(') { $line + 1 }
        })
        if ($replaceLines.Count -ne 1) { throw 'The integration history hook requires one explicit atomic state replacement.' }
        $hookLine = $replaceLines[0]
        if ($config.TransactionHook -eq 'AfterStateReplaceCrash') {
            $afterLines = @(for ($line = $hookLine; $line -lt $sourceLines.Length; $line++) {
                if ($sourceLines[$line] -match '^\s*\$ownedTemporary = \$false') { $line + 1 }
            })
            if ($afterLines.Count -ne 1) { throw 'The integration history hook requires the explicit post-replacement ownership update.' }
            $hookLine = $afterLines[0]
        }
        $null = Set-PSBreakpoint -Script $historySource -Line $hookLine -Action {
            [IO.File]::WriteAllText($config.HookMarkerPath, (@{ Hook = $config.TransactionHook } | ConvertTo-Json))
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
    if ($config.Action -in @('Discover', 'InteractivePreview')) {
        $discoveryPromptState = @{ Count = 0; SelectionIndex = 0 }
        function Read-Host {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'This bounded test worker supplies only expected URL and numbered-choice application UI responses; web cmdlet host prompts still fail under NonInteractive.')]
            [CmdletBinding()]
            param([string]$Prompt)

            $discoveryPromptState.Count++
            if ($discoveryPromptState.Count -eq 1 -and $Prompt -eq 'Paste RSS feed URL OR podcast page URL') {
                return $config.FeedUrl
            }
            if ($Prompt -eq 'Choose feed number (1-2)' -and $discoveryPromptState.SelectionIndex -lt @($config.Selection).Count) {
                $selection = @($config.Selection)[$discoveryPromptState.SelectionIndex]
                $discoveryPromptState.SelectionIndex++
                return $selection
            }
            throw 'Unexpected or repeated application prompt in discovery integration worker.'
        }
    }
    switch ($config.Action) {
        'Discover' {
            . $config.ProductScript
            $initial = Get-FeedUrlInteractive
            $resolved = Resolve-PodcastItems -Feeds @($initial.Url) -InitialResolution $initial
            $result.ResolvedUrl = $resolved.Url
            $result.ItemCount = @($resolved.Items).Count
            $result.Kind = $resolved.Kind
            $result.FinalUri = $resolved.FinalUri.AbsoluteUri
            $result.Candidates = @($resolved.Candidates)
            $result.PromptCount = $discoveryPromptState.Count
        }
        'Source' {
            . $config.ProductScript
            if ($config.ReuseResponse) {
                $response = Invoke-PodcastWebRequest -Uri $config.FeedUrl
                $source = Resolve-PodcastSource -Uri $config.FeedUrl -Response $response
            }
            else { $source = Resolve-PodcastSource -Uri $config.FeedUrl }
            $result.ResolvedUrl = $source.Url
            $result.Kind = $source.Kind
            $result.FinalUri = $source.FinalUri.AbsoluteUri
            $result.Candidates = @($source.Candidates)
            $result.ItemCount = @($source.Items).Count
            $result.ContentLength = $source.Content.Length
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
        'Preview' {
            & $config.ProductScript -Mode $config.Mode -FeedUrl $config.FeedUrl -OutputPath $config.OutputPath -WhatIf
        }
        'InteractivePreview' {
            & $config.ProductScript -Mode $config.Mode -OutputPath $config.OutputPath -WhatIf
            $result.PromptCount = $discoveryPromptState.Count
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
