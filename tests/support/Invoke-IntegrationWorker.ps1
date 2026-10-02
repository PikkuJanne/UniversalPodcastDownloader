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
            if ($config.BoundaryJunctionPath -and $config.BoundaryStage -eq 'AfterTransfer') {
                $boundaryInjectionState = @{ Count = 0 }
                function Invoke-WebRequest {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'This integration wrapper uses the real web cmdlet and inserts an owned junction only after the first media response has finished.')]
                    [CmdletBinding()]
                    param([string]$Uri, [string]$OutFile, [switch]$UseBasicParsing)

                    $response = Microsoft.PowerShell.Utility\Invoke-WebRequest @PSBoundParameters
                    if ($OutFile -and $boundaryInjectionState.Count -eq 0) {
                        $null = New-Item -ItemType Junction -Path $config.BoundaryJunctionPath -Target $config.BoundaryJunctionTarget -ErrorAction Stop
                        $boundaryInjectionState.Count++
                    }
                    return $response
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
$result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $config.ResultPath -Encoding UTF8
if ($result.Succeeded) { exit 0 }
exit 1
