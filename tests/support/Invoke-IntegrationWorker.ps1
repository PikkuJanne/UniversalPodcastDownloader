param([Parameter(Mandatory)][string]$ConfigPath)

# Run under -NonInteractive in a bounded, owned child process. A legacy web
# confirmation becomes an observable failure instead of hanging a test runner.
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$result = [ordered]@{
    Succeeded = $false
    EngineVersion = $PSVersionTable.PSVersion.ToString()
    BasicParsingDefault = [bool]$config.BasicParsing
    ErrorMessage = $null
    ErrorId = $null
}

if ($config.BasicParsing) {
    # Test-only compatibility injection. It is reported in every result and is
    # never evidence that the untouched Windows PowerShell entrypoint works.
    $PSDefaultParameterValues['Invoke-WebRequest:UseBasicParsing'] = $true
}

try {
    switch ($config.Action) {
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
            & $config.ProductScript -Mode All -FeedUrl $config.FeedUrl -OutputPath $config.OutputPath -Verbose
        }
        default { throw 'Unknown integration worker action.' }
    }
    $result.Succeeded = $true
}
catch {
    $result.ErrorMessage = $_.Exception.Message
    $result.ErrorId = $_.FullyQualifiedErrorId
}
$result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $config.ResultPath -Encoding UTF8
if ($result.Succeeded) { exit 0 }
exit 1
