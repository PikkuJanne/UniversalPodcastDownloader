param([Parameter(Mandatory)][string]$ConfigPath)

$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
$result = [ordered]@{
    Succeeded = $false; ErrorMessage = $null; FailureKind = $null
    Bytes = 0; Content = $null; ElapsedSeconds = 0
}
$settings = @{}
foreach ($property in $config.Settings.PSObject.Properties) { $settings[$property.Name] = $property.Value }
$watch = [Diagnostics.Stopwatch]::StartNew()
try {
    if ($config.Kind -eq 'Metadata') {
        . $config.ProductScript
        $policy = New-PodcastTransportPolicy @settings
        $response = Invoke-PodcastMetadataRequest -Uri $config.Url -Policy $policy
        $result.Bytes = $response.Bytes
        $result.Content = $response.Content
    }
    elseif ($config.Kind -eq 'Media') {
        & $config.ProductScript -Mode All -FeedUrl $config.Url -OutputPath $config.OutputPath @settings
    }
    else { throw 'Unknown transport test action.' }
    $result.Succeeded = $true
}
catch {
    $result.ErrorMessage = $_.Exception.Message
    $exception = $_.Exception
    while ($null -ne $exception) {
        if ($exception.Data.Contains('Kind')) { $result.FailureKind = [string]$exception.Data['Kind'] }
        $exception = $exception.InnerException
    }
}
finally {
    $result.ElapsedSeconds = $watch.Elapsed.TotalSeconds
    [IO.File]::WriteAllText($config.ResultPath, ($result | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
}
if ($result.Succeeded) { exit 0 }
exit 1
