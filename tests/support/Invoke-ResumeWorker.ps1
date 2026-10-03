param([Parameter(Mandatory)][string]$ConfigPath)

$ErrorActionPreference = 'Stop'
$config = [IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
. (Join-Path $PSScriptRoot 'WorkerRunProjection.ps1')
$workerExit = 1
$maxAttempts = if ($null -ne $config.PSObject.Properties['MaxAttempts']) { [int]$config.MaxAttempts } else { 1 }
$result = [ordered]@{ Succeeded = $false; ErrorMessage = $null }
try {
    if ($config.Interrupt) {
        $mediaSource = Join-Path (Split-Path $config.ProductScript -Parent) 'src/MediaRequest.ps1'
        $sourceLines = [IO.File]::ReadAllLines($mediaSource)
        $readLines = @(for ($index = 0; $index -lt $sourceLines.Length; $index++) {
            if ($sourceLines[$index] -match 'while.*Read-PodcastResponseChunk') { $index + 1 }
        })
        if ($readLines.Count -ne 1) { throw 'Resume interruption requires one explicit media response read loop.' }
        # Stop at the next read, after the previous write and checkpoint have
        # completed. This is a debugger hook in an owned worker, not app behavior.
        $null = Set-PSBreakpoint -Script $mediaSource -Line $readLines[0] -Action {
            if ($DestinationStream.Length -gt 0) {
                $DestinationStream.Flush($true)
                [IO.File]::WriteAllText($config.MarkerPath, (@{
                    PartialPath = $DestinationStream.Name
                    Bytes = $DestinationStream.Length
                } | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
                [Diagnostics.Process]::GetCurrentProcess().Kill()
            }
        }
    }
    $published = @(& $config.ProductScript -Mode All -FeedUrl $config.FeedUrl -OutputPath $config.OutputPath `
        -MaxAttempts $maxAttempts -HeaderTimeoutSeconds 3 -IdleTimeoutSeconds 5 -RetryBudgetSeconds 10 -PassThru)
    $projection = Get-UpdWorkerRunProjection -Output $published
    $result.Succeeded = $projection.Succeeded
    $result.ErrorMessage = $projection.ErrorMessage
    $result.RunResult = $projection.RunResult
    $workerExit = $projection.ExitCode
}
catch { $result.ErrorMessage = $_.Exception.Message }
[IO.File]::WriteAllText($config.ResultPath, ($result | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
exit $workerExit
