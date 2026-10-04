# Development-only projection for callers of the public script boundary.
function Get-UpdWorkerRunProjection {
    param([Parameter(Mandatory)][object[]]$Output)
    if ($Output.Count -ne 1 -or $null -eq $Output[0].PSObject.Properties['ExitCode']) {
        throw 'The application must publish exactly one structured run result when PassThru is requested.'
    }
    $run = $Output[0]
    [pscustomobject]@{
        Succeeded = ([int]$run.ExitCode -eq 0)
        ExitCode = [int]$run.ExitCode
        ErrorMessage = if ([int]$run.ExitCode -eq 0) { $null } else { [string]$run.Message }
        RunResult = $run
    }
}
