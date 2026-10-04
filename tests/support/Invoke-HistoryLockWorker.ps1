param(
    [Parameter(Mandatory)][string]$ProductScript,
    [Parameter(Mandatory)][string]$ArchiveRoot,
    [Parameter(Mandatory)][string]$ReadyPath,
    [ValidateSet('Show', 'Archive')][string]$Scope = 'Show'
)

# The parent tracks and terminates only this test process. A real OS handle,
# rather than a mocked lock, exercises exclusion and process-death release.
$ErrorActionPreference = 'Stop'
. $ProductScript
$heldStream = if ($Scope -eq 'Archive') { Enter-PodcastArchiveLock -Root $ArchiveRoot }
else { (Enter-PodcastHistoryLock -Root $ArchiveRoot).Stream }
try {
    [IO.File]::WriteAllText($ReadyPath, 'locked')
    Start-Sleep -Seconds 45
}
finally { $heldStream.Dispose() }
