#requires -Version 5.1

function New-PodcastConfigResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only a private in-memory command result.')]
    [CmdletBinding()]
    param([switch]$Preview, [Nullable[bool]]$Changed, [object[]]$Shows = @(),
        [ValidateSet(0, 1, 130)][int]$ExitCode = 0, [string]$Message)
    $rows = @($Shows | Where-Object { $null -ne $_ } | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Mode = $_.Mode; CustomCount = $_.CustomCount; FeedConfigured = [bool]$_.FeedConfigured }
    })
    return [pscustomobject]@{
        Type = 'Podcast.ConfigResult'; SchemaVersion = 1; Preview = [bool]$Preview
        Changed = $Changed; Complete = ($ExitCode -eq 0); ExitCode = $ExitCode
        Status = $(if ($ExitCode -eq 130) { 'cancelled' } elseif ($ExitCode -eq 1) { 'fatal' } elseif ($Preview) { 'preview' } else { 'success' })
        Shows = @($rows); Message = $Message
    }
}

function Get-PodcastSavedCommandErrorMessage {
    [CmdletBinding()]
    param($ErrorObject)
    $allowed = @(
        'Saved-show configuration is malformed or unsupported. Preserved the configuration.',
        'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.',
        'Saved-show configuration path is unsafe or inaccessible.',
        'Saved-show credentials are unavailable for the current Windows user.',
        'Saved-show configuration is in use by another operation.',
        'The requested saved show was not found.',
        'Saved-show name must be a token of 1 to 64 letters, digits, underscores or hyphens.',
        'Saved-show output must be an absolute supported Windows path.',
        'Saved-show configuration supports at most 100 shows.',
        'Saved-show export already exists or its destination is unavailable.',
        'Choose exactly one saved-show operation.',
        'Saved-show operations cannot be combined with direct feed or legacy arguments.',
        'The saved-show operation contains unsupported arguments.',
        'Saving a show requires an explicit nonblank FeedUrl.',
        'ShowName requires one name, or Batch for several names.',
        'ConfigPath requires a saved-show operation.',
        '-NonInteractive cannot be combined with -Confirm.',
        '-CustomCount must be a positive integer.',
        '-CustomCount conflicts with an explicit Latest or All mode.',
        "Mode 'Custom' requires -CustomCount with a value >= 1.",
        'Batch run options are invalid or unsupported.'
    )
    $cause = if ($ErrorObject -is [Management.Automation.ErrorRecord]) { $ErrorObject.Exception } else { $ErrorObject }
    for ($depth = 0; $cause -is [Exception] -and $depth -lt 16; $depth++) {
        if ($allowed -ccontains $cause.Message) { return $cause.Message }
        $cause = $cause.InnerException
    }
    return 'The saved-show operation failed. Check its local configuration and retry.'
}

function Invoke-PodcastCommand {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Options)
    $keys = @($Options.Keys)
    $selectors = @('SaveShow', 'ListShows', 'RemoveShow', 'ExportShows', 'ShowName', 'Batch')
    $savedRequested = @($keys | Where-Object {
        $_ -eq 'ConfigPath' -or ($_ -in $selectors -and ($_ -notin @('ListShows', 'Batch') -or [bool]$Options[$_]))
    }).Count -gt 0
    if (-not $savedRequested) {
        # Explicitly disabled new switches retain the ordinary command path.
        $ordinary = @{}
        foreach ($key in $keys) { if ($key -notin @('ListShows', 'Batch')) { $ordinary[$key] = $Options[$key] } }
        return Invoke-PodcastRun @ordinary
    }

    $namedRun = $false
    $preview = if ($keys -contains 'WhatIf') { [bool]$Options['WhatIf'] } else { [bool]$WhatIfPreference }
    try {
        $active = @($selectors | Where-Object {
            $keys -contains $_ -and ($_ -notin @('ListShows', 'Batch') -or [bool]$Options[$_])
        })
        if ($active.Count -eq 0) { throw 'ConfigPath requires a saved-show operation.' }
        $batchRequested = $active -contains 'Batch'
        $namedRun = $active -contains 'ShowName'
        if (($batchRequested -and @($active | Where-Object { $_ -notin @('Batch', 'ShowName') }).Count -gt 0) -or
            (-not $batchRequested -and $active.Count -ne 1)) { throw 'Choose exactly one saved-show operation.' }
        $operation = if ($batchRequested) { 'Batch' } else { $active[0] }
        if (@($keys | Where-Object { $_ -like 'Legacy*' }).Count -gt 0 -or
            ($operation -ne 'SaveShow' -and $keys -contains 'FeedUrl')) {
            throw 'Saved-show operations cannot be combined with direct feed or legacy arguments.'
        }
        $common = @('ConfigPath', 'NonInteractive', 'WhatIf', 'Confirm', 'Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction', 'ProgressAction')
        $runtime = @(Get-PodcastSavedRunOptionName)
        $allowed = $common + @($active)
        if ($operation -eq 'SaveShow') { $allowed += @('FeedUrl', 'Mode', 'CustomCount', 'OutputPath') }
        if ($operation -in @('ShowName', 'Batch')) { $allowed += $runtime }
        if (@($keys | Where-Object { $_ -notin $allowed -and $_ -notin @('ListShows', 'Batch') }).Count -gt 0) {
            throw 'The saved-show operation contains unsupported arguments.'
        }
        if ($Options['NonInteractive'] -and $Options['Confirm']) { throw '-NonInteractive cannot be combined with -Confirm.' }
        $configPath = if ($keys -contains 'ConfigPath') { [string]$Options['ConfigPath'] } else { Get-PodcastSavedShowConfigPath }
        $confirmation = @{}
        if ($keys -contains 'WhatIf') { $confirmation.WhatIf = [bool]$Options['WhatIf'] }
        if ($keys -contains 'Confirm') { $confirmation.Confirm = [bool]$Options['Confirm'] }
        if ($Options['NonInteractive']) { $confirmation.Confirm = $false }

        if ($operation -in @('ShowName', 'Batch')) {
            $overrides = @{}
            foreach ($key in $runtime) { if ($keys -contains $key) { $overrides[$key] = $Options[$key] } }
            if ($operation -eq 'Batch') {
                $batchArguments = @{ ConfigPath = $configPath; RunOptions = $overrides }
                if ($namedRun) { $batchArguments.ShowName = $Options['ShowName'] }
                return Invoke-PodcastBatchRun @batchArguments @confirmation
            }
            $names = @($Options['ShowName'])
            if ($names.Count -ne 1 -or $null -eq $names[0] -or $names[0] -isnot [string] -or [string]::IsNullOrWhiteSpace($names[0])) {
                throw 'ShowName requires one name, or Batch for several names.'
            }
            $saved = Get-PodcastSavedShow -ConfigPath $configPath -Name $names[0]
            $runArguments = Resolve-PodcastSavedRunOptions -RunOptions $overrides -SavedShow $saved
            if ($keys -contains 'NonInteractive') { $runArguments.NonInteractive = [bool]$Options['NonInteractive'] }
            return Invoke-PodcastRun @runArguments @confirmation
        }
        if ($operation -eq 'ListShows') {
            $rows = @(Get-PodcastSavedShowList -ConfigPath $configPath)
            # Helpers return one array object, including empty/singleton lists.
            if ($rows.Count -eq 1 -and $rows[0] -is [array]) { $rows = @($rows[0]) }
            Write-Host ('Saved shows: {0}' -f $rows.Count)
            foreach ($row in $rows) {
                $selection = if ($row.Mode -eq 'Custom') { 'Custom ({0})' -f $row.CustomCount } else { $row.Mode }
                Write-Host ('  {0}: {1}' -f $row.Name, $selection)
            }
            return New-PodcastConfigResult -Preview:$preview -Changed:$false -Shows $rows -Message 'Saved-show settings listed; feed addresses remain protected.'
        }
        if ($operation -eq 'SaveShow') {
            if ($keys -notcontains 'FeedUrl' -or [string]::IsNullOrWhiteSpace([string]$Options['FeedUrl'])) {
                throw 'Saving a show requires an explicit nonblank FeedUrl.'
            }
            $saveArguments = @{ ConfigPath = $configPath; Name = $Options['SaveShow']; FeedUrl = $Options['FeedUrl'] }
            foreach ($key in @('Mode', 'CustomCount', 'OutputPath')) { if ($keys -contains $key) { $saveArguments[$key] = $Options[$key] } }
            if ($keys -notcontains 'OutputPath') { $saveArguments.OutputPath = "$env:USERPROFILE\Downloads\Podcasts" }
            $changed = Save-PodcastSavedShow @saveArguments @confirmation
        }
        elseif ($operation -eq 'RemoveShow') {
            $changed = Remove-PodcastSavedShow -ConfigPath $configPath -Name $Options['RemoveShow'] @confirmation
        }
        else {
            $changed = Export-PodcastSavedShowConfig -ConfigPath $configPath -Path $Options['ExportShows'] @confirmation
        }
        Write-Host $changed.Message
        return New-PodcastConfigResult -Preview:$changed.Preview -Changed:$changed.Changed -Message $changed.Message
    }
    catch {
        $cancelled = Test-PodcastCancellation -ErrorObject $_
        $message = if ($cancelled) { 'Saved-show command cancelled; retained settings and completed media were preserved.' }
            else { Get-PodcastSavedCommandErrorMessage -ErrorObject $_ }
        $result = if ($Options['Batch']) { New-PodcastBatchResult -Fatal:(-not $cancelled) -Cancelled:$cancelled -Preview:$preview -Message $message }
            elseif ($namedRun) { New-PodcastRunResult -Fatal:(-not $cancelled) -Cancelled:$cancelled -Preview:$preview -Message $message }
            else { New-PodcastConfigResult -Preview:$preview -ExitCode $(if ($cancelled) { 130 } else { 1 }) -Message $message }
        # Presentation is secondary to the established operation result.
        try { $null = Write-Host ('[ERROR] ' + $message) -ForegroundColor Red }
        catch { $null = $_ }
        return $result
    }
}
