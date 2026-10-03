#requires -Version 5.1

function Get-PodcastSavedRunOptionName {
    [CmdletBinding()]
    param()

    # Routing and strict primitive validation share one argument allowlist.
    'Mode', 'CustomCount', 'OutputPath', 'KeepAwake', 'MaxFeedPages', 'MaxAttempts',
        'HeaderTimeoutSeconds', 'IdleTimeoutSeconds', 'RetryBudgetSeconds', 'BaseDelaySeconds', 'MaxDelaySeconds'
}

function Get-PodcastBatchOverride {
    [CmdletBinding()]
    param([AllowNull()][Collections.IDictionary]$RunOptions)
    $accepted = @(Get-PodcastSavedRunOptionName)
    $options = @{}
    if ($null -eq $RunOptions) { return $options }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($key in @($RunOptions.Keys)) {
        if ($key -isnot [string] -or $key -notin $accepted -or -not $seen.Add($key)) {
            throw 'Batch run options are invalid or unsupported.'
        }
        $value = $RunOptions[$key]
        if ($key -eq 'Mode') {
            if ($value -isnot [string] -or $value -notin @('Latest', 'Custom', 'All')) { throw 'Batch run options are invalid or unsupported.' }
            $options.Mode = @('Latest', 'Custom', 'All') | Where-Object { $_ -eq $value }
        }
        elseif ($key -eq 'OutputPath') {
            if ($value -isnot [string] -or [string]::IsNullOrWhiteSpace($value)) { throw 'Batch run options are invalid or unsupported.' }
            $options.OutputPath = $value
        }
        elseif ($key -eq 'KeepAwake') {
            if ($value -isnot [bool] -and $value -isnot [Management.Automation.SwitchParameter]) { throw 'Batch run options are invalid or unsupported.' }
            $options.KeepAwake = [bool]$value
        }
        else {
            if ($value -isnot [System.Byte] -and $value -isnot [System.Int16] -and $value -isnot [System.Int32] -and $value -isnot [System.Int64] -and
                $value -isnot [System.Single] -and $value -isnot [System.Double] -and $value -isnot [System.Decimal]) {
                throw 'Batch run options are invalid or unsupported.'
            }
            $number = [double]$value
            if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) { throw 'Batch run options are invalid or unsupported.' }
            $minimum = 0.0; $maximum = 3600.0
            if ($key -in @('CustomCount', 'MaxFeedPages', 'MaxAttempts')) {
                $minimum = 1.0
                $maximum = if ($key -eq 'CustomCount') { 2147483647.0 } elseif ($key -eq 'MaxFeedPages') { 100.0 } else { 10.0 }
                if ($number -ne [Math]::Truncate($number)) { throw 'Batch run options are invalid or unsupported.' }
            }
            elseif ($key -in @('HeaderTimeoutSeconds', 'IdleTimeoutSeconds')) { $minimum = 0.001; $maximum = 86400.0 }
            elseif ($key -eq 'RetryBudgetSeconds') { $maximum = 86400.0 }
            if ($number -lt $minimum -or $number -gt $maximum) { throw 'Batch run options are invalid or unsupported.' }
            $options[$key] = if ($key -in @('CustomCount', 'MaxFeedPages', 'MaxAttempts')) { [int]$number } else { $number }
        }
    }
    if ($options.ContainsKey('CustomCount') -and $options.ContainsKey('Mode') -and $options.Mode -ne 'Custom') {
        throw '-CustomCount conflicts with an explicit Latest or All mode.'
    }
    return $options
}

function Resolve-PodcastSavedRunOptions {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Options names the existing command argument set.')]
    [CmdletBinding()]
    param([AllowNull()][Collections.IDictionary]$RunOptions, [Parameter(Mandatory)]$SavedShow)
    $overrides = Get-PodcastBatchOverride -RunOptions $RunOptions
    $arguments = @{ FeedUrl = $SavedShow.FeedUrl; OutputPath = $SavedShow.OutputPath; Mode = $SavedShow.Mode }
    if ($SavedShow.Mode -eq 'Custom' -and $null -ne $SavedShow.CustomCount) { $arguments.CustomCount = $SavedShow.CustomCount }
    foreach ($key in $overrides.Keys) { $arguments[$key] = $overrides[$key] }
    if ($overrides.ContainsKey('CustomCount') -and -not $overrides.ContainsKey('Mode')) { $arguments.Mode = 'Custom' }
    if ($arguments.Mode -ne 'Custom') { $arguments.Remove('CustomCount') }
    $cli = Resolve-PodcastCliOptions -BoundParameters $arguments -Mode $arguments.Mode `
        -CustomCount $arguments.CustomCount -FeedUrl $arguments.FeedUrl
    $arguments.Mode = $cli.Mode
    return $arguments
}

function ConvertTo-PodcastBatchShowResult {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$RunResult)
    if ($RunResult.Type -cne 'Podcast.RunResult' -or $RunResult.SchemaVersion -ne 1 -or $RunResult.ExitCode -notin @(0, 1, 2, 130)) {
        throw 'The saved show did not return a valid run result.'
    }
    $safe = [ordered]@{}
    foreach ($field in @('Type', 'SchemaVersion', 'Mode', 'Preview', 'Complete', 'ExitCode', 'Status', 'Planned',
        'Downloaded', 'VerifiedSkipped', 'LegacyUnverified', 'Conflicts', 'Failed', 'Deferred', 'Cancelled',
        'CatalogueComplete', 'CatalogueStopReason', 'PagesFetched', 'Message')) {
        if ($null -eq $RunResult.PSObject.Properties[$field]) { throw 'The saved show did not return a valid run result.' }
        $safe[$field] = $RunResult.$field
    }
    $episodes = [Collections.Generic.List[object]]::new()
    foreach ($episode in @($RunResult.Episodes)) {
        if ($null -eq $episode) { continue }
        $episodes.Add([pscustomobject]@{ EpisodeId=$episode.EpisodeId; Outcome=$episode.Outcome; Bytes=$episode.Bytes;
            Verification=$episode.Verification; Attempts=$episode.Attempts; Message=$episode.Message })
    }
    $plan = [Collections.Generic.List[object]]::new()
    foreach ($entry in @($RunResult.Plan)) {
        if ($null -eq $entry) { continue }
        $plan.Add([pscustomobject]@{ EpisodeId=$entry.EpisodeId; IdentitySource=$entry.IdentitySource; Source=$entry.Source; Recorded=$entry.Recorded })
    }
    $safe.Episodes = @($episodes.ToArray()); $safe.Plan = @($plan.ToArray())
    return [pscustomobject]$safe
}

function New-PodcastBatchResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only a private in-memory result value.')]
    [CmdletBinding()]
    param([string[]]$SelectedName = @(), [object[]]$Shows = @(), [switch]$Preview, [switch]$Fatal, [switch]$Cancelled, [string]$Message)
    $Shows = @($Shows | Where-Object { $null -ne $_ })
    $totals = @{}
    foreach ($field in @('Planned', 'Downloaded', 'VerifiedSkipped', 'LegacyUnverified', 'Conflicts', 'Failed', 'Deferred', 'Cancelled')) { $totals[$field] = [long]0 }
    $successful = 0; $incomplete = 0; $failed = 0; $stopped = 0
    foreach ($show in $Shows) {
        foreach ($field in @($totals.Keys)) { $totals[$field] += [long]$show.Result.$field }
        switch ([int]$show.Result.ExitCode) { 0 { $successful++ }; 1 { $failed++ }; 2 { $incomplete++ }; 130 { $stopped++ } }
    }
    $unstarted = @($SelectedName | Select-Object -Skip $Shows.Count)
    $exitCode = if ($Cancelled -or $stopped -gt 0) { 130 } elseif ($Fatal) { 1 } elseif ($failed + $incomplete -gt 0) { 2 } else { 0 }
    $status = if ($exitCode -eq 130) { 'cancelled' } elseif ($exitCode -eq 1) { 'fatal' }
        elseif ($exitCode -eq 2) { 'incomplete' } elseif ($Preview) { 'preview' } else { 'success' }
    if ([string]::IsNullOrEmpty($Message)) {
        $Message = switch ($status) {
            'cancelled' { 'Batch cancelled; completed shows and retained recovery evidence were preserved.' }
            'fatal' { 'Saved-show configuration could not be read. Check its format, version and local permissions.' }
            'incomplete' { 'Batch incomplete; one or more saved shows did not complete.' }
            'preview' { 'Batch preview completed.' }
            default { 'Batch completed.' }
        }
    }
    return [pscustomobject]@{
        Type='Podcast.BatchResult'; SchemaVersion=1; Preview=[bool]$Preview; Complete=($exitCode -eq 0 -and $unstarted.Count -eq 0)
        ExitCode=$exitCode; Status=$status; SelectedShows=@($SelectedName).Count; ProcessedShows=$Shows.Count
        UnstartedShowCount=$unstarted.Count; UnstartedShows=@($unstarted); Shows=@($Shows)
        SuccessfulShows=$successful; IncompleteShows=$incomplete; FatalShows=$failed; CancelledShows=$stopped
        Planned=$totals.Planned; Downloaded=$totals.Downloaded; VerifiedSkipped=$totals.VerifiedSkipped
        LegacyUnverified=$totals.LegacyUnverified; Conflicts=$totals.Conflicts; Failed=$totals.Failed
        Deferred=$totals.Deferred; Cancelled=$totals.Cancelled; Message=$Message
    }
}

function Write-PodcastBatchSummary {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Result, [switch]$IncludeFailure, [switch]$PreservePrimary)
    $messages = [Collections.Generic.List[string]]::new()
    if ($IncludeFailure) { $messages.Add('[ERROR] ' + $Result.Message) }
    $messages.Add(('Batch summary: selected {0}; processed {1}; successful {2}; incomplete {3}; fatal {4}; cancelled {5}; unstarted {6}.' -f
        $Result.SelectedShows, $Result.ProcessedShows, $Result.SuccessfulShows, $Result.IncompleteShows, $Result.FatalShows,
        $Result.CancelledShows, $Result.UnstartedShowCount))
    $messages.Add(('Batch episodes: planned {0}; downloaded {1}; verified-skipped {2}; legacy-unverified {3}; conflicts {4}; failed {5}; deferred {6}; cancelled {7}.' -f
        $Result.Planned, $Result.Downloaded, $Result.VerifiedSkipped, $Result.LegacyUnverified, $Result.Conflicts,
        $Result.Failed, $Result.Deferred, $Result.Cancelled))
    foreach ($message in $messages) {
        try { $null = Write-Host $message }
        catch {
            # A secondary presentation failure must not replace an established
            # fatal/cancelled result. Normal-summary Control+C still cancels.
            if (-not $PreservePrimary -and (Test-PodcastCancellation -ErrorObject $_)) { throw }
        }
    }
}

function Invoke-PodcastBatchRun {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$ConfigPath, [AllowNull()][AllowEmptyCollection()][string[]]$ShowName,
        [AllowNull()][Collections.IDictionary]$RunOptions)
    $selectedNames = @(); $shows = [Collections.Generic.List[object]]::new(); $preview = [bool]$WhatIfPreference
    $failureMessage = 'Saved-show configuration could not be read. Check its format, version and local permissions.'
    try {
        $failureMessage = 'Batch run options are invalid or unsupported.'
        $overrides = Get-PodcastBatchOverride -RunOptions $RunOptions
        $failureMessage = 'Saved-show configuration could not be read. Check its format, version and local permissions.'
        $configuration = Read-PodcastSavedShowConfig -ConfigPath $ConfigPath
        $available = @($configuration.shows)
        $failureMessage = 'The batch show selection is invalid. Choose unique existing saved shows.'
        if ($available.Count -gt 100) { throw $failureMessage }
        $selected = [Collections.Generic.List[object]]::new()
        if ($PSBoundParameters.ContainsKey('ShowName')) {
            if ($null -eq $ShowName -or $ShowName.Count -eq 0 -or $ShowName.Count -gt 100) { throw $failureMessage }
            $requested = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($name in $ShowName) {
                if ([string]::IsNullOrWhiteSpace($name) -or -not $requested.Add($name)) { throw $failureMessage }
                $match = @($available | Where-Object { $_.name -eq $name })
                if ($match.Count -ne 1) { throw $failureMessage }
                $selected.Add($match[0])
            }
        }
        else { foreach ($show in $available) { $selected.Add($show) } }
        $selectedNames = @($selected | ForEach-Object { $_.name })
        $failureMessage = 'Batch run options are invalid or unsupported.'
        foreach ($stored in $selected) {
            $settings = [pscustomobject]@{ FeedUrl=''; OutputPath=$stored.output_path; Mode=$stored.mode; CustomCount=$stored.custom_count }
            $null = Resolve-PodcastSavedRunOptions -RunOptions $overrides -SavedShow $settings
        }
        if ($selected.Count -gt 0) {
            $confirmed = $PSCmdlet.ShouldProcess(('Selected saved shows: {0}' -f $selected.Count), 'Run sequential podcast batch')
            if (-not $confirmed -and -not $preview) {
                return New-PodcastBatchResult -SelectedName $selectedNames -Preview -Message 'Batch was not started because confirmation was declined.'
            }
        }
        foreach ($stored in $selected) {
            try {
                $savedShow = Get-PodcastSavedShow -ConfigPath $ConfigPath -Name $stored.name
                $childArguments = Resolve-PodcastSavedRunOptions -RunOptions $overrides -SavedShow $savedShow
                $childArguments.NonInteractive = $true; $childArguments.Confirm = $false; $childArguments.WhatIf = $preview
                # One child owns and releases all of its archive/native resources
                # before the next saved show is resolved or invoked.
                $childResult = Invoke-PodcastRun @childArguments
                $childResult = ConvertTo-PodcastBatchShowResult -RunResult $childResult
            }
            catch {
                if ($_.Exception.Message -in @('Saved-show configuration is malformed or unsupported. Preserved the configuration.',
                    'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.')) {
                    $failureMessage = 'Saved-show configuration could not be read. Check its format, version and local permissions.'
                    throw
                }
                $cancelled = Test-PodcastCancellation -ErrorObject $_
                $safeMessage = if ($cancelled) { 'Run cancelled; retained media and recovery evidence were preserved.' }
                    elseif ($_.Exception.Message -ceq 'Saved-show credentials are unavailable for the current Windows user.') { $_.Exception.Message }
                    else { 'The saved show failed. Check its local settings and retry.' }
                $childResult = New-PodcastRunResult -Mode $stored.mode -Preview:$preview -Fatal:(-not $cancelled) -Cancelled:$cancelled -Message $safeMessage
            }
            $shows.Add([pscustomobject]@{ Name=$stored.name; Result=$childResult })
            if ($childResult.ExitCode -eq 130) { break }
        }
        $result = New-PodcastBatchResult -SelectedName $selectedNames -Shows $shows.ToArray() -Preview:$preview
        Write-PodcastBatchSummary -Result $result
        return $result
    }
    catch {
        $cancelled = Test-PodcastCancellation -ErrorObject $_
        $result = New-PodcastBatchResult -SelectedName $selectedNames -Shows $shows.ToArray() -Preview:$preview `
            -Fatal:(-not $cancelled) -Cancelled:$cancelled -Message $(if ($cancelled) { 'Batch cancelled; completed shows and retained recovery evidence were preserved.' } else { $failureMessage })
        Write-PodcastBatchSummary -Result $result -IncludeFailure -PreservePrimary
        return $result
    }
}
