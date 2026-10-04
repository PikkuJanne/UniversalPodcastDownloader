BeforeAll {
    $script:batchRepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:batchRepositoryRoot 'UniversalPodcastDownloader.ps1')
    $batchSourcePath = Join-Path $script:batchRepositoryRoot 'src/Batch.ps1'
    if (Test-Path -LiteralPath $batchSourcePath) { . $batchSourcePath }
    if (-not ('UPDBatchDeclineHost' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Globalization;
using System.Security;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Host;
public sealed class UPDBatchDeclineHost : PSHost {
    private readonly Guid id = Guid.NewGuid();
    private readonly DeclineUI ui = new DeclineUI();
    public int ConfirmationCount { get { return ui.ConfirmationCount; } }
    public override string Name { get { return "Owned batch confirmation fixture"; } }
    public override Version Version { get { return new Version(1,0); } }
    public override Guid InstanceId { get { return id; } }
    public override PSHostUserInterface UI { get { return ui; } }
    public override CultureInfo CurrentCulture { get { return CultureInfo.InvariantCulture; } }
    public override CultureInfo CurrentUICulture { get { return CultureInfo.InvariantCulture; } }
    public override void SetShouldExit(int code) { throw new InvalidOperationException("A batch unit must not exit its host."); }
    public override void EnterNestedPrompt() { throw new InvalidOperationException("No nested prompts in batch units."); }
    public override void ExitNestedPrompt() { }
    public override void NotifyBeginApplication() { }
    public override void NotifyEndApplication() { }
    private sealed class DeclineUI : PSHostUserInterface {
        public int ConfirmationCount;
        public override PSHostRawUserInterface RawUI { get { return null; } }
        public override string ReadLine() { throw new InvalidOperationException("No guided input in batch units."); }
        public override SecureString ReadLineAsSecureString() { throw new InvalidOperationException("No credentials in batch units."); }
        public override void Write(string value) { }
        public override void Write(ConsoleColor fg, ConsoleColor bg, string value) { }
        public override void WriteLine(string value) { }
        public override void WriteErrorLine(string value) { }
        public override void WriteDebugLine(string value) { }
        public override void WriteProgress(long sourceId, ProgressRecord value) { }
        public override void WriteVerboseLine(string value) { }
        public override void WriteWarningLine(string value) { }
        public override Dictionary<string,PSObject> Prompt(string caption, string message, Collection<FieldDescription> fields) { throw new InvalidOperationException("No field prompt in batch units."); }
        public override PSCredential PromptForCredential(string caption, string message, string user, string target) { throw new InvalidOperationException("No credential prompt in batch units."); }
        public override PSCredential PromptForCredential(string caption, string message, string user, string target, PSCredentialTypes types, PSCredentialUIOptions options) { throw new InvalidOperationException("No credential prompt in batch units."); }
        public override int PromptForChoice(string caption, string message, Collection<ChoiceDescription> choices, int defaultChoice) {
            ConfirmationCount++;
            for (int i=0; i<choices.Count; i++) if (choices[i].Label.Replace("&", "").Equals("No", StringComparison.OrdinalIgnoreCase)) return i;
            throw new InvalidOperationException("The fixture expected an actual No confirmation choice.");
        }
    }
}
'@
    }
}

Describe 'A047 validated selection and private batch outcomes' -Tag 'Unit', 'A047' {
    BeforeEach {
        $script:batchOwnedOutput = Join-Path $TestDrive 'owned-output'
        $script:batchConfiguration = [pscustomobject]@{ schema_version=1; shows=@(
            [pscustomobject]@{ name='alpha'; mode='Latest'; custom_count=$null; output_path=$script:batchOwnedOutput; feed_protected='cipher-canary-alpha' },
            [pscustomobject]@{ name='beta'; mode='Custom'; custom_count=2; output_path=$script:batchOwnedOutput; feed_protected='cipher-canary-beta' },
            [pscustomobject]@{ name='gamma'; mode='All'; custom_count=$null; output_path=$script:batchOwnedOutput; feed_protected='cipher-canary-gamma' }
        ) }
        $script:batchInvocations = [Collections.Generic.List[object]]::new()
        $script:batchExpressionExecuted = $false
        Mock Read-Host { throw 'Batch units must not prompt.' }
        Mock Get-PodcastHttpClient { throw 'Batch units must not contact a network.' }
        Mock Write-Host {}
        Mock Read-PodcastSavedShowConfig { $script:batchConfiguration }
        Mock Get-PodcastSavedShow {
            $stored = @($script:batchConfiguration.shows | Where-Object { $_.name -eq $Name })[0]
            [pscustomobject]@{ Name=$stored.name; FeedUrl=('https://feeds.example.invalid/' + $stored.name + '.xml');
                OutputPath=$stored.output_path; Mode=$stored.mode; CustomCount=$stored.custom_count }
        }
        Mock Invoke-PodcastRun {
            $script:batchInvocations.Add([pscustomobject]@{ FeedUrl=$FeedUrl; Mode=$Mode; CustomCount=$CustomCount;
                OutputPath=$OutputPath; WhatIf=$WhatIf; Confirm=$Confirm; NonInteractive=$NonInteractive; KeepAwake=$KeepAwake;
                MaxAttempts=$MaxAttempts; MaxFeedPages=$MaxFeedPages; HeaderTimeoutSeconds=$HeaderTimeoutSeconds;
                IdleTimeoutSeconds=$IdleTimeoutSeconds; RetryBudgetSeconds=$RetryBudgetSeconds;
                BaseDelaySeconds=$BaseDelaySeconds; MaxDelaySeconds=$MaxDelaySeconds })
            New-PodcastRunResult -Mode $Mode -Preview:$WhatIf
        }
    }

    It 'returns stable empty arrays for an absent or empty configuration' {
        $script:batchConfiguration.shows = @()
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'absent.json') -Confirm:$false
        $result.ExitCode | Should -Be 0
        $result.Complete | Should -BeTrue
        $result.SelectedShows | Should -Be 0
        $result.ProcessedShows | Should -Be 0
        $result.UnstartedShowCount | Should -Be 0
        $result.Shows -is [array] | Should -BeTrue
        $result.UnstartedShows -is [array] | Should -BeTrue
        @($result.Shows).Count | Should -Be 0
        @($result.UnstartedShows).Count | Should -Be 0
        Should -Invoke Get-PodcastSavedShow -Times 0 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }

    It 'uses explicit request order and case-insensitive lookup while preserving stored names' {
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName @('GAMMA','ALPHA') -Confirm:$false
        @($result.Shows | ForEach-Object { $_.Name }) | Should -Be @('gamma','alpha')
        @($script:batchInvocations | ForEach-Object { $_.FeedUrl }) | Should -Be @('https://feeds.example.invalid/gamma.xml','https://feeds.example.invalid/alpha.xml')
        $result.SelectedShows | Should -Be 2
        $result.ProcessedShows | Should -Be 2
        foreach ($call in $script:batchInvocations) { $call.NonInteractive | Should -BeTrue; $call.Confirm | Should -BeFalse }
    }

    It 'rejects <Kind> selection before resolving credentials or invoking a child' -ForEach @(
        @{ Kind='null'; Names=$null }, @{ Kind='empty'; Names=@() }, @{ Kind='blank'; Names=@(' ') },
        @{ Kind='duplicate-case'; Names=@('alpha','ALPHA') }, @{ Kind='unknown'; Names=@('missing') },
        @{ Kind='oversized'; Names=@(1..101 | ForEach-Object { 'show' + $_ }) }
    ) {
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName $Names -Confirm:$false
        $result.ExitCode | Should -Be 1
        $result.Status | Should -BeExactly 'fatal'
        $result.Message | Should -BeExactly 'The batch show selection is invalid. Choose unique existing saved shows.'
        $result.ProcessedShows | Should -Be 0
        Should -Invoke Get-PodcastSavedShow -Times 0 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
        Test-Path -LiteralPath $script:batchOwnedOutput | Should -BeFalse
    }

    It 'rejects disallowed option <Key> before configuration reads or child side effects' -ForEach @(
        @{ Key='FeedUrl' }, @{ Key='LegacyPath' }, @{ Key='LegacyAction' }, @{ Key='DiagnosticExportPath' },
        @{ Key='PassThru' }, @{ Key='ShowName' }, @{ Key='ConfigPath' }, @{ Key='NonInteractive' },
        @{ Key='Confirm' }, @{ Key='WhatIf' }, @{ Key='ErrorAction' }
    ) {
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -RunOptions @{ $Key='https://canary.invalid/private-token' } -Confirm:$false
        $result.ExitCode | Should -Be 1
        $result.Message | Should -BeExactly 'Batch run options are invalid or unsupported.'
        ($result | ConvertTo-Json -Depth 10) | Should -Not -Match 'private-token|canary.invalid'
        Should -Invoke Read-PodcastSavedShowConfig -Times 0 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }

    It 'rejects invalid primitive <Kind> before any child' -ForEach @(
        @{ Kind='count-zero'; Options=@{CustomCount=0} }, @{ Kind='count-fraction'; Options=@{CustomCount=1.5} },
        @{ Kind='attempts-too-many'; Options=@{MaxAttempts=11} }, @{ Kind='pages-too-many'; Options=@{MaxFeedPages=101} },
        @{ Kind='zero-header-wait'; Options=@{HeaderTimeoutSeconds=0} }, @{ Kind='negative-idle-wait'; Options=@{IdleTimeoutSeconds=-1} },
        @{ Kind='nan'; Options=@{RetryBudgetSeconds=[double]::NaN} }, @{ Kind='infinity'; Options=@{MaxDelaySeconds=[double]::PositiveInfinity} },
        @{ Kind='blank-output'; Options=@{OutputPath=' '} }, @{ Kind='string-switch'; Options=@{KeepAwake='false'} },
        @{ Kind='script'; Options=@{Mode={ $script:batchExpressionExecuted=$true; 'All' }} },
        @{ Kind='conflicting-mode-count'; Options=@{Mode='All';CustomCount=2} }
    ) {
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -RunOptions $Options -Confirm:$false
        $result.ExitCode | Should -Be 1
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
        Should -Invoke Get-PodcastSavedShow -Times 0 -Exactly
        $script:batchExpressionExecuted | Should -BeFalse
        Test-Path -LiteralPath $script:batchOwnedOutput | Should -BeFalse
    }

    It 'passes the complete allowed transport override set and count-only Custom without mutating caller data' {
        $overrideRoot = Join-Path $TestDrive 'override-output'
        $options = @{ CustomCount=3; OutputPath=$overrideRoot; KeepAwake=$true; MaxFeedPages=7; MaxAttempts=2;
            HeaderTimeoutSeconds=1.5; IdleTimeoutSeconds=2.5; RetryBudgetSeconds=8.0; BaseDelaySeconds=0.0; MaxDelaySeconds=0.5 }
        $beforeOptions = $options | ConvertTo-Json -Compress
        $beforeConfiguration = $script:batchConfiguration | ConvertTo-Json -Depth 6 -Compress
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -RunOptions $options -Confirm:$false
        $result.ExitCode | Should -Be 0
        $result.ProcessedShows | Should -Be 3
        foreach ($call in $script:batchInvocations) {
            $call.Mode | Should -BeExactly 'Custom'; $call.CustomCount | Should -Be 3
            $call.OutputPath | Should -BeExactly $overrideRoot; $call.KeepAwake | Should -BeTrue
            $call.MaxFeedPages | Should -Be 7; $call.MaxAttempts | Should -Be 2
            $call.HeaderTimeoutSeconds | Should -Be 1.5; $call.IdleTimeoutSeconds | Should -Be 2.5
            $call.RetryBudgetSeconds | Should -Be 8; $call.BaseDelaySeconds | Should -Be 0; $call.MaxDelaySeconds | Should -Be 0.5
        }
        ($options | ConvertTo-Json -Compress) | Should -BeExactly $beforeOptions
        ($script:batchConfiguration | ConvertTo-Json -Depth 6 -Compress) | Should -BeExactly $beforeConfiguration
    }

    It 'uses explicit <Mode> without accidentally forwarding the saved Custom count' -ForEach @(@{Mode='Latest'},@{Mode='All'}) {
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName beta -RunOptions @{Mode=$Mode} -Confirm:$false
        $result.ExitCode | Should -Be 0
        $script:batchInvocations[0].Mode | Should -BeExactly $Mode
        Should -Invoke Invoke-PodcastRun -Times 1 -Exactly -ParameterFilter { $CustomCount -eq 0 }
    }

    It 'retains a stored Custom count but rejects an unusable Mode Custom for a later show before the first child' {
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName beta -RunOptions @{Mode='Custom'} -Confirm:$false
        $result.ExitCode | Should -Be 0
        $script:batchInvocations[0].CustomCount | Should -Be 2
        $script:batchInvocations.Clear()
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName @('beta','alpha') -RunOptions @{Mode='Custom'} -Confirm:$false
        $result.ExitCode | Should -Be 1
        $result.ProcessedShows | Should -Be 0
        @($result.UnstartedShows) | Should -Be @('beta','alpha')
        $script:batchInvocations.Count | Should -Be 0
        Should -Invoke Get-PodcastSavedShow -Times 1 -Exactly
    }

    It 'isolates missing credentials, excludes the raw failure and still runs the next selected show' {
        Mock Get-PodcastSavedShow {
            if ($Name -eq 'alpha') { throw 'Saved-show credentials are unavailable for the current Windows user.' }
            [pscustomobject]@{Name=$Name;FeedUrl='https://feeds.example.invalid/beta.xml';OutputPath=$script:batchOwnedOutput;Mode='Custom';CustomCount=2}
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName @('alpha','beta') -Confirm:$false
        $result.ExitCode | Should -Be 2
        $result.FatalShows | Should -Be 1
        $result.SuccessfulShows | Should -Be 1
        $result.ProcessedShows | Should -Be 2
        $result.Shows[0].Result.Message | Should -BeExactly 'Saved-show credentials are unavailable for the current Windows user.'
        Should -Invoke Invoke-PodcastRun -Times 1 -Exactly
    }

    It 'projects safe result fields and summary counts without legacy data, URL, cipher or path leakage' {
        Mock Invoke-PodcastRun {
            $run = New-PodcastRunResult -Mode Latest -Planned 1 -EpisodeResults @(New-PodcastEpisodeResult -EpisodeId ('a'*64) -Outcome downloaded -Bytes 8) `
                -LegacyResult ([pscustomobject]@{Path='legacy-path-canary';Url='https://secret.example.invalid/raw-token'})
            $run | Add-Member NoteProperty OutputPath 'output-path-canary'
            $run | Add-Member NoteProperty Cipher 'cipher-canary-alpha'
            $run
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName alpha -Confirm:$false
        $result.ExitCode | Should -Be 0
        $result.Downloaded | Should -Be 1
        $serialized = $result | ConvertTo-Json -Depth 10
        $serialized | Should -Not -Match 'LegacyResult|legacy-path-canary|output-path-canary|cipher-canary|raw-token|https://'
        $result.Shows[0].Result.Episodes -is [array] | Should -BeTrue
        $result.Shows[0].Result.Plan -is [array] | Should -BeTrue
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { $Object -like 'Batch summary: selected 1; processed 1;*' -and $Object -notmatch 'alpha|https://|owned-output' }
    }

    It 'keeps every episode outcome count distinct while isolating an incomplete catalogue and fatal show' {
        Mock Invoke-PodcastRun {
            if ($FeedUrl.EndsWith('/alpha.xml')) {
                $episodes = @('downloaded','verified_skip','legacy_unverified','conflict','failed','deferred') | ForEach-Object {
                    New-PodcastEpisodeResult -EpisodeId ('a'*64) -Outcome $_
                }
                return New-PodcastRunResult -Mode Latest -Planned 6 -EpisodeResults $episodes -Catalogue ([pscustomobject]@{Complete=$false;StopReason='page_failed';PagesFetched=2})
            }
            New-PodcastRunResult -Mode Custom -Fatal -Message 'Safe authored failure.'
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName @('alpha','beta') -Confirm:$false
        $result.ExitCode | Should -Be 2
        $result.Status | Should -BeExactly 'incomplete'
        $result.IncompleteShows | Should -Be 1
        $result.FatalShows | Should -Be 1
        $result.Planned | Should -Be 6
        foreach ($field in @('Downloaded','VerifiedSkipped','LegacyUnverified','Conflicts','Failed','Deferred')) { $result.$field | Should -Be 1 }
        $result.Cancelled | Should -Be 0
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
            $Object -eq 'Batch episodes: planned 6; downloaded 1; verified-skipped 1; legacy-unverified 1; conflicts 1; failed 1; deferred 1; cancelled 0.'
        }
    }

    It 'stops on a returned cleanup cancellation while retaining committed download counts and only remaining names' {
        Mock Invoke-PodcastRun {
            New-PodcastRunResult -Mode Latest -Cancelled -Planned 1 -EpisodeResults @(New-PodcastEpisodeResult -EpisodeId ('a'*64) -Outcome downloaded -Bytes 8)
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 130
        $result.CancelledShows | Should -Be 1
        $result.Downloaded | Should -Be 1
        $result.Cancelled | Should -Be 0
        $result.SelectedShows | Should -Be 3
        $result.ProcessedShows | Should -Be 1
        $result.UnstartedShowCount | Should -Be 2
        @($result.UnstartedShows) | Should -Be @('beta','gamma')
        Should -Invoke Get-PodcastSavedShow -Times 1 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 1 -Exactly
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
            $Object -eq 'Batch episodes: planned 1; downloaded 1; verified-skipped 0; legacy-unverified 0; conflicts 0; failed 0; deferred 0; cancelled 0.'
        }
    }

    It 'preserves typed <Kind> cancellation and does not invent cancelled episode rows' -ForEach @(
        @{Kind='during-credentials'}, @{Kind='during-child'}, @{Kind='during-summary'}
    ) {
        if ($Kind -eq 'during-credentials') { Mock Get-PodcastSavedShow { throw [OperationCanceledException]::new('private-cancellation-canary') } }
        elseif ($Kind -eq 'during-child') { Mock Invoke-PodcastRun { throw [OperationCanceledException]::new('private-cancellation-canary') } }
        else { Mock Write-Host { throw [OperationCanceledException]::new('private-cancellation-canary') } }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 130
        $result.Status | Should -BeExactly 'cancelled'
        $result.Cancelled | Should -Be 0
        ($result | ConvertTo-Json -Depth 10) | Should -Not -Match 'private-cancellation-canary'
        if ($Kind -ne 'during-summary') {
            $result.ProcessedShows | Should -Be 1
            @($result.UnstartedShows) | Should -Be @('beta','gamma')
        }
        else { $result.ProcessedShows | Should -Be 3; $result.CancelledShows | Should -Be 0 }
    }

    It 'stops globally on a structural configuration change after a completed show' {
        Mock Get-PodcastSavedShow {
            if ($Name -eq 'beta') { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
            [pscustomobject]@{Name=$Name;FeedUrl='https://feeds.example.invalid/alpha.xml';OutputPath=$script:batchOwnedOutput;Mode='Latest';CustomCount=$null}
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 1
        $result.ProcessedShows | Should -Be 1
        $result.SuccessfulShows | Should -Be 1
        @($result.UnstartedShows) | Should -Be @('beta','gamma')
        Should -Invoke Invoke-PodcastRun -Times 1 -Exactly
        Should -Invoke Get-PodcastSavedShow -Times 2 -Exactly
    }

    It 'maps an ordinary raw child failure safely and continues without leaking publisher content' {
        Mock Invoke-PodcastRun {
            if ($FeedUrl.EndsWith('/alpha.xml')) { throw [IO.IOException]::new('https://secret.example.invalid/private-key output-path-canary') }
            New-PodcastRunResult -Mode Custom
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName @('alpha','beta') -Confirm:$false
        $result.ExitCode | Should -Be 2
        $result.FatalShows | Should -Be 1
        $result.SuccessfulShows | Should -Be 1
        ($result | ConvertTo-Json -Depth 10) | Should -Not -Match 'private-key|secret.example|output-path-canary'
        $result.Shows[0].Result.Message | Should -BeExactly 'The saved show failed. Check its local settings and retry.'
    }

    It 'runs sequentially and permits the next show to acquire the exact resource released by the previous child' {
        $script:batchResourcePath = Join-Path $TestDrive 'owned-exclusive-resource.lock'
        $script:batchActiveChildren = 0
        $script:batchMaxActive = 0
        Mock Invoke-PodcastRun {
            $handle = [IO.File]::Open($script:batchResourcePath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            $script:batchActiveChildren++
            $script:batchMaxActive = [Math]::Max($script:batchMaxActive, $script:batchActiveChildren)
            try { New-PodcastRunResult -Mode $Mode }
            finally { $handle.Dispose(); $script:batchActiveChildren-- }
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 0
        $result.ProcessedShows | Should -Be 3
        $script:batchMaxActive | Should -Be 1
        $script:batchActiveChildren | Should -Be 0
        $guard = [IO.File]::Open($script:batchResourcePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $guard.Dispose()
    }

    It 'uses long aggregate counters without overflow across large legitimate per-show plans' {
        Mock Invoke-PodcastRun { New-PodcastRunResult -Mode $Mode -Planned ([int]::MaxValue) }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.Planned | Should -Be ([long][int]::MaxValue * 3)
        $result.Planned -is [long] | Should -BeTrue
        $result.ExitCode | Should -Be 0
    }

    It 'treats an ordinary summary rendering error as advisory after child outcomes are established' {
        Mock Write-Host { throw [IO.IOException]::new('private-render-canary') }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 0
        $result.ProcessedShows | Should -Be 3
        ($result | ConvertTo-Json -Depth 10) | Should -Not -Match 'private-render-canary'
    }

    It 'declines one actual outer confirmation without resolving credentials or running any child' {
        $decliningHost = [UPDBatchDeclineHost]::new()
        $runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($decliningHost)
        $powershell = [Management.Automation.PowerShell]::Create()
        try {
            $runspace.Open(); $powershell.Runspace = $runspace
            $null = $powershell.AddScript(@'
param($root)
. (Join-Path $root 'src/RunResult.ps1')
. (Join-Path $root 'src/Batch.ps1')
function Read-PodcastSavedShowConfig {
    param($ConfigPath)
    [pscustomobject]@{schema_version=1;shows=@([pscustomobject]@{name='alpha';mode='Latest';custom_count=$null;output_path=$ConfigPath})}
}
function Get-PodcastSavedShow { throw 'Credentials must not be resolved after declined confirmation.' }
function Invoke-PodcastRun { throw 'A child must not run after declined confirmation.' }
Invoke-PodcastBatchRun -ConfigPath 'owned-unused-config.json' -Confirm
'@).AddArgument($script:batchRepositoryRoot)
            $output = @($powershell.Invoke())
            $powershell.HadErrors | Should -BeFalse
            $output.Count | Should -Be 1
            $output[0].Type | Should -BeExactly 'Podcast.BatchResult'
            $output[0].Preview | Should -BeTrue
            $output[0].Complete | Should -BeFalse
            $output[0].ExitCode | Should -Be 0
            $output[0].SelectedShows | Should -Be 1
            $output[0].ProcessedShows | Should -Be 0
            $output[0].UnstartedShowCount | Should -Be 1
            @($output[0].UnstartedShows) | Should -Be @('alpha')
            $decliningHost.ConfirmationCount | Should -Be 1
        }
        finally { $powershell.Dispose(); $runspace.Dispose() }
    }

    It 'returns a global fatal setup result for a malformed read without exposing raw errors or starting work' {
        Mock Read-PodcastSavedShowConfig { throw [IO.IOException]::new('https://secret.example.invalid/raw-config-token path-canary') }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 1
        $result.Status | Should -BeExactly 'fatal'
        $result.ProcessedShows | Should -Be 0
        ($result | ConvertTo-Json -Depth 8) | Should -Not -Match 'secret.example|raw-config-token|path-canary'
        Should -Invoke Get-PodcastSavedShow -Times 0 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }

    It 'returns only one batch value when a console sink accidentally emits pipeline output' {
        Mock Write-Host { 'owned accidental host output' }
        $result = @(Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false)
        $result.Count | Should -Be 1
        $result[0].Type | Should -BeExactly 'Podcast.BatchResult'
        $result[0].ProcessedShows | Should -Be 3
    }

    It 'desired batch failure emits fixed setup guidance and honest zero-work summaries' {
        Mock Read-PodcastSavedShowConfig { throw [IO.IOException]::new('https://secret.example.invalid/raw-setup-token path-canary') }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 1
        $result.ProcessedShows | Should -Be 0
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
            $Object -eq '[ERROR] Saved-show configuration could not be read. Check its format, version and local permissions.'
        }
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
            $Object -eq 'Batch summary: selected 0; processed 0; successful 0; incomplete 0; fatal 0; cancelled 0; unstarted 0.'
        }
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
            $Object -eq 'Batch episodes: planned 0; downloaded 0; verified-skipped 0; legacy-unverified 0; conflicts 0; failed 0; deferred 0; cancelled 0.'
        }
        ($result | ConvertTo-Json -Depth 10) | Should -Not -Match 'raw-setup-token|secret.example|path-canary'
    }

    It 'desired batch failure reports retained completion counts after a structural error between children' {
        Mock Get-PodcastSavedShow {
            if ($Name -eq 'beta') { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
            [pscustomobject]@{Name='alpha';FeedUrl='https://feeds.example.invalid/alpha.xml';OutputPath=$script:batchOwnedOutput;Mode='Latest';CustomCount=$null}
        }
        Mock Invoke-PodcastRun {
            New-PodcastRunResult -Mode Latest -Planned 1 -EpisodeResults @(New-PodcastEpisodeResult -EpisodeId ('a'*64) -Outcome downloaded -Bytes 8)
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 1
        $result.Downloaded | Should -Be 1
        $result.ProcessedShows | Should -Be 1
        @($result.UnstartedShows) | Should -Be @('beta','gamma')
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
            $Object -eq '[ERROR] Saved-show configuration could not be read. Check its format, version and local permissions.'
        }
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
            $Object -eq 'Batch summary: selected 3; processed 1; successful 1; incomplete 0; fatal 0; cancelled 0; unstarted 2.'
        }
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
            $Object -eq 'Batch episodes: planned 1; downloaded 1; verified-skipped 0; legacy-unverified 0; conflicts 0; failed 0; deferred 0; cancelled 0.'
        }
    }

    It 'preserves primary setup cancellation despite a <RendererFailure> renderer failure' -ForEach @(
        @{RendererFailure='ordinary'}, @{RendererFailure='typed cancellation'}
    ) {
        Mock Read-PodcastSavedShowConfig { throw [OperationCanceledException]::new('private-primary-cancel-canary') }
        if ($RendererFailure -eq 'ordinary') { Mock Write-Host { throw [IO.IOException]::new('private-render-canary') } }
        else { Mock Write-Host { throw [OperationCanceledException]::new('private-render-canary') } }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 130
        $result.Status | Should -BeExactly 'cancelled'
        $result.Message | Should -BeExactly 'Batch cancelled; completed shows and retained recovery evidence were preserved.'
        $result.ProcessedShows | Should -Be 0
        $result.Cancelled | Should -Be 0
        @($result.Shows).Count | Should -Be 0
        ($result | ConvertTo-Json -Depth 10) | Should -Not -Match 'private-primary-cancel-canary|private-render-canary'
        Should -Invoke Write-Host -Times 3 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }

    It 'preserves primary structural failure and committed counts despite a <RendererFailure> renderer failure' -ForEach @(
        @{RendererFailure='ordinary'}, @{RendererFailure='typed cancellation'}
    ) {
        Mock Get-PodcastSavedShow {
            if ($Name -eq 'beta') { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
            [pscustomobject]@{Name='alpha';FeedUrl='https://feeds.example.invalid/alpha.xml';OutputPath=$script:batchOwnedOutput;Mode='Latest';CustomCount=$null}
        }
        Mock Invoke-PodcastRun {
            New-PodcastRunResult -Mode Latest -Planned 1 -EpisodeResults @(New-PodcastEpisodeResult -EpisodeId ('a'*64) -Outcome downloaded -Bytes 8)
        }
        if ($RendererFailure -eq 'ordinary') { Mock Write-Host { throw [IO.IOException]::new('private-render-canary') } }
        else { Mock Write-Host { throw [OperationCanceledException]::new('private-render-canary') } }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        $result.ExitCode | Should -Be 1
        $result.Status | Should -BeExactly 'fatal'
        $result.Message | Should -BeExactly 'Saved-show configuration could not be read. Check its format, version and local permissions.'
        $result.ProcessedShows | Should -Be 1
        $result.Downloaded | Should -Be 1
        $result.Cancelled | Should -Be 0
        $result.Shows[0].Result.Episodes[0].Bytes | Should -Be 8
        @($result.UnstartedShows) | Should -Be @('beta','gamma')
        ($result | ConvertTo-Json -Depth 10) | Should -Not -Match 'private-render-canary'
        Should -Invoke Write-Host -Times 3 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 1 -Exactly
    }
}

Describe 'A047 desired bounded sequential batch behavior' -Tag 'Unit', 'A047' {
    BeforeEach {
        Mock Read-Host { throw 'Batch units must not prompt.' }
        Mock Get-PodcastHttpClient { throw 'Batch units must not contact a network.' }
        Mock Write-Host {}
        Mock Read-PodcastSavedShowConfig {
            [pscustomobject]@{ schema_version=1; shows=@(
                [pscustomobject]@{ name='alpha'; mode='Latest'; custom_count=$null; output_path=(Join-Path $TestDrive 'owned') },
                [pscustomobject]@{ name='beta'; mode='Latest'; custom_count=$null; output_path=(Join-Path $TestDrive 'owned') }
            ) }
        }
        Mock Get-PodcastSavedShow {
            [pscustomobject]@{ Name=$Name; FeedUrl=('https://feeds.example.invalid/' + $Name + '.xml');
                OutputPath=(Join-Path $TestDrive 'owned'); Mode='Latest'; CustomCount=$null }
        }
    }

    It 'previews one saved show through the existing run without media or state writes' {
        Mock Invoke-PodcastRun {
            New-PodcastRunResult -Mode Latest -Preview -Planned 1
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -ShowName alpha -WhatIf
        $result.Type | Should -BeExactly 'Podcast.BatchResult'
        $result.SchemaVersion | Should -Be 1
        $result.Preview | Should -BeTrue
        $result.ExitCode | Should -Be 0
        $result.SelectedShows | Should -Be 1
        $result.ProcessedShows | Should -Be 1
        $result.Shows -is [array] | Should -BeTrue
        @($result.Shows).Count | Should -Be 1
        $result.Shows[0].Name | Should -BeExactly 'alpha'
        Should -Invoke Invoke-PodcastRun -Times 1 -Exactly -ParameterFilter { $WhatIf -and $NonInteractive -and -not $Confirm }
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
        Test-Path -LiteralPath (Join-Path $TestDrive 'owned') | Should -BeFalse
    }

    It 'isolates the first show fatal result and processes the later show in configuration order' {
        $script:batchCalls = [Collections.Generic.List[string]]::new()
        Mock Invoke-PodcastRun {
            $script:batchCalls.Add($FeedUrl)
            if ($FeedUrl.EndsWith('/alpha.xml')) { return New-PodcastRunResult -Mode Latest -Fatal -Message 'The requested show could not be processed.' }
            New-PodcastRunResult -Mode Latest -Planned 1 -EpisodeResults @(New-PodcastEpisodeResult -EpisodeId ('b' * 64) -Outcome downloaded -Bytes 8)
        }
        $result = Invoke-PodcastBatchRun -ConfigPath (Join-Path $TestDrive 'shows.json') -Confirm:$false
        @($script:batchCalls) | Should -Be @('https://feeds.example.invalid/alpha.xml', 'https://feeds.example.invalid/beta.xml')
        $result.ExitCode | Should -Be 2
        $result.Complete | Should -BeFalse
        $result.SelectedShows | Should -Be 2
        $result.ProcessedShows | Should -Be 2
        $result.FatalShows | Should -Be 1
        $result.SuccessfulShows | Should -Be 1
        $result.Downloaded | Should -Be 1
        @($result.UnstartedShows).Count | Should -Be 0
        Should -Invoke Read-Host -Times 0 -Exactly
    }
}
