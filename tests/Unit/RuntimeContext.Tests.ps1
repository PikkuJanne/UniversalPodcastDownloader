BeforeAll {
    $script:RuntimeContextRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    function Invoke-RuntimeContextProbe {
        param([Parameter(Mandatory)][string]$Body, [switch]$AllowCaughtError)
        $powershell = [Management.Automation.PowerShell]::Create()
        try {
            $null = $powershell.AddScript("param(`$repository)`r`n" + $Body).AddArgument($script:RuntimeContextRepository)
            $output = @($powershell.Invoke())
            if (-not $AllowCaughtError) { $powershell.HadErrors | Should -BeFalse }
            $output.Count | Should -Be 1
            return $output[0]
        }
        finally { $powershell.Dispose() }
    }
}

Describe 'A048 per-invocation runtime context' -Tag 'Unit', 'A048' {
    It 'desired import preserves caller transport and page-limit variables' {
        $observation = Invoke-RuntimeContextProbe -Body @'
$script:PodcastTransportPolicy = [pscustomobject]@{Marker='owned-transport-sentinel'}
$script:PodcastMaxFeedPages = 37
$originalPolicy = $script:PodcastTransportPolicy
. (Join-Path $repository 'UniversalPodcastDownloader.ps1')
[pscustomobject]@{SamePolicy=[object]::ReferenceEquals($originalPolicy,$script:PodcastTransportPolicy);MaxPages=$script:PodcastMaxFeedPages}
'@
        $observation.SamePolicy | Should -BeTrue
        $observation.MaxPages | Should -Be 37
    }

    It 'desired fatal run preserves caller transport and page-limit variables' {
        $observation = Invoke-RuntimeContextProbe -Body @'
. (Join-Path $repository 'UniversalPodcastDownloader.ps1')
function Write-Host {}
function Get-PodcastHttpClient { throw 'A runtime-context unit must not create a network client.' }
$script:PodcastTransportPolicy = [pscustomobject]@{Marker='owned-transport-sentinel'}
$script:PodcastMaxFeedPages = 37
$originalPolicy = $script:PodcastTransportPolicy
$result = Invoke-PodcastRun -NonInteractive -WhatIf -MaxAttempts 1 -MaxFeedPages 2
[pscustomobject]@{ExitCode=$result.ExitCode;SamePolicy=[object]::ReferenceEquals($originalPolicy,$script:PodcastTransportPolicy);MaxPages=$script:PodcastMaxFeedPages}
'@
        $observation.ExitCode | Should -Be 1
        $observation.SamePolicy | Should -BeTrue
        $observation.MaxPages | Should -Be 37
    }

    It 'keeps caller preferences and standalone defaults after a <Outcome> run' -ForEach @(
        @{Outcome='completed preview';ExpectedExit=0}, @{Outcome='fatal';ExpectedExit=1}, @{Outcome='cancelled';ExpectedExit=130}
    ) {
        $body = @'
. (Join-Path $repository 'UniversalPodcastDownloader.ps1')
function Write-Host {}
function Read-Host { throw 'A runtime-context unit must not prompt.' }
function Get-PodcastHttpClient { throw 'A runtime-context unit must not create a network client.' }
function Invoke-PodcastMetadataRequest {
    param($Uri,$Policy)
    $script:ObservedPolicies.Add($Policy)
    if ($script:CancelMetadata) { throw [OperationCanceledException]::new('owned-cancellation-canary') }
    [pscustomobject]@{Content='<rss><channel><title>Owned context fixture</title><item><title>Episode</title><guid>context-unit</guid><enclosure url="https://media.example.invalid/context.mp3" type="audio/mpeg" /></item></channel></rss>';FinalUri=[uri]$Uri;ContentType='application/rss+xml'}
}
$script:OriginalCatalogueResolver = (Get-Command Resolve-PodcastCatalogue).ScriptBlock
function Resolve-PodcastCatalogue {
    param($InitialResolution,$MaxPages,$Policy)
    $script:ObservedPageLimits.Add([int]$MaxPages)
    & $script:OriginalCatalogueResolver @PSBoundParameters
}
$script:ObservedPolicies = [Collections.Generic.List[object]]::new()
$script:ObservedPageLimits = [Collections.Generic.List[int]]::new()
$script:PodcastTransportPolicy = [pscustomobject]@{Marker='owned-transport-sentinel'}
$script:PodcastMaxFeedPages = 37
$originalPolicy = $script:PodcastTransportPolicy
$ErrorActionPreference = 'Continue'
$ConfirmPreference = 'High'
$ProgressPreference = 'SilentlyContinue'
$WhatIfPreference = $false
$LASTEXITCODE = 19
$script:CancelMetadata = '__OUTCOME__' -eq 'cancelled'
$runArguments = @{NonInteractive=$true;WhatIf=$true;OutputPath='__OUTPUT__';MaxAttempts=1;MaxFeedPages=2}
if ('__OUTCOME__' -ne 'fatal') { $runArguments.FeedUrl='https://feed.example.invalid/context.xml' }
$result = Invoke-PodcastRun @runArguments
$runAttempts = if ($script:ObservedPolicies.Count) { $script:ObservedPolicies[0].MaxAttempts } else { $null }
$runPages = if ($script:ObservedPageLimits.Count) { $script:ObservedPageLimits[0] } else { $null }
$script:CancelMetadata = $false
$null = Invoke-PodcastWebRequest -Uri 'https://feed.example.invalid/standalone.xml'
$standaloneAttempts = $script:ObservedPolicies[$script:ObservedPolicies.Count-1].MaxAttempts
$null = Resolve-PodcastItems -Feeds @('https://feed.example.invalid/default-pages.xml')
[pscustomobject]@{ExitCode=$result.ExitCode;Complete=$result.Complete;Preview=$result.Preview;SamePolicy=[object]::ReferenceEquals($originalPolicy,$script:PodcastTransportPolicy);MaxPages=$script:PodcastMaxFeedPages;RunAttempts=$runAttempts;RunPages=$runPages;StandaloneAttempts=$standaloneAttempts;StandalonePages=$script:ObservedPageLimits[$script:ObservedPageLimits.Count-1];ErrorPreference=[string]$ErrorActionPreference;ConfirmPreference=[string]$ConfirmPreference;ProgressPreference=[string]$ProgressPreference;WhatIfPreference=[bool]$WhatIfPreference;LastExit=$LASTEXITCODE}
'@
        $ownedOutput = Join-Path $TestDrive 'context-preview-only'
        $body = $body.Replace('__OUTCOME__', $Outcome).Replace('__OUTPUT__', $ownedOutput.Replace("'", "''"))
        $observation = Invoke-RuntimeContextProbe -Body $body
        $observation.ExitCode | Should -Be $ExpectedExit
        $observation.Preview | Should -BeTrue
        $observation.SamePolicy | Should -BeTrue
        $observation.MaxPages | Should -Be 37
        $observation.StandaloneAttempts | Should -Be 3
        $observation.StandalonePages | Should -Be 20
        $observation.ErrorPreference | Should -BeExactly 'Continue'
        $observation.ConfirmPreference | Should -BeExactly 'High'
        $observation.ProgressPreference | Should -BeExactly 'SilentlyContinue'
        $observation.WhatIfPreference | Should -BeFalse
        $observation.LastExit | Should -Be 19
        Test-Path -LiteralPath $ownedOutput | Should -BeFalse
        if ($Outcome -eq 'completed preview') {
            $observation.Complete | Should -BeTrue
            $observation.RunAttempts | Should -Be 1
            $observation.RunPages | Should -Be 2
        }
        elseif ($Outcome -eq 'cancelled') { $observation.RunAttempts | Should -Be 1 }
    }

    It 'desired guided metadata cancellation propagates before another input prompt' {
        $observation = Invoke-RuntimeContextProbe -AllowCaughtError -Body @'
. (Join-Path $repository 'UniversalPodcastDownloader.ps1')
function Write-Host {}
function Get-PodcastHttpClient { throw 'A guided runtime-context unit must not create a network client.' }
$script:InputCalls = 0
$script:MetadataCalls = 0
function Read-Host {
    $script:InputCalls++
    if ($script:InputCalls -gt 1) { throw [InvalidOperationException]::new('A cancelled metadata request must not prompt again.') }
    'https://feed.example.invalid/guided-cancel.xml'
}
function Invoke-PodcastMetadataRequest {
    $script:MetadataCalls++
    throw [OperationCanceledException]::new('owned-guided-cancellation-canary')
}
$failure = $null
$source = $null
try { $source = Get-FeedUrlInteractive }
catch { $failure = $_ }
[pscustomobject]@{Cancelled=(Test-PodcastCancellation -ErrorObject $failure);InputCalls=$script:InputCalls;MetadataCalls=$script:MetadataCalls;NoSource=($null -eq $source)}
'@
        $observation.Cancelled | Should -BeTrue
        $observation.InputCalls | Should -Be 1
        $observation.MetadataCalls | Should -Be 1
        $observation.NoSource | Should -BeTrue
    }
}
