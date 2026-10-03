BeforeAll {
    $script:DiagnosticLifecycleRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $PSScriptRoot '../../src/PathSafety.ps1')
    . (Join-Path $PSScriptRoot '../../src/Diagnostics.ps1')
}

Describe 'A048 diagnostic correlation lifecycle' -Tag 'Unit', 'A048' {
    BeforeEach { $null = Initialize-PodcastDiagnostics -Preview }
    AfterEach { Close-PodcastDiagnostics }

    It 'desired close releases exact URL correlation keys while retaining the original dictionary' {
        $null = Get-PodcastSafeUrl -Url 'https://feed.example.invalid/private-path?token=correlation-canary'
        $retainedMap = $script:PodcastDiagnosticRequestIds
        $retainedMap.Count | Should -Be 1
        Close-PodcastDiagnostics
        $retainedMap.Count | Should -Be 0
        [object]::ReferenceEquals($retainedMap, $script:PodcastDiagnosticRequestIds) | Should -BeTrue
    }

    It 'desired close releases correlation keys even when writer disposal fails' {
        $null = Get-PodcastSafeUrl -Url 'https://feed.example.invalid/private-path?token=disposal-canary'
        $retainedMap = $script:PodcastDiagnosticRequestIds
        $writer = [pscustomobject]@{ DisposeCalls = 0 }
        $writer | Add-Member -MemberType ScriptMethod -Name Dispose -Value {
            $this.DisposeCalls++
            throw [IO.IOException]::new('owned-disposal-canary')
        }
        $script:PodcastDiagnostics.Writer = $writer
        Close-PodcastDiagnostics
        $writer.DisposeCalls | Should -Be 1
        $retainedMap.Count | Should -Be 0
        [object]::ReferenceEquals($retainedMap, $script:PodcastDiagnosticRequestIds) | Should -BeTrue
        $script:PodcastDiagnostics.Writer = $null
    }

    It 'starts a new correlation after close without discarding safe events or the diagnostic context' {
        $url = 'https://feed.example.invalid/private-path?token=new-correlation-canary'
        $first = Get-PodcastSafeUrl -Url $url
        Write-PodcastDiagnostic -Message 'An authored safe event.'
        $retainedContext = $script:PodcastDiagnostics
        $eventCount = $retainedContext.Events.Count
        Close-PodcastDiagnostics
        $second = Get-PodcastSafeUrl -Url $url
        $second | Should -Not -Be $first
        $second | Should -Match '^feed\.example\.invalid \[url [a-f0-9]{32}\]$'
        [object]::ReferenceEquals($retainedContext, $script:PodcastDiagnostics) | Should -BeTrue
        $retainedContext.Events.Count | Should -Be $eventCount
        $script:PodcastDiagnosticRequestIds.Count | Should -Be 1
    }

    It 'keeps import and close without initialized diagnostics free of new script state' {
        $powershell = [Management.Automation.PowerShell]::Create()
        try {
            $null = $powershell.AddScript(@'
param($repository)
$names = @('PodcastDiagnostics','PodcastDiagnosticPreview','PodcastDiagnosticRequestIds')
. (Join-Path $repository 'src/Diagnostics.ps1')
Close-PodcastDiagnostics
[pscustomobject]@{Created=@($names | Where-Object { $null -ne $ExecutionContext.SessionState.PSVariable.Get('script:' + $_) }).Count}
'@).AddArgument($script:DiagnosticLifecycleRepository)
            $output = @($powershell.Invoke())
            $powershell.HadErrors | Should -BeFalse
            $output.Count | Should -Be 1
            $output[0].Created | Should -Be 0
        }
        finally { $powershell.Dispose() }
    }

    It 'allows a strict safe export after close while releasing the raw correlation map' {
        $ownedRoot = Join-Path $TestDrive 'post-close-log'
        $context = Initialize-PodcastDiagnostics -Roots $ownedRoot
        $null = Get-PodcastSafeUrl -Url 'https://feed.example.invalid/private-path?token=export-canary'
        Write-PodcastDiagnostic -Message 'An authored safe export event.'
        Close-PodcastDiagnostics
        $context.Writer | Should -BeNullOrEmpty
        $script:PodcastDiagnosticRequestIds.Count | Should -Be 0
        $exportPath = Join-Path $TestDrive 'closed-export.json'
        Export-PodcastDiagnostics -Path $exportPath -Confirm:$false
        $json = [IO.File]::ReadAllText($exportPath)
        $exported = $json | ConvertFrom-Json
        $exported.schema_version | Should -Be 1
        $exported.run_id | Should -BeExactly $context.RunId
        @($exported.events).Count | Should -Be $context.Events.Count
        $json | Should -Not -Match 'export-canary|private-path|feed.example.invalid|authored safe export event|Writer|Path'
    }

    It 'retains the preview export prohibition after closing correlations' {
        $null = Get-PodcastSafeUrl -Url 'https://feed.example.invalid/private-path?token=preview-canary'
        Close-PodcastDiagnostics
        $script:PodcastDiagnosticRequestIds.Count | Should -Be 0
        $exportPath = Join-Path $TestDrive 'preview-export.json'
        Export-PodcastDiagnostics -Path $exportPath -Confirm:$false
        Test-Path -LiteralPath $exportPath | Should -BeFalse
        $script:PodcastDiagnostics.Preview | Should -BeTrue
        $script:PodcastDiagnosticPreview | Should -BeTrue
    }
}
