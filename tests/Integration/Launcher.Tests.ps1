BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')

    function New-UpdLauncherFixture {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates synthetic launcher files only inside the marked owned integration root.')]
        [CmdletBinding()]
        param([Parameter(Mandatory)]$Context, [switch]$MissingScript)

        $folder = Join-Path $Context.Root ('Launcher space ' + [char]0x00E4 + ' & (fixture)!')
        $null = [IO.Directory]::CreateDirectory($folder)
        $launcher = Join-Path $folder 'UniversalPodcastDownloader.bat'
        [IO.File]::Copy((Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.bat'), $launcher)
        if (-not $MissingScript) {
            $spy = @'
param([string]$Payload, [int]$ExitCode = 0, [switch]$NonInteractive)
$observation = [pscustomobject]@{
    Payload = $Payload
    ExitCode = $ExitCode
    NonInteractive = [bool]$NonInteractive
    Launcher = $env:UPD_LAUNCHER
    ParameterCount = $PSBoundParameters.Count
    Engine = $PSVersionTable.PSVersion.ToString()
}
[IO.File]::WriteAllText($env:UPD_LAUNCHER_OBSERVATION, ($observation | ConvertTo-Json -Compress), [Text.UTF8Encoding]::new($false))
Write-Host 'Synthetic launcher child completed.'
exit $ExitCode
'@
            [IO.File]::WriteAllText((Join-Path $folder 'UniversalPodcastDownloader.ps1'), $spy)
        }
        return $launcher
    }

    function Invoke-UpdLauncherFixture {
        param(
            [Parameter(Mandatory)]$Context,
            [Parameter(Mandatory)][string]$Launcher,
            [string]$Arguments = '',
            [switch]$InheritedErrorLevel
        )

        # Raw command text is limited to synthetic test literals. The outer
        # quotes are CMD /S /C framing; each data argument keeps caller quoting.
        $start = [Diagnostics.ProcessStartInfo]::new()
        $start.FileName = Join-Path $env:WINDIR 'System32/cmd.exe'
        $start.Arguments = '/d /s /v:off /c ""' + $Launcher + '" ' + $Arguments + '"'
        $start.WorkingDirectory = $Context.Root
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $start.RedirectStandardInput = $true
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        foreach ($name in @('LOCALAPPDATA', 'USERPROFILE', 'TEMP', 'TMP')) {
            $start.EnvironmentVariables[$name] = $Context.Root
        }
        $start.EnvironmentVariables['PSModulePath'] = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules'
        $start.EnvironmentVariables['UPD_LAUNCHER'] = 'caller-value'
        $observationPath = Join-Path $Context.Root 'launcher-observation.json'
        $start.EnvironmentVariables['UPD_LAUNCHER_OBSERVATION'] = $observationPath
        if ($InheritedErrorLevel) { $start.EnvironmentVariables['ERRORLEVEL'] = '77' }
        $process = [Diagnostics.Process]::new()
        $process.StartInfo = $start
        if (-not $process.Start()) { throw 'Could not start the owned launcher fixture.' }
        $owned = [pscustomobject]@{
            Process = $process
            Output = $process.StandardOutput.ReadToEndAsync()
            ErrorOutput = $process.StandardError.ReadToEndAsync()
        }
        $Context.Processes.Add($owned)
        $exitedWithoutInput = $process.WaitForExit(15000)
        # A flawed launcher can still pause after its child exits. Release only
        # this tracked fixture's input, then let normal owned cleanup handle it.
        if (-not $exitedWithoutInput) { $process.StandardInput.WriteLine('') }
        $process.StandardInput.Close()
        if (-not $process.WaitForExit(5000)) { throw 'Owned launcher fixture did not exit after input release.' }
        [pscustomobject]@{
            ExitCode = $process.ExitCode
            ExitedWithoutInput = $exitedWithoutInput
            Stdout = $owned.Output.Result
            Stderr = $owned.ErrorOutput.Result
            Observation = $(if ([IO.File]::Exists($observationPath)) { [IO.File]::ReadAllText($observationPath) | ConvertFrom-Json } else { $null })
        }
    }
}

Describe 'A040/A041 real Windows batch launcher plumbing' {
    BeforeEach {
        $script:launcherContext = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
    }

    AfterEach {
        if ($script:launcherContext) { Remove-UpdIntegrationContext -Context $script:launcherContext }
    }

    It 'A041 launcher forwarding keeps a PowerShell expression literal and returns without a pause' {
        $launcher = New-UpdLauncherFixture -Context $script:launcherContext
        $canary = Join-Path $script:launcherContext.Root 'must-not-execute.txt'
        $payload = "`$(Set-Content -LiteralPath '$canary' -Value injected); & Write-Host private-canary!"
        $run = Invoke-UpdLauncherFixture -Context $script:launcherContext -Launcher $launcher -Arguments ('-Payload "' + $payload + '" -ExitCode 2 -NonInteractive')
        $run.ExitCode | Should -Be 2
        $run.ExitedWithoutInput | Should -BeTrue
        $run.Observation.Payload | Should -BeExactly $payload
        $run.Observation.NonInteractive | Should -BeTrue
        $run.Observation.Launcher | Should -BeExactly '1'
        $run.Stdout | Should -Not -Match 'Press any key|Done\. You can close'
        Test-Path -LiteralPath $canary | Should -BeFalse
    }

    It 'A040 preserves native child exit code <Code> without a CMD pause' -TestCases @(
        @{ Code = 0 }
        @{ Code = 1 }
        @{ Code = 2 }
        @{ Code = 130 }
    ) {
        param($Code)
        $launcher = New-UpdLauncherFixture -Context $script:launcherContext
        $run = Invoke-UpdLauncherFixture -Context $script:launcherContext -Launcher $launcher -Arguments ('-ExitCode ' + $Code)
        $run.ExitCode | Should -Be $Code
        $run.ExitedWithoutInput | Should -BeTrue
        $run.Observation.ExitCode | Should -Be $Code
        $run.Observation.Engine | Should -Match '^5\.1\.'
        $run.Stdout | Should -Not -Match 'Press any key|Done\. You can close'
    }

    It 'A041 argument-free launch supplies the process-local guided marker' {
        $launcher = New-UpdLauncherFixture -Context $script:launcherContext
        $run = Invoke-UpdLauncherFixture -Context $script:launcherContext -Launcher $launcher
        $run.ExitCode | Should -Be 0
        $run.Observation.ParameterCount | Should -Be 0
        $run.Observation.Launcher | Should -BeExactly '1'
        $run.Stdout | Should -Not -Match 'Done\. You can close'
    }

    It 'A041 missing colocated script is a visible fatal error without a false completion or pause' {
        $launcher = New-UpdLauncherFixture -Context $script:launcherContext -MissingScript
        $run = Invoke-UpdLauncherFixture -Context $script:launcherContext -Launcher $launcher -Arguments '-NonInteractive'
        $run.ExitCode | Should -Be 1
        $run.ExitedWithoutInput | Should -BeTrue
        ($run.Stdout + $run.Stderr) | Should -Match 'ERROR: UniversalPodcastDownloader\.ps1 is missing next to the launcher\.'
        $run.Stdout | Should -Not -Match 'Press any key|Done\. You can close|\[OK\]'
        $run.Observation | Should -BeNullOrEmpty
    }

    It 'A040 an inherited ERRORLEVEL variable cannot replace the actual child process code' {
        $launcher = New-UpdLauncherFixture -Context $script:launcherContext
        $run = Invoke-UpdLauncherFixture -Context $script:launcherContext -Launcher $launcher -Arguments '-ExitCode 2' -InheritedErrorLevel
        $run.ExitCode | Should -Be 2
        $run.ExitedWithoutInput | Should -BeTrue
    }
}
