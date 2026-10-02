Describe 'A005: development runner failure handling' -Tag 'Unit', 'A005' {
    BeforeAll {
        $script:RunnerRepo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $script:RunnerTools = Join-Path $script:RunnerRepo '.dev-tools/Modules'
        $script:RunnerEngine = (Get-Process -Id $PID).Path

        function Invoke-UpdRunnerCheck {
            param([string]$ScriptName, [string[]]$Arguments)
            $start = New-Object System.Diagnostics.ProcessStartInfo
            $start.FileName = $script:RunnerEngine
            $values = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
                (Join-Path $script:RunnerRepo "scripts/$ScriptName")) + $Arguments
            $start.Arguments = ($values | ForEach-Object { '"' + $_ + '"' }) -join ' '
            $start.UseShellExecute = $false
            $start.CreateNoWindow = $true
            $start.RedirectStandardOutput = $true
            $start.RedirectStandardError = $true
            $start.EnvironmentVariables.Remove('GITHUB_STEP_SUMMARY')
            if ($PSVersionTable.PSEdition -eq 'Desktop') {
                $start.EnvironmentVariables['PSModulePath'] = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules'
            }
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $start
            try {
                [void]$process.Start()
                $stdout = $process.StandardOutput.ReadToEndAsync()
                $stderr = $process.StandardError.ReadToEndAsync()
                if (-not $process.WaitForExit(30000)) {
                    $process.Kill()
                    $process.WaitForExit()
                    throw 'Owned runner check timed out.'
                }
                # Error formatting wraps at different widths on local/hosted PS5.
                # Assert message content and exit status independently of layout.
                [pscustomobject]@{ Code = $process.ExitCode; Text = [regex]::Replace($stdout.Result + $stderr.Result, '\s+', ' ') }
            }
            finally { $process.Dispose() }
        }
    }

    BeforeEach {
        $script:RunnerFixture = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory((Join-Path $script:RunnerFixture 'Unit'))
    }

    It 'fails when pinned Pester is missing instead of using a global version' {
        $result = Invoke-UpdRunnerCheck Test.ps1 @('-Suite', 'Unit', '-TestRoot', $script:RunnerFixture,
            '-ToolsPath', (Join-Path $TestDrive 'missing-modules'))
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'Missing\s+pinned\s+Pester\s+5\.7\.1'
    }

    It 'fails when a requested suite contains no test files' {
        $result = Invoke-UpdRunnerCheck Test.ps1 @('-Suite', 'Unit', '-TestRoot', $script:RunnerFixture, '-ToolsPath', $script:RunnerTools)
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'No tests found for suite Unit'
    }

    It 'reports <Label> and propagates its exit status' -TestCases @(
        @{ Label = 'a passing selection'; Body = "It 'sample' { 1 | Should -Be 1 }"; Code = 0; Counts = 'passed=1 failed=0 skipped=0' },
        @{ Label = 'a failing selection'; Body = "It 'sample' { 1 | Should -Be 2 }"; Code = 1; Counts = 'passed=0 failed=1 skipped=0' },
        @{ Label = 'an entirely skipped selection'; Body = "It 'sample' -Skip { 1 | Should -Be 1 }"; Code = 1; Counts = 'passed=0 failed=0 skipped=1' }
    ) {
        param($Label, $Body, $Code, $Counts)
        $Label | Should -Not -BeNullOrEmpty
        Set-Content -LiteralPath (Join-Path $script:RunnerFixture 'Unit/Sample.Tests.ps1') -Value ("Describe 'miniature runner fixture' { " + $Body + ' }')
        $result = Invoke-UpdRunnerCheck Test.ps1 @('-Suite', 'Unit', '-TestRoot', $script:RunnerFixture, '-ToolsPath', $script:RunnerTools)
        $result.Code | Should -Be $Code -Because $result.Text
        $result.Text | Should -Match $Counts
    }

    It 'fails when a test filter selects nothing' {
        Set-Content -LiteralPath (Join-Path $script:RunnerFixture 'Unit/Sample.Tests.ps1') -Value "Describe 'fixture' { It 'sample' { 1 | Should -Be 1 } }"
        $result = Invoke-UpdRunnerCheck Test.ps1 @('-Suite', 'Unit', '-TestRoot', $script:RunnerFixture,
            '-ToolsPath', $script:RunnerTools, '-Filter', '*no-matching-test*')
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'No tests executed'
    }

    It 'fails All when a whole suite is <Label>' -TestCases @(
        @{ Label = 'empty despite a test filename'; Source = '# empty discovery fixture'; ErrorText = 'No tests discovered in suite Integration' },
        @{ Label = 'entirely skipped'; Source = "Describe 'integration fixture' { It 'skipped' -Skip { 1 | Should -Be 1 } }"; ErrorText = 'No tests executed in suite Integration' }
    ) {
        param($Label, $Source, $ErrorText)
        $Label | Should -Not -BeNullOrEmpty
        Set-Content -LiteralPath (Join-Path $script:RunnerFixture 'Unit/Sample.Tests.ps1') -Value "Describe 'unit fixture' { It 'passes' { 1 | Should -Be 1 } }"
        [void][IO.Directory]::CreateDirectory((Join-Path $script:RunnerFixture 'Integration'))
        Set-Content -LiteralPath (Join-Path $script:RunnerFixture 'Integration/Sample.Tests.ps1') -Value $Source
        $result = Invoke-UpdRunnerCheck Test.ps1 @('-Suite', 'All', '-TestRoot', $script:RunnerFixture, '-ToolsPath', $script:RunnerTools)
        $result.Code | Should -Be 1 -Because $result.Text
        $result.Text | Should -Match $ErrorText
    }

    It 'fails when pinned PSScriptAnalyzer is missing' {
        $result = Invoke-UpdRunnerCheck Analyze.ps1 @('-ToolsPath', (Join-Path $TestDrive 'missing-modules'))
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'Missing\s+pinned\s+PSScriptAnalyzer\s+1\.24\.0'
    }

    It 'fails lint for a new warning outside the legacy baseline' {
        $file = Join-Path $TestDrive 'new-warning.ps1'
        Set-Content -LiteralPath $file -Value "Invoke-Expression 'synthetic-string-never-executed'"
        $result = Invoke-UpdRunnerCheck Analyze.ps1 @('-ToolsPath', $script:RunnerTools, '-Path', $file)
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'PSAvoidUsingInvokeExpression'
    }
}
