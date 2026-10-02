Describe 'A004: importing the actual downloader script' -Tag 'Unit', 'A004' {
    BeforeAll {
        $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    }

    It 'defines helpers without prompts, requests, output, writes or preference changes' {
        Mock Read-Host { throw 'Import attempted to prompt.' }
        Mock Invoke-WebRequest { throw 'Import attempted a network request.' }
        Mock Add-Type { throw 'Import attempted to initialize the media client.' }
        Mock New-Item { throw 'Import attempted to create an item.' }
        Mock Set-Content { throw 'Import attempted to write a file.' }
        Mock Add-Content { throw 'Import attempted to append a file.' }
        Mock Write-Host {}
        Mock Write-Progress {}

        $previousErrorPreference = $ErrorActionPreference
        $previousProgressPreference = $global:ProgressPreference
        $previousLogVariable = Get-Variable LogFile -Scope Script -ErrorAction SilentlyContinue
        $previousLogValue = if ($previousLogVariable) { $previousLogVariable.Value } else { $null }
        $originalChildren = @(Get-ChildItem -LiteralPath $TestDrive -Force -Recurse).Count
        $unusedOutput = Join-Path $TestDrive 'import-must-not-create-this'
        try {
            $ErrorActionPreference = 'Continue'
            $global:ProgressPreference = 'SilentlyContinue'
            $script:LogFile = 'import-log-sentinel'

            $output = @(. $script:DownloaderPath -OutputPath $unusedOutput)

            $output.Count | Should -Be 0
            $ErrorActionPreference | Should -Be 'Continue'
            $global:ProgressPreference | Should -Be 'SilentlyContinue'
            $script:LogFile | Should -Be 'import-log-sentinel'
            (Get-Command Resolve-PodcastItems).CommandType | Should -Be 'Function'
            (Get-Command Get-EpisodeData).CommandType | Should -Be 'Function'
            (Get-Command New-EpisodeFileName).CommandType | Should -Be 'Function'
            Test-Path -LiteralPath $unusedOutput | Should -BeFalse
            @(Get-ChildItem -LiteralPath $TestDrive -Force -Recurse).Count | Should -Be $originalChildren
            Should -Invoke Read-Host -Times 0 -Exactly -Scope It
            Should -Invoke Invoke-WebRequest -Times 0 -Exactly -Scope It
            Should -Invoke Add-Type -Times 0 -Exactly -Scope It
            Should -Invoke New-Item -Times 0 -Exactly -Scope It
            Should -Invoke Set-Content -Times 0 -Exactly -Scope It
            Should -Invoke Add-Content -Times 0 -Exactly -Scope It
            Should -Invoke Write-Host -Times 0 -Exactly -Scope It
            Should -Invoke Write-Progress -Times 0 -Exactly -Scope It
        }
        finally {
            $ErrorActionPreference = $previousErrorPreference
            $global:ProgressPreference = $previousProgressPreference
            if ($previousLogVariable) {
                $script:LogFile = $previousLogValue
            }
            else {
                Remove-Variable LogFile -Scope Script -ErrorAction SilentlyContinue
            }
        }
    }
}
