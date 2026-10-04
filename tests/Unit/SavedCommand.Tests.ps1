BeforeAll {
    . (Join-Path $PSScriptRoot '../../UniversalPodcastDownloader.ps1')
}

Describe 'A046/A047 saved command boundary and unchanged one-off routing' {
    BeforeEach {
        $script:commandCalls = [Collections.Generic.List[object]]::new()
        Mock Write-Host { }
        Mock Get-PodcastSavedShowConfigPath { Join-Path $TestDrive 'private/shows.json' }
        Mock Get-PodcastSavedShow { [pscustomobject]@{ Name='news'; FeedUrl='https://example.invalid/private?token=PRIVATE_SECRET'; OutputPath=$TestDrive; Mode='Custom'; CustomCount=7 } }
        Mock Invoke-PodcastRun {
            $script:commandCalls.Add([pscustomobject]@{ Feed=$FeedUrl; Mode=$Mode; Count=$CustomCount; Output=$OutputPath; Quiet=[bool]$NonInteractive; KeepAwake=[bool]$KeepAwake; Preview=[bool]$WhatIfPreference })
            $effectiveMode = if ($Mode) { $Mode } else { 'Latest' }
            New-PodcastRunResult -Mode $effectiveMode -Preview:$WhatIfPreference -Message 'Safe run result.'
        }
        Mock Save-PodcastSavedShow { [pscustomobject]@{ Changed=$false; Preview=$true; Name=$Name; Message='Saved-show configuration preview; no changes made.' } }
        Mock Remove-PodcastSavedShow { [pscustomobject]@{ Changed=$false; Preview=$true; Name=$Name; Message='Saved-show configuration preview; no changes made.' } }
        Mock Export-PodcastSavedShowConfig { [pscustomobject]@{ Changed=$false; Preview=$true; Message='Saved-show export preview; no changes made.' } }
        Mock Get-PodcastSavedShowList { return ,@() }
        Mock Invoke-PodcastBatchRun { New-PodcastBatchResult -Preview:$WhatIfPreference }
    }

    It 'does not read or resolve any saved settings for the ordinary one-off API route' {
        $result = @(Invoke-PodcastCommand -Options @{ FeedUrl='https://example.invalid/feed'; NonInteractive=$true; CustomCount=3; Mode='Custom' })
        $result.Count | Should -Be 1
        $result[0].Type | Should -Be 'Podcast.RunResult'
        $script:commandCalls.Count | Should -Be 1
        Should -Invoke Get-PodcastSavedShowConfigPath -Times 0 -Exactly
        Should -Invoke Get-PodcastSavedShow -Times 0 -Exactly
        Should -Invoke Save-PodcastSavedShow -Times 0 -Exactly
    }

    It 'rejects <Case> before config resolution, archive work or mutation' -TestCases @(
        @{ Case='two management operations'; Options=@{ SaveShow='news'; ListShows=$true; FeedUrl='https://example.invalid/feed' } },
        @{ Case='direct feed mixed with named show'; Options=@{ ShowName='news'; FeedUrl='https://example.invalid/feed' } },
        @{ Case='legacy mixed with batch'; Options=@{ Batch=$true; LegacyPath='D:\copied' } },
        @{ Case='management with a runtime request'; Options=@{ ListShows=$true; KeepAwake=$true } },
        @{ Case='config without a selector'; Options=@{ ConfigPath='D:\copied\shows.json'; FeedUrl='https://example.invalid/feed' } }
    ) {
        param($Options)
        $result = Invoke-PodcastCommand -Options $Options
        $result.ExitCode | Should -Be 1
        Should -Invoke Get-PodcastSavedShowConfigPath -Times 0 -Exactly
        Should -Invoke Get-PodcastSavedShow -Times 0 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
        Should -Invoke Save-PodcastSavedShow -Times 0 -Exactly
        Should -Invoke Invoke-PodcastBatchRun -Times 0 -Exactly
    }

    It 'requires one explicit name for a single saved-show request' {
        $result = Invoke-PodcastCommand -Options @{ ShowName=@('news','other'); ConfigPath=(Join-Path $TestDrive 'private/shows.json') }
        $result.ExitCode | Should -Be 1
        Should -Invoke Get-PodcastSavedShow -Times 0 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }

    It 'keeps explicitly disabled saved switches on the ordinary path without forwarding unknown parameters' {
        Mock Read-PodcastSavedShowConfig { throw 'Unexpected configuration access.' }
        $result = Invoke-PodcastCommand -Options @{ FeedUrl='https://example.invalid/feed'; NonInteractive=$true; Batch=$false; ListShows=$false }
        $result.ExitCode | Should -Be 0
        Should -Invoke Invoke-PodcastRun -Times 1 -Exactly -ParameterFilter { $FeedUrl -eq 'https://example.invalid/feed' -and $NonInteractive }
        Should -Invoke Get-PodcastSavedShowConfigPath -Times 0 -Exactly
        Should -Invoke Read-PodcastSavedShowConfig -Times 0 -Exactly
        Should -Invoke Invoke-PodcastBatchRun -Times 0 -Exactly
    }

    It 'honors explicit preview over inherited preference on rejected commands before configuration access' {
        $previousWhatIf = $WhatIfPreference
        try {
            $WhatIfPreference = $true
            foreach ($explicitPreview in @($true, $false)) {
                $result = Invoke-PodcastCommand -Options @{ Batch=$true; FeedUrl='https://example.invalid/private'; WhatIf=$explicitPreview }
                $result.ExitCode | Should -Be 1
                $result.Preview | Should -Be $explicitPreview
            }
        }
        finally { $WhatIfPreference = $previousWhatIf }
        Should -Invoke Get-PodcastSavedShowConfigPath -Times 0 -Exactly
        Should -Invoke Get-PodcastSavedShow -Times 0 -Exactly
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }

    It 'runs the stored selection without returning its private address or output path' {
        $result = Invoke-PodcastCommand -Options @{ ShowName='news'; NonInteractive=$true; KeepAwake=$true }
        $result.ExitCode | Should -Be 0
        $script:commandCalls[0].Mode | Should -Be 'Custom'
        $script:commandCalls[0].Count | Should -Be 7
        $script:commandCalls[0].Quiet | Should -BeTrue
        $script:commandCalls[0].KeepAwake | Should -BeTrue
        ($result | ConvertTo-Json -Depth 8) | Should -Not -Match 'PRIVATE_SECRET|example\.invalid|FeedUrl|OutputPath'
    }

    It 'uses a count-only override and retires a stored count for explicit Latest' {
        $null = Invoke-PodcastCommand -Options @{ ShowName='news'; CustomCount=2; NonInteractive=$true }
        $null = Invoke-PodcastCommand -Options @{ ShowName='news'; Mode='Latest'; NonInteractive=$true }
        $script:commandCalls[0].Mode | Should -Be 'Custom'
        $script:commandCalls[0].Count | Should -Be 2
        $script:commandCalls[1].Mode | Should -Be 'Latest'
        $script:commandCalls[1].Count | Should -Be 0
    }

    It 'returns stable zero, one and several safe list rows' {
        foreach ($size in @(0,1,3)) {
            $script:listRows = @(for ($i=0; $i -lt $size; $i++) { [pscustomobject]@{ Name=('show'+$i); Mode='Latest'; CustomCount=$null; FeedConfigured=$true; FeedUrl='PRIVATE_SECRET'; feed_protected='PRIVATE_CIPHER'; output_path='PRIVATE_PATH' } })
            Mock Get-PodcastSavedShowList { return ,$script:listRows }
            $result = @(Invoke-PodcastCommand -Options @{ ListShows=$true })
            $result.Count | Should -Be 1
            $result[0].Shows -is [array] | Should -BeTrue
            $result[0].Shows.Count | Should -Be $size
            ($result | ConvertTo-Json -Depth 8) | Should -Not -Match 'PRIVATE_SECRET|PRIVATE_CIPHER|PRIVATE_PATH'
        }
    }

    It 'passes a count-only save without inventing an explicit Latest argument' {
        $result = Invoke-PodcastCommand -Options @{ SaveShow='news'; FeedUrl='https://example.invalid/feed'; CustomCount=4; OutputPath=$TestDrive; WhatIf=$true }
        $result.Type | Should -Be 'Podcast.ConfigResult'
        $result.Preview | Should -BeTrue
        $result.Changed | Should -BeFalse
        Should -Invoke Save-PodcastSavedShow -Times 1 -Exactly -ParameterFilter { $CustomCount -eq 4 -and $FeedUrl -eq 'https://example.invalid/feed' }
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }

    It 'passes explicit subset and confirmation to batch without treating them as runtime overrides' {
        $result = Invoke-PodcastCommand -Options @{ Batch=$true; ShowName=@('second','first'); CustomCount=2; WhatIf=$true; Confirm=$false }
        $result.Type | Should -Be 'Podcast.BatchResult'
        Should -Invoke Invoke-PodcastBatchRun -Times 1 -Exactly -ParameterFilter { $ShowName[0] -eq 'second' -and $ShowName[1] -eq 'first' -and $RunOptions.CustomCount -eq 2 -and -not $RunOptions.Contains('Confirm') -and -not $RunOptions.Contains('WhatIf') }
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }

    It 'reports fixed credential failure and preserves typed cancellation without exposing exception strings' {
        foreach ($cancelled in @($false,$true)) {
            $script:commandCancelled = $cancelled
            Mock Get-PodcastSavedShow {
                if ($script:commandCancelled) { throw [OperationCanceledException]::new('PRIVATE_SECRET') }
                throw 'Saved-show credentials are unavailable for the current Windows user.'
            }
            $result = Invoke-PodcastCommand -Options @{ ShowName='news'; NonInteractive=$true }
            $result.ExitCode | Should -Be $(if ($cancelled) { 130 } else { 1 })
            $result.Message | Should -Not -Match 'PRIVATE_SECRET'
            Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
        }
    }

    It 'desired dispatcher renderer guard preserves primary results after a <RendererFailure> display failure' -ForEach @(
        @{ RendererFailure='ordinary' }, @{ RendererFailure='typed' }
    ) {
        $script:commandRenderCancelled = $RendererFailure -eq 'typed'
        $script:commandRenderCalls = 0
        Mock Write-Host {
            $script:commandRenderCalls++
            if ($script:commandRenderCancelled) { throw [OperationCanceledException]::new('PRIVATE_RENDER_SECRET') }
            throw [IO.IOException]::new('PRIVATE_RENDER_SECRET')
        }
        Mock Get-PodcastSavedShow { throw [OperationCanceledException]::new('PRIVATE_PRIMARY_SECRET') }
        foreach ($case in @(
            @{ Options=@{ Batch=$true; LegacyPath='owned-unused' }; Type='Podcast.BatchResult'; Exit=1 },
            @{ Options=@{ ListShows=$true; RemoveShow='news' }; Type='Podcast.ConfigResult'; Exit=1 },
            @{ Options=@{ ShowName='news'; NonInteractive=$true }; Type='Podcast.RunResult'; Exit=130 }
        )) {
            $result = @(Invoke-PodcastCommand -Options $case.Options)
            $result.Count | Should -Be 1
            $result[0].Type | Should -Be $case.Type
            $result[0].ExitCode | Should -Be $case.Exit
            ($result | ConvertTo-Json -Depth 8) | Should -Not -Match 'PRIVATE_RENDER_SECRET|PRIVATE_PRIMARY_SECRET'
        }
        $script:commandRenderCalls | Should -Be 3
        Should -Invoke Invoke-PodcastRun -Times 0 -Exactly
    }
}
