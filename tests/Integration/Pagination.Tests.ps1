BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    . (Join-Path $repositoryRoot 'UniversalPodcastDownloader.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned loopback feed-pagination integration fixtures.' }
    $script:paginationPython = $python.Source
    $sample = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
    $script:paginationHash = (Get-FileHash -LiteralPath $sample -Algorithm SHA256).Hash
    $script:paginationLength = (Get-Item -LiteralPath $sample).Length

    function Get-UpdPaginationArchive {
        param([Parameter(Mandatory)][string]$OutputPath)
        $states = @(Get-ChildItem -LiteralPath $OutputPath -Recurse -Force -File -Filter 'state.json')
        $states.Count | Should -Be 1
        $archiveRoot = Split-Path $states[0].DirectoryName -Parent
        [pscustomobject]@{ Root = $archiveRoot; State = Read-PodcastHistory -Root $archiveRoot }
    }

    function Assert-UpdPaginationIncomplete {
        param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)][int]$Downloads)
        $Run.Result.Succeeded | Should -BeFalse
        $Run.ExitCode | Should -Be 1
        $Run.Stdout | Should -Match ('Downloaded\s+: ' + $Downloads)
        $Run.Stdout | Should -Match 'Feed catalogue incomplete:'
        $Run.Result.ErrorMessage | Should -Match 'Feed catalogue incomplete; accessible selected episodes were processed, but advertised pages remain unresolved\.'
        $Run.Stdout | Should -Not -Match 'Run completed|\[OK\]'
    }
}

Describe 'A037/A038 explicit feed pagination at real pipeline boundaries' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:paginationPython
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'deduplicates accessible entries and reports a cycle after processing those entries' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/page-1.xml'
        @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 2
        Assert-UpdPaginationIncomplete -Run $run -Downloads 2
        $archive = Get-UpdPaginationArchive -OutputPath $run.OutputPath
        @($archive.State.episodes).Count | Should -Be 2
        foreach ($record in $archive.State.episodes) {
            $record.status | Should -Be 'transfer_verified'
            (Get-FileHash -LiteralPath (Join-Path $archive.Root $record.relative_path) -Algorithm SHA256).Hash | Should -Be $script:paginationHash
        }
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/page-1.xml' | Should -Be 1
        $stats.'/feeds/page-2.xml' | Should -Be 1
        $stats.'/media/ok.mp3' | Should -Be 2
    }

    It 'previews a cyclic catalogue with a visible warning and no media requests or persistent writes' {
        $preview = Invoke-UpdIntegrationWorker -Context $context -Action Preview -FeedPath '/feeds/page-1.xml'
        $preview.Result.Succeeded | Should -BeTrue -Because ($preview.Stdout + $preview.Stderr + $preview.Result.ErrorMessage)
        $preview.Stdout | Should -Match 'Feed catalogue incomplete:'
        Test-Path -LiteralPath $preview.OutputPath | Should -BeFalse
        @(Get-ChildItem -LiteralPath $context.Root -Recurse -Force -File -Filter '*.log').Count | Should -Be 0
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/page-1.xml' | Should -Be 1
        $stats.'/feeds/page-2.xml' | Should -Be 1
        $stats.'/media/ok.mp3' | Should -BeNullOrEmpty
    }

    It 'deduplicates <Kind> <Relation> pages, preserves original bytes and skips the same recorded files on repeat' -TestCases @(
        @{ Kind = 'RSS'; Relation = 'next'; FeedPath = '/feeds/pagination-rss.xml'; PagePath = '/feeds/pagination-rss-page-2.xml'; IdentitySource = 'rss-guid' }
        @{ Kind = 'Atom'; Relation = 'next'; FeedPath = '/feeds/pagination-atom.xml'; PagePath = '/feeds/pagination-atom-page-2.xml'; IdentitySource = 'atom-id' }
        @{ Kind = 'RSS'; Relation = 'prev-archive'; FeedPath = '/feeds/pagination-archive.xml'; PagePath = '/feeds/pagination-rss-page-2.xml'; IdentitySource = 'rss-guid' }
    ) {
        param($Kind, $Relation, $FeedPath, $PagePath, $IdentitySource)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath
        $first.Result.Succeeded | Should -BeTrue -Because ($Kind + ' ' + $Relation + ': ' + $first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $first.Stdout | Should -Match 'Downloaded\s+: 3'
        $first.Stdout | Should -Not -Match 'Feed catalogue incomplete:'
        $archive = Get-UpdPaginationArchive -OutputPath $first.OutputPath
        @($archive.State.episodes).Count | Should -Be 3
        $before = @($archive.State.episodes | Sort-Object episode_id | ForEach-Object {
            $_.status | Should -Be 'transfer_verified'
            $_.identity_source | Should -Be $IdentitySource
            $file = Get-Item -LiteralPath (Join-Path $archive.Root $_.relative_path)
            $file.Length | Should -Be $script:paginationLength
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be $script:paginationHash
            $_.episode_id + '|' + $file.FullName + '|' + $file.LastWriteTimeUtc.Ticks
        })
        $repeat = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath
        $repeat.Result.Succeeded | Should -BeTrue -Because ($repeat.Stdout + $repeat.Stderr + $repeat.Result.ErrorMessage)
        $repeat.Stdout | Should -Match 'Skipped\s+: 3'
        $after = Get-UpdPaginationArchive -OutputPath $repeat.OutputPath
        @($after.State.episodes).Count | Should -Be 3
        $snapshot = @($after.State.episodes | Sort-Object episode_id | ForEach-Object {
            $file = Get-Item -LiteralPath (Join-Path $after.Root $_.relative_path)
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash | Should -Be $script:paginationHash
            $_.episode_id + '|' + $file.FullName + '|' + $file.LastWriteTimeUtc.Ticks
        })
        ($snapshot -join "`n") | Should -Be ($before -join "`n")
        @(Get-ChildItem -LiteralPath $repeat.OutputPath -Recurse -File -Filter '*.mp3').Count | Should -Be 3
        $stats = Get-UpdFixtureState -Context $context
        $stats.PSObject.Properties[$FeedPath].Value | Should -Be 2
        $stats.PSObject.Properties[$PagePath].Value | Should -Be 2
        $stats.'/media/pagination.mp3' | Should -Be 3
    }

    It 'visits every advertised page before <Mode> chooses the newer later-page entry' -TestCases @(
        @{ Mode = 'Latest'; Count = 1; Names = @('new') }
        @{ Mode = 'Custom'; Count = 2; Names = @('new', 'shared') }
    ) {
        param($Mode, $Count, $Names)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/pagination-rss.xml' -Mode $Mode -CustomCount $Count
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Stdout | Should -Match ('Downloaded\s+: ' + $Count)
        $archive = Get-UpdPaginationArchive -OutputPath $run.OutputPath
        @($archive.State.episodes).Count | Should -Be $Count
        for ($index = 0; $index -lt $Count; $index++) {
            $expected = Get-PodcastEpisodeIdentity -FeedId $archive.State.feed_id -Episode ([pscustomobject]@{
                Guid = 'pagination-' + $Names[$index]; AtomId = $null; Url = $context.BaseUrl + '/media/pagination.mp3?id=' + $Names[$index]
            })
            $archive.State.episodes[$index].episode_id | Should -Be $expected.Id
            (Get-FileHash -LiteralPath (Join-Path $archive.Root $archive.State.episodes[$index].relative_path) -Algorithm SHA256).Hash | Should -Be $script:paginationHash
        }
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/pagination-rss.xml' | Should -Be 1
        $stats.'/feeds/pagination-rss-page-2.xml' | Should -Be 1
        $stats.'/media/pagination.mp3' | Should -Be $Count
    }

    It 'processes accessible entries in <Mode> but preserves an incomplete outcome after a <Problem> continuation' -TestCases @(
        @{ Problem = '404'; Mode = 'Latest'; Count = 1; FeedPath = '/feeds/pagination-gap.xml'; Target = '/pagination/missing.xml' }
        @{ Problem = '404'; Mode = 'Custom'; Count = 2; FeedPath = '/feeds/pagination-gap.xml'; Target = '/pagination/missing.xml' }
        @{ Problem = 'malformed XML'; Mode = 'All'; Count = 1; FeedPath = '/feeds/pagination-malformed.xml'; Target = '/feeds/malformed.xml' }
        @{ Problem = 'HTML'; Mode = 'All'; Count = 1; FeedPath = '/feeds/pagination-html.xml'; Target = '/show/not-feed' }
    ) {
        param($Problem, $Mode, $Count, $FeedPath, $Target)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath -Mode $Mode -CustomCount $Count
        Assert-UpdPaginationIncomplete -Run $run -Downloads 1
        $archive = Get-UpdPaginationArchive -OutputPath $run.OutputPath
        @($archive.State.episodes).Count | Should -Be 1 -Because ($Problem + ' does not erase the accessible first-page entry.')
        $archive.State.episodes[0].status | Should -Be 'transfer_verified'
        (Get-FileHash -LiteralPath (Join-Path $archive.Root $archive.State.episodes[0].relative_path) -Algorithm SHA256).Hash | Should -Be $script:paginationHash
        $stats = Get-UpdFixtureState -Context $context
        $stats.PSObject.Properties[$FeedPath].Value | Should -Be 1
        $stats.PSObject.Properties[$Target].Value | Should -Be 1
        $stats.'/media/pagination.mp3' | Should -Be 1
    }

    It 'uses the redirected effective URI and inherited xml:base while reusing the initial response once' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Resolve -FeedPath '/feeds/pagination-relative.xml' -ReuseResponse -MaxFeedPages 3
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Result.ResolvedUrl | Should -Be ($context.BaseUrl + '/feeds/pagination-relative.xml')
        $run.Result.ItemCount | Should -Be 2
        (@($run.Result.EpisodeTitles) -join '|') | Should -Be 'Pagination shared|Pagination new'
        $run.Result.Catalogue.Complete | Should -BeTrue
        $run.Result.Catalogue.PagesFetched | Should -Be 2
        $run.Result.Catalogue.ItemsSeen | Should -Be 2
        $run.Result.Catalogue.DuplicateCount | Should -Be 0
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/pagination-relative.xml' | Should -Be 1
        $stats.'/pagination/redirected/start.xml' | Should -Be 1
        $stats.'/pagination/redirected/catalog/page-2.xml' | Should -Be 1
        $stats.'/media/pagination.mp3' | Should -BeNullOrEmpty
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
    }

    It 'does not choose between distinct advertised continuation links and reports accessible scope as incomplete' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/pagination-ambiguous.xml'
        Assert-UpdPaginationIncomplete -Run $run -Downloads 1
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/pagination-rss-page-2.xml' | Should -BeNullOrEmpty
        $stats.'/feeds/pagination-atom-page-2.xml' | Should -BeNullOrEmpty
        $stats.'/media/pagination.mp3' | Should -Be 1
    }

    It 'enforces the configured page budget before fetching a continuation and exposes the limited catalogue' {
        $resolution = Invoke-UpdIntegrationWorker -Context $context -Action Resolve -FeedPath '/feeds/pagination-rss.xml' -MaxFeedPages 1
        $resolution.Result.Succeeded | Should -BeTrue -Because ($resolution.Stdout + $resolution.Stderr + $resolution.Result.ErrorMessage)
        $resolution.Result.ItemCount | Should -Be 2
        $resolution.Result.Catalogue.Complete | Should -BeFalse
        $resolution.Result.Catalogue.StopReason | Should -Be 'page_limit'
        $resolution.Result.Catalogue.PagesFetched | Should -Be 1
        $resolution.Result.Catalogue.ItemsSeen | Should -Be 2
        $resolution.Result.Catalogue.MaxPages | Should -Be 1
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/pagination-rss.xml' -MaxFeedPages 1
        Assert-UpdPaginationIncomplete -Run $run -Downloads 2
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/pagination-rss.xml' | Should -Be 2
        $stats.'/feeds/pagination-rss-page-2.xml' | Should -BeNullOrEmpty
        $stats.'/media/pagination.mp3' | Should -Be 2
    }

    It 'ignores unsupported relations, unqualified RSS links and HTTP Link pagination without crawling' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/pagination-ignored.xml'
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Stdout | Should -Match 'Downloaded\s+: 1'
        $run.Stdout | Should -Not -Match 'Feed catalogue incomplete:'
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/pagination-rss-page-2.xml' | Should -BeNullOrEmpty
        $stats.'/feeds/pagination-atom-page-2.xml' | Should -BeNullOrEmpty
        $stats.'/media/pagination.mp3' | Should -Be 1
    }

    It 'rejects an unsafe advertised target before fetching it while retaining an incomplete result' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/pagination-unsafe.xml'
        Assert-UpdPaginationIncomplete -Run $run -Downloads 1
        ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage) | Should -Not -Match 'file:///|synthetic-pagination-must-not-read'
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/pagination-unsafe.xml' | Should -Be 1
        $stats.'/media/pagination.mp3' | Should -Be 1
    }
}
