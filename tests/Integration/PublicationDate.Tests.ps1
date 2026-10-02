BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    . (Join-Path $repositoryRoot 'UniversalPodcastDownloader.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned loopback publication-date integration fixtures.' }
    $script:publicationPython = $python.Source
    $sample = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
    $script:publicationMediaHash = (Get-FileHash -LiteralPath $sample -Algorithm SHA256).Hash
    $script:publicationMediaLength = (Get-Item -LiteralPath $sample).Length

    function Get-UpdPublicationArchive {
        param([Parameter(Mandatory)][string]$OutputPath)
        $states = @(Get-ChildItem -LiteralPath $OutputPath -Recurse -Force -File -Filter 'state.json')
        $states.Count | Should -Be 1
        $archiveRoot = Split-Path $states[0].DirectoryName -Parent
        [pscustomobject]@{
            Root = $archiveRoot
            Path = $states[0].FullName
            State = Read-PodcastHistory -Root $archiveRoot
        }
    }
}

Describe 'A033/A034 publication dates at real download and history boundaries' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:publicationPython
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'Latest downloads the newest published Atom entry with its Atom ID under <Culture>' -TestCases @(
        @{ Culture = 'en-US' }
        @{ Culture = 'de-DE' }
        @{ Culture = 'fi-FI' }
    ) {
        param($Culture)
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/atom.xml' -Mode Latest -Culture $Culture
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Stdout | Should -Match 'Downloaded\s+: 1'
        $archive = Get-UpdPublicationArchive -OutputPath $run.OutputPath
        @($archive.State.episodes).Count | Should -Be 1
        $record = $archive.State.episodes[0]
        $record.identity_source | Should -Be 'atom-id'
        $expected = Get-PodcastEpisodeIdentity -FeedId $archive.State.feed_id -Episode ([pscustomobject]@{
            Guid = $null; AtomId = 'urn:fixture:new'; Url = $context.BaseUrl + '/media/ok.mp3?id=new'
        })
        $record.episode_id | Should -Be $expected.Id
        $record.relative_path | Should -Match '^2026-09-29 - '
        $media = Get-Item -LiteralPath (Join-Path $archive.Root $record.relative_path)
        $media.Length | Should -Be $script:publicationMediaLength
        (Get-FileHash -LiteralPath $media.FullName -Algorithm SHA256).Hash | Should -Be $script:publicationMediaHash
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
    }

    It 'Custom keeps source order for equal instants and undated entries after dated entries' {
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/publication-order.xml' -Mode Custom -CustomCount 3 -Culture 'fi-FI'
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $run.Stdout | Should -Match 'Downloaded\s+: 3'
        $archive = Get-UpdPublicationArchive -OutputPath $run.OutputPath
        @($archive.State.episodes).Count | Should -Be 3
        $expectedGuids = @('publication-tie-first', 'publication-tie-second', 'publication-missing-first')
        for ($index = 0; $index -lt $expectedGuids.Count; $index++) {
            $record = $archive.State.episodes[$index]
            $expected = Get-PodcastEpisodeIdentity -FeedId $archive.State.feed_id -Episode ([pscustomobject]@{
                Guid = $expectedGuids[$index]; AtomId = $null; Url = $context.BaseUrl + '/media/ok.mp3'
            })
            $record.episode_id | Should -Be $expected.Id
            if ($index -lt 2) { $record.relative_path | Should -Match '^2026-09-01 - ' }
            else { $record.relative_path | Should -Not -Match '^\d{4}-\d{2}-\d{2} - ' }
            (Get-FileHash -LiteralPath (Join-Path $archive.Root $record.relative_path) -Algorithm SHA256).Hash | Should -Be $script:publicationMediaHash
        }
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 3
    }

    It 'previews without writes or media and uses the UTC calendar day across midnight when downloading' {
        $preview = Invoke-UpdIntegrationWorker -Context $context -Action Preview -FeedPath '/feeds/publication-midnight.xml' -Mode Latest -Culture 'fi-FI'
        $preview.Result.Succeeded | Should -BeTrue -Because ($preview.Stdout + $preview.Stderr + $preview.Result.ErrorMessage)
        Test-Path -LiteralPath $preview.OutputPath | Should -BeFalse
        @(Get-ChildItem -LiteralPath $context.Root -Recurse -File -Filter '*.log').Count | Should -Be 0
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -BeNullOrEmpty
        $run = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath '/feeds/publication-midnight.xml' -Mode Latest -Culture 'fi-FI'
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $media = @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3')
        $media.Count | Should -Be 1
        $media[0].Name | Should -Match '^2026-08-31 - '
        (Get-FileHash -LiteralPath $media[0].FullName -Algorithm SHA256).Hash | Should -Be $script:publicationMediaHash
        $stats = Get-UpdFixtureState -Context $context
        $stats.'/feeds/publication-midnight.xml' | Should -Be 2
        $stats.'/media/ok.mp3' | Should -Be 1
    }

    It 'retains historical recorded paths and bytes after publication metadata changes for <Kind>' -TestCases @(
        @{ Kind = 'RSS'; FeedPath = '/feeds/publication-history-rss.xml'; IdentitySource = 'rss-guid' }
        @{ Kind = 'Atom'; FeedPath = '/feeds/publication-history-atom.xml'; IdentitySource = 'atom-id' }
    ) {
        param($Kind, $FeedPath, $IdentitySource)
        $first = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath -Culture 'en-US'
        $first.Result.Succeeded | Should -BeTrue -Because ($first.Stdout + $first.Stderr + $first.Result.ErrorMessage)
        $archive = Get-UpdPublicationArchive -OutputPath $first.OutputPath
        @($archive.State.episodes).Count | Should -Be 1
        $record = $archive.State.episodes[0]
        $record.identity_source | Should -Be $IdentitySource -Because ($Kind + ' publisher identifiers remain stable when publication metadata changes.')
        $record.relative_path | Should -Match '^2026-08-31 - '
        $originalPath = Join-Path $archive.Root $record.relative_path

        # Seed only this owned synthetic archive with a verified historical
        # local-day filename. Product runs must honor its recorded destination.
        $historicalName = '2026-09-01' + $record.relative_path.Substring(10)
        $historicalPath = Join-Path $archive.Root $historicalName
        [IO.File]::Move($originalPath, $historicalPath)
        $historyLock = Enter-PodcastHistoryLock -Root $archive.Root
        try {
            $historicalState = Read-PodcastHistory -Root $archive.Root
            $historicalState.episodes[0].relative_path = $historicalName
            $historicalState.generation++
            $null = Write-PodcastHistory -Lock $historyLock -State $historicalState
        }
        finally { $historyLock.Stream.Dispose() }
        $historicalFile = Get-Item -LiteralPath $historicalPath
        $originalStamp = $historicalFile.LastWriteTimeUtc
        $originalHash = (Get-FileHash -LiteralPath $historicalPath -Algorithm SHA256).Hash
        $originalHash | Should -Be $script:publicationMediaHash
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing -TimeoutSec 10

        $second = Invoke-UpdIntegrationWorker -Context $context -Action Download -FeedPath $FeedPath -Culture 'de-DE'
        $second.Result.Succeeded | Should -BeTrue -Because ($second.Stdout + $second.Stderr + $second.Result.ErrorMessage)
        $second.Stdout | Should -Match 'Skipped\s+: 1'
        $after = Get-UpdPublicationArchive -OutputPath $second.OutputPath
        $after.Path | Should -Be $archive.Path
        @($after.State.episodes).Count | Should -Be 1
        $after.State.episodes[0].episode_id | Should -Be $record.episode_id
        $after.State.episodes[0].relative_path | Should -Be $historicalName
        $after.State.episodes[0].status | Should -Be 'transfer_verified'
        $files = @(Get-ChildItem -LiteralPath $second.OutputPath -Recurse -File -Filter '*.mp3')
        $files.Count | Should -Be 1
        $files[0].FullName | Should -Be $historicalPath
        $files[0].LastWriteTimeUtc | Should -Be $originalStamp
        $files[0].Length | Should -Be $script:publicationMediaLength
        (Get-FileHash -LiteralPath $historicalPath -Algorithm SHA256).Hash | Should -Be $originalHash
        Test-Path -LiteralPath $originalPath | Should -BeFalse
        $stats = Get-UpdFixtureState -Context $context
        $stats.PSObject.Properties[$FeedPath].Value | Should -Be 2
        $stats.'/media/ok.mp3' | Should -Be 1
    }
}
