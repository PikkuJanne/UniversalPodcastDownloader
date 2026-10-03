BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    . (Join-Path $repositoryRoot 'UniversalPodcastDownloader.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for the owned loopback legacy integration fixtures.' }
    $script:legacyPython = $python.Source
    $script:legacySample = Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'
    $script:legacyName = '2026-09-01 - Original episode.mp3'

    function Get-UpdLegacySnapshot {
        param([Parameter(Mandatory)][string]$Root)

        return @(Get-ChildItem -LiteralPath $Root -Recurse -Force | Sort-Object FullName | ForEach-Object {
            $relative = $_.FullName.Substring($Root.Length).TrimStart([char[]]'\/')
            if ($_.PSIsContainer) { 'directory|' + $relative }
            else { 'file|' + $relative + '|' + $_.Length + '|' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
        })
    }

    function Assert-UpdLegacyOriginal {
        param([Parameter(Mandatory)]$Archive)

        foreach ($original in $Archive.Originals) {
            $path = Join-Path $Archive.Root $original.RelativePath
            Test-Path -LiteralPath $path -PathType Leaf | Should -BeTrue -Because ('every original must survive: ' + $original.RelativePath)
            (Get-Item -LiteralPath $path).Length | Should -Be $original.Bytes
            (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -Be $original.Sha256
        }
    }

    function New-UpdLegacyCopy {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an isolated synthetic archive copy only within its caller-owned marked test root.')]
        [CmdletBinding()]
        param([Parameter(Mandatory)]$Context)

        $source = Join-Path $Context.Root 'synthetic-source'
        $output = Join-Path $Context.Root 'output'
        $root = Join-Path $output 'Original history show'
        $null = [IO.Directory]::CreateDirectory($source)
        $null = [IO.Directory]::CreateDirectory($root)
        [IO.File]::Copy($script:legacySample, (Join-Path $source $script:legacyName))
        [IO.File]::Copy($script:legacySample, (Join-Path $source 'unmatched.mp3'))
        [IO.File]::WriteAllBytes((Join-Path $source 'empty.mp3'), [byte[]]@())
        [IO.File]::WriteAllText((Join-Path $source 'text.mp3'), '<html>Synthetic error response, not audio.</html>')
        [IO.File]::WriteAllBytes((Join-Path $source 'unfinished.part'), [byte[]]@(73, 68, 51, 0, 0))
        [IO.File]::WriteAllText((Join-Path $source 'owner-notes.txt'), 'Preserve this original note as well as all media.')
        $feedId = Get-PodcastNameHash -IdentityKey ('feed:' + $Context.BaseUrl + '/feeds/history.xml')
        $episode = [pscustomobject]@{
            Title = 'Original episode'
            Guid = 'history-stable-001'
            AtomId = $null
            Url = $Context.BaseUrl + '/media/history.mp3?signature=original&part=1'
            PubDate = [datetime]'2026-09-01T12:00:00Z'
        }
        $identity = Get-PodcastEpisodeIdentity -Episode $episode -FeedId $feedId
        # Copy from a separate synthetic source, then record every original file.
        Get-ChildItem -LiteralPath $source -File | ForEach-Object { [IO.File]::Copy($_.FullName, (Join-Path $root $_.Name)) }
        $originals = @(Get-ChildItem -LiteralPath $root -Recurse -Force -File | ForEach-Object {
            [pscustomobject]@{
                RelativePath = $_.FullName.Substring($root.Length).TrimStart([char[]]'\/')
                Bytes = $_.Length
                Sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            }
        })
        [pscustomobject]@{ Root = $root; OutputPath = $output; Source = $source; EpisodeId = $identity.Id; Originals = $originals; FeedPath = '/feeds/history.xml' }
    }

    function Invoke-UpdLegacyWorker {
        param(
            [Parameter(Mandatory)]$Context,
            [Parameter(Mandatory)]$Archive,
            [ValidateSet('Preview', 'Adopt', 'Redownload', 'Rollback')][string]$Action,
            [string]$EpisodeId,
            [string]$File,
            [switch]$OmitSha256,
            [string]$Checkpoint,
            [switch]$Normal,
            [switch]$PreviewOnly
        )

        $token = [guid]::NewGuid().ToString('N')
        $resultPath = Join-Path $Context.Root ($token + '-legacy-result.json')
        $configPath = Join-Path $Context.Root ($token + '-legacy-config.json')
        $config = @{
            ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'
            FeedUrl = $Context.BaseUrl + $Archive.FeedPath
            OutputPath = $Archive.OutputPath
            ResultPath = $resultPath
            LegacyPath = if ($Normal) { $null } else { $Archive.Root }
            LegacyAction = $Action
            LegacyEpisodeId = $EpisodeId
            LegacyFile = $File
            LegacySha256 = if ($Action -eq 'Adopt' -and $File -and -not $OmitSha256) { @($Archive.Originals | Where-Object RelativePath -eq $File)[0].Sha256.ToLowerInvariant() } else { $null }
            LegacyCheckpoint = $Checkpoint
            PreviewOnly = [bool]$PreviewOnly
        }
        $config | ConvertTo-Json | Set-Content -LiteralPath $configPath -Encoding UTF8
        $engineName = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $worker = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engineName) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-LegacyWorker.ps1'), '-ConfigPath', $configPath
        )
        if (-not $worker.Process.HasExited -and -not $worker.Process.WaitForExit(30000)) {
            $worker.Process.Kill()
            $worker.Process.WaitForExit()
            throw 'Legacy integration child timed out after 30 seconds; only its owned process was stopped.'
        }
        $stdout = $worker.Output.Result
        $stderr = $worker.ErrorOutput.Result
        if (-not (Test-Path -LiteralPath $resultPath)) {
            throw "Legacy integration child produced no result. Exit=$($worker.Process.ExitCode); stdout=$stdout; stderr=$stderr"
        }
        [pscustomobject]@{
            ExitCode = $worker.Process.ExitCode
            Result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
            Stdout = $stdout
            Stderr = $stderr
        }
    }

    function Get-UpdLegacyState {
        param([Parameter(Mandatory)]$Archive)

        return (Get-Content -LiteralPath (Join-Path $Archive.Root '.upd/state.json') -Raw | ConvertFrom-Json)
    }

    function Assert-UpdLegacyNoMediaRequest {
        param([Parameter(Mandatory)]$Context)

        $stats = Get-UpdFixtureState -Context $Context
        @($stats.PSObject.Properties | Where-Object { $_.Name -like '/media/*' }).Count | Should -Be 0
    }
}

Describe 'Legacy archive migration through actual CLI processes' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:legacyPython
        $script:legacy = New-UpdLegacyCopy -Context $context
    }

    AfterEach {
        if ($context) { Remove-UpdIntegrationContext -Context $context }
    }

    It 'A017 A018 normal WhatIf preserves the entire copied output tree and makes no media requests' {
        $before = Get-UpdLegacySnapshot -Root $legacy.OutputPath
        $run = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Normal -PreviewOnly
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        (Get-UpdLegacySnapshot -Root $legacy.OutputPath) -join "`n" | Should -Be ($before -join "`n")
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 normal WhatIf does not create an absent output root' {
        $legacy.OutputPath = Join-Path $context.Root 'absent-output'
        $run = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Normal -PreviewOnly
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        Test-Path -LiteralPath $legacy.OutputPath | Should -BeFalse
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 <Action> preview inventories every original without writing state logs locks or media' -TestCases @(
        @{ Action = $null }
        @{ Action = 'Preview' }
    ) {
        param($Action)
        $before = Get-UpdLegacySnapshot -Root $legacy.OutputPath
        $parameters = @{ Context = $context; Archive = $legacy }
        if ($Action) { $parameters.Action = $Action }
        $run = Invoke-UpdLegacyWorker @parameters
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $plan = @($run.Result.Output)[-1]
        $plan.Root | Should -Be $legacy.Root
        @($plan.Files).Count | Should -Be $legacy.Originals.Count
        @($plan.Episodes).Count | Should -Be 1
        $plan.Episodes[0].EpisodeId | Should -Be $legacy.EpisodeId
        $plan.Episodes[0].Classification | Should -Be 'unverified'
        $plan.Episodes[0].SuggestedPath | Should -Be $script:legacyName
        @($plan.Files | Where-Object { $_.Classification -eq 'conflict' }).Count | Should -BeGreaterThan 2
        ($plan | ConvertTo-Json -Depth 12) | Should -Not -Match 'transfer_verified|signature=|history-stable-001|http://'
        (Get-UpdLegacySnapshot -Root $legacy.OutputPath) -join "`n" | Should -Be ($before -join "`n")
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 <Action> WhatIf returns a read-only preview even when an explicit operation is selected' -TestCases @(
        @{ Action = 'Adopt' }
        @{ Action = 'Redownload' }
    ) {
        param($Action)
        $before = Get-UpdLegacySnapshot -Root $legacy.OutputPath
        $run = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action $Action -EpisodeId $legacy.EpisodeId -File $script:legacyName -PreviewOnly
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        @($run.Result.Output)[-1].Episodes[0].EpisodeId | Should -Be $legacy.EpisodeId
        (Get-UpdLegacySnapshot -Root $legacy.OutputPath) -join "`n" | Should -Be ($before -join "`n")
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 an ordinary run flags a title-matched legacy archive instead of downloading a second archive' {
        $before = Get-UpdLegacySnapshot -Root $legacy.OutputPath
        $run = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Normal
        $run.Result.Succeeded | Should -BeFalse
        ($run.Stdout + $run.Result.ErrorMessage) | Should -Match 'legacy|unverified|review'
        (Get-UpdLegacySnapshot -Root $legacy.OutputPath) -join "`n" | Should -Be ($before -join "`n")
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A018 adoption requires an explicit <Missing> selection' -TestCases @(
        @{ Missing = 'episode'; IncludeEpisode = $false; IncludeFile = $true; OmitSha256 = $false }
        @{ Missing = 'file'; IncludeEpisode = $true; IncludeFile = $false; OmitSha256 = $false }
        @{ Missing = 'reviewed digest'; IncludeEpisode = $true; IncludeFile = $true; OmitSha256 = $true }
    ) {
        param($Missing, $IncludeEpisode, $IncludeFile, $OmitSha256)
        $parameters = @{ Context = $context; Archive = $legacy; Action = 'Adopt'; OmitSha256 = $OmitSha256 }
        if ($IncludeEpisode) { $parameters.EpisodeId = $legacy.EpisodeId }
        if ($IncludeFile) { $parameters.File = $script:legacyName }
        $run = Invoke-UpdLegacyWorker @parameters
        $run.Result.Succeeded | Should -BeFalse -Because ('owner choice must include the ' + $Missing)
        Test-Path -LiteralPath (Join-Path $legacy.Root '.upd/state.json') | Should -BeFalse
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A018 a file changed after preview cannot be adopted with the reviewed original digest' {
        $preview = Invoke-UpdLegacyWorker -Context $context -Archive $legacy
        $preview.Result.Succeeded | Should -BeTrue -Because ($preview.Stdout + $preview.Stderr + $preview.Result.ErrorMessage)
        Assert-UpdLegacyOriginal -Archive $legacy
        $path = Join-Path $legacy.Root $script:legacyName
        $bytes = [IO.File]::ReadAllBytes($path)
        $bytes[$bytes.Length - 1] = $bytes[$bytes.Length - 1] -bxor 1
        [IO.File]::WriteAllBytes($path, $bytes)
        $before = Get-UpdLegacySnapshot -Root $legacy.OutputPath
        $adopt = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Adopt -EpisodeId $legacy.EpisodeId -File $script:legacyName
        $adopt.Result.Succeeded | Should -BeFalse
        $adopt.Result.RunResult.Status | Should -Be 'fatal'
        $adopt.Result.ErrorMessage | Should -Match 'Private error details were omitted'
        (Get-UpdLegacySnapshot -Root $legacy.OutputPath) -join "`n" | Should -Be ($before -join "`n")
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 explicit adoption records local confidence and repeat runs remain adopted without a media request' {
        $run = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Adopt -EpisodeId $legacy.EpisodeId -File $script:legacyName
        $run.Result.Succeeded | Should -BeTrue -Because ($run.Stdout + $run.Stderr + $run.Result.ErrorMessage)
        $state = Get-UpdLegacyState -Archive $legacy
        $state.schema_version | Should -Be 2
        @($state.episodes).Count | Should -Be 1
        $state.episodes[0].status | Should -Be 'adopted'
        $state.episodes[0].relative_path | Should -Be $script:legacyName
        $state.episodes[0].verification.method | Should -Be 'owner-approved-local-signature'
        ($state.episodes[0].verification.notes -join ' ') | Should -Match 'complet|publisher|local'
        $state.episodes[0].local_sha256 | Should -Be ((Get-FileHash -LiteralPath (Join-Path $legacy.Root $script:legacyName) -Algorithm SHA256).Hash)
        @($run.Result.Output)[-1].Checkpoint | Should -Match '^legacy-[a-f0-9]{32}\.json$'
        $again = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Normal
        $again.Result.Succeeded | Should -BeFalse
        $again.ExitCode | Should -Be 2
        $again.Result.RunResult.LegacyUnverified | Should -Be 1
        $again.Stdout | Should -Match 'Adopted\s+: 1'
        (Get-UpdLegacyState -Archive $legacy).episodes[0].status | Should -Be 'adopted'
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 adoption refuses <File> while preserving all original files' -TestCases @(
        @{ File = 'empty.mp3' }
        @{ File = 'text.mp3' }
        @{ File = 'unfinished.part' }
    ) {
        param($File)
        $run = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Adopt -EpisodeId $legacy.EpisodeId -File $File
        $run.Result.Succeeded | Should -BeFalse
        if (Test-Path -LiteralPath (Join-Path $legacy.Root '.upd/state.json')) {
            @((Get-UpdLegacyState -Archive $legacy).episodes | Where-Object { $_.status -eq 'adopted' }).Count | Should -Be 0
        }
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 ambiguous filenames remain conflicts until the owner selects one exact file' {
        $episode = [pscustomobject]@{ Title = 'Original episode'; Url = $context.BaseUrl + '/media/history.mp3'; PubDate = [datetime]'2026-09-01T12:00:00Z' }
        $modernName = New-EpisodeFileName -Episode $episode -IdentityHash $legacy.EpisodeId
        $duplicate = Join-Path $legacy.Root $modernName
        [IO.File]::Copy($script:legacySample, $duplicate)
        $legacy.Originals += [pscustomobject]@{ RelativePath = $modernName; Bytes = (Get-Item -LiteralPath $duplicate).Length; Sha256 = (Get-FileHash -LiteralPath $duplicate -Algorithm SHA256).Hash }
        $before = Get-UpdLegacySnapshot -Root $legacy.OutputPath
        $preview = Invoke-UpdLegacyWorker -Context $context -Archive $legacy
        $preview.Result.Succeeded | Should -BeTrue -Because ($preview.Stdout + $preview.Stderr + $preview.Result.ErrorMessage)
        $row = @($preview.Result.Output)[-1].Episodes[0]
        $row.Classification | Should -Be 'conflict'
        @($row.Candidates).Count | Should -Be 2
        $row.SuggestedPath | Should -BeNullOrEmpty
        (Get-UpdLegacySnapshot -Root $legacy.OutputPath) -join "`n" | Should -Be ($before -join "`n")
        $adopt = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Adopt -EpisodeId $legacy.EpisodeId -File $script:legacyName
        $adopt.Result.Succeeded | Should -BeTrue -Because ($adopt.Stdout + $adopt.Stderr + $adopt.Result.ErrorMessage)
        (Get-UpdLegacyState -Archive $legacy).episodes[0].relative_path | Should -Be $script:legacyName
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 a <Mutation> adopted file requires review on every rerun without automatic download' -TestCases @(
        @{ Mutation = 'changed' }
        @{ Mutation = 'missing' }
    ) {
        param($Mutation)
        $adopt = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Adopt -EpisodeId $legacy.EpisodeId -File $script:legacyName
        $adopt.Result.Succeeded | Should -BeTrue -Because ($adopt.Stdout + $adopt.Stderr + $adopt.Result.ErrorMessage)
        Assert-UpdLegacyOriginal -Archive $legacy
        $path = Join-Path $legacy.Root $script:legacyName
        if ($Mutation -eq 'changed') {
            $bytes = [IO.File]::ReadAllBytes($path)
            $bytes[$bytes.Length - 1] = $bytes[$bytes.Length - 1] -bxor 1
            [IO.File]::WriteAllBytes($path, $bytes)
        }
        else { [IO.File]::Move($path, (Join-Path $context.Root 'owner-moved-original.mp3')) }
        $afterOwnerChange = Get-UpdLegacySnapshot -Root $legacy.Root | Where-Object { $_ -notmatch '\|\.upd(?:\\|\|)|\.log\|' }
        foreach ($attempt in 1..2) {
            $run = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Normal
            $run.Result.Succeeded | Should -BeFalse -Because ('attempt ' + $attempt + ' still needs an owner decision')
            $run.ExitCode | Should -Be 2
            $run.Result.RunResult.Status | Should -Be 'incomplete'
            $run.Result.RunResult.Conflicts | Should -Be 1
            $run.Result.RunResult.Episodes[0].Outcome | Should -Be 'conflict'
            (Get-UpdLegacyState -Archive $legacy).episodes[0].status | Should -Not -Be 'transfer_verified'
            $afterRun = Get-UpdLegacySnapshot -Root $legacy.Root | Where-Object { $_ -notmatch '\|\.upd(?:\\|\|)|\.log\|' }
            $afterRun -join "`n" | Should -Be ($afterOwnerChange -join "`n")
        }
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 remote title URL and media bytes changes never overwrite an adopted original' {
        $legacy.FeedPath = '/feeds/legacy-changing.xml'
        $feedId = Get-PodcastNameHash -IdentityKey ('feed:' + $context.BaseUrl + $legacy.FeedPath)
        $legacy.EpisodeId = (Get-PodcastEpisodeIdentity -Episode ([pscustomobject]@{ Guid = 'history-stable-001' }) -FeedId $feedId).Id
        $adopt = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Adopt -EpisodeId $legacy.EpisodeId -File $script:legacyName
        $adopt.Result.Succeeded | Should -BeTrue -Because ($adopt.Stdout + $adopt.Stderr + $adopt.Result.ErrorMessage)
        $null = Invoke-WebRequest -Uri ($context.BaseUrl + '/__recover') -Method Post -UseBasicParsing -TimeoutSec 10
        $run = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Normal
        $run.Result.Succeeded | Should -BeFalse
        $run.ExitCode | Should -Be 2
        $run.Result.RunResult.LegacyUnverified | Should -Be 1
        $run.Stdout | Should -Match 'Adopted\s+: 1'
        (Get-UpdLegacyState -Archive $legacy).episodes[0].status | Should -Be 'adopted'
        Assert-UpdLegacyOriginal -Archive $legacy
        Assert-UpdLegacyNoMediaRequest -Context $context
    }

    It 'A017 A018 explicit redownload uses a separate safe target and rollback restores metadata only' {
        $download = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Redownload -EpisodeId $legacy.EpisodeId
        $download.Result.Succeeded | Should -BeTrue -Because ($download.Stdout + $download.Stderr + $download.Result.ErrorMessage)
        $checkpoint = @($download.Result.Output)[-1].Checkpoint
        $checkpoint | Should -Match '^legacy-[a-f0-9]{32}\.json$'
        $state = Get-UpdLegacyState -Archive $legacy
        @($state.episodes).Count | Should -Be 1
        $state.episodes[0].status | Should -Be 'transfer_verified'
        $state.episodes[0].relative_path | Should -Not -Be $script:legacyName
        $downloadedPath = Join-Path $legacy.Root $state.episodes[0].relative_path
        Test-Path -LiteralPath $downloadedPath | Should -BeTrue
        $downloadedHash = (Get-FileHash -LiteralPath $downloadedPath -Algorithm SHA256).Hash
        $downloadedHash | Should -Be ((Get-FileHash -LiteralPath $script:legacySample -Algorithm SHA256).Hash)
        Assert-UpdLegacyOriginal -Archive $legacy
        $beforeRollbackPreview = Get-UpdLegacySnapshot -Root $legacy.OutputPath
        $preview = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Rollback -Checkpoint $checkpoint -PreviewOnly
        $preview.Result.Succeeded | Should -BeTrue -Because ($preview.Stdout + $preview.Stderr + $preview.Result.ErrorMessage)
        (Get-UpdLegacySnapshot -Root $legacy.OutputPath) -join "`n" | Should -Be ($beforeRollbackPreview -join "`n")
        $rollback = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Rollback -Checkpoint $checkpoint
        $rollback.Result.Succeeded | Should -BeTrue -Because ($rollback.Stdout + $rollback.Stderr + $rollback.Result.ErrorMessage)
        $restored = Get-UpdLegacyState -Archive $legacy
        $restored.generation | Should -BeGreaterThan $state.generation
        @($restored.episodes).Count | Should -Be 0
        (Get-FileHash -LiteralPath $downloadedPath -Algorithm SHA256).Hash | Should -Be $downloadedHash
        Test-Path -LiteralPath (Join-Path $legacy.Root ('.upd/' + $checkpoint)) | Should -BeTrue
        Assert-UpdLegacyOriginal -Archive $legacy
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }

    It 'A017 A018 redownload preserves an occupied modern destination as well as the historical file' {
        $episode = [pscustomobject]@{ Title = 'Original episode'; Url = $context.BaseUrl + '/media/history.mp3'; PubDate = [datetime]'2026-09-01T12:00:00Z' }
        $modernName = New-EpisodeFileName -Episode $episode -IdentityHash $legacy.EpisodeId
        $occupied = Join-Path $legacy.Root $modernName
        [IO.File]::WriteAllText($occupied, 'Synthetic unknown original occupying the modern destination.')
        $legacy.Originals += [pscustomobject]@{ RelativePath = $modernName; Bytes = (Get-Item -LiteralPath $occupied).Length; Sha256 = (Get-FileHash -LiteralPath $occupied -Algorithm SHA256).Hash }
        $download = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Redownload -EpisodeId $legacy.EpisodeId
        $download.Result.Succeeded | Should -BeTrue -Because ($download.Stdout + $download.Stderr + $download.Result.ErrorMessage)
        $record = (Get-UpdLegacyState -Archive $legacy).episodes[0]
        $record.status | Should -Be 'transfer_verified'
        $record.relative_path | Should -Not -Be $script:legacyName
        $record.relative_path | Should -Not -Be $modernName
        (Get-FileHash -LiteralPath (Join-Path $legacy.Root $record.relative_path) -Algorithm SHA256).Hash | Should -Be ((Get-FileHash -LiteralPath $script:legacySample -Algorithm SHA256).Hash)
        Assert-UpdLegacyOriginal -Archive $legacy
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }

    It 'A017 A018 explicit redownload selects one identity from a six-episode feed without downloading the rest' {
        $legacy.FeedPath = '/feeds/collisions.xml'
        $feedId = Get-PodcastNameHash -IdentityKey ('feed:' + $context.BaseUrl + $legacy.FeedPath)
        $legacy.EpisodeId = (Get-PodcastEpisodeIdentity -Episode ([pscustomobject]@{ Guid = 'one' }) -FeedId $feedId).Id
        $historicalName = '2026-09-01 - Same title.mp3'
        $historicalPath = Join-Path $legacy.Root $historicalName
        [IO.File]::Copy($script:legacySample, $historicalPath)
        $legacy.Originals += [pscustomobject]@{ RelativePath = $historicalName; Bytes = (Get-Item -LiteralPath $historicalPath).Length; Sha256 = (Get-FileHash -LiteralPath $historicalPath -Algorithm SHA256).Hash }
        $preview = Invoke-UpdLegacyWorker -Context $context -Archive $legacy
        $preview.Result.Succeeded | Should -BeTrue -Because ($preview.Stdout + $preview.Stderr + $preview.Result.ErrorMessage)
        @(@($preview.Result.Output)[-1].Episodes).Count | Should -Be 6
        $download = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Redownload -EpisodeId $legacy.EpisodeId
        $download.Result.Succeeded | Should -BeTrue -Because ($download.Stdout + $download.Stderr + $download.Result.ErrorMessage)
        $state = Get-UpdLegacyState -Archive $legacy
        @($state.episodes).Count | Should -Be 1
        $state.episodes[0].episode_id | Should -Be $legacy.EpisodeId
        $state.episodes[0].status | Should -Be 'transfer_verified'
        $state.episodes[0].relative_path | Should -Not -Be $historicalName
        Assert-UpdLegacyOriginal -Archive $legacy
        (Get-UpdFixtureState -Context $context).'/media/ok.mp3' | Should -Be 1
    }

    It 'A017 A018 redownloading an adopted episode preserves its bound file and rollback restores adopted confidence' {
        $adopt = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Adopt -EpisodeId $legacy.EpisodeId -File $script:legacyName
        $adopt.Result.Succeeded | Should -BeTrue -Because ($adopt.Stdout + $adopt.Stderr + $adopt.Result.ErrorMessage)
        $download = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Redownload -EpisodeId $legacy.EpisodeId
        $download.Result.Succeeded | Should -BeTrue -Because ($download.Stdout + $download.Stderr + $download.Result.ErrorMessage)
        $state = Get-UpdLegacyState -Archive $legacy
        $state.episodes[0].status | Should -Be 'transfer_verified'
        $state.episodes[0].relative_path | Should -Not -Be $script:legacyName
        Assert-UpdLegacyOriginal -Archive $legacy
        $checkpoint = @($download.Result.Output)[-1].Checkpoint
        $rollback = Invoke-UpdLegacyWorker -Context $context -Archive $legacy -Action Rollback -Checkpoint $checkpoint
        $rollback.Result.Succeeded | Should -BeTrue -Because ($rollback.Stdout + $rollback.Stderr + $rollback.Result.ErrorMessage)
        $restored = Get-UpdLegacyState -Archive $legacy
        $restored.episodes[0].status | Should -Be 'adopted'
        $restored.episodes[0].relative_path | Should -Be $script:legacyName
        Test-Path -LiteralPath (Join-Path $legacy.Root $state.episodes[0].relative_path) | Should -BeTrue
        Assert-UpdLegacyOriginal -Archive $legacy
        (Get-UpdFixtureState -Context $context).'/media/history.mp3' | Should -Be 1
    }
}
