BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    $git = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python -or -not $git) { throw 'Development-only release smoke fixtures require Python and Git.' }
    $script:releaseGit = $git.Source
    $script:releasePython = $python.Source
    $script:releaseEngine = Join-Path $PSHOME $(if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' })

    function Invoke-UpdReleaseGit {
        param([string]$Root, [string[]]$Arguments)
        $output = @(& $script:releaseGit -c core.autocrlf=false -c core.safecrlf=false -C $Root @Arguments 2>&1)
        if ($LASTEXITCODE -ne 0) { throw ('Owned release Git fixture failed: ' + ($output -join "`n")) }
        return $output
    }

    function Invoke-UpdReleaseBuild {
        param($Context, [string]$SourceRoot, [string]$OutputName)
        $child = Start-UpdOwnedProcess -Context $Context -FilePath $script:releaseEngine -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
            (Join-Path $SourceRoot 'scripts/Build-Release.ps1'), '-OutputDirectory', (Join-Path $Context.Root $OutputName)
        )
        if (-not $child.Process.WaitForExit(60000)) { throw 'Owned release build exceeded its bounded wait.' }
        if ($child.Process.ExitCode -ne 0) { throw ('Owned release build failed: ' + $child.Output.Result + $child.ErrorOutput.Result) }
        return ($child.Output.Result | ConvertFrom-Json)
    }

    function Invoke-UpdReleaseObservation {
        param($Context, [ValidateSet('Import', 'Portable', 'HelpAndLauncher')][string]$Action)
        $id = [guid]::NewGuid().ToString('N')
        $configPath = Join-Path $Context.Root ($id + '-config.json')
        $resultPath = Join-Path $Context.Root ($id + '-result.json')
        $config = @{ Root = $Context.Root; Token = $Context.Token; BaseUrl = $Context.BaseUrl; Action = $Action;
            ProductScript = Join-Path $package 'UniversalPodcastDownloader.ps1'; ResultPath = $resultPath;
            OutputPath = Join-Path $Context.Root ($id.Substring(0, 6) + '-a') }
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        $workerName = if ($Action -eq 'HelpAndLauncher') { 'Invoke-ReleasePackageWorker.ps1' } else { 'Invoke-ArchitectureWorker.ps1' }
        $workerPath = Join-Path $Context.Root $workerName
        if (-not [IO.File]::Exists($workerPath)) {
            [IO.File]::Copy((Join-Path $repositoryRoot ('tests/support/' + $workerName)), $workerPath)
        }
        $child = Start-UpdOwnedProcess -Context $Context -FilePath $script:releaseEngine -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $workerPath, '-ConfigPath', $configPath
        )
        if (-not $child.Process.WaitForExit(45000)) { throw 'Owned extracted release worker exceeded its bounded wait.' }
        if (-not [IO.File]::Exists($resultPath)) { throw ('Extracted release worker produced no report: ' + $child.Output.Result + $child.ErrorOutput.Result) }
        [pscustomobject]@{ ExitCode = $child.Process.ExitCode; Stdout = $child.Output.Result; Stderr = $child.ErrorOutput.Result;
            Result = ([IO.File]::ReadAllText($resultPath) | ConvertFrom-Json); OutputPath = $config.OutputPath }
    }

    function Get-UpdReleaseEntryContent {
        param($Entry)
        $stream = $Entry.Open()
        $memory = [IO.MemoryStream]::new()
        try { $stream.CopyTo($memory); return ,$memory.ToArray() }
        finally { $memory.Dispose(); $stream.Dispose() }
    }

    function Get-UpdReleaseBytesHash {
        param([byte[]]$Bytes)
        $sha = [Security.Cryptography.SHA256]::Create()
        try { return ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant() }
        finally { $sha.Dispose() }
    }
}

Describe 'A051/A052 release ZIP workflow and layered extraction support' {
    BeforeAll {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        $sourceRoot = Join-Path $context.Root 'committed source'
        $null = [IO.Directory]::CreateDirectory($sourceRoot)
        # Snapshot current tracked source and explicit intended packaging additions.
        # Commit only this marked synthetic copy; never require or alter a clean
        # developer checkout, or package a pretend runtime in place of the ZIP.
        $sourceFiles = @(Invoke-UpdReleaseGit -Root $repositoryRoot -Arguments @('ls-files')) + @(
            'scripts/Build-Release.ps1', 'tools/release-package.json', 'CHANGELOG.md', 'RELEASE.md'
        )
        $sourceFiles = @($sourceFiles | Sort-Object -Unique)
        foreach ($relative in $sourceFiles) {
            $target = Join-Path $sourceRoot $relative
            $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
            [IO.File]::Copy((Join-Path $repositoryRoot $relative), $target)
        }
        [IO.File]::AppendAllText((Join-Path $sourceRoot '.gitignore'), "`n/poison/`n", [Text.UTF8Encoding]::new($false))
        $null = Invoke-UpdReleaseGit -Root $sourceRoot -Arguments @('init', '--quiet')
        for ($index = 0; $index -lt $sourceFiles.Count; $index += 40) {
            $last = [Math]::Min($index + 39, $sourceFiles.Count - 1)
            $null = Invoke-UpdReleaseGit -Root $sourceRoot -Arguments (@('add', '--') + $sourceFiles[$index..$last])
        }
        $null = Invoke-UpdReleaseGit -Root $sourceRoot -Arguments @('-c', 'user.name=UPD owned release fixture',
            '-c', 'user.email=upd-fixture@example.invalid', '-c', 'commit.gpgsign=false', 'commit', '--quiet', '-m', 'Owned synthetic release source')
        $script:releaseSourceCommit = (Invoke-UpdReleaseGit -Root $sourceRoot -Arguments @('rev-parse', 'HEAD')) -join ''
        $script:releaseSourceTree = (Invoke-UpdReleaseGit -Root $sourceRoot -Arguments @('rev-parse', 'HEAD^{tree}')) -join ''
        $poisonFiles = @('.local/private-feed-url.txt', '.dev-tools/private-config.json', 'private.log', 'unknown.part',
            'artifacts/previous.zip', 'poison/media.mp3', 'poison/history/state.json', 'poison/shows.json')
        foreach ($relative in $poisonFiles) {
            $target = Join-Path $sourceRoot $relative
            $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
            [IO.File]::WriteAllText($target, 'UPD_PRIVATE_PACKAGE_CANARY_http://private.invalid/?token=SECRET')
        }
        if (@(Invoke-UpdReleaseGit -Root $sourceRoot -Arguments @('status', '--porcelain')).Count -ne 0) {
            throw 'The synthetic committed release source must remain clean despite ignored poison inputs.'
        }
        $script:releaseFirstBuild = Invoke-UpdReleaseBuild -Context $context -SourceRoot $sourceRoot -OutputName 'build one'
        $script:releaseSecondBuild = Invoke-UpdReleaseBuild -Context $context -SourceRoot $sourceRoot -OutputName 'build two'
        $expectedPayload = @('UniversalPodcastDownloader.ps1', 'UniversalPodcastDownloader.bat', 'LICENSE', 'README.md',
            'CHANGELOG.md', 'RELEASE.md', 'UniversalPodcastDownloader.ico', 'UniversalPodcastDownloader_icon.png',
            'UniversalPodcastDownloader_poster.png') + @(Get-ChildItem -LiteralPath (Join-Path $sourceRoot 'src') -File -Filter '*.ps1' | ForEach-Object { 'src/' + $_.Name })
        [Array]::Sort($expectedPayload, [StringComparer]::Ordinal)
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        Start-UpdFixtureServer -Context $context -PythonPath $script:releasePython
    }
    BeforeEach {
        $package = Join-Path $context.Root ('extracted application ' + [char]0x00e5 + ' ' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        [IO.Compression.ZipFile]::ExtractToDirectory($script:releaseFirstBuild.zipPath, $package)
    }
    AfterAll { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'builds two traceable candidates with exact sorted inventory, original MIT and validated hashes excluding owned poison and development history' {
        $expectedPayload.Count | Should -Be 36
        $script:releaseSourceCommit | Should -Match '^[0-9a-f]{40}$'
        $script:releaseFirstBuild.sourceCommit | Should -BeExactly $script:releaseSourceCommit
        $script:releaseFirstBuild.sourceTree | Should -BeExactly $script:releaseSourceTree
        $script:releaseSecondBuild.sourceCommit | Should -BeExactly $script:releaseSourceCommit
        $script:releaseFirstBuild.version | Should -BeExactly '0.1.0-rc.1'
        foreach ($build in @($script:releaseFirstBuild, $script:releaseSecondBuild)) {
            $manifestBytes = [IO.File]::ReadAllBytes($build.manifestPath)
            $manifest = [Text.Encoding]::UTF8.GetString($manifestBytes) | ConvertFrom-Json
            $manifest.schemaVersion | Should -Be 1
            $manifest.packageName | Should -BeExactly 'UniversalPodcastDownloader'
            $manifest.version | Should -BeExactly $script:releaseFirstBuild.version
            $manifest.releaseStatus | Should -BeExactly 'UNRELEASED_CANDIDATE'
            $manifest.sourceCommit | Should -BeExactly $script:releaseSourceCommit
            $manifest.sourceTree | Should -BeExactly $script:releaseSourceTree
            $manifest.sourceRepository | Should -BeExactly 'https://github.com/PikkuJanne/UniversalPodcastDownloader'
            $manifest.sourceCommitUrl | Should -BeExactly ('https://github.com/PikkuJanne/UniversalPodcastDownloader/commit/' + $script:releaseSourceCommit)
            ($manifest.files.path -join "`n") | Should -BeExactly ($expectedPayload -join "`n")
            $records = [IO.File]::ReadAllLines($build.checksumPath)
            $records.Count | Should -Be 2
            @(($records | ForEach-Object { ($_ -split '  ', 2)[1] }) | Select-Object -Unique).Count | Should -Be 2
            foreach ($record in $records) {
                $record | Should -Match '^([0-9a-f]{64})  ([^\\/]+)$'
                $parts = $record -split '  ', 2
                $parts[1] | Should -BeIn @([IO.Path]::GetFileName($build.zipPath), 'manifest.json')
                (Get-FileHash -LiteralPath (Join-Path $build.outputDirectory $parts[1]) -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $parts[0]
            }
            $archive = [IO.Compression.ZipFile]::OpenRead($build.zipPath)
            try {
                $names = @($archive.Entries.FullName)
                $names.Count | Should -Be ($expectedPayload.Count + 1)
                @($names | Select-Object -Unique).Count | Should -Be $names.Count
                @($names | Where-Object { $_ -match '(^/|\\|(^|/)\.\.?(/|$)|:|/$)' }).Count | Should -Be 0
                $sortedExpected = @($expectedPayload) + 'manifest.json'
                [Array]::Sort($sortedExpected, [StringComparer]::Ordinal)
                ($names -join "`n") | Should -BeExactly ($sortedExpected -join "`n")
                (Get-UpdReleaseBytesHash -Bytes (Get-UpdReleaseEntryContent -Entry $archive.GetEntry('manifest.json'))) | Should -BeExactly (Get-UpdReleaseBytesHash -Bytes $manifestBytes)
                foreach ($file in $manifest.files) {
                    $entry = $archive.GetEntry($file.path)
                    $entry.LastWriteTime.DateTime | Should -Be ([datetime]'1980-01-01T00:00:00')
                    $bytes = Get-UpdReleaseEntryContent -Entry $entry
                    $entry.Length | Should -Be $file.size
                    $bytes.Length | Should -Be $file.size
                    (Get-UpdReleaseBytesHash -Bytes $bytes) | Should -BeExactly $file.sha256
                    (Get-FileHash -LiteralPath (Join-Path $sourceRoot $file.path) -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $file.sha256
                    [Text.Encoding]::UTF8.GetString($bytes) | Should -Not -Match 'UPD_PRIVATE_PACKAGE_CANARY'
                }
                @($names | Where-Object { $_ -match '(?i)(^\.git/|^tests/|^tools/|^scripts/|^docs/|\.log$|\.part|\.mp3$|state\.json|shows\.json|(^|/)history/|fixture|secret)' }).Count | Should -Be 0
            }
            finally { $archive.Dispose() }
        }
        (Get-UpdReleaseBytesHash -Bytes ([IO.File]::ReadAllBytes($script:releaseFirstBuild.manifestPath))) | Should -BeExactly (Get-UpdReleaseBytesHash -Bytes ([IO.File]::ReadAllBytes($script:releaseSecondBuild.manifestPath)))
        (Get-FileHash -LiteralPath (Join-Path $package 'LICENSE') -Algorithm SHA256).Hash | Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $repositoryRoot 'LICENSE') -Algorithm SHA256).Hash
        [IO.File]::ReadAllText((Join-Path $package 'LICENSE')) | Should -Match 'MIT License'
        foreach ($relative in $poisonFiles) { Test-Path -LiteralPath (Join-Path $sourceRoot $relative) | Should -BeTrue }
        $entryAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $package 'UniversalPodcastDownloader.ps1'), [ref]$null, [ref]$null)
        $imports = @($entryAst.FindAll({ param($node) $node -is [Management.Automation.Language.StringConstantExpressionAst] -and $node.Value -match '^src/.+\.ps1$' }, $true).Value)
        $imports.Count | Should -Be 24
        foreach ($relative in $imports) { $relative | Should -BeIn $expectedPayload }
        foreach ($nested in @('src/Preflight.ps1', 'src/ResumePolicy.ps1', 'src/TransportPolicy.ps1')) { $nested | Should -BeIn $expectedPayload }
    }

    It 'imports an actual freshly extracted ZIP outside the checkout without initialization or caller-state changes' {
        $package | Should -Match ('extracted application ' + [char]0x00e5)
        $package.StartsWith($repositoryRoot, [StringComparison]::OrdinalIgnoreCase) | Should -BeFalse
        $run = Invoke-UpdReleaseObservation -Context $context -Action Import
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.HostSurvived | Should -BeTrue
        $run.Result.ImportOutput | Should -Be 0
        $run.Result.ImportCallCount | Should -Be 0
        $run.Result.ImportChangedTree | Should -BeFalse
        $run.Result.ImportPowerType | Should -BeFalse
        $run.Result.ImportCreatedPolicy | Should -BeFalse
        $run.Result.ImportCreatedPageLimit | Should -BeFalse
        $run.Result.PreferencesPreserved | Should -BeTrue
        $run.Result.PackageChanged | Should -BeFalse
        $run.Result.RuntimePythonAvailable | Should -BeFalse
    }

    It 'downloads original loopback bytes from the extracted ZIP then verifies repeat and no-write preview with closed handles' {
        $requestsBefore = (Get-UpdFixtureState -Context $context).'/media/ok.mp3'
        $run = Invoke-UpdReleaseObservation -Context $context -Action Portable
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        @($run.Result.Runs).Count | Should -Be 3
        @($run.Result.Runs | Where-Object ExitCode -ne 0).Count | Should -Be 0 -Because ($run.Stdout + $run.Stderr + ($run.Result.Runs | ConvertTo-Json -Depth 20))
        $run.Result.Runs[0].Downloaded | Should -Be 1
        $run.Result.Runs[1].VerifiedSkipped | Should -Be 1
        $run.Result.Runs[2].Preview | Should -BeTrue
        @($run.Result.Runs | Where-Object ExitCode -ne 0).Count | Should -Be 0
        @($run.Result.CorrelationCounts | Where-Object { $_ -ne 0 }).Count | Should -Be 0
        $run.Result.HandlesClosed | Should -BeTrue
        $run.Result.PreviewChangedTree | Should -BeFalse
        $run.Result.PackageChanged | Should -BeFalse
        $run.Result.RuntimePythonAvailable | Should -BeFalse
        $run.Result.RunCreatedOptionState | Should -BeFalse
        $run.Result.RunPreferencesPreserved | Should -BeTrue
        $media = @(Get-ChildItem -LiteralPath $run.OutputPath -Recurse -File -Filter '*.mp3')
        $media.Count | Should -Be 1
        (Get-FileHash -LiteralPath $media[0].FullName -Algorithm SHA256).Hash | Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3') -Algorithm SHA256).Hash
        ((Get-UpdFixtureState -Context $context).'/media/ok.mp3' - [int]$requestsBefore) | Should -Be 1
    }

    It 'renders packaged help with Git Python Pester and analyzer unavailable and executes the real batch preview and fatal exit paths' {
        $requestsBefore = (Get-UpdFixtureState -Context $context).'/media/ok.mp3'
        $run = Invoke-UpdReleaseObservation -Context $context -Action HelpAndLauncher
        $run.ExitCode | Should -Be 0 -Because ($run.Stdout + $run.Stderr + $run.Result.Error)
        $run.Result.Engine | Should -BeExactly $PSVersionTable.PSVersion.ToString()
        $run.Result.RuntimeGitAvailable | Should -BeFalse
        $run.Result.RuntimePythonAvailable | Should -BeFalse
        $run.Result.RuntimePesterAvailable | Should -BeFalse
        $run.Result.RuntimeAnalyzerAvailable | Should -BeFalse
        $run.Result.HostStartup.ExitCode | Should -Be 0
        $run.Result.NativeHostCachePath | Should -BeExactly (Join-Path $context.Root 'home/AppData/Local/Microsoft/Windows/PowerShell/StartupProfileData-NonInteractive')
        @($run.Result.HostStartupTreeDelta | Where-Object {
            $_.InputObject -notlike 'directory:*' -and
            -not $_.InputObject.StartsWith($run.Result.NativeHostCachePath + ':', [StringComparison]::OrdinalIgnoreCase)
        }).Count | Should -Be 0 -Because ($run.Result.HostStartupTreeDelta | ConvertTo-Json -Depth 20)
        $run.Result.HelpAuthored | Should -BeTrue
        $run.Result.HelpExampleCount | Should -BeGreaterOrEqual 1
        $run.Result.Help | Should -Match 'SYNOPSIS'
        $run.Result.Help | Should -Match 'PARAMETERS'
        $run.Result.Help | Should -Match 'EXAMPLE'
        $run.Result.Preview.ExitCode | Should -Be 0 -Because ($run.Result.Preview.Stdout + $run.Result.Preview.Stderr)
        $run.Result.Fatal.ExitCode | Should -Be 1
        $run.Result.Preview.Stdout | Should -Match 'Update archive history and logs; download selected missing episodes'
        $run.Result.Fatal.Stdout | Should -Not -Match '\[OK\]|Done\. You can close'
        $run.Result.PreviewChangedTree | Should -BeFalse -Because ($run.Result.PreviewTreeDelta | ConvertTo-Json -Depth 20)
        $run.Result.PackageChanged | Should -BeFalse
        Test-Path -LiteralPath $run.OutputPath | Should -BeFalse
        ((Get-UpdFixtureState -Context $context).'/media/ok.mp3' - [int]$requestsBefore) | Should -Be 0
    }
}
