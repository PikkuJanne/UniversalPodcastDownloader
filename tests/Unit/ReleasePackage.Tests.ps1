Describe 'A051/A052 release packaging safety and traceability' -Tag 'Unit', 'A051', 'A052' {
    BeforeAll {
        $script:PackageRepo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $script:PackageEngine = (Get-Process -Id $PID).Path
        $script:PackageBuilder = [IO.File]::ReadAllBytes((Join-Path $script:PackageRepo 'scripts/Build-Release.ps1'))
        $fixtureConfig = [IO.File]::ReadAllText((Join-Path $script:PackageRepo 'tools/release-package.json')) | ConvertFrom-Json
        $fixtureConfig.version = '0.1.0-rc.1'
        $fixtureConfig.releaseStatus = 'UNRELEASED_CANDIDATE'
        $script:PackageConfig = $fixtureConfig | ConvertTo-Json -Depth 5
        $script:PackagePaths = @($fixtureConfig.files)
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        function Invoke-UpdPackageFixtureGit {
            param([string[]]$Arguments)
            $result = & git --no-replace-objects -C $script:PackageFixture @Arguments 2>&1
            if ($LASTEXITCODE -ne 0) { throw "Fixture Git failed: $result" }
            return ($result -join "`n").Trim()
        }

        function Save-UpdPackageFixtureCommit {
            param([string[]]$Paths)
            $null = Invoke-UpdPackageFixtureGit -Arguments (@('add', '--') + $Paths)
            $null = Invoke-UpdPackageFixtureGit -Arguments @('-c', 'user.name=UPD fixture', '-c', 'user.email=upd-test@example.invalid',
                '-c', 'commit.gpgsign=false', 'commit', '-q', '-m', 'Owned synthetic release source')
            return Invoke-UpdPackageFixtureGit -Arguments @('rev-parse', 'HEAD')
        }

        function Invoke-UpdPackageBuild {
            param([string[]]$Arguments, [switch]$WrapDiagnosticWords)
            $start = New-Object Diagnostics.ProcessStartInfo
            $start.FileName = $script:PackageEngine
            $values = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
                (Join-Path $script:PackageFixture 'scripts/Build-Release.ps1')) + $Arguments
            $start.Arguments = ($values | ForEach-Object { '"' + $_ + '"' }) -join ' '
            $start.WorkingDirectory = $script:PackageFixture
            $start.UseShellExecute = $false
            $start.CreateNoWindow = $true
            $start.RedirectStandardOutput = $true
            $start.RedirectStandardError = $true
            $start.EnvironmentVariables.Remove('GITHUB_STEP_SUMMARY')
            if ($PSVersionTable.PSEdition -eq 'Desktop') {
                $start.EnvironmentVariables['PSModulePath'] = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules'
            }
            $process = New-Object Diagnostics.Process
            $process.StartInfo = $start
            try {
                [void]$process.Start()
                $stdout = $process.StandardOutput.ReadToEndAsync()
                $stderr = $process.StandardError.ReadToEndAsync()
                if (-not $process.WaitForExit(60000)) {
                    $process.Kill()
                    $process.WaitForExit()
                    throw 'Owned release builder timed out.'
                }
                $data = $null
                if ($process.ExitCode -eq 0) { $data = $stdout.Result | ConvertFrom-Json }
                $rawText = $stdout.Result + $stderr.Result
                # Native error display wraps words according to the host width
                # and source-path length. Selected refusal cases force wrapping
                # of their actual captured diagnostic to exercise that format
                # without depending on local or hosted temporary-path lengths.
                $diagnostic = $stderr.Result
                if ($WrapDiagnosticWords) { $diagnostic = [regex]::Replace($diagnostic, '[ \t]+', "`r`n    ") }
                return [pscustomobject]@{ Code = $process.ExitCode; Data = $data; RawText = $rawText;
                    Text = [regex]::Replace(($stdout.Result + $diagnostic), '\s+', ' ') }
            }
            finally { $process.Dispose() }
        }

        function Invoke-UpdPackageFixtureConfigWrite {
            param([object]$Config)
            [IO.File]::WriteAllText((Join-Path $script:PackageFixture 'tools/release-package.json'),
                ($Config | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
            $null = Save-UpdPackageFixtureCommit -Paths @('tools/release-package.json')
        }
    }

    BeforeEach {
        $script:PackageFixture = Join-Path $TestDrive ([guid]::NewGuid().ToString('N').Substring(0, 8))
        [void][IO.Directory]::CreateDirectory((Join-Path $script:PackageFixture 'scripts'))
        [void][IO.Directory]::CreateDirectory((Join-Path $script:PackageFixture 'tools'))
        [IO.File]::WriteAllBytes((Join-Path $script:PackageFixture 'scripts/Build-Release.ps1'), $script:PackageBuilder)
        [IO.File]::WriteAllText((Join-Path $script:PackageFixture 'tools/release-package.json'), $script:PackageConfig)
        [IO.File]::WriteAllText((Join-Path $script:PackageFixture '.gitignore'), "/artifacts/`n/.dev-tools/`n")
        foreach ($path in $script:PackagePaths) {
            $fullPath = Join-Path $script:PackageFixture $path
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($fullPath))
            [IO.File]::WriteAllBytes($fullPath, [byte[]](0..255))
        }
        $null = Invoke-UpdPackageFixtureGit -Arguments @('init', '-q')
        $null = Invoke-UpdPackageFixtureGit -Arguments @('config', 'core.autocrlf', 'false')
        $script:PackageFixtureCommit = Save-UpdPackageFixtureCommit -Paths @($script:PackagePaths + '.gitignore' +
            'scripts/Build-Release.ps1' + 'tools/release-package.json')
        $script:PackageOutput = Join-Path $script:PackageFixture 'artifacts/releases'
    }

    It 'exports exact committed binary bytes and sorted traceable inventory without tracked or ignored poison' {
        foreach ($path in @('tests/private-feed.txt', 'logs/podcast.log', 'history.json', 'saved-settings.json', 'media/private.mp3')) {
            $fullPath = Join-Path $script:PackageFixture $path
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($fullPath))
            [IO.File]::WriteAllText($fullPath, 'secret synthetic poison')
        }
        $script:PackageFixtureCommit = Save-UpdPackageFixtureCommit -Paths @('tests/private-feed.txt', 'logs/podcast.log',
            'history.json', 'saved-settings.json', 'media/private.mp3')
        [void][IO.Directory]::CreateDirectory((Join-Path $script:PackageFixture '.dev-tools'))
        [IO.File]::WriteAllText((Join-Path $script:PackageFixture '.dev-tools/untracked-secret'), 'ignored poison')
        $result = Invoke-UpdPackageBuild
        $result.Code | Should -Be 0 -Because $result.Text
        $manifest = [IO.File]::ReadAllText($result.Data.manifestPath) | ConvertFrom-Json
        $manifest.files.Count | Should -Be 36
        ($manifest.files.path -join "`n") | Should -Be ($script:PackagePaths -join "`n")
        $manifest.sourceCommit | Should -BeExactly $script:PackageFixtureCommit
        $manifest.sourceCommitUrl | Should -BeExactly ('https://github.com/PikkuJanne/UniversalPodcastDownloader/commit/' + $script:PackageFixtureCommit)
        $manifest.sourceTree | Should -BeExactly (Invoke-UpdPackageFixtureGit -Arguments @('rev-parse', 'HEAD^{tree}'))
        $archive = [IO.Compression.ZipFile]::OpenRead($result.Data.zipPath)
        try {
            $archive.Entries.Count | Should -Be 37
            foreach ($record in $manifest.files) {
                $entry = $archive.GetEntry($record.path)
                $entry.Length | Should -Be 256
                $entry.LastWriteTime.Year | Should -Be 1980
                $stream = $entry.Open()
                $buffer = New-Object IO.MemoryStream
                try {
                    $stream.CopyTo($buffer)
                    [Convert]::ToBase64String($buffer.ToArray()) | Should -BeExactly ([Convert]::ToBase64String([byte[]](0..255)))
                }
                finally { $stream.Dispose(); $buffer.Dispose() }
                $record.sha256 | Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $script:PackageFixture $record.path) -Algorithm SHA256).Hash.ToLowerInvariant()
            }
        }
        finally { $archive.Dispose() }
    }

    It 'rejects <Label> source changes before creating an output' -TestCases @(
        @{ Label = 'unstaged runtime'; Path = 'src/Batch.ps1'; Stage = $false },
        @{ Label = 'staged release metadata'; Path = 'tools/release-package.json'; Stage = $true },
        @{ Label = 'untracked private configuration'; Path = 'private-settings.json'; Stage = $false }
    ) {
        param($Label, $Path, $Stage)
        $Label | Should -Not -BeNullOrEmpty
        [IO.File]::WriteAllText((Join-Path $script:PackageFixture $Path), 'changed synthetic source')
        if ($Stage) { $null = Invoke-UpdPackageFixtureGit -Arguments @('add', '--', $Path) }
        $result = Invoke-UpdPackageBuild -WrapDiagnosticWords
        $result.Code | Should -Be 1
        Test-Path -LiteralPath $script:PackageOutput | Should -BeFalse
        $result.Text | Should -Match 'Release source must be clean'
    }

    It 'rejects a <Label> allowlist committed to the selected source' -TestCases @(
        @{ Label = 'missing module'; Mutation = 'missing' },
        @{ Label = 'duplicate file'; Mutation = 'duplicate' },
        @{ Label = 'forbidden extra secret'; Mutation = 'extra' },
        @{ Label = 'wrong ordering'; Mutation = 'ordering' },
        @{ Label = 'path traversal'; Mutation = 'traversal' }
    ) {
        param($Label, $Mutation)
        $Label | Should -Not -BeNullOrEmpty
        $config = $script:PackageConfig | ConvertFrom-Json
        switch ($Mutation) {
            'missing' { $config.files = @($config.files | Where-Object { $_ -ne 'src/Batch.ps1' }) }
            'duplicate' { $config.files += 'LICENSE' }
            'extra' { $config.files += 'saved-settings.json' }
            'ordering' { [array]::Reverse($config.files) }
            'traversal' { $config.files[0] = '../private-settings.json' }
        }
        Invoke-UpdPackageFixtureConfigWrite -Config $config
        $result = Invoke-UpdPackageBuild
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'exact root and complete src inventory'
        Test-Path -LiteralPath $script:PackageOutput | Should -BeFalse
    }

    It 'rejects an omitted runtime source module even when the allowlist otherwise looks valid' {
        [IO.File]::WriteAllText((Join-Path $script:PackageFixture 'src/Undeclared.ps1'), 'synthetic runtime module')
        $null = Save-UpdPackageFixtureCommit -Paths @('src/Undeclared.ps1')
        $result = Invoke-UpdPackageBuild
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'exactly 27 runtime'
    }

    It 'rejects a missing committed license without substituting working checkout bytes' {
        $null = Invoke-UpdPackageFixtureGit -Arguments @('rm', '-q', '--', 'LICENSE')
        $null = Invoke-UpdPackageFixtureGit -Arguments @('-c', 'user.name=UPD fixture', '-c', 'user.email=upd-test@example.invalid',
            '-c', 'commit.gpgsign=false', 'commit', '-q', '-m', 'Remove owned synthetic license')
        $result = Invoke-UpdPackageBuild -WrapDiagnosticWords
        $result.Code | Should -Be 1
        Test-Path -LiteralPath $script:PackageOutput | Should -BeFalse
        $result.Text | Should -Match 'Missing committed payload file: LICENSE'
    }

    It 'rejects a symlink Git blob even with a clean non-symlink working file' {
        $null = Invoke-UpdPackageFixtureGit -Arguments @('config', 'core.symlinks', 'false')
        [IO.File]::WriteAllText((Join-Path $script:PackageFixture 'LICENSE'), 'README.md')
        $blob = Invoke-UpdPackageFixtureGit -Arguments @('hash-object', '-w', '--', 'LICENSE')
        $null = Invoke-UpdPackageFixtureGit -Arguments @('update-index', '--cacheinfo', ('120000,' + $blob + ',LICENSE'))
        $null = Invoke-UpdPackageFixtureGit -Arguments @('-c', 'user.name=UPD fixture', '-c', 'user.email=upd-test@example.invalid',
            '-c', 'commit.gpgsign=false', 'commit', '-q', '-m', 'Owned synthetic symlink index')
        (Invoke-UpdPackageFixtureGit -Arguments @('status', '--porcelain')) | Should -BeNullOrEmpty
        $result = Invoke-UpdPackageBuild
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'regular Git blob: LICENSE'
    }

    It 'rejects <Label> SourceCommit values' -TestCases @(
        @{ Label = 'short ID'; Value = 'abc123' },
        @{ Label = 'missing full ID'; Value = '0000000000000000000000000000000000000000' },
        @{ Label = 'revision expression'; Value = 'HEAD~1' }
    ) {
        param($Label, $Value)
        $Label | Should -Not -BeNullOrEmpty
        $result = Invoke-UpdPackageBuild -Arguments @('-SourceCommit', $Value)
        $result.Code | Should -Be 1
        Test-Path -LiteralPath $script:PackageOutput | Should -BeFalse
    }

    It 'uses the selected earlier commit version and bytes rather than current committed metadata' {
        $config = $script:PackageConfig | ConvertFrom-Json
        $config.version = '0.1.0-rc.2'
        Invoke-UpdPackageFixtureConfigWrite -Config $config
        $result = Invoke-UpdPackageBuild -Arguments @('-SourceCommit', $script:PackageFixtureCommit)
        $result.Code | Should -Be 0 -Because $result.Text
        $result.Data.version | Should -BeExactly '0.1.0-rc.1'
        $result.Data.sourceCommit | Should -BeExactly $script:PackageFixtureCommit
    }

    It 'builds a stable <Version> package from its exact committed status and source' -TestCases @(
        @{ Version = '1.0.0' }, @{ Version = '0.0.0' }
    ) {
        param($Version)
        $config = $script:PackageConfig | ConvertFrom-Json
        $config.version = $Version
        $config.releaseStatus = 'STABLE_RELEASE'
        Invoke-UpdPackageFixtureConfigWrite -Config $config
        $result = Invoke-UpdPackageBuild
        $result.Code | Should -Be 0 -Because $result.Text
        $result.Data.version | Should -BeExactly $Version
        $result.Data.sourceCommit | Should -BeExactly (Invoke-UpdPackageFixtureGit -Arguments @('rev-parse', 'HEAD'))
        [IO.Path]::GetFileName($result.Data.zipPath) | Should -BeExactly ('UniversalPodcastDownloader-' + $Version + '.zip')
        $manifest = [IO.File]::ReadAllText($result.Data.manifestPath) | ConvertFrom-Json
        $manifest.version | Should -BeExactly $Version
        $manifest.releaseStatus | Should -BeExactly 'STABLE_RELEASE'
        $manifest.files.Count | Should -Be 36
    }

    It 'rejects <Label> version and status pairs before creating output' -TestCases @(
        @{ Label = 'stable version with candidate status'; Version = '1.0.0'; Status = 'UNRELEASED_CANDIDATE' },
        @{ Label = 'rc version with stable status'; Version = '0.1.0-rc.1'; Status = 'STABLE_RELEASE' },
        @{ Label = 'unsupported release status'; Version = '1.0.0'; Status = 'RELEASED' },
        @{ Label = 'version path traversal'; Version = '../0.1.0-rc.1'; Status = 'UNRELEASED_CANDIDATE' },
        @{ Label = 'zero rc ordinal'; Version = '0.1.0-rc.0'; Status = 'UNRELEASED_CANDIDATE' },
        @{ Label = 'leading zero rc ordinal'; Version = '0.1.0-rc.01'; Status = 'UNRELEASED_CANDIDATE' },
        @{ Label = 'leading zero major'; Version = '01.0.0'; Status = 'STABLE_RELEASE' },
        @{ Label = 'leading zero minor'; Version = '1.00.0'; Status = 'STABLE_RELEASE' },
        @{ Label = 'leading zero patch'; Version = '1.0.00'; Status = 'STABLE_RELEASE' },
        @{ Label = 'unsupported prerelease'; Version = '1.0.0-beta.1'; Status = 'STABLE_RELEASE' },
        @{ Label = 'unsupported build metadata'; Version = '1.0.0+build.1'; Status = 'STABLE_RELEASE' },
        @{ Label = 'trailing version line break'; Version = "1.0.0`n"; Status = 'STABLE_RELEASE' }
    ) {
        param($Label, $Version, $Status)
        $Label | Should -Not -BeNullOrEmpty
        $config = $script:PackageConfig | ConvertFrom-Json
        $config.version = $Version
        $config.releaseStatus = $Status
        Invoke-UpdPackageFixtureConfigWrite -Config $config
        $result = Invoke-UpdPackageBuild
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'Release version and status must pair'
        Test-Path -LiteralPath $script:PackageOutput | Should -BeFalse
    }

    It 'rejects <Label> output paths without changing source files' -TestCases @(
        @{ Label = 'repository root'; PathKind = 'repo' },
        @{ Label = 'tracked src directory'; PathKind = 'src' },
        @{ Label = 'repository ancestor'; PathKind = 'parent' },
        @{ Label = 'volume root'; PathKind = 'volume' },
        @{ Label = 'existing file'; PathKind = 'file' }
    ) {
        param($Label, $PathKind)
        $Label | Should -Not -BeNullOrEmpty
        $path = switch ($PathKind) {
            'repo' { $script:PackageFixture }
            'src' { Join-Path $script:PackageFixture 'src' }
            'parent' { Split-Path $script:PackageFixture -Parent }
            'volume' { [IO.Path]::GetPathRoot($script:PackageFixture) }
            'file' { Join-Path $script:PackageFixture 'LICENSE' }
        }
        $originalHash = (Get-FileHash -LiteralPath (Join-Path $script:PackageFixture 'LICENSE')).Hash
        $result = Invoke-UpdPackageBuild -Arguments @('-OutputDirectory', $path)
        $result.Code | Should -Be 1
        (Get-FileHash -LiteralPath (Join-Path $script:PackageFixture 'LICENSE')).Hash | Should -BeExactly $originalHash
    }

    It 'rejects a junction output ancestor and preserves its unknown target' {
        $target = Join-Path $TestDrive ('outside-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $link = Join-Path $script:PackageFixture 'artifacts'
        [void][IO.Directory]::CreateDirectory($target)
        [IO.File]::WriteAllText((Join-Path $target 'unknown.txt'), 'preserve')
        $null = New-Item -ItemType Junction -Path $link -Target $target
        try {
            $result = Invoke-UpdPackageBuild
            $result.Code | Should -Be 1
            $result.Text | Should -Match 'Reparse point'
            [IO.File]::ReadAllText((Join-Path $target 'unknown.txt')) | Should -BeExactly 'preserve'
            @(Get-ChildItem -LiteralPath $target).Count | Should -Be 1
        }
        finally { [IO.Directory]::Delete($link) }
    }

    It 'never overwrites an existing candidate directory or unknown file' {
        $candidate = Join-Path $script:PackageOutput ('UniversalPodcastDownloader-0.1.0-rc.1-' + $script:PackageFixtureCommit)
        [void][IO.Directory]::CreateDirectory($candidate)
        [IO.File]::WriteAllText((Join-Path $candidate 'unknown.txt'), 'preserve')
        $result = Invoke-UpdPackageBuild
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'already exists; nothing overwritten'
        [IO.File]::ReadAllText((Join-Path $candidate 'unknown.txt')) | Should -BeExactly 'preserve'
        @(Get-ChildItem -LiteralPath $candidate).Count | Should -Be 1
        @(Get-ChildItem -LiteralPath $script:PackageOutput -Force).Count | Should -Be 1
    }

    It 'checksums identify later tampering of both ZIP and sidecar manifest' {
        $result = Invoke-UpdPackageBuild
        $result.Code | Should -Be 0 -Because $result.Text
        $sums = [IO.File]::ReadAllLines($result.Data.checksumPath)
        $sums.Count | Should -Be 2
        foreach ($index in @(0, 1)) {
            $line = $sums[$index] -split '  ', 2
            $path = Join-Path $result.Data.outputDirectory $line[1]
            (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $line[0]
            $stream = [IO.File]::Open($path, [IO.FileMode]::Append)
            try { $stream.WriteByte(42) } finally { $stream.Dispose() }
            (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() | Should -Not -BeExactly $line[0]
        }
    }
}
