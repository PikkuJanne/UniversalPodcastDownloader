Describe 'A053/A054 release workflow policy and candidate readiness' -Tag 'Unit', 'A053', 'A054' {
    BeforeAll {
        $script:WorkflowRepo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $script:WorkflowEngine = (Get-Process -Id $PID).Path
        $script:WorkflowText = [IO.File]::ReadAllText((Join-Path $script:WorkflowRepo '.github/workflows/test.yml'))

        # This deliberately limited reader is a fail-closed policy regression
        # helper for this repository's block YAML. It is not a general YAML
        # parser or an independent GitHub Actions syntax validator.
        function ConvertFrom-UpdWorkflowPolicyYaml {
            param([string]$Text)
            function ConvertFrom-UpdPolicyScalar {
                param([string]$Value)
                $valueText = $Value.Trim()
                if ($valueText -eq '{}') { return @{} }
                if ($valueText -match '^\[([a-zA-Z0-9_, -]*)\]$') {
                    return ,@($Matches[1].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                }
                if ($valueText -match '^''(?:[^'']|'''')*''$') { return $valueText.Substring(1, $valueText.Length - 2).Replace("''", "'") }
                if ($valueText -match '^"[^"\\]*"$') { return $valueText.Substring(1, $valueText.Length - 2) }
                if ($valueText -match '^(true|false)$') { return $valueText -eq 'true' }
                if ($valueText -match '^[0-9]+$') { return [int]$valueText }
                if ($valueText -match '^[&*!|>\[{]' -or $valueText -match '^---|^\.\.\.' -or $valueText -match ':\s') {
                    throw "Unsupported workflow policy scalar: $valueText"
                }
                return $valueText
            }
            function Read-UpdPolicyNode {
                param([object[]]$Rows, [int]$Start, [int]$Indent)
                $sequence = $Rows[$Start].Text.StartsWith('- ')
                $node = if ($sequence) { New-Object Collections.ArrayList } else { @{} }
                $index = $Start
                while ($index -lt $Rows.Count -and $Rows[$index].Indent -eq $Indent) {
                    $row = $Rows[$index]
                    $itemText = $row.Text
                    if ($sequence) {
                        if (-not $itemText.StartsWith('- ')) { throw 'Mixed sequence and mapping in workflow policy YAML.' }
                        $itemText = $itemText.Substring(2)
                        if ($itemText -notmatch '^([a-zA-Z_][a-zA-Z0-9_-]*):(?:\s+(.*))?$') {
                            [void]$node.Add((ConvertFrom-UpdPolicyScalar $itemText))
                            $index++
                            continue
                        }
                        $key = $Matches[1]; $valueText = $Matches[2]
                        $item = @{}
                        if ($null -ne $row.Block) { $item[$key] = $row.Block }
                        elseif ($valueText) { $item[$key] = ConvertFrom-UpdPolicyScalar $valueText }
                        else { throw 'Nested sequence key must have an explicit scalar in workflow policy YAML.' }
                        $index++
                        if ($index -lt $Rows.Count -and $Rows[$index].Indent -gt $Indent) {
                            if ($Rows[$index].Indent -ne ($Indent + 2)) { throw 'Unsupported sequence indentation in workflow policy YAML.' }
                            $more = Read-UpdPolicyNode -Rows $Rows -Start $index -Indent ($Indent + 2)
                            if ($more.Value -isnot [Collections.IDictionary]) { throw 'Expected mapping after workflow sequence key.' }
                            foreach ($moreKey in $more.Value.Keys) {
                                if ($item.ContainsKey($moreKey)) { throw "Duplicate workflow policy key: $moreKey" }
                                $item[$moreKey] = $more.Value[$moreKey]
                            }
                            $index = $more.Next
                        }
                        [void]$node.Add($item)
                        continue
                    }
                    if ($itemText -notmatch '^([a-zA-Z_][a-zA-Z0-9_-]*):(?:\s+(.*))?$') { throw "Unsupported workflow mapping: $itemText" }
                    $key = $Matches[1]; $valueText = $Matches[2]
                    if ($node.ContainsKey($key)) { throw "Duplicate workflow policy key: $key" }
                    $index++
                    if ($null -ne $row.Block) { $node[$key] = $row.Block }
                    elseif ($valueText) { $node[$key] = ConvertFrom-UpdPolicyScalar $valueText }
                    elseif ($index -lt $Rows.Count -and $Rows[$index].Indent -gt $Indent) {
                        if ($Rows[$index].Indent -ne ($Indent + 2)) { throw 'Unsupported mapping indentation in workflow policy YAML.' }
                        $child = Read-UpdPolicyNode -Rows $Rows -Start $index -Indent ($Indent + 2)
                        $node[$key] = $child.Value; $index = $child.Next
                    }
                    else { $node[$key] = $null }
                }
                if ($index -lt $Rows.Count -and $Rows[$index].Indent -gt $Indent) { throw 'Unexpected nested workflow policy content.' }
                return @{ Value = $node; Next = $index }
            }
            $rawLines = @($Text -split '\r?\n')
            $rows = New-Object Collections.ArrayList
            for ($lineIndex = 0; $lineIndex -lt $rawLines.Count; $lineIndex++) {
                $line = $rawLines[$lineIndex]
                if ($line.Contains("`t")) { throw 'Tabs are unsupported in workflow policy YAML.' }
                if ($line -match '^\s*(#.*)?$') { continue }
                $indent = $line.Length - $line.TrimStart().Length
                $lineText = $line.TrimStart()
                $block = $null
                if ($lineText -match '^([a-zA-Z_][a-zA-Z0-9_-]*):\s+\|$') {
                    $key = $Matches[1]
                    $blockLines = New-Object Collections.ArrayList
                    while (($lineIndex + 1) -lt $rawLines.Count) {
                        $next = $rawLines[$lineIndex + 1]
                        $nextIndent = $next.Length - $next.TrimStart().Length
                        if ($next.Trim() -and $nextIndent -le $indent) { break }
                        $lineIndex++
                        if ($next.Trim() -and $nextIndent -lt ($indent + 2)) { throw 'Invalid workflow literal block indentation.' }
                        [void]$blockLines.Add($(if ($next.Length -ge ($indent + 2)) { $next.Substring($indent + 2) } else { '' }))
                    }
                    $block = $blockLines -join "`n"
                    $lineText = $key + ': |'
                }
                else {
                    # Repository comments appear only outside quoted scalars.
                    # Reject ambiguous quoting rather than silently misread it.
                    if ($lineText -notmatch ':\s+[''\"]' -and $lineText -match '\s+#') { $lineText = ($lineText -split '\s+#', 2)[0] }
                }
                [void]$rows.Add(@{ Indent = $indent; Text = $lineText; Block = $block })
            }
            if (-not $rows.Count -or $rows[0].Indent -ne 0) { throw 'Workflow policy root must be a block mapping.' }
            $parsed = Read-UpdPolicyNode -Rows @($rows) -Start 0 -Indent 0
            if ($parsed.Next -ne $rows.Count -or $parsed.Value -isnot [Collections.IDictionary]) { throw 'Unconsumed workflow policy YAML.' }
            return $parsed.Value
        }

        function Assert-UpdWorkflowPolicy {
            param([string]$Text)
            $workflow = ConvertFrom-UpdWorkflowPolicyYaml $Text
            if ($workflow.permissions -isnot [Collections.IDictionary] -or $workflow.permissions.Count -ne 0) { throw 'Workflow token defaults must deny permissions.' }
            $triggers = @($workflow.on.Keys)
            if ($triggers.Count -ne 3 -or @($triggers | Where-Object { $_ -notin @('push', 'pull_request', 'workflow_dispatch') }).Count) { throw 'Only nonprivileged reviewed triggers are permitted.' }
            if ($workflow.on.push.Count -ne 1 -or @($workflow.on.push.branches).Count -ne 1 -or $workflow.on.push.branches[0] -cne 'main' -or
                $null -ne $workflow.on.pull_request -or $null -ne $workflow.on.workflow_dispatch) { throw 'Trigger scope must exclude tags, inputs and privileged events.' }
            if ($workflow.env.Count -ne 1 -or $workflow.env.UPD_SOURCE_COMMIT -cne '${{ github.event.pull_request.head.sha || github.sha }}') { throw 'Build source must be the selected event commit.' }
            if (-not $workflow.concurrency.group -or $workflow.concurrency.'cancel-in-progress' -ne $true) { throw 'Workflow concurrency must be bounded.' }
            if ($workflow.jobs.Count -ne 2 -or -not $workflow.jobs.ContainsKey('windows') -or -not $workflow.jobs.ContainsKey('candidate')) { throw 'Both engine validation and candidate preparation jobs are required.' }
            $pins = @{
                'actions/checkout' = '3d3c42e5aac5ba805825da76410c181273ba90b1'
                'actions/setup-python' = '5fda3b95a4ea91299a34e894583c3862153e4b97'
                'actions/upload-artifact' = '043fb46d1a93c77aae656e7c1c64a875d1fc6a0a'
            }
            foreach ($job in $workflow.jobs.Values) {
                if ($job.permissions.Count -ne 1 -or $job.permissions.contents -cne 'read') { throw 'Jobs may only read repository contents.' }
                if ($job.'runs-on' -cne 'windows-2022' -or $job.'timeout-minutes' -lt 1 -or $job.'timeout-minutes' -gt 40) { throw 'Jobs must use reviewed bounded Windows runners.' }
                foreach ($unsupported in @('environment', 'container', 'services', 'secrets', 'continue-on-error', 'uses', 'env')) {
                    if ($job.ContainsKey($unsupported)) { throw "Unreviewed job capability: $unsupported" }
                }
                foreach ($step in $job.steps) {
                    if ($step.ContainsKey('continue-on-error') -or $step.ContainsKey('env')) { throw 'Steps cannot suppress failure or introduce credentials.' }
                    if ($step.uses) {
                        if ($step.uses -cnotmatch '^([^@]+)@([0-9a-f]{40})$' -or -not $pins.ContainsKey($Matches[1]) -or $pins[$Matches[1]] -cne $Matches[2]) { throw 'Actions must use reviewed immutable commits.' }
                        if ($step.uses.StartsWith('actions/checkout@')) {
                            if ($step.with.'persist-credentials' -ne $false -or $step.with.ref -cne '${{ env.UPD_SOURCE_COMMIT }}' -or $step.with.ContainsKey('token')) { throw 'Checkout must use the exact source without retained credentials.' }
                        }
                    }
                    if ($step.run) {
                        $tokens = $null; $parseErrors = $null
                        $null = [Management.Automation.Language.Parser]::ParseInput($step.run, [ref]$tokens, [ref]$parseErrors)
                        if ($parseErrors.Count) { throw 'Every workflow PowerShell block must parse in the current supported engine.' }
                        if ($step.run -match '(?im)\b(?:git\s+(?:push|tag)|gh\s+(?:release|repo|api)|Invoke-(?:WebRequest|RestMethod)|GITHUB_TOKEN|GH_TOKEN)\b') { throw 'Workflow cannot publish, alter GitHub settings or consume explicit credentials.' }
                    }
                }
            }
            if ($Text -match '(?i)secrets\s*\.|github\s*\.\s*token') { throw 'Workflow cannot reference secrets or expose the workflow token.' }
            $windows = $workflow.jobs.windows
            if ($windows.strategy.'max-parallel' -ne 2 -or $windows.strategy.'fail-fast' -ne $false -or $windows.strategy.matrix.include.Count -ne 2 -or
                (@($windows.strategy.matrix.include.shell | Sort-Object) -join ',') -cne 'powershell,pwsh') { throw 'Both supported Windows engines must validate before packaging.' }
            $testSteps = @($windows.steps | Where-Object { $_.run -match '(?m)^\./scripts/Test\.ps1 -Suite All\s*$' })
            $analysisSteps = @($windows.steps | Where-Object { $_.run -match '(?m)^\./scripts/Analyze\.ps1\s*$' })
            if ($testSteps.Count -ne 1 -or $analysisSteps.Count -ne 1 -or $testSteps[0].ContainsKey('if') -or
                $analysisSteps[0].if -cne '${{ !cancelled() }}') { throw 'Complete engine tests and analysis are required.' }
            $candidate = $workflow.jobs.candidate
            if (@($candidate.needs).Count -ne 1 -or @($candidate.needs)[0] -cne 'windows' -or $candidate.if -cne '${{ success() && needs.windows.result == ''success'' }}') { throw 'Candidate must wait for successful complete engine checks.' }
            $prepare = @($candidate.steps | Where-Object { $_.id -eq 'prepared' })
            if ($prepare.Count -ne 1 -or $prepare[0].ContainsKey('if') -or $prepare[0].run -notmatch '\./scripts/Prepare-ReleaseCandidate\.ps1' -or
                $prepare[0].run -notmatch '-SourceCommit\s+\$env:UPD_SOURCE_COMMIT' -or $prepare[0].run -notmatch '-GitHubOutput') { throw 'Candidate upload outputs must follow the actual readiness gate.' }
            $uploads = @($candidate.steps | Where-Object { $_.uses -like 'actions/upload-artifact@*' })
            $allUploads = @($workflow.jobs.Values | ForEach-Object { $_.steps } | Where-Object { $_.uses -like 'actions/upload-artifact@*' })
            if ($uploads.Count -ne 1 -or $allUploads.Count -ne 1) { throw 'Exactly one candidate artifact upload is permitted.' }
            $upload = $uploads[0]
            if ($upload.ContainsKey('if') -or $upload.with.'if-no-files-found' -cne 'error' -or $upload.with.'include-hidden-files' -ne $false -or
                $upload.with.overwrite -ne $false -or $upload.with.'retention-days' -ne 7 -or $upload.with.'compression-level' -ne 0) { throw 'Artifact upload must fail closed and retain an immutable bounded candidate.' }
            $paths = @($upload.with.path.Trim() -split '\r?\n')
            $allowedPaths = @('${{ steps.prepared.outputs.zipPath }}', '${{ steps.prepared.outputs.manifestPath }}', '${{ steps.prepared.outputs.checksumPath }}')
            if ($paths.Count -ne 3 -or @($paths | Where-Object { $_ -cnotin $allowedPaths }).Count -or @($paths | Select-Object -Unique).Count -ne 3) { throw 'Only the three validated candidate files may be uploaded.' }
            if ($upload.with.name -cne 'candidate-${{ env.UPD_SOURCE_COMMIT }}-${{ github.run_id }}-${{ github.run_attempt }}') { throw 'Candidate artifact must identify exact source and attempt.' }
            $prepareIndex = [array]::IndexOf(@($candidate.steps), $prepare[0]); $uploadIndex = [array]::IndexOf(@($candidate.steps), $upload)
            if ($prepareIndex -ge $uploadIndex) { throw 'Readiness validation must precede upload.' }
        }

        . (Join-Path $script:WorkflowRepo 'scripts/Prepare-ReleaseCandidate.ps1')
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $script:CandidateFixture = Join-Path $TestDrive 'source'
        [void][IO.Directory]::CreateDirectory((Join-Path $script:CandidateFixture 'scripts'))
        [void][IO.Directory]::CreateDirectory((Join-Path $script:CandidateFixture 'tools'))
        $script:CandidateConfig = [IO.File]::ReadAllText((Join-Path $script:WorkflowRepo 'tools/release-package.json')) | ConvertFrom-Json
        foreach ($path in $script:CandidateConfig.files) {
            $fullPath = Join-Path $script:CandidateFixture $path
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($fullPath))
            [IO.File]::WriteAllBytes($fullPath, [byte[]](0..255))
        }
        foreach ($path in @('scripts/Build-Release.ps1', 'scripts/Prepare-ReleaseCandidate.ps1', 'tools/release-package.json')) {
            [IO.File]::WriteAllBytes((Join-Path $script:CandidateFixture $path), [IO.File]::ReadAllBytes((Join-Path $script:WorkflowRepo $path)))
        }
        [IO.File]::WriteAllText((Join-Path $script:CandidateFixture '.gitignore'), "/artifacts/`n")

        function Invoke-UpdCandidateFixtureGit {
            param([string[]]$Arguments)
            $result = & git --no-replace-objects -C $script:CandidateFixture @Arguments 2>&1
            if ($LASTEXITCODE -ne 0) { throw "Owned candidate fixture Git failed: $result" }
            return ($result -join "`n").Trim()
        }
        $null = Invoke-UpdCandidateFixtureGit @('init', '-q')
        $null = Invoke-UpdCandidateFixtureGit @('config', 'core.autocrlf', 'false')
        $null = Invoke-UpdCandidateFixtureGit (@('add', '--') + @($script:CandidateConfig.files) + @('.gitignore',
            'scripts/Build-Release.ps1', 'scripts/Prepare-ReleaseCandidate.ps1', 'tools/release-package.json'))
        $null = Invoke-UpdCandidateFixtureGit @('-c', 'user.name=UPD fixture', '-c', 'user.email=upd-test@example.invalid',
            '-c', 'commit.gpgsign=false', 'commit', '-q', '-m', 'Owned synthetic release readiness source')
        $script:CandidateCommit = Invoke-UpdCandidateFixtureGit @('rev-parse', 'HEAD')
        $script:CandidateTree = Invoke-UpdCandidateFixtureGit @('rev-parse', 'HEAD^{tree}')

        function Invoke-UpdCandidateFixturePrepare {
            param([string]$Commit, [string]$OutputDirectory, [string]$GitHubOutput)
            $start = New-Object Diagnostics.ProcessStartInfo
            $start.FileName = $script:WorkflowEngine
            $values = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
                (Join-Path $script:CandidateFixture 'scripts/Prepare-ReleaseCandidate.ps1'), '-SourceCommit', $Commit,
                '-OutputDirectory', $OutputDirectory)
            if ($GitHubOutput) { $values += @('-GitHubOutput', $GitHubOutput) }
            $start.Arguments = ($values | ForEach-Object { '"' + $_ + '"' }) -join ' '
            $start.WorkingDirectory = $script:CandidateFixture
            $start.UseShellExecute = $false; $start.CreateNoWindow = $true
            $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
            $start.EnvironmentVariables.Remove('GITHUB_STEP_SUMMARY')
            if ($PSVersionTable.PSEdition -eq 'Desktop') {
                $start.EnvironmentVariables['PSModulePath'] = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules'
            }
            $process = New-Object Diagnostics.Process
            $process.StartInfo = $start
            try {
                [void]$process.Start()
                $stdout = $process.StandardOutput.ReadToEndAsync(); $stderr = $process.StandardError.ReadToEndAsync()
                if (-not $process.WaitForExit(60000)) { $process.Kill(); $process.WaitForExit(); throw 'Owned candidate wrapper timed out.' }
                $data = if ($process.ExitCode -eq 0) { $stdout.Result | ConvertFrom-Json } else { $null }
                return [pscustomobject]@{ Code = $process.ExitCode; Data = $data; Text = [regex]::Replace($stdout.Result + $stderr.Result, '\s+', ' ') }
            }
            finally { $process.Dispose() }
        }
        $script:PreparedOutput = Join-Path $TestDrive 'prepared-outputs.txt'
        [IO.File]::WriteAllText($script:PreparedOutput, 'owned-existing-output' + "`n")
        $baselineDirectory = Join-Path $script:CandidateFixture 'artifacts/baseline'
        $script:PreparedCandidate = Invoke-UpdCandidateFixturePrepare -Commit $script:CandidateCommit -OutputDirectory $baselineDirectory -GitHubOutput $script:PreparedOutput
        if ($script:PreparedCandidate.Code -ne 0) { throw "Synthetic candidate preparation failed: $($script:PreparedCandidate.Text)" }

        function Invoke-UpdCandidateCopy {
            $parent = Join-Path $script:CandidateFixture ('artifacts/probe-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            [void][IO.Directory]::CreateDirectory($parent)
            $destination = Join-Path $parent ([IO.Path]::GetFileName($script:PreparedCandidate.Data.outputDirectory))
            Copy-Item -LiteralPath $script:PreparedCandidate.Data.outputDirectory -Destination $destination -Recurse
            return $destination
        }
        function Invoke-UpdCandidateChecksumWrite {
            param([string]$Directory)
            $zipName = $script:CandidateConfig.packageName + '-' + $script:CandidateConfig.version + '.zip'
            $text = (Get-FileHash -LiteralPath (Join-Path $Directory $zipName) -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + $zipName + "`n" +
                (Get-FileHash -LiteralPath (Join-Path $Directory 'manifest.json') -Algorithm SHA256).Hash.ToLowerInvariant() + "  manifest.json`n"
            [IO.File]::WriteAllText((Join-Path $Directory 'SHA256SUMS'), $text, (New-Object Text.UTF8Encoding($false)))
        }
        function Write-UpdCandidateZipEntry {
            param([string]$Directory, [string]$EntryName, [byte[]]$Bytes, [switch]$Duplicate)
            $zipPath = Join-Path $Directory ($script:CandidateConfig.packageName + '-' + $script:CandidateConfig.version + '.zip')
            $archive = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Update)
            try {
                $old = $archive.GetEntry($EntryName)
                if ($old -and -not $Duplicate) { $old.Delete() }
                $entry = $archive.CreateEntry($EntryName)
                $stream = $entry.Open()
                try { $stream.Write($Bytes, 0, $Bytes.Length) } finally { $stream.Dispose() }
            }
            finally { $archive.Dispose() }
        }
        function Write-UpdCandidateManifest {
            param([string]$Directory, [object]$Manifest)
            $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes(($Manifest | ConvertTo-Json -Depth 8))
            [IO.File]::WriteAllBytes((Join-Path $Directory 'manifest.json'), $bytes)
            Write-UpdCandidateZipEntry -Directory $Directory -EntryName 'manifest.json' -Bytes $bytes
            Invoke-UpdCandidateChecksumWrite -Directory $Directory
        }
    }

    It 'accepts the actual workflow with least privilege, both engines and validated candidate outputs' {
        { Assert-UpdWorkflowPolicy $script:WorkflowText } | Should -Not -Throw
    }

    It 'refuses <Label> workflow changes before they can authorize candidate upload' -TestCases @(
        @{ Label = 'privileged pull request execution'; Old = '  pull_request:'; New = '  pull_request_target:' },
        @{ Label = 'token default elevation'; Old = 'permissions: {}'; New = "permissions:`n  contents: write" },
        @{ Label = 'job token elevation'; Old = '    contents: read'; New = '    contents: write' },
        @{ Label = 'mutable checkout action'; Old = 'actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1'; New = 'actions/checkout@main' },
        @{ Label = 'retained checkout credentials'; Old = 'persist-credentials: false'; New = 'persist-credentials: true' },
        @{ Label = 'moving checkout ref'; Old = 'ref: ${{ env.UPD_SOURCE_COMMIT }}'; New = 'ref: main' },
        @{ Label = 'missing complete test prerequisite'; Old = './scripts/Test.ps1 -Suite All'; New = './scripts/Test.ps1 -Suite Unit' },
        @{ Label = 'skipped analysis prerequisite'; Old = 'if: ${{ !cancelled() }}'; New = 'if: false' },
        @{ Label = 'skipped dependency result'; Old = "success() && needs.windows.result == 'success'"; New = 'always()' },
        @{ Label = 'wildcard artifact capture'; Old = '${{ steps.prepared.outputs.zipPath }}'; New = 'artifacts/**' },
        @{ Label = 'hidden artifact capture'; Old = 'include-hidden-files: false'; New = 'include-hidden-files: true' },
        @{ Label = 'artifact replacement'; Old = 'overwrite: false'; New = 'overwrite: true' },
        @{ Label = 'missing artifact tolerance'; Old = 'if-no-files-found: error'; New = 'if-no-files-found: warn' },
        @{ Label = 'readiness bypass'; Old = './scripts/Prepare-ReleaseCandidate.ps1'; New = './scripts/Build-Release.ps1' },
        @{ Label = 'explicit secret reference'; Old = 'name: Test foundations'; New = 'name: ${{ secrets.PUBLIC_TOKEN }}' }
    ) {
        param($Label, $Old, $New)
        $Label | Should -Not -BeNullOrEmpty
        $script:WorkflowText.Contains($Old) | Should -BeTrue -Because 'a mutation probe must actually change the checked workflow'
        $changed = $script:WorkflowText.Replace($Old, $New)
        { Assert-UpdWorkflowPolicy $changed } | Should -Throw
    }

    It 'fails closed for ambiguous or unsupported <Label> YAML policy syntax' -TestCases @(
        @{ Label = 'duplicate permissions key'; Extra = "`npermissions: {}`n" },
        @{ Label = 'anchor and alias'; Extra = "`nshared: &shared`n  contents: read`ncopy: *shared`n" },
        @{ Label = 'second document'; Extra = "`n---`npermissions: {}`n" }
    ) {
        param($Label, $Extra)
        $Label | Should -Not -BeNullOrEmpty
        $Extra | Should -Not -BeNullOrEmpty
        { ConvertFrom-UpdWorkflowPolicyYaml ($script:WorkflowText + $Extra) } | Should -Throw
    }

    It 'builds a genuine synthetic candidate and exports only three validated immutable file paths' {
        $script:PreparedCandidate.Code | Should -Be 0
        $ready = Test-UpdReleaseCandidate -Repository $script:CandidateFixture -SourceCommit $script:CandidateCommit -CandidateDirectory $script:PreparedCandidate.Data.outputDirectory
        $ready.sourceCommit | Should -BeExactly $script:CandidateCommit
        $ready.sourceTree | Should -BeExactly $script:CandidateTree
        $ready.version | Should -BeExactly $script:CandidateConfig.version
        $lines = @([IO.File]::ReadAllLines($script:PreparedOutput))
        $lines.Count | Should -Be 4
        $lines[0] | Should -BeExactly 'owned-existing-output'
        $lines[1] | Should -BeExactly ('zipPath=' + $ready.zipPath)
        $lines[2] | Should -BeExactly ('manifestPath=' + $ready.manifestPath)
        $lines[3] | Should -BeExactly ('checksumPath=' + $ready.checksumPath)
        @(Get-ChildItem -LiteralPath $ready.outputDirectory -Force).Count | Should -Be 3
    }

    It 'refuses <Label> output poisoning while preserving the candidate and unknown file' -TestCases @(
        @{ Label = 'extra private log'; Name = 'private-feed.log' },
        @{ Label = 'hidden private state'; Name = '.podcast-history.json' },
        @{ Label = 'nested archive folder'; Name = 'media'; Directory = $true }
    ) {
        param($Label, $Name, $Directory)
        $Label | Should -Not -BeNullOrEmpty
        $candidate = Invoke-UpdCandidateCopy
        $poison = Join-Path $candidate $Name
        if ($Directory) { [void][IO.Directory]::CreateDirectory($poison); $poison = Join-Path $poison 'unknown.mp3' }
        [IO.File]::WriteAllText($poison, 'preserve unknown synthetic private material')
        if ($Name.StartsWith('.')) { [IO.File]::SetAttributes($poison, [IO.FileAttributes]::Hidden) }
        { Test-UpdReleaseCandidate -Repository $script:CandidateFixture -SourceCommit $script:CandidateCommit -CandidateDirectory $candidate } | Should -Throw
        [IO.File]::ReadAllText($poison) | Should -BeExactly 'preserve unknown synthetic private material'
    }

    It 'refuses hash-consistent <Label> manifest substitutions' -TestCases @(
        @{ Label = 'different source commit'; Field = 'sourceCommit'; Value = '0000000000000000000000000000000000000000' },
        @{ Label = 'different source tree'; Field = 'sourceTree'; Value = '0000000000000000000000000000000000000000' },
        @{ Label = 'stable release claim'; Field = 'releaseStatus'; Value = 'RELEASED' },
        @{ Label = 'different version'; Field = 'version'; Value = '0.1.0-rc.2' },
        @{ Label = 'foreign source URL'; Field = 'sourceCommitUrl'; Value = 'https://example.invalid/commit/claimed-source' }
    ) {
        param($Label, $Field, $Value)
        $Label | Should -Not -BeNullOrEmpty
        $candidate = Invoke-UpdCandidateCopy
        $manifest = [IO.File]::ReadAllText((Join-Path $candidate 'manifest.json')) | ConvertFrom-Json
        $manifest.$Field = $Value
        Write-UpdCandidateManifest -Directory $candidate -Manifest $manifest
        { Test-UpdReleaseCandidate -Repository $script:CandidateFixture -SourceCommit $script:CandidateCommit -CandidateDirectory $candidate } | Should -Throw
    }

    It 'refuses <Label> ZIP changes even when sidecar checksums are recomputed' -TestCases @(
        @{ Label = 'extra traversal entry'; Entry = '../private.mp3' },
        @{ Label = 'duplicate licensed entry'; Entry = 'LICENSE'; Duplicate = $true },
        @{ Label = 'payload byte substitution'; Entry = 'src/Batch.ps1' },
        @{ Label = 'different internal manifest'; Entry = 'manifest.json' }
    ) {
        param($Label, $Entry, $Duplicate)
        $Label | Should -Not -BeNullOrEmpty
        $candidate = Invoke-UpdCandidateCopy
        Write-UpdCandidateZipEntry -Directory $candidate -EntryName $Entry -Bytes ([byte[]](1, 2, 3)) -Duplicate:$Duplicate
        Invoke-UpdCandidateChecksumWrite -Directory $candidate
        { Test-UpdReleaseCandidate -Repository $script:CandidateFixture -SourceCommit $script:CandidateCommit -CandidateDirectory $candidate } | Should -Throw
    }

    It 'refuses substituted payload bytes with a self-consistent ZIP, both manifests and checksum list' {
        $candidate = Invoke-UpdCandidateCopy
        $bytes = [byte[]](1, 2, 3)
        Write-UpdCandidateZipEntry -Directory $candidate -EntryName 'src/Batch.ps1' -Bytes $bytes
        $manifest = [IO.File]::ReadAllText((Join-Path $candidate 'manifest.json')) | ConvertFrom-Json
        $record = @($manifest.files | Where-Object { $_.path -ceq 'src/Batch.ps1' })[0]
        $record.size = $bytes.Length
        $stream = New-Object IO.MemoryStream (,$bytes)
        try { $record.sha256 = Get-UpdCandidateStreamHash -Stream $stream } finally { $stream.Dispose() }
        Write-UpdCandidateManifest -Directory $candidate -Manifest $manifest
        { Test-UpdReleaseCandidate -Repository $script:CandidateFixture -SourceCommit $script:CandidateCommit -CandidateDirectory $candidate } | Should -Throw
    }

    It 'refuses a candidate directory junction without changing its target' {
        $candidate = Invoke-UpdCandidateCopy
        $link = Join-Path (Split-Path $candidate -Parent) 'junction'
        $null = New-Item -ItemType Junction -Path $link -Target $candidate
        $hash = (Get-FileHash -LiteralPath (Join-Path $candidate 'manifest.json')).Hash
        try {
            { Test-UpdReleaseCandidate -Repository $script:CandidateFixture -SourceCommit $script:CandidateCommit -CandidateDirectory $link } | Should -Throw
            (Get-FileHash -LiteralPath (Join-Path $candidate 'manifest.json')).Hash | Should -BeExactly $hash
        }
        finally { [IO.Directory]::Delete($link) }
    }

    It 'refuses <Label> source selections without building or exporting Actions outputs' -TestCases @(
        @{ Label = 'revision expression'; Commit = 'HEAD' },
        @{ Label = 'missing full commit'; Commit = '0000000000000000000000000000000000000000' }
    ) {
        param($Label, $Commit)
        $Label | Should -Not -BeNullOrEmpty
        $directory = Join-Path $script:CandidateFixture ('artifacts/refused-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $output = Join-Path $TestDrive ('refused-' + [guid]::NewGuid().ToString('N') + '.txt')
        [IO.File]::WriteAllText($output, 'preserve existing output')
        $result = Invoke-UpdCandidateFixturePrepare -Commit $Commit -OutputDirectory $directory -GitHubOutput $output
        $result.Code | Should -Be 1 -Because $result.Text
        Test-Path -LiteralPath $directory | Should -BeFalse
        [IO.File]::ReadAllText($output) | Should -BeExactly 'preserve existing output'
    }

    It 'rejects a line-injected Actions path before appending any outputs' {
        $candidate = [pscustomobject]@{ zipPath = "safe.zip`nmaliciousOutput=value"; manifestPath = 'manifest.json'; checksumPath = 'SHA256SUMS' }
        $output = Join-Path $TestDrive 'injected-output.txt'
        [IO.File]::WriteAllText($output, 'preserve')
        { Write-UpdCandidateGitHubOutput -Candidate $candidate -Path $output } | Should -Throw
        [IO.File]::ReadAllText($output) | Should -BeExactly 'preserve'
    }

    It 'refuses dirty tracked source before building or appending candidate outputs' {
        $sourcePath = Join-Path $script:CandidateFixture 'README.md'
        $original = [IO.File]::ReadAllBytes($sourcePath)
        $directory = Join-Path $script:CandidateFixture 'artifacts/dirty-refused'
        $output = Join-Path $TestDrive 'dirty-refused-output.txt'
        [IO.File]::WriteAllText($output, 'preserve existing output')
        try {
            [IO.File]::WriteAllText($sourcePath, 'uncommitted synthetic replacement')
            $result = Invoke-UpdCandidateFixturePrepare -Commit $script:CandidateCommit -OutputDirectory $directory -GitHubOutput $output
            $result.Code | Should -Be 1 -Because $result.Text
            Test-Path -LiteralPath $directory | Should -BeFalse
            [IO.File]::ReadAllText($output) | Should -BeExactly 'preserve existing output'
        }
        finally { [IO.File]::WriteAllBytes($sourcePath, $original) }
    }

    It 'refuses Actions output overlap with candidate <Field> without altering candidate bytes' -TestCases @(
        @{ Field = 'zipPath' }, @{ Field = 'manifestPath' }, @{ Field = 'checksumPath' }
    ) {
        param($Field)
        $candidate = Invoke-UpdCandidateCopy
        $ready = Test-UpdReleaseCandidate -Repository $script:CandidateFixture -SourceCommit $script:CandidateCommit -CandidateDirectory $candidate
        $path = $ready.$Field
        $original = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        { Write-UpdCandidateGitHubOutput -Candidate $ready -Path $path } | Should -Throw
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -BeExactly $original
    }

    It 'refuses hidden <Label> before an upload can silently omit candidate files' -TestCases @(
        @{ Label = 'allowlisted sidecar'; Child = 'manifest.json' },
        @{ Label = 'directory component'; Child = '' }
    ) {
        param($Label, $Child)
        $Label | Should -Not -BeNullOrEmpty
        $candidate = Invoke-UpdCandidateCopy
        if (-not $Child) {
            $hiddenParent = Join-Path (Split-Path $candidate -Parent) '.hidden'
            [void][IO.Directory]::CreateDirectory($hiddenParent)
            $hiddenCandidate = Join-Path $hiddenParent ([IO.Path]::GetFileName($candidate))
            Copy-Item -LiteralPath $candidate -Destination $hiddenCandidate -Recurse
            $candidate = $hiddenCandidate
        }
        $path = if ($Child) { Join-Path $candidate $Child } else { $candidate }
        $original = [IO.File]::GetAttributes($path)
        try {
            [IO.File]::SetAttributes($path, ($original -bor [IO.FileAttributes]::Hidden))
            { Test-UpdReleaseCandidate -Repository $script:CandidateFixture -SourceCommit $script:CandidateCommit -CandidateDirectory $candidate } | Should -Throw
        }
        finally { [IO.File]::SetAttributes($path, $original) }
    }
}
