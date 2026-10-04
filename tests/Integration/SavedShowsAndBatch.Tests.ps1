BeforeAll {
    $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $repositoryRoot 'tests/support/IntegrationHarness.ps1')
    $python = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) { throw 'Python is required for owned saved-show and batch fixtures.' }
    $script:savedPython = $python.Source
    $script:savedAudioHash = (Get-FileHash -LiteralPath (Join-Path $repositoryRoot 'tools/codex-handoff/fixtures/silence.mp3') -Algorithm SHA256).Hash
    if ($PSVersionTable.PSEdition -eq 'Core') { Add-Type -AssemblyName System.Security.Cryptography.ProtectedData }
    else { Add-Type -AssemblyName System.Security }

    function Invoke-UpdSavedWorker {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][hashtable]$Options, [long]$CancelAfterBytes = 0)
        $id = [guid]::NewGuid().ToString('N')
        $workerConfig = Join-Path $Context.Root ($id + '-saved-worker.json')
        $payload = @{ Root = $Context.Root; Token = $Context.Token; ProductScript = Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1';
            SavedConfigPath = Join-Path $Context.Root 'config/shows.json'; ResultPath = Join-Path $Context.Root ($id + '-saved-result.json');
            Options = $Options; CancelAfterBytes = $CancelAfterBytes }
        [IO.File]::WriteAllText($workerConfig, ($payload | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
        $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $owned = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engine) -ArgumentList @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
            (Join-Path $Context.RepositoryRoot 'tests/support/Invoke-SavedShowsWorker.ps1'), '-ConfigPath', $workerConfig)
        if (-not $owned.Process.WaitForExit(60000)) { throw 'Owned saved-show worker completion timed out.' }
        if (-not [IO.File]::Exists($payload.ResultPath)) { throw ('Owned saved-show worker produced no result: ' + $owned.ErrorOutput.Result) }
        [pscustomobject]@{ ExitCode = $owned.Process.ExitCode; Report = [IO.File]::ReadAllText($payload.ResultPath) | ConvertFrom-Json;
            Stdout = $owned.Output.Result; Stderr = $owned.ErrorOutput.Result }
    }

    function Assert-UpdSavedResult {
        param([Parameter(Mandatory)]$Run, [int]$ExitCode = 0, [string]$Type = 'Podcast.ConfigResult')
        $Run.Report.Error | Should -BeNullOrEmpty -Because ($Run.Stdout + $Run.Stderr)
        $Run.Report.HostSurvived | Should -BeTrue
        $Run.ExitCode | Should -Be $ExitCode -Because ($Run.Stdout + $Run.Stderr)
        $Run.Report.Result.ExitCode | Should -Be $ExitCode
        $Run.Report.Result.Type | Should -Be $Type
        $Run.Report.Result.SchemaVersion | Should -Be 1
        # Existing preview Source intentionally contains a host and opaque URL token.
        $Run.Report.Result | ConvertTo-Json -Depth 12 | Should -Not -Match 'https?://|[A-Za-z]:\\|feed_protected|output_path|LegacyResult|FAKE_SAVED_SECRET'
    }

    function Invoke-UpdSavedEntry {
        param([Parameter(Mandatory)]$Context, [ValidateSet('SavePreview', 'BatchPreview')][string]$Action)
        $arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
            (Join-Path $Context.RepositoryRoot 'UniversalPodcastDownloader.ps1'),
            '-ConfigPath', (Join-Path $Context.Root 'config/shows.json'), '-NonInteractive', '-WhatIf', '-PassThru')
        if ($Action -eq 'SavePreview') {
            $arguments += @('-SaveShow', 'alpha', '-FeedUrl', ($Context.BaseUrl + '/feeds/single.xml'),
                '-OutputPath', (Join-Path $Context.Root 'alpha'), '-Mode', 'All')
        }
        else { $arguments += '-Batch' }
        $engine = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $owned = Start-UpdOwnedProcess -Context $Context -FilePath (Join-Path $PSHOME $engine) -ArgumentList $arguments
        if (-not $owned.Process.WaitForExit(30000)) { throw 'Owned saved-show script entry completion timed out.' }
        [pscustomobject]@{ ExitCode = $owned.Process.ExitCode; Stdout = $owned.Output.Result; Stderr = $owned.ErrorOutput.Result }
    }

    function Add-UpdSavedFixture {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Stores only an explicit loopback fixture in its marked temporary configuration through the actual product dispatcher.')]
        [CmdletBinding()]
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name, [string]$FeedPath = '/feeds/single.xml', [string]$OutputName)
        if (-not $OutputName) { $OutputName = $Name }
        $run = Invoke-UpdSavedWorker -Context $Context -Options @{ SaveShow = $Name; FeedUrl = $Context.BaseUrl + $FeedPath;
            OutputPath = Join-Path $Context.Root $OutputName; Mode = 'All' }
        Assert-UpdSavedResult -Run $run
        $run.Report.ProtectCalls | Should -Be 1
        $run.Report.ConfigLockCalls | Should -Be 1
        $run.Report.Result.Changed | Should -BeTrue
    }

    function Assert-UpdSavedAcl {
        param([Parameter(Mandatory)][string]$Path, [switch]$Directory)
        $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User
        $acl = Get-Acl -LiteralPath $Path
        $acl.GetOwner([Security.Principal.SecurityIdentifier]).Value | Should -Be $sid.Value
        $acl.AreAccessRulesProtected | Should -BeTrue
        $rules = @($acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))
        $rules.Count | Should -Be 1
        $rules[0].IdentityReference.Value | Should -Be $sid.Value
        $rules[0].AccessControlType | Should -Be ([Security.AccessControl.AccessControlType]::Allow)
        $rules[0].FileSystemRights | Should -Be ([Security.AccessControl.FileSystemRights]::FullControl)
        $rules[0].IsInherited | Should -BeFalse
        $rules[0].PropagationFlags | Should -Be ([Security.AccessControl.PropagationFlags]::None)
        $expected = if ($Directory) { [Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit' } else { [Security.AccessControl.InheritanceFlags]::None }
        $rules[0].InheritanceFlags | Should -Be $expected
    }

    function Get-UpdSavedTree {
        param([Parameter(Mandatory)]$Context)
        $rows = foreach ($name in @('config', 'alpha', 'beta', 'local')) {
            $path = Join-Path $Context.Root $name
            if (-not (Test-Path -LiteralPath $path)) { [pscustomobject]@{ Path = $name; Missing = $true }; continue }
            foreach ($entry in @((Get-Item -LiteralPath $path)) + @(Get-ChildItem -LiteralPath $path -Recurse -Force)) {
                [pscustomobject]@{ Path = $entry.FullName.Substring($Context.Root.Length); Directory = $entry.PSIsContainer;
                    Modified = $entry.LastWriteTimeUtc.Ticks; Hash = $(if (-not $entry.PSIsContainer) { (Get-FileHash -LiteralPath $entry.FullName -Algorithm SHA256).Hash } else { $null }) }
            }
        }
        ConvertTo-Json -InputObject @($rows | Sort-Object Path) -Depth 4 -Compress
    }

    function Assert-UpdSavedAclRestore {
        param([Parameter(Mandatory)]$OriginalAcl, [Parameter(Mandatory)]$RestoredAcl)
        # Same complete descriptor proof as the owned run-safety ACL fixtures.
        $originalDescriptor = [Security.AccessControl.RawSecurityDescriptor]::new($OriginalAcl.GetSecurityDescriptorBinaryForm(), 0)
        $restoredDescriptor = [Security.AccessControl.RawSecurityDescriptor]::new($RestoredAcl.GetSecurityDescriptorBinaryForm(), 0)
        $restoredDescriptor.Owner.Value | Should -BeExactly $originalDescriptor.Owner.Value
        $restoredDescriptor.Group.Value | Should -BeExactly $originalDescriptor.Group.Value
        $originalFlags = [int]$originalDescriptor.ControlFlags
        $restoredFlags = [int]$restoredDescriptor.ControlFlags
        $autoInherited = [int][Security.AccessControl.ControlFlags]::DiscretionaryAclAutoInherited
        ($restoredFlags -eq $originalFlags -or $restoredFlags -eq ($originalFlags -bor $autoInherited)) |
            Should -BeTrue -Because 'Windows may only add the discretionary ACL automatic-inheritance bookkeeping flag'
        $restoredDescriptor.ResourceManagerControl | Should -Be $originalDescriptor.ResourceManagerControl
        foreach ($aclName in @('DiscretionaryAcl', 'SystemAcl')) {
            $originalEntries = $originalDescriptor.$aclName
            $restoredEntries = $restoredDescriptor.$aclName
            ($null -ne $restoredEntries) | Should -Be ($null -ne $originalEntries)
            if ($null -eq $originalEntries) { continue }
            $restoredEntries.Revision | Should -Be $originalEntries.Revision
            $restoredEntries.Count | Should -Be $originalEntries.Count
            $originalBytes = New-Object byte[] $originalEntries.BinaryLength
            $restoredBytes = New-Object byte[] $restoredEntries.BinaryLength
            $originalEntries.GetBinaryForm($originalBytes, 0)
            $restoredEntries.GetBinaryForm($restoredBytes, 0)
            [Convert]::ToBase64String($restoredBytes) | Should -BeExactly ([Convert]::ToBase64String($originalBytes))
        }
    }

    function Get-UpdSavedArchive {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name)
        $paths = @(Get-ChildItem -LiteralPath (Join-Path $Context.Root $Name) -Recurse -Force -File -Filter 'state.json')
        $paths.Count | Should -Be 1
        [pscustomobject]@{ Root = Split-Path $paths[0].DirectoryName -Parent; State = [IO.File]::ReadAllText($paths[0].FullName) | ConvertFrom-Json }
    }

    function Assert-UpdSavedOriginal {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name)
        $archive = Get-UpdSavedArchive -Context $Context -Name $Name
        $archive.State.episodes.Count | Should -Be 1
        $archive.State.episodes[0].status | Should -Be 'transfer_verified'
        (Get-FileHash -LiteralPath (Join-Path $archive.Root $archive.State.episodes[0].relative_path) -Algorithm SHA256).Hash | Should -Be $script:savedAudioHash
    }
}

Describe 'A046/A047 saved-show configuration and sequential batches through actual owned processes' {
    BeforeEach {
        $context = New-UpdIntegrationContext -RepositoryRoot $repositoryRoot
        Start-UpdFixtureServer -Context $context -PythonPath $script:savedPython
        $script:savedPath = Join-Path $context.Root 'config/shows.json'
    }
    AfterEach { if ($context) { Remove-UpdIntegrationContext -Context $context } }

    It 'accepts SaveShow WhatIf at the actual script parameter and exit boundary without creating configuration' {
        $before = Get-UpdSavedTree -Context $context
        $entry = Invoke-UpdSavedEntry -Context $context -Action SavePreview
        $entry.ExitCode | Should -Be 0 -Because ($entry.Stdout + $entry.Stderr)
        $entry.Stderr | Should -BeNullOrEmpty
        $entry.Stdout | Should -Match 'Podcast\.ConfigResult'
        $entry.Stdout | Should -Not -Match 'https?://|FAKE_SAVED_SECRET|\[ERROR\]'
        (Get-UpdSavedTree -Context $context) | Should -Be $before
        Test-Path -LiteralPath $script:savedPath | Should -BeFalse
    }

    It 'accepts an empty Batch WhatIf at the actual script parameter and exit boundary without creating configuration' {
        $before = Get-UpdSavedTree -Context $context
        $entry = Invoke-UpdSavedEntry -Context $context -Action BatchPreview
        $entry.ExitCode | Should -Be 0 -Because ($entry.Stdout + $entry.Stderr)
        $entry.Stderr | Should -BeNullOrEmpty
        $entry.Stdout | Should -Match 'Podcast\.BatchResult'
        $entry.Stdout | Should -Not -Match 'https?://|FAKE_SAVED_SECRET|\[ERROR\]'
        (Get-UpdSavedTree -Context $context) | Should -Be $before
        Test-Path -LiteralPath $script:savedPath | Should -BeFalse
    }

    It 'protects the actual feed with CurrentUser DPAPI and ACLs, exports only safe metadata, and verifies a named repeat' {
        $secretUrl = $context.BaseUrl + '/feeds/single.xml?owned-secret=FAKE_SAVED_SECRET'
        $saved = Invoke-UpdSavedWorker -Context $context -Options @{ SaveShow = 'alpha'; FeedUrl = $secretUrl; OutputPath = Join-Path $context.Root 'alpha'; Mode = 'All' }
        Assert-UpdSavedResult -Run $saved
        Assert-UpdSavedAcl -Path (Split-Path $script:savedPath -Parent) -Directory
        Assert-UpdSavedAcl -Path $script:savedPath
        $text = [IO.File]::ReadAllText($script:savedPath)
        $text | Should -Not -Match 'FAKE_SAVED_SECRET|https?://|127\.0\.0\.1'
        $stored = $text | ConvertFrom-Json
        $entropy = [Text.Encoding]::UTF8.GetBytes('UniversalPodcastDownloader.SavedShows.Schema1')
        $plain = [Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($stored.shows[0].feed_protected), $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
        [Text.Encoding]::UTF8.GetString($plain) | Should -Be $secretUrl
        $configHash = (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash
        $listed = Invoke-UpdSavedWorker -Context $context -Options @{ ListShows = $true }
        Assert-UpdSavedResult -Run $listed
        @($listed.Report.Result.Shows).Count | Should -Be 1
        $listed.Report.Result.Shows[0].Name | Should -Be 'alpha'
        $listed.Report.Result.Shows[0].FeedConfigured | Should -BeTrue
        $exportPath = Join-Path (Split-Path $script:savedPath -Parent) 'safe-export.json'
        $export = Invoke-UpdSavedWorker -Context $context -Options @{ ExportShows = $exportPath }
        Assert-UpdSavedResult -Run $export
        Assert-UpdSavedAcl -Path $exportPath
        $exportText = [IO.File]::ReadAllText($exportPath)
        $exportText | Should -Not -Match 'FAKE_SAVED_SECRET|https?://|feed_protected|output_path|[A-Za-z]:\\'
        $exportText | Should -Not -Match ([regex]::Escape($stored.shows[0].feed_protected))
        ($exportText | ConvertFrom-Json).shows[0].Name | Should -Be 'alpha'
        $exportHash = (Get-FileHash -LiteralPath $exportPath -Algorithm SHA256).Hash
        $existingExport = Invoke-UpdSavedWorker -Context $context -Options @{ ExportShows = $exportPath }
        Assert-UpdSavedResult -Run $existingExport -ExitCode 1
        (Get-FileHash -LiteralPath $exportPath -Algorithm SHA256).Hash | Should -Be $exportHash
        $first = Invoke-UpdSavedWorker -Context $context -Options @{ ShowName = @('alpha') }
        Assert-UpdSavedResult -Run $first -Type 'Podcast.RunResult'
        $first.Report.Result.Downloaded | Should -Be 1
        Assert-UpdSavedOriginal -Context $context -Name 'alpha'
        $repeat = Invoke-UpdSavedWorker -Context $context -Options @{ ShowName = @('alpha') }
        Assert-UpdSavedResult -Run $repeat -Type 'Podcast.RunResult'
        $repeat.Report.Result.VerifiedSkipped | Should -Be 1
        @($repeat.Report.RequestTrace | Where-Object Kind -eq 'media').Count | Should -Be 0
        (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash | Should -Be $configHash
    }

    It 'preserves complete owned trees for management, named and batch WhatIf without protection, locks, media or power' {
        $missing = Invoke-UpdSavedWorker -Context $context -Options @{ SaveShow = 'alpha'; FeedUrl = $context.BaseUrl + '/feeds/single.xml'; OutputPath = Join-Path $context.Root 'alpha'; WhatIf = $true }
        Assert-UpdSavedResult -Run $missing
        Test-Path -LiteralPath (Split-Path $script:savedPath -Parent) | Should -BeFalse
        $missing.Report.ProtectCalls | Should -Be 0
        $missing.Report.ConfigLockCalls | Should -Be 0
        Add-UpdSavedFixture -Context $context -Name 'alpha'
        Add-UpdSavedFixture -Context $context -Name 'beta'
        $before = Get-UpdSavedTree -Context $context
        $operations = @(
            @{ SaveShow = 'extra'; FeedUrl = $context.BaseUrl + '/feeds/single.xml'; OutputPath = Join-Path $context.Root 'beta'; WhatIf = $true },
            @{ RemoveShow = 'alpha'; WhatIf = $true },
            @{ ExportShows = Join-Path (Split-Path $script:savedPath -Parent) 'preview-export.json'; WhatIf = $true },
            @{ ShowName = @('alpha'); WhatIf = $true; KeepAwake = $true },
            @{ Batch = $true; WhatIf = $true; KeepAwake = $true })
        foreach ($options in $operations) {
            $run = Invoke-UpdSavedWorker -Context $context -Options $options
            $type = if ($options.Batch) { 'Podcast.BatchResult' } elseif ($options.ShowName) { 'Podcast.RunResult' } else { 'Podcast.ConfigResult' }
            Assert-UpdSavedResult -Run $run -Type $type
            $run.Report.Result.Preview | Should -BeTrue
            $run.Report.ProtectCalls | Should -Be 0
            $run.Report.ConfigLockCalls | Should -Be 0
            $run.Report.PowerRequested | Should -Be 0
            @($run.Report.RequestTrace | Where-Object Kind -eq 'media').Count | Should -Be 0
            (Get-UpdSavedTree -Context $context) | Should -Be $before
        }
    }

    It 'preserves strict invalid configuration without running shows: <Kind>' -ForEach @(
        @{ Kind = 'malformed'; Document = '{"schema_version":1,"shows":[' },
        @{ Kind = 'unknown schema'; Document = '{"schema_version":99,"shows":[]}' },
        @{ Kind = 'extra executable-looking data'; Document = '{"schema_version":1,"shows":[],"command":"$(New-Item injected-owned-file)"}' }) {
        Add-UpdSavedFixture -Context $context -Name 'alpha'
        [IO.File]::WriteAllText($script:savedPath, $Document, [Text.UTF8Encoding]::new($false))
        $hash = (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash
        $listed = Invoke-UpdSavedWorker -Context $context -Options @{ ListShows = $true }
        Assert-UpdSavedResult -Run $listed -ExitCode 1
        $listed.Report.Result.Message | Should -Be 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
        $batch = Invoke-UpdSavedWorker -Context $context -Options @{ Batch = $true }
        Assert-UpdSavedResult -Run $batch -ExitCode 1 -Type 'Podcast.BatchResult'
        $batch.Report.Result.ProcessedShows | Should -Be 0
        @($batch.Report.RequestTrace).Count | Should -Be 0
        $batch.Report.ConfigLockCalls | Should -Be 0
        Test-Path -LiteralPath (Join-Path $context.Root 'injected-owned-file') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $context.Root 'alpha') | Should -BeFalse
        (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash | Should -Be $hash
    }

    It 'isolates an unavailable encrypted feed while a later saved show completes unchanged original bytes' {
        Add-UpdSavedFixture -Context $context -Name 'unavailable'
        Add-UpdSavedFixture -Context $context -Name 'healthy'
        $stored = [IO.File]::ReadAllText($script:savedPath) | ConvertFrom-Json
        $stored.shows[0].feed_protected = 'AQIDBA=='
        [IO.File]::WriteAllText($script:savedPath, ($stored | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
        $hash = (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash
        $listed = Invoke-UpdSavedWorker -Context $context -Options @{ ListShows = $true }
        Assert-UpdSavedResult -Run $listed
        @($listed.Report.Result.Shows).Count | Should -Be 2
        $run = Invoke-UpdSavedWorker -Context $context -Options @{ Batch = $true }
        Assert-UpdSavedResult -Run $run -ExitCode 2 -Type 'Podcast.BatchResult'
        $run.Report.Result.ProcessedShows | Should -Be 2
        $run.Report.Result.FatalShows | Should -Be 1
        $run.Report.Result.SuccessfulShows | Should -Be 1
        $run.Report.Result.Shows[0].Result.Message | Should -Be 'Saved-show credentials are unavailable for the current Windows user.'
        @($run.Report.RequestTrace | Where-Object Show -eq 'unavailable').Count | Should -Be 0
        Assert-UpdSavedOriginal -Context $context -Name 'healthy'
        (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash | Should -Be $hash
    }

    It 'rejects an unsafe owned configuration directory ACL and preserves the bytes and restored original permissions' {
        Add-UpdSavedFixture -Context $context -Name 'alpha'
        $directory = Split-Path $script:savedPath -Parent
        $originalAcl = Get-Acl -LiteralPath $directory
        $info = [IO.DirectoryInfo]::new($directory)
        $hash = (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash
        try {
            $unsafe = [Security.AccessControl.DirectorySecurity]::new()
            $unsafe.SetSecurityDescriptorBinaryForm($originalAcl.GetSecurityDescriptorBinaryForm(), [Security.AccessControl.AccessControlSections]::Access)
            $everyone = [Security.Principal.SecurityIdentifier]::new('S-1-1-0')
            $unsafe.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($everyone, [Security.AccessControl.FileSystemRights]::ReadAndExecute, [Security.AccessControl.AccessControlType]::Allow))
            if ('System.IO.FileSystemAclExtensions' -as [type]) { [IO.FileSystemAclExtensions]::SetAccessControl($info, $unsafe) }
            else { $info.SetAccessControl($unsafe) }
            $run = Invoke-UpdSavedWorker -Context $context -Options @{ ListShows = $true }
            Assert-UpdSavedResult -Run $run -ExitCode 1
            $run.Report.Result.Message | Should -Be 'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.'
            @($run.Report.RequestTrace).Count | Should -Be 0
            $run.Report.ConfigLockCalls | Should -Be 0
            (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash | Should -Be $hash
        }
        finally {
            # Only the DACL was changed. Marking only Access for restoration avoids
            # requesting SeSecurityPrivilege to write the untouched system ACL.
            $restore = [Security.AccessControl.DirectorySecurity]::new()
            $restore.SetSecurityDescriptorBinaryForm($originalAcl.GetSecurityDescriptorBinaryForm(), [Security.AccessControl.AccessControlSections]::Access)
            if ('System.IO.FileSystemAclExtensions' -as [type]) { [IO.FileSystemAclExtensions]::SetAccessControl($info, $restore) }
            else { $info.SetAccessControl($restore) }
            Assert-UpdSavedAclRestore -OriginalAcl $originalAcl -RestoredAcl (Get-Acl -LiteralPath $directory)
            Assert-UpdSavedAcl -Path $directory -Directory
        }
    }

    It 'treats a valid saved output containing PowerShell syntax as a literal path rather than executing it' {
        # Keep the actual Windows root within the product's identifier budget.
        $literalName = '$(mkdir x)'
        Add-UpdSavedFixture -Context $context -Name 'literal' -OutputName $literalName
        $configHash = (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash
        $run = Invoke-UpdSavedWorker -Context $context -Options @{ ShowName = @('literal') }
        Assert-UpdSavedResult -Run $run -Type 'Podcast.RunResult'
        Assert-UpdSavedOriginal -Context $context -Name $literalName
        Test-Path -LiteralPath (Join-Path $context.Root $literalName) | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $context.Root 'x') | Should -BeFalse
        (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash | Should -Be $configHash
    }

    It 'runs real requests sequentially and continues after fatal feed and invalid-body failures with honest totals' {
        Add-UpdSavedFixture -Context $context -Name 'first'
        Add-UpdSavedFixture -Context $context -Name 'badfeed' -FeedPath '/feeds/malformed.xml'
        Add-UpdSavedFixture -Context $context -Name 'badbody' -FeedPath '/feeds/format-html.xml'
        Add-UpdSavedFixture -Context $context -Name 'last'
        $hash = (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash
        $run = Invoke-UpdSavedWorker -Context $context -Options @{ Batch = $true; MaxAttempts = 1 }
        Assert-UpdSavedResult -Run $run -ExitCode 2 -Type 'Podcast.BatchResult'
        $result = $run.Report.Result
        $result.SelectedShows | Should -Be 4
        $result.ProcessedShows | Should -Be 4
        $result.UnstartedShowCount | Should -Be 0
        $result.SuccessfulShows | Should -Be 2
        $result.FatalShows | Should -Be 1
        $result.IncompleteShows | Should -Be 1
        $result.Downloaded | Should -Be 2
        $result.Failed | Should -Be 1
        ($result.Shows.Name -join ',') | Should -Be 'first,badfeed,badbody,last'
        $run.Report.PeakActiveRuns | Should -Be 1
        (($run.Report.RunTrace | ForEach-Object { $_.Stage + ':' + $_.Name }) -join ',') | Should -Be 'start:first,end:first,start:badfeed,end:badfeed,start:badbody,end:badbody,start:last,end:last'
        (($run.Report.RequestTrace | Where-Object Kind -eq 'metadata').Show -join ',') | Should -Be 'first,badfeed,badbody,last'
        foreach ($start in @($run.Report.RunTrace | Where-Object Stage -eq 'start')) { $start.NonInteractive | Should -BeTrue; $start.Confirm | Should -BeFalse }
        Assert-UpdSavedOriginal -Context $context -Name 'first'
        Assert-UpdSavedOriginal -Context $context -Name 'last'
        @(Get-ChildItem -LiteralPath (Join-Path $context.Root 'badbody') -Recurse -File -Filter '*.mp3').Count | Should -Be 0
        (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash | Should -Be $hash
    }

    It 'honors subset request order, verifies repeats and preserves an existing changed media file without overwrite' {
        foreach ($name in @('alpha', 'beta', 'gamma')) { Add-UpdSavedFixture -Context $context -Name $name }
        $configHash = (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash
        $first = Invoke-UpdSavedWorker -Context $context -Options @{ Batch = $true; ShowName = @('gamma', 'alpha') }
        Assert-UpdSavedResult -Run $first -Type 'Podcast.BatchResult'
        ($first.Report.Result.Shows.Name -join ',') | Should -Be 'gamma,alpha'
        (($first.Report.RequestTrace | Where-Object Kind -eq 'metadata').Show -join ',') | Should -Be 'gamma,alpha'
        Test-Path -LiteralPath (Join-Path $context.Root 'beta') | Should -BeFalse
        Assert-UpdSavedOriginal -Context $context -Name 'alpha'
        Assert-UpdSavedOriginal -Context $context -Name 'gamma'
        $repeat = Invoke-UpdSavedWorker -Context $context -Options @{ Batch = $true; ShowName = @('alpha', 'gamma') }
        Assert-UpdSavedResult -Run $repeat -Type 'Podcast.BatchResult'
        $repeat.Report.Result.VerifiedSkipped | Should -Be 2
        @($repeat.Report.RequestTrace | Where-Object Kind -eq 'media').Count | Should -Be 0
        $archive = Get-UpdSavedArchive -Context $context -Name 'alpha'
        $media = Join-Path $archive.Root $archive.State.episodes[0].relative_path
        $stream = [IO.File]::Open($media, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try { $stream.WriteByte(42) } finally { $stream.Dispose() }
        $changedHash = (Get-FileHash -LiteralPath $media -Algorithm SHA256).Hash
        $conflict = Invoke-UpdSavedWorker -Context $context -Options @{ Batch = $true; ShowName = @('alpha', 'gamma') }
        Assert-UpdSavedResult -Run $conflict -ExitCode 2 -Type 'Podcast.BatchResult'
        $conflict.Report.Result.Conflicts | Should -Be 1
        $conflict.Report.Result.VerifiedSkipped | Should -Be 1
        @($conflict.Report.RequestTrace | Where-Object Kind -eq 'media').Count | Should -Be 0
        (Get-FileHash -LiteralPath $media -Algorithm SHA256).Hash | Should -Be $changedHash
        Assert-UpdSavedOriginal -Context $context -Name 'gamma'
        (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash | Should -Be $configHash
    }

    It 'stops at a real byte-callback cancellation with 130 and leaves the next show unstarted and owned recovery handles closed' {
        Add-UpdSavedFixture -Context $context -Name 'cancel' -FeedPath '/resume/feed/valid'
        Add-UpdSavedFixture -Context $context -Name 'later'
        $hash = (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash
        $run = Invoke-UpdSavedWorker -Context $context -Options @{ Batch = $true; MaxAttempts = 1 } -CancelAfterBytes 8192
        Assert-UpdSavedResult -Run $run -ExitCode 130 -Type 'Podcast.BatchResult'
        $run.Report.CancelBytes | Should -Be 8192
        $run.Report.PartialClosed | Should -BeTrue
        $run.Report.LockClosed | Should -BeTrue
        $run.Report.Result.SelectedShows | Should -Be 2
        $run.Report.Result.ProcessedShows | Should -Be 1
        $run.Report.Result.CancelledShows | Should -Be 1
        $run.Report.Result.Cancelled | Should -Be 1
        $run.Report.Result.UnstartedShowCount | Should -Be 1
        @($run.Report.Result.UnstartedShows)[0] | Should -Be 'later'
        @($run.Report.RequestTrace | Where-Object Show -eq 'later').Count | Should -Be 0
        Test-Path -LiteralPath (Join-Path $context.Root 'later') | Should -BeFalse
        $sidecars = @(Get-ChildItem -LiteralPath (Join-Path $context.Root 'cancel') -Recurse -Force -File -Filter 'resume-*.json')
        $sidecars.Count | Should -Be 1
        $checkpoint = [IO.File]::ReadAllText($sidecars[0].FullName) | ConvertFrom-Json
        $checkpoint.offset | Should -Be 8192
        $partial = Join-Path (Split-Path $sidecars[0].DirectoryName -Parent) $checkpoint.partial_name
        (Get-Item -LiteralPath $partial).Length | Should -Be 8192
        (Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash.ToLowerInvariant() | Should -Be $checkpoint.prefix_sha256
        (Get-FileHash -LiteralPath $script:savedPath -Algorithm SHA256).Hash | Should -Be $hash
    }
}
