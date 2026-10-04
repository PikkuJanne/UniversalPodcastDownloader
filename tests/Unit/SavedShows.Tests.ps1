BeforeAll {
    $script:savedShowRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:savedShowRepository 'src/NetworkPolicy.ps1')
    . (Join-Path $script:savedShowRepository 'src/RunResult.ps1')
    $savedShowSource = Join-Path $script:savedShowRepository 'src/SavedShows.ps1'
    if (Test-Path -LiteralPath $savedShowSource) { . $savedShowSource }
    Mock Invoke-WebRequest { throw 'Saved-show units must not access the network.' }
    if (-not ('UPD.SavedShowFaultStream' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.IO;
namespace UPD {
    public sealed class SavedShowFaultStream : FileStream {
        public static int WrittenBytes;
        public static bool CancellationWritten;
        public readonly bool CancelWrite;
        public readonly bool FailDispose;
        public SavedShowFaultStream(string path, bool cancelWrite, bool failDispose)
            : base(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None) {
            CancelWrite = cancelWrite; FailDispose = failDispose;
        }
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, Math.Min(count, 8));
            WrittenBytes += Math.Min(count, 8);
            CancellationWritten = CancelWrite;
            if (CancelWrite) throw new OperationCanceledException("private write cancellation canary");
            throw new IOException("private write failure canary");
        }
        protected override void Dispose(bool disposing) {
            base.Dispose(disposing);
            if (disposing && FailDispose) throw new IOException("private disposal canary");
        }
    }
}
'@
    }
    $script:savedShowRealOpen = (Get-Command Open-PodcastSavedShowFile).ScriptBlock

    function New-UpdSavedShowFixture {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only synthetic configuration in Pester-owned TestDrive for assertions.')]
        [CmdletBinding()]
        param([string]$Directory)
        $path = Join-Path $Directory (('private-' + [guid]::NewGuid().ToString('N').Substring(0, 8)) + '/shows.json')
        $null = Save-PodcastSavedShow -ConfigPath $path -Name 'daily' -FeedUrl 'https://feeds.example.invalid/private-canary?token=token-canary' -OutputPath (Join-Path $Directory 'archive-canary')
        return $path
    }

    function Set-UpdSavedShowJson {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Corrupts only a synthetic owned config fixture to test strict preservation.')]
        [CmdletBinding()]
        param([string]$Path, [string]$Json)
        [IO.File]::WriteAllText($Path, $Json, [Text.UTF8Encoding]::new($false))
    }

    function Assert-UpdSavedShowAclRestore {
        param([Parameter(Mandatory)]$OriginalAcl, [Parameter(Mandatory)]$RestoredAcl)
        $original = [Security.AccessControl.RawSecurityDescriptor]::new($OriginalAcl.GetSecurityDescriptorBinaryForm(), 0)
        $restored = [Security.AccessControl.RawSecurityDescriptor]::new($RestoredAcl.GetSecurityDescriptorBinaryForm(), 0)
        $restored.Owner.Value | Should -BeExactly $original.Owner.Value
        $restored.Group.Value | Should -BeExactly $original.Group.Value
        $originalFlags = [int]$original.ControlFlags
        $restoredFlags = [int]$restored.ControlFlags
        $autoInherited = [int][Security.AccessControl.ControlFlags]::DiscretionaryAclAutoInherited
        ($restoredFlags -eq $originalFlags -or $restoredFlags -eq ($originalFlags -bor $autoInherited)) |
            Should -BeTrue -Because 'Windows may only add the discretionary ACL automatic-inheritance bookkeeping flag'
        $restored.ResourceManagerControl | Should -Be $original.ResourceManagerControl
        foreach ($aclName in @('DiscretionaryAcl', 'SystemAcl')) {
            $originalEntries = $original.$aclName
            $restoredEntries = $restored.$aclName
            ($null -ne $restoredEntries) | Should -Be ($null -ne $originalEntries)
            if ($null -eq $originalEntries) { continue }
            $restoredEntries.Revision | Should -Be $originalEntries.Revision
            $restoredEntries.Count | Should -Be $originalEntries.Count
            $originalBytes = [byte[]]::new($originalEntries.BinaryLength)
            $restoredBytes = [byte[]]::new($restoredEntries.BinaryLength)
            $originalEntries.GetBinaryForm($originalBytes, 0)
            $restoredEntries.GetBinaryForm($restoredBytes, 0)
            [Convert]::ToBase64String($restoredBytes) | Should -BeExactly ([Convert]::ToBase64String($originalBytes))
        }
    }
}

Describe 'A046 saved shows desired storage baseline' {
    It 'persists a feed URL only as current-user protected bytes and privately retrieves it' {
        $configPath = Join-Path $TestDrive 'protected/shows.json'
        $feedUrl = 'https://feeds.example.invalid/private-path-canary?token=private-token-canary'
        $null = Save-PodcastSavedShow -ConfigPath $configPath -Name 'daily' -FeedUrl $feedUrl -OutputPath (Join-Path $TestDrive 'archive')
        $raw = [IO.File]::ReadAllText($configPath)
        $raw | Should -Not -Match 'private-path-canary|private-token-canary|https://'
        (Get-PodcastSavedShow -ConfigPath $configPath -Name 'DAILY').FeedUrl | Should -BeExactly $feedUrl
    }

    It 'does not create configuration directories or invoke protection under WhatIf' {
        $configPath = Join-Path $TestDrive 'preview/shows.json'
        $result = Save-PodcastSavedShow -ConfigPath $configPath -Name 'daily' -FeedUrl 'https://feeds.example.invalid/feed' -OutputPath (Join-Path $TestDrive 'archive') -WhatIf
        $result.Preview | Should -BeTrue
        Test-Path -LiteralPath (Split-Path $configPath -Parent) | Should -BeFalse
    }
}

Describe 'A046 saved shows strict validation and safe lifecycle' {
    It 'uses isolated LocalAppData by default without creating any configuration' {
        $path = Get-PodcastSavedShowConfigPath
        $path | Should -BeExactly (Join-Path $env:LOCALAPPDATA 'UniversalPodcastDownloader/saved-shows/shows.json')
        $config = Read-PodcastSavedShowConfig
        $config.schema_version | Should -Be 1
        ($config.shows -is [array]) | Should -BeTrue
        $config.shows.Count | Should -Be 0
        [IO.Directory]::Exists([IO.Path]::GetDirectoryName($path)) | Should -BeFalse
    }

    It 'returns stable empty and singleton safe lists without decryption' {
        $path = Join-Path $TestDrive 'lists/shows.json'
        $empty = Get-PodcastSavedShowList -ConfigPath $path
        ($empty -is [array]) | Should -BeTrue
        $empty.Count | Should -Be 0
        $null = Save-PodcastSavedShow -ConfigPath $path -Name 'daily' -FeedUrl 'https://feeds.example.invalid/feed?secret=secret-canary' -OutputPath (Join-Path $TestDrive 'archive')
        Mock Unprotect-PodcastSavedShowFeed { throw 'Listing must not decrypt.' }
        $items = Get-PodcastSavedShowList -ConfigPath $path
        ($items -is [array]) | Should -BeTrue
        $items.Count | Should -Be 1
        ($items[0].PSObject.Properties.Name -join ',') | Should -BeExactly 'Name,Mode,CustomCount,FeedConfigured'
        ($items | ConvertTo-Json) | Should -Not -Match 'secret-canary|feed_protected|OutputPath|https://'
        Should -Invoke Unprotect-PodcastSavedShowFeed -Times 0
    }

    It 'creates protected current-user-only directory, configuration and persistent lock ACLs' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        try { $sid = $identity.User.Value } finally { $identity.Dispose() }
        foreach ($ownedPath in @([IO.Path]::GetDirectoryName($path), $path, ($path + '.lock'))) {
            $acl = Get-Acl -LiteralPath $ownedPath
            $acl.AreAccessRulesProtected | Should -BeTrue
            $acl.GetOwner([Security.Principal.SecurityIdentifier]).Value | Should -BeExactly $sid
            $rules = @($acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))
            $rules.Count | Should -Be 1
            $rules[0].IdentityReference.Value | Should -BeExactly $sid
            $rules[0].FileSystemRights | Should -Be ([Security.AccessControl.FileSystemRights]::FullControl)
            $rules[0].AccessControlType | Should -Be ([Security.AccessControl.AccessControlType]::Allow)
            $rules[0].IsInherited | Should -BeFalse
        }
    }

    It 'upserts names case-insensitively in original order and preserves another ciphertext' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $null = Save-PodcastSavedShow -ConfigPath $path -Name 'weekly' -FeedUrl 'https://feeds.example.invalid/weekly' -OutputPath (Join-Path $TestDrive 'archive') -Mode All
        $before = Read-PodcastSavedShowConfig -ConfigPath $path
        $null = Save-PodcastSavedShow -ConfigPath $path -Name 'DAILY' -FeedUrl 'https://feeds.example.invalid/changed' -OutputPath (Join-Path $TestDrive 'archive') -CustomCount 3
        $after = Read-PodcastSavedShowConfig -ConfigPath $path
        ($after.shows.name -join ',') | Should -BeExactly 'DAILY,weekly'
        $after.shows[1].feed_protected | Should -BeExactly $before.shows[1].feed_protected
        $show = Get-PodcastSavedShow -ConfigPath $path -Name 'daily'
        $show.Mode | Should -BeExactly Custom
        $show.CustomCount | Should -Be 3
    }

    It 'normalizes a case-insensitive save mode while retaining strict canonical stored data' {
        $path = Join-Path $TestDrive 'canonical-mode/shows.json'
        $null = Save-PodcastSavedShow -ConfigPath $path -Name daily -FeedUrl 'https://feeds.example.invalid/feed' -OutputPath (Join-Path $TestDrive 'archive') -Mode all
        (Get-PodcastSavedShow -ConfigPath $path -Name daily).Mode | Should -BeExactly All
        $config = Read-PodcastSavedShowConfig -ConfigPath $path
        $config.shows[0].mode = 'all'
        Set-UpdSavedShowJson -Path $path -Json (ConvertTo-Json -InputObject $config -Depth 5)
        { Read-PodcastSavedShowConfig -ConfigPath $path } | Should -Throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
    }

    It 'roundtrips a long supported URL with its larger protected representation inside the document bound' {
        $path = Join-Path $TestDrive 'unicode-feed/shows.json'
        $feed = 'https://feeds.example.invalid/' + ('a' * 32000)
        $null = Save-PodcastSavedShow -ConfigPath $path -Name daily -FeedUrl $feed -OutputPath (Join-Path $TestDrive 'archive')
        (Read-PodcastSavedShowConfig -ConfigPath $path).shows[0].feed_protected.Length | Should -BeGreaterThan 32768
        (Get-PodcastSavedShow -ConfigPath $path -Name daily).FeedUrl | Should -BeExactly $feed
    }

    It 'removes the last show into a validated empty config and retains unchanged stale lock contents' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        [IO.File]::WriteAllText(($path + '.lock'), 'opaque stale lock text; not a PID policy')
        $lockHash = (Get-FileHash -LiteralPath ($path + '.lock')).Hash
        $result = Remove-PodcastSavedShow -ConfigPath $path -Name 'DAILY'
        $result.Changed | Should -BeTrue
        (Read-PodcastSavedShowConfig -ConfigPath $path).shows.Count | Should -Be 0
        (Get-FileHash -LiteralPath ($path + '.lock')).Hash | Should -BeExactly $lockHash
        [IO.File]::Exists($path) | Should -BeTrue
    }

    It 'exports only safe metadata into a protected new file and preserves the config' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $before = (Get-FileHash -LiteralPath $path).Hash
        $target = Join-Path $TestDrive 'safe-export.json'
        $result = Export-PodcastSavedShowConfig -ConfigPath $path -Path $target
        $result.Changed | Should -BeTrue
        $raw = [IO.File]::ReadAllText($target)
        $raw | Should -Not -Match 'private-canary|token-canary|archive-canary|feed_protected|output_path|https://'
        $export = ConvertFrom-Json -InputObject $raw
        $export.schema_version | Should -Be 1
        ($export.shows[0].PSObject.Properties.Name -join ',') | Should -BeExactly 'Name,Mode,CustomCount,FeedConfigured'
        Assert-PodcastSavedShowAccess -Path $target
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $before
        { Export-PodcastSavedShowConfig -ConfigPath $path -Path $target } | Should -Throw 'Saved-show export already exists or its destination is unavailable.'
        [IO.File]::ReadAllText($target) | Should -BeExactly $raw
    }

    It 'does not decrypt or create files when save, remove and export are previews' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $before = (Get-FileHash -LiteralPath $path).Hash
        Mock Protect-PodcastSavedShowFeed { throw 'Preview must not protect.' }
        Mock Enter-PodcastSavedShowLock { throw 'Preview must not lock.' }
        $save = Save-PodcastSavedShow -ConfigPath $path -Name 'daily' -FeedUrl 'https://feeds.example.invalid/new' -OutputPath (Join-Path $TestDrive 'archive') -WhatIf
        $remove = Remove-PodcastSavedShow -ConfigPath $path -Name 'daily' -WhatIf
        $export = Export-PodcastSavedShowConfig -ConfigPath $path -Path (Join-Path $TestDrive 'preview-export.json') -WhatIf
        foreach ($result in @($save, $remove, $export)) { $result.Preview | Should -BeTrue; $result.Changed | Should -BeFalse }
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $before
        [IO.File]::Exists((Join-Path $TestDrive 'preview-export.json')) | Should -BeFalse
        Should -Invoke Protect-PodcastSavedShowFeed -Times 0
        Should -Invoke Enter-PodcastSavedShowLock -Times 0
    }

    It 'isolates unavailable ciphertext to one show while preserving structural reads and another credential' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $null = Save-PodcastSavedShow -ConfigPath $path -Name 'good' -FeedUrl 'https://feeds.example.invalid/good' -OutputPath (Join-Path $TestDrive 'archive')
        $config = Read-PodcastSavedShowConfig -ConfigPath $path
        $config.shows[0].feed_protected = 'AAAA'
        Set-UpdSavedShowJson -Path $path -Json (ConvertTo-Json -InputObject $config -Depth 5)
        $hash = (Get-FileHash -LiteralPath $path).Hash
        (Read-PodcastSavedShowConfig -ConfigPath $path).shows.Count | Should -Be 2
        { Get-PodcastSavedShow -ConfigPath $path -Name daily } | Should -Throw 'Saved-show credentials are unavailable for the current Windows user.'
        (Get-PodcastSavedShow -ConfigPath $path -Name good).FeedUrl | Should -BeExactly 'https://feeds.example.invalid/good'
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $hash
    }

    It 'rejects a decrypted user-information URL through the existing network policy' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $config = Read-PodcastSavedShowConfig -ConfigPath $path
        $config.shows[0].feed_protected = Protect-PodcastSavedShowFeed -FeedUrl 'https://private-user:private-password@feeds.example.invalid/feed'
        Set-UpdSavedShowJson -Path $path -Json (ConvertTo-Json -InputObject $config -Depth 5)
        { Get-PodcastSavedShow -ConfigPath $path -Name daily } | Should -Throw 'Saved-show credentials are unavailable for the current Windows user.'
    }

    It 'rejects malformed schema data <Label> without changing it' -ForEach @(
        @{Label='unknown-version'; Json='{"schema_version":2,"shows":[]}'},
        @{Label='fractional-version'; Json='{"schema_version":1.0,"shows":[]}'},
        @{Label='string-version'; Json='{"schema_version":"1","shows":[]}'},
        @{Label='unknown-top-field'; Json='{"schema_version":1,"shows":[],"secret":"private-canary"}'},
        @{Label='wrong-show-type'; Json='{"schema_version":1,"shows":{}}'},
        @{Label='duplicate-key'; Json='{"schema_version":1,"schema_version":1,"shows":[]}'},
        @{Label='escaped-key'; Json='{"schema_\u0076ersion":1,"shows":[]}'},
        @{Label='trailing-object-comma'; Json='{"schema_version":1,"shows":[],}'},
        @{Label='trailing-array-comma'; Json='{"schema_version":1,"shows":[{},]}'},
        @{Label='comment'; Json='{"schema_version":1,/* private-canary */"shows":[]}'},
        @{Label='multiple-roots'; Json='{"schema_version":1,"shows":[]}{"schema_version":1,"shows":[]}'},
        @{Label='string-expression'; Json='"$(New-Item injected-canary)"'},
        @{Label='empty'; Json=''}
    ) {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        Set-UpdSavedShowJson -Path $path -Json $Json
        $hash = (Get-FileHash -LiteralPath $path).Hash
        { Read-PodcastSavedShowConfig -ConfigPath $path } | Should -Throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
        { Save-PodcastSavedShow -ConfigPath $path -Name daily -FeedUrl 'https://feeds.example.invalid/new' -OutputPath (Join-Path $TestDrive 'archive') } | Should -Throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $hash
        [IO.File]::Exists((Join-Path $TestDrive 'injected-canary')) | Should -BeFalse
    }

    It 'rejects invalid show records <Label> as data' -ForEach @(
        @{Label='unknown-field'; Field='extra'; Value='private-canary'},
        @{Label='bad-name'; Field='name'; Value='$(private-canary)'},
        @{Label='bad-base64'; Field='feed_protected'; Value='not base64!'},
        @{Label='relative-output'; Field='output_path'; Value='relative/private-canary'},
        @{Label='unknown-mode'; Field='mode'; Value='Execute'},
        @{Label='latest-count'; Field='custom_count'; Value=2}
    ) {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $config = Read-PodcastSavedShowConfig -ConfigPath $path
        $config.shows[0] | Add-Member -MemberType NoteProperty -Name $Field -Value $Value -Force
        Set-UpdSavedShowJson -Path $path -Json (ConvertTo-Json -InputObject $config -Depth 5)
        { Read-PodcastSavedShowConfig -ConfigPath $path } | Should -Throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
    }

    It 'rejects case-insensitive duplicate names and a configuration over 1 MiB' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $config = Read-PodcastSavedShowConfig -ConfigPath $path
        $copy = $config.shows[0].PSObject.Copy(); $copy.name = 'DAILY'
        $config.shows = @($config.shows[0], $copy)
        Set-UpdSavedShowJson -Path $path -Json (ConvertTo-Json -InputObject $config -Depth 5)
        { Read-PodcastSavedShowConfig -ConfigPath $path } | Should -Throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
        [IO.File]::WriteAllBytes($path, [byte[]]::new(1048577))
        { Read-PodcastSavedShowConfig -ConfigPath $path } | Should -Throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
        (Get-Item -LiteralPath $path).Length | Should -Be 1048577
    }

    It 'allows a 100-show update but refuses a 101st show without changing the config' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $config = Read-PodcastSavedShowConfig -ConfigPath $path
        $sample = $config.shows[0]
        $config.shows = @(1..100 | ForEach-Object { $item = $sample.PSObject.Copy(); $item.name = "show$_"; $item })
        Set-UpdSavedShowJson -Path $path -Json (ConvertTo-Json -InputObject $config -Depth 5)
        $hash = (Get-FileHash -LiteralPath $path).Hash
        { Save-PodcastSavedShow -ConfigPath $path -Name overflow -FeedUrl 'https://feeds.example.invalid/new' -OutputPath (Join-Path $TestDrive 'archive') } | Should -Throw 'Saved-show configuration supports at most 100 shows.'
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $hash
        $null = Save-PodcastSavedShow -ConfigPath $path -Name SHOW1 -FeedUrl 'https://feeds.example.invalid/updated' -OutputPath (Join-Path $TestDrive 'archive')
        (Read-PodcastSavedShowConfig -ConfigPath $path).shows.Count | Should -Be 100
        (Get-PodcastSavedShow -ConfigPath $path -Name show1).FeedUrl | Should -BeExactly 'https://feeds.example.invalid/updated'
    }

    It 'rejects unsafe config path <Path> before creating anything' -ForEach @(
        @{Path='relative/shows.json'}, @{Path='C:shows.json'}, @{Path='C:\safe\..\shows.json'},
        @{Path='C:\safe\shows.json:stream'}, @{Path='\\?\C:\shows.json'}, @{Path='C:\safe.\shows.json'},
        @{Path='C:\safe\bad"name.json'}, @{Path='C:\safe\CON.json'}, @{Path='C:\safe\NUL\shows.json'}
    ) {
        { Read-PodcastSavedShowConfig -ConfigPath $Path } | Should -Throw 'Saved-show configuration path is unsafe or inaccessible.'
    }

    It 'reserves the generated temporary basename budget before creating any long path' {
        $parent = Join-Path $TestDrive ('x' * (211 - $TestDrive.Length - 1))
        $path = Join-Path $parent 'shows.json'
        $path.Length | Should -BeLessThan 241
        { Save-PodcastSavedShow -ConfigPath $path -Name daily -FeedUrl 'https://feeds.example.invalid/feed' -OutputPath (Join-Path $TestDrive 'archive') } | Should -Throw 'Saved-show configuration path is unsafe or inaccessible.'
        [IO.Directory]::Exists($parent) | Should -BeFalse
    }

    It 'rejects invalid UTF-8 without replacing the original bytes' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        [IO.File]::WriteAllBytes($path, [byte[]]@(0xff, 0xfe, 0x7b, 0x7d))
        $hash = (Get-FileHash -LiteralPath $path).Hash
        { Read-PodcastSavedShowConfig -ConfigPath $path } | Should -Throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $hash
    }

    It 'refuses a real owned junction before reading or writing its target configuration' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $junction = Join-Path $TestDrive 'config-junction'
        $hash = (Get-FileHash -LiteralPath $path).Hash
        try {
            $null = New-Item -ItemType Junction -Path $junction -Target ([IO.Path]::GetDirectoryName($path)) -ErrorAction Stop
            $aliased = Join-Path $junction 'shows.json'
            { Read-PodcastSavedShowConfig -ConfigPath $aliased } | Should -Throw 'Saved-show configuration path is unsafe or inaccessible.'
            { Save-PodcastSavedShow -ConfigPath $aliased -Name daily -FeedUrl 'https://feeds.example.invalid/new' -OutputPath (Join-Path $TestDrive 'archive') } | Should -Throw 'Saved-show configuration path is unsafe or inaccessible.'
            (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $hash
        }
        finally {
            if ([IO.Directory]::Exists($junction)) { [IO.Directory]::Delete($junction) }
        }
    }

    It 'rejects explicit mode/count conflicts before configuration creation' -ForEach @(
        @{Options=@{Mode='Latest';CustomCount=2}; Message='-CustomCount conflicts with an explicit Latest or All mode.'},
        @{Options=@{Mode='All';CustomCount=2}; Message='-CustomCount conflicts with an explicit Latest or All mode.'},
        @{Options=@{Mode='Custom'}; Message="Mode 'Custom' requires -CustomCount with a value >= 1."},
        @{Options=@{CustomCount=0}; Message='-CustomCount must be a positive integer.'}
    ) {
        $path = Join-Path $TestDrive 'invalid/shows.json'
        { Save-PodcastSavedShow -ConfigPath $path -Name daily -FeedUrl 'https://feeds.example.invalid/feed' -OutputPath (Join-Path $TestDrive 'archive') @Options } | Should -Throw $Message
        [IO.Directory]::Exists([IO.Path]::GetDirectoryName($path)) | Should -BeFalse
    }

    It 'refuses broad existing <Target> permissions without repairing or rewriting them' -ForEach @(@{Target='file'}, @{Target='directory'}) {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $targetPath = if ($Target -eq 'file') { $path } else { [IO.Path]::GetDirectoryName($path) }
        $original = Get-Acl -LiteralPath $targetPath
        $changed = Get-Acl -LiteralPath $targetPath
        $everyone = [Security.Principal.SecurityIdentifier]::new('S-1-1-0')
        $changed.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($everyone, [Security.AccessControl.FileSystemRights]::Read, [Security.AccessControl.AccessControlType]::Allow))
        $hash = (Get-FileHash -LiteralPath $path).Hash
        try {
            $info = if ($Target -eq 'file') { [IO.FileInfo]::new($targetPath) } else { [IO.DirectoryInfo]::new($targetPath) }
            if ('System.IO.FileSystemAclExtensions' -as [type]) { [IO.FileSystemAclExtensions]::SetAccessControl($info, $changed) }
            else { $info.SetAccessControl($changed) }
            $unsafe = (Get-Acl -LiteralPath $targetPath).Sddl
            { Read-PodcastSavedShowConfig -ConfigPath $path } | Should -Throw 'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.'
            { Save-PodcastSavedShow -ConfigPath $path -Name daily -FeedUrl 'https://feeds.example.invalid/new' -OutputPath (Join-Path $TestDrive 'archive') } | Should -Throw 'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.'
            (Get-Acl -LiteralPath $targetPath).Sddl | Should -BeExactly $unsafe
            (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $hash
        }
        finally {
            $restore = if ($Target -eq 'file') { [Security.AccessControl.FileSecurity]::new() } else { [Security.AccessControl.DirectorySecurity]::new() }
            $restore.SetSecurityDescriptorBinaryForm($original.GetSecurityDescriptorBinaryForm(), [Security.AccessControl.AccessControlSections]::Access)
            if ('System.IO.FileSystemAclExtensions' -as [type]) { [IO.FileSystemAclExtensions]::SetAccessControl($info, $restore) }
            else { $info.SetAccessControl($restore) }
            Assert-PodcastSavedShowAccess -Path $targetPath -Directory:($Target -eq 'directory')
            Assert-UpdSavedShowAclRestore -OriginalAcl $original -RestoredAcl (Get-Acl -LiteralPath $targetPath)
        }
    }

    It 'refuses a real exclusive lock and preserves the existing config and lock contents' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $before = (Get-FileHash -LiteralPath $path).Hash
        $lock = Enter-PodcastSavedShowLock -ConfigPath $path
        try {
            { Save-PodcastSavedShow -ConfigPath $path -Name daily -FeedUrl 'https://feeds.example.invalid/new' -OutputPath (Join-Path $TestDrive 'archive') } | Should -Throw 'Saved-show configuration is in use by another operation.'
            (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $before
        }
        finally { $lock.Dispose() }
        $reopened = Enter-PodcastSavedShowLock -ConfigPath $path
        $reopened.Dispose()
    }
}

Describe 'A046 saved shows primary cancellation and transactional cleanup' {
    It 'preserves typed cancellation while checking actual existing permissions and closes no unrelated handle' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        Mock Get-PodcastSavedShowIdentity { throw [OperationCanceledException]::new('private access cancellation canary') }
        $errorRecord = $null
        try { $null = Read-PodcastSavedShowConfig -ConfigPath $path } catch { $errorRecord = $_ }
        (Test-PodcastCancellation -ErrorObject $errorRecord) | Should -BeTrue
        $guard = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $guard.Dispose()
    }

    It 'preserves typed cancellation from record-path validation after opening the real config' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        Mock Resolve-PodcastSavedShowPath { throw [OperationCanceledException]::new('private output validation cancellation canary') } -ParameterFilter { $Output }
        $errorRecord = $null
        try { $null = Read-PodcastSavedShowConfig -ConfigPath $path } catch { $errorRecord = $_ }
        (Test-PodcastCancellation -ErrorObject $errorRecord) | Should -BeTrue
        $guard = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $guard.Dispose()
    }

    It 'preserves typed cancellation through DPAPI initialization' {
        Mock Initialize-PodcastSavedShowProtection { throw [OperationCanceledException]::new('private protection cancellation canary') }
        $errorRecord = $null
        try { $null = Protect-PodcastSavedShowFeed -FeedUrl 'https://feeds.example.invalid/private' } catch { $errorRecord = $_ }
        (Test-PodcastCancellation -ErrorObject $errorRecord) | Should -BeTrue
    }

    It 'preserves cancellation during actual config read and closes its real read handle' {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $hash = (Get-FileHash -LiteralPath $path).Hash
        Mock Assert-PodcastSavedShowConfig { throw [OperationCanceledException]::new('private validation cancellation canary') }
        $errorRecord = $null
        try { $null = Read-PodcastSavedShowConfig -ConfigPath $path } catch { $errorRecord = $_ }
        (Test-PodcastCancellation -ErrorObject $errorRecord) | Should -BeTrue
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $hash
        $guard = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $guard.Dispose()
    }

    It 'preserves typed cancellation through decrypted URL validation' {
        $protected = Protect-PodcastSavedShowFeed -FeedUrl 'https://feeds.example.invalid/private'
        Mock Get-PodcastRequestUri { throw [OperationCanceledException]::new('private URI cancellation canary') }
        $errorRecord = $null
        try { $null = Unprotect-PodcastSavedShowFeed -Protected $protected } catch { $errorRecord = $_ }
        (Test-PodcastCancellation -ErrorObject $errorRecord) | Should -BeTrue
    }

    It 'preserves the original generation and cleans a real partially written temp after <Label>' -ForEach @(
        @{Label='ordinary write failure'; Cancel=$false; DisposeFailure=$false},
        @{Label='typed write cancellation plus ordinary disposal failure'; Cancel=$true; DisposeFailure=$true}
    ) {
        $path = New-UpdSavedShowFixture -Directory $TestDrive
        $before = (Get-FileHash -LiteralPath $path).Hash
        $script:savedFaultCancel = $Cancel
        $script:savedFaultDispose = $DisposeFailure
        [UPD.SavedShowFaultStream]::WrittenBytes = 0
        [UPD.SavedShowFaultStream]::CancellationWritten = $false
        Mock Open-PodcastSavedShowFile {
            param($Path, $Mode)
            $null = $Mode
            $real = & $script:savedShowRealOpen -Path $Path -Mode CreateNew
            $real.Dispose()
            return [UPD.SavedShowFaultStream]::new($Path, $script:savedFaultCancel, $script:savedFaultDispose)
        } -ParameterFilter { $Path -like '*.tmp' }
        $errorRecord = $null
        try { $null = Save-PodcastSavedShow -ConfigPath $path -Name daily -FeedUrl 'https://feeds.example.invalid/new' -OutputPath (Join-Path $TestDrive 'archive') } catch { $errorRecord = $_ }
        [UPD.SavedShowFaultStream]::WrittenBytes | Should -Be 8
        [UPD.SavedShowFaultStream]::CancellationWritten | Should -Be $Cancel
        if ($Cancel) { (Test-PodcastCancellation -ErrorObject $errorRecord) | Should -BeTrue -Because $errorRecord.Exception.ToString() }
        else { $errorRecord.Exception.Message | Should -BeExactly 'Saved-show configuration path is unsafe or inaccessible.' }
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $before
        @(Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($path)) -Filter '*.tmp').Count | Should -Be 0
        $guard = Enter-PodcastSavedShowLock -ConfigPath $path
        $guard.Dispose()
    }
}
