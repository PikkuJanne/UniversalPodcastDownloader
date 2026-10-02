BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/PathSafety.ps1')
    . (Join-Path $script:RepositoryRoot 'src/Naming.ps1')
    . (Join-Path $script:RepositoryRoot 'src/HistoryStore.ps1')
    . (Join-Path $script:RepositoryRoot 'src/NetworkPolicy.ps1')
    . (Join-Path $script:RepositoryRoot 'src/ResumeStore.ps1')
    . (Join-Path $script:RepositoryRoot 'src/MediaValidation.ps1')
    . (Join-Path $script:RepositoryRoot 'src/MediaRequest.ps1')
    . (Join-Path $script:RepositoryRoot 'src/MediaTransfer.ps1')
    . (Join-Path $script:RepositoryRoot 'src/HistoryWorkflow.ps1')
    Mock Get-PodcastHttpClient { throw 'Resume store units must not open a network client.' }
    Mock Invoke-WebRequest { throw 'Resume store units must not request the network.' }

    function Get-TestResumeDigest {
        param([AllowEmptyCollection()][byte[]]$Bytes)
        $sha = [Security.Cryptography.SHA256]::Create()
        try { return ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant() }
        finally { $sha.Dispose() }
    }

    function Write-TestResumeJson {
        param([Parameter(Mandatory)][string]$Json)
        [IO.File]::WriteAllText($script:SidecarPath, $Json, [Text.UTF8Encoding]::new($false))
    }
}

Describe 'A027/A029: strict resume store and durable prefix evidence' -Tag 'Unit', 'A027', 'A029' {
    BeforeEach {
        $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:Lock = Enter-PodcastHistoryLock -Root $script:Root
        $script:Prefix = [byte[]]@(1, 2, 3, 4, 5, 6, 7, 8)
        $script:State = [pscustomobject]@{
            schema_version = 1; feed_id = 'a' * 64; episode_id = 'b' * 64
            relative_path = 'episode.mp3'; partial_name = '.upd-' + [guid]::NewGuid().ToString('N') + '.tmp'
            request_fingerprint = 'c' * 64; final_uri_fingerprint = 'd' * 64
            etag = '"entity-1"'; total_length = [long]16; content_type = 'audio/mpeg'
            content_encoding = 'identity'; offset = [long]8; prefix_sha256 = Get-TestResumeDigest $script:Prefix
        }
        $script:PartialPath = Join-Path $script:Root $script:State.partial_name
        [IO.File]::WriteAllBytes($script:PartialPath, $script:Prefix)
        $script:SidecarPath = Get-PodcastResumeStatePath -Root $script:Root -EpisodeId $script:State.episode_id
        $script:OpenParameters = @{
            Lock = $script:Lock; State = $script:State; FeedId = $script:State.feed_id
            EpisodeId = $script:State.episode_id; RelativePath = $script:State.relative_path
            RequestFingerprint = $script:State.request_fingerprint
        }
    }
    AfterEach { if ($script:Lock -and $script:Lock.Stream) { $script:Lock.Stream.Dispose() } }

    It 'round trips strict evidence and opens only its exact verified prefix exclusively' {
        $saved = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $read = Read-PodcastResumeState -Root $script:Root -EpisodeId $saved.episode_id
        (ConvertTo-Json $read -Compress) | Should -BeExactly (ConvertTo-Json $saved -Compress)
        $stream = Open-PodcastResumePartial @script:OpenParameters
        try {
            $stream.Position | Should -Be 8
            $stream.Length | Should -Be 8
            $stream.CanWrite | Should -BeTrue
            $stream.CanRead | Should -BeTrue
            { [IO.File]::Open($script:PartialPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read) } | Should -Throw
            (Get-PodcastResumeStreamHash -Stream $stream) | Should -BeExactly $script:State.prefix_sha256
            $stream.Position | Should -Be 8
        }
        finally { $stream.Dispose() }
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
    }

    It 'rejects unsupported resume field <Case>' -ForEach @(
        @{ Case = 'unknown schema'; Change = { param($s) $s.schema_version = 2 } },
        @{ Case = 'text schema'; Change = { param($s) $s.schema_version = '1' } },
        @{ Case = 'unknown field'; Change = { param($s) $s | Add-Member NoteProperty secret 'private-token' } },
        @{ Case = 'missing field'; Change = { param($s) $s.PSObject.Properties.Remove('prefix_sha256') } },
        @{ Case = 'invalid digest'; Change = { param($s) $s.prefix_sha256 = 'x' * 64 } },
        @{ Case = 'raw URL identity'; Change = { param($s) $s.request_fingerprint = 'https://private.invalid/?private-token' } },
        @{ Case = 'nested target'; Change = { param($s) $s.relative_path = 'nested\episode.mp3' } },
        @{ Case = 'reserved target'; Change = { param($s) $s.relative_path = 'NUL.mp3' } },
        @{ Case = 'metadata target'; Change = { param($s) $s.relative_path = '.upd' } },
        @{ Case = 'unowned partial'; Change = { param($s) $s.partial_name = 'download.part' } },
        @{ Case = 'traversal partial'; Change = { param($s) $s.partial_name = '..\episode.mp3' } },
        @{ Case = 'weak validator'; Change = { param($s) $s.etag = 'W/"entity-1"' } },
        @{ Case = 'absent validator'; Change = { param($s) $s.etag = $null } },
        @{ Case = 'compressed representation'; Change = { param($s) $s.content_encoding = 'gzip' } },
        @{ Case = 'nonstring representation'; Change = { param($s) $s.content_encoding = @('identity') } },
        @{ Case = 'multipart representation'; Change = { param($s) $s.content_type = 'multipart/byteranges' } },
        @{ Case = 'fractional offset'; Change = { param($s) $s.offset = 2.5 } },
        @{ Case = 'text offset'; Change = { param($s) $s.offset = '8' } },
        @{ Case = 'negative offset'; Change = { param($s) $s.offset = -1 } },
        @{ Case = 'offset beyond total'; Change = { param($s) $s.offset = 17 } },
        @{ Case = 'zero total'; Change = { param($s) $s.total_length = 0 } }
    ) {
        & $Change $script:State
        { Write-PodcastResumeState -Lock $script:Lock -State $script:State } | Should -Throw
        Test-Path -LiteralPath $script:SidecarPath | Should -BeFalse
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
    }

    It 'preserves ambiguous raw resume JSON <Case> without printing private values' -ForEach @(
        @{ Case = 'duplicate'; Transform = { param($j) $j.Replace('"schema_version":1', '"schema_version":1,"schema_version":1') } },
        @{ Case = 'escaped property'; Transform = { param($j) $j.Replace('"schema_version"', '"schema_vers\u0069on"') } },
        @{ Case = 'case collision'; Transform = { param($j) $j.Replace('"schema_version":1', '"schema_version":1,"SCHEMA_VERSION":1') } },
        @{ Case = 'array root'; Transform = { param($j) '[' + $j + ']' } },
        @{ Case = 'truncated'; Transform = { param($j) $j.Substring(0, $j.Length - 2) } },
        @{ Case = 'unknown private field'; Transform = { param($j) $j.Replace('"schema_version":1', '"schema_version":1,"private":"private-token"') } }
    ) {
        $json = & $Transform (ConvertTo-Json $script:State -Compress)
        Write-TestResumeJson -Json $json
        $caught = $null
        try { Read-PodcastResumeState -Root $script:Root -EpisodeId $script:State.episode_id }
        catch { $caught = $_ }
        $caught | Should -Not -BeNullOrEmpty
        $caught.Exception.Message | Should -Match 'preserved'
        $caught.Exception.Message | Should -Not -Match 'private-token'
        [IO.File]::ReadAllText($script:SidecarPath) | Should -BeExactly $json
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
    }

    It 'preserves sidecars rejected for byte size or invalid UTF-8 <Case>' -ForEach @(
        @{ Case = 'empty'; Bytes = [byte[]]@() },
        @{ Case = 'oversize'; Bytes = [byte[]]::new(16385) },
        @{ Case = 'invalid UTF-8'; Bytes = [byte[]]@(0xc3, 0x28) }
    ) {
        [IO.File]::WriteAllBytes($script:SidecarPath, $Bytes)
        { Read-PodcastResumeState -Root $script:Root -EpisodeId $script:State.episode_id } | Should -Throw '*preserved*'
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:SidecarPath)) | Should -BeExactly ([Convert]::ToBase64String($Bytes))
    }

    It 'rejects mismatched <Field> without opening another episodes bytes' -ForEach @(
        @{ Field = 'FeedId'; Value = 'e' * 64 },
        @{ Field = 'EpisodeId'; Value = 'e' * 64 },
        @{ Field = 'RelativePath'; Value = 'other.mp3' }
    ) {
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $before = [IO.File]::ReadAllText($script:SidecarPath)
        $script:OpenParameters[$Field] = $Value
        { Open-PodcastResumePartial @script:OpenParameters } | Should -Throw '*identity*preserved*'
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
        [IO.File]::ReadAllText($script:SidecarPath) | Should -BeExactly $before
    }

    It 'checks the full episode identity when sidecar path prefixes collide' {
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $otherId = ('b' * 32) + ('e' * 32)
        { Read-PodcastResumeState -Root $script:Root -EpisodeId $otherId } | Should -Throw '*preserved*'
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
    }

    It 'preserves valid old bytes and requests a fresh transfer for a changed request URL' {
        $script:OpenParameters.RequestFingerprint = 'e' * 64
        Open-PodcastResumePartial @script:OpenParameters | Should -BeNullOrEmpty
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
    }

    It 'preserves inconsistent partial bytes for <Case>' -ForEach @(
        @{ Case = 'shorter file'; Bytes = [byte[]]@(1, 2, 3) },
        @{ Case = 'uncheckpointed tail'; Bytes = [byte[]]@(1, 2, 3, 4, 5, 6, 7, 8, 9) },
        @{ Case = 'same-size modification'; Bytes = [byte[]]@(8, 7, 6, 5, 4, 3, 2, 1) }
    ) {
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $before = [IO.File]::ReadAllText($script:SidecarPath)
        [IO.File]::WriteAllBytes($script:PartialPath, $Bytes)
        { Open-PodcastResumePartial @script:OpenParameters } | Should -Throw '*checkpoint*preserved*'
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $Bytes
        [IO.File]::ReadAllText($script:SidecarPath) | Should -BeExactly $before
    }

    It 'preserves a sidecar whose named partial is missing' {
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $before = [IO.File]::ReadAllText($script:SidecarPath)
        [IO.File]::Delete($script:PartialPath)
        { Open-PodcastResumePartial @script:OpenParameters } | Should -Throw '*preserved*'
        [IO.File]::ReadAllText($script:SidecarPath) | Should -BeExactly $before
        Test-Path -LiteralPath $script:PartialPath | Should -BeFalse
    }

    It 'does not promote full or empty checkpoints to completed media <Length>' -ForEach @(@{ Length = 0 }, @{ Length = 16 }) {
        $bytes = [byte[]]::new($Length)
        [IO.File]::WriteAllBytes($script:PartialPath, $bytes)
        $script:State.offset = [long]$Length
        $script:State.prefix_sha256 = Get-TestResumeDigest $bytes
        Open-PodcastResumePartial @script:OpenParameters | Should -BeNullOrEmpty
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:PartialPath)) | Should -BeExactly ([Convert]::ToBase64String($bytes))
        Test-Path -LiteralPath (Join-Path $script:Root 'episode.mp3') | Should -BeFalse
    }

    It 'requires a live archive-specific writer lock before state or partial mutations' {
        $script:Lock.Stream.Dispose()
        { Write-PodcastResumeState -Lock $script:Lock -State $script:State } | Should -Throw '*held*'
        { Open-PodcastResumePartial @script:OpenParameters } | Should -Throw '*held*'
        Test-Path -LiteralPath $script:SidecarPath | Should -BeFalse
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
    }

    It 'rejects a lock handle that belongs to a different archive' {
        $otherRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $other = Enter-PodcastHistoryLock -Root $otherRoot
        try {
            $other.Root = $script:Root
            { Write-PodcastResumeState -Lock $other -State $script:State } | Should -Throw '*does not belong*'
            Test-Path -LiteralPath $script:SidecarPath | Should -BeFalse
        }
        finally { $other.Stream.Dispose() }
    }

    It 'retains exact old sidecar and partial when a new owned transfer replaces the active reference' {
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $before = [IO.File]::ReadAllText($script:SidecarPath)
        $next = $script:State.PSObject.Copy()
        $next.partial_name = '.upd-' + [guid]::NewGuid().ToString('N') + '.tmp'
        [IO.File]::WriteAllBytes((Join-Path $script:Root $next.partial_name), $script:Prefix)
        $null = Write-PodcastResumeState -Lock $script:Lock -State $next -Expected $script:State
        (Read-PodcastResumeState -Root $script:Root -EpisodeId $next.episode_id).partial_name | Should -BeExactly $next.partial_name
        $old = @(Get-ChildItem -LiteralPath (Join-Path $script:Root '.upd') -Filter '*.old.json' -File)
        $old.Count | Should -Be 1
        [IO.File]::ReadAllText($old[0].FullName) | Should -BeExactly $before
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
    }

    It 'refuses stale or unexpected existing evidence without replacing it' {
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $before = [IO.File]::ReadAllText($script:SidecarPath)
        { Write-PodcastResumeState -Lock $script:Lock -State $script:State } | Should -Throw '*changed*'
        $stale = $script:State.PSObject.Copy()
        $stale.offset = 7
        { Write-PodcastResumeState -Lock $script:Lock -State $script:State -Expected $stale } | Should -Throw '*changed*'
        { Remove-PodcastResumeState -Lock $script:Lock -State $stale } | Should -Throw '*changed*'
        [IO.File]::ReadAllText($script:SidecarPath) | Should -BeExactly $before
    }

    It 'preserves the prior checkpoint when atomic replacement is denied and leaves unknown files alone' {
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $before = [IO.File]::ReadAllText($script:SidecarPath)
        $unknown = Join-Path $script:Root '.upd\resume-unknown.tmp'
        [IO.File]::WriteAllText($unknown, 'unknown bytes')
        $reader = [IO.File]::Open($script:SidecarPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        try {
            $next = $script:State.PSObject.Copy()
            $next.offset = 9
            { Write-PodcastResumeState -Lock $script:Lock -State $next -Expected $script:State } | Should -Throw
            [IO.File]::ReadAllText($script:SidecarPath) | Should -BeExactly $before
            [IO.File]::ReadAllText($unknown) | Should -BeExactly 'unknown bytes'
            @(Get-ChildItem -LiteralPath (Join-Path $script:Root '.upd') -Filter '*.tmp' -File).Count | Should -Be 1
            [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
        }
        finally { $reader.Dispose() }
    }

    It 'hashes the first nonempty checkpoint then defers small changes until terminal checkpoint' {
        $script:State.offset = [long]0
        $script:State.prefix_sha256 = Get-TestResumeDigest ([byte[]]@())
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $stream = [IO.File]::Open($script:PartialPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        try {
            $stream.Position = $stream.Length
            $session = [pscustomobject]@{ Lock = $script:Lock; State = $script:State; Stream = $stream; PartialName = $script:State.partial_name }
            Update-PodcastResumeCheckpoint -Session $session
            $session.State.offset | Should -Be 8
            $session.State.prefix_sha256 | Should -BeExactly (Get-TestResumeDigest $script:Prefix)
            $stream.WriteByte(9)
            Update-PodcastResumeCheckpoint -Session $session
            $session.State.offset | Should -Be 8
            (Read-PodcastResumeState -Root $script:Root -EpisodeId $session.State.episode_id).offset | Should -Be 8
            Update-PodcastResumeCheckpoint -Session $session -Force
            $session.State.offset | Should -Be 9
            $stream.Position | Should -Be 9
            $session.State.prefix_sha256 | Should -BeExactly (Get-TestResumeDigest ([byte[]]@(1, 2, 3, 4, 5, 6, 7, 8, 9)))
        }
        finally { $stream.Dispose() }
    }

    It 'removes only matching completed-transfer metadata and preserves partial or unknown media' {
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $unknown = Join-Path $script:Root 'unknown.part'
        [IO.File]::WriteAllText($unknown, 'unclaimed media')
        Remove-PodcastResumeState -Lock $script:Lock -State $script:State
        Test-Path -LiteralPath $script:SidecarPath | Should -BeFalse
        [IO.File]::ReadAllBytes($script:PartialPath) | Should -Be $script:Prefix
        [IO.File]::ReadAllText($unknown) | Should -BeExactly 'unclaimed media'
    }

    It 'preserves a previously resumed partial when a competing final file blocks placement after sidecar retirement' {
        $script:CompleteMedia = [IO.File]::ReadAllBytes((Join-Path $script:RepositoryRoot 'tools/codex-handoff/fixtures/silence.mp3'))
        $script:Prefix = [byte[]]$script:CompleteMedia[0..99]
        [IO.File]::WriteAllBytes($script:PartialPath, $script:Prefix)
        $uri = 'https://media.example.invalid/episode.mp3'
        $script:State.request_fingerprint = Get-PodcastResumeUriFingerprint -Uri ([uri]$uri)
        $script:State.final_uri_fingerprint = $script:State.request_fingerprint
        $script:State.offset = [long]$script:Prefix.Length
        $script:State.total_length = [long]$script:CompleteMedia.Length
        $script:State.prefix_sha256 = Get-TestResumeDigest $script:Prefix
        $null = Write-PodcastResumeState -Lock $script:Lock -State $script:State
        $script:FinalPath = Join-Path $script:Root $script:State.relative_path
        $script:PreparedResult = $null
        $script:FinalCheckpoint = $null

        Mock Invoke-PodcastMediaRequest {
            & $OnResponse ([pscustomobject]@{
                ResumeSupported = $true; FinalUriFingerprint = $script:State.final_uri_fingerprint
                ETag = $script:State.etag; TotalLength = $script:State.total_length; ContentType = 'audio/mpeg'
            })
            $DestinationStream.Write($script:CompleteMedia, [int]$Resume.Offset, ($script:CompleteMedia.Length - [int]$Resume.Offset))
            & $OnProgress $DestinationStream.Length
            [pscustomobject]@{
                Completed = $true; Bytes = $script:CompleteMedia.Length
                ContentLength = $script:CompleteMedia.Length; ContentType = 'audio/mpeg'
            }
        }
        $context = [pscustomobject]@{ Lock = $script:Lock; FeedId = $script:State.feed_id; EpisodeId = $script:State.episode_id }
        $beforeFinalize = {
            param($Result)
            $script:PreparedResult = $Result
            $script:FinalCheckpoint = Read-PodcastResumeState -Root $script:Root -EpisodeId $script:State.episode_id
            [IO.File]::WriteAllText($script:FinalPath, 'competing final bytes')
        }

        { Invoke-PodcastMediaTransfer -Uri $uri -Root $script:Root -RelativePath $script:State.relative_path `
                -ResumeContext $context -BeforeFinalize $beforeFinalize } | Should -Throw

        $script:PreparedResult.Bytes | Should -Be $script:CompleteMedia.Length
        $script:PreparedResult.Sha256 | Should -BeExactly (Get-TestResumeDigest $script:CompleteMedia)
        $script:FinalCheckpoint.offset | Should -Be $script:CompleteMedia.Length
        $script:FinalCheckpoint.partial_name | Should -BeExactly $script:State.partial_name
        Test-Path -LiteralPath $script:SidecarPath | Should -BeFalse
        [IO.File]::ReadAllText($script:FinalPath) | Should -BeExactly 'competing final bytes'
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:PartialPath)) | Should -BeExactly ([Convert]::ToBase64String($script:CompleteMedia))
        @(Get-ChildItem -LiteralPath $script:Root -Filter '.upd-*.tmp' -File).Count | Should -Be 1
        Should -Invoke Invoke-PodcastMediaRequest -Times 1 -Exactly -ParameterFilter {
            $Resume.Offset -eq 100 -and $DestinationStream.Name -eq $script:PartialPath
        }
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
    }
}
