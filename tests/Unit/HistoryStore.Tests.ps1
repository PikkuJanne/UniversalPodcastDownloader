BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/PathSafety.ps1')
    . (Join-Path $script:RepositoryRoot 'src/Naming.ps1')
    . (Join-Path $script:RepositoryRoot 'src/HistoryStore.ps1')
    Mock Invoke-WebRequest { throw 'History tests must not make network requests.' }

    function New-TestHistory {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Constructs an in-memory fixture only.')]
        [CmdletBinding()]
        param()

        $state = New-PodcastHistory -FeedId ('a' * 64) -FeedAliasFingerprint ('a' * 64)
        $state.generation = 1
        $state.episodes = @([pscustomobject]@{
            episode_id = Get-PodcastNameHash -IdentityKey ('episode:' + ('a' * 64) + ':' + ('c' * 64))
            identity_source = 'rss-guid'
            identity_fingerprint = 'c' * 64
            relative_path = 'episode.mp3'
            status = 'transfer_verified'
            bytes = 128
            local_sha256 = 'd' * 64
            completed_utc = '2026-10-02T10:00:00Z'
            verification = [pscustomobject]@{ method = 'completed-http-length-and-signature'; media_kind = 'mpeg-audio'; notes = @() }
        })
        return $state
    }

    function Write-TestRawHistory {
        param([string]$Root, [string]$Json, [string]$Name = 'state.json')
        $directory = Join-Path $Root '.upd'
        $null = [IO.Directory]::CreateDirectory($directory)
        [IO.File]::WriteAllText((Join-Path $directory $Name), $Json, (New-Object Text.UTF8Encoding($false)))
    }
}

Describe 'A015: strict versioned history validation' -Tag 'Unit', 'A015' {
    It 'constructs an initial state with unique explicit feed aliases' {
        $state = New-PodcastHistory -FeedId ('a' * 64) -FeedAliasFingerprint ('a' * 64)
        $state.generation | Should -Be 0
        $state.feed_alias_fingerprints.Count | Should -Be 1
        $state.episodes.Count | Should -Be 0
        { Assert-PodcastHistory -State $state -Root $TestDrive -AllowInitial } | Should -Not -Throw
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Throw '*generation*'
    }

    It 'rejects invalid state field <Case>' -ForEach @(
        @{ Case = 'unknown version'; Change = { param($s) $s.schema_version = 3 } },
        @{ Case = 'text version'; Change = { param($s) $s.schema_version = '1' } },
        @{ Case = 'fractional generation'; Change = { param($s) $s.generation = 1.5 } },
        @{ Case = 'text generation'; Change = { param($s) $s.generation = '1' } },
        @{ Case = 'negative generation'; Change = { param($s) $s.generation = -1 } },
        @{ Case = 'exhausted generation'; Change = { param($s) $s.generation = [long]::MaxValue } },
        @{ Case = 'feed raw URL'; Change = { param($s) $s.feed_id = 'https://synthetic.invalid/feed?token=private' } },
        @{ Case = 'alias scalar'; Change = { param($s) $s.feed_alias_fingerprints = $s.feed_id } },
        @{ Case = 'alias duplicate'; Change = { param($s) $s.feed_alias_fingerprints = @($s.feed_id, $s.feed_id) } },
        @{ Case = 'original alias missing'; Change = { param($s) $s.feed_alias_fingerprints = @('e' * 64) } },
        @{ Case = 'excess aliases'; Change = { param($s) $s.feed_alias_fingerprints = @($s.feed_id) * 129 } },
        @{ Case = 'episode scalar'; Change = { param($s) $s.episodes = $s.episodes[0] } },
        @{ Case = 'episode null'; Change = { param($s) $s.episodes = $null } },
        @{ Case = 'unknown field'; Change = { param($s) $s | Add-Member NoteProperty secret 'synthetic' } },
        @{ Case = 'missing field'; Change = { param($s) $s.PSObject.Properties.Remove('generation') } },
        @{ Case = 'episode duplicate'; Change = { param($s) $s.episodes = @($s.episodes[0], $s.episodes[0]) } },
        @{ Case = 'nonhex identity'; Change = { param($s) $s.episodes[0].episode_id = 'q' * 64 } },
        @{ Case = 'identity fingerprint mismatch'; Change = { param($s) $s.episodes[0].episode_id = 'e' * 64 } },
        @{ Case = 'invalid source'; Change = { param($s) $s.episodes[0].identity_source = 'title' } },
        @{ Case = 'rooted path'; Change = { param($s) $s.episodes[0].relative_path = 'C:\synthetic\episode.mp3' } },
        @{ Case = 'traversal path'; Change = { param($s) $s.episodes[0].relative_path = '..\episode.mp3' } },
        @{ Case = 'nested path'; Change = { param($s) $s.episodes[0].relative_path = 'nested\episode.mp3' } },
        @{ Case = 'reserved path'; Change = { param($s) $s.episodes[0].relative_path = 'NUL.mp3' } },
        @{ Case = 'metadata path'; Change = { param($s) $s.episodes[0].relative_path = '.upd' } },
        @{ Case = 'unsupported status'; Change = { param($s) $s.episodes[0].status = 'complete' } },
        @{ Case = 'string bytes'; Change = { param($s) $s.episodes[0].bytes = '128' } },
        @{ Case = 'negative bytes'; Change = { param($s) $s.episodes[0].bytes = -1 } },
        @{ Case = 'zero completed bytes'; Change = { param($s) $s.episodes[0].bytes = 0 } },
        @{ Case = 'missing completed hash'; Change = { param($s) $s.episodes[0].local_sha256 = $null } },
        @{ Case = 'invalid timestamp'; Change = { param($s) $s.episodes[0].completed_utc = '2026-19-55T10:00:00Z' } },
        @{ Case = 'non UTC timestamp'; Change = { param($s) $s.episodes[0].completed_utc = '2026-10-02T10:00:00+02:00' } },
        @{ Case = 'method URL'; Change = { param($s) $s.episodes[0].verification.method = 'https://synthetic.invalid' } },
        @{ Case = 'unknown verification'; Change = { param($s) $s.episodes[0].verification.method = 'none' } },
        @{ Case = 'unrecognized completed media'; Change = { param($s) $s.episodes[0].verification.media_kind = 'unknown' } },
        @{ Case = 'raw URL notes'; Change = { param($s) $s.episodes[0].verification.notes = @('https://synthetic.invalid?token=private') } },
        @{ Case = 'overlong notes'; Change = { param($s) $s.episodes[0].verification.notes = @('x' * 257) } },
        @{ Case = 'too many notes'; Change = { param($s) $s.episodes[0].verification.notes = @('note') * 17 } },
        @{ Case = 'scalar notes'; Change = { param($s) $s.episodes[0].verification.notes = 'note' } }
    ) {
        $state = New-TestHistory
        & $Change $state
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Throw
    }

    It 'accepts bounded evidence and nullable evidence for failed records' {
        $state = New-TestHistory
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Not -Throw
        $state.episodes[0].status = 'failed'
        $state.episodes[0].bytes = $null
        $state.episodes[0].local_sha256 = $null
        $state.episodes[0].completed_utc = $null
        $state.episodes[0].verification.media_kind = $null
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Not -Throw
    }

    It 'rejects destinations that differ only by case' {
        $state = New-TestHistory
        $second = New-TestHistory
        $second.episodes[0].identity_fingerprint = 'e' * 64
        $second.episodes[0].episode_id = Get-PodcastNameHash -IdentityKey ('episode:' + $state.feed_id + ':' + ('e' * 64))
        $second.episodes[0].relative_path = 'EPISODE.MP3'
        $state.episodes += $second.episodes[0]
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Throw '*colliding*'
    }

    It 'preserves corrupt or ambiguous JSON <Case>' -ForEach @(
        @{ Case = 'truncated'; Json = '{"schema_version":1' },
        @{ Case = 'unknown version'; Json = '{"schema_version":3,"feed_id":"a","feed_alias_fingerprints":[],"generation":1,"episodes":[]}' },
        @{ Case = 'duplicate property'; Json = '{"schema_version":1,"schema_version":2}' },
        @{ Case = 'escaped property'; Json = '{"schema_vers\u0069on":1}' },
        @{ Case = 'mixed case property'; Json = '{"schema_version":1,"SCHEMA_VERSION":1}' },
        @{ Case = 'too deeply nested'; Json = ('[' * 13) + (' ]' * 13) },
        @{ Case = 'wrong root'; Json = '[]' },
        @{ Case = 'secret-bearing invalid JSON'; Json = '{"schema_version":"https://synthetic.invalid/?secret=hidden",}' }
    ) {
        $root = Join-Path $TestDrive $Case
        Write-TestRawHistory -Root $root -Json $Json
        { Read-PodcastHistory -Root $root } | Should -Throw '*Preserved*'
        [IO.File]::ReadAllText((Join-Path $root '.upd\state.json')) | Should -BeExactly $Json
        try { Read-PodcastHistory -Root $root } catch { $_.Exception.Message | Should -Not -Match 'secret=hidden' }
    }

    It 'rejects an array wrapped valid object and timezone coercion on every engine' {
        $root = Join-Path $TestDrive 'coercion'
        $json = ConvertTo-Json -InputObject (New-TestHistory) -Depth 8 -Compress
        Write-TestRawHistory -Root $root -Json ('[' + $json + ']')
        { Read-PodcastHistory -Root $root } | Should -Throw
        Write-TestRawHistory -Root $root -Json ($json.Replace('2026-10-02T10:00:00Z', '2026-10-02T10:00:00+02:00'))
        { Read-PodcastHistory -Root $root } | Should -Throw
    }

    It 'checks history size before parsing and preserves the file' {
        $root = Join-Path $TestDrive 'oversize'
        $null = [IO.Directory]::CreateDirectory((Join-Path $root '.upd'))
        $file = Join-Path $root '.upd\state.json'
        $stream = [IO.File]::Create($file)
        try { $stream.SetLength(16777217) } finally { $stream.Dispose() }
        { Read-PodcastHistory -Root $root } | Should -Throw '*Preserved*'
        (Get-Item -LiteralPath $file).Length | Should -Be 16777217
    }
}

Describe 'A015: exclusive and atomic history persistence' -Tag 'Unit', 'A015' {
    It 'returns absent history without creating directories' {
        $root = Join-Path $TestDrive 'absent'
        Read-PodcastHistory -Root $root | Should -BeNullOrEmpty
        Test-Path -LiteralPath $root | Should -BeFalse
    }

    It 'refuses a second writer and reuses the persistent lock after release' {
        $root = Join-Path $TestDrive 'exclusive'
        $lock = Enter-PodcastHistoryLock -Root $root
        try { { Enter-PodcastHistoryLock -Root $root } | Should -Throw '*writer lock*' }
        finally { $lock.Stream.Dispose() }
        $path = Join-Path $root '.upd\writer.lock'
        Test-Path -LiteralPath $path | Should -BeTrue
        [IO.File]::WriteAllText($path, 'persistent lock fixture')
        $next = Enter-PodcastHistoryLock -Root $root
        $next.Stream.Dispose()
        [IO.File]::ReadAllText($path) | Should -BeExactly 'persistent lock fixture'
    }

    It 'round trips exact valid evidence and preserves the preceding generation as backup' {
        $root = Join-Path $TestDrive 'roundtrip'
        $lock = Enter-PodcastHistoryLock -Root $root
        try {
            $state = New-TestHistory
            $null = Write-PodcastHistory -Lock $lock -State $state
            $first = [IO.File]::ReadAllText((Join-Path $root '.upd\state.json'))
            $read = Read-PodcastHistory -Root $root
            $read.generation | Should -Be 1
            $read.episodes[0].completed_utc | Should -BeExactly '2026-10-02T10:00:00Z'
            $state.generation = 2
            $null = Write-PodcastHistory -Lock $lock -State $state
            [IO.File]::ReadAllText((Join-Path $root '.upd\state.json.bak')) | Should -BeExactly $first
            (Read-PodcastHistory -Root $root).generation | Should -Be 2
            $state.generation = 3
            $null = Write-PodcastHistory -Lock $lock -State $state
            (Read-PodcastHistoryFile -Root $root -RelativePath '.upd\state.json.bak').generation | Should -Be 2
            @(Get-ChildItem -LiteralPath (Join-Path $root '.upd') -Filter '*.tmp').Count | Should -Be 0
        }
        finally { $lock.Stream.Dispose() }
    }

    It 'rejects stale generations and a changed feed without changing persisted bytes' {
        $root = Join-Path $TestDrive 'stale'
        $lock = Enter-PodcastHistoryLock -Root $root
        try {
            $state = New-TestHistory
            $null = Write-PodcastHistory -Lock $lock -State $state
            $before = [IO.File]::ReadAllText((Join-Path $root '.upd\state.json'))
            { Write-PodcastHistory -Lock $lock -State $state } | Should -Throw '*generation*'
            $state.generation = 2
            $lock.Generation = 0
            { Write-PodcastHistory -Lock $lock -State $state } | Should -Throw '*generation*'
            $lock.Generation = 1
            $state.feed_id = 'e' * 64
            $state.feed_alias_fingerprints = @($state.feed_id)
            $state.episodes[0].episode_id = Get-PodcastNameHash -IdentityKey ('episode:' + $state.feed_id + ':' + $state.episodes[0].identity_fingerprint)
            { Write-PodcastHistory -Lock $lock -State $state } | Should -Throw '*feed identity*'
            [IO.File]::ReadAllText((Join-Path $root '.upd\state.json')) | Should -BeExactly $before
        }
        finally { $lock.Stream.Dispose() }
    }

    It 'requires a live archive-specific writer handle' {
        $root = Join-Path $TestDrive 'closed'
        $state = New-TestHistory
        { Write-PodcastHistory -Lock ([pscustomobject]@{ Root = $root; Stream = $null; Generation = 0 }) -State $state } | Should -Throw '*held*'
        $lock = Enter-PodcastHistoryLock -Root $root
        $lock.Stream.Dispose()
        { Write-PodcastHistory -Lock $lock -State $state } | Should -Throw '*held*'
        Test-Path -LiteralPath (Join-Path $root '.upd\state.json') | Should -BeFalse
    }

    It 'preserves primary state when the replacement is denied and cleans only its own temporary' {
        $root = Join-Path $TestDrive 'denied'
        $lock = Enter-PodcastHistoryLock -Root $root
        $reader = $null
        try {
            $state = New-TestHistory
            $null = Write-PodcastHistory -Lock $lock -State $state
            $path = Join-Path $root '.upd\state.json'
            $before = [IO.File]::ReadAllText($path)
            $orphan = Join-Path $root '.upd\state.unknown.tmp'
            [IO.File]::WriteAllText($orphan, 'unclaimed fixture')
            $reader = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            $state.generation = 2
            { Write-PodcastHistory -Lock $lock -State $state } | Should -Throw
            [IO.File]::ReadAllText($path) | Should -BeExactly $before
            [IO.File]::ReadAllText($orphan) | Should -BeExactly 'unclaimed fixture'
            @(Get-ChildItem -LiteralPath (Join-Path $root '.upd') -Filter '*.tmp').Count | Should -Be 1
        }
        finally {
            if ($null -ne $reader) { $reader.Dispose() }
            $lock.Stream.Dispose()
        }
    }

    It 'fails closed when only a backup exists and when that backup is corrupt' {
        $root = Join-Path $TestDrive 'backup-only'
        $json = ConvertTo-Json -InputObject (New-TestHistory) -Depth 8
        Write-TestRawHistory -Root $root -Json $json -Name 'state.json.bak'
        { Read-PodcastHistory -Root $root } | Should -Throw '*primary is missing*'
        { Enter-PodcastHistoryLock -Root $root } | Should -Throw '*primary is missing*'
        [IO.File]::ReadAllText((Join-Path $root '.upd\state.json.bak')) | Should -BeExactly $json
        Write-TestRawHistory -Root $root -Json $json
        Write-TestRawHistory -Root $root -Json '{broken' -Name 'state.json.bak'
        { Read-PodcastHistory -Root $root } | Should -Throw '*Preserved*'
        [IO.File]::ReadAllText((Join-Path $root '.upd\state.json.bak')) | Should -BeExactly '{broken'
    }

    It 'releases the writer handle when corrupt history stops acquisition' {
        $root = Join-Path $TestDrive 'corrupt-lock'
        Write-TestRawHistory -Root $root -Json '{broken'
        { Enter-PodcastHistoryLock -Root $root } | Should -Throw '*Preserved*'
        $stream = New-Object IO.FileStream((Join-Path $root '.upd\writer.lock'), [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $stream.Dispose()
    }

    It 'rejects reparse metadata before reading or opening any writer' {
        $root = Join-Path $TestDrive 'reparse'
        $metadata = Join-Path $root '.upd'
        Mock Get-PodcastPathAttribute { [IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::Directory } -ParameterFilter { $LiteralPath -eq $metadata }
        { Read-PodcastHistory -Root $root } | Should -Throw '*reparse*'
        { Enter-PodcastHistoryLock -Root $root } | Should -Throw '*reparse*'
        Test-Path -LiteralPath $root | Should -BeFalse
    }
}
