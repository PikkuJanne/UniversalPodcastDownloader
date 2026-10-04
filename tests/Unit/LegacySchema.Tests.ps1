BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/PathSafety.ps1')
    . (Join-Path $script:RepositoryRoot 'src/Naming.ps1')
    . (Join-Path $script:RepositoryRoot 'src/HistoryStore.ps1')
    Mock Invoke-WebRequest { throw 'Legacy schema tests must not make network requests.' }

    function New-TestLegacyHistory {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Constructs an in-memory fixture only.')]
        [CmdletBinding()]
        param([int]$SchemaVersion = 2, [string]$Status = 'adopted')

        $state = New-PodcastHistory -FeedId ('a' * 64) -FeedAliasFingerprint ('a' * 64) -SchemaVersion $SchemaVersion
        $state.generation = 1
        $state.episodes = @([pscustomobject]@{
            episode_id = Get-PodcastNameHash -IdentityKey ('episode:' + ('a' * 64) + ':' + ('c' * 64))
            identity_source = 'rss-guid'
            identity_fingerprint = 'c' * 64
            relative_path = 'existing.mp3'
            status = $Status
            bytes = 128
            local_sha256 = 'd' * 64
            completed_utc = '2026-10-02T10:00:00Z'
            verification = [pscustomobject]@{
                method = 'owner-approved-local-signature'
                media_kind = 'mpeg-audio'
                notes = @('local-signature-only; transfer-completeness-unverified')
            }
        })
        return $state
    }
}

Describe 'A017: explicit legacy history schema evolution' -Tag 'Unit', 'A017' {
    It 'keeps new ordinary histories at version 1 and creates version 2 only when requested' {
        $ordinary = New-PodcastHistory -FeedId ('a' * 64) -FeedAliasFingerprint ('a' * 64)
        $ordinary.schema_version | Should -Be 1
        $legacy = New-PodcastHistory -FeedId ('a' * 64) -FeedAliasFingerprint ('a' * 64) -SchemaVersion 2
        $legacy.schema_version | Should -Be 2
        { Assert-PodcastHistory -State $legacy -Root $TestDrive -AllowInitial } | Should -Not -Throw
    }

    It 'keeps version 1 <Status> invalid' -ForEach @(
        @{ Status = 'adopted' },
        @{ Status = 'unverified' }
    ) {
        $state = New-TestLegacyHistory -SchemaVersion 1 -Status $Status
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Throw '*unsupported episode status*'
    }

    It 'accepts version 2 unverified records with nullable evidence' {
        $state = New-TestLegacyHistory -Status 'unverified'
        $episode = $state.episodes[0]
        $episode.bytes = $null
        $episode.local_sha256 = $null
        $episode.completed_utc = $null
        $episode.verification.method = 'legacy-inventory'
        $episode.verification.media_kind = $null
        $episode.verification.notes = @('historical-filename-only')
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Not -Throw
    }

    It 'keeps optional evidence type validation for unverified records' {
        $state = New-TestLegacyHistory -Status 'unverified'
        $state.episodes[0].bytes = '128'
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Throw '*byte count*'
    }

    It 'promotes a validated copy without changing generation, source objects or disk' {
        $root = Join-Path $TestDrive 'no-mutation'
        $state = New-TestLegacyHistory -SchemaVersion 1 -Status 'transfer_verified'
        $state.episodes[0].verification.method = 'completed-http-length-and-signature'
        $state.episodes[0].verification.notes = @('source-note')
        $before = ConvertTo-Json -InputObject $state -Depth 8 -Compress
        $promoted = ConvertTo-PodcastHistoryV2 -State $state -Root $root
        $promoted.schema_version | Should -Be 2
        $promoted.generation | Should -Be $state.generation
        $promoted.episodes[0].status | Should -BeExactly 'transfer_verified'
        $promoted.episodes[0].verification.notes[0] = 'copy-note'
        $promoted.episodes[0].relative_path = 'copy.mp3'
        $promoted.feed_alias_fingerprints[0] = 'e' * 64
        (ConvertTo-Json -InputObject $state -Depth 8 -Compress) | Should -BeExactly $before
        Test-Path -LiteralPath $root | Should -BeFalse
    }

    It 'rejects newer or corrupt state before promotion' {
        $state = New-TestLegacyHistory
        $state.schema_version = 3
        { ConvertTo-PodcastHistoryV2 -State $state -Root $TestDrive } | Should -Throw '*Unsupported history schema*'
        $state.schema_version = 1
        { ConvertTo-PodcastHistoryV2 -State $state -Root $TestDrive } | Should -Throw '*unsupported episode status*'
    }

    It 'retains version 1 reads and ordinary writes until explicit promotion' {
        $root = Join-Path $TestDrive 'v1-retained'
        $lock = Enter-PodcastHistoryLock -Root $root
        try {
            $state = New-TestLegacyHistory -SchemaVersion 1 -Status 'transfer_verified'
            $state.episodes[0].verification.method = 'completed-http-length-and-signature'
            $null = Write-PodcastHistory -Lock $lock -State $state
            $path = Join-Path $root '.upd\state.json'
            $before = [IO.File]::ReadAllBytes($path)
            $read = Read-PodcastHistory -Root $root
            $read.schema_version | Should -Be 1
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) | Should -BeExactly ([Convert]::ToBase64String($before))
            $read.generation = 2
            $null = Write-PodcastHistory -Lock $lock -State $read
            (Read-PodcastHistory -Root $root).schema_version | Should -Be 1
            (Read-PodcastHistoryFile -Root $root -RelativePath '.upd\state.json.bak').schema_version | Should -Be 1
        }
        finally { $lock.Stream.Dispose() }
    }

    It 'atomically promotes to version 2 while preserving the previous version 1 backup bytes' {
        $root = Join-Path $TestDrive 'promotion'
        $lock = Enter-PodcastHistoryLock -Root $root
        try {
            $state = New-PodcastHistory -FeedId ('a' * 64) -FeedAliasFingerprint ('a' * 64)
            $state.generation = 1
            $null = Write-PodcastHistory -Lock $lock -State $state
            $before = [IO.File]::ReadAllText((Join-Path $root '.upd\state.json'))
            $promoted = ConvertTo-PodcastHistoryV2 -State $state -Root $root
            $promoted.generation = 2
            $promoted.episodes = (New-TestLegacyHistory).episodes
            $null = Write-PodcastHistory -Lock $lock -State $promoted
            $read = Read-PodcastHistory -Root $root
            $read.schema_version | Should -Be 2
            $read.episodes[0].status | Should -BeExactly 'adopted'
            $read.episodes[0].verification.notes[0] | Should -BeExactly 'local-signature-only; transfer-completeness-unverified'
            [IO.File]::ReadAllText((Join-Path $root '.upd\state.json.bak')) | Should -BeExactly $before
            (Read-PodcastHistoryFile -Root $root -RelativePath '.upd\state.json.bak').schema_version | Should -Be 1
        }
        finally { $lock.Stream.Dispose() }
    }

    It 'rejects a schema downgrade write without changing primary or backup' {
        $root = Join-Path $TestDrive 'downgrade'
        $lock = Enter-PodcastHistoryLock -Root $root
        try {
            $state = New-PodcastHistory -FeedId ('a' * 64) -FeedAliasFingerprint ('a' * 64) -SchemaVersion 2
            $state.generation = 1
            $null = Write-PodcastHistory -Lock $lock -State $state
            $state.generation = 2
            $null = Write-PodcastHistory -Lock $lock -State $state
            $before = [IO.File]::ReadAllText((Join-Path $root '.upd\state.json'))
            $backupBefore = [IO.File]::ReadAllText((Join-Path $root '.upd\state.json.bak'))
            $state.generation = 3
            $state.schema_version = 1
            { Write-PodcastHistory -Lock $lock -State $state } | Should -Throw '*downgrade*'
            [IO.File]::ReadAllText((Join-Path $root '.upd\state.json')) | Should -BeExactly $before
            [IO.File]::ReadAllText((Join-Path $root '.upd\state.json.bak')) | Should -BeExactly $backupBefore
        }
        finally { $lock.Stream.Dispose() }
    }

    It 'rejects a manually downgraded primary paired with a version 2 backup' {
        $root = Join-Path $TestDrive 'downgraded-pair'
        $lock = Enter-PodcastHistoryLock -Root $root
        try {
            $state = New-PodcastHistory -FeedId ('a' * 64) -FeedAliasFingerprint ('a' * 64) -SchemaVersion 2
            $state.generation = 1
            $null = Write-PodcastHistory -Lock $lock -State $state
            $state.generation = 2
            $null = Write-PodcastHistory -Lock $lock -State $state
        }
        finally { $lock.Stream.Dispose() }
        $state.schema_version = 1
        $json = ConvertTo-Json -InputObject $state -Depth 8 -Compress
        [IO.File]::WriteAllText((Join-Path $root '.upd\state.json'), $json, (New-Object Text.UTF8Encoding($false)))
        { Read-PodcastHistory -Root $root } | Should -Throw '*backup contradicts*'
        [IO.File]::ReadAllText((Join-Path $root '.upd\state.json')) | Should -BeExactly $json
    }
}

Describe 'A018: adoption evidence is distinct from transfer completion' -Tag 'Unit', 'A018' {
    It 'accepts positive local evidence with explicit owner approval and limited confidence' {
        $state = New-TestLegacyHistory
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Not -Throw
    }

    It 'rejects incomplete or misleading adopted evidence: <Case>' -ForEach @(
        @{ Case = 'missing bytes'; Change = { param($e) $e.bytes = $null } },
        @{ Case = 'zero bytes'; Change = { param($e) $e.bytes = 0 } },
        @{ Case = 'missing digest'; Change = { param($e) $e.local_sha256 = $null } },
        @{ Case = 'missing timestamp'; Change = { param($e) $e.completed_utc = $null } },
        @{ Case = 'name-only method'; Change = { param($e) $e.verification.method = 'historical-filename' } },
        @{ Case = 'claimed completed HTTP transfer'; Change = { param($e) $e.verification.method = 'completed-http-length-and-signature' } },
        @{ Case = 'claimed completed EOF transfer'; Change = { param($e) $e.verification.method = 'completed-eof-and-signature' } },
        @{ Case = 'unknown signature'; Change = { param($e) $e.verification.media_kind = 'unknown' } },
        @{ Case = 'missing signature'; Change = { param($e) $e.verification.media_kind = $null } },
        @{ Case = 'missing confidence'; Change = { param($e) $e.verification.notes = @() } },
        @{ Case = 'claimed cryptographic completeness'; Change = { param($e) $e.verification.notes = @('cryptographically-complete') } }
    ) {
        $state = New-TestLegacyHistory
        & $Change $state.episodes[0]
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Throw
    }

    It 'does not allow adoption evidence to masquerade as <Status>' -ForEach @(
        @{ Status = 'prepared' },
        @{ Status = 'transfer_verified' }
    ) {
        $state = New-TestLegacyHistory -Status $Status
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Throw '*unsupported verification method*'
    }

    It 'keeps genuine transfer evidence supported in version 2' {
        $state = New-TestLegacyHistory -Status 'transfer_verified'
        $state.episodes[0].verification.method = 'completed-eof-and-signature'
        $state.episodes[0].verification.notes = @()
        { Assert-PodcastHistory -State $state -Root $TestDrive } | Should -Not -Throw
    }
}
