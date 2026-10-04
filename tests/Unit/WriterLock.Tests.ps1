BeforeAll {
    $script:writerLockRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:writerLockRepository 'src/PathSafety.ps1')
    . (Join-Path $script:writerLockRepository 'src/Diagnostics.ps1')
    . (Join-Path $script:writerLockRepository 'src/HistoryStore.ps1')
    . (Join-Path $script:writerLockRepository 'src/HistoryWorkflow.ps1')
    Mock Invoke-WebRequest { throw 'Writer lock units must not request the network.' }

    function New-WriterLockFixture {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only synthetic lock files in the Pester-owned TestDrive.')]
        [CmdletBinding()]
        param([string]$Root, [ValidateSet('History', 'Archive')][string]$Scope)

        $null = [IO.Directory]::CreateDirectory($Root)
        $lockPath = Join-Path $Root '.upd-archive.lock'
        if ($Scope -eq 'History') {
            $null = [IO.Directory]::CreateDirectory((Join-Path $Root '.upd'))
            $lockPath = Join-Path $Root '.upd/writer.lock'
        }
        $bytes = [Text.Encoding]::UTF8.GetBytes('sensitive-lock-canary; pid=999999999; no ownership evidence')
        [IO.File]::WriteAllBytes($lockPath, $bytes)
        return [pscustomobject]@{ Root = $Root; LockPath = $lockPath; Original = [Convert]::ToBase64String($bytes) }
    }
}

Describe 'A042/A043 writer lock behavior' {
    It 'reports an actual held <Scope> handle as busy without changing persistent contents' -ForEach @(
        @{ Scope = 'History'; Message = 'The podcast archive writer lock is in use. Wait for the current writer to finish, then retry.' }
        @{ Scope = 'Archive'; Message = 'The archive selection lock is in use. Wait for the current writer to finish, then retry.' }
    ) {
        $fixture = New-WriterLockFixture -Root (Join-Path $TestDrive ('busy-' + $Scope)) -Scope $Scope
        $first = if ($Scope -eq 'History') { Enter-PodcastHistoryLock -Root $fixture.Root } else { Enter-PodcastArchiveLock -Root $fixture.Root }
        try {
            $failure = $null
            try {
                if ($Scope -eq 'History') { $null = Enter-PodcastHistoryLock -Root $fixture.Root }
                else { $null = Enter-PodcastArchiveLock -Root $fixture.Root }
            }
            catch { $failure = $_ }
            $failure | Should -Not -BeNullOrEmpty
            $failure.Exception.Message | Should -BeExactly $Message
            Get-PodcastDiagnosticError -Error $failure | Should -BeExactly $Message
            $failure.Exception.Message | Should -Not -Match 'sensitive-lock-canary|pid=|TestDrive'
        }
        finally {
            if ($Scope -eq 'History') { $first.Stream.Dispose() }
            else { $first.Dispose() }
        }
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($fixture.LockPath)) | Should -BeExactly $fixture.Original
    }

    It 'reports an inaccessible <Scope> lock separately from a busy writer' -ForEach @(
        @{ Scope = 'History'; Message = 'The podcast archive writer lock is inaccessible. Check destination permissions and retry.' }
        @{ Scope = 'Archive'; Message = 'The archive selection lock is inaccessible. Check output-folder permissions and retry.' }
    ) {
        $fixture = New-WriterLockFixture -Root (Join-Path $TestDrive ('readonly-' + $Scope)) -Scope $Scope
        $originalAttributes = [IO.File]::GetAttributes($fixture.LockPath)
        [IO.File]::SetAttributes($fixture.LockPath, ($originalAttributes -bor [IO.FileAttributes]::ReadOnly))
        try {
            $failure = $null
            try {
                if ($Scope -eq 'History') { $null = Enter-PodcastHistoryLock -Root $fixture.Root }
                else { $null = Enter-PodcastArchiveLock -Root $fixture.Root }
            }
            catch { $failure = $_ }
            $failure | Should -Not -BeNullOrEmpty
            $failure.Exception.Message | Should -BeExactly $Message
            Get-PodcastDiagnosticError -Error $failure | Should -BeExactly $Message
            $failure.Exception.Message | Should -Not -Match 'sensitive-lock-canary|pid=|TestDrive'
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($fixture.LockPath)) | Should -BeExactly $fixture.Original
        }
        finally { [IO.File]::SetAttributes($fixture.LockPath, $originalAttributes) }
    }

    It 'reacquires a released <Scope> lock without interpreting or truncating stale text' -ForEach @(
        @{ Scope = 'History' }, @{ Scope = 'Archive' }
    ) {
        $fixture = New-WriterLockFixture -Root (Join-Path $TestDrive ('stale-' + $Scope)) -Scope $Scope
        foreach ($iteration in 1..2) {
            $null = $iteration
            if ($Scope -eq 'History') {
                $held = Enter-PodcastHistoryLock -Root $fixture.Root
                try {
                    $held.PSObject.TypeNames | Should -Contain 'UPD.PodcastHistoryLock'
                    $held.Root | Should -BeExactly ([IO.Path]::GetFullPath($fixture.Root))
                    $held.Stream | Should -BeOfType ([IO.FileStream])
                    $held.Generation | Should -Be 0
                }
                finally { $held.Stream.Dispose() }
            }
            else {
                $held = Enter-PodcastArchiveLock -Root $fixture.Root
                try { $held | Should -BeOfType ([IO.FileStream]) }
                finally { $held.Dispose() }
            }
            Test-Path -LiteralPath $fixture.LockPath | Should -BeTrue
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($fixture.LockPath)) | Should -BeExactly $fixture.Original
        }
    }

    It 'allows separate show writer handles while output-root selection is serialized independently' {
        $outputRoot = Join-Path $TestDrive 'independent-shows'
        $left = New-WriterLockFixture -Root (Join-Path $outputRoot 'one') -Scope History
        $right = New-WriterLockFixture -Root (Join-Path $outputRoot 'two') -Scope History
        $leftLock = $null; $rightLock = $null; $selection = $null
        try {
            $leftLock = Enter-PodcastHistoryLock -Root $left.Root
            $rightLock = Enter-PodcastHistoryLock -Root $right.Root
            $selection = Enter-PodcastArchiveLock -Root $outputRoot
            $leftLock.Stream.CanWrite | Should -BeTrue
            $rightLock.Stream.CanWrite | Should -BeTrue
            $selection.CanWrite | Should -BeTrue
        }
        finally {
            if ($null -ne $selection) { $selection.Dispose() }
            if ($null -ne $rightLock) { $rightLock.Stream.Dispose() }
            if ($null -ne $leftLock) { $leftLock.Stream.Dispose() }
        }
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($left.LockPath)) | Should -BeExactly $left.Original
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($right.LockPath)) | Should -BeExactly $right.Original
    }

    It 'preserves typed cancellation from post-open state acquisition and releases its handle' {
        $fixture = New-WriterLockFixture -Root (Join-Path $TestDrive 'cancel-state') -Scope History
        Mock Read-PodcastHistory { throw [OperationCanceledException]::new('sensitive-lock-canary cancellation') }
        $failure = $null
        try { $null = Enter-PodcastHistoryLock -Root $fixture.Root }
        catch { $failure = $_ }
        $failure | Should -Not -BeNullOrEmpty
        Get-PodcastWriterLockMessage -ErrorObject $failure | Should -BeNullOrEmpty
        $proof = [IO.File]::Open($fixture.LockPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        try { $proof.CanWrite | Should -BeTrue }
        finally { $proof.Dispose() }
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($fixture.LockPath)) | Should -BeExactly $fixture.Original
    }

    It 'keeps unsafe metadata path guards ahead of directory or lock creation' {
        $root = Join-Path $TestDrive 'unsafe-metadata'
        $metadata = Join-Path $root '.upd'
        Mock Get-PodcastPathAttribute { [IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::Directory } -ParameterFilter { $LiteralPath -eq $metadata }
        { Enter-PodcastHistoryLock -Root $root } | Should -Throw '*reparse*'
        Test-Path -LiteralPath $root | Should -BeFalse
    }
}

Describe 'A042/A043 bounded native writer lock classification' {
    It 'classifies <Scope> <Kind> without publishing private error text' -ForEach @(
        @{ Scope = 'History'; Kind = 'sharing HRESULT'; Cause = [IO.IOException]::new('sensitive-lock-canary', -2147024864); Busy = $true }
        @{ Scope = 'Archive'; Kind = 'sharing HRESULT'; Cause = [IO.IOException]::new('sensitive-lock-canary', -2147024864); Busy = $true }
        @{ Scope = 'History'; Kind = 'byte-lock HRESULT'; Cause = [IO.IOException]::new('sensitive-lock-canary', -2147024863); Busy = $true }
        @{ Scope = 'Archive'; Kind = 'byte-lock HRESULT'; Cause = [IO.IOException]::new('sensitive-lock-canary', -2147024863); Busy = $true }
        @{ Scope = 'History'; Kind = 'wrapped sharing HRESULT'; Cause = [Exception]::new('outer sensitive-lock-canary', [IO.IOException]::new('inner sensitive-lock-canary', -2147024864)); Busy = $true }
        @{ Scope = 'Archive'; Kind = 'wrapped sharing HRESULT'; Cause = [Exception]::new('outer sensitive-lock-canary', [IO.IOException]::new('inner sensitive-lock-canary', -2147024864)); Busy = $true }
        @{ Scope = 'History'; Kind = 'ordinary IOException with misleading text'; Cause = [IO.IOException]::new('Another writer is active. sensitive-lock-canary'); Busy = $false }
        @{ Scope = 'Archive'; Kind = 'ordinary IOException with misleading text'; Cause = [IO.IOException]::new('Another writer is active. sensitive-lock-canary'); Busy = $false }
        @{ Scope = 'History'; Kind = 'sharing code outside Win32 HRESULT'; Cause = [IO.IOException]::new('sensitive-lock-canary', 32); Busy = $false }
        @{ Scope = 'Archive'; Kind = 'sharing code outside Win32 HRESULT'; Cause = [IO.IOException]::new('sensitive-lock-canary', 32); Busy = $false }
        @{ Scope = 'History'; Kind = 'non-IO sharing HRESULT'; Cause = [Runtime.InteropServices.ExternalException]::new('sensitive-lock-canary', -2147024864); Busy = $false }
        @{ Scope = 'Archive'; Kind = 'non-IO sharing HRESULT'; Cause = [Runtime.InteropServices.ExternalException]::new('sensitive-lock-canary', -2147024864); Busy = $false }
        @{ Scope = 'History'; Kind = 'access denied'; Cause = [UnauthorizedAccessException]::new('sensitive-lock-canary'); Busy = $false }
        @{ Scope = 'Archive'; Kind = 'access denied'; Cause = [UnauthorizedAccessException]::new('sensitive-lock-canary'); Busy = $false }
        @{ Scope = 'History'; Kind = 'untyped cancellation text'; Cause = 'OperationCanceledException sensitive-lock-canary'; Busy = $false }
        @{ Scope = 'Archive'; Kind = 'untyped cancellation text'; Cause = 'OperationCanceledException sensitive-lock-canary'; Busy = $false }
    ) {
        $message = Get-PodcastWriterLockMessage -ErrorObject $Cause -Scope $Scope
        $subject = if ($Scope -eq 'Archive') { 'The archive selection lock' } else { 'The podcast archive writer lock' }
        $expected = if ($Busy) { $subject + ' is in use. Wait for the current writer to finish, then retry.' }
            elseif ($Scope -eq 'Archive') { $subject + ' is inaccessible. Check output-folder permissions and retry.' }
            else { $subject + ' is inaccessible. Check destination permissions and retry.' }
        $message | Should -BeExactly $expected
        $message | Should -Not -Match 'sensitive-lock-canary|pid=|https?://'
    }

    It 'unwraps an ErrorRecord while refusing to call directory creation a writer conflict' {
        $cause = [IO.IOException]::new('sensitive-lock-canary', -2147024864)
        $record = [Management.Automation.ErrorRecord]::new($cause, 'writer-lock', [Management.Automation.ErrorCategory]::ResourceBusy, $null)
        Get-PodcastWriterLockMessage -ErrorObject $record | Should -BeExactly 'The podcast archive writer lock is in use. Wait for the current writer to finish, then retry.'
        Get-PodcastWriterLockMessage -ErrorObject $record -UnavailableOnly | Should -BeExactly 'The podcast archive writer lock is inaccessible. Check destination permissions and retry.'
    }

    It 'preserves <Kind> typed cancellation instead of assigning a lock message' -ForEach @(
        @{ Kind = 'operation'; Cause = [OperationCanceledException]::new('sensitive-lock-canary') }
        @{ Kind = 'task'; Cause = [Threading.Tasks.TaskCanceledException]::new('sensitive-lock-canary') }
        @{ Kind = 'pipeline'; Cause = [Management.Automation.PipelineStoppedException]::new() }
        @{ Kind = 'wrapped'; Cause = [Exception]::new('sensitive-lock-canary', [OperationCanceledException]::new()) }
    ) {
        Get-PodcastWriterLockMessage -ErrorObject $Cause | Should -BeNullOrEmpty
        Get-PodcastWriterLockMessage -ErrorObject $Cause -Scope Archive -UnavailableOnly | Should -BeNullOrEmpty
    }

    It 'bounds native exception traversal instead of searching arbitrary deep causes' {
        $cause = [IO.IOException]::new('sensitive-lock-canary', -2147024864)
        foreach ($depth in 1..16) { $cause = [Exception]::new(('level ' + $depth), $cause) }
        Get-PodcastWriterLockMessage -ErrorObject $cause | Should -BeExactly 'The podcast archive writer lock is inaccessible. Check destination permissions and retry.'
    }
}
