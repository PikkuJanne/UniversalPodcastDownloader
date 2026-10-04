BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/PathSafety.ps1')
    . (Join-Path $script:RepositoryRoot 'src/Diagnostics.ps1')
    Mock Invoke-WebRequest { throw 'Diagnostic unit tests must not make network requests.' }
}

Describe 'A020: safe URL and diagnostic text boundaries' -Tag 'Unit', 'A020' {
    BeforeEach { $null = Initialize-PodcastDiagnostics -Preview }
    AfterEach { Close-PodcastDiagnostics }

    It 'retains only host and an opaque ID for <Case>' -ForEach @(
        @{ Case = 'userinfo'; Url = 'https://private-user:private-pass@podcast.invalid/feed'; HostName = 'podcast.invalid' },
        @{ Case = 'unlabelled opaque path'; Url = 'https://podcast.invalid/9f6a174e12bd64d28c72874a1/feed.xml'; HostName = 'podcast.invalid' },
        @{ Case = 'query'; Url = 'https://podcast.invalid/feed?token=private-query&x=other-secret'; HostName = 'podcast.invalid' },
        @{ Case = 'fragment'; Url = 'https://podcast.invalid/feed#private-fragment'; HostName = 'podcast.invalid' },
        @{ Case = 'escaped path'; Url = 'https://podcast.invalid/%2Fprivate-path%3Fa%3Dtoken'; HostName = 'podcast.invalid' },
        @{ Case = 'FTP credentials'; Url = 'ftp://private-user:private-pass@podcast.invalid/private-path'; HostName = 'podcast.invalid' },
        @{ Case = 'custom scheme'; Url = 'private-app://private-user:private-pass@podcast.invalid/private-path'; HostName = 'podcast.invalid' },
        @{ Case = 'IPv6 credentials'; Url = 'http://private-user:private-pass@[::1]:9876/private-path'; HostName = '::1' },
        @{ Case = 'URN'; Url = 'urn:private-token:private-body'; HostName = 'unknown-host' },
        @{ Case = 'data URL'; Url = 'data:text/plain,private-body'; HostName = 'unknown-host' },
        @{ Case = 'malformed'; Url = 'https://private-user:private-pass@[/private-path'; HostName = 'unknown-host' },
        @{ Case = 'relative'; Url = '/private-path?secret=private-query'; HostName = 'unknown-host' }
    ) {
        $display = Get-PodcastSafeUrl -Url $Url
        $display | Should -Match ('^' + [regex]::Escape($HostName) + ' \[url [a-f0-9]{32}\]$')
        $display | Should -Not -Match 'private-|9f6a174e12bd64d28c72874a1|9876|feed|token'
        (Get-PodcastSafeUrl -Url $Url) | Should -BeExactly $display
    }

    It 'uses random correlations per run and for distinct exact URLs' {
        $first = Get-PodcastSafeUrl 'https://podcast.invalid/secret-a'
        $second = Get-PodcastSafeUrl 'https://podcast.invalid/secret-b'
        $second | Should -Not -Be $first
        $null = Initialize-PodcastDiagnostics -Preview
        (Get-PodcastSafeUrl 'https://podcast.invalid/secret-a') | Should -Not -Be $first
    }

    It 'bounds the runtime URL correlation dictionary' {
        foreach ($number in 1..270) { $null = Get-PodcastSafeUrl ('https://podcast.invalid/secret-' + $number) }
        $script:PodcastDiagnosticRequestIds.Count | Should -BeLessOrEqual 256
        $before = $script:PodcastDiagnosticRequestIds.Count
        $null = Get-PodcastSafeUrl ('https://podcast.invalid/' + ('x' * 32768))
        $script:PodcastDiagnosticRequestIds.Count | Should -Be $before
    }

    It 'redacts every header name including unknown and folded headers' {
        $raw = "Authorization: Bearer private-authorization`r`nCookie: session=private-cookie`r`nX-Unfamiliar: private-custom`r`n continuation-private`r`nSome-Arbitrary-Field: private-arbitrary"
        $safe = Protect-PodcastDiagnosticText $raw
        $safe | Should -Not -Match 'private-|Bearer|session=|Authorization|Cookie|X-Unfamiliar|Some-Arbitrary'
        $safe | Should -Be '[header omitted]  [header omitted]  [header omitted]  [header omitted]'
    }

    It 'redacts complete URL tokens in authored text and strips control characters' {
        $safe = Protect-PodcastDiagnosticText ("Request https://private-user:private-pass@podcast.invalid/opaque-secret?private-query#private-fragment`nNext`e[31m`0")
        $safe | Should -Match 'Request podcast.invalid \[url [a-f0-9]{32}\]'
        $safe | Should -Not -Match 'private-|opaque-secret|[\p{Cc}\p{Cf}\p{Cs}]'
    }

    It 'omits URL tokens with schemes without hosts and malformed host syntax' {
        $safe = Protect-PodcastDiagnosticText 'Values urn:private-token data:text/plain,private-secret https://[malformed/private-token'
        $safe | Should -Not -Match 'private-|malformed|text/plain'
    }

    It 'keeps numeric application summaries but rejects free-form lookalikes' {
        $summary = 'Summary: Downloaded=1, Skipped=2, Failed=3, Adopted=4'
        (Protect-PodcastDiagnosticText $summary) | Should -BeExactly $summary
        (Protect-PodcastDiagnosticText 'Summary: Downloaded=private-token, Skipped=2, Failed=3, Adopted=4') | Should -Be '[header omitted]'
    }

    It 'bounds diagnostic messages after redaction' {
        (Protect-PodcastDiagnosticText ('x' * 3000)).Length | Should -BeLessOrEqual 2060
        $text = ('x' * 2047) + [char]::ConvertFromUtf32(0x1f3a7) + 'suffix'
        $safe = Protect-PodcastDiagnosticText $text
        { (New-Object Text.UTF8Encoding($false, $true)).GetBytes($safe) } | Should -Not -Throw
    }

    It 'summarizes errors without reading their raw body or mutating the original exception' {
        $exception = New-Object IO.IOException("private-bare-token https://podcast.invalid/private-path`r`nAuthorization: private-header")
        $record = New-Object Management.Automation.ErrorRecord($exception, 'private-error-id', [Management.Automation.ErrorCategory]::NotSpecified, 'private-target')
        $record.ErrorDetails = New-Object Management.Automation.ErrorDetails('private-details')
        $original = $exception.Message
        (Get-PodcastDiagnosticError -Error $record) | Should -Be 'The operation failed (file or transport error). Private error details were omitted.'
        [object]::ReferenceEquals($record.Exception, $exception) | Should -BeTrue
        $exception.Message | Should -BeExactly $original
        $record.ErrorDetails.Message | Should -Be 'private-details'
        $record.TargetObject | Should -Be 'private-target'
    }
}

Describe 'A020: best-effort diagnostic lifecycle' -Tag 'Unit', 'A020' {
    BeforeEach {
        $script:PreviousDiagnosticError = [Console]::Error
        $script:CapturedDiagnosticError = New-Object IO.StringWriter
        [Console]::SetError($script:CapturedDiagnosticError)
    }
    AfterEach {
        Close-PodcastDiagnostics
        [Console]::SetError($script:PreviousDiagnosticError)
        $script:CapturedDiagnosticError.Dispose()
    }

    It 'creates a unique UTF-8 startup log and a distinct run on immediate restart' {
        $root = Join-Path $TestDrive 'logs'
        $first = Initialize-PodcastDiagnostics -Roots $root
        $unicode = 'Unicode ' + [char]0x00e4 + [char]0x4e2d + [char]::ConvertFromUtf32(0x1f3a7)
        Write-PodcastDiagnostic -Message $unicode
        Close-PodcastDiagnostics
        $bytes = [IO.File]::ReadAllBytes($first.Path)
        $bytes[0] | Should -Not -Be 239
        ((New-Object Text.UTF8Encoding($false, $true)).GetString($bytes)) | Should -Match $unicode
        $second = Initialize-PodcastDiagnostics -Roots $root
        $second.RunId | Should -Not -Be $first.RunId
        $second.Path | Should -Not -Be $first.Path
        [IO.File]::Exists($first.Path) | Should -BeTrue
        [IO.File]::Exists($second.Path) | Should -BeTrue
    }

    It 'tries the second destination without exposing the rejected path or exception' {
        $badRoot = Join-Path $TestDrive 'private-path-is-file'
        [IO.File]::WriteAllText($badRoot, 'synthetic fixture')
        $goodRoot = Join-Path $TestDrive 'fallback'
        $context = Initialize-PodcastDiagnostics -Roots @($badRoot, $goodRoot)
        [IO.Path]::GetDirectoryName($context.Path) | Should -Be $goodRoot
        $script:CapturedDiagnosticError.ToString() | Should -Be ''
    }

    It 'reports sanitized stderr and retains in-memory events when every destination fails' {
        $badRoot = Join-Path $TestDrive 'private-path-is-file'
        [IO.File]::WriteAllText($badRoot, 'synthetic fixture')
        $context = Initialize-PodcastDiagnostics -Roots @($badRoot)
        Write-PodcastDiagnostic 'Request https://podcast.invalid/private-token' -Level WARN
        $context.Path | Should -BeNullOrEmpty
        $context.Events.Count | Should -Be 2
        $script:CapturedDiagnosticError.ToString() | Should -Match 'logging is unavailable'
        $script:CapturedDiagnosticError.ToString() | Should -Not -Match 'private-|IOException|Exception'
    }

    It 'never replaces the original error when the active log stream fails' {
        $context = Initialize-PodcastDiagnostics -Roots (Join-Path $TestDrive 'closed-stream')
        $context.Writer.Dispose()
        $original = New-Object InvalidOperationException('private-original-error')
        $caught = $null
        try {
            try { throw $original }
            finally { Write-PodcastDiagnostic -Message 'The operation failed.' -Level ERROR }
        }
        catch { $caught = $_.Exception }
        [object]::ReferenceEquals($caught, $original) | Should -BeTrue
        $caught.Message | Should -Be 'private-original-error'
        $context.Writer | Should -BeNullOrEmpty
        $script:CapturedDiagnosticError.ToString() | Should -Match 'original error is preserved'
        $script:CapturedDiagnosticError.ToString() | Should -Not -Match 'private-original-error'
        Write-PodcastDiagnostic 'Later request https://podcast.invalid/private-token' -Level ERROR
        $script:CapturedDiagnosticError.ToString() | Should -Match '\[ERROR\] Later request podcast.invalid \[url [a-f0-9]{32}\]'
        $script:CapturedDiagnosticError.ToString() | Should -Not -Match 'private-token'
        { Close-PodcastDiagnostics } | Should -Not -Throw
    }

    It 'keeps memory bounded and rejects level and code injection' {
        $context = Initialize-PodcastDiagnostics -Preview
        foreach ($number in 1..270) { Write-PodcastDiagnostic -Message ('Event ' + $number) }
        Write-PodcastDiagnostic -Message 'Safe event.' -Level 'private-level' -Code 'private-code'
        $context.Events.Count | Should -Be 256
        $context.Events[255].Level | Should -Be 'INFO'
        $context.Events[255].Code | Should -Be 'message'
    }

    It 'moves to an existing show log and replays bounded startup events while retaining the startup file' {
        $context = Initialize-PodcastDiagnostics -Roots (Join-Path $TestDrive 'startup')
        $startupPath = $context.Path
        Write-PodcastDiagnostic 'Discovery reached podcast.invalid.'
        $showRoot = Join-Path $TestDrive 'show'
        $null = [IO.Directory]::CreateDirectory($showRoot)
        Set-PodcastDiagnosticArchive -Root $showRoot
        $context.StartupPath | Should -Be $startupPath
        $context.Path | Should -Not -Be $startupPath
        [IO.Path]::GetDirectoryName($context.Path) | Should -Be $showRoot
        Close-PodcastDiagnostics
        [IO.File]::ReadAllText($context.Path) | Should -Match 'Discovery reached podcast.invalid'
        [IO.File]::ReadAllText($startupPath) | Should -Match 'Discovery reached podcast.invalid'
        [IO.Path]::GetFileName($context.Path) | Should -Match $context.RunId
    }

    It 'retains the usable startup log when the show directory is absent' {
        $context = Initialize-PodcastDiagnostics -Roots (Join-Path $TestDrive 'startup')
        $startupPath = $context.Path
        $absentShow = Join-Path $TestDrive 'absent-show'
        { Set-PodcastDiagnosticArchive -Root $absentShow } | Should -Not -Throw
        $context.Path | Should -Be $startupPath
        [IO.Directory]::Exists($absentShow) | Should -BeFalse
        Close-PodcastDiagnostics
        [IO.File]::ReadAllText($startupPath) | Should -Match 'retaining the startup log'
    }

    It 'creates no directories logs or export in preview mode' {
        $root = Join-Path $TestDrive 'absent-preview'
        $context = Initialize-PodcastDiagnostics -Preview -Roots $root
        Write-PodcastDiagnostic 'Preview event.'
        Set-PodcastDiagnosticArchive -Root $root
        Export-PodcastDiagnostics -Path (Join-Path $TestDrive 'preview.json')
        $context.Writer | Should -BeNullOrEmpty
        [IO.Directory]::Exists($root) | Should -BeFalse
        [IO.File]::Exists((Join-Path $TestDrive 'preview.json')) | Should -BeFalse
        $script:CapturedDiagnosticError.ToString() | Should -Be ''
    }
}

Describe 'A020: diagnostic export allowlist' -Tag 'Unit', 'A020' {
    BeforeEach { $null = Initialize-PodcastDiagnostics -Roots (Join-Path $TestDrive 'logs') }
    AfterEach { Close-PodcastDiagnostics }

    It 'exports only validated metadata and event codes even with adversarial context fields and messages' {
        $context = $script:PodcastDiagnostics
        $context | Add-Member NoteProperty RuntimeState @{ URL = 'https://podcast.invalid/private-runtime'; Authorization = 'private-auth' }
        $context | Add-Member NoteProperty History @{ Path = 'private-history'; episodes = @('private-episode') }
        $context | Add-Member NoteProperty Configuration 'private-config'
        Write-PodcastDiagnostic 'Safe authored event.' -Level WARN
        $context.Events[1].Message = 'private-bare-message'
        $context.Events[1] | Add-Member NoteProperty Headers @{ Cookie = 'private-cookie' }
        $context.Events.Add([pscustomobject]@{ TimestampUtc = [datetime]::UtcNow; Level = 'private-level'; Code = 'message'; Message = 'private-body' })
        $context.Events.Add([pscustomobject]@{ TimestampUtc = 'private-time'; Level = 'INFO'; Code = 'message'; Message = 'private-body' })
        $context.Events.Add([pscustomobject]@{ TimestampUtc = [datetime]::UtcNow; Level = 'INFO'; Code = 'private-code'; Message = 'private-body' })
        $context.Events.Add([pscustomobject]@{ TimestampUtc = [datetime]::UtcNow; Level = 'INFO'; Code = 'message'; Message = 'https://podcast.invalid/private-secret'; Extra = 'private-extra' })
        Close-PodcastDiagnostics
        [IO.File]::AppendAllText($context.Path, 'private-log-content')
        $historyPath = Join-Path $TestDrive 'state.json'
        [IO.File]::WriteAllText($historyPath, 'private-history-file')
        $export = Join-Path $TestDrive 'safe.json'
        Export-PodcastDiagnostics -Path $export
        $text = [IO.File]::ReadAllText($export)
        $text | Should -Not -Match 'private-|https:|"Message"\s*:|RuntimeState|History|Configuration|Headers|Cookie|Authorization|Writer|Path'
        $payload = $text | ConvertFrom-Json
        $payload.run_id | Should -Be $context.RunId
        $payload.events.Count | Should -Be 3
        @($payload.PSObject.Properties.Name) | Should -Be @('schema_version', 'run_id', 'started_utc', 'engine_version', 'events')
        @($payload.events[0].PSObject.Properties.Name) | Should -Be @('timestamp_utc', 'level', 'code')
        [IO.File]::ReadAllBytes($export)[0] | Should -Not -Be 239
    }

    It 'refuses invalid run identifiers before writing an export' {
        $script:PodcastDiagnostics.RunId = 'private-run-token'
        $export = Join-Path $TestDrive 'invalid-run.json'
        { Export-PodcastDiagnostics -Path $export } | Should -Throw '*metadata is invalid*'
        [IO.File]::Exists($export) | Should -BeFalse
    }

    It 'does not overwrite an existing user file' {
        $export = Join-Path $TestDrive 'existing.json'
        [IO.File]::WriteAllText($export, 'existing-owner-content')
        { Export-PodcastDiagnostics -Path $export } | Should -Throw '*could not be written*'
        [IO.File]::ReadAllText($export) | Should -Be 'existing-owner-content'
    }

    It 'honors WhatIf without creating directories or a file' {
        $directory = Join-Path $TestDrive 'whatif-absent'
        Export-PodcastDiagnostics -Path (Join-Path $directory 'safe.json') -WhatIf
        [IO.Directory]::Exists($directory) | Should -BeFalse
    }
}
