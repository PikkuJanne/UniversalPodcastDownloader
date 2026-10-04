BeforeAll {
    $retryRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $retryRepository 'src/TransportPolicy.ps1')
}

Describe 'A048 shared retry-wait context' -Tag 'Unit', 'A048' {
    BeforeEach {
        $script:RetryContextNow = [DateTimeOffset]'2026-10-03T10:00:00Z'
        $script:RetryContextClockCalls = 0
        $script:RetryContextAttempts = 0
        $script:RetryContextFailure = $null
    }

    It 'keeps an already observed zero-delay retry from sampling a later clock instant' {
        $policy = New-PodcastTransportPolicy -BaseDelaySeconds 0 -MaxDelaySeconds 0 -RetryBudgetSeconds 1 -Clock {
            $script:RetryContextClockCalls++
            if ($script:RetryContextClockCalls -gt 2) { $script:RetryContextNow = $script:RetryContextNow.AddHours(1) }
            $script:RetryContextNow
        } -Delay { throw 'A zero-delay retry must not sleep.' }
        $result = Invoke-PodcastTransportOperation -Policy $policy -Operation {
            $script:RetryContextAttempts++
            if ($script:RetryContextAttempts -eq 1) {
                throw (New-PodcastTransportException -Kind Connection -Message 'Synthetic connection failure.' -Retryable $true)
            }
            $script:RetryContextNow
        }
        $result | Should -Be ([DateTimeOffset]'2026-10-03T10:00:00Z')
        $script:RetryContextAttempts | Should -Be 2
        $script:RetryContextClockCalls | Should -Be 2
    }

    It 'preserves an injected <Kind> delay failure without a new attempt or added metadata' -ForEach @(
        @{Kind='cancellation'}, @{Kind='transport'}, @{Kind='io'}
    ) {
        $script:RetryContextFailure = switch ($Kind) {
            'cancellation' { [OperationCanceledException]::new('Synthetic delay cancellation.') }
            'transport' { New-PodcastTransportException -Kind Deferred -Message 'Synthetic custom delay failure.' }
            'io' { [IO.IOException]::new('Synthetic delay provider failure.') }
        }
        $policy = New-PodcastTransportPolicy -Clock { $script:RetryContextNow } -Delay {
            throw $script:RetryContextFailure
        }
        $caught = $null
        try {
            Invoke-PodcastTransportOperation -Policy $policy -Operation {
                $script:RetryContextAttempts++
                throw (New-PodcastTransportException -Kind Connection -Message 'Synthetic connection failure.' -Retryable $true)
            }
        }
        catch {
            $caught = $_.Exception
            for ($depth = 0; $depth -lt 16 -and $null -ne $caught.InnerException; $depth++) { $caught = $caught.InnerException }
        }
        [object]::ReferenceEquals($caught, $script:RetryContextFailure) | Should -BeTrue
        $caught.Data.Contains('Attempts') | Should -BeFalse
        $caught.Data.Contains('PodcastAttempts') | Should -BeFalse
        $script:RetryContextAttempts | Should -Be 1
        if ($Kind -eq 'cancellation') { Test-PodcastCancellation -ErrorObject $caught | Should -BeTrue }
    }
}
