Describe 'A058 release-readiness reviewed-input policy' -Tag 'Unit', 'A058' {
    BeforeAll {
        $script:ReadinessRepo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $script:ReadinessScript = Join-Path $script:ReadinessRepo 'scripts/Test-ReleaseReadiness.ps1'
        $script:ReadinessEngine = (Get-Process -Id $PID).Path
        . $script:ReadinessScript
        $script:ReadinessCanonicalText = [IO.File]::ReadAllText((Join-Path $script:ReadinessRepo 'docs/codex/ACCEPTANCE_CASES.json'))

        function Invoke-UpdReadyReviewFixture {
            $canonical = $script:ReadinessCanonicalText | ConvertFrom-Json
            foreach ($row in $canonical.cases) { $row.status = 'passed' }
            $source = 'a' * 40
            $engineRuns = @('PS51', 'PS7') | ForEach-Object {
                @{ name = $_; suite = 'All'; version = $(if ($_ -eq 'PS51') { '5.1.20348.5622' } else { '7.6.6' });
                    sourceCommit = $source; total = 1791; passed = 1791; failed = 0; skipped = 0; notRun = 0; inconclusive = 0;
                    analysis = @{ files = 116; parseErrors = 0; newFindings = 0; baselineWarnings = 0 } }
            }
            $rows = foreach ($case in $canonical.cases) {
                $kind = @{ unit = 'product_unit_mocked'; integration = 'product_loopback_integration';
                    manual = 'manual_source_review'; workflow = 'workflow' }[$case.type]
                if ($case.id -in @('A041', 'A044', 'A045')) { $kind = 'actual_windows_console' }
                $proof = @(@{ path = 'docs/codex/evidence/UPD-0001.md'; line = 1; kind = $kind;
                    scope = 'Synthetic reviewed-input fixture; no product observation claimed.';
                    engines = @('PS51', 'PS7'); sourceCommit = $source })
                if ($case.id -eq 'A041') {
                    $proof += @{ path = 'docs/codex/evidence/UPD-0001.md'; line = 1; kind = 'actual_windows_gui';
                        observation = 'explorer_double_click'; scope = 'Synthetic Explorer observation fixture.'; engines = @(); sourceCommit = $source }
                }
                if ($case.id -eq 'A045') {
                    $proof[0].observation = 'physical_ctrl_c'
                    $proof += @{ path = 'docs/codex/evidence/UPD-0001.md'; line = 1; kind = 'actual_native_api';
                        scope = 'Synthetic native-API review fixture.'; engines = @('PS51', 'PS7'); sourceCommit = $source }
                }
                @{ id = $case.id; task = $case.task; type = $case.type;
                    status = $(if ($case.id -eq 'A060') { 'owner_gate' } else { 'passed' });
                    canonicalStatusAtAudit = 'not_run'; evidence = $proof; requiredUnrun = @(); limits = @();
                    ownerScopeChange = $null; currentSourceBinding = @{ sourceCommit = $source; basis = 'reviewed_unchanged_source' } }
            }
            $review = @{ schemaVersion = 1; sourceCommit = $source; sourceTree = 'b' * 40;
                candidateVersion = '0.1.0-rc.1'; candidateStatus = 'UNRELEASED_CANDIDATE'; assessmentStatus = 'NOT_READY';
                verification = @{ testedSourceCommit = $source; candidateRuntimeFingerprint = 'c' * 64;
                    testedRuntimeFingerprint = 'c' * 64; sourceBinding = 'exact'; equivalenceEvidence = $null; engines = @($engineRuns) };
                cases = @($rows) }
            # JSON round-trip models the same objects received by the CLI.
            return @{ Review = ($review | ConvertTo-Json -Depth 15 | ConvertFrom-Json); Canonical = $canonical }
        }

        function Invoke-UpdReadyAssessment {
            param([object]$Fixture)
            return Test-UpdReleaseReadiness -Review $Fixture.Review -Canonical $Fixture.Canonical -RepositoryRoot $script:ReadinessRepo
        }

        function Invoke-UpdReadinessCliProbe {
            param([object]$Fixture, [string]$Name)
            $reviewPath = Join-Path $TestDrive ($Name + '-review.json')
            $canonicalPath = Join-Path $TestDrive ($Name + '-canonical.json')
            [IO.File]::WriteAllText($reviewPath, ($Fixture.Review | ConvertTo-Json -Depth 15))
            [IO.File]::WriteAllText($canonicalPath, ($Fixture.Canonical | ConvertTo-Json -Depth 15))
            $output = & $script:ReadinessEngine -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $script:ReadinessScript -ReviewPath $reviewPath -CanonicalPath $canonicalPath
            $code = $LASTEXITCODE
            return @{ ExitCode = $code; Report = (($output -join "`n") | ConvertFrom-Json) }
        }
    }

    BeforeEach { $script:ReadinessFixture = Invoke-UpdReadyReviewFixture }

    It 'allows an explicitly complete synthetic review only as ready_for_owner_review' {
        $report = Invoke-UpdReadyAssessment $script:ReadinessFixture
        $report.assessment | Should -BeExactly 'ready_for_owner_review'
        $report.satisfiedCases | Should -Be 59
        $report.issues.Count | Should -Be 0
        $report.releaseAuthorized | Should -BeFalse
        $report.published | Should -BeFalse
        $report.ownerGate | Should -Match 'separate owner authorization'
    }

    It 'does not trust the top-level readiness label or old audit status' {
        $script:ReadinessFixture.Review.assessmentStatus = 'RELEASED'
        $script:ReadinessFixture.Review.cases[0].canonicalStatusAtAudit = 'not_run'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).assessment | Should -BeExactly 'ready_for_owner_review'
    }

    It 'keeps explicitly unrun A041 A045 and A059 conditions outstanding' {
        # Acceptance records can advance after real observations. Keep this
        # refusal regression independent of their current reviewed status.
        foreach ($id in @('A041', 'A045', 'A059')) {
            $row = @($script:ReadinessFixture.Review.cases | Where-Object { $_.id -eq $id })[0]
            $row.status = 'not_run'
            $row.requiredUnrun = @('Explicitly unrun acceptance condition in this synthetic policy fixture.')
            @($script:ReadinessFixture.Canonical.cases | Where-Object { $_.id -eq $id })[0].status = 'not_run'
        }
        $script:ReadinessFixture.Review.cases[40].evidence = @($script:ReadinessFixture.Review.cases[40].evidence | Where-Object { $_.kind -ne 'actual_windows_gui' })
        $script:ReadinessFixture.Review.cases[44].evidence[0].PSObject.Properties.Remove('observation')
        $report = Invoke-UpdReadyAssessment $script:ReadinessFixture
        $report.assessment | Should -BeExactly 'not_ready'
        foreach ($id in @('A041', 'A045', 'A059')) {
            @($report.issues | Where-Object { $_.caseId -eq $id }).Count | Should -BeGreaterThan 0
        }
        $report.releaseAuthorized | Should -BeFalse
    }

    It 'rejects a <Status> mandatory review status despite all-green run totals' -ForEach @(
        @{ Status = 'not_run' }, @{ Status = 'failed' }, @{ Status = 'skipped' },
        @{ Status = 'pending' }, @{ Status = 'deferred' }, @{ Status = 'owner_gate' }
    ) {
        $script:ReadinessFixture.Review.cases[0].status = $Status
        $report = Invoke-UpdReadyAssessment $script:ReadinessFixture
        $report.assessment | Should -BeExactly 'not_ready'
        @($report.issues | Where-Object { $_.code -eq 'case_not_passed' -and $_.caseId -eq 'A001' }).Count | Should -Be 1
    }

    It 'rejects canonical not-run status even when the review labels it passed' {
        $script:ReadinessFixture.Canonical.cases[58].status = 'not_run'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'case_not_passed'
    }

    It 'rejects a missing mandatory case' {
        $script:ReadinessFixture.Review.cases = @($script:ReadinessFixture.Review.cases | Where-Object { $_.id -ne 'A022' })
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'case_missing_or_duplicate'
    }

    It 'rejects a duplicate case replacing a different case' {
        $script:ReadinessFixture.Review.cases[1].id = 'A001'
        $report = Invoke-UpdReadyAssessment $script:ReadinessFixture
        $report.issues.code | Should -Contain 'case_inventory'
        $report.issues.code | Should -Contain 'case_missing_or_duplicate'
    }

    It 'rejects an unknown added case' {
        $script:ReadinessFixture.Review.cases += [pscustomobject]@{ id = 'A999' }
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'case_inventory'
    }

    It 'rejects changed canonical type or task metadata' {
        $script:ReadinessFixture.Review.cases[0].task = 'UPD-9999'
        $script:ReadinessFixture.Review.cases[3].type = 'workflow'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'case_definition'
    }

    It 'rejects helper-only evidence for product <Type> acceptance' -ForEach @(
        @{ Type = 'unit' }, @{ Type = 'integration' }, @{ Type = 'manual' }
    ) {
        $row = @($script:ReadinessFixture.Review.cases | Where-Object { $_.type -eq $Type })[0]
        foreach ($entry in $row.evidence) { $entry.kind = 'helper' }
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'evidence_layer'
    }

    It 'rejects source-review-only replacement of a physical Windows observation' {
        $script:ReadinessFixture.Review.cases[40].evidence[0].kind = 'manual_source_review'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'evidence_layer'
    }

    It 'rejects an unrun physical condition even when the case is labeled passed' {
        $script:ReadinessFixture.Review.cases[44].requiredUnrun = @('Physical keyboard Ctrl+C while keep-awake is active.')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'required_condition_unrun'
    }

    It 'rejects terminal-launcher-only A041 proof after status and unrun flags are cleared' {
        $script:ReadinessFixture.Review.cases[40].evidence = @($script:ReadinessFixture.Review.cases[40].evidence[0])
        $report = Invoke-UpdReadyAssessment $script:ReadinessFixture
        $report.issues.code | Should -Contain 'manual_observation'
        $report.assessment | Should -BeExactly 'not_ready'
    }

    It 'rejects typed-cancellation-only A045 proof after status and unrun flags are cleared' {
        $script:ReadinessFixture.Review.cases[44].evidence[0].PSObject.Properties.Remove('observation')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'manual_observation'
    }

    It 'requires both engines on the physical Ctrl+C observation itself' {
        $script:ReadinessFixture.Review.cases[44].evidence[0].engines = @('PS7')
        $other = $script:ReadinessFixture.Review.cases[44].evidence[0] | ConvertTo-Json -Depth 5 | ConvertFrom-Json
        $other.PSObject.Properties.Remove('observation')
        $other.engines = @('PS51', 'PS7')
        $script:ReadinessFixture.Review.cases[44].evidence += $other
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'manual_observation'
    }

    It 'rejects a missing required-unrun list rather than inferring empty' {
        $script:ReadinessFixture.Review.cases[0].PSObject.Properties.Remove('requiredUnrun')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'required_condition_unrun'
    }

    It 'rejects implicit scope approval or deferral' {
        $script:ReadinessFixture.Review.cases[40].ownerScopeChange = 'Assumed acceptable from passing tests.'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'scope_change'
    }

    It 'rejects missing concrete evidence' {
        $script:ReadinessFixture.Review.cases[0].evidence = @()
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'evidence_missing'
    }

    It 'rejects evidence <Fault> locations' -ForEach @(
        @{ Fault = 'missing'; Path = 'docs/codex/evidence/no-such-release-proof.md'; Line = 1 },
        @{ Fault = 'traversal'; Path = 'docs/../AGENTS.md'; Line = 1 },
        @{ Fault = 'outside'; Path = 'C:/unknown.md'; Line = 1 },
        @{ Fault = 'unrecorded line'; Path = 'docs/codex/evidence/UPD-0001.md'; Line = 999999 }
    ) {
        $script:ReadinessFixture.Review.cases[0].evidence[0].path = $Path
        $script:ReadinessFixture.Review.cases[0].evidence[0].line = $Line
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'evidence_location'
    }

    It 'rejects missing product engine evidence even if helper evidence has both engines' {
        $row = @($script:ReadinessFixture.Review.cases | Where-Object { $_.type -eq 'unit' })[0]
        $row.evidence[0].engines = @('PS7')
        $extra = $row.evidence[0] | ConvertTo-Json -Depth 5 | ConvertFrom-Json
        $extra.kind = 'helper'; $extra.engines = @('PS51', 'PS7')
        $row.evidence += $extra
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'evidence_engines'
    }

    It 'rejects a missing supported full-run engine' {
        $script:ReadinessFixture.Review.verification.engines = @($script:ReadinessFixture.Review.verification.engines[1])
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'engine_inventory'
    }

    It 'rejects duplicate engines rather than counting two runs' {
        $script:ReadinessFixture.Review.verification.engines[0].name = 'PS7'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'engine_inventory'
    }

    It 'rejects a focused Unit run despite positive all-pass counts' {
        $script:ReadinessFixture.Review.verification.engines[0].suite = 'Unit'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'run_incomplete'
    }

    It 'rejects full-run <Field> counts' -ForEach @(
        @{ Field = 'failed' }, @{ Field = 'skipped' }, @{ Field = 'notRun' }, @{ Field = 'inconclusive' }
    ) {
        $script:ReadinessFixture.Review.verification.engines[0].$Field = 1
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'run_incomplete'
    }

    It 'rejects absent counts instead of interpreting null as zero' {
        $script:ReadinessFixture.Review.verification.engines[0].PSObject.Properties.Remove('skipped')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'run_incomplete'
    }

    It 'rejects missing inconclusive metadata rather than assuming zero' {
        $script:ReadinessFixture.Review.verification.engines[0].PSObject.Properties.Remove('inconclusive')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'run_incomplete'
    }

    It 'rejects empty all-pass run counts and numeric strings' {
        $script:ReadinessFixture.Review.verification.engines[0].total = 0
        $script:ReadinessFixture.Review.verification.engines[0].passed = 0
        $script:ReadinessFixture.Review.verification.engines[1].passed = '1791'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'run_incomplete'
    }

    It 'rejects analyzer findings' {
        $script:ReadinessFixture.Review.verification.engines[0].analysis.newFindings = 1
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'analysis_incomplete'
    }

    It 'rejects engine version and source disagreements' {
        $script:ReadinessFixture.Review.verification.engines[0].version = '7.6.6'
        $script:ReadinessFixture.Review.verification.engines[1].sourceCommit = 'd' * 40
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'engine_source'
    }

    It 'rejects malformed dotted engine versions' {
        $script:ReadinessFixture.Review.verification.engines[0].version = '5.1.7..'
        $script:ReadinessFixture.Review.verification.engines[1].version = '7..'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'engine_source'
    }

    It 'rejects a stale case review binding' {
        $script:ReadinessFixture.Review.cases[0].currentSourceBinding.sourceCommit = 'd' * 40
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'case_source'
    }

    It 'rejects a different tested revision labeled exact' {
        $script:ReadinessFixture.Review.verification.testedSourceCommit = 'd' * 40
        foreach ($engine in $script:ReadinessFixture.Review.verification.engines) { $engine.sourceCommit = 'd' * 40 }
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'source_binding'
    }

    It 'rejects runtime equivalence without a reviewed evidence location' {
        $script:ReadinessFixture.Review.verification.testedSourceCommit = 'd' * 40
        foreach ($engine in $script:ReadinessFixture.Review.verification.engines) { $engine.sourceCommit = 'd' * 40 }
        $script:ReadinessFixture.Review.verification.sourceBinding = 'reviewed_equivalent'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'source_binding'
    }

    It 'allows explicitly reviewed unchanged runtime while preserving the tested source identity' {
        $script:ReadinessFixture.Review.verification.testedSourceCommit = 'd' * 40
        foreach ($engine in $script:ReadinessFixture.Review.verification.engines) { $engine.sourceCommit = 'd' * 40 }
        $script:ReadinessFixture.Review.verification.sourceBinding = 'reviewed_equivalent'
        $script:ReadinessFixture.Review.verification.equivalenceEvidence = [pscustomobject]@{
            path = 'docs/codex/evidence/UPD-0001.md'; line = 1; scope = 'Synthetic explicit review of equal runtime fingerprints; development/docs-only difference.'
        }
        $report = Invoke-UpdReadyAssessment $script:ReadinessFixture
        $report.assessment | Should -BeExactly 'ready_for_owner_review'
        $report.testedSourceCommit | Should -BeExactly ('d' * 40)
        $report.sourceCommit | Should -BeExactly ('a' * 40)
    }

    It 'rejects changed runtime fingerprints despite equivalence wording' {
        $script:ReadinessFixture.Review.verification.candidateRuntimeFingerprint = 'd' * 64
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'runtime_binding'
    }

    It 'rejects an expected candidate source mismatch' {
        $report = Test-UpdReleaseReadiness -Review $script:ReadinessFixture.Review -Canonical $script:ReadinessFixture.Canonical -RepositoryRoot $script:ReadinessRepo -ExpectedSourceCommit ('d' * 40)
        $report.issues.code | Should -Contain 'candidate_source'
    }

    It 'rejects a stable-release candidate label' {
        $script:ReadinessFixture.Review.candidateStatus = 'RELEASED'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'candidate_status'
    }

    It 'rejects one-element arrays in scalar <Field> metadata' -ForEach @(
        @{ Field = 'schemaVersion'; Value = @(1); Code = 'schema' },
        @{ Field = 'candidateStatus'; Value = @('UNRELEASED_CANDIDATE'); Code = 'candidate_status' },
        @{ Field = 'candidateVersion'; Value = @('0.1.0-rc.1'); Code = 'candidate_status' },
        @{ Field = 'sourceCommit'; Value = @('a' * 40); Code = 'candidate_source' }
    ) {
        $script:ReadinessFixture.Review.$Field = $Value
        $report = Invoke-UpdReadyAssessment $script:ReadinessFixture
        $report.assessment | Should -BeExactly 'not_ready'
        $report.issues.code | Should -Contain $Code
    }

    It 'rejects a one-element array source-binding enumeration' {
        $script:ReadinessFixture.Review.verification.sourceBinding = @('exact')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'source_binding'
    }

    It 'rejects a numeric-string schema version' {
        $script:ReadinessFixture.Review.schemaVersion = '1'
        $script:ReadinessFixture.Canonical.schema_version = '1'
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'schema'
    }

    It 'rejects a one-element array current-review binding basis' {
        $script:ReadinessFixture.Review.cases[0].currentSourceBinding.basis = @('reviewed_unchanged_source')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'case_source'
    }

    It 'rejects one-element arrays in engine <Field> metadata' -ForEach @(
        @{ Field = 'name'; Value = @('PS51') }, @{ Field = 'version'; Value = @('5.1.20348.5622') },
        @{ Field = 'sourceCommit'; Value = @('a' * 40) }
    ) {
        $script:ReadinessFixture.Review.verification.engines[0].$Field = $Value
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'engine_source'
    }

    It 'rejects one-element arrays in acceptance <Field> metadata' -ForEach @(
        @{ Field = 'id'; Value = @('A001'); Code = 'case_inventory' },
        @{ Field = 'status'; Value = @('passed'); Code = 'case_not_passed' },
        @{ Field = 'type'; Value = @('workflow'); Code = 'case_definition' },
        @{ Field = 'task'; Value = @('UPD-0001'); Code = 'case_definition' }
    ) {
        $script:ReadinessFixture.Review.cases[0].$Field = $Value
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain $Code
    }

    It 'rejects a one-element array canonical status' {
        $script:ReadinessFixture.Canonical.cases[0].status = @('passed')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'case_not_passed'
    }

    It 'rejects a one-element array evidence kind' {
        $script:ReadinessFixture.Review.cases[0].evidence[0].kind = @('workflow')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'evidence_kind'
    }

    It 'rejects an array observation marker instead of interpreting it as literal Explorer proof' {
        $script:ReadinessFixture.Review.cases[40].evidence[1].observation = @('explorer_double_click')
        (Invoke-UpdReadyAssessment $script:ReadinessFixture).issues.code | Should -Contain 'manual_observation'
    }

    It 'does not convert a passed A060 publication claim into release authorization' {
        $script:ReadinessFixture.Review.cases[59].status = 'passed'
        $report = Invoke-UpdReadyAssessment $script:ReadinessFixture
        $report.issues.code | Should -Contain 'owner_gate'
        $report.releaseAuthorized | Should -BeFalse
    }

    It 'returns CLI exit zero only for a complete synthetic review' {
        $result = Invoke-UpdReadinessCliProbe $script:ReadinessFixture 'ready'
        $result.ExitCode | Should -Be 0
        $result.Report.assessment | Should -BeExactly 'ready_for_owner_review'
        $result.Report.published | Should -BeFalse
    }

    It 'returns CLI exit two for outstanding acceptance despite all-green tests' {
        $script:ReadinessFixture.Review.cases[58].status = 'not_run'
        $result = Invoke-UpdReadinessCliProbe $script:ReadinessFixture 'not-ready'
        $result.ExitCode | Should -Be 2
        $result.Report.assessment | Should -BeExactly 'not_ready'
    }

    It 'returns CLI exit one for invalid input JSON' {
        $inputPath = Join-Path $TestDrive 'bad-review.json'
        [IO.File]::WriteAllText($inputPath, '{ invalid JSON')
        $priorPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            $null = & $script:ReadinessEngine -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $script:ReadinessScript -ReviewPath $inputPath 2>&1
            $code = $LASTEXITCODE
        }
        finally { $ErrorActionPreference = $priorPreference }
        $code | Should -Be 1
    }
}
