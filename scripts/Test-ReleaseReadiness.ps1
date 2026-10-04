#requires -Version 5.1
<#
.SYNOPSIS
Evaluates reviewed release-candidate acceptance metadata without publication.
.DESCRIPTION
Offline development-only policy check. It reads the current canonical register
and reviewed evidence metadata; it does not authenticate observations, source
hashes or equivalence reviews. Test counts cannot satisfy individual cases.
.PARAMETER ReviewPath
Reviewed JSON input; defaults to docs/codex/RELEASE_ACCEPTANCE.json.
.PARAMETER CanonicalPath
Current canonical JSON register; defaults to docs/codex/ACCEPTANCE_CASES.json.
.PARAMETER ExpectedSourceCommit
Optional full candidate revision that the reviewed input must match. The input
retains the original tested revision when reviewed runtime equivalence is used.
.NOTES
Exit 0 means ready_for_owner_review, 2 means not_ready, and 1 means the input
could not be read/evaluated. All outputs retain releaseAuthorized=false and
published=false. A060 owner authorization is separate from this policy check.
#>
[CmdletBinding()]
param(
    [string]$ReviewPath,
    [string]$CanonicalPath,
    [string]$ExpectedSourceCommit
)

# Development-only policy evaluation. Evidence and source equivalence are
# reviewed inputs, not independently authenticated observations or authority
# to publish. This script does not run the downloader, Git or a network client.
function Get-UpdReadinessField {
    param([object]$Value, [string]$Name)
    if ($null -eq $Value) { return $null }
    if ($Value -is [Collections.IDictionary]) { return ,($Value[$Name]) }
    $property = $Value.PSObject.Properties[$Name]
    if ($null -ne $property) { return ,($property.Value) }
    return $null
}

function Test-UpdReadinessString {
    param([object]$Value, [string[]]$Allowed)
    return ($Value -is [string] -and $Value -cin $Allowed)
}

function Test-UpdReadinessInteger {
    param([object]$Value, [long]$Minimum = 0)
    return (($Value -is [int] -or $Value -is [long]) -and $Value -ge $Minimum)
}

function Test-UpdReadinessEmptyList {
    param([object]$Value, [string]$Name)
    if ($null -eq $Value) { return $false }
    $list = $null
    if ($Value -is [Collections.IDictionary]) { $list = $Value[$Name] }
    elseif ($null -ne $Value.PSObject.Properties[$Name]) { $list = $Value.PSObject.Properties[$Name].Value }
    # Explicit [] must survive field access; normal PowerShell pipeline output
    # collapses both [] and a missing property to null.
    return ($null -ne $list -and $list -is [Collections.IList] -and $list.Count -eq 0)
}

function Test-UpdReadinessEvidenceLocation {
    param([object]$Evidence, [string]$RepositoryRoot)
    $path = Get-UpdReadinessField $Evidence 'path'
    $line = Get-UpdReadinessField $Evidence 'line'
    $scope = Get-UpdReadinessField $Evidence 'scope'
    if ($path -isnot [string] -or $path -notmatch '^docs/[A-Za-z0-9_./-]+\.md$' -or
        @($path.Split('/') | Where-Object { $_ -eq '..' -or $_ -eq '.' -or $_ -eq '' }).Count -gt 0 -or
        -not (Test-UpdReadinessInteger $line 1) -or $scope -isnot [string] -or -not $scope.Trim()) { return $false }
    $root = [IO.Path]::GetFullPath($RepositoryRoot).TrimEnd('\', '/')
    $target = [IO.Path]::GetFullPath((Join-Path $root $path))
    if (-not $target.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -or
        -not [IO.File]::Exists($target)) { return $false }
    $ancestor = Get-Item -LiteralPath $target -Force
    while ($ancestor.FullName -ne $root) {
        if (($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $false }
        $ancestor = Get-Item -LiteralPath (Split-Path $ancestor.FullName -Parent) -Force
    }
    return $line -le [IO.File]::ReadAllLines($target).Length
}

function Test-UpdReleaseReadiness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Review,
        [Parameter(Mandatory)][object]$Canonical,
        [Parameter(Mandatory)][string]$RepositoryRoot,
        [string]$ExpectedSourceCommit
    )
    $issues = New-Object Collections.ArrayList
    function Add-UpdReadinessIssue {
        param([string]$Code, [string]$CaseId, [string]$Message)
        [void]$issues.Add([pscustomobject][ordered]@{ code = $Code; caseId = $CaseId; message = $Message })
    }
    $source = Get-UpdReadinessField $Review 'sourceCommit'
    $tree = Get-UpdReadinessField $Review 'sourceTree'
    $reviewSchema = Get-UpdReadinessField $Review 'schemaVersion'
    $canonicalSchema = Get-UpdReadinessField $Canonical 'schema_version'
    if (-not (Test-UpdReadinessInteger $reviewSchema 1) -or $reviewSchema -ne 1 -or
        -not (Test-UpdReadinessInteger $canonicalSchema 1) -or $canonicalSchema -ne 1) {
        Add-UpdReadinessIssue 'schema' '' 'Both reviewed inputs must use schema version 1.'
    }
    if ($source -isnot [string] -or $source -cnotmatch '^[a-f0-9]{40}$' -or
        $tree -isnot [string] -or $tree -cnotmatch '^[a-f0-9]{40}$' -or
        ($ExpectedSourceCommit -and $source -cne $ExpectedSourceCommit)) {
        Add-UpdReadinessIssue 'candidate_source' '' 'Candidate source must be a full immutable revision matching the expected source.'
    }
    if (-not (Test-UpdReadinessString (Get-UpdReadinessField $Review 'candidateStatus') @('UNRELEASED_CANDIDATE')) -or
        -not (Test-UpdReadinessString (Get-UpdReadinessField $Review 'candidateVersion') @('0.1.0-rc.1'))) {
        Add-UpdReadinessIssue 'candidate_status' '' 'This gate assesses the unreleased 0.1.0-rc.1 candidate only.'
    }
    $verification = Get-UpdReadinessField $Review 'verification'
    $testedSource = Get-UpdReadinessField $verification 'testedSourceCommit'
    $candidateFingerprint = Get-UpdReadinessField $verification 'candidateRuntimeFingerprint'
    $testedFingerprint = Get-UpdReadinessField $verification 'testedRuntimeFingerprint'
    $binding = Get-UpdReadinessField $verification 'sourceBinding'
    $equivalence = Get-UpdReadinessField $verification 'equivalenceEvidence'
    if ($testedSource -isnot [string] -or $testedSource -cnotmatch '^[a-f0-9]{40}$' -or
        $candidateFingerprint -isnot [string] -or $candidateFingerprint -cnotmatch '^[a-f0-9]{64}$' -or
        $testedFingerprint -isnot [string] -or $testedFingerprint -cnotmatch '^[a-f0-9]{64}$' -or
        $candidateFingerprint -cne $testedFingerprint) {
        Add-UpdReadinessIssue 'runtime_binding' '' 'Reviewed candidate and tested runtime fingerprints must be present and equal.'
    }
    if (Test-UpdReadinessString $binding @('exact')) {
        if ($testedSource -cne $source -or $null -ne $equivalence) {
            Add-UpdReadinessIssue 'source_binding' '' 'Exact binding requires the same candidate and tested revision with no equivalence claim.'
        }
    }
    elseif (Test-UpdReadinessString $binding @('reviewed_equivalent')) {
        if ($testedSource -ceq $source -or -not (Test-UpdReadinessEvidenceLocation $equivalence $RepositoryRoot)) {
            Add-UpdReadinessIssue 'source_binding' '' 'Different source revisions require an explicit reviewed runtime-equivalence evidence location.'
        }
    }
    else { Add-UpdReadinessIssue 'source_binding' '' 'A reviewed exact or reviewed_equivalent source binding is required.' }

    $engines = @((Get-UpdReadinessField $verification 'engines'))
    $engineNames = @($engines | ForEach-Object { Get-UpdReadinessField $_ 'name' })
    if ($engines.Count -ne 2 -or @($engineNames | Where-Object { $_ -ceq 'PS51' }).Count -ne 1 -or
        @($engineNames | Where-Object { $_ -ceq 'PS7' }).Count -ne 1) {
        Add-UpdReadinessIssue 'engine_inventory' '' 'One complete PS51 run and one complete PS7 run are required.'
    }
    foreach ($engine in $engines) {
        $name = Get-UpdReadinessField $engine 'name'
        $version = Get-UpdReadinessField $engine 'version'
        $total = Get-UpdReadinessField $engine 'total'
        $passed = Get-UpdReadinessField $engine 'passed'
        $failed = Get-UpdReadinessField $engine 'failed'
        $skipped = Get-UpdReadinessField $engine 'skipped'
        $notRun = Get-UpdReadinessField $engine 'notRun'
        $inconclusive = Get-UpdReadinessField $engine 'inconclusive'
        $analysis = Get-UpdReadinessField $engine 'analysis'
        if (-not (Test-UpdReadinessString $name @('PS51', 'PS7')) -or $version -isnot [string] -or
            (($name -ceq 'PS51') -and ($version -notmatch '^5\.1\.[0-9]+\.[0-9]+$')) -or
            (($name -ceq 'PS7') -and ($version -notmatch '^7\.[0-9]+\.[0-9]+(?:\.[0-9]+)?$')) -or
            (Get-UpdReadinessField $engine 'sourceCommit') -isnot [string] -or
            (Get-UpdReadinessField $engine 'sourceCommit') -cne $testedSource) {
            Add-UpdReadinessIssue 'engine_source' '' 'Engine version and tested source must match the reviewed run.'
        }
        if (-not (Test-UpdReadinessString (Get-UpdReadinessField $engine 'suite') @('All')) -or
            -not (Test-UpdReadinessInteger $total 1) -or -not (Test-UpdReadinessInteger $passed 1) -or
            -not (Test-UpdReadinessInteger $failed) -or -not (Test-UpdReadinessInteger $skipped) -or
            -not (Test-UpdReadinessInteger $notRun) -or -not (Test-UpdReadinessInteger $inconclusive) -or $passed -ne $total -or
            $failed -ne 0 -or $skipped -ne 0 -or $notRun -ne 0 -or $inconclusive -ne 0) {
            Add-UpdReadinessIssue 'run_incomplete' '' 'Each full run must have explicit positive totals, all passed, and zero failed, skipped, not-run and inconclusive tests.'
        }
        if (-not (Test-UpdReadinessInteger (Get-UpdReadinessField $analysis 'files') 1) -or
            -not (Test-UpdReadinessInteger (Get-UpdReadinessField $analysis 'parseErrors')) -or
            -not (Test-UpdReadinessInteger (Get-UpdReadinessField $analysis 'newFindings')) -or
            -not (Test-UpdReadinessInteger (Get-UpdReadinessField $analysis 'baselineWarnings')) -or
            (Get-UpdReadinessField $analysis 'parseErrors') -ne 0 -or
            (Get-UpdReadinessField $analysis 'newFindings') -ne 0 -or
            (Get-UpdReadinessField $analysis 'baselineWarnings') -ne 0) {
            Add-UpdReadinessIssue 'analysis_incomplete' '' 'Both engines require complete analysis with zero parse errors, findings and baseline warnings.'
        }
    }

    $canonicalRows = @((Get-UpdReadinessField $Canonical 'cases'))
    $reviewRows = @((Get-UpdReadinessField $Review 'cases'))
    $expectedIds = @(1..60 | ForEach-Object { 'A{0:D3}' -f $_ })
    foreach ($inventory in @(@{ name = 'canonical'; rows = $canonicalRows }, @{ name = 'review'; rows = $reviewRows })) {
        $ids = @($inventory.rows | ForEach-Object { Get-UpdReadinessField $_ 'id' })
        foreach ($inventoryRow in $inventory.rows) {
            if (-not (Test-UpdReadinessString (Get-UpdReadinessField $inventoryRow 'id') $expectedIds)) {
                Add-UpdReadinessIssue 'case_inventory' '' ($inventory.name + ' case IDs must be literal A001-A060 strings.')
            }
        }
        if ($ids.Count -ne 60 -or @($ids | Where-Object { $_ -cnotin $expectedIds }).Count -gt 0 -or
            @($ids | Sort-Object -Unique).Count -ne 60) {
            Add-UpdReadinessIssue 'case_inventory' '' ($inventory.name + ' must contain exactly one of each A001-A060 case.')
        }
    }
    $satisfied = 0
    foreach ($id in $expectedIds) {
        $canonicalMatches = @($canonicalRows | Where-Object { (Get-UpdReadinessField $_ 'id') -ceq $id })
        $reviewMatches = @($reviewRows | Where-Object { (Get-UpdReadinessField $_ 'id') -ceq $id })
        if ($canonicalMatches.Count -ne 1 -or $reviewMatches.Count -ne 1) {
            Add-UpdReadinessIssue 'case_missing_or_duplicate' $id 'Exactly one canonical case and one reviewed assessment are required.'
            continue
        }
        $canonicalCase = $canonicalMatches[0]
        $row = $reviewMatches[0]
        $before = $issues.Count
        $type = Get-UpdReadinessField $row 'type'
        $status = Get-UpdReadinessField $row 'status'
        $rowTask = Get-UpdReadinessField $row 'task'
        $canonicalTask = Get-UpdReadinessField $canonicalCase 'task'
        $canonicalType = Get-UpdReadinessField $canonicalCase 'type'
        if ($rowTask -isnot [string] -or $canonicalTask -isnot [string] -or $canonicalTask -cnotmatch '^UPD-[0-9]{4}$' -or
            $rowTask -cne $canonicalTask -or $type -cne $canonicalType -or
            -not (Test-UpdReadinessString $type @('unit', 'integration', 'manual', 'workflow')) -or
            -not (Test-UpdReadinessString $canonicalType @('unit', 'integration', 'manual', 'workflow'))) {
            Add-UpdReadinessIssue 'case_definition' $id 'Task and evidence type must match the canonical acceptance case.'
        }
        $currentBinding = Get-UpdReadinessField $row 'currentSourceBinding'
        if ((Get-UpdReadinessField $currentBinding 'sourceCommit') -isnot [string] -or
            (Get-UpdReadinessField $currentBinding 'sourceCommit') -cne $source -or
            -not (Test-UpdReadinessString (Get-UpdReadinessField $currentBinding 'basis') @('current_product_rerun', 'reviewed_unchanged_source', 'current_manual_review'))) {
            Add-UpdReadinessIssue 'case_source' $id 'Each assessment must explicitly bind its review to this candidate source.'
        }
        if ($id -ceq 'A060') {
            if (-not (Test-UpdReadinessString $status @('owner_gate'))) { Add-UpdReadinessIssue 'owner_gate' $id 'Publication authority remains a separate owner gate.' }
            continue
        }
        if (-not (Test-UpdReadinessString $status @('passed')) -or
            -not (Test-UpdReadinessString (Get-UpdReadinessField $canonicalCase 'status') @('passed'))) {
            Add-UpdReadinessIssue 'case_not_passed' $id 'Both the current canonical case and its reviewed assessment must be passed.'
        }
        if (-not (Test-UpdReadinessEmptyList $row 'requiredUnrun')) {
            Add-UpdReadinessIssue 'required_condition_unrun' $id 'Required unrun conditions must be explicitly empty before readiness.'
        }
        if ($null -ne (Get-UpdReadinessField $row 'ownerScopeChange')) {
            Add-UpdReadinessIssue 'scope_change' $id 'This gate cannot approve a scope change or infer a deferral.'
        }
        $evidence = @((Get-UpdReadinessField $row 'evidence'))
        if ($evidence.Count -eq 0 -or $null -eq $evidence[0]) {
            Add-UpdReadinessIssue 'evidence_missing' $id 'Concrete reviewed evidence is required for each mandatory case.'
        }
        $kinds = @()
        $enginesByKind = @{}
        foreach ($entry in $evidence) {
            if (-not (Test-UpdReadinessEvidenceLocation $entry $RepositoryRoot)) {
                Add-UpdReadinessIssue 'evidence_location' $id 'Evidence must reference an existing repository Markdown line with a concrete scope.'
            }
            $kind = Get-UpdReadinessField $entry 'kind'
            if (-not (Test-UpdReadinessString $kind @('product_unit_mocked', 'product_loopback_integration', 'actual_windows_console', 'actual_windows_gui', 'actual_native_api', 'manual_source_review', 'workflow', 'helper'))) {
                Add-UpdReadinessIssue 'evidence_kind' $id 'Evidence layer must be explicit and recognized.'
            }
            $kinds += $kind
            $entrySource = Get-UpdReadinessField $entry 'sourceCommit'
            if ($null -ne $entrySource -and ($entrySource -isnot [string] -or $entrySource -cnotmatch '^[a-f0-9]{40}$')) {
                Add-UpdReadinessIssue 'evidence_source' $id 'Recorded historical evidence source must remain a full revision or explicit null.'
            }
            $entryEngines = @((Get-UpdReadinessField $entry 'engines'))
            if (@($entryEngines | Where-Object { $null -ne $_ -and $_ -cnotin @('PS51', 'PS7') }).Count -gt 0) {
                Add-UpdReadinessIssue 'evidence_engines' $id 'Evidence engine names must use PS51 and PS7.'
            }
            # A helper's engine counts cannot satisfy product or physical evidence.
            if (Test-UpdReadinessString $kind @('product_unit_mocked', 'product_loopback_integration', 'actual_windows_console', 'actual_native_api')) {
                if (-not $enginesByKind.ContainsKey($kind)) { $enginesByKind[$kind] = @() }
                $enginesByKind[$kind] += $entryEngines
            }
        }
        $requiredKind = @{ unit = 'product_unit_mocked'; integration = 'product_loopback_integration'; workflow = 'workflow' }
        if (Test-UpdReadinessString $type @('manual')) {
            if (@($kinds | Where-Object { $_ -cin @('actual_windows_console', 'actual_windows_gui', 'actual_native_api', 'manual_source_review') }).Count -eq 0) {
                Add-UpdReadinessIssue 'evidence_layer' $id 'Manual acceptance needs an actual manual observation or explicit source review.'
            }
            $physicalKinds = @{ A041 = @('actual_windows_console'); A044 = @('actual_windows_console'); A045 = @('actual_windows_console', 'actual_native_api') }
            foreach ($physicalKind in $physicalKinds[$id]) {
                if ($physicalKind -cnotin $kinds) { Add-UpdReadinessIssue 'evidence_layer' $id 'This manual case requires its actual Windows console or native API observation.' }
            }
            if ($id -ceq 'A041' -and @($evidence | Where-Object {
                (Test-UpdReadinessString (Get-UpdReadinessField $_ 'kind') @('actual_windows_gui')) -and
                (Test-UpdReadinessString (Get-UpdReadinessField $_ 'observation') @('explorer_double_click'))
            }).Count -eq 0) {
                Add-UpdReadinessIssue 'manual_observation' $id 'A041 needs an explicitly observed Explorer double-click; terminal launcher checks do not prove it.'
            }
            if ($id -ceq 'A045') {
                $ctrlCEngines = @($evidence | Where-Object {
                    (Test-UpdReadinessString (Get-UpdReadinessField $_ 'kind') @('actual_windows_console')) -and
                    (Test-UpdReadinessString (Get-UpdReadinessField $_ 'observation') @('physical_ctrl_c'))
                } | ForEach-Object { (Get-UpdReadinessField $_ 'engines') })
                if ('PS51' -cnotin $ctrlCEngines -or 'PS7' -cnotin $ctrlCEngines) {
                    Add-UpdReadinessIssue 'manual_observation' $id 'A045 needs physical keyboard Ctrl+C observations in both Windows engines; typed cancellation is distinct.'
                }
            }
        }
        elseif ($type -is [string] -and $requiredKind.ContainsKey($type) -and $requiredKind[$type] -cnotin $kinds) {
            Add-UpdReadinessIssue 'evidence_layer' $id 'Helper or bundle validation cannot substitute for the canonical product or workflow evidence layer.'
        }
        $engineKinds = if (Test-UpdReadinessString $type @('unit', 'integration')) { @($requiredKind[$type]) }
            else { @($kinds | Where-Object { $_ -cin @('actual_windows_console', 'actual_native_api') } | Sort-Object -Unique) }
        foreach ($engineKind in $engineKinds) {
            if ('PS51' -cnotin $enginesByKind[$engineKind] -or 'PS7' -cnotin $enginesByKind[$engineKind]) {
                Add-UpdReadinessIssue 'evidence_engines' $id 'Applicable product and physical observations must include both supported Windows engines.'
            }
        }
        if ($issues.Count -eq $before) { $satisfied++ }
    }
    $ready = $issues.Count -eq 0
    return [pscustomobject][ordered]@{
        schemaVersion = 1
        assessment = $(if ($ready) { 'ready_for_owner_review' } else { 'not_ready' })
        sourceCommit = $source
        testedSourceCommit = $testedSource
        mandatoryCases = 59
        satisfiedCases = $satisfied
        ownerGate = 'A060: separate owner authorization; publication not evaluated'
        releaseAuthorized = $false
        published = $false
        evidenceTrust = 'offline evaluation of reviewed metadata; no authentication of observations or publication authority'
        issues = @($issues.ToArray())
    }
}

if ($MyInvocation.InvocationName -eq '.') { return }
$ErrorActionPreference = 'Stop'
try {
    $repo = Split-Path $PSScriptRoot -Parent
    if (-not $ReviewPath) { $ReviewPath = Join-Path $repo 'docs/codex/RELEASE_ACCEPTANCE.json' }
    if (-not $CanonicalPath) { $CanonicalPath = Join-Path $repo 'docs/codex/ACCEPTANCE_CASES.json' }
    $reviewInput = [IO.File]::ReadAllText([IO.Path]::GetFullPath($ReviewPath)) | ConvertFrom-Json
    $canonicalInput = [IO.File]::ReadAllText([IO.Path]::GetFullPath($CanonicalPath)) | ConvertFrom-Json
    $result = Test-UpdReleaseReadiness -Review $reviewInput -Canonical $canonicalInput -RepositoryRoot $repo -ExpectedSourceCommit $ExpectedSourceCommit
    $result | ConvertTo-Json -Depth 10
    if ($result.assessment -eq 'ready_for_owner_review') { exit 0 }
    exit 2
}
catch {
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
    exit 1
}
