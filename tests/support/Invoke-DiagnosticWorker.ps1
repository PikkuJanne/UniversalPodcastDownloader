param([Parameter(Mandatory)][string]$ConfigPath)

# This child owns only the paths in its marked integration context. Isolating
# both environment locations prevents startup diagnostics reaching user data.
$ErrorActionPreference = 'Stop'
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$env:LOCALAPPDATA = Join-Path $config.Root 'local'
$env:TEMP = Join-Path $config.Root 'temp'
$env:TMP = $env:TEMP
$result = [ordered]@{
    Succeeded = $false
    ErrorMessage = $null
    ErrorType = $null
    OriginalExceptionPreserved = $false
    Output = @()
    RunId = $null
}

if ($config.BlockPrimary -or $config.BlockBoth) {
    $null = [IO.Directory]::CreateDirectory($config.Root)
    [IO.File]::WriteAllText($env:LOCALAPPDATA, 'Owned synthetic obstruction.')
}
if ($config.BlockBoth) { [IO.File]::WriteAllText($env:TEMP, 'Owned synthetic obstruction.') }

$parameters = @{
    FeedUrl = $config.FeedUrl
    OutputPath = Join-Path $config.Root 'output'
    Mode = 'All'
    Verbose = $true
}
if ($config.Action -eq 'InvalidOutput') {
    $null = [IO.Directory]::CreateDirectory($config.Root)
    [IO.File]::WriteAllText($parameters.OutputPath, 'Preserve this synthetic existing file.')
}
if ($config.Action -eq 'Preview') { $parameters.WhatIf = $true }
if ($config.Action -eq 'LegacyPreview') {
    $parameters.LegacyPath = Join-Path $parameters.OutputPath 'Synthetic legacy show'
    $null = [IO.Directory]::CreateDirectory($parameters.LegacyPath)
    $parameters.LegacyAction = 'Preview'
}

# Allow the engine to format the actual entrypoint failure. No test catch or
# result serializer can hide accidental raw URL/exception output on this path.
if ($config.Action -eq 'RawFailure') {
    & $config.ProductScript @parameters
    exit 0
}

try {
    if ($config.Action -in @('ApiWrite', 'AppendFailure')) {
        . $config.ProductScript
        Initialize-PodcastDiagnostics -Roots @($config.LogRoot)
        $result.RunId = $script:PodcastDiagnostics.RunId
        $unicode = 'Synthetic Unicode ' + [char]0x00e4 + [char]0x65e5 + [char]::ConvertFromUtf32(0x1f3a7)
        Write-PodcastDiagnostic -Message $unicode
        if ($config.Action -eq 'AppendFailure') {
            # A disposed writer is a deterministic disk-writer failure seam.
            # Keep the primary exception object alive while diagnostics fail.
            $script:PodcastDiagnostics.Writer.Dispose()
            $primary = [InvalidOperationException]::new('Original synthetic operation failure.')
            try { throw $primary }
            catch {
                $original = $_.Exception
                Write-PodcastDiagnostic -Message 'Secondary diagnostic write failed.' -Level ERROR
                $result.OriginalExceptionPreserved = [object]::ReferenceEquals($original, $primary)
                throw
            }
        }
        Close-PodcastDiagnostics
    }
    else { $result.Output = @(& $config.ProductScript @parameters) }
    $result.Succeeded = $true
}
catch {
    $result.ErrorMessage = $_.Exception.Message
    $result.ErrorType = $_.Exception.GetType().FullName
}
finally {
    if (Get-Command Close-PodcastDiagnostics -ErrorAction SilentlyContinue) { Close-PodcastDiagnostics }
}
[IO.File]::WriteAllText($config.ResultPath, ($result | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
if ($result.Succeeded) { exit 0 }
exit 1
