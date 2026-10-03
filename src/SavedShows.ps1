#requires -Version 5.1

function Get-PodcastSavedShowConfigPath {
    [CmdletBinding()]
    param([string]$ConfigPath)
    if (-not $ConfigPath) {
        $localRoot = $env:LOCALAPPDATA
        if (-not $localRoot) { throw 'Saved-show configuration path is unsafe or inaccessible.' }
        $ConfigPath = Join-Path $localRoot 'UniversalPodcastDownloader/saved-shows/shows.json'
    }
    $path = Resolve-PodcastSavedShowPath -Path $ConfigPath
    # Reserve the longer generated temp basename and persistent lock suffix
    # within the shared Windows/Framework path budget before creating anything.
    if ([IO.Path]::GetDirectoryName($path).Length -gt 210) { throw 'Saved-show configuration path is unsafe or inaccessible.' }
    return $path
}

function Resolve-PodcastSavedShowPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [switch]$Output)
    $message = if ($Output) { 'Saved-show output must be an absolute supported Windows path.' } else { 'Saved-show configuration path is unsafe or inaccessible.' }
    if ($Path.Length -gt 240 -or $Path -match '[\x00-\x1f"<>|?*]' -or $Path -match '^\\\\[?.]\\' -or
        $Path -notmatch '^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+(?:[\\/]|$))') { throw $message }
    $tail = if ($Path -match '^[A-Za-z]:') { $Path.Substring(2) } else { $Path }
    if ($tail.Contains(':') -or $tail -match '(?:^|[\\/])(?:\.{1,2}|[^\\/]*[ .])(?:[\\/]|$)') { throw $message }
    if ($tail -match '(?:^|[\\/])(?:CON|PRN|AUX|NUL|CONIN\$|CONOUT\$|COM[1-9\u00b9\u00b2\u00b3]|LPT[1-9\u00b9\u00b2\u00b3])(?:\.[^\\/]*|[ .]*)(?:[\\/]|$)') { throw $message }
    try { return [IO.Path]::GetFullPath($Path) }
    catch { if (Test-PodcastCancellation -ErrorObject $_) { throw }; throw $message }
}

function Assert-PodcastSavedShowPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $current = $Path
    while ($current) {
        try { $attributes = [IO.File]::GetAttributes($current) }
        catch [IO.FileNotFoundException] { $attributes = $null }
        catch [IO.DirectoryNotFoundException] { $attributes = $null }
        catch { if (Test-PodcastCancellation -ErrorObject $_) { throw }; throw 'Saved-show configuration path is unsafe or inaccessible.' }
        if ($null -ne $attributes -and ($attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Saved-show configuration path is unsafe or inaccessible.'
        }
        $parent = [IO.Path]::GetDirectoryName($current.TrimEnd('\', '/'))
        if (-not $parent -or $parent -eq $current) { break }
        $current = $parent
    }
}

function Get-PodcastSavedShowSecurity {
    [CmdletBinding()]
    param([switch]$Directory)
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Saved-show credentials are unavailable for the current Windows user.' }
    $sid = Get-PodcastSavedShowIdentity
    $security = if ($Directory) { [Security.AccessControl.DirectorySecurity]::new() } else { [Security.AccessControl.FileSecurity]::new() }
    $security.SetOwner($sid)
    $security.SetAccessRuleProtection($true, $false)
    $inheritance = if ($Directory) { [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit' } else { [Security.AccessControl.InheritanceFlags]::None }
    $rule = [Security.AccessControl.FileSystemAccessRule]::new($sid, [Security.AccessControl.FileSystemRights]::FullControl,
        $inheritance, [Security.AccessControl.PropagationFlags]::None, [Security.AccessControl.AccessControlType]::Allow)
    $security.AddAccessRule($rule)
    return $security
}

function Get-PodcastSavedShowIdentity {
    [CmdletBinding()]
    param()
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try { return $identity.User }
    finally { $identity.Dispose() }
}

function Assert-PodcastSavedShowAccess {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [switch]$Directory)
    Assert-PodcastSavedShowPath -Path $Path
    try {
        $info = if ($Directory) { [IO.DirectoryInfo]::new($Path) } else { [IO.FileInfo]::new($Path) }
        $security = if ('System.IO.FileSystemAclExtensions' -as [type]) { [IO.FileSystemAclExtensions]::GetAccessControl($info) } else { $info.GetAccessControl() }
        $sid = Get-PodcastSavedShowIdentity
        $rules = @($security.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))
        $expectedInheritance = if ($Directory) { [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit' } else { [Security.AccessControl.InheritanceFlags]::None }
        $valid = $security.AreAccessRulesProtected -and $security.GetOwner([Security.Principal.SecurityIdentifier]).Value -eq $sid.Value -and $rules.Count -eq 1
        if ($valid) {
            $rule = $rules[0]
            $valid = $rule.IdentityReference.Value -eq $sid.Value -and -not $rule.IsInherited -and
                $rule.AccessControlType -eq [Security.AccessControl.AccessControlType]::Allow -and
                $rule.FileSystemRights -eq [Security.AccessControl.FileSystemRights]::FullControl -and
                $rule.InheritanceFlags -eq $expectedInheritance -and $rule.PropagationFlags -eq [Security.AccessControl.PropagationFlags]::None
        }
    }
    catch { if (Test-PodcastCancellation -ErrorObject $_) { throw }; throw 'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.' }
    if (-not $valid) { throw 'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.' }
}

function Initialize-PodcastSavedShowDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    Assert-PodcastSavedShowPath -Path $Path
    if ([IO.Directory]::Exists($Path)) { Assert-PodcastSavedShowAccess -Path $Path -Directory; return }
    $missing = [Collections.Generic.Stack[string]]::new()
    $current = $Path
    while (-not [IO.Directory]::Exists($current)) {
        $missing.Push($current)
        $current = [IO.Path]::GetDirectoryName($current)
        if (-not $current) { throw 'Saved-show configuration path is unsafe or inaccessible.' }
    }
    $security = Get-PodcastSavedShowSecurity -Directory
    while ($missing.Count) {
        $next = $missing.Pop()
        Assert-PodcastSavedShowPath -Path $next
        try {
            if ('System.IO.FileSystemAclExtensions' -as [type]) { [IO.FileSystemAclExtensions]::Create([IO.DirectoryInfo]::new($next), $security) }
            else { $null = [IO.Directory]::CreateDirectory($next, $security) }
        }
        catch { if (Test-PodcastCancellation -ErrorObject $_) { throw }; throw 'Saved-show configuration path is unsafe or inaccessible.' }
        Assert-PodcastSavedShowAccess -Path $next -Directory
    }
}

function Open-PodcastSavedShowFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [IO.FileMode]$Mode = [IO.FileMode]::CreateNew)
    $security = Get-PodcastSavedShowSecurity
    if ('System.IO.FileSystemAclExtensions' -as [type]) {
        return [IO.FileSystemAclExtensions]::Create([IO.FileInfo]::new($Path), $Mode, [Security.AccessControl.FileSystemRights]::FullControl,
            [IO.FileShare]::None, 4096, [IO.FileOptions]::None, $security)
    }
    return [IO.FileStream]::new($Path, $Mode, [Security.AccessControl.FileSystemRights]::FullControl, [IO.FileShare]::None, 4096, [IO.FileOptions]::None, $security)
}

function Enter-PodcastSavedShowLock {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ConfigPath)
    $path = $ConfigPath + '.lock'
    Assert-PodcastSavedShowPath -Path $path
    if ([IO.File]::Exists($path)) { Assert-PodcastSavedShowAccess -Path $path }
    $stream = $null
    try {
        $stream = Open-PodcastSavedShowFile -Path $path -Mode OpenOrCreate
        Assert-PodcastSavedShowAccess -Path $path
        return $stream
    }
    catch {
        $failure = $_
        if ($stream) { try { $stream.Dispose() } catch { $null = $_ } }
        if (Test-PodcastCancellation -ErrorObject $failure) { throw $failure.Exception }
        $exception = $failure.Exception
        while ($exception) {
            if ($exception -is [IO.IOException] -and (($exception.HResult -band 0xffff) -in @(32, 33))) {
                throw 'Saved-show configuration is in use by another operation.'
            }
            $exception = $exception.InnerException
        }
        if ($failure.Exception.Message -eq 'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.') { throw $failure.Exception }
        throw 'Saved-show configuration path is unsafe or inaccessible.'
    }
}

function Assert-PodcastSavedShowFailure {
    [CmdletBinding()]
    param($ErrorObject, [Parameter(Mandatory)][string]$Fallback)
    if ($null -eq $ErrorObject) { return }
    if (Test-PodcastCancellation -ErrorObject $ErrorObject) { throw $ErrorObject.Exception }
    $safe = @(
        'Saved-show configuration is malformed or unsupported. Preserved the configuration.',
        'Saved-show configuration permissions are unsafe. Current-user-only protected access is required.',
        'Saved-show configuration path is unsafe or inaccessible.',
        'Saved-show credentials are unavailable for the current Windows user.',
        'Saved-show configuration is in use by another operation.',
        'The requested saved show was not found.',
        'Saved-show configuration supports at most 100 shows.',
        'Saved-show export already exists or its destination is unavailable.'
    )
    if ($safe -contains $ErrorObject.Exception.Message) { throw $ErrorObject.Exception.Message }
    throw $Fallback
}

function Assert-PodcastSavedShowJson {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Json)
    # A bounded strict lexer rejects duplicate keys, comments, escaped keys and
    # trailing commas before the two engines' JSON readers can disagree.
    $pattern = '[ \t\r\n]+|"(?:[^"\\\x00-\x1f]|\\(?:["\\/bfnrt]|u[0-9a-fA-F]{4}))*"|true|false|null|-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?|[{}\[\]:,]'
    $lexer = [regex]::new($pattern, [Text.RegularExpressions.RegexOptions]::CultureInvariant, [TimeSpan]::FromSeconds(2))
    $stack = [Collections.Generic.Stack[object]]::new()
    $offset = 0; $tokens = 0; $rootSeen = $false
    while ($offset -lt $Json.Length) {
        $token = $lexer.Match($Json, $offset)
        if (-not $token.Success -or $token.Index -ne $offset -or (++$tokens) -gt 8192) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
        $offset += $token.Length
        $value = $token.Value
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if (-not $stack.Count) {
            if ($rootSeen -or $value -ne '{') { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
            $rootSeen = $true
        }
        else {
            $frame = $stack.Peek()
            if ($value -eq '}' -or $value -eq ']') {
                $kind = if ($value -eq '}') { '{' } else { '[' }
                if ($frame.Kind -ne $kind -or $frame.State -notin @('first', 'end')) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
                $null = $stack.Pop(); continue
            }
            if ($frame.Kind -eq '{' -and $frame.State -in @('first', 'key')) {
                if ($value -cnotmatch '^"([a-z_][a-z0-9_]*)"$' -or -not $frame.Names.Add($Matches[1])) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
                $frame.State = 'colon'; continue
            }
            if ($frame.State -eq 'colon') {
                if ($value -ne ':') { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
                $frame.State = 'value'; continue
            }
            if ($frame.State -eq 'end') {
                if ($value -ne ',') { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
                $frame.State = if ($frame.Kind -eq '{') { 'key' } else { 'value' }; continue
            }
            if ($frame.State -notin @('value', 'first') -or $value -in @(':', ',')) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
            $frame.State = 'end'
        }
        if ($value -eq '{' -or $value -eq '[') {
            if ($stack.Count -ge 8) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
            $stack.Push([pscustomobject]@{ Kind = $value; State = 'first'; Names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) })
        }
    }
    if (-not $rootSeen -or $stack.Count) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
}

function Assert-PodcastSavedShowProperty {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Value, [Parameter(Mandatory)][string[]]$Names)
    if ($Value -isnot [pscustomobject] -or @($Value.PSObject.Properties).Count -ne $Names.Count) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
    foreach ($name in $Value.PSObject.Properties.Name) {
        if ($Names -cnotcontains $name) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
    }
}

function Assert-PodcastSavedShowName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Name)
    if ($Name -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$') { throw 'Saved-show name must be a token of 1 to 64 letters, digits, underscores or hyphens.' }
}

function Assert-PodcastSavedShowConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Config)
    Assert-PodcastSavedShowProperty -Value $Config -Names @('schema_version', 'shows')
    if (($Config.schema_version -isnot [int] -and $Config.schema_version -isnot [long]) -or $Config.schema_version -ne 1 -or
        $Config.shows -isnot [array] -or $Config.shows.Count -gt 100) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
    $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($show in $Config.shows) {
        Assert-PodcastSavedShowProperty -Value $show -Names @('name', 'feed_protected', 'output_path', 'mode', 'custom_count')
        if ($show.name -isnot [string] -or $show.name -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$' -or -not $names.Add($show.name) -or
            $show.feed_protected -isnot [string] -or $show.feed_protected.Length -lt 4 -or $show.feed_protected.Length -gt 196608 -or
            $show.feed_protected -cnotmatch '^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$' -or
            $show.output_path -isnot [string] -or $show.mode -isnot [string] -or @('Latest', 'All', 'Custom') -cnotcontains $show.mode) {
            throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
        }
        try { $null = Resolve-PodcastSavedShowPath -Path $show.output_path -Output }
        catch { if (Test-PodcastCancellation -ErrorObject $_) { throw }; throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
        if (($show.mode -eq 'Custom' -and (($show.custom_count -isnot [int] -and $show.custom_count -isnot [long]) -or $show.custom_count -lt 1 -or $show.custom_count -gt [int]::MaxValue)) -or
            ($show.mode -ne 'Custom' -and $null -ne $show.custom_count)) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
    }
}

function Read-PodcastSavedShowConfig {
    [CmdletBinding()]
    param([string]$ConfigPath)
    $path = Get-PodcastSavedShowConfigPath -ConfigPath $ConfigPath
    Assert-PodcastSavedShowPath -Path $path
    if (-not [IO.File]::Exists($path)) {
        if ([IO.Directory]::Exists($path)) { throw 'Saved-show configuration path is unsafe or inaccessible.' }
        return [pscustomobject]@{ schema_version = 1; shows = @() }
    }
    Assert-PodcastSavedShowAccess -Path ([IO.Path]::GetDirectoryName($path)) -Directory
    Assert-PodcastSavedShowAccess -Path $path
    $stream = $null; $reader = $null; $failure = $null; $config = $null
    try {
        $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        if ($stream.Length -lt 1 -or $stream.Length -gt 1048576) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
        $reader = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false, $true), $false)
        $json = $reader.ReadToEnd()
        Assert-PodcastSavedShowJson -Json $json
        $config = ConvertFrom-Json -InputObject $json -ErrorAction Stop
        Assert-PodcastSavedShowConfig -Config $config
    }
    catch { $failure = $_ }
    finally {
        try { if ($reader) { $reader.Dispose() } elseif ($stream) { $stream.Dispose() } }
        catch { if ($null -eq $failure) { $failure = $_ } }
    }
    Assert-PodcastSavedShowFailure -ErrorObject $failure -Fallback 'Saved-show configuration is malformed or unsupported. Preserved the configuration.'
    return $config
}

function Initialize-PodcastSavedShowProtection {
    [CmdletBinding()]
    param()
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Saved-show credentials are unavailable for the current Windows user.' }
    if (-not ('System.Security.Cryptography.ProtectedData' -as [type])) { Add-Type -AssemblyName System.Security -ErrorAction Stop }
}

function Protect-PodcastSavedShowFeed {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FeedUrl)
    $plain = $null
    try {
        Initialize-PodcastSavedShowProtection
        $plain = [Text.UTF8Encoding]::new($false, $true).GetBytes($FeedUrl)
        $entropy = [Text.Encoding]::UTF8.GetBytes('UniversalPodcastDownloader.SavedShows.Schema1')
        return [Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect($plain, $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser))
    }
    catch { if (Test-PodcastCancellation -ErrorObject $_) { throw }; throw 'Saved-show credentials are unavailable for the current Windows user.' }
    finally { if ($plain) { [Array]::Clear($plain, 0, $plain.Length) } }
}

function Unprotect-PodcastSavedShowFeed {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Protected)
    $plain = $null
    try {
        Initialize-PodcastSavedShowProtection
        $entropy = [Text.Encoding]::UTF8.GetBytes('UniversalPodcastDownloader.SavedShows.Schema1')
        $plain = [Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($Protected), $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
        $url = [Text.UTF8Encoding]::new($false, $true).GetString($plain)
        $null = Get-PodcastRequestUri -Uri $url
        return $url
    }
    catch { if (Test-PodcastCancellation -ErrorObject $_) { throw }; throw 'Saved-show credentials are unavailable for the current Windows user.' }
    finally { if ($plain) { [Array]::Clear($plain, 0, $plain.Length) } }
}

function Write-PodcastSavedShowConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ConfigPath, [Parameter(Mandatory)]$Config)
    Assert-PodcastSavedShowConfig -Config $Config
    $bytes = [Text.UTF8Encoding]::new($false, $true).GetBytes((ConvertTo-Json -InputObject $Config -Depth 5))
    if ($bytes.Length -gt 1048576) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
    $temporary = Join-Path ([IO.Path]::GetDirectoryName($ConfigPath)) ('shows-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $writer = $null; $ownedTemporary = $false; $failure = $null
    try {
        Assert-PodcastSavedShowAccess -Path ([IO.Path]::GetDirectoryName($ConfigPath)) -Directory
        $writer = Open-PodcastSavedShowFile -Path $temporary
        $ownedTemporary = $true
        Assert-PodcastSavedShowAccess -Path $temporary
        $writer.Write($bytes, 0, $bytes.Length); $writer.Flush($true); $writer.Dispose(); $writer = $null
        Assert-PodcastSavedShowPath -Path $ConfigPath
        if ([IO.File]::Exists($ConfigPath)) {
            Assert-PodcastSavedShowAccess -Path $ConfigPath
            [IO.File]::Replace($temporary, $ConfigPath, [NullString]::Value)
        }
        else { [IO.File]::Move($temporary, $ConfigPath) }
        $ownedTemporary = $false
        Assert-PodcastSavedShowAccess -Path $ConfigPath
    }
    catch { $failure = $_ }
    finally {
        try { if ($writer) { $writer.Dispose() } }
        catch { if ($null -eq $failure) { $failure = $_ } }
        try {
            if ($ownedTemporary -and [IO.File]::Exists($temporary)) {
                if ([IO.Path]::GetDirectoryName($temporary) -ne [IO.Path]::GetDirectoryName($ConfigPath)) { throw 'Saved-show configuration path is unsafe or inaccessible.' }
                Assert-PodcastSavedShowPath -Path $temporary
                Assert-PodcastSavedShowAccess -Path ([IO.Path]::GetDirectoryName($temporary)) -Directory
                [IO.File]::Delete($temporary)
            }
        }
        catch { if ($null -eq $failure) { $failure = $_ } }
    }
    Assert-PodcastSavedShowFailure -ErrorObject $failure -Fallback 'Saved-show configuration path is unsafe or inaccessible.'
}

function Get-PodcastSavedShow {
    [CmdletBinding()]
    param([string]$ConfigPath, [Parameter(Mandatory)][string]$Name)
    Assert-PodcastSavedShowName -Name $Name
    $config = Read-PodcastSavedShowConfig -ConfigPath $ConfigPath
    $show = @($config.shows | Where-Object name -eq $Name)
    if ($show.Count -ne 1) { throw 'The requested saved show was not found.' }
    return [pscustomobject]@{ Name = $show[0].name; FeedUrl = (Unprotect-PodcastSavedShowFeed -Protected $show[0].feed_protected);
        OutputPath = $show[0].output_path; Mode = $show[0].mode; CustomCount = $show[0].custom_count }
}

function Get-PodcastSavedShowList {
    [CmdletBinding()]
    param([string]$ConfigPath)
    $config = Read-PodcastSavedShowConfig -ConfigPath $ConfigPath
    $items = @($config.shows | ForEach-Object { [pscustomobject]@{ Name = $_.name; Mode = $_.mode; CustomCount = $_.custom_count; FeedConfigured = $true } })
    return ,$items
}

function Save-PodcastSavedShow {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param([string]$ConfigPath, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$FeedUrl,
        [Parameter(Mandatory)][string]$OutputPath, [string]$Mode = 'Latest', [Nullable[int]]$CustomCount)
    Assert-PodcastSavedShowName -Name $Name
    $path = Get-PodcastSavedShowConfigPath -ConfigPath $ConfigPath
    $output = Resolve-PodcastSavedShowPath -Path $OutputPath -Output
    $null = Get-PodcastRequestUri -Uri $FeedUrl
    if (@('Latest', 'All', 'Custom') -notcontains $Mode) { throw 'Saved-show configuration is malformed or unsupported. Preserved the configuration.' }
    $Mode = switch ($Mode) { 'Latest' { 'Latest' }; 'All' { 'All' }; 'Custom' { 'Custom' } }
    $hasCount = $PSBoundParameters.ContainsKey('CustomCount')
    if ($hasCount -and ($null -eq $CustomCount -or $CustomCount -lt 1)) { throw '-CustomCount must be a positive integer.' }
    if ($hasCount -and $PSBoundParameters.ContainsKey('Mode') -and $Mode -ne 'Custom') { throw '-CustomCount conflicts with an explicit Latest or All mode.' }
    if ($hasCount -and -not $PSBoundParameters.ContainsKey('Mode')) { $Mode = 'Custom' }
    if ($Mode -eq 'Custom' -and -not $hasCount) { throw "Mode 'Custom' requires -CustomCount with a value >= 1." }
    $initial = Read-PodcastSavedShowConfig -ConfigPath $path
    if (-not $PSCmdlet.ShouldProcess('Saved-show configuration', 'Store saved show')) { return [pscustomobject]@{ Changed = $false; Preview = $true; Name = $Name; Message = 'Saved-show configuration preview; no changes made.' } }
    $null = $initial
    Initialize-PodcastSavedShowDirectory -Path ([IO.Path]::GetDirectoryName($path))
    $lock = Enter-PodcastSavedShowLock -ConfigPath $path
    $failure = $null; $result = $null
    try {
        $config = Read-PodcastSavedShowConfig -ConfigPath $path
        $existing = @($config.shows | Where-Object name -eq $Name)
        if (-not $existing.Count -and $config.shows.Count -ge 100) { throw 'Saved-show configuration supports at most 100 shows.' }
        $show = [pscustomobject]@{ name = $Name; feed_protected = (Protect-PodcastSavedShowFeed -FeedUrl $FeedUrl); output_path = $output; mode = $Mode; custom_count = $CustomCount }
        if ($existing.Count) {
            $config.shows = @($config.shows | ForEach-Object { if ($_.name -eq $Name) { $show } else { $_ } })
        }
        else { $config.shows = @($config.shows) + @($show) }
        Write-PodcastSavedShowConfig -ConfigPath $path -Config $config
        $result = [pscustomobject]@{ Changed = $true; Preview = $false; Name = $Name; Message = 'Saved show stored.' }
    }
    catch { $failure = $_ }
    finally { try { $lock.Dispose() } catch { if ($null -eq $failure) { $failure = $_ } } }
    Assert-PodcastSavedShowFailure -ErrorObject $failure -Fallback 'Saved-show configuration path is unsafe or inaccessible.'
    return $result
}

function Remove-PodcastSavedShow {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param([string]$ConfigPath, [Parameter(Mandatory)][string]$Name)
    Assert-PodcastSavedShowName -Name $Name
    $path = Get-PodcastSavedShowConfigPath -ConfigPath $ConfigPath
    $config = Read-PodcastSavedShowConfig -ConfigPath $path
    if (-not @($config.shows | Where-Object name -eq $Name).Count) { throw 'The requested saved show was not found.' }
    if (-not $PSCmdlet.ShouldProcess('Saved-show configuration', 'Remove saved show')) { return [pscustomobject]@{ Changed = $false; Preview = $true; Name = $Name; Message = 'Saved-show configuration preview; no changes made.' } }
    $lock = Enter-PodcastSavedShowLock -ConfigPath $path
    $failure = $null; $result = $null
    try {
        $config = Read-PodcastSavedShowConfig -ConfigPath $path
        if (-not @($config.shows | Where-Object name -eq $Name).Count) { throw 'The requested saved show was not found.' }
        $config.shows = @($config.shows | Where-Object name -ne $Name)
        Write-PodcastSavedShowConfig -ConfigPath $path -Config $config
        $result = [pscustomobject]@{ Changed = $true; Preview = $false; Name = $Name; Message = 'Saved show removed.' }
    }
    catch { $failure = $_ }
    finally { try { $lock.Dispose() } catch { if ($null -eq $failure) { $failure = $_ } } }
    Assert-PodcastSavedShowFailure -ErrorObject $failure -Fallback 'Saved-show configuration path is unsafe or inaccessible.'
    return $result
}

function Export-PodcastSavedShowConfig {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param([string]$ConfigPath, [Parameter(Mandatory)][string]$Path)
    $target = Resolve-PodcastSavedShowPath -Path $Path
    Assert-PodcastSavedShowPath -Path $target
    $directory = [IO.Path]::GetDirectoryName($target)
    if (-not [IO.Directory]::Exists($directory) -or [IO.File]::Exists($target) -or [IO.Directory]::Exists($target)) { throw 'Saved-show export already exists or its destination is unavailable.' }
    $items = Get-PodcastSavedShowList -ConfigPath $ConfigPath
    if (-not $PSCmdlet.ShouldProcess('Saved-show metadata export', 'Create sanitized export')) { return [pscustomobject]@{ Changed = $false; Preview = $true; Name = $null; Message = 'Saved-show configuration preview; no changes made.' } }
    $bytes = [Text.UTF8Encoding]::new($false, $true).GetBytes((ConvertTo-Json -InputObject ([pscustomobject]@{ schema_version = 1; shows = $items }) -Depth 4))
    $writer = $null; $failure = $null; $result = $null
    try {
        Assert-PodcastSavedShowPath -Path $target
        $writer = Open-PodcastSavedShowFile -Path $target
        $writer.Write($bytes, 0, $bytes.Length); $writer.Flush($true)
        Assert-PodcastSavedShowAccess -Path $target
        $result = [pscustomobject]@{ Changed = $true; Preview = $false; Name = $null; Message = 'Sanitized saved-show metadata exported.' }
    }
    catch { $failure = $_ }
    finally { try { if ($writer) { $writer.Dispose() } } catch { if ($null -eq $failure) { $failure = $_ } } }
    Assert-PodcastSavedShowFailure -ErrorObject $failure -Fallback 'Saved-show export already exists or its destination is unavailable.'
    return $result
}
