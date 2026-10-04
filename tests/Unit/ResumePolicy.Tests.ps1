BeforeAll {
    $script:RepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:RepositoryRoot 'src/NetworkPolicy.ps1')
    . (Join-Path $script:RepositoryRoot 'src/MediaRequest.ps1')
    Add-Type -AssemblyName System.Net.Http
    Mock Get-PodcastHttpClient { return $script:Client }

    function Get-ResumeTestResponse {
        param([int]$Status = 206, [string]$Range = 'bytes 2-3/4', [string]$ETag = '"entity-1"',
            [byte[]]$Body = @(3, 4), [string]$ContentType = 'audio/mpeg', [string]$Encoding, $Length = 2)
        $response = [Net.Http.HttpResponseMessage]::new([Enum]::ToObject([Net.HttpStatusCode], $Status))
        $response.Content = [Net.Http.ByteArrayContent]::new($Body)
        if ($null -ne $Length) { $response.Content.Headers.ContentLength = $Length }
        if ($Range) { $null = $response.Content.Headers.TryAddWithoutValidation('Content-Range', $Range) }
        if ($ETag) { $null = $response.Headers.TryAddWithoutValidation('ETag', $ETag) }
        if ($ContentType) { $null = $response.Content.Headers.TryAddWithoutValidation('Content-Type', $ContentType) }
        if ($Encoding) { $null = $response.Content.Headers.TryAddWithoutValidation('Content-Encoding', $Encoding) }
        return $response
    }

    function Get-ResumeTestClient {
        param([object[]]$Responses)
        $client = [pscustomobject]@{ Responses = $Responses; Requests = [Collections.Generic.List[object]]::new(); Disposed = $false }
        $client | Add-Member ScriptMethod SendAsync {
            param($request, $completion, $token)
            $null = $completion; $null = $token
            $response = $this.Responses[[Math]::Min($this.Requests.Count, $this.Responses.Count - 1)]
            $this.Requests.Add($request)
            $task = [Threading.Tasks.TaskCompletionSource[object]]::new()
            $task.SetResult($response)
            return $task.Task
        }
        $client | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        return $client
    }
}

Describe 'A027: validated media resume' -Tag 'Unit', 'A027' {
    BeforeEach {
        $script:Uri = [Uri]'https://media.invalid/episode.mp3'
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $fingerprint = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($script:Uri.AbsoluteUri)))).Replace('-', '').ToLowerInvariant() }
        finally { $sha.Dispose() }
        $script:Resume = [pscustomobject]@{ Offset = [long]2; TotalLength = [long]4; ETag = '"entity-1"'; ContentType = 'audio/mpeg'; FinalUriFingerprint = $fingerprint }
        $script:Responses = @((Get-ResumeTestResponse))
        $script:Client = Get-ResumeTestClient -Responses $script:Responses
        $script:Destination = [IO.MemoryStream]::new()
        $script:Destination.Write([byte[]]@(1, 2), 0, 2)
        $script:Infos = [Collections.Generic.List[object]]::new()
        $script:Progress = [Collections.Generic.List[long]]::new()
    }
    AfterEach {
        $script:Destination.Dispose()
        foreach ($response in $script:Responses) { $response.Dispose() }
    }

    It 'appends only an exactly validated complete tail and reports total bytes' {
        $result = Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume -OnResponse {
            param($Info)
            $script:Destination.Length | Should -Be 2
            $script:Infos.Add($Info)
        } -OnProgress { param($Bytes) $script:Progress.Add($Bytes) }
        $result.Completed | Should -BeTrue
        $result.Bytes | Should -Be 4
        $result.ContentLength | Should -Be 4
        $script:Destination.ToArray() | Should -Be @(1, 2, 3, 4)
        $script:Infos.Count | Should -Be 1
        $script:Infos[0].Action | Should -Be 'Append'
        $script:Infos[0].ResumeSupported | Should -BeTrue
        $script:Progress.ToArray() | Should -Be @(4)
        [string]$script:Client.Requests[0].Headers.Range | Should -BeExactly 'bytes=2-'
        [string]$script:Client.Requests[0].Headers.IfRange | Should -BeExactly '"entity-1"'
    }

    It 'requires restart without touching partial bytes for <Case>' -ForEach @(
        @{ Case = 'ignored range'; Response = @{ Status = 200; Range = '' } },
        @{ Case = '416 matching total'; Response = @{ Status = 416; Range = 'bytes */4' } },
        @{ Case = '416 missing total'; Response = @{ Status = 416; Range = '' } },
        @{ Case = 'changed ETag'; Response = @{ ETag = '"entity-2"' } },
        @{ Case = 'case changed ETag'; Response = @{ ETag = '"Entity-1"' } },
        @{ Case = 'weak ETag'; Response = @{ ETag = 'W/"entity-1"' } },
        @{ Case = 'absent ETag'; Response = @{ ETag = '' } },
        @{ Case = 'wrong offset'; Response = @{ Range = 'bytes 1-3/4'; Length = 3 } },
        @{ Case = 'wrong total'; Response = @{ Range = 'bytes 2-4/5'; Length = 3 } },
        @{ Case = 'shorter tail'; Response = @{ Range = 'bytes 2-2/4'; Length = 1 } },
        @{ Case = 'unknown total'; Response = @{ Range = 'bytes 2-3/*' } },
        @{ Case = 'invalid interval'; Response = @{ Range = 'bytes 3-2/4' } },
        @{ Case = 'malformed range'; Response = @{ Range = 'secret-token-not-a-range' } },
        @{ Case = 'missing range'; Response = @{ Range = '' } },
        @{ Case = 'different unit'; Response = @{ Range = 'items 2-3/4' } },
        @{ Case = 'multiple ranges'; Response = @{ Range = 'bytes 2-2/4,3-3/4' } },
        @{ Case = 'wrong content length'; Response = @{ Length = 3 } },
        @{ Case = 'changed type'; Response = @{ ContentType = 'audio/ogg' } },
        @{ Case = 'missing type'; Response = @{ ContentType = '' } },
        @{ Case = 'multipart'; Response = @{ ContentType = 'multipart/byteranges; boundary=secret-token' } },
        @{ Case = 'compressed body'; Response = @{ Encoding = 'gzip' } }
    ) {
        $script:Responses[0].Dispose()
        $script:Responses = @((Get-ResumeTestResponse @Response))
        $script:Client = Get-ResumeTestClient -Responses $script:Responses
        $result = Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume -OnResponse { throw 'Callback must not run.' } -OnProgress { throw 'Callback must not run.' }
        $result.Completed | Should -BeFalse
        $result.RestartRequired | Should -BeTrue
        $result.RestartReason | Should -Not -Match 'secret-token|entity-[12]'
        $script:Destination.ToArray() | Should -Be @(1, 2)
    }

    It 'rejects a partial stream with a mismatched offset before opening a client' {
        $script:Destination.Position = 0
        { Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume } | Should -Throw '*does not match*'
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
        $script:Destination.ToArray() | Should -Be @(1, 2)
    }

    It 'rejects repeated singleton response headers before appending <Header>' -ForEach @(
        @{ Header = 'ETag'; Value = '"entity-1"' }, @{ Header = 'Content-Range'; Value = 'bytes 2-3/4' }
    ) {
        if ($Header -eq 'ETag') { $null = $script:Responses[0].Headers.TryAddWithoutValidation($Header, $Value) }
        else { $null = $script:Responses[0].Content.Headers.TryAddWithoutValidation($Header, $Value) }
        $result = Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume
        $result.RestartRequired | Should -BeTrue
        $script:Destination.ToArray() | Should -Be @(1, 2)
    }

    It 'rejects unsafe stored validators before opening a client for <Tag>' -ForEach @(
        @{ Tag = 'W/"entity-1"' }, @{ Tag = 'unquoted' }, @{ Tag = '*' }, @{ Tag = '"hello world"' }, @{ Tag = "`"secret`nheader`"" }, @{ Tag = '' }
    ) {
        $script:Resume.ETag = $Tag
        { Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume } | Should -Throw '*does not match*'
        Should -Invoke Get-PodcastHttpClient -Times 0 -Exactly
        $script:Destination.ToArray() | Should -Be @(1, 2)
    }

    It 'preserves checkpoint failures as local errors before any append' {
        $caught = $null
        try {
            Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume -OnResponse {
                throw [IO.IOException]::new('Synthetic local checkpoint failure.')
            }
        }
        catch { $caught = $_ }
        (Get-PodcastTransportFailure -ErrorObject $caught) | Should -BeNullOrEmpty
        $caught.Exception.Message | Should -Be 'Synthetic local checkpoint failure.'
        $script:Destination.ToArray() | Should -Be @(1, 2)
    }

    It 'does not checkpoint a chunk whose destination write failed' {
        $script:Destination.Dispose()
        $script:Destination = [IO.MemoryStream]::new([byte[]]@(1, 2), $true)
        $script:Destination.Position = 2
        { Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume -OnProgress { param($Bytes) $script:Progress.Add($Bytes) } } | Should -Throw
        $script:Progress.Count | Should -Be 0
        $script:Destination.ToArray() | Should -Be @(1, 2)
    }

    It 'never writes a response chunk beyond the validated range' {
        $script:Responses[0].Dispose()
        $script:Responses = @((Get-ResumeTestResponse -Body @(3, 4, 5)))
        $script:Client = Get-ResumeTestClient -Responses $script:Responses
        { Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume } | Should -Throw '*exceeds*range*'
        $script:Destination.ToArray() | Should -Be @(1, 2)
    }

    It 'reports a short received tail as transient without claiming completion' {
        $script:Responses[0].Dispose()
        $script:Responses = @((Get-ResumeTestResponse -Body @(3)))
        $script:Client = Get-ResumeTestClient -Responses $script:Responses
        $caught = $null
        try { Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume }
        catch { $caught = Get-PodcastTransportFailure -ErrorObject $_ }
        $caught.Data['Kind'] | Should -Be 'IncompleteBody'
        $caught.Data['Retryable'] | Should -BeTrue
        $script:Destination.ToArray() | Should -Be @(1, 2, 3)
    }

    It 'sends Range and If-Range only to the stored final target after a redirect' {
        $redirect = Get-ResumeTestResponse -Status 302 -Range ''
        $redirect.Headers.Location = $script:Uri
        $script:Responses = @($redirect, $script:Responses[0])
        $script:Client = Get-ResumeTestClient -Responses $script:Responses
        $result = Invoke-PodcastMediaRequest -Uri 'https://other.invalid/start' -DestinationStream $script:Destination -Resume $script:Resume
        $result.Completed | Should -BeTrue
        $script:Client.Requests[0].Headers.Range | Should -BeNullOrEmpty
        $script:Client.Requests[0].Headers.IfRange | Should -BeNullOrEmpty
        [string]$script:Client.Requests[1].Headers.Range | Should -Be 'bytes=2-'
        [string]$script:Client.Requests[1].Headers.IfRange | Should -Be '"entity-1"'
    }

    It 'requires fresh transfer for a changed final target even with matching ETag' {
        $redirect = Get-ResumeTestResponse -Status 302 -Range ''
        $redirect.Headers.Location = [Uri]'https://other.invalid/changed'
        $script:Responses = @($redirect, $script:Responses[0])
        $script:Client = Get-ResumeTestClient -Responses $script:Responses
        $result = Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -Resume $script:Resume
        $result.RestartRequired | Should -BeTrue
        $result.RestartReason | Should -Be 'response-target-changed'
        $script:Client.Requests[1].Headers.Range | Should -BeNullOrEmpty
        $script:Client.Requests[1].Headers.IfRange | Should -BeNullOrEmpty
        $script:Destination.ToArray() | Should -Be @(1, 2)
    }

    It 'offers durable resume metadata before the first fresh write' {
        $script:Destination.SetLength(0)
        $script:Destination.Position = 0
        $script:Responses[0].Dispose()
        $script:Responses = @((Get-ResumeTestResponse -Status 200 -Range '' -ContentType 'Audio/MPEG; parameter=value'))
        $script:Client = Get-ResumeTestClient -Responses $script:Responses
        $result = Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -OnResponse {
            param($Info)
            $script:Destination.Length | Should -Be 0
            $script:Infos.Add($Info)
        }
        $result.Completed | Should -BeTrue
        $script:Infos[0].Action | Should -Be 'Fresh'
        $script:Infos[0].ResumeSupported | Should -BeTrue
        $script:Infos[0].ETag | Should -BeExactly '"entity-1"'
        $script:Infos[0].ContentType | Should -BeExactly 'audio/mpeg'
        $script:Infos[0].TotalLength | Should -Be 2
        $script:Infos[0].FinalUriFingerprint | Should -Be $script:Resume.FinalUriFingerprint
    }

    It 'never offers Last-Modified or a weak/absent ETag as resume evidence <Tag>' -ForEach @(
        @{ Tag = 'W/"entity-1"' }, @{ Tag = '' }
    ) {
        $script:Destination.SetLength(0)
        $script:Destination.Position = 0
        $script:Responses[0].Dispose()
        $response = Get-ResumeTestResponse -Status 200 -Range '' -ETag $Tag
        $response.Content.Headers.LastModified = [DateTimeOffset]'2026-10-02T10:00:00Z'
        $script:Responses = @($response)
        $script:Client = Get-ResumeTestClient -Responses $script:Responses
        $null = Invoke-PodcastMediaRequest -Uri $script:Uri.AbsoluteUri -DestinationStream $script:Destination -OnResponse { param($Info) $script:Infos.Add($Info) }
        $script:Infos[0].ResumeSupported | Should -BeFalse
    }
}
