#requires -Version 5.1

function ConvertFrom-PodcastFeedXml {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Content)

    # Check the string even when a caller bypasses the bounded HTTP adapter.
    if ($Content.Length -gt 8388608) { throw 'Feed XML is invalid or exceeds safe parser limits.' }
    $settings = [Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.MaxCharactersInDocument = 8388608
    $settings.MaxCharactersFromEntities = 1024
    $settings.ValidationType = [Xml.ValidationType]::None
    $settings.CloseInput = $true
    $reader = $null
    $textReader = $null
    try {
        # Preflight depth and node/attribute counts before allocating a DOM.
        # Both passes use the same DTD, resolver and character restrictions.
        $textReader = [IO.StringReader]::new($Content)
        $reader = [Xml.XmlReader]::Create($textReader, $settings)
        $nodes = 0
        while ($reader.Read()) {
            $nodes += 1 + $reader.AttributeCount
            if ($reader.Depth -gt 64 -or $nodes -gt 100000) { throw 'XML resource limit.' }
        }
        $reader.Dispose()
        $reader = $null
        $textReader.Dispose()
        $textReader = [IO.StringReader]::new($Content)
        $reader = [Xml.XmlReader]::Create($textReader, $settings)
        $document = [Xml.XmlDocument]::new()
        $document.XmlResolver = $null
        $document.Load($reader)
        return ,$document
    }
    catch {
        # Parser errors can contain private feed text or external identifiers.
        throw 'Feed XML is invalid or exceeds safe parser limits.'
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $textReader) { $textReader.Dispose() }
    }
}
