BeforeAll {
    $script:DownloaderPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'UniversalPodcastDownloader.ps1'
    . $script:DownloaderPath
}

Describe 'A024: explicit bounded feed XML parsing' -Tag 'Unit', 'A024' {
    It 'rejects a DTD instead of expanding its internal entity into an episode' {
        Mock Invoke-PodcastWebRequest {
            [pscustomobject]@{ Content = '<!DOCTYPE rss [<!ENTITY title "Expanded title">]><rss><channel><item><title>&title;</title></item></channel></rss>' }
        }
        { Resolve-PodcastItems -Feeds 'https://feed.example.invalid/rss' } | Should -Throw 'Source XML is invalid or exceeds safe parser limits.'
    }

    It 'rejects external and internal DTDs with a fixed error' -TestCases @(
        @{ Xml = '<!DOCTYPE rss SYSTEM "https://entity.example.invalid/private"><rss />' }
        @{ Xml = '<!DOCTYPE rss [<!ENTITY x SYSTEM "file:///C:/private.txt">]><rss>&x;</rss>' }
        @{ Xml = '<!DOCTYPE rss [<!ENTITY x "secret">]><rss>&x;</rss>' }
    ) {
        param($Xml)
        $inputXml = $Xml
        { ConvertFrom-PodcastFeedXml -Content $inputXml } | Should -Throw 'Feed XML is invalid or exceeds safe parser limits.'
    }

    It 'rejects excess nesting before building the document' {
        { ConvertFrom-PodcastFeedXml -Content (('<a>' * 66) + ('</a>' * 66)) } | Should -Throw '*safe parser limits*'
    }

    It 'rejects excess node and attribute count before building the document' {
        { ConvertFrom-PodcastFeedXml -Content ('<rss>' + ('<item x="1" />' * 50001) + '</rss>') } | Should -Throw '*safe parser limits*'
    }

    It 'rejects an oversized string even without the HTTP adapter' {
        { ConvertFrom-PodcastFeedXml -Content ('<rss>' + ('x' * 8388608) + '</rss>') } | Should -Throw '*safe parser limits*'
    }

    It 'parses valid XML with namespaces and predefined character references' {
        $xml = ConvertFrom-PodcastFeedXml -Content '<feed xmlns="http://www.w3.org/2005/Atom"><entry><title>A &amp; B &#x1F600;</title></entry></feed>'
        $xml.DocumentElement.LocalName | Should -Be 'feed'
        $xml.DocumentElement.FirstChild.InnerText | Should -Be ('A & B ' + [char]::ConvertFromUtf32(0x1f600))
    }

    It 'does not interpret XInclude or schema locations as fetch instructions' {
        $xml = ConvertFrom-PodcastFeedXml -Content '<rss xmlns:xi="http://www.w3.org/2001/XInclude" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="https://entity.example.invalid/schema"><xi:include href="https://entity.example.invalid/private" /></rss>'
        $xml.DocumentElement.FirstChild.LocalName | Should -Be 'include'
    }

    It 'omits private content from malformed XML errors' {
        { ConvertFrom-PodcastFeedXml -Content '<privateCanary><mismatch></privateCanary>' } | Should -Throw 'Feed XML is invalid or exceeds safe parser limits.'
    }
}
