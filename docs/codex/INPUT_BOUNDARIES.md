# Input and preview boundaries

UPD-0107 applies the following policy to the existing PowerShell downloader. The limits are the project's compatibility and resource choices. They do not establish that the previous parser defaults were exploitable.

## Preview

An ordinary `-WhatIf` run resolves the feed, parses it, validates episode targets and reads local history to build a plan. It may also read and hash existing files for legacy review. Legacy `Preview` is read-only; `Adopt`, `Redownload` and `Rollback` with `-WhatIf` also validate their explicit selections before returning a plan.

| Operation | Preview behavior |
|---|---|
| Retrieve the selected feed/show page and permitted redirects | Allowed under the metadata limits below |
| Read history, checkpoints and local media for review | Allowed |
| Follow planned enclosure URLs or start the media-transfer path | Not performed |
| Create/change output directories, media, history, checkpoints or configuration | Not performed |
| Create lock files, startup/show logs or diagnostic exports | Not performed |
| Change keep-awake or power settings | Not performed; opt-in KeepAwake is suppressed during preview |
| Report a completed download | Not performed |

Preview can contact metadata servers. The supplied URL or an allowed redirect can return unexpected content, including audio; the bounded metadata reader may read those bytes before rejecting the response as a feed. The preview guarantee covers avoiding planned enclosure requests, media-transfer execution and persistent downloader writes. It cannot determine a server's content before requesting it.

Preview does not promise that a URL is reachable later, that media will transfer successfully, or that local state will remain unchanged by another process. Normal execution and changing legacy actions pass their ShouldProcess decision before archive writes, then recheck local state under writer protection. Explicit confirmation and diagnostics retain the behavior in [DIAGNOSTICS.md](DIAGNOSTICS.md).

## Request targets and redirects

`src/NetworkPolicy.ps1` supplies the shared policy for pages, feeds and media:

- Targets must be well-formed absolute HTTP or HTTPS URLs with a host and at most 32,768 characters. Backslashes, ASCII control characters and unescaped ASCII spaces are rejected. A URL's user-information component, such as `user:password@host`, is rejected. File, FTP, data and other schemes are unsupported.
- Private-network addresses and loopback hosts are allowed. This is an owner-operated local downloader, so the policy does not restrict requests to public internet destinations.
- Initial HTTP is allowed. Responses with status 301, 302, 303, 307 or 308 can redirect. Each target is validated before another request; HTTPS-to-HTTP downgrade is rejected. At most five redirects are followed for one retrieval. Missing, invalid, unsupported or excessive redirects fail the retrieval.
- Automatic redirect following, default request credentials and cookie handling are disabled. Requests do not reuse Windows login credentials or response cookies. There is no browser-session import or custom authentication/header feature.
- TLS certificate verification and system proxy defaults remain in use. The downloader does not install a certificate exception or alter machine-wide TLS/proxy settings.
- Signed URL path/query values remain available for the actual request. Diagnostic display strings are separate and contain only a hostname and opaque request ID.

Cross-origin redirects are allowed when the target satisfies this policy. Each request is constructed from the accepted target without copying credentials or cookies from the previous response. The policy does not infer that an unfamiliar host belongs to a publisher.

Redirects do not silently alter archive identity: the original requested feed URL remains the identity input. Relative feed links discovered in HTML are resolved against that page's final effective URL and validated before use. All parsed episode enclosure targets are validated during planning, including when `-WhatIf` prevents enclosure transfers.

## Metadata and parser limits

Feed and HTML retrieval is separate from large streamed media. `Invoke-PodcastMetadataRequest` bounds the metadata body before handing text to the parser. `ConvertFrom-PodcastFeedXml` also bounds direct string input, so callers cannot bypass XML limits by skipping the network helper.

Metadata must be a complete HTTP 200 response without Content-Range. The reader rejects an excessive declared Content-Length before buffering, enforces the byte limit while streaming even without a declared length, and checks received bytes against Content-Length when present. Requests ask for `Accept-Encoding: identity`; other content encodings, including gzip, are refused rather than decompressed.

Text decoding uses a BOM first, then an HTTP charset, then supported XML encoding detection/declaration, and otherwise UTF-8. Supported encodings are UTF-8, UTF-16 and UTF-32 in either byte order, ASCII, ISO-8859-1 and Windows-1252. Invalid byte sequences and unsupported encodings fail. The XML declaration scan examines at most the first 512 bytes. The internal `MaximumBytes` parameter can lower the body cap but cannot raise it; it is not a new main-script option.

| Resource | Limit or policy |
|---|---|
| Metadata response body | 8 MiB (8,388,608 bytes) |
| Direct HTML discovery input | 8,388,608 characters |
| XML input string / `MaxCharactersInDocument` | 8,388,608 characters |
| `MaxCharactersFromEntities` | Set to 1,024; DTDs remain prohibited, and built-in character references use the document limit |
| XML reader depth | At most 64, with the root element at depth 0 |
| XML reader nodes plus attributes | At most 100,000 |
| HTML discovery regex | 250 ms per matching operation |
| DTD handling | `DtdProcessing.Prohibit` |
| External XML resolution | Null resolver on both the reader and document |

The XML helper first reads the complete document with a streaming reader to check depth and node/attribute counts. It then loads the DOM through a second reader with the same settings. DTDs are rejected on encounter; external entity, DTD and schema references are not fetched. Errors use a safe local message rather than including private XML text.

The metadata cap does not limit episode audio to 8 MiB. Media still streams into an owned temporary file and must pass the existing completion/framing/signature checks before final placement. Parser limits bound accepted input size and structure; they are not a promise of a fixed total run time or a full media decoder. Implemented policies include [bounded retries and header/idle timeouts](TRANSPORT_POLICY.md), [validator-aware resume](RESUME_POLICY.md), [bounded pagination](FEED_PAGINATION.md), and deterministic supported-audio selection documented in the README.

## Compatibility and evidence

UPD-0203 adds [shared feed discovery](FEED_DISCOVERY.md) for both entry paths. The scanner keeps the 8 MiB character limit and 250 ms regex timeouts, adds a 100,000 token limit, ignores inert/raw text markup, validates decoded discovery/base targets, and stops at one selected feed. Final page/base URIs do not replace original direct-feed identities. Numbered guided choices and explicit CLI ambiguity errors occur before archive locks/writes; already fetched metadata is reused.

The request and XML helpers use built-in .NET APIs available to Windows PowerShell 5.1 and PowerShell 7 on Windows. They add no runtime package, external binary or persistent configuration. Imports define helpers without starting requests or writing files.

Primary API references and their check date are in [SOURCES.md](SOURCES.md), especially S7 and S11-S14. Task evidence records actual mocked/loopback tests, engine versions and limitations separately from these policy statements. No real archive or private subscription is required to verify the boundaries.
