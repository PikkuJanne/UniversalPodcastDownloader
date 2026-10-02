# Synthetic test helper kit

This is development tooling, not the downloader and not a product acceptance suite. Python 3.9+ standard library is sufficient; it is never a runtime dependency of the PowerShell application. The MP3 contains generated silence. No real podcast content or credentials are included.

From this directory, run helper self-tests:

```powershell
python -B -m unittest discover -s selftests -v
```

Start a loopback server with an ephemeral port:

```powershell
python -B .\fixture_server.py --port 0
```

It prints JSON containing its actual `base_url`. A test harness may pass `--ready-file` pointing to a **new file in its own temporary directory**; the server refuses to overwrite an existing ready file. Stop only the process your test started. The server binds only `127.0.0.1`, never an external interface. Do not expose it through a proxy/tunnel. No web fetching, package installation or administrator listener configuration is needed.

Feed placeholders `{{BASE_URL}}` and `{{AUDIO_BYTES}}` are replaced when served. Standalone parsing tests can replace them in memory with a loopback test URL. `.invalid` entries in the media metadata fixture intentionally have no downloadable bytes: mock them for unit tests and never enable arbitrary external egress.

## Routes

| Route | Behavior |
|---|---|
| `/feeds/single.xml`, `/feeds/empty.xml`, `/feeds/collisions.xml` | Named RSS fixtures |
| `/feeds/atom.xml`, `/feeds/dates.xml`, `/feeds/media.xml` | Atom/date/enclosure fixtures |
| `/feeds/page-1.xml`, `/feeds/page-2.xml` | Duplicate across pages, with a next-link cycle |
| `/feeds/dtd.xml`, `/feeds/malformed.xml` | Rejected XML cases |
| `/feeds/external-http.xml`, `/feeds/external-file.xml?file=<encoded file URI>` | DTD entity declarations; HTTP references `/entity/never`, file URI is reflected into XML only and never read by the server |
| `/feeds/internal-dtd.xml`, `/entity/never` | Internal entity expansion fixture; countable external-entity sentinel response |
| `/feeds/oversized.xml`, `/feeds/oversized-no-length.xml` | Finite valid XML just above 8 MiB, with Content-Length or connection-close framing |
| `/feeds/deep.xml`, `/feeds/many-nodes.xml` | Valid XML exceeding the proposed depth 64 / node 100,000 parser budgets |
| `/feeds/redirect-media.xml`, `/feeds/redirect-file-media.xml` | Feeds pointing to a valid or forbidden media redirect |
| `/show`, `/show/not-feed` | Multiple candidates/base/relative links; false feed substring |
| `/redirect/show` | Redirect to the show page |
| `/redirect/loop` | Same-URL redirect loop |
| `/redirect/feed`, `/redirect/file`, `/redirect/userinfo` | Relative valid feed target, forbidden file scheme, or synthetic embedded credentials |
| `/redirect/cookie`, `/feeds/no-cookie.xml` | Set a synthetic cookie then redirect from IPv4 to localhost; records only a `/credential-received` counter if cookie/authorization headers reach the destination |
| `/media/redirect.mp3`, `/media/redirect-file.mp3` | Relative media redirect or forbidden file target |
| `/media/ok.mp3` | Silent MP3; strong ETag; correct single byte-range handling |
| `/media/ignore-range.mp3` | Ignores Range and returns full body |
| `/media/changed.mp3` | Changed entity and different strong ETag; stale If-Range produces full body |
| `/media/weak.mp3`, `/media/no-validator.mp3` | Weak/absent validation evidence |
| `/media/bad-range.mp3` | Intentionally inconsistent Content-Range |
| `/media/always-416.mp3` | Always 416, including a Content-Range total |
| `/media/truncated.mp3` | Disconnects after one third of declared body; deliberately omits a validator to retain failed-attempt cleanup coverage |
| `/media/empty.mp3`, `/media/html.mp3` | Empty or HTML masquerading as audio |
| `/media/octet-stream` | Valid MP3 with generic Content-Type, no filename extension |
| `/media/no-length.mp3` | Valid close-delimited media without Content-Length |
| `/media/stall.mp3` | Body starts then pauses for 0.6 seconds |
| `/media/stall-headers.mp3` | Response headers delayed 0.6 seconds |
| `/retry/once.mp3` | First request 503 with Retry-After: 1, then successful media |
| `/status/403`, `/status/404`, `/status/429`, `/status/503` | Status tests; transient codes include Retry-After |
| `/__stats` | Loopback-only path request counters (no query values) |
| `/transport/feed/<scenario>` | RSS pointing to the same named media scenario; ordinary successful metadata |
| `/transport/metadata/<scenario>`, `/transport/media/<scenario>` | Bounded transient/permanent status, Retry-After delta/date, header/idle stall, active progress and one-time truncated response fixtures |
| `/transport/{metadata,media}/redirect-wait`, `/transport/{metadata,media}/redirect-defer` | Relative redirects with a one-second or five-second Retry-After; named `redirect-target` returns complete content |
| `/__transport` | Fixed transport route names, request arrival timestamps and advertised earliest retry times; no query/header values |
| `/resume/feed/<scenario>`, `/resume/media/<scenario>` | Stable synthetic RSS plus a 64-copy silent MP3; first response pauses after 64 KiB until the owned test releases `POST /__recover`. Named scenarios cover strong/weak/absent validators, missing length, ignored/changed/invalid ranges, 416 and renewed synthetic signatures |
| `/__resume` | Fixed resume route names, bounded numeric Range and known synthetic If-Range values, presence flags and an original/renewed enum; no raw query or arbitrary header values |
| `/resume/media/fresh-truncated-once`, `/resume/media/range-truncated-once` | A short first 200 body or short second-request 206 tail, followed by valid range handling, exercises durable checkpoints after caught failures |
| `POST /__reset` | Resets test counters; unsupported otherwise |

The fixed short delay exists to make tests quick. Configure a suitably smaller idle timeout only in the integration test. Inject clocks/delays for unit tests rather than waiting real long intervals. HEAD returns headers without a media body. The server handles HTTP/1.0 close-delimited responses intentionally; it is not a full production HTTP implementation.

The legacy `/media/recover.mp3` and `/media/interrupt.mp3` routes also omit ETag. Their tests continue to cover safe fresh restart and preservation of unknown crash partials. Validator-bearing durable resume is exercised separately by the `/resume/` scenarios.

Self-tests verify helper semantics and readiness-file safety. They do not exercise the PowerShell tool, interrupted archive recovery, TLS/proxy behavior, Windows file APIs or PowerShell versions. Codex must add those product tests in the planned tasks.
