# Safe resume policy (UPD-0202)

Resume is automatic for a recorded episode under the existing archive writer lock and confirmed download operation. It adds no command-line option or runtime dependency. Original media bytes, history schemas 1/2, prepared completion evidence and no-overwrite final placement remain the completion protocol.

## Eligibility and response decisions

A fresh HTTP 200 can create resume evidence only with a positive Content-Length, one strong quoted ASCII ETag of at most 1,024 characters, a parsed non-multipart media type and identity content encoding. Weak or absent ETags, Last-Modified alone and unknown total length use ordinary fresh transfers. Accept-Ranges is only advisory and is not required.

An append requires all of the following before any response bytes reach the partial:

- The sidecar belongs to the full feed and episode identities and the exact planned filename.
- The original request URI fingerprint still matches, including query data; the exclusively opened partial has exactly the saved offset and SHA-256.
- The offset is greater than zero and less than the saved total.
- Range `bytes=<offset>-` and If-Range carry the saved strong ETag only at the previously observed final URI. Redirect hops use fresh request headers; a validator is never sent to a different target.
- HTTP 206 returns the same final URI fingerprint, ETag, canonical media type and identity encoding. Content-Range must be a single byte range from the saved offset through total minus one, with that exact known total. Content-Length, when supplied, must equal the remaining tail length.

After appending, verified EOF must show exactly the validated tail length. Completion checks then use the combined file length and original media validation; a short tail is not completion.

An ignored Range (200), changed or missing validator/representation, changed final target, inconsistent range, multipart response or 416 closes the old stream and makes one fresh GET into a new temporary file. The old partial remains byte-for-byte intact. A fixed diagnostic explains the restart without including headers or URLs. The outer UPD-0201 policy alone controls retries and backoff. A full GET that fails still follows its ordinary classification; an unsolicited 206 or 416 cannot become a completed transfer.

The policy deliberately requires the complete remaining tail and repeated representation headers. RFC 9110 permits some shorter ranges or omitted headers; those cases conservatively restart here. Media type comparison ignores parameters after parsing. A local full-length partial, matching hash or 416 total never proves that an earlier response reached verified EOF. Full/empty checkpoints require a fresh GET, preserving their old partial.

## Sidecar schema and privacy

The active sidecar is `.upd/resume-<first 32 hex characters of episode ID>.json`. Its full identity detects filename collisions. Schema 1 requires exactly these fields:

| Field | Meaning |
| --- | --- |
| `schema_version` | Integer 1; independent of history schema |
| `feed_id`, `episode_id` | Full lower-case SHA-256 identities |
| `relative_path` | One validated destination filename |
| `partial_name` | Owned sibling `.upd-<32 lower-case hex characters>.tmp` |
| `request_fingerprint`, `final_uri_fingerprint` | SHA-256 of each absolute request URI, retaining query distinctions |
| `etag` | Exact strong validator, needed for If-Range |
| `total_length` | Positive integer representation length |
| `content_type`, `content_encoding` | Canonical media type and literal `identity` |
| `offset`, `prefix_sha256` | Durable byte count and SHA-256 of that entire prefix |

Reads enforce strict UTF-8, a 16 KiB limit, nesting limit 12, exact names/types, no duplicate or escaped property names, safe destinations and supported schema. JSON is data only. No raw URI or publisher ID is persisted. ETags, relative filenames and fingerprints are sensitive local metadata; sidecars and their snapshots are excluded from diagnostic exports. They are not uploaded.

## Checkpoint and recovery rules

The current partial stays open with ReadWrite access and FileShare.None. Hashing uses that same handle and restores its position. Before the first fresh body byte, eligible headers create a flushed zero-offset sidecar; resumed headers retain the verified offset. The first successful nonempty chunk is checkpointed. Later progress checkpoints occur when the length reaches the larger of 8 MiB or twice the previous checkpoint offset. Success and caught network failures force a final checkpoint. Each checkpoint flushes the media to disk, hashes it, flushes a unique owned metadata temporary, and uses File.Replace (or initial no-overwrite File.Move). Geometric spacing makes progress-checkpoint hashing within an attempt proportional to file size. Recovery, headers and forced checkpoints can each hash an existing prefix again, even when little new data arrived.

A caught transient network failure can therefore resume from its latest written prefix on the next bounded attempt. An abrupt process kill is resumable only when the on-disk file length and digest exactly match its last committed checkpoint. Extra uncheckpointed tail bytes, shorter/corrupt files, missing partials, corrupt sidecars, identity conflicts or unavailable atomic replacement fail conservatively and preserve the files for review. Recovery never truncates an uncertain tail or guesses a new offset.

Switching to a new eligible partial atomically preserves the exact old sidecar as a unique `.upd/resume-<random 32 hex characters>.old.json`. The old partial remains. Ordinary progress replaces only the active sidecar. An ineligible fresh fallback may leave the old active sidecar and its partial even after success; only a sidecar matching the completed partial is retired. The downloader does not age out snapshots or unknown partials. A fresh ineligible failed transfer cleans only its newly reserved file. A file whose ownership commit was attempted is preserved if that commit's outcome is uncertain.

After whole-file validation, hashing and the prepared history commit, the active matching sidecar is removed before final no-overwrite placement. This prevents a crash from leaving an active checkpoint that references a partial already moved to its final name. A crash before placement leaves an unclaimed partial and safely starts fresh; a crash afterward uses the prepared history evidence. Sidecar cleanup never deletes media. Existing path/reparse checks and lock checks apply at read, write and finalization boundaries.

This validates process-crash recovery and ordinary concurrent-writer exclusion. It does not claim power-loss durability on every filesystem, live UNC testing, or protection against a privileged filesystem attacker. Local digests establish consistency, not publisher authenticity.
