# Audio selection and original file bytes

UPD-0205 implements bounded audio selection and filename extension resolution. The downloader stores the exact received bytes. It does not transcode, rewrite tags, decode a complete track, prove playability or authenticate a publisher. HTTP completion/length and a recognized prefix provide limited transfer evidence.

## Enclosure planning

RSS enclosures and Atom enclosure links are retained in source order. Supported audio MIME hints are MP3 (`audio/mpeg`, `audio/mp3`, `audio/x-mp3`, `audio/x-mpeg`, `audio/x-mpegmp3`), M4A (`audio/mp4`, `audio/m4a`, `audio/x-m4a`), Ogg (`audio/ogg`, `audio/opus`, `audio/vorbis`, `audio/speex`), WAVE (`audio/wav`, `audio/wave`, `audio/x-wav`, `audio/vnd.wave`) and FLAC (`audio/flac`, `audio/x-flac`). MIME comparisons ignore case, parameters and surrounding whitespace.

Missing/generic MIME (`application/octet-stream`, `binary/octet-stream`, `application/ogg`, `application/mp4`) may use `.mp3`, `.m4a`, `.ogg`, `.opus`, `.wav` or `.flac` URL hints. Other explicit MIME types, including video and unsupported audio codecs, are ineligible. Select the first known audio candidate in source order; a later MIME hint does not reorder an earlier eligible audio URL. When none is known, the first generic enclosure without a supported audio hint may be selected provisionally and must subsequently pass the signature checks. A legacy GUID/link fallback is considered only when no enclosure URL was supplied and it has a supported audio extension. URLs remain subject to the shared HTTP(S) policy. Selection never requests enclosures during parsing or preview.

The selected URL remains exact for URL-based identity; GUID/Atom ID priority is unchanged. A publisher changing the selected URL of an episode without a stable publisher ID changes its identity under the existing policy. This work does not infer equivalence between distinct media URLs or silently adopt old files.

## Completed byte checks

| Received prefix | Canonical new extension | Limited evidence |
| --- | --- | --- |
| MPEG Layer III, optionally after a bounded ID3 skip | `.mp3` | Plausible frame header and first-frame length; Layer I/II are unsupported |
| RIFF/WAVE | `.wav` | Declared RIFF container fits stored bytes |
| FLAC | `.flac` | STREAMINFO header layout |
| Ogg Opus, Vorbis or Speex | `.ogg` | First complete BOS packet identification within lacing boundaries |
| ISO BMFF with M4A brand, or bounded audio handler evidence | `.m4a` | M4A convention or structurally located `soun`; observed `vide` rejected |

The existing read budget is at most 65,536 bytes: a 61,440-byte prefix plus up to 4,096 bytes after an ID3 tag. ISO BMFF traversal uses only that prefix, bounded box visits/depth and actual nested boundaries. Generic containers without sufficient audio evidence are `ambiguous_media`; recognized video/unsupported MPEG layers are `unsupported_media`. Unknown binary is `unrecognized_media`. HTML/XML/JSON/plain text is rejected as `non_audio_text`, even with an audio header. Missing/invalid transport completion, HTTP framing or empty bodies fail before media classification.

M4A branding is only a compatibility indication: the registry permits audio and video in that brand. An unobserved video track or corrupt later payload can escape these modest checks. A generic MP4 whose required audio structure is outside the inspection prefix is rejected rather than scanned without a bound. Ogg checks identify the initial audio codec packet; later pages, chained streams, checksums and decoding are not verified. WAVE/FLAC checks are similarly modest structural probes. Synthetic test bodies prove classification, byte preservation and transaction behavior, not that every fixture is playable.

Response/feed MIME and URL extensions are advisory for recognized audio bytes. Contradictions create fixed warnings without repeating untrusted header/URL text. A misleading audio MIME cannot make text or unknown bytes pass; a generic or incorrect header cannot by itself reject a recognized supported audio prefix. Publisher enclosure length is advisory. Content-Disposition never supplies a destination name.

## Paths, resume and history

Planning uses a MIME/URL extension hint, with `.mp3` as the provisional extension when unknown. Preview cannot predict every final extension and performs no media request or persistent write. A new episode's validated bytes resolve its final extension before prepared history is saved. The final name is rebuilt with the original episode identity and filename budget. History ownership, containment, reparse checks and no-overwrite placement apply to the resolved path. The prepared record stores that exact final path, size and SHA-256 before rename; a subsequent run can reconcile an already placed file.

An established history record retains its relative path, including a historical extension that disagrees with newly received bytes. Such a transfer records a warning; no existing file is renamed or overwritten. A matching completed file still skips after hashing its retained bytes. A prior failed/missing allocation with both byte count and digest absent has no completed-byte evidence and may resolve its final extension after successful transfer; its checkpoint remains bound to the original provisional path until retirement. Prepared/completed/adopted evidence and unverified/conflict bindings remain pinned. Explicit legacy Redownload creates a separate allocation with its own hash and may resolve that new file's extension while preserving original media.

Resume checkpoints continue to own the exact provisional planning path throughout the attempt. They are never rewritten to a sniffed extension. The existing order saves prepared evidence, retires the checkpoint, then places the file. A crash after corrected-path preparation but before checkpoint retirement can leave a missing final file and a checkpoint bound to the previous provisional path. Recovery then stops on the exact-path mismatch, preserving evidence for review. It never appends under changed ownership. A crash after placement still reconciles from the prepared final path. These are filesystem transaction checks, not universal power-loss guarantees.

A completed body that fails audio classification may already have a valid owned byte checkpoint. Its partial and sidecar remain for review under the existing resume policy; completion of those bytes does not confer audio or history-completion evidence. An uncheckpointed temporary file created by the failed attempt is cleaned up. Neither case places final audio or records a verified transfer.
