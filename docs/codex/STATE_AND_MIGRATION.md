# Durable state, completion and legacy migration

## Implemented in UPD-0104 and UPD-0105

Each show uses `.upd/state.json`, with `.upd/state.json.bak` retaining the previous validated generation. Ordinary new histories use schema 1; an explicit legacy action promotes a validated copy to schema 2. Normal runs keep schema-1 archives at schema 1, and preview never promotes or writes history. Schema 2 adds `adopted` and `unverified` evidence levels. Writes reject schema downgrades, including a manually downgraded primary paired with a newer-schema backup.

The original exact feed URL's full SHA-256 identifies the feed and its first alias. Discovery scans immediate output folders for an exact alias fingerprint; it preserves the established directory when the title changes. Alias associations are explicit state entries, never inferred from titles, redirects or stripped URL queries. There is no alias-management CLI yet. Duplicate alias ownership, unreadable history or reparse directories stop discovery without creating a replacement archive.

Episode identity prefers RSS GUID, then Atom ID, then the exact media request URL. The identity source and its full fingerprint are persisted; the episode ID hashes that fingerprint with the feed ID. Exact source strings remain in memory. The planner rejects contradictory entries in the complete current snapshot before applying Latest/Custom selection. Across refreshes, a stable publisher ID binds the existing relative filename even if metadata changes. Without one, a new URL can create a new identity. Publisher reuse of a GUID across separate refreshes cannot be reliably distinguished from an edit.

Both schemas require exactly these top-level fields: `schema_version`, `feed_id`, `feed_alias_fingerprints`, `generation`, `episodes`. Episode records require `episode_id`, `identity_source`, `identity_fingerprint`, `relative_path`, `status`, `bytes`, `local_sha256`, `completed_utc`, and `verification` (`method`, `media_kind`, `notes`). Destinations are single Windows filenames, unique under case-insensitive comparison. Schema-1 statuses are `prepared`, `transfer_verified`, `missing`, `conflict`, `failed`; schema 2 also permits `adopted` and `unverified`. Prepared/completed records require positive bytes, a digest, UTC timestamp and recognized completion/signature evidence. Adopted records require positive bytes, a digest, UTC timestamp, a recognized local media signature, method `owner-approved-local-signature` and the exact confidence note `local-signature-only; transfer-completeness-unverified`. Adoption evidence cannot satisfy a `prepared` or `transfer_verified` record. History contains no raw feed URLs, media URLs or publisher IDs. Its relative filenames and fingerprints remain sensitive local archive data and are excluded from diagnostic exports.

Reads enforce strict UTF-8, 16 MiB, nesting depth 12, at most 100,000 episodes, 128 aliases and 16 notes of 256 characters each. Unknown/missing/duplicate fields, escaped property names, unsupported versions, invalid types and contradictory primary/backup pairs fail closed. Persisted generations start at 1; in-memory new state starts at 0. Schema promotion requires an explicit legacy action. There is no automatic backup restoration; preserve both files for inspection on error.

The brief output-root `.upd-archive.lock` covers a second discovery pass, show creation and initial state commit. The selected show then holds `.upd/writer.lock` through its run. Both are persistent files opened with exclusive OS handles; process exit/death releases ownership without deleting the file or trusting a PID. Writes require that show lock, validate the current generation, serialize to an owned `.upd/state.<GUID>.tmp`, flush and use same-volume File.Replace (or no-overwrite File.Move for first creation). No non-atomic replacement fallback is used. The previous valid state becomes the backup. Unclaimed temporary files remain untouched.

After media transfer validation, the downloader hashes the closed temporary file and commits `prepared` evidence before no-overwrite final placement. It then commits `transfer_verified`. A restart hashes the final file and compares bytes/digest against prepared or completed evidence before a verified skip; this closes the placement/history crash gap. Missing media with transfer evidence is safely downloaded again at its recorded path. Changed or unknown planned destinations become conflicts and produce an incomplete run. Adopted files are skipped only when their recorded bytes and SHA-256 still match, and the summary counts them separately as `Adopted`. Missing or changed adopted files and `unverified` records require review without automatic redownload or promotion to transfer evidence. Hashing denies write/delete sharing while the read handle is open. These checks do not promise a sandbox against a privileged concurrent filesystem attacker.

Local Windows tests exercise actual process death before/after media placement and state replacement, plus lock contention/release. Legacy tests use synthetic copied archives and compare every original filename, byte count and SHA-256 before and after preview, adoption, redownload and rollback. Hardware power-loss durability and live UNC filesystems remain unverified. UPD-0202 adds [validator-aware resume](RESUME_POLICY.md) without changing these completion evidence levels.

## Completion is an evidence level

A nonzero file is not necessarily complete. A matching historical filename is not identity. A local SHA-256 digest proves the bytes stayed the same relative to that digest, not that the publisher supplied the whole correct episode. Audio sniffing rejects obvious wrong data but is not a full decoder or a cryptographic guarantee. Use names like `transfer_verified`, `adopted`, `unverified` and `conflict` rather than an unexplained single boolean.

The HTTP body's expected byte count, when meaningful for the delivered representation, is stronger transfer evidence than the RSS enclosure length. Publisher metadata can be stale, and dynamically inserted content can change media. Mismatch against an enclosure estimate is a warning/review signal, not an automatic deletion/replacement rule. Missing length is supported when stream completion and other checks are satisfactory.

## State and checkpoint files

`src/HistoryStore.ps1` defines and validates the current schemas. Keep the existing show directory and media names. Migration stores additional full-history snapshots as `.upd/legacy-<32 lowercase hex characters>.json`; it never copies media into a checkpoint. These snapshots use the same strict schema reader and limits as live history. `src/ResumeStore.ps1` separately validates schema-1 resume sidecars and their owned byte checkpoints. Local subscription configuration remains future work.

Unknown newer schemas and corrupt primary, backup or selected checkpoint files fail clearly and remain untouched. History is parsed as data, never evaluated as PowerShell.

### Data retention and diagnostic boundary

Keep history, its backup and migration checkpoints with the archive: they support verification, identity binding and explicit metadata rollback. The downloader does not age them out or delete originals, unknown partials or old checkpoints. Removing history can remove the evidence needed for a verified skip or a recovery decision.

Diagnostics have separate retention. Startup logs, show logs and optional JSON exports remain until the owner deletes them; no automatic upload, rotation or deletion is implemented. Older logs may contain raw credentials and are never ingested or re-exported. Current logs avoid raw URLs, titles and paths, but hostnames and identity fingerprints may still reveal or correlate subscriptions. The restricted export includes none of those archive fields. See [DIAGNOSTICS.md](DIAGNOSTICS.md).

Legacy inventory deliberately returns exact filenames, titles and identity data needed for local adoption decisions. Those objects, checkpoint result values, shell input history, transcripts and caller-inspected original exceptions are sensitive local review data. They are excluded from diagnostic export; manually saving or sharing them is a separate owner action.

## Identity policy

UPD-0204 date normalization changes chronological selection and new UTC filename prefixes only. Existing stable identities keep their recorded relative destinations. The established publisher-ID lookup remains compatible. Old-priority machine-local date hints are retained solely for unverified legacy review; see [PUBLICATION_DATES.md](PUBLICATION_DATES.md). No media/history is renamed or promoted when a date parses differently.

Prefer a stable feed-scoped RSS GUID or Atom ID. Avoid title/date as the primary key. Handle repeated identical entries across pages but flag contradictory reuse of an identifier rather than silently merging unrelated records. Store a stable local feed ID and explicit alias associations so a publisher/feed title change does not create a new archive automatically.

Fallback identity may use a conservative URI/metadata fingerprint. Do not remove all query parameters from the actual request or assume all query values are disposable tracking tokens. Without a stable publisher identity, changes in title, URL and signing tokens can be genuinely ambiguous; document this limit and preserve conflicting files. Do not assert universal perfect deduplication. A digest suffix solves collisions only when backed by a full key and conflict check; a short hash alone is not a guarantee.

Persist only relative destinations and validate containment when reading state. History is not trusted merely because the program wrote a prior version. State that contains raw URL credentials is sensitive local data; prefer fingerprints and separate protected secret storage. Config should avoid secret-bearing URLs in exported examples and warn/protect private entries. A Windows-native protected-secret option can be considered without imposing a new module dependency; document portability limits.

## Safe write/finalize sequence

Acquire archive writer protection before modifying state/download ownership. Reserve a new owned temporary file next to the destination or verify an existing resume checkpoint under its exclusive handle. After transfer validation, associate its episode identity and byte/hash evidence with a durable `prepared` record before final placement. Retire the matching resume sidecar before moving the validated partial to its final path. Never adopt a random `.part` left by another program based solely on extension.

Stream to the temporary path, close it and validate. Recheck the destination. Finalize using a no-overwrite operation. Then update state through a temporary state file plus same-volume replacement/backup strategy supported on the chosen Windows runtimes. Keep the previous good state until the new record is valid. Release handles in finally. Document filesystem/UNC constraints; a rename is not a universal guarantee against hardware power-loss corruption.

Test crash gaps: before final rename; after rename before state commit; during state replacement; after cancellation; with concurrent writers; with changed/deleted completed files. Reconciliation must never overwrite a final file because history lagged. Both history and migration use the existing exclusive writer handles.

## Writer conflicts and preflight (UPD-0302)

The output-root selection lock is brief; the per-show writer lock stays held through transfer and history updates. Two runs for one show are excluded with a fixed message advising a retry after the writer finishes. Sharing conflicts and inaccessible locks have distinct guidance. Different shows under the same output root can transfer concurrently once root selection finishes; unrelated roots are independent. Ordinary OS handle release after process exit permits reopening a persistent lock file. Its contents, age and any stored PID are not authority: the downloader neither removes stale lock files nor kills a process based on them.

Confirmed ordinary and legacy changes perform an owned write probe before destination/state writes. For a missing root, the probe uses its nearest existing ordinary ancestor, then the actual show is checked under its writer lock. CreateNew collisions remain unclaimed; only the successfully reserved DeleteOnClose probe is disposed. Preview and declined confirmation make no probes or archive writes. File creation does not establish separate directory-creation rights; actual directory failures retain fixed permission guidance. Access checks can race later permission changes.

After validating response headers and before copying media bytes, compare known remaining response bytes with available destination capacity. A validated 206 needs only its remaining tail; RSS enclosure sizes are advisory. AvailableFreeSpace reflects the caller's quota on a ready local/mapped drive. Arbitrary UNC and failed queries are unknown, and an unknown response length cannot support a size comparison. Those observations produce fixed explanatory output and ordinary transfer validation continues. Observed zero bytes is not unknown. No extra HEAD/GET is added; accepted headers and initial/failed-attempt history may precede a size rejection. This snapshot cannot guarantee capacity throughout a transfer.

Catchable cancellation preserves partials and attempts an actual-byte checkpoint before closing a resumable stream. Cleanup attempts every acquired run/media/legacy handle and preserves the primary failure; a failing Dispose cannot guarantee release. If the final checkpoint write fails, prior evidence and a longer partial stay for review and strict resume refuses their mismatch. A retained unknown-length or uncheckpointed partial receives no fabricated sidecar and remains unclaimed on a fresh restart. Never delete unknown partials to recover space. See [UPD-0302 evidence](evidence/UPD-0302.md) for the real process/ACL/transfer observations and limits.

## Resume policy (UPD-0202)

Resume requires a valid owned sidecar, exact local prefix size/hash, matching feed/episode/request/target identities, a strong ETag, a known total and matching representation. Only a validated complete-tail 206 can append. Ignored ranges, changed validators/representation and 416 use one fresh GET into a new partial while preserving the old files. Corrupt sidecars or uncheckpointed crash tails stop for review. Full or empty checkpoints are not completion evidence. See [RESUME_POLICY.md](RESUME_POLICY.md) for the schema, geometric checkpoint schedule, atomic replacement, privacy and crash rules. Unknown partials are never inferred to be owned from their names.

Resume mechanics depend on HTTP semantics [S6 in SOURCES.md]. The supplied HTTP scenarios are adversarial tests for this project's implementation. They are not proof that all real providers implement ranges correctly.

## Implemented legacy review and migration

`-LegacyPath` selects an existing immediate show directory under `-OutputPath`, as either an absolute path or folder name. It requires an explicit `-FeedUrl`; the legacy workflow uses the complete current feed snapshot, independent of the ordinary Latest/Custom selection. Conflicting feed ownership, invalid paths and reparse directories fail before a decision is applied. The [README examples](../../README.md#review-and-migrate-a-legacy-archive) show the commands and returned objects.

`-LegacyAction Preview` is the default. It may fetch the feed and read local files, but never creates output directories, media, state, logs, configuration or lock files. `-WhatIf` also leaves the output tree unchanged for normal downloads and all legacy actions. A changing legacy action with `-WhatIf` still validates its explicit selection and reports its proposed operation.

Inventory includes immediate ordinary files regardless of extension. It hashes readable originals and holds a read-only handle while hashing and inspecting up to 64 KiB for a recognizable local signature. It reports unsafe/reparse paths without following them and does not recurse into child directories or `.upd`. Files named as partials, empty bodies and obvious text remain conflicts. Unknown plausible media remains `unverified`; no observation claims an observed HTTP transfer.

Matching covers the original title/date naming scheme and its short SHA-1 fallback, UPD-0102 full-hash suffixes and current feed-scoped identity suffixes. Names only identify candidates. A suggestion requires exactly one candidate file for one episode, with no other history record owning that path. Repeated title/date names, multiple copies and ownership collisions remain conflicts. The plan exposes `Files` and `Episodes`; episode rows include the full `EpisodeId`, candidate paths and a nullable `SuggestedPath`. Existing confirmed history paths are labelled `recorded`, with `RecordedStatus` retaining the actual evidence level.

An ordinary run checks the proposed archive and, for a new feed binding, the old title-only folder. If selected unbound episodes could cause new downloads while that folder contains unclaimed media, it returns a review plan under `-WhatIf` or stops with a review message. The check repeats under writer protection before initial state writes. A folder title is a reason to request review; it never establishes the feed or episode identity. Existing archives with an unrelated recorded feed are not claimed by title.

In an established archive, an abandoned file with the exact `.upd-<32 lowercase hex characters>.tmp` transfer naming pattern is preserved without blocking a fresh transfer. Its name grants no completion or adoption evidence. Only a validated resume sidecar and matching bytes can authorize reuse. Other unclaimed partial names remain part of legacy review.

### Adopt

`-LegacyAction Adopt` requires a full `-LegacyEpisodeId`, an exact `-LegacyFile` basename and the reviewed `-LegacySha256` from the preview. Only one current episode and one plausible local file can be selected. The owner may explicitly resolve ambiguous names or choose a differently named file. A path owned by another episode is refused. Existing completed/adopted evidence must be handled through explicit redownload rather than silently rebound.

The workflow repeats inventory and validates the chosen bytes while holding the archive writer locks, then keeps a read-only media handle through the state commit. A changed reviewed digest, partial name, empty/text file or unrecognized signature prevents adoption. It records `adopted` with the limited confidence described above and preserves the original path and bytes. No media request or remote size comparison supplies adoption evidence.

### Redownload

`-LegacyAction Redownload -LegacyEpisodeId <full ID>` downloads only the explicit current episode to a new filename with a fresh full-hash suffix. Existing filenames and bytes survive, including an earlier file at the usual modern destination. The existing prepared/finalize history protocol records the new observed transfer. A changed remote body never authorizes overwriting an original. Other unmatched files stay available for later review; the action does not adopt or redownload the rest of the archive.

### Checkpoints and rollback

Before its final metadata change, each changing legacy action saves a unique `.upd/legacy-<id>.json` checkpoint and prints its basename. Successful results also return it as `Checkpoint`. The snapshot contains the entire pre-action history metadata, including unrelated episode records; it contains no media bytes. A first migration initializes schema-2 empty history before saving that checkpoint. Existing schema-1 metadata is copied into schema 2 without discarding records; its prior generation remains protected by the ordinary atomic backup mechanism.

`-LegacyAction Rollback -LegacyCheckpoint <basename>` requires a readable checkpoint for the same feed and exact aliases whose generation is no newer than the current state. Its preview includes `Operation.RestoreRecords`. Rollback restores all snapshot records in a new current generation, keeps schema 2 and the feed association, and saves another checkpoint before restoring. Thus rolling back an older checkpoint also removes later metadata decisions. Rolling back the first adoption leaves an empty schema-2 history with the same feed association.

Rollback never deletes, renames or restores original media, separate redownloads, logs or existing checkpoints. Files whose bindings were removed remain on disk and need review. A failed redownload can use its printed checkpoint for metadata recovery; media already placed before a failure remains preserved. Corrupt state still fails closed, so checkpoint rollback is not a bypass around an unreadable primary/backup. Preserve those files for inspection instead of replacing them automatically.
