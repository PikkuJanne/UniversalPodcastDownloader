# Durable state, completion and legacy migration

## Completion is an evidence level

A nonzero file is not necessarily complete. A matching historical filename is not identity. A local SHA-256 digest proves the bytes stayed the same relative to that digest, not that the publisher supplied the whole correct episode. Audio sniffing rejects obvious wrong data but is not a full decoder or a cryptographic guarantee. Use names like `transfer_verified`, `adopted`, `unverified` and `conflict` rather than an unexplained single boolean.

The HTTP body's expected byte count, when meaningful for the delivered representation, is stronger transfer evidence than the RSS enclosure length. Publisher metadata can be stale, and dynamically inserted content can change media. Mismatch against an enclosure estimate is a warning/review signal, not an automatic deletion/replacement rule. Missing length is supported when stream completion and other checks are satisfactory.

## Suggested state layout — implement and version, do not copy blindly

Keep local state beneath each established podcast directory, for example `.upd/state.json`, plus owned partial/resume records. Keep an owner-only local subscription index/config outside source control. Existing output folders stay in place unless explicitly migrated.

```json
{
  "schema_version": 1,
  "feed_id": "stable-local-feed-identifier",
  "feed_alias_fingerprints": [],
  "generation": 1,
  "episodes": [
    {
      "episode_id": "feed-scoped-stable-identifier",
      "identity_source": "rss-guid",
      "identity_fingerprint": "digest-not-raw-secret",
      "relative_path": "2026-09-01 - Example.mp3",
      "status": "transfer_verified",
      "bytes": 12345,
      "local_sha256": null,
      "validator": null,
      "completed_utc": "2026-10-01T12:00:00Z",
      "verification": {"method": "completed-transfer", "notes": []}
    }
  ]
}
```

This is an illustrative contract, not a production schema or a file to place in someone's archive. Define field limits, unknown-version behavior, serialization depth and migration strategy. Never evaluate strings as PowerShell. Unknown newer schema: preserve and fail clearly. Corrupt state: preserve it, do not silently initialize an empty successful history.

## Identity policy

Prefer a stable feed-scoped RSS GUID or Atom ID. Avoid title/date as the primary key. Handle repeated identical entries across pages but flag contradictory reuse of an identifier rather than silently merging unrelated records. Store a stable local feed ID and explicit alias associations so a publisher/feed title change does not create a new archive automatically.

Fallback identity may use a conservative URI/metadata fingerprint. Do not remove all query parameters from the actual request or assume all query values are disposable tracking tokens. Without a stable publisher identity, changes in title, URL and signing tokens can be genuinely ambiguous; document this limit and preserve conflicting files. Do not assert universal perfect deduplication. A digest suffix solves collisions only when backed by a full key and conflict check; a short hash alone is not a guarantee.

Persist only relative destinations and validate containment when reading state. History is not trusted merely because the program wrote a prior version. State that contains raw URL credentials is sensitive local data; prefer fingerprints and separate protected secret storage. Config should avoid secret-bearing URLs in exported examples and warn/protect private entries. A Windows-native protected-secret option can be considered without imposing a new module dependency; document portability limits.

## Safe write/finalize sequence

Acquire archive writer protection before modifying state/download ownership. Reserve an owned temporary file next to the destination on the same volume. Associate its episode identity, expected entity evidence and byte count with a versioned sidecar. Never adopt a random `.part` left by another program based solely on extension.

Stream to the temporary path, close it and validate. Recheck the destination. Finalize using a no-overwrite operation. Then update state through a temporary state file plus same-volume replacement/backup strategy supported on the chosen Windows runtimes. Keep the previous good state until the new record is valid. Release handles in finally. Document filesystem/UNC constraints; a rename is not a universal guarantee against hardware power-loss corruption.

Test crash gaps: before final rename; after rename before state commit; during state replacement; after cancellation; with concurrent writers; with changed/deleted completed files. Reconciliation must never overwrite a final file because history lagged. The basic exclusive writer primitive lands with history, not only in later usability polish.

## Resume policy

Resume only program-owned partials with consistent episode and entity evidence. Prefer a suitable strong validator; otherwise restart conservatively. Validate resumed response/offset/total before appending. Full-body responses replace/restart the owned partial safely, not append. Inconsistent range, changed entity, weak/absent evidence, corrupted sidecars and 416 require an explicit tested branch. Partial files belonging to the user or another program are not cleanup targets.

Resume mechanics depend on HTTP semantics [S6 in SOURCES.md]. The supplied HTTP scenarios are adversarial tests for this project's implementation. They are not proof that all real providers implement ranges correctly.

## Legacy adoption procedure

1. **Inventory without mutation.** Match existing names and potential identities; classify obvious empty/bad files and ambiguous collisions. Hash originals in a copied test archive to prove safety.
2. **Preview.** List would-adopt, unverified, ambiguous and would-download actions. Do not create history/log/config folders in WhatIf.
3. **Explicit choice.** Adoption can accept an existing file with a clearly stated confidence level; redownload goes to a separate safe target and never overwrites the old file by default. A remote size comparison is not guaranteed proof of identical media.
4. **Record, do not rewrite.** Preserve all original filenames/bytes. Bind accepted identities to existing paths; leave uncertain files untouched. Record adoption policy and observed evidence.
5. **Recovery.** Preserve prior state for rollback. Rollback restores owned metadata only, never deletes originals. A rerun should offer reconciliation rather than blindly redownload the whole archive.

For normal runs, report unknown existing files as `legacy_unverified`/conflict and an incomplete/review-needed outcome, not as verified skips. Do not make migration an automatic destructive startup operation.
