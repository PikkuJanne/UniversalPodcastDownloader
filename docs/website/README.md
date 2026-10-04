# Portable distribution content

This directory prepares text and metadata for later website integration. It contains no framework, hosting configuration, domain/DNS change, deployment or hosted downloader. No site repository or destination has been selected. Creating or copying these files does not publish a release or authorize a live site.

| File | Purpose |
| --- | --- |
| [PRODUCT_PAGE.md](PRODUCT_PAGE.md) | Product copy with current candidate status, requirements, usage, privacy, limits, support and license |
| [version.example.json](version.example.json) | Explicit unpublished version/asset placeholders; data only, not an application update feed |
| [SCREENSHOT_CAPTURE.md](SCREENSHOT_CAPTURE.md) | Real local TUI capture and privacy/provenance instructions |

The example is `0.1.0-rc.1` / `UNRELEASED_CANDIDATE`. `currentPublishedVersion`, source identifiers, tag, publication time and all release/download/checksum/manifest/note URLs stay `null`. Expected filenames are labels, not download links. The repository/support URLs are navigation links, not released assets. No screenshot is embedded or claimed captured; its metadata remains unassigned. Keep these distinctions visible when importing the content elsewhere.

The product copy reflects implemented source and current user guidance. Preparing it does not complete final clean-clone acceptance or owner publication review. [Release requirements](../codex/RELEASE_AND_WEBSITE.md), [draft notes](../releases/DRAFT-0.1.0-rc.1.md) and the [approval checklist](../releases/APPROVAL_CHECKLIST.md) define the remaining review/action gates. CI artifacts are expiring review candidates; do not advertise them as stable/public release downloads.

## Maintain links after an authorized publication

1. Obtain completed candidate acceptance and explicit owner authority for the exact source, version, tag, asset bytes and publication action. Website deployment has separate scope. Do not infer authority from this content or a PR review, and do not create even a GitHub draft release for an absent tag without tag authority.
2. Verify the actual published release and remote tag against the approved full source commit/tree. Download the exact published ZIP, manifest and `SHA256SUMS`; compare their hashes, the embedded manifest and every payload entry with the reviewed source. Record the release/run/asset identities and evidence. Changed source or bytes need renewed review.
3. Create/update the eventual site's real version record from those verified facts. Keep this file explicitly named/documented as an example if retaining it. Set version and published state according to the actual release; an approved published prerelease must remain visibly a prerelease and must not be described as stable. Set `currentPublishedVersion` only to the published version being presented, record the exact existing tag, full source commit/tree, immutable commit URL and actual UTC ISO-8601 publication time.
4. Copy the exact release asset URLs from GitHub's published release for the ZIP, manifest and checksum list. Use the exact tag-specific release page for notes, or its verified version-specific notes asset. Store the observed ZIP/manifest SHA-256 values. Do not invent URLs, use `/releases/latest`, a moving branch ZIP or an expired Actions artifact as a version download. Leave any unavailable asset URL `null` and display it as unavailable.
5. Update the product page's status, version and download/checksum/source/note links together. Rewrite repository-relative README/RELEASE/CHANGELOG/LICENSE links for the destination using shipped documentation or the verified immutable source commit; check every resulting link. Keep compatibility, privacy, support and known-limit claims aligned with the reviewed version. Include a real inspected screenshot only with its actual source/engine/capture record; see [capture instructions](SCREENSHOT_CAPTURE.md).
6. Independently download through the final page links and verify the same bytes again. Record the approved action and observed outcome; if any field/link cannot be verified, retain the candidate/unavailable wording. This maintenance procedure is manual; it adds no runtime updater, telemetry or scheduled publication.

## Metadata field meanings

`schemaVersion` versions this illustrative data shape. `version` names the declared candidate; `releaseStatus` controls how that record is presented. `currentPublishedVersion` being `null` means this example announces no published version. Null source/tag/date/URL/hash fields are deliberately unknown, not instructions to derive values from the newest branch or release.

`assets` names the three separately reviewed portable files. The ZIP SHA-256 differs from the outer Actions artifact digest; the sidecar manifest checksum covers its complete bytes. Hashes detect differences and do not authenticate the publisher. `sourceCommitUrl` must point to the exact full commit, not a branch. `releaseNotesUrl`/`releaseUrl` refer to verified published material only. Screenshot fields stay null until a reviewed real capture exists.

Validate the JSON as data, compare its version/filenames with [release-package.json](../../tools/release-package.json), review candidate/null invariants and check relative links before handing the files to a future site maintainer. No local file, private feed URL, configuration, history, media, raw log or secret should become website metadata or an uploaded asset.
