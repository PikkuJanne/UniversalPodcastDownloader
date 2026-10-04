# Release and website preparation

## Public candidate requirements

Preserve the MIT license and existing artwork unless the owner requests a change. Select a version based on actual existing tags/releases at implementation time; do not invent a prior release series or assume version 1.0.0 is free. The handoff's 1.0.0 is **only a bundle version**, not the application version.

Ship a portable ZIP with the root launcher/script plus any runtime module files, license, quick start and meaningful release notes. Include version/source identifiers and SHA-256 checksums. The build must not rely on files omitted from the ZIP or a developer's checkout layout. Test a fresh extraction in a path with spaces and non-ASCII characters. Keep Python, Pester, test fixtures, generated media, personal logs, history/config and handoff material out of the user release.

A checksum helps compare downloaded bytes; it is not by itself publisher authentication. Signing can improve provenance but requires an actual signing workflow and owner decision. Do not create self-signed certificates and imply public trust. Do not buy certificates/services or modify trusted certificate stores automatically. If code is unsigned, document that honestly and explain the launcher/process policy without recommending machine-wide security weakening.

## Draft release checklist

- All advertised features are implemented and supported by current engine evidence.
- All mandatory acceptance cases are resolved; any owner-approved scope deferral is clearly reflected in feature claims.
- Version, source commit, package manifest and checksum list agree; all imported modules are included.
- Fresh extraction works; copied legacy files retain original hashes after migration tests.
- Privacy/security limitations and partial-catalogue behavior are explicit.
- Draft notes describe improvements, compatibility, recovery changes and known limitations.
- No tag/publication action occurs without explicit authorization; creating a draft must not implicitly create a new release tag.

## Website deliverables — content first

There is no selected tools-site repository, domain or host in this request. Create portable source content under an appropriate docs directory, not an unrelated live site. A static Markdown page and release metadata are sufficient until the website architecture is chosen. The site is a distribution/support page, not a hosted podcast downloader.

Suggested page sections: product name and one-sentence purpose; current released version and Windows requirements; verified download/checksum/source links; real TUI screenshot; three-step quick start; core modes; output/recovery; privacy; limitations; license; changelog and bug report instructions. Use a real capture from the implemented program rather than generated imagery. Do not fabricate screenshots before the UI exists.

Use explicit placeholders such as `UNRELEASED_CANDIDATE` in unpublished metadata. Do not present a planned feature or draft build as stable. After a release is actually approved and published, update links to the exact tested GitHub release asset rather than a moving main-branch download. No automatic update checker/network beacon is added to the tool as part of a product page.

Privacy wording should reflect implementation: local downloads/state; network requests to user-selected feed/media hosts; no product telemetry unless explicitly designed, disclosed and approved. Avoid an absolute claim that third-party podcast hosts cannot log requests. Private feed URLs and redacted diagnostic sharing deserve clear instructions. Explain that the software license does not automatically license podcast recordings.

## Approval record

Record owner approval scope, date, exact candidate/source/tag, and the requested action before merging/publishing/deploying. After an approved action, verify the resulting GitHub artifact/link or deployed page and record outcome. Do not interpret an approved PR review as automatic permission for a hosting purchase or unrelated settings change.
