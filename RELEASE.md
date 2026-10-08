# Portable release and source guidance

**1.0.0** is the first stable distribution of the completed application. Use the [v1.0.0 release page](https://github.com/PikkuJanne/UniversalPodcastDownloader/releases/tag/v1.0.0) for the published ZIP, manifest, checksums and release notes. See [changelog](CHANGELOG.md), [usage and recovery](README.md) and the original [MIT license](LICENSE). The earlier `0.1.0-rc.1` packaging candidate and its acceptance observations retain their historical identities.

## Use the extracted application

Extract the ZIP into a new ordinary directory, including all of `src`. Keep the `.ps1` and `.bat` files together. Double-click `UniversalPodcastDownloader.bat` for the guided flow, or run the script in Windows PowerShell 5.1/PowerShell 7 on Windows. `Get-Help .\UniversalPodcastDownloader.ps1 -Full` works from this directory; the README contains the quick start and verified examples. The application needs no Git, Python, Pester, package manager, FFmpeg or checkout directory. Build tooling is separate.

The script is unsigned. The batch launcher uses process-only `-ExecutionPolicy Bypass` for its child; it does not change machine/user execution policy. Respect managed Windows restrictions. Do not weaken machine-wide policy or modify trusted certificate stores to run it.

## Check bytes and source

The build produces a versioned ZIP, a JSON manifest and `SHA256SUMS`. Compare SHA-256 values with `Get-FileHash -LiteralPath <file> -Algorithm SHA256` before extraction. The checksum list covers the ZIP and sidecar manifest. The same manifest bytes are inside the ZIP as `manifest.json`; it records version, release channel, full source commit/tree, repository URL, and the size/SHA-256 of every declared source file. `STABLE_RELEASE` identifies a stable semantic version; `UNRELEASED_CANDIDATE` identifies an rc version. Neither channel value alone proves publication: verify the matching published GitHub release and tag. The manifest cannot hash itself; the external checksum covers its bytes. Validate the extracted files against its list if checking individual contents.

Use the manifest's exact source commit URL when reviewing or obtaining source; do not substitute a moving branch download. README links to extended policy/evidence open online at the previously verified runtime/documentation checkpoint. Those reference documents are outside the portable ZIP, and the application does not read them at startup. The new manifest identifies this package's complete committed source.

Checksums detect differing bytes when compared with a trusted expected value. They do not authenticate the publisher if an attacker can replace both the ZIP and checksums. No Authenticode signature, public trust or signing certificate is claimed. Entries use ordinal path order, a fixed 1980 ZIP timestamp and zero external attributes. Inventory and source bytes are deterministic. Repeated builds with the same engine/compression implementation are checked for equal bytes; equality across different compression implementations is not guaranteed.

## Build locally from committed source

Build tooling requires Git and Windows PowerShell 5.1 or PowerShell 7. From a clean committed checkout, run:

```powershell
pwsh -NoProfile -NonInteractive -File .\scripts\Build-Release.ps1
```

Native Windows PowerShell can run the same script. The default parent is ignored `artifacts/releases`; each build creates a child named `UniversalPodcastDownloader-<version>-<full source commit>` containing the ZIP, manifest and checksum list. Use `-OutputDirectory <directory>` to choose another parent. Read `Get-Help .\scripts\Build-Release.ps1 -Full` for the exact interface. Version and inventory are declared in `tools/release-package.json`. The builder exports Git objects, not mutable working-tree content, and refuses dirty source or unsafe output locations. Existing candidates are never overwritten. Use a separate output parent for each repeat build. Its result reports the actual generated paths.

Only the declared runtime, LICENSE, README, CHANGELOG, this guide and original artwork are shipped. Development modules/scripts, tests, fixtures, handoff history, Git internals, logs, downloaded media, saved configuration, archive history and resume sidecars are excluded. Stable and rc builds use the same source, containment and inventory checks. A local build neither creates a tag/release nor uploads files.

## Switch or roll back application code

Keep the earlier ZIP, manifest and checksum list together. Verify and extract it into a separate new directory; select that directory's entry point to switch application code. Retain the current candidate as a separate copy. Do not overlay runtime files from different versions.

Application-code rollback does not roll back archive/configuration formats or bytes. Preserve media, `.upd` history/backups/checkpoints, partials and resume sidecars before trying an earlier version with explicit copied synthetic data. Older code may refuse newer schemas. Never delete state to force acceptance or blindly restore metadata over a live archive. Legacy rollback in the README is a reviewed metadata operation, not restoration of media or software. Saved feed credentials depend on the Windows user/profile and need explicit re-saving when moved.

Candidate CI validates the tested immutable source and stores only the ZIP, manifest and checksums for review; it creates no tag or GitHub release, including for stable-version builds. Release notes and authorization records are maintained under `docs/releases` and are outside the portable runtime. The owner explicitly requested v1.0.0 publication on 8 October 2026. Website assets and deployment remain a later project.
