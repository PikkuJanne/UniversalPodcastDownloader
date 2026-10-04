# Optional saved shows and sequential batch

Implemented by UPD-0304 on 3 October 2026. Windows PowerShell 5.1 and PowerShell 7 on Windows use the same bundled helpers; no application runtime package, daemon, scheduler or account service is required.

## Commands

The normal argument-free guidance and direct FeedUrl workflow do not access saved configuration. Saving a show stores settings without fetching metadata or media. Choose a short name, an HTTP(S) URL accepted by the existing transport policy, an absolute output root and Latest/All/Custom selection. URL user information is rejected; signed path/query data is preserved exactly inside the protected value.

```powershell
.\UniversalPodcastDownloader.ps1 -SaveShow news -FeedUrl 'https://example.invalid/feed.xml' -OutputPath 'D:\Podcasts' -CustomCount 5
.\UniversalPodcastDownloader.ps1 -ListShows
.\UniversalPodcastDownloader.ps1 -ShowName news -NonInteractive
.\UniversalPodcastDownloader.ps1 -Batch -NonInteractive
.\UniversalPodcastDownloader.ps1 -Batch -WhatIf
.\UniversalPodcastDownloader.ps1 -RemoveShow news -WhatIf
.\UniversalPodcastDownloader.ps1 -ExportShows 'D:\Share\shows-summary.json'
```

These examples require replacing the reserved URL/output path; the export parent must already exist. Save the same name again to replace that row's settings in its existing position. Names are unique ignoring case, start with a letter/digit, and contain 1-64 letters, digits, underscores or hyphens. At most 100 rows are supported. Removing a row preserves all archive media, history, checkpoints and logs.

Exactly one saved selector is allowed, except Batch may combine with ShowName. ConfigPath applies only to saved operations. Direct FeedUrl is allowed only for SaveShow, and saved operations reject legacy actions. Save accepts FeedUrl, OutputPath and selection options; List/Remove/Export reject download options. Named runs and Batch accept the runtime overrides below. NonInteractive rejects explicit enabled Confirm; Batch always makes each child noninteractive.

Batch runs all saved rows in stored order. A subset follows the explicit array order and rejects null/empty, repeated, unknown or more than 100 names before execution. In a PowerShell session:

```powershell
& .\UniversalPodcastDownloader.ps1 -Batch -ShowName @('news', 'other') -NonInteractive
. .\UniversalPodcastDownloader.ps1
$result = Invoke-PodcastCommand -Options @{ Batch=$true; ShowName=@('other','news'); WhatIf=$true }
```

The native Windows PowerShell -File boundary and batch launcher cannot pass a multi-element string array; use all shows, one name, or an in-process array. Microsoft documents this native boundary in [about_PowerShell_exe](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_powershell_exe?view=powershell-5.1).

## Strict schema and storage

The default path is `%LOCALAPPDATA%\UniversalPodcastDownloader\saved-shows\shows.json`. ConfigPath selects an absolute ordinary drive/UNC path, at most 240 characters, with a parent at most 210 characters to reserve lock/temporary room; use a new dedicated private parent directory. Device paths/names, alternate data streams, control/quote/wildcard characters, dot/dot-dot components and reparse configuration paths are refused. Stored output roots must be absolute supported Windows paths; execution retains all existing containment, reparse, lock and no-overwrite checks.

The internal schema is exactly `schema_version: 1` plus a `shows` array. Each row has exactly `name`, `feed_protected`, `output_path`, `mode`, `custom_count`. Custom requires a positive Int32 count; Latest/All require a null count. The reader bounds the file at 1 MiB, token count at 8,192 and nesting at eight, rejects invalid UTF-8, extra/duplicate/escaped property names, unexpected types, comments, trailing commas and unsupported versions. Strings remain data throughout; there is no evaluation or configuration-driven command execution. The private reader/getter exposes protected storage or the decrypted URL to in-process callers and is not a sharing/export interface.

Missing configuration is an empty selection. Malformed or newer configuration is preserved and fails clearly; there is no automatic reset, partial acceptance, migration or backup restoration. Do not manually edit ciphertext. To repair unreadable credentials for a particular row, explicitly save that name again with its URL while using the same Windows identity. A structural error must first be inspected by the owner.

Configuration directories and files are created with a protected DACL granting FullControl only to the current user's SID, owned by that SID. Existing configuration parents/files must match this policy; the application does not broaden access or silently repair permissions. The existing ancestors of a newly created private directory retain their ACLs. Creation uses the platform's ACL-aware APIs so sensitive files are not first created with a broad inherited DACL. [Microsoft ACL-aware creation API](https://learn.microsoft.com/en-us/dotnet/api/system.io.filesystemaclextensions.create)

Mutation holds a persistent `<config>.lock` through an exclusive file handle, re-reads and validates the current file under that lock, writes and flushes a newly owned protected sibling temporary, then uses same-volume File.Replace or no-overwrite Move for initial creation. Lock files are retained and their contents/PIDs are never authority. Only an owned safe temporary can be cleaned up. Atomic replacement requires filesystem support; UNC paths remain subject to share/filesystem/ACL availability and neither process tests nor flush establishes hardware power-loss durability. No configuration backup is exported or stored automatically. [Microsoft File.Replace](https://learn.microsoft.com/en-us/dotnet/api/system.io.file.replace)

## Protection and sharing limits

Every complete feed URL is protected using Windows DPAPI CurrentUser with fixed application/schema entropy. This covers public URLs as well as URLs containing signed path/query credentials. The Windows profile must be available; unavailable/corrupted protected data returns a fixed credential error without exposing the original bytes. Moving ciphertext to a different account/profile is not a supported credential transfer; re-save the URL there. PowerShell already supplies the built-in framework API; the downloader installs no cryptography package. [Microsoft DPAPI guidance](https://learn.microsoft.com/en-us/dotnet/standard/security/how-to-use-data-protection), [CurrentUser scope](https://learn.microsoft.com/en-us/dotnet/api/system.security.cryptography.dataprotectionscope)

DPAPI CurrentUser permits another process running as that same user to decrypt. The directory ACL and encryption do not protect against that user, same-user malware, administrator compromise or access to the running process. Decrypted URLs exist in process memory while settings and runs are used; immutable strings cannot be reliably erased. Shell input/history/transcripts may retain the original save command. Saved names and output paths remain plaintext local settings, and safe names can still disclose subscription choices. Existing diagnostic hostname/identity correlation limits continue to apply.

List, public management results and ExportShows contain only Name, Mode, CustomCount and FeedConfigured. Export omits URL, ciphertext, output path and unknown properties, uses a new protected file in an existing ordinary directory and never overwrites. Its schema-1 projection is deliberately not accepted as the storage schema; it is a reviewable subscription summary, not a restorable credential backup. Review names before sharing. A failure after export creation may leave its owned partial file; later export refuses that occupied path. Errors use fixed private guidance rather than raw path/credential exceptions.

## Overrides, batch outcomes and preview

Named and batch runs accept Mode, CustomCount, OutputPath, KeepAwake, MaxFeedPages, MaxAttempts, HeaderTimeoutSeconds, IdleTimeoutSeconds, RetryBudgetSeconds, BaseDelaySeconds and MaxDelaySeconds. Overrides apply only to that invocation; configuration is unchanged. A positive count alone changes the effective mode to Custom. Explicit Latest/All discards a stored Custom count. Explicit Custom uses the stored count only if the saved mode was Custom; otherwise supply a positive count. Explicit Latest/All together with a count is invalid. Transport bounds remain the existing single-show bounds. DiagnosticExportPath and arbitrary child/common options are not batch runtime overrides.

Each batch child finishes and attempts all resource cleanup before the next starts; a failing Dispose/native cleanup cannot guarantee release. A show with unavailable credentials, failed metadata, invalid media, conflicting files or unresolved catalogue does not erase successful work or prevent later shows. Overall exit 2 reports any isolated fatal/incomplete child; the summary distinguishes show failures from episode failures and does not invent episode outcomes for a feed that could not be read. Invalid configuration, selection or overrides stop batch setup with 1. A structural configuration error observed between children stops with 1 and preserves processed counts. Catchable cancellation stops with 130 and preserves completed work, recovery evidence and unstarted names. Empty configuration succeeds with zero selected shows. Forced termination and every physical Ctrl+C path cannot guarantee a result.

Batch result schema-1 includes show totals, combined Planned/Downloaded/VerifiedSkipped/LegacyUnverified/Conflicts/Failed/Deferred/Cancelled counts and safe child RunResults. Child projection omits LegacyResult and arbitrary properties. Unknown/adopted local media retains its established conflict/unverified behavior; batch does not adopt, roll back or overwrite it. See [CLI_RESULTS.md](CLI_RESULTS.md) for result arrays, precedence and entry-script PassThru behavior.

WhatIf may decrypt selected URLs, retrieve bounded feed/page metadata and read existing local evidence for every selected show. It requests no planned enclosure, writes no configuration/directories/locks/archive/logs/exports and makes no keep-awake request. Management WhatIf/declined confirmation also bypasses protection and creation work. A declined batch confirmation does not start/decrypt children and reports their unstarted names; preview 0 is not completed media. Invalid preview input can still fail with 1, and incomplete child planning keeps aggregate 2.
