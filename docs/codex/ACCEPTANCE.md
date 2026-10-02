# Acceptance matrix

Canonical statuses are in ACCEPTANCE_CASES.json. A001-A036 passed; A037-A060 remain NOT RUN. Helper results and historical defect characterizations are separate from downloader acceptance. See [UPD-0205 evidence](evidence/UPD-0205.md).

| ID | Task | Type | Status | Expected result |
|---|---|---|---|---|
| A001 | UPD-0001 | workflow | passed | Correct repository proven; unrelated changes preserved; baseline drift recorded; no automatic reset or stash. |
| A002 | UPD-0001 | workflow | passed | Only new handoff files copied; unchanged files recognized; conflicting files cause no overwrite; repeat application is idempotent. |
| A003 | UPD-0001 | workflow | passed | Remote branch SHA equals local HEAD; draft PR created/reused or an authentication/permission blocker is reported honestly. |
| A004 | UPD-0002 | unit | passed | No prompts, network calls, directories, downloads or preference mutations occur. |
| A005 | UPD-0002 | workflow | passed | Versions and engine are recorded; missing test tools produce a clear failure, not an empty green result; no runtime dependency added. |
| A006 | UPD-0101 | unit | passed | Selections are arrays; valid counts and progress; empty input is handled explicitly; no null/divide-by-zero errors on either Windows engine. |
| A007 | UPD-0101 | integration | passed | No legacy DOM parsing or security confirmation prompt; existing system security policy is not changed. |
| A008 | UPD-0102 | unit | passed | Valid deterministic components; identifiers retained after shortening; Windows limits handled clearly. |
| A009 | UPD-0102 | integration | passed | No escape from output root, no metadata-controlled drive/UNC path and no silent traversal through unsafe reparse points. |
| A010 | UPD-0102 | unit | passed | No accidental skip or overwrite; case-insensitive collisions also handled. |
| A011 | UPD-0103 | integration | passed | Only an owned partial file exists after interruption; final name absent until completion; rerun never counts partial as complete. |
| A012 | UPD-0103 | integration | passed | Invalid bodies fail; valid EOF-terminated audio can succeed; RSS enclosure length mismatch is advisory, not by itself proof of corruption. |
| A013 | UPD-0103 | integration | passed | No existing file is overwritten; temporary file is on destination volume; all streams close on failure. |
| A014 | UPD-0104 | unit | passed | Identity remains stable where evidence permits; exact media request URLs are not destructively normalized; distinct feeds cannot collide by title. |
| A015 | UPD-0104 | integration | passed | Previous valid state retained; reconciliation is conservative; corrupted history is preserved and reported, not silently reset. |
| A016 | UPD-0104 | integration | passed | History is checked against disk; missing/changed media is not blindly skipped; no unknown file is deleted. |
| A017 | UPD-0105 | integration | passed | Existing file hashes remain unchanged; verified/adopted/unverified/conflict states are distinguishable; no blanket redownload or deletion. |
| A018 | UPD-0105 | integration | passed | No claim of cryptographically verified completeness; adoption requires an explicit policy/owner choice; remote content changes do not trigger overwrites. |
| A019 | UPD-0106 | integration | passed | A startup diagnostic exists in a writable fallback location or stderr reports why logging is unavailable; original failure is preserved. |
| A020 | UPD-0106 | unit | passed | Shareable output contains no raw secret-bearing URL/header; use host plus opaque ID by default; runtime state containing secrets is excluded from exports. |
| A021 | UPD-0106 | integration | passed | Consistent UTF-8, unique run IDs; log failure never replaces the primary error; best-effort console fallback is clear. |
| A022 | UPD-0107 | integration | passed | Feed-only reads allowed as documented; no media request, folder/file/history/config/log/lock creation or keep-awake mutation; no subsequent false completion. |
| A023 | UPD-0107 | unit | passed | Only policy-approved HTTP(S) requests; TLS verification preserved; cross-origin credentials not forwarded; private local tests remain possible. |
| A024 | UPD-0107 | integration | passed | DTD rejected before expansion; no entity HTTP/file access; bounded response and parser resources; valid feeds parse. |
| A025 | UPD-0201 | integration | passed | Attempts bounded; transient retries differ from permanent failures; Retry-After does not cause an early retry; wait beyond configured budget becomes deferred/failed. |
| A026 | UPD-0201 | integration | passed | Connection/header and idle-body limits work; active long downloads are not killed by an inappropriate short overall timeout. |
| A027 | UPD-0202 | integration | passed | Result equals original media byte-for-byte; offset, total and validator checked; state belongs to this transfer. |
| A028 | UPD-0202 | integration | passed | Never blindly append; safely restart with explicit logging or leave partial for review. |
| A029 | UPD-0202 | integration | passed | No false completion, cross-episode concatenation or unsafe deletion; validated recovery/fresh transfer only. |
| A030 | UPD-0203 | integration | passed | Shared resolution rules produce same candidates; already fetched feed response reused. |
| A031 | UPD-0203 | unit | passed | Correct final base URI; all candidates deduplicated; interactive choice or clear deterministic/noninteractive policy. |
| A032 | UPD-0203 | unit | passed | Root-aware classification and useful distinct diagnostics; substring matches alone do not prove a feed. |
| A033 | UPD-0204 | unit | passed | Culture-independent chronological ordering; deterministic ties/missing dates; documented filename-date convention. |
| A034 | UPD-0204 | unit | passed | Use published when available; updated only a fallback; Atom ID collected. |
| A035 | UPD-0205 | unit | passed | Deterministic supported-audio choice; correct extension; no transcode; unsupported/ambiguous content explained. |
| A036 | UPD-0205 | integration | passed | Header alone neither guarantees validity nor rejects all generic audio; bounded sniffing, no promise of full decode/integrity validation. |
| A037 | UPD-0206 | integration | not_run | Deduplicated accessible entries only; cycle/budget stop visible; no silent claim of complete historical catalogue. |
| A038 | UPD-0206 | integration | not_run | Documented bounded selection; page-order assumptions explicit; retrieval gap yields incomplete result rather than false success. |
| A039 | UPD-0301 | integration | not_run | Infer Custom when count alone; reject conflicting explicit mode/count; fail without Read-Host when noninteractive. |
| A040 | UPD-0301 | integration | not_run | Documented 0/2/1/130 outcomes where cancellation is catchable; launcher preserves code; no [OK] after partial failure; module does not exit host. |
| A041 | UPD-0301 | manual | not_run | Simple interactive flow retained; arguments not evaluated as code; quiet/noninteractive path does not pause; missing script error visible. |
| A042 | UPD-0302 | integration | not_run | Exclusive writer protection for same archive; safe lock release/recovery; no crash-corrupt manifest; independent destination policy documented. |
| A043 | UPD-0302 | integration | not_run | Useful early errors; no existing files changed; handles closed; partial state preserved safely; no unknown PID killed. |
| A044 | UPD-0303 | manual | not_run | Progress never reports misleading 100% before success; episode/byte information honest; readable noninteractive summary. |
| A045 | UPD-0303 | manual | not_run | Off by default; temporary and restored in finally; no machine-wide power setting change; limitations on hard kill documented. |
| A046 | UPD-0304 | integration | not_run | Versioned validated config with safe permissions; clear failure on malformed input; no execution of strings; no plaintext secrets exported. |
| A047 | UPD-0304 | integration | not_run | Sequential bounded downloads, failure isolation, combined honest counts/exit; WhatIf does not mutate any show. |
| A048 | UPD-0401 | workflow | not_run | No new runtime packages, syntax incompatible with 5.1, destructive defaults, duplicated untested engines or global state leaks. |
| A049 | UPD-0402 | manual | not_run | Commands work, functions/parameters exist, capabilities match tests, no promised unreleased feature or machine-wide security bypass. |
| A050 | UPD-0402 | manual | not_run | All scope is explicit; legacy unverified files not presented as verified; repair procedure avoids deleting originals. |
| A051 | UPD-0403 | workflow | not_run | Deterministic file inventory and version; checksums validate; includes imported modules/license/help; excludes logs, media, secrets, test fixtures and handoff history. |
| A052 | UPD-0403 | manual | not_run | No hidden checkout dependency; source version traceable; original MIT license retained. |
| A053 | UPD-0404 | workflow | not_run | Minimal job permissions, reviewed SHA-pinned actions, no pull_request_target execution of untrusted code, secret-safe artifacts; no automatic public release. |
| A054 | UPD-0404 | workflow | not_run | Draft/artifacts ready; no tag push, release publication, merge, branch deletion, visibility/settings change or purchase occurs. |
| A055 | UPD-0405 | manual | not_run | Current-version placeholder not falsely advertised as released; features/limitations/privacy/license/support accurate; real UI not generated imagery. |
| A056 | UPD-0405 | workflow | not_run | Portable static content/metadata only; no domain, hosting, DNS or live site changes. |
| A057 | UPD-0501 | manual | not_run | All mandatory cases have concrete evidence; user archive untouched; unrun/manual cases remain clearly outstanding. |
| A058 | UPD-0501 | workflow | not_run | Gate fails; bundle/server validation never substitutes for downloader acceptance. |
| A059 | UPD-0502 | workflow | not_run | Report ready-for-owner-review, remaining approvals and exact branch/PR/artifact/version identifiers; leave draft private/unpublished as applicable. |
| A060 | UPD-0502 | workflow | not_run | No inferred authority for other destructive/publication operations; audit trail records scope and resulting identifiers. |
