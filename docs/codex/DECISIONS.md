# Initial design decisions

These are proposed project defaults for this improvement project. Record any necessary deviation with reason and evidence; do not silently replace the owner's workflow.

| ID | Decision | Rationale |
|---|---|---|
| D001 | Improve existing PowerShell tool; no rewrite. | Preserve a successful simple workflow. |
| D002 | Windows PowerShell 5.1 and PowerShell 7 on Windows. | Existing advertised compatibility remains a requirement. |
| D003 | No new runtime third-party dependencies. | Keep double-click installation approachable. Development test tools are separate. |
| D004 | One task per Codex thread; one feature branch per milestone. | Keep context bounded and GitHub continuation reliable. |
| D005 | Local active-machine testing is authoritative; CI is supplementary. | Windows behavior must be observed, not inferred from Linux helper tests. |
| D006 | Unknown legacy files are unverified, never silently removed/replaced. | Protect the current archive. |
| D007 | Local versioned JSON first; no database server. | Sufficient for modest single-user archives if locking/atomic update is correct. |
| D008 | Sequential downloads by default. | Avoid connection pressure and needless concurrency complexity. |
| D009 | Hashes help identity/integrity checking but do not prove publisher authenticity. | Do not overstate what local hashes establish. |
| D010 | Preview may read feed/HTML but does not request media or persist anything. | Clear safe preview contract. |
| D011 | CustomCount alone implies Custom; contradictory explicit Mode rejected. | Remove current surprising parameter behavior. |
| D012 | Exit codes: 0 success; 1 setup/fatal; 2 incomplete; 130 catchable cancellation. | Make automation interpret outcomes. Module helpers return objects/errors rather than exiting the host. |
| D013 | HTTP response framing and RSS enclosure length have different authority. | Metadata estimates can be stale; avoid false corruption claims. |
| D014 | Website means product/distribution content, not hosted downloading. | No server storage, hosting selection or unrelated-site changes. |
| D015 | Release tags/publication, merges/settings and paid signing remain owner-gated. | Review and explicit publication control. |

## Future decision entry

ID; date; task; problem; considered alternatives; choice; compatibility/security/archive implications; evidence; owner approval where needed.
