# UPD-0502 — human Windows launcher and Ctrl+C observations

Observed 4 October 2026 (Europe/Berlin). The human completed the Explorer launcher flow and physical keyboard Ctrl+C tests in both supported engines. These are actual Windows Sandbox guest observations, with preserved synthetic archives and separate human confirmations.

The observed candidate is source **ed3cd45ddbc5f3d485ad29062420b11d3fc02fd8**, tree **86cbb1c5907aa37b5db1af56bbbfd7913fe326af**. Independent immutable Git/package audits establish equality of all **29 runtime files and 36 exported payloads** with tested **6a34eb60486e30e7a2631dc0bdbb4d69e643c2b9**; runtime fingerprint **709367285219c0ac0f4f1b8be2d346d6b91d5d6885b3b3afc27372352f745ca1**. Original full local All1791/analysis116 retain that tested6a identity. The new A041 Explorer and A045 physical Ctrl+C observations belong to ed3; no historical run is relabelled.

The [committed observation metadata](UPD-0502-MANUAL.json) records candidate identities, raw-report hashes, human outcomes, native lease fields, file inventories and the retained recovery limitation. Actual reports and media snapshots are preserved in the marked host directory `D:\projects\UPD manual ä-122202ec\Sandbox observations`; an additional raw JSON copy is ignored `.dev-tools/upd0502-manual-records`. They survive closing Sandbox. The compact record contains synthetic guest data, not private subscriptions or the real archive.

## Exact observed package

| Field | Verified value |
| --- | --- |
| Version / status | `0.1.0-rc.1` / `UNRELEASED_CANDIDATE` |
| Exact-source CI | [37180637487](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37180637487), attempt 1, all three jobs successful |
| Artifact | [11295685474](https://github.com/PikkuJanne/UniversalPodcastDownloader/actions/runs/37180637487/artifacts/11295685474) |
| Artifact name | `candidate-ed3cd45ddbc5f3d485ad29062420b11d3fc02fd8-37180637487-1` |
| Outer artifact | 6,217,665 bytes; SHA256 `79be50ff76513afb8c4faf06a4985cdfdcf136ac3e823d09d1edc9aee54903a3` |
| Portable ZIP | SHA256 `6ddc6c7e3a2b315ada19cd7c298d4c02284bb6177948b55fc2918b6e717d3ff2` |
| Manifest | SHA256 `df4ca4eac0b242ac5e75f1281abd8e46513bbb83957df631b397bd03fc1a5c26` |
| SHA256SUMS | SHA256 `5f6014ea657b251970caded740cf4250cfde059f6b57463bc37da98d87c80a29` |
| Retention | Expires 11 October 2026 at 08:00:23 Europe/Berlin (06:00:23 UTC) |

This ZIP was independently downloaded/audited before the manual session: API/upload/download identity, all 36 immutable Git payloads, 37 ZIP entries, original MIT bytes, embedded/sidecar manifests and checksums agreed. Hosted PS5.1.20348.5622 and PS7 7.6.6 each passed All1862 with failed/skipped/inconclusive/not_run0 and analysis118 with parse/new/baseline0. Those hosted observations remain distinct from the guest manual cases below.

## Isolation and environment

The preliminary host Explorer environment sentinel returned STOP: Explorer reused the ordinary profile. No downloader was run through that attempt. The subsequent Windows Sandbox configuration mapped only owned inputs and observations plus a read-only development Python directory. External networking and device redirection were disabled; the trusted fixture bound guest loopback. The host profile/archive was not mapped and no account, registry or machine-wide power/execution-policy setting was changed.

Guest account `WDAGUtilityAccount`, profile `C:\Users\WDAGUtilityAccount`, OS `Microsoft Windows NT 10.0.26100.0`. Native Windows PowerShell **5.1.26100.9444/Desktop** and portable PowerShell **7.6.5/Core** used actual interactive ConsoleHosts without redirected output. All 658 portable PS7 runtime files were compared against the bundled runtime before use. Python **3.12.14** is only the development fixture/recording runtime and is not shipped or required by the downloader.

Every guest extraction compared source/tree, manifest and all 36 payload sizes/hashes before execution. Product bytes were unchanged. Test helpers/server/runtime were outside the extracted package. The original argument-free batch defaults to the isolated guest Downloads archive. Callable interruption runs use fresh marked output directories explicitly.

## A041 — actual Explorer flow

In guest Explorer at `C:\UPD manual ä\Podcast test ü`, the human followed the instruction to double-click the unchanged `UniversalPodcastDownloader.bat` directly without arguments. At its guided feed prompt they supplied synthetic `http://127.0.0.1:49670/feeds/single.xml`, then pressed Enter at the count prompt for one episode. The reported summary was **Downloaded1 / Skipped0 / Adopted0 / Failed0**, followed by `[OK] Done. Files are in the selected podcast archive.`

The console stayed at `Press Enter to close`. The human subsequently reported **“Original downloader closed with Enter.”** This supplies the Explorer/association event, guided interaction, retained console and normal closing observation that terminal-invocation component checks did not supply.

The first process-inspection collector returned `Access is denied` before retaining process metadata. Its exact failing API was not recorded; no admin/permission change or product relaunch was used. A replacement file-only collector preserved the completed archive while the closing prompt remained open. **The numeric exit code of this Explorer-launched process is unavailable/null.** Historical unchanged-payload batch forwarding/quiet/missing-script/numeric-exit observations remain separate evidence; no fresh numeric0 is invented here.

File checks verified both pre/post package equality, snapshot byte equality, exclusive reopening of all six archive files, exactly one finalized synthetic MP3 (**1,689 bytes**, SHA256 **fcf79a0b8481d71b4f20c4dcb27f68f262ee2f5455e5f07d8f9fa1a3edabdd33**), matching `transfer_verified` history and no active partial/checkpoint. No full audio decoding, accessibility or physical display approval is inferred.

## A045 — actual keyboard interruption and native restoration

Each fresh interactive guest console dot-sourced the unchanged package. Observation-only forwarding wrappers retained the actual returned keep-awake lease and progress context; they used the real default native Windows API and did not inject cancellation or call cleanup. The human ran this command by itself:

```powershell
$global:A045Probe.RunOutput = @(Invoke-PodcastRun -FeedUrl $global:A045Probe.Feed -OutputPath $global:A045Probe.Archive -Mode All -KeepAwake -MaxAttempts 1 -IdleTimeoutSeconds 30 -RetryBudgetSeconds 0)
```

While Receiving showed nonzero bytes, the human physically pressed Ctrl+C once. After the PS prompt returned, they entered the separate observation command in that same console:

```powershell
. 'C:\UPD-Inputs\Audit-A045.ps1'
```

For **each engine**, the human separately confirmed **“Yes, I pressed Ctrl+C during Receiving.”** The audit first captured native lease fields before any manual cleanup; it never invoked Stop, cancellation or restoration. Both original callable result arrays were empty after the host interruption; neither numeric130 nor a process exit is claimed.

| Post-interruption observation | Windows PS5.1 | PS7 |
| --- | --- | --- |
| Actual engine | 5.1.26100.9444/Desktop | 7.6.5/Core |
| Audit UTC | 2026-10-04T11:31:58.6690866Z | 2026-10-04T11:44:50.943741Z |
| Activation/restoration thread | 6576 / 6576 | 4688 / 4688 |
| Native activation / prior state | true / `0x80000000` | true / `0x80000000` |
| Active / restored / worker alive afterward | false / true / false | false / true / false |
| Received and retained bytes | 483,328 of 2,026,800 | 139,264 of 2,026,800 |
| Closed progress / reopened all file handles | true / true | true / true |
| Final MP3 / verified history entries | absent / 0 | absent / 0 |
| Durable checkpoint offset | 65,536 | 65,536 |
| Whole partial matches checkpoint | false | false |

The test-only copied loopback fixture repeats the known synthetic MP3 1,200 times, advertises 2,026,800 bytes and a strong ETag, sends 65,536 bytes initially and then 8,192 bytes every 0.5 seconds. This keeps real .NET reads responsive and gives the human time to press Ctrl+C. The original tracked fixture was not changed. Original fixture SHA256 **75fd3d7ddb8cda5b9bd22268a69b95602339fe91303c4e17a3bd0d858ee061e5**; modified test-only copy **f00fd684c83248702e150d824f2c733cbc08692584ab78b28866b86406f15e99**; retained diff **04457f60fe584989c8ebdf73f5e5a72a82734ed35964b95a2508f023a3d21b93**.

All six preserved files per engine match their actual reported sizes/hashes. Both complete partials match the independently generated trusted fixture prefix. Their first65,536 bytes match the durable checkpoint SHA256 **48fad4d77558d80fe0cceff7483aa37ce09ce271aa69e5dd667215701c4d00f9**. Full PS5.1 partial SHA256 **58b68848422ae850e5bd087f852cc889cdb15c90cf4fd4941fe207c71fabb373**; PS7 partial SHA256 **dd1ffdbbf370bf6b42d0b59c166bba7e38afe444d44935eaf0291c032631047e**.

**The interrupted files are not automatically resumable.** Their saved checkpoints cover an earlier prefix; the larger uncheckpointed tails are preserved for review. The existing [resume policy](../RESUME_POLICY.md) and [strict opening check](../../../src/ResumeStore.ps1) reject whole-file length/hash mismatches and never truncate or guess an offset. Physical Ctrl+C catchability/final checkpointing was already qualified by that policy. The collector's stronger `matchingNonzeroCheckpoint` criterion remains false; no report was changed to manufacture a resumable file. A045's actual opted-in native restoration, progress/handle cleanup, retained bytes and lack of false completion are observed; successful automatic resume is not claimed.

The retained `$Error` entries for nonexistent prospective media/temporary paths are caught existence probes; `auditErrors` is empty. The absent callable result/final checkpoint is consistent with a ConsoleHost pipeline abort, but that control-path explanation is an inference. Native effects were measured **inside the Windows guest only**. Physical host sleep/lid prevention, hard kill, power loss, arbitrary native cleanup failures, live publisher behaviour and UNC durability remain outside these observations.

## Acceptance boundary

A041's literal Explorer condition and A045's actual physical Ctrl+C/native cleanup conditions are now observed without a waiver or policy change. Historical evidence remains source-qualified. The [UPD-0502 handoff](UPD-0502.md), canonical register and reviewed matrix reconcile the current result; A060 still requires separate explicit action-specific authority and verified outcome. No merge, tag, release or deployment is authorized by these tests.
