# Parity matrix — .NET XactCopy vs native C++ port

Audit date: 2026-07-23 (Phase 4e landed). Status values: **Done** (behavioral
parity, validated), **Partial** (subset ported, gap listed), **Missing** (not
started).

Phase 4e delivered: JobCatalogStore + models (cross-compat both ways), the
JobManagerService, the Job Manager dialog + TextPromptDialog, main-window jobs
wiring (menus un-greyed, run lifecycle, queue drain, startup interrupt marking),
Explorer context-menu registration + `--from-explorer[-folder]` launch args
(source-folder and selected-items modes), the missing main-window controls
(continue/skip-known-bad/wait-media/fragile checks; buffer/retries/timeout
numerics; throughput/ETA/buffer-usage/rescue/job-summary/journal telemetry
labels; diagnostics strip gated on `UiShowDiagnostics`).

## 1. Core + IPC

| Feature | Status | Notes |
| --- | --- | --- |
| Models/enums (CopyJobOptions/Result/ProgressSnapshot, all enums) | Done | Byte-parity via goldens |
| IPC envelope + all 11 messages | Done | Golden + InteropProbe A–I |
| Named-pipe framing, CurrentUserOnly DACL, 16 MB cap | Done | |
| STJ serializer parity (escaping, times, indented mode) | Done | Golden files from real STJ |
| Heartbeats, telemetry throttling, log rate caps | Done | Exact log strings asserted |

## 2. Storage

| Feature | Status | Notes |
| --- | --- | --- |
| JobJournalStore (snapshots, backups, mirrors, ledger, anchors, HMAC) | Done | Bidirectional CrossTests |
| BadRangeMapStore (envelope, DPAPI key, legacy fallback) | Done | Bidirectional CrossTests |
| Tamper trust fallback semantics | Done | Both directions |
| JobCatalogStore (`jobs\catalog.json`) | Done | Plain STJ indented + string enums, `.tmp` atomic move, normalization + legacy-queue migration + queue-entry dedup, schema v2. Cross-compat both ways (`-CrossTests`); reads the real .NET catalog live in the Job Manager |

## 3. Worker engine

| Feature | Status | Notes |
| --- | --- | --- |
| Resumable copy, rescue passes, salvage, retries/policies | Done | InteropProbe byte-exact |
| ScanOnly precise + fast (parallel, precise fallback) | Done | |
| Bad-range-map integration (seed/skip/persist) | Done | .NET store loads native maps |
| Parallel small-file phase, adaptive buffers, CopyFileEx path | Done | |
| Verification (full + sampled, sha256/512) | Done | |
| Fault injector (env rules) | Done | Same rule grammar |
| Raw-disk experimental scan backend | Partial | Degrades to standard reads with log — identical to .NET non-elevated behavior; elevated raw path not ported (accepted niche) |

## 4. UI infrastructure

| Feature | Status | Notes |
| --- | --- | --- |
| WorkerSupervisor (heartbeat, stall formula, auto-recovery) | Done | Kill-mid-job test resumes byte-exact |
| RecoveryService/StateStore + RunOnce autostart | Done | .NET-compatible recovery-state.json |
| AppSettings (DOM-preserving settings.json) | Done | Unknown keys survive |
| Theming (dark/light, accent, menus, UAH bar, owner-drawn controls) | Done | |
| ComCtl32 v6 + PMv2 manifest, WM_DPICHANGED | Done | |
| Single-instance guard | Partial | Surfaces existing window; does **not** forward launch args (`HandleForwardedLaunch`) to the running instance |
| Settings dialog (8 pages, ~65 fields) | Done | DOM-preserving save; beautified nav rail (panel sidebar + accent selection bar) |
| About dialog | Done | Axiom-style: app icon, title, version/build/author/licence, components, Check-for-Updates + OK |
| RecoveryPromptForm | Partial | Native uses a MessageBox (Resume/Discard/Later equivalent); no themed form with run details |
| TextPromptDialog | Done | Themed one-line prompt (Save-As-Job / Rename / Duplicate) |
| TaskbarProgressController (ITaskbarList3) | Done | Normal/Paused/Error states wired into the job lifecycle (`taskbar_progress.h`) |
| TooltipScenarioFormatter / tooltips | Missing | |
| UiAppearanceManager consumption | Partial | Theme/accent/log colorize/log font + `UiShowDiagnostics` consumed; density, scale %, chrome, grid lines, progress style stored but not applied |
| WindowChromeManager / WindowIconHelper | Partial | Dark titlebar done; custom chrome/icon helpers not ported |

## 5. Main window controls vs .NET MainForm

| .NET control | Native | Notes |
| --- | --- | --- |
| Source/destination pickers | Done | |
| Mode / engine / overwrite selections | Done | Combos |
| Verification | Done | Native combo (VB: checkbox + settings mode) |
| Salvage / resume / map / adaptive checks | Done | |
| Continue-on-error check | Done | |
| Skip-known-bad-ranges check | Done | Gated on Use-bad-range-map like .NET |
| Wait-for-media check | Done | |
| Fragile-mode check | Done | |
| Buffer MB / retries / timeout numerics | Done | Seeded from settings; overrideable per run |
| Scan Bad Blocks button | Done | Via Tools menu |
| Start/pause/resume/cancel | Done | Buttons + menu |
| Overall + current progress bars | Done | |
| Overall stats / current file labels | Done | |
| Bytes label | Done | In stats |
| Throughput label (EWMA 0.65/0.35 + avg) | Done | |
| ETA label (smoothed-speed estimate) | Done | Same reference-speed selection + format |
| Buffer usage label | Done | Configured size + average utilization |
| Rescue telemetry label | Done | Pass / regions / remaining bytes |
| Job summary label | Done | |
| Journal path label | Done | Deterministic path on start, result path on finish |
| Diagnostics strip (`UiShowDiagnostics`) | Done | Native counters (active/scan-workers/recovered/skipped); .NET's UI-render timings are not applicable |
| Log (severity colors, fonts) | Done | Listbox vs .NET ListView; virtualization N/A at native volumes |
| Progress animation (marquee during scan) | Missing | Cosmetic |
| Taskbar progress | Missing | |
| Crash-resume prompt | Done | Message-box form |
| Menu bar (full structure, state-aware) | Done | Save-As-Job / Job-Manager / Run-Next-Queued now wired; Check-Updates greyed (updater = phase 5) |

## 6. Jobs subsystem (Done — Phase 4e core)

| Feature | Status | Notes |
| --- | --- | --- |
| JobCatalog models | Done | ManagedJob, ManagedJobQueueEntry, ManagedJobRun (+ nullable LastAttemptUtc/FinishedUtc/Result), ManagedJobRunStatus (Queued/Running/Paused/Completed/Failed/Cancelled/Interrupted), JobCatalog schema v2 |
| JobManagerService | Done | save/rename/duplicate/delete, queue ops (enqueue/remove/move/clear/dequeue), run lifecycle (create ad-hoc + for-job, running/paused/resumed/completed/interrupted marks), history (recent runs, delete, clear). Headless-tested in `xactcopy_supervisor_tests` |
| Job Manager dialog | Done | Unified ListView grid (Type/Name/State/Queue/Trigger/Started/Updated/Source/Destination/Summary) with custom-draw state colors + dark header, view + run-status filters, search debounce, details pane, Run Now / Queue / Remove / Top / Up / Down / Bottom / Rename / Duplicate / Delete Job / Delete Run / Open Journal / Clear Queue / Clear History / Refresh, 3 s auto-refresh, double-click + Del/F5/Enter keys |
| MainForm integration | Done | Ad-hoc run per manual copy/scan ("Manual Copy"/"manual", "Bad Block Scan"/"scan"), mark_run_running w/ deterministic journal path, paused/resumed, completed w/ result, mark-any-running-interrupted at startup + interrupt on exit-mid-run, resume-interrupted run, queued-auto drain after completion, `AutoRunQueuedJobsOnStartup` |

## 7. Explorer integration (Done)

HKCU `Software\Classes\{Directory,Drive,*,Directory\Background}\shell\XactCopy`:
caption "Copy with XactCopy", `Icon` = exe, `MultiSelectModel=Player` (except
Background), command `"exe" --from-explorer "%1" %*` (Background:
`--from-explorer-folder "%V"`). `ExplorerIntegrationService` registers/unregisters
byte-identically to the .NET service; synced on startup + settings save from
`EnableExplorerContextMenu` (only touches the registry when state differs).
Launch args `--from-explorer[=]` / `--from-explorer-folder[=]` parsed into
`LaunchOptions`; `apply_explorer_launch_options` sets the source box
(source-folder mode) or computes the common root + relative selection set
(selected-items mode → `SelectedRelativePaths`), then prompts for a destination.
Remaining: forwarding args to an already-running instance (single-instance
guard still just surfaces the window).

## 8. Updates (Done — check + download + verify + apply)

`update_service.h`: WinHTTP GET of the GitHub releases API (`UpdateReleaseUrl`,
default `.../repos/Wimukthi/XactCopy/releases/latest`), JSON parse (incl.
`assets[]` with name/browser_download_url/size/`digest`), version compare vs
`kNativeVersion`. `select_best_asset` (win/arch/zip/exe scoring),
`download_asset` (WinHTTP streamed to file + progress callback + streaming
SHA-256), `resolve_asset_sha256` (asset digest first, then checksum sidecar
assets — `.sha256`/`.sha256sum`/`checksums.txt`, `SHA256 (file) = hash` styles),
`format_bytes`. Check-Updates menu wired + `CheckUpdatesOnLaunch` silent startup
check.

`update_dialog.h`: themed `UpdateDialog` on the shared ModalHost (resizable,
dark title bar, owner-drawn buttons / progress bar / notes list) — release
summary (current/latest/package), markdown-ish release notes (word-wrapped,
headings bold, bullets), Download && Install / View Release / Close(Cancel).
Download runs on a background thread posting progress + done messages; verifies
SHA-256 before applying. Apply flow mirrors UpdateForm: `.exe` asset →
`ShellExecuteW` + app exit; `.zip` asset → writes `apply-update.ps1` (wait for
PID exit → `ZipFile.ExtractToDirectory` → resolve payload dir → robocopy
backup/copy with restore-on-failure → relaunch → cleanup), launched via
`ShellExecuteW` so it survives app exit; then `on_update_done` posts `WM_CLOSE`
to exit. Verified live against GitHub (check + asset selection). Remaining:
channel/prerelease filtering (accepted niche).

## Remaining after Phase 4e

Deferred to Phase 5 (updater + packaging) or accepted as cosmetic:

- UpdateService + Update dialog (check + download + verify + apply) — **Done**.
- Taskbar progress (ITaskbarList3 Normal/Paused/Error) — **Done**.
- App icon on the exe + every window/dialog title bar (`app_icon.h`, `app.rc`
  IDI_XACTCOPY + VERSIONINFO) — **Done**. Exe is now `XactCopy.exe`.
- Axiom-style About dialog + beautified Settings nav rail — **Done**.
- Comprehensive shell integration (`explorer_integration.h`): copy verbs, a
  "Scan for Bad Blocks" verb (Drive/Directory, `--scan-from-explorer`),
  Applications\XactCopy.exe (Open-with + drag-drop), App Paths, Send-To
  shortcut — **Done** (register/unregister round-trip verified).
- Inno Setup release pipeline (`installer\XactCopy.iss` +
  `installer\build-installer.ps1`, mirrors AxiomCompress) — **Done**
  (produces `XactCopySetup-<ver>-win-x64.exe`).
- Progress marquee animation during scan — cosmetic.
- Themed RecoveryPromptForm (native uses a MessageBox today).
- Forwarding launch args to an already-running instance.
- Appearance settings still stored-but-unapplied: density, window chrome style,
  grid lines.
- Tooltips / TooltipScenarioFormatter.
