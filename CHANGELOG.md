# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

## [1.0.8.5] - 2026-02-19

### Changed

- Completed a broad codebase documentation pass across hand-written source and tests:
  - Added consistent file headers and XML summaries for key types and members.
  - Improved maintainability and code navigation without altering runtime behavior.
- Recovery/rescue naming polish:
  - User-facing/runtime references now use `Rescue Engine` instead of `AegisRescueCore`.
  - Updated related UI tooltips/section labels and architecture/readme wording.

## [1.0.8.4] - 2026-02-19

### Fixed

- Supervisor recovery process containment hardening:
  - Added launched-worker PID tracking and orphan-worker cleanup on shutdown/recovery.
  - Recovery now aborts instead of spawning a new worker when the previous stuck worker cannot be terminated cleanly.
  - Prevents duplicate stuck-worker accumulation and repeated same-file restarts during severe media hangs.

## [1.0.8.3] - 2026-02-19

### Fixed

- Fragile-media restart/resume hardening for stalled files:
  - Journal entries left in `InProgress` after worker stall/recovery are now normalized before the next run pass.
  - In fragile mode (`Skip file on first read error`), stale in-progress entries are promoted to non-retry failed entries and skipped, preventing repeated hammering of the same unstable file after worker restarts.
  - Added explicit worker log signal when stale entries are promoted to fragile non-retry state.

## [1.0.8.2] - 2026-02-19

### Added

- Fragile-media hardening controls:
  - New `Fragile media mode` run option and defaults.
  - New `Skip file on first read error` behavior for ultra-unstable media.
  - New `Persist fragile skips across resume` option using journal-backed `do-not-retry` markers.
  - New fragile failure circuit-breaker controls (time window, threshold, cooldown).

### Changed

- Fragile mode now applies conservative runtime safety presets automatically:
  - Disables raw-disk experimental scan backend.
  - Forces retries to `0`.
  - Forces skip-on-first-read-error behavior.
  - Caps timeout to a short bounded range to avoid prolonged lockups on bad media.
- UI/settings/supervisor/recovery option cloning now carries fragile-mode policies end-to-end across new runs, resumed runs, and queued jobs.

### Fixed

- Repeated re-hammering of known-failing files on fragile disks:
  - Files that trip fragile first-read failure can be marked non-retry in the journal and skipped on resume.
  - Circuit-breaker cooldown now throttles repeated failure bursts to keep progression responsive under severe media instability.

## [1.0.8.1] - 2026-02-19

### Added

- Operations log severity color-coding:
  - Main operations log now classifies lines (info/success/warning/error/critical) and renders them with severity-aware colors.
  - New Appearance setting to enable/disable severity coloring (`Color-code log by severity`).

### Changed

- Destination metadata hot path now uses Win32 `GetFileAttributesExW` for existence/size/last-write checks used by overwrite/completed-file decisions, reducing per-file metadata overhead.
- Source directory scanning now prefers Win32 enumeration (`FindFirstFileEx`/`FindNextFile`) with managed fallback for unsupported cases.
- Rescue range state updates now mutate only affected segments with localized merge, reducing full-list rebuild churn on large range maps.

### Fixed

- Cancellation hang hardening:
  - Worker IPC sends are now frame-safe under cancellation (no mid-frame cancellation token abort), preventing malformed IPC payloads.
  - Supervisor now handles repeated malformed IPC frames as stream desync, recovers deterministically, and force-finalizes active runs instead of waiting indefinitely for cancellation acknowledgement.
  - When channel loss happens after user cancel, supervisor force-stops the worker and completes the run state as cancelled with a clear reason.

## [1.0.7.0] - 2026-02-19

### Changed

- Scan-mode small-file performance:
  - Added a dedicated `ScanSmallFast` path so tiny files in scan mode avoid the heavier multi-pass rescue pipeline when no persisted rescue coverage exists.
  - Retained rescue-state accounting/progress semantics so bad-range detection fidelity remains unchanged.

### Fixed

- Scan throughput regression on large small-file sets:
  - Bad-range map persistence is now batched during scan runs (interval-based) instead of forcing a full map write per file.
  - Final map flush remains durable at run end to preserve crash-safe behavior.

## [1.0.6.8] - 2026-02-18

### Fixed

- Operations log rendering hardening across consecutive jobs:
  - Reset now clears pending queued log messages and dispatch flags before a new run begins, preventing stale queued lines from corrupting first repaint after completion.
  - Virtual log view size updates now use a dedicated refresh path that forces safe redraw/invalidate for virtual list content.
  - Diagnostics row now refreshes immediately after log reset so telemetry state does not lag behind the visible log view.

## [1.0.6.7] - 2026-02-18

### Fixed

- Updater apply cleanup hardening:
  - Zip payload root resolution now targets the actual app payload directory to avoid leaving wrapper folders (for example `win-x64`) inside the install directory.
  - Apply script now removes stale wrapper folders created by previous update packages.
  - Post-apply cleanup now removes temporary updater staging/script files after restart handoff.
  - Cancel/failure close paths now perform best-effort cleanup of temporary updater directories.

## [1.0.6.5] - 2026-02-18

### Fixed

- Startup crash hardening:
  - `Program.Main` now configures `SetCompatibleTextRenderingDefault(False)` before color-mode initialization and guards the call with a safe fallback logger path.
  - Prevents `InvalidOperationException` (`SetCompatibleTextRenderingDefault must be called before the first IWin32Window object is created`) from terminating launch in edge restart/recovery paths.

## [1.0.6.3] - 2026-02-18

### Changed

- Codebase comment pass:
  - Added focused explanatory comments around operation-mode semantics, root-safe path normalization, clone/snapshot boundaries, and bad-range-map durability logic.
- Documentation refresh to match current implementation:
  - Rewrote `ARCHITECTURE.md` to document current runtime topology, scan mode, bad-range maps, integrity hardening, and worker/UI responsibilities.
  - Updated `README.md` highlights/build instructions and documentation links.
  - Updated `CONTRIBUTING.md` and `docs/RELEASE_PROCESS.md` with current contribution/release expectations.

## [1.0.6.2] - 2026-02-17

### Added

- Scan-only bad-block workflow:
  - Added explicit `ScanOnly` operation mode in shared job options/contracts.
  - Added `Scan Bad Blocks` start action on the main window for a read-only scan run.
  - Added persistent bad-range map model/storage (`BadRangeMap`, `BadRangeMapFileEntry`, `BadRangeMapStore`) so scan data can be reused by later copy runs.

### Changed

- Main UI and worker option pipelines now carry scan mode end-to-end (UI, supervisor, recovery, and job manager cloning paths).
- Path normalization now preserves drive roots (`D:\`) instead of trimming to drive-relative forms (`D:`), preventing accidental remap to current-working-directory paths.
- Running window title now reflects mode (`Scanning` vs `Copying`) during active runs.

### Fixed

- Scan-only runs no longer execute copy-only overwrite/completed-file skip checks, preventing false `Skipped (destination is newer or same age)` outcomes.
- Starting a scan-only run no longer writes normalized destination fallback back into the destination textbox.
- Selecting a drive-root source path now remains stable and does not drift to the executable/current directory path.

## [1.0.5.9] - 2026-02-17

### Added

- Journal durability and anti-corruption hardening:
  - Multi-generation snapshot backups (`.bak1`..`.bak3`) for each journal save.
  - Mirrored journal snapshots under `%LocalAppData%\XactCopy\journals-mirror`.
  - Append-only journal ledger (`.ledger`) with framed records and checksum validation.
  - Hash-chained, HMAC-signed ledger records and signed anchor metadata (`.anchor`) to detect tampering and rollback.

### Changed

- Journal load/recovery selection is now trust-aware and sequence-aware:
  - Prefers the newest signed snapshot recorded in trusted ledger history.
  - Falls back across primary, backup, and mirror candidates when corruption or loss is detected.
- Added journal seed-state caching in the storage layer to avoid repeated full-ledger scans during frequent save flushes.
- Journal normalization on load/save now enforces stable defaults for IDs/timestamps/file-entry maps before persistence decisions.

### Fixed

- Resume reliability is improved when the primary journal file is modified or partially corrupted: signed backup/mirror states are now used automatically.
- Legacy plain-JSON journals remain loadable when signed metadata does not yet exist, preserving backward compatibility.

## [1.0.5.8] - 2026-02-17

### Added

- Media identity guards for both source and destination paths:
  - Run-time options now carry expected media identities (`ExpectedSourceIdentity` / `ExpectedDestinationIdentity`).
  - New runs capture baseline identity (volume serial for local roots, normalized `\\server\share` identity for UNC roots) when available.
  - Worker availability checks now validate identity, preventing writes to a different device accidentally mounted on the same path/drive letter.

### Changed

- Supervisor recovery now forces `ResumeFromJournal = True` before re-dispatching a job after worker disconnect/crash, ensuring journal-based continuation semantics.
- Availability/error classification now uses Win32/HResult-aware checks (not only message text), improving handling reliability for localized Windows environments.
- Journal initialization now probes writability before run start and automatically falls back to `%TEMP%\XactCopy\journals` when the default journal directory is unavailable.
- Worker now emits destination free-space preflight warnings when estimated payload exceeds currently available space.
- Added remap-capable resume plumbing:
  - New run options `ResumeJournalPathHint` and `AllowJournalRootRemap` let resumed jobs load an existing journal and continue even after source/destination path remap.
  - Completion flow now offers a guided remap-and-resume prompt for media/path-style failures.
- Added explicit hardening policies for lock/AV contention and source mutation:
  - `WaitForFileLockRelease`
  - `TreatAccessDeniedAsContention`
  - `LockContentionProbeInterval`
  - `SourceMutationPolicy` (`FailFile`, `SkipFile`, `WaitForReappearance`)
- Performance settings now expose new default controls for the above policies.

### Safety

- Saved job templates now clear run-scoped media identity fields to avoid stale identity locks between unrelated runs.
- Write path now fails fast on non-recoverable destination errors (`disk full`, `access denied`) with explicit user-facing messages instead of exhausting generic retry loops.
- Read salvage now avoids masking media/path-unavailable conditions by escalating unavailability failures as hard read errors.
- Media identity matching is now remap-safe for local volumes by comparing volume serial identity (with compatibility for legacy identity strings that included drive letters).
- Source-file disappearance during active runs is now policy-driven rather than implicitly treated as generic media outage.

## [1.0.5.0] - 2026-02-16

### Changed

- Worker data path now uses handle-based `RandomAccess` for chunk reads and writes, reducing hot-path seek overhead and stream-position churn.
- Added adaptive rescue pass tuning that adjusts chunk size and retry aggressiveness by observed bad-region density.
- Added reverse-direction rescue sweeps (`TrimSweepReverse` and descending `RetryBad`) to improve recovery behavior on fragmented/borderline media.
- Added a dedicated small-file fast path that lowers per-file overhead while preserving retry, salvage, progress, and journal guarantees.

### Performance

- Improved throughput consistency and lower CPU/GC overhead during both high-speed healthy copies and recovery-heavy rescue scenarios.

## [1.0.4.9] - 2026-02-16

### Changed

- Worker copy pipeline hot path was optimized to reduce allocation and stream-open overhead:
  - Added pooled I/O buffers (`ArrayPool`) for rescue/salvage/verification read paths.
  - Added per-file transfer sessions to reuse source and destination `FileStream` instances across chunk operations.
  - Refactored chunk read/write retry paths to reuse shared buffers and invalidate/reopen streams only on media/retry failures.

### Performance

- Improved sustained throughput consistency and reduced GC pressure during long-running copy jobs, especially for large datasets and rescue passes.

## [1.0.4.8] - 2026-02-15

### Changed

- Settings dialog open/render path was optimized to reduce first-frame layout churn (cached property reflection, debounced dirty-state recompute, faster page switching, and broader double-buffer usage).
- Job Manager grid pipeline was reworked for scale and responsiveness:
  - Switched to `DataGridView` virtual mode with a cached row model (`CellValueNeeded`) instead of row-by-row `Rows.Add` rebuilds.
  - Added debounced filter/search refresh and batched row-count swaps.
  - Added O(1) ID lookups for selected-item details and journal actions.

### Fixed

- Settings window no longer flashes a light frame while opening in dark mode.
- Job Manager refresh, filtering, and selection-detail updates are now significantly faster on large history datasets.

## [1.0.4.7] - 2026-02-15

### Added

- Advanced jobs data model with schema versioning (`SchemaVersion = 2`) and explicit queue entry records (`QueueEntries`) including enqueue metadata.
- Legacy catalog migration path from `QueuedJobIds` to new queue-entry records.
- Job service APIs for queue entry inspection, queue reordering, targeted dequeue, queue clearing, run deletion, and history clearing.
- New Job Manager UI redesign (`Jobs Console`) with:
  - Single-grid unified view (`All Items`, `Saved Jobs`, `Queue`, `Run History`)
  - Search and run-status filtering
  - Queue controls (run now, remove, move up/down, clear)
  - Run-history actions (delete run, open journal, clear history)
  - Details pane for selected item metadata
  - Persisted window/grid/splitter layout

### Changed

- Main-form queue execution now dequeues queue-entry work items (not just raw job IDs), preserving queue-entry metadata in created runs.
- Runs now record queue linkage (`QueueEntryId`) and queue-attempt index (`QueueAttempt`) for better traceability in history.

### Fixed

- Job-catalog load/save tests now validate queue-entry persistence and legacy queue migration behavior.
- Job Manager layout spacing was tightened and splitter restore timing corrected so details panel visibility and footer alignment remain stable.
- Filter-row alignment was refined, including `Refresh` button positioning.

## [1.0.4.1] - 2026-02-15

### Added

- New overwrite policy option: `Always ask for each conflict` in Copy Defaults.
- Conflict prompt flow for `Ask` policy now prompts per destination collision (`Yes` overwrite, `No` skip, `Cancel` abort start).

### Changed

- Source and destination text input rows were adjusted to avoid vertical wrapper stretch artifacts in dark mode.

### Fixed

- Explorer context menu sync now only applies and logs when the registration state actually changes, preventing repeated `enabled` log lines on unrelated settings saves.

## [1.0.4.0] - 2026-02-14

### Added

- New dedicated About dialog (Help -> About XactCopy) with the XactCopy banner logo and themed presentation.
- About dialog now includes author attribution (`Wimukthi Bandara`) and a live system snapshot (OS/runtime/architecture/CPU/RAM).

### Changed

- About dialog version string is now centered along the bottom of the logo banner.
- About dialog content was streamlined to remove duplicated metadata text.

### Fixed

- Dark-mode theming now applies correctly to `RichTextBox` controls used by the About dialog.
- `Assets/logo.png` is now packaged with app output so the About logo loads correctly in published builds.

## [1.0.3.7] - 2026-02-14

### Added

- README now includes a screenshots section for the main window and appearance settings page.
- Screenshot assets were added under `docs/screenshots/` for release documentation.

## [1.0.3.6] - 2026-02-14

### Added

- New Diagnostics settings category for worker telemetry policy and UI responsiveness tuning.
- Worker telemetry profiles (`Normal`, `Verbose`, `Debug`) with per-run effective progress/log throttling.
- Main-window live diagnostics strip (render latency, event coalescing, log queue depth, dropped/suppressed counts).
- Appearance settings expansion:
  - Accent source (`Auto`, `System`, `Custom`) and custom accent color picker.
  - UI density (`Compact`, `Normal`, `Comfortable`) and UI scale (`90%`, `100%`, `110%`, `125%`).
  - Operations log font family/size.
  - Grid appearance controls (alternating rows, row height, header style).
  - Main status-row visibility controls (Buffer, Rescue, Diagnostics).
  - Progress-bar style (`Thin`, `Standard`, `Thick`) and optional percentage overlay.
  - Window chrome mode (`Themed` vs `Standard`) with DWM fallback handling.

### Changed

- Main UI now batches worker log rendering in timed chunks to avoid flooding the UI thread during high-event copy sessions.
- Dynamic transfer status labels were switched to fixed-layout rendering to reduce repeated autosize/layout churn under fast progress updates.
- Runtime log view on main form is now virtualized (ListView virtual mode) for smoother scrolling and lower UI memory pressure on very large runs.
- Worker now emits telemetry summaries for progress/log suppression behavior at run end.
- Theme application now consumes appearance profile settings (accent + grid presentation + density/scale pass).

### Fixed

- Recovery heartbeat persistence is now queued off the UI thread, removing synchronous state writes from the progress repaint path.
- Container double-buffering is now applied across core layouts/grids to reduce resize and repaint jitter.
- Explorer shell-launch flow now opens destination folder picker automatically after source is populated from `Copy with XactCopy` (cancel keeps current destination).
- Main operations-log tooltip now uses concise wording (`Operations log`).
- Settings page host now scrolls correctly when appearance page content exceeds the visible area, preventing clipped controls.
- Restart prompt reasons now include key UI settings (`Window chrome mode`, `UI scale`, `UI density`) in addition to theme/startup behaviors.

## [1.0.2.7] - 2026-02-14

### Changed

- Settings dialog now tracks dirty state live and keeps `Save` disabled until there are actual changes.
- Restart-required settings are now identified during save, with a clear reason list shown to the user.

### Fixed

- Theme mode changes are now treated as restart-required for full startup-level color-mode consistency.
- After saving restart-required settings, XactCopy now offers immediate relaunch (or defers safely if a run is active).

## [1.0.2.4] - 2026-02-13

### Added

- New `AegisRescueCore` multi-pass recovery engine for per-file resilient copy (`FastScan`, `TrimSweep`, `Scrape`, `RetryBad`) with split-on-failure bad-block isolation.
- Journal persistence now tracks per-file rescue ranges (`Pending`, `Good`, `Bad`, `Recovered`) and last rescue pass for stronger crash-resume continuity.
- Copy progress IPC/model telemetry now includes rescue pass, bad-region count, and remaining unrecovered bytes.
- Performance settings now expose `AegisRescueCore` pass tuning defaults (per-pass chunk sizes, split minimum, and pass retries).

### Changed

- Main window live status now displays rescue telemetry and includes active rescue-pass details in the running title text.
- Write path no longer flushes every chunk, reducing overhead during sustained throughput.
- Runtime log is now event-focused (retries/errors/rescue summaries) instead of per-file/pass noise, and the UI log buffer auto-trims to recent events.
- Progress rendering now coalesces high-frequency worker snapshots and caps redraw rate to keep UI telemetry smooth during very fast copies.

### Fixed

- Main progress labels/bars no longer remain at zero until completion under bursty high-throughput runs.
- Final UI progress state now snaps to the job result to prevent completion percentages from stopping short of expected totals.

## [1.0.1.3] - 2026-02-13

### Fixed

- Updater apply phase no longer gets stuck showing `Canceling...` while waiting to exit/restart.
- Release notes text in updater now normalizes escaped newlines for readable formatting.

## [1.0.0.9] - 2026-02-13

### Added

- In-app updater flow now supports downloading release assets and applying updates in-place (zip/exe), with progress UI.

### Changed

- `Check for Updates` now opens the installer dialog for available releases instead of only prompting to open the release page.

## [1.0.0.7] - 2026-02-13

### Changed

- Worker binary renamed from `XactCopy.Worker` to `XactCopyExecutive`.
- Worker now uses the same application icon as the main XactCopy UI.
- Supervisor worker launch resolution updated to prefer `XactCopyExecutive` and keep legacy fallback support.

### Fixed

- Running the worker manually without `--pipe` no longer throws an unhandled exception.
- Standalone worker launch now exits cleanly with a user-facing message.

## [1.0.0.3] - 2026-02-13

### Changed

- Default update release URL now points to `https://api.github.com/repos/Wimukthi/XactCopy/releases/latest`.
- Settings UI placeholder for release URL now reflects the real XactCopy endpoint.
- Existing empty update URL values are normalized to the default endpoint.

## [1.0.0.1] - 2026-02-13

### Added

- Initial public release.
- Out-of-process resilient copy worker with supervisor restart handling.
- Journal resume/recovery, pause/resume/cancel, telemetry, and adaptive buffer controls.
- Explorer integration, settings dialog, dark mode, update checking, and job manager.
