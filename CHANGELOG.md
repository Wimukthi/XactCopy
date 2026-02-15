# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

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
