# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

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
