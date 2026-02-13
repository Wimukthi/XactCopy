# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

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
