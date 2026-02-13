# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

### Added

- Explorer context-menu argument handling hardened for reliable path passing.
- Title bar run telemetry and Windows taskbar progress integration.
- One-grid job manager with persisted layout and column state.
- Recovery startup and interrupted-run resume prompt flow.
- Explorer shell integration settings and update/recovery categories in Settings.
- App/window icon integration from `Icons/xactcopy.ico`.

### Changed

- Settings dialog layout and control alignment cleanup.
- Main/settings synchronization for copy defaults, including salvage fill pattern.
- Progress bar theming and dark-mode rendering polish.
- UI tooltips expanded across main window, settings, job manager, and recovery prompt.

### Fixed

- Explorer verb launch path parsing edge cases.
- Settings layout/splitter initialization crash paths.
- Horizontal scrolling behavior for large logs and grids.
