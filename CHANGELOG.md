# Changelog

All notable changes to XactCopy are documented in this file. The format is based
on [Keep a Changelog](https://keepachangelog.com/); this project uses date-stamped
releases.

## [2.0.0] — Native application

The application has been rewritten as a native C++/Win32 program. The `2.0.0`
version marks this rewrite and supersedes the `1.x` VB.NET releases.

XactCopy is now a native C++/Win32 application: a themed desktop UI that
supervises a separate copy/scan worker over a versioned named-pipe protocol. No
managed runtime is required to run it.

### Copy & scan

- Resilient copy engine with resumable, journaled transfers.
- Multi-strategy rescue passes over damaged regions, best-effort salvage of
  partially readable blocks, configurable retry/back-off, and a fragile-media
  mode.
- Copy engines: **Auto**, **Native Fast**, and **Managed Rescue**.
- Overwrite policies: overwrite, skip existing, overwrite-if-newer, and ask.
- Bad-block scanning with precise and fast parallel profiles.
- Bad-range maps: known-bad regions are saved and skipped on later runs.
- Integrity verification with SHA‑256/512, full or sampled.
- Adaptive buffering, wait-for-media, and continue-on-error handling.

### Application

- Native dark/light theming with accent colour, dark title bars and menus, and
  owner-drawn controls; per-monitor DPI scaling; taskbar progress.
- Job Manager: saved jobs, a run queue, and run history with journal links.
- Crash recovery: interrupted runs are detected and can be resumed from their
  journal on next launch.
- Settings across eight pages, preserved across upgrades.
- Live telemetry: throughput with running average, ETA, buffer utilization, and
  rescue-pass progress, plus a severity-coloured log.

### Integration & updates

- Comprehensive Explorer integration: *Copy with XactCopy* and *Scan for Bad
  Blocks* context-menu verbs, Open-with / drag-and-drop registration, a *Send to*
  shortcut, and a Run-dialog entry.
- In-app updater: checks GitHub releases, downloads, verifies against the
  published SHA‑256, and installs in place.
- Windows Installer package built with Inno Setup.

### Notes

- The previous VB.NET implementation is preserved on the `vbnet-legacy` branch and
  the `vbnet-final` tag. Its on-disk journals, bad-range maps, and settings remain
  compatible with this release.
