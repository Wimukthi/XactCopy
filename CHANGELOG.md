# Changelog

All notable changes to XactCopy are documented in this file. The format is based
on [Keep a Changelog](https://keepachangelog.com/); this project uses date-stamped
releases.

## [2.0.0.6]

### Removed

- The .NET interop tooling is retired along with the VB.NET implementation it
  tested: `tools/GoldenGen`, `tools/StorageProbe`, `tools/InteropProbe`, the
  `build.ps1 -CrossTests` suite, and the cross-compat modes in the storage test
  binary. `tests/golden/` stays as frozen wire-format fixtures, and the catalog
  fixture those probes shared now drives a native round-trip test instead, so no
  coverage is lost. Reading journals, bad-range maps, and settings written by the
  1.x releases is unaffected and still tested.

### Changed

- **Journals are compressed.** Snapshots over 64 KB are stored as an
  XPRESS_HUFF container instead of raw JSON; smaller journals stay plain text so
  the common case is still readable. Combined with the slimmer encoding below, a
  real 196,519-entry whole-drive journal drops from 96.3 MB to 4.1 MB, and its
  full on-disk artifact set (snapshot, three backups, and the mirror copies of
  each) from 578 MB to 24.9 MB. Journals written by earlier builds are detected
  and read as before, so upgrading keeps existing resume state.
- **Journals are substantially smaller and far cheaper to write.** Journal file
  entries no longer store members that hold their default value, and no longer
  repeat the relative path that already keys the entry. On a real 196,519-entry
  whole-drive journal this cuts the snapshot from 96.3 MB to 54.2 MB. Both the
  native and .NET readers treat an absent member as its default, so journals
  written by either implementation still load identically on the other.
- **Journal flush frequency now scales with journal size.** A snapshot is
  rewritten in full on every flush, so on a large job the previous fixed 500 ms
  cadence cost more I/O than the copy itself. Flushes are now budgeted against
  measured save cost, holding journal I/O near 5% of run time. Small journals are
  unaffected and keep their sub-second cadence.
- Rotating a journal backup renames the previous snapshot instead of copying it,
  removing a full file read and write per save on both the primary and the mirror.

## [2.0.0.5]

### Fixed

- Explorer file and folder context-menu actions now show the exact selected item
  as the source and always copy only that selection. The internal parent folder
  used by the worker is no longer exposed in the Source box.

## [2.0.0.4]

### Fixed

- The update checker no longer reports the currently installed release as a new
  update. Product-version metadata, updater comparisons, About text, and
  installer packaging now share one version definition to prevent drift.

## [2.0.0.3]

### Changed

- Native Windows light/dark mode, High Contrast, title bars, menus, and common
  control theming now use the shared `Wimukthi.Win32Theme` framework.
- XactCopy retains its existing palette, Fluent icons, progress indicators, and
  application-specific owner-drawn controls.
- GCC and MSVC builds accept `-ThemeRoot` when the framework is not checked out
  beside XactCopy.

### Packaging

- The installer includes the framework and Darkmodelib license notices, and the
  repository documents where to obtain the corresponding source.

## [2.0.0.2]

### Fixed

- **Copying several selected items now works.** Windows Explorer runs a
  context-menu command once per selected item, so picking three files started
  three copies of XactCopy and stacked up a destination prompt for each — and
  only the last item survived. Selections are now gathered into a single job with
  one source folder and one destination prompt, whether you pick files, folders,
  or a mix.

### Added

- **Add Files…** next to the source box opens a multi-select picker, and files or
  folders can be **dragged onto the window**. Both add to the current selection.
- A summary under the source box shows what is queued — e.g. *"3 items selected
  (2 files, 1 folder) — only these will be copied"* — with a **Clear** button.
  Choosing a source folder with Browse clears the selection and copies the whole
  folder again.

### Changed

- The installer follows your Windows light/dark setting.

## [2.0.0.1]

### Fixed

- **Explorer context menu did not pick up the selection.** Launch arguments were
  parsed from a command line that excludes the executable path, so the first
  switch was discarded and the selected file or folder never reached the Source
  box. Right-clicking a file, folder, or drive now fills in the source correctly.
- Launching from Explorer while XactCopy is **already running** no longer throws
  the selection away — the running window picks it up instead of just surfacing.
- The Job Manager title bar now shows the XactCopy icon.
- The Job Manager details pane no longer jumps back to the top while you are
  reading it; the three-second refresh leaves an unchanged pane alone.
- The About window no longer clips the bottom of the "XactCopy" title.

### Changed

- **All message boxes and prompts are now dark-themed**, matching the rest of the
  application instead of falling back to the system dialog.
- Buttons carry **Fluent icons**, and checkboxes and glyphs are **antialiased**.
- The Job Manager grid now draws row and column separators.
- Update settings moved out of Settings: the "check automatically" toggle now
  lives in the About window.
- Application icons are now **multi-resolution** (16–256 px), so they stay sharp
  in the title bar, Alt-Tab, and Explorer. The worker executable carries the icon
  and version details too.

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
