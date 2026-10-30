# Changelog

All notable changes to XactCopy are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
Versions are four-part (`major.minor.build.revision`) and tagged `v<version>`.

## [2.2.0.0] — 2026-10-30

### Changed

- **One task list.** The separate "What to copy" list is gone. Every add
  (+ Files, + Folder, Paste, a drop, Explorer) makes a task straight away and
  selects it; the task list moved to the left.
- **Settings beside the list.** The right column shows the selected task's
  settings, and changes are saved to it as you make them. With no task
  selected (click below the list, or Esc) it holds the **New task settings**
  that the next add uses, including an optional destination that skips the
  "where to?" question. Add to tasks, Edit in builder, and the Tools and Tasks
  menu commands built on them are gone; File has **New Task from Files /
  Folder / Paste**, and **Ctrl+Enter** queues the selected task.
- **Action box.** The Copy / Assess readability tabs became an **Action** box at
  the left of the toolbar, for what the next add does, and an Action box at
  the top of each task's settings for what that task does.
- **Now running card.** Progress, speed, ETA, Pause, and Stop sit on a card at
  the bottom right that follows the running task whatever is selected, with
  the buffer, rescue, journal, and diagnostics lines below. A finished run's
  figures stay until the next run.
- Options are always shown; the collapsible Options section is gone. The
  Now running card always shows its bars and figures, below a line that
  separates it from the task's settings.

### Added

- **Tasks files.** File → **Export Tasks…** saves every task to a readable
  `.xctasks` JSON file, and **Import Tasks…** (or dropping one on the window)
  adds the tasks it lists, unqueued. Files can be written by hand: options
  left out come from your defaults. The format is in the User Guide.
- **Remove** on the task toolbar removes the selected task.
- A one-folder task has **Copy the folder's contents, not the folder** under
  its Source, replacing the old right-click option.

### Fixed

- A null window handle was repainted when a run finished, which asked Windows
  to redraw every window on the desktop.
- The Source box and the (now visible) Action box shared a control ID.

## [2.1.5.0] — 2026-10-29

### Changed

- **Edit in builder** (and double-clicking a task) adds the task to the list
  as one more row instead of replacing the list, and adding that row again
  updates the task instead of making a copy. Its folders show as folders and
  its own Resume setting is kept.
- **Stop** holds the queue instead of starting the next task. A task that fails
  to start goes back to the front of a held queue; a later part of a
  multi-part task waits to run from that part.
- Tools **Assess Readable Files…** is now **Queue as Assessment Tasks** and no
  longer leaves the window in Assess mode. Tasks **Save Current Options As
  Pinned Task…** is now **Add as Pinned Tasks**, which is what it did.
- Leaving Verified Copy or Recover Media by changing one option keeps that
  profile's engine. Verified Copy no longer carries the plaintext or
  recovered-overwrite overrides from Settings.

### Fixed

- History: Del no longer deletes the run in progress, and Clear history waits
  for it to finish. Older runs of the same source and destination no longer
  show the latest run's journal, report, or range map, and a run whose journal
  retention removed offers no Resume. Runs that finished with skipped, failed,
  or recovered files can be resumed.
- Stop works while waiting for administrator approval; a late approval then
  starts nothing.
- Stop, Clear history, Resume Interrupted Run, Check for Updates, and the
  bad-range map tools are enabled only when they can act, including while a
  worker starts. The pane's Pause / Resume button follows the run's state.
- Inspect and Clear Source Bad-Range Map act on the selected row and name the
  source.
- A long Source box is read in full instead of being cut at 4,096 characters.
- Ctrl+A and right-click in the list update the Source and Destination boxes;
  right-clicking below the last task no longer acts on the previous selection.
- If a task cannot be saved, rows that were not saved stay in the list.
- Adding items that are already listed or do not exist says so instead of
  doing nothing.

## [2.1.4.0] — 2026-10-29

### Fixed

- Running a task that needs direct NTFS reads no longer puts its files and
  folders back on the left-hand list. The administrator window that takes over
  now shows the list the first window had, so rows not yet added to tasks are
  kept too. Resuming a run no longer replaces the list either.
- History's **Report…**, **Range map**, and **Delete** are now disabled until a
  run is selected instead of doing nothing when clicked.
- **Range map** lists files with unreadable or recovered bytes first, with a
  count, so in a copy they are no longer buried among fully read files. When a
  run recorded no ranges it explains why instead of only saying so.

## [2.1.3.0] — 2026-10-28

### Added

- **Low-space warning.** Copy rows are checked against the free space where
  they are going (rows bound for the same drive are totalled) as sizes and
  destinations come in. Rows that may not fit show "Not enough space", the list
  footer says so, and the log gets one warning. Nothing is blocked: the size is
  an estimate and files the copy replaces give space back.

### Fixed

- Settings no longer opens with a white, unpainted area around the page list.

## [2.1.2.0] — 2026-10-28

### Added

- **Destination asked when adding.** In Copy mode every add (+ Files,
  + Folder, Paste, drag-and-drop, Explorer) asks where the items go;
  cancelling adds nothing. Each row keeps its own destination, shown in the
  list, and becomes a task with that destination.
- **Source box** under the list, with **Browse…** to pick other files or
  folders for the selected row. Source and Destination edit the selected row
  and can never be left empty; a source path that does not exist is refused
  and the row is put back.
- **Browse…** for the journal folder in Settings.
- The main window title shows the version.

### Changed

- Settings has five pages instead of seven (Verification moved into Copy
  Defaults, the Explorer switch into Recovery & Startup), is shorter, and
  balances sections across its two columns.

### Fixed

- No white flash when switching Settings pages: every page is built once and
  swapped with redraw suspended.

## [2.1.1.0] — 2026-10-28

### Added

- **Direct NTFS reads for copies.** The **Read method** choice now applies to
  copies as well as assessments (one setting for both; the Recover Media
  profile turns it on). Reads go straight to the volume and are abandoned at
  the timeout instead of stalling inside the filesystem on failing media. A
  file is read directly only when Windows confirms its whole length holds
  written data; compressed, encrypted, very small, pre-allocated, and
  read-only files are read through Windows, and the log names each one.
- **Journal folder** setting (Settings → Diagnostics). New runs write their
  journals there; the mirror copy stays in your profile, runs started before a
  change still resume, and journal clean-up covers both folders.

### Changed

- Options follow what each pipeline actually uses. **Skip known-bad ranges**
  now covers what "Use bad-range map" did (that box had no effect on its own),
  copies gain a visible **Record bad ranges**, and in Assess mode
  Fragile-media mode locks **Scan** on *Precise*, which it always forced.
- The copy profiles say what they mean: Verified Copy turns recording and
  direct reads off; Recover Media turns them on along with salvage,
  continue-on-error, fragile pacing, adaptive 1 MB reads, and wait for media.

### Fixed

- Raw volume reads flush a file's cached writes first, so a recently written
  file is never read stale.

## [2.1.0.0] — 2026-10-28

### Changed

- **New main window.** A **Copy | Assess readability** switch picks the mode;
  the choices below it are captioned and relabel themselves for scanning, and
  switching modes never moves the layout. The item list and the task list share
  one top and height, and splitters size the log and the two columns (both
  remembered; double-click resets).
- **What to copy** groups everything one action adds (several folders picked in
  one dialog, a paste, a drop) into one row. **Add to tasks**, beside the list,
  turns the selected rows (or all of them) into tasks; Ctrl+Enter adds and
  queues. A row spanning several folders or drives becomes one task with
  several parts, run back to back.
- **Tasks and History replace the Job Manager window.** Tasks lists the running
  task, then the queue in run order, then everything else; two buttons follow
  the selected row (Run / Queue, Start now / Unqueue, Pause / Stop), and **Hold**
  lets the current task finish before the queue waits. Run and Start now start
  the chosen task even while the queue is held. History lists every run with
  its details and journal findings, plus Resume, Journal, Report, Range map,
  and Delete. Tasks, the queue, and history live in the signed job catalog.
- **Pinned tasks** always stay; tasks added from the builder follow an on-exit
  choice (Ask / Keep / Discard, with "Remember my choice" and a setting).
  Existing saved jobs become pinned tasks, and leftover unstarted tasks from
  earlier builds move into the queue once.
- **Explorer** gains **Assess with XactCopy** next to **Copy with XactCopy** on
  files, folders, drives, and folder backgrounds; each opens in its mode.
- **Direct NTFS assessments elevate automatically.** A task, a queued task, or a
  crash-recovery resume that needs raw volume reads relaunches XactCopy as
  administrator and continues the same run. Links and folders opened from an
  elevated window still open as the signed-in user.
- **Settings** were audited: choices the main window already shows are saved
  with **File → Save as defaults** instead of being repeated, the Job Manager
  grid settings are gone, and interrupted-run handling is one choice (Ask /
  Resume automatically / Leave in History).

### Fixed

- Paths longer than 260 characters copy and verify correctly.
- Full verification reads the destination from disk (unbuffered) instead of
  from the file cache.
- Sleep and resume no longer make the watchdog restart a healthy worker.
- Failure, error, mismatch, and warning log lines are never rate-limited away.
- Closing with a run active shows one prompt that names a copy or an
  assessment. The interrupted-run prompt can be told not to ask at startup;
  History keeps offering Resume.
- Salvage and "Treat Access Denied as contention" can no longer be combined
  into a run the worker refuses.
- Task rows, list borders, and owner-drawn controls no longer show overlapping
  text, white flashes, or stale slivers while resizing or switching modes.

### Performance

- Resume no longer re-verifies finished files twice; small files copy in
  parallel by default; per-file media checks are cached; verification hashes
  source and destination concurrently.
- The log is virtual (owner-data), filters by severity, searches, and saves
  without blocking the window.

## [2.0.0.8] — 2026-08-18

### Fixed

- Interrupted copy and assessment runs now retain their immutable launch
  options in both run history and the journal. Resume restores the operation
  mode, scan backend/profile, conflict and verification policy, rescue/map
  choices, worker tuning, timeouts, and selected-source scope before restarting.
- Fast assessments checkpoint completed files while workers are active, and
  precise assessments bind and checkpoint partial byte-range coverage so a
  restart resumes useful work instead of beginning at the first file.
- Parallel native-copy cancellation now commits completed publications before
  unwinding, and cancellation results can no longer erase a valid journal path.
- Job Manager run details now show authenticated journal progress plus every
  recorded unreadable/recovered or failed filename. Interrupted, cancelled,
  incomplete, and failed runs with journal state expose a **Resume** action.
- Fragile-mode first-read stops now retain the failed primary-data offset and
  attempted length in both the journal and bad-range map for assessments and
  copies. A first observation remains diagnostic; repeated matching evidence is
  still required before it can become a skip hint.
- Missing, locked, access-denied, and unavailable sources are classified before
  fragile-media handling, so those operational failures cannot be mislabeled as
  physical bad ranges. The bad-range inspector also cross-references the latest
  assessment journal and lists failures that could not be localized to bytes.

## [2.0.0.7] — 2026-08-06

### Added

- **Raw disk scanning** can read allocated NTFS file extents directly for
  damaged-media scans, with standard file I/O retained as the safe fallback.
  The experimental path is opt-in and requires the appropriate Windows access.
- **Persistent run defaults** can be saved explicitly from the main window, so
  reusable scan, recovery, buffer, retry, timeout, and verification choices
  survive a restart without persisting job-specific paths.

### Changed

- The log window now provides horizontal scrolling for long paths and messages.
- Job Manager keeps the populated row and column grid while removing the heavy
  column rules that previously continued through the empty list tail.
- Raw-scan, settings, worker, storage, and UI regression coverage was expanded.

## [2.0.0.6] — 2026-07-30

### Changed

- **Journals are compressed.** Snapshots over 64 KB are stored as an XPRESS_HUFF
  container instead of raw JSON; smaller journals stay plain text so the common
  case is still readable. Combined with the slimmer encoding below, a real
  196,519-entry whole-drive journal drops from 96.3 MB to 4.1 MB, and its full
  on-disk artifact set — snapshot, three backups, and the mirror of each — from
  578 MB to 24.9 MB. Journals written by earlier builds are detected and read as
  before, so upgrading keeps existing resume state.
- **Journal entries are slimmer.** Entries no longer store members holding their
  default value, and no longer repeat the relative path that already keys the
  entry. On that same 196,519-entry journal this cuts the snapshot from 96.3 MB
  to 54.2 MB. An absent member reads back as its default, so journals written by
  either implementation still load identically on the other.
- **Journal flush frequency scales with journal size.** A snapshot is rewritten
  in full on every flush, so on a large job the previous fixed 500 ms cadence
  cost more I/O than the copy itself. Flushes are now budgeted against measured
  save cost, holding journal I/O near 5% of run time. Small journals are
  unaffected and keep their sub-second cadence.
- Rotating a journal backup renames the previous snapshot instead of copying it,
  removing a full file read and write per save on both the primary and the
  mirror.

### Removed

- The .NET interop tooling, retired along with the VB.NET implementation it
  tested: `tools/GoldenGen`, `tools/StorageProbe`, `tools/InteropProbe`, the
  `build.ps1 -CrossTests` suite, and the cross-compat modes in the storage test
  binary. `tests/golden/` remains as frozen wire-format fixtures, and the catalog
  fixture those probes shared now drives a native round-trip test instead, so no
  coverage is lost. Reading journals, bad-range maps, and settings written by the
  1.x releases is unaffected and still tested.

## [2.0.0.5] — 2026-07-27

### Fixed

- Explorer file and folder context-menu actions now show the exact selected item
  as the source and always copy only that selection. The internal parent folder
  used by the worker is no longer exposed in the Source box.

## [2.0.0.4] — 2026-07-27

### Fixed

- The update checker no longer reports the currently installed release as a new
  update. Product-version metadata, updater comparisons, About text, and
  installer packaging now share one version definition to prevent drift.

## [2.0.0.3] — 2026-07-27

### Changed

- Native light/dark mode, High Contrast, title bars, menus, and common-control
  theming now come from the shared `Wimukthi.Win32Theme` framework. XactCopy
  keeps its own palette, Fluent icons, progress indicators, and application-
  specific owner-drawn controls.
- GCC and MSVC builds accept `-ThemeRoot` when the framework is not checked out
  beside XactCopy.
- The installer includes the framework and Darkmodelib license notices, and the
  repository documents where to obtain the corresponding source.

## [2.0.0.2] — 2026-07-25

### Added

- **Add Files…** next to the source box opens a multi-select picker, and files or
  folders can be **dragged onto the window**. Both add to the current selection.
- A summary under the source box shows what is queued — e.g. *"3 items selected
  (2 files, 1 folder) — only these will be copied"* — with a **Clear** button.
  Choosing a source folder with Browse clears the selection and copies the whole
  folder again.

### Fixed

- **Copying several selected items now works.** Windows Explorer runs a
  context-menu command once per selected item, so picking three files started
  three copies of XactCopy and stacked up a destination prompt for each — and
  only the last item survived. Selections are now gathered into a single job with
  one source folder and one destination prompt, whether you pick files, folders,
  or a mix.

### Changed

- The installer follows the Windows light/dark setting.

## [2.0.0.1] — 2026-07-25

### Fixed

- **The Explorer context menu did not pick up the selection.** Launch arguments
  were parsed from a command line that excludes the executable path, so the first
  switch was discarded and the selected file or folder never reached the Source
  box. Right-clicking a file, folder, or drive now fills in the source correctly.
- Launching from Explorer while XactCopy is **already running** no longer throws
  the selection away — the running window picks it up instead of just surfacing.
- The Job Manager title bar now shows the XactCopy icon.
- The Job Manager details pane no longer jumps back to the top while you are
  reading it; the three-second refresh leaves an unchanged pane alone.
- The About window no longer clips the bottom of the "XactCopy" title.

### Changed

- **All message boxes and prompts are dark-themed**, matching the rest of the
  application instead of falling back to the system dialog.
- Buttons carry **Fluent icons**, and checkboxes and glyphs are antialiased.
- The Job Manager grid draws row and column separators.
- Update settings moved out of Settings: the *check automatically* toggle now
  lives in the About window.
- Application icons are **multi-resolution** (16–256 px), so they stay sharp in
  the title bar, Alt-Tab, and Explorer. The worker executable carries the icon
  and version details too.

## [2.0.0.0] — 2026-07-25

The application was rewritten as a native C++/Win32 program: a themed desktop UI
supervising a separate copy/scan worker over a versioned named-pipe protocol. No
managed runtime is required to run it. `2.0.0.0` marks this rewrite and
supersedes the 1.x VB.NET releases.

### Copy & scan

- Resilient copy engine with resumable, journaled transfers.
- Multi-strategy rescue passes over damaged regions, best-effort salvage of
  partially readable blocks, configurable retry/back-off, and a fragile-media
  mode.
- Copy engines: **Auto**, **Native Fast**, and **Managed Rescue**.
- Overwrite policies: overwrite, skip existing, overwrite-if-newer, and ask.
- Bad-block scanning with precise and fast parallel profiles.
- Optional elevated raw-volume reads for allocated files on local NTFS volumes,
  with per-file fallback to standard reads and no bad-range format changes.
- Bad-range maps: known-bad regions are saved and skipped on later runs.
- Integrity verification with SHA-256/512, full or sampled.
- Adaptive buffering, wait-for-media, and continue-on-error handling.

### Application

- Native dark/light theming with accent colour, dark title bars and menus, and
  owner-drawn controls; per-monitor DPI scaling; taskbar progress.
- Job Manager: saved jobs, a run queue, and run history with journal links.
- Crash recovery: interrupted runs are detected and can be resumed from their
  journal on next launch.
- Settings across seven pages, preserved across upgrades.
- Live telemetry: throughput with running average, ETA, buffer utilisation, and
  rescue-pass progress, plus a severity-coloured log.

### Integration & updates

- Explorer integration: *Copy with XactCopy* and *Scan for Bad Blocks* verbs,
  Open-with and drag-and-drop registration, a *Send to* shortcut, and a
  Run-dialog entry.
- In-app updater: checks GitHub releases, downloads, verifies against the
  published SHA-256, and installs in place.
- Windows installer package built with Inno Setup.

### Notes

- The previous VB.NET implementation is preserved on the `vbnet-legacy` branch
  and the `vbnet-final` tag. Its on-disk journals, bad-range maps, and settings
  remain compatible with this release.

[2.1.3.0]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.1.3.0
[2.1.2.0]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.1.2.0
[2.1.1.0]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.1.1.0
[2.1.0.0]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.1.0.0
[2.0.0.8]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.8
[2.0.0.7]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.7
[2.0.0.6]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.6
[2.0.0.5]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.5
[2.0.0.4]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.4
[2.0.0.3]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.3
[2.0.0.2]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.2
[2.0.0.1]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.1
[2.0.0.0]: https://github.com/Wimukthi/XactCopy/releases/tag/v2.0.0.0
