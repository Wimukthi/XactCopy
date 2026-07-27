# XactCopy User Guide

This guide explains every mode, option, and workflow in XactCopy. If you just
want to get going, see the Quick Start in the [README](../README.md).

## Contents

- [The main window](#the-main-window)
- [Copy mode](#copy-mode)
- [Scan Bad Blocks mode](#scan-bad-blocks-mode)
- [Resilience options](#resilience-options)
- [Verification](#verification)
- [Bad-range maps](#bad-range-maps)
- [The Job Manager](#the-job-manager)
- [Crash recovery](#crash-recovery)
- [Explorer integration](#explorer-integration)
- [Settings](#settings)
- [Updates](#updates)
- [Where XactCopy stores its data](#where-xactcopy-stores-its-data)

## The main window

Pick a **Source** and **Destination** with the *Browse…* buttons (or type/paste a
path), choose a mode and options, and press **Start**. While a job runs the
window shows two progress bars (current file and overall), throughput with a
running average, an ETA, buffer utilization, rescue-pass telemetry, and a live,
colour-coded log. **Pause**, **Resume**, and **Cancel** are always available;
input controls are locked during a run so the job can't be changed underneath
itself.

## Copy mode

Copies everything under **Source** into **Destination**, resiliently.

**Engine** selects the copy strategy:

- **Auto** *(default)* — starts fast and automatically escalates to the rescue
  engine when it encounters errors. Best for most situations.
- **Native Fast** — a straight, high-throughput copy for healthy media.
- **Managed Rescue** — the full damaged-media engine from the start: use it when
  you already know the source is failing.

**Overwrite** controls what happens when a destination file already exists:

| Choice | Behavior |
| --- | --- |
| **Overwrite** | Always replace the destination file |
| **Skip existing** | Never touch an existing destination file |
| **Overwrite if newer** | Replace only when the source is newer |
| **Ask** | Prompt for each conflict |

## Scan Bad Blocks mode

Reads the **Source** without writing anything and records which regions are
unreadable. Use it to survey a suspect drive, or to build a bad-range map before
a copy. Scanning runs in a precise profile or a faster parallel profile
(configurable under **Settings → Performance**); the fast profile falls back to
precise reads around any fault it finds so the map stays accurate.

A destination is not required in this mode.

## Resilience options

These toggles shape how XactCopy handles trouble:

- **Salvage unreadable blocks** — when a region can't be read cleanly, recover as
  much of it as possible (down to the sector) instead of discarding the whole
  block. Leaves any truly unreadable bytes clearly accounted for.
- **Resume from journal** — continue an interrupted job from its journal instead
  of restarting. On by default; every job is journaled regardless.
- **Use bad-range map** — consult (and update) a saved map of known-bad regions
  for this source. See [Bad-range maps](#bad-range-maps).
- **Skip known-bad ranges** — with a map loaded, don't re-read regions already
  known to be bad. This is what makes a second pass fast.
- **Adaptive buffer** — size the read buffer dynamically from measured throughput
  and latency, backing off on slow/latent media.
- **Continue on error** — keep going past a file that ultimately can't be copied,
  rather than stopping the whole job.
- **Wait for media** — if the source or destination disappears (a drive spins
  down, a disc is ejected), pause and wait for it to come back instead of
  failing.
- **Fragile-media mode** — minimise stress on a dying drive: gentler retry
  timing and access patterns, trading speed for a better chance of getting the
  data off before the drive fails completely.

**Buffer (MB)**, **Retries**, and **Timeout (s)** tune the read buffer size, how
many times a failed read is retried, and how long an operation may stall before
it's treated as failed.

### Rescue passes

When the engine hits damage it runs a sequence of *rescue passes* over the
affected regions — fast sweeps first, then progressively more aggressive
strategies (forward and reverse trim sweeps, fine-grained scraping, and targeted
retries of the worst spots), splitting failing ranges and adapting to the density
of errors. The rescue telemetry line reports the current pass, how many regions
remain, and the bytes still outstanding.

## Verification

**Verify** checks that what landed at the destination matches the source:

- **None** — no verification (fastest).
- **Sampled** — hash a representative sample of each file.
- **Full** — hash the entire file (SHA‑256/512).

Verification runs after the data is written and reports any mismatch in the log
and job summary.

## Bad-range maps

A **bad-range map** is a saved record of which byte ranges of a particular source
are unreadable. Build one with **Scan Bad Blocks**, or let a copy populate it as
it goes. On later runs, enable **Use bad-range map** and **Skip known-bad ranges**
so XactCopy spends its effort on recoverable data instead of grinding on sectors
it already knows are dead. Maps are matched to their source and are integrity-
protected on disk.

## The Job Manager

**Jobs → Job Manager** opens a console for saved jobs, the run queue, and history:

- **Saved jobs** — save a configured copy/scan as a named job, then run, rename,
  duplicate, or delete it.
- **Queue** — line jobs up to run back-to-back; reorder with Top/Up/Down/Bottom,
  or clear the queue. Queued jobs can auto-run on startup (a Settings option).
- **History** — every run is recorded with its type, state (Running, Completed,
  Failed, Cancelled, Interrupted…), timing, and a link to its journal. Filter by
  view or status, search, and open a run's journal to inspect it.

The grid auto-refreshes; double-click a run to open its journal, and use
Del/F5/Enter as shortcuts.

## Crash recovery

Because every run is journaled, XactCopy can recover from an unclean shutdown. If
a run was active when XactCopy or Windows went down, the next launch detects the
interrupted run and offers to **Resume** it (continuing from the journal),
**Discard** it, or decide **Later**.

## Explorer integration

Enable **Settings → Explorer Integration** to add XactCopy to the Windows shell:

- **Copy with XactCopy** — on files, folders, drives, and folder backgrounds.
  A selected item is shown as the exact source and only that item is copied.
  Multi-select copies the selected items together from their shared source
  folder. Using the folder-background command copies the full folder.
- **Scan for Bad Blocks with XactCopy** — on folders and drives; opens XactCopy
  ready to scan that target.
- **Open with / drag-and-drop** — XactCopy is registered as an application, so you
  can drop files onto it or pick it from *Open with*.
- **Send to → XactCopy** and a **Run-dialog** entry (`XactCopy`).

All shell entries are per-user and are removed cleanly when you turn the option
off.

## Settings

Settings is organised into pages:

- **Appearance** — theme (System/Light/Dark), accent colour, window chrome, UI
  density and scale, log font and severity colouring, and progress-bar style.
- **Copy Defaults** — the engine, overwrite, verification, and resilience options
  a new job starts with.
- **Performance** — buffer sizing, scan profile and worker count, and throughput
  limits.
- **Diagnostics** — extra on-screen counters and status rows.
- **Verification** — default hashing behavior.
- **Updates** — the release channel/URL and whether to check on launch.
- **Recovery & Startup** — crash-resume behavior and whether queued jobs run at
  startup.
- **Explorer Integration** — the shell entries described above.

Settings are saved to a JSON file and preserved across upgrades.

## Updates

**Help → About → Check for Updates** (or **Tools → Check for Updates**) asks
GitHub for the latest release. If a newer version exists, XactCopy shows the
release notes and can **Download & Install** it: the package is downloaded,
verified against its published SHA‑256, and applied in place, after which
XactCopy restarts on the new version. You can also just open the release page in
your browser.

## Where XactCopy stores its data

XactCopy keeps its journals, bad-range maps, job catalog, recovery state, and
security keys under `%LOCALAPPDATA%\XactCopy\`, and its settings in a JSON file in
the same area. These are integrity-protected (HMAC-signed journals, a
DPAPI-protected map key) and are shared by the app and its worker process.
