# XactCopy User Guide

Everything XactCopy does, option by option. If you only want to get going, the
Quick Start in the [README](../README.md) is enough.

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
- [Menu reference](#menu-reference)
- [Command line](#command-line)
- [Where XactCopy stores its data](#where-xactcopy-stores-its-data)

## The main window

![The XactCopy main window during a copy](images/main-window.png)

Pick a **Source** and **Destination** with *Browse…*, or type or paste a path.
*Add Files…* opens a multi-select picker, and files or folders dropped onto the
window are added the same way. When a specific set of items is selected rather
than a whole folder, a summary line appears under the source box — for example
*"3 items selected (2 files, 1 folder) — only these will be copied"* — with a
**Clear** button. Choosing a source folder with *Browse* clears the selection and
copies the whole folder again.

Below the paths are the four mode combos (**Mode**, **Engine**, **Overwrite**,
**Verify**), the resilience toggles, and the numeric tuning fields.

While a job runs the window shows:

- Two progress bars: the current file and the job overall.
- A counter line: files done, bytes done, and the current rescue pass.
- Throughput with a running average, an **ETA**, buffer utilisation, and
  rescue-pass telemetry.
- A **Journal** path for the run, and a diagnostics row.
- A live, severity-coloured operations log.

**Start**, **Pause**, and **Cancel** are always available; the input controls
lock during a run so the job cannot be changed underneath itself.

## Copy mode

Copies everything under **Source** into **Destination**, resiliently.

**Engine** selects the copy strategy:

| Engine | Behaviour |
| --- | --- |
| **Auto** *(default)* | Starts fast and escalates to the rescue engine on the first error. Best for most situations. |
| **Native Fast** | A straight, high-throughput copy for healthy media. |
| **Managed Rescue** | The full damaged-media engine from the start. Use it when you already know the source is failing. |

**Overwrite** controls what happens when a destination file already exists:

| Choice | Behaviour |
| --- | --- |
| **Overwrite** | Always replace the destination file |
| **Skip existing** | Never touch an existing destination file |
| **Overwrite if newer** | Replace only when the source is newer |
| **Ask** | Prompt for each conflict |

## Scan Bad Blocks mode

Reads the **Source** without writing anything and records which regions are
unreadable. Use it to survey a suspect drive, or to build a bad-range map before
a copy. No destination is required.

Scanning runs in a precise profile or a faster parallel one (**Settings →
Performance → Scan profile**); the fast profile falls back to precise reads
around any fault it finds, so the resulting map stays accurate.

## Resilience options

These toggles shape how XactCopy handles trouble.

| Option | What it does |
| --- | --- |
| **Salvage unreadable blocks** | When a region can't be read cleanly, recover as much of it as possible — down to the sector — instead of discarding the whole block. Unreadable bytes are filled with a known pattern and accounted for. |
| **Resume from journal** | Continue an interrupted job from its journal instead of restarting. On by default; every job is journaled either way. |
| **Use bad-range map** | Consult, and update, a saved map of known-bad regions for this source. See [Bad-range maps](#bad-range-maps). |
| **Skip known-bad ranges** | With a map loaded, don't re-read regions already known to be bad. This is what makes a second pass fast. |
| **Adaptive buffer** | Size the read buffer dynamically from measured throughput and latency, backing off on slow or high-latency media. |
| **Continue on error** | Keep going past a file that ultimately can't be copied, instead of stopping the whole job. |
| **Wait for media** | If the source or destination disappears — a drive spins down, a disc is ejected — pause and wait for it to come back rather than failing. |
| **Fragile-media mode** | Minimise stress on a dying drive: gentler retry timing and access patterns, trading speed for a better chance of getting the data off before the drive fails completely. |

**Buffer (MB)**, **Retries**, and **Timeout (s)** set the read buffer size, how
many times a failed read is retried, and how long an operation may stall before
it counts as failed.

### Rescue passes

When the engine hits damage it runs a sequence of *rescue passes* over the
affected regions: fast sweeps first, then progressively more aggressive
strategies — forward and reverse trim sweeps, fine-grained scraping, and targeted
retries of the worst spots — splitting failing ranges and adapting to the density
of errors. The rescue telemetry line reports the current pass, how many regions
remain, and the bytes still outstanding.

## Verification

**Verify** checks that what landed at the destination matches the source:

| Mode | Behaviour |
| --- | --- |
| **None** | No verification. Fastest. |
| **Sampled** | Hash a representative sample of each file. |
| **Full** | Hash the entire file. |

The hash algorithm (SHA-256 or SHA-512) and the sampling chunk size and count are
set under **Settings → Verification**. Verification runs after the data is
written, and any mismatch is reported in the log and the job summary.

## Bad-range maps

A **bad-range map** records which byte ranges of a particular source are
unreadable. Build one with **Scan Bad Blocks**, or let a copy populate it as it
goes. On later runs, enable **Use bad-range map** and **Skip known-bad ranges** so
XactCopy spends its effort on recoverable data instead of grinding on sectors it
already knows are dead.

Maps are bound to the source they were made for and are integrity-protected on
disk, so a map can never be applied to the wrong media. **Settings → Copy
Defaults** controls whether runs update the map and how old a map may be before
it is ignored.

## The Job Manager

**Jobs → Job Manager…** opens the Jobs Console: saved jobs, the run queue, and
run history in one grid.

![The XactCopy Job Manager](images/job-manager.png)

- **Saved jobs** — save a configured copy or scan as a named job (**Jobs → Save
  Current Options As Job…**), then run, rename, duplicate, or delete it.
- **Queue** — line jobs up to run back-to-back; reorder with Top/Up/Down/Bottom,
  or clear the queue. Queued jobs can run automatically at startup (**Settings →
  Recovery & Startup**).
- **History** — every run is recorded with its type, state (Running, Completed,
  Failed, Cancelled, Interrupted…), timings, source and destination, and a
  summary. **Open Journal** opens the run's journal.

Filter with the **View** and **Run Status** combos or the **Search** box. The
grid refreshes automatically; double-click a run to open its journal, and use
<kbd>Del</kbd>, <kbd>F5</kbd>, and <kbd>Enter</kbd> as shortcuts.

## Crash recovery

Because every run is journaled, XactCopy can recover from an unclean shutdown. If
a run was active when XactCopy or Windows went down, the next launch detects the
interrupted run and offers to **Resume** it — continuing from the journal —
**Discard** it, or decide **Later**.

**Settings → Recovery & Startup** controls whether you are prompted, whether
interrupted runs resume automatically, whether the prompt keeps reappearing until
resolved, and whether XactCopy restarts itself at the next logon after an
interruption.

## Explorer integration

Enable **Settings → Explorer Integration** to add XactCopy to the Windows shell:

- **Copy with XactCopy** — on files, folders, drives, and folder backgrounds. A
  selected item is shown as the exact source and only that item is copied.
  Multi-select copies the selected items together from their shared source
  folder. The folder-background command copies the whole folder.
- **Scan for Bad Blocks with XactCopy** — on folders and drives; opens XactCopy
  ready to scan that target.
- **Open with / drag-and-drop** — XactCopy is registered as an application, so
  you can drop files onto it or pick it from *Open with*.
- **Send to → XactCopy**, and a **Run** dialog entry (`XactCopy`).

All shell entries are per-user and are removed cleanly when you turn the option
off. Launching from Explorer while XactCopy is already running hands the
selection to the running window rather than starting a second copy.

## Settings

![XactCopy settings](images/settings.png)

**Tools → Settings…** opens seven pages. Everything here sets *defaults* for new
jobs; the main window's controls override them per run.

### Appearance

Theme mode (Dark, System, Classic), accent source (Auto, System, Custom) and a
custom accent colour, and window chrome (themed or standard title bar). Layout
covers UI density and scale. The Operations Log section sets the log font, size,
and severity colouring. Grid Appearance covers alternating rows, row height, and
header style. Status & Progress toggles the buffer and rescue status rows,
percentage text on progress bars, and the progress-bar style.

### Copy Defaults

Default run behaviour — resume from journal, salvage, continue on file errors,
preserve source timestamps, copy empty directories, wait for source/destination,
wait for lock release, and whether *Access Denied* counts as contention.

Policies — overwrite policy, symlink handling (skip or follow), and the fill
pattern used for unrecoverable bytes (zero, `0xFF`, or random).

Bad Range Map — whether maps are used and updated, whether known-bad ranges are
skipped, and the maximum age of a map before it is ignored. **Experimental raw
disk scan backend** currently has no effect: raw-disk reads are not implemented
in the native build, and a scan with it enabled logs that it is falling back to
standard file reads.

### Performance

Transfer tuning — adaptive buffer, transfer engine policy, scan profile, worker
process priority, manual buffer size, retry count, operation and per-file
timeouts, a throughput cap, small-file worker count and threshold, and scan
worker count.

Fragile Media Guard — enable fragile mode by default, skip a file on the first
read error, persist those skips across a resume, and the failure window,
threshold, and cooldown that trigger the guard.

Contention & Source Mutation — lock-probe interval, and what to do when the
source changes mid-run (fail the file, skip it, or wait for it to reappear).

Rescue Engine Tuning — per-pass chunk sizes (FastScan, TrimSweep, Scrape,
RetryBad), the minimum split size, and per-pass retry counts. `0` means *auto*
for each; leave these alone unless you are tuning for a specific drive.

### Diagnostics

Worker telemetry profile (Normal, Verbose, Debug), progress interval, and a log
rate cap. UI diagnostics strip and its refresh interval, and the maximum number
of lines the log keeps. Journal storage — whether backups are discarded when a
run completes, how long journals are kept, and how many are always retained.

### Verification

Whether verification is on by default, full or sampled mode, SHA-256 or SHA-512,
and the sampled-mode chunk size and count.

### Recovery & Startup

Auto-start at the next logon after an interruption, automatically run queued jobs
at startup, prompt to resume interrupted runs, auto-resume without prompting,
keep prompting until resolved, and the heartbeat write interval.

### Explorer Integration

A single switch for the shell entries described in
[Explorer integration](#explorer-integration).

Settings are stored as JSON. Unknown keys are preserved, so a newer and an older
build can share the same file without either losing settings.

## Updates

**Help → About XactCopy** shows the exact version and build — the two things any
bug report needs — along with the **Check automatically** switch that decides
whether XactCopy looks for updates on launch.

![The XactCopy About window](images/about.png)

**Check for Updates…** there, or **Tools → Check for Updates…**, asks GitHub for
the latest release now.

![The XactCopy update dialog](images/update.png)

If a newer version exists, XactCopy shows the release notes and the package name.
**Download & Install** downloads the package, verifies it against its published
SHA-256, and applies it in place, after which XactCopy restarts on the new
version. **View Release** opens the release page in your browser instead.

XactCopy will not install a package it cannot check: if no SHA-256 is published
for the release asset, the update is refused rather than applied unverified.

## Menu reference

| Menu | Items |
| --- | --- |
| **File** | Start / Pause / Resume / Cancel Copy · Open Journal Folder · Open Crash Folder · Exit |
| **Tools** | Scan for Bad Blocks… · Settings… · Check for Updates… |
| **Jobs** | Save Current Options As Job… · Job Manager… · Resume Interrupted Job · Run Next Queued Job |
| **Help** | About XactCopy |

## Command line

`XactCopy.exe` accepts the switches the Explorer verbs use. They are documented
here because they are also useful for scripting; there is no separate console
interface.

| Switch | Effect |
| --- | --- |
| `--source <path>` | Set the source folder. |
| `--from-explorer <path> [<path>…]` | Copy exactly these items. Their shared parent becomes the internal copy root. |
| `--from-explorer-folder <path>` | Copy the whole folder (what the folder-background verb uses). |
| `--scan-from-explorer <path> [<path>…]` | Open in Scan Bad Blocks mode targeting these items. |
| `--resume-interrupted` | Show the resume prompt for an interrupted run on launch. Also accepted as `--force-resume-prompt`. |
| `--recovery-autostart` | Marks the launch as the post-interruption auto-start. Set by XactCopy itself; not meant to be typed. |

`--from-explorer`, `--scan-from-explorer`, and `--from-explorer-folder` also
accept the `=` form, e.g. `--from-explorer="D:\Recovery"`.

Only one instance runs at a time. Launching a second one forwards its arguments
to the running window and exits.

`XactCopyExecutive.exe` is the worker process. It takes `--pipe <name>` and is
started by `XactCopy.exe`; running it directly does nothing useful.

## Where XactCopy stores its data

Everything lives under `%LOCALAPPDATA%\XactCopy\`:

| Path | Contents |
| --- | --- |
| `journals\` | Per-run journals, their rotating backups, and mirrors |
| `badmaps\` | Bad-range maps |
| `jobs\catalog.json` | Saved jobs, the queue, and run history |
| `runtime\recovery-state.json` | The active run and clean-shutdown marker |
| `security\` | HMAC keys — the journal key as raw bytes, the map key DPAPI-protected for the current user |
| `settings.json` | Application settings |

**File → Open Journal Folder** opens `journals\` and **File → Open Crash Folder**
opens `runtime\`. Journals are HMAC-signed and hash-chained, so tampering is
detectable; see [ARCHITECTURE.md](ARCHITECTURE.md) for the formats.
