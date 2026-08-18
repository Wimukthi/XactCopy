# Troubleshooting

Common situations, what XactCopy is actually doing, and what to change. For what
each option means, see the [User Guide](USER_GUIDE.md).

## Getting diagnostic information

- The **operations log** in the main window is colour-coded by severity and is
  the first place to look. Increase its detail with **Settings → Diagnostics →
  Worker telemetry profile** (Normal → Verbose → Debug).
- The **journal** for a run records every file and its state.
  **File → Open Journal Folder**, or select a run in the Job Manager and press
  **Open Journal**.
- **Help → About XactCopy** has the exact version and build, which any bug report
  needs.

## Copy problems

### The copy stops on a file it can't read

Turn on **Continue on error** only when recovering the remaining files is more
important than producing one complete backup generation. The failed file stays
visible and the run is **Incomplete**; earlier files may already have been
published. Leave it off when a partial destination must never be mistaken for a
complete backup.

### It's grinding for hours on a small region

That is the rescue engine doing its job, but you can bound it:

- If the device is stable enough for another full read, run **Assess Readable
  Files**, then use **Use bad-range map** and **Skip known-bad ranges** on later
  passes. Do not assess first on a rapidly deteriorating device: recover the
  most valuable data before spending reads on diagnosis.
- Lower **Retries** and **Timeout (s)** so each bad region gives up sooner.
- Turn on **Fragile-media mode** if the drive is audibly struggling — it trades
  speed for a better chance of finishing before the drive dies.

### The drive gets worse while copying

Stop, and re-run with **Fragile-media mode** on. Under **Settings → Performance →
Fragile Media Guard** you can also make XactCopy skip a file on its first read
error, persist those skips across a resume, and cool down after a burst of
failures. A first-read stop records the failed read request as diagnostic range
evidence, but it is not used as a future skip hint until a later matching
observation confirms it. Resume restores the interrupted task's original
settings and continues its authenticated journal rather than silently using the
current main-window defaults.

### Some files copied but are corrupt

Set **Verify** to **Full** and re-run. Verification hashes the destination
against the source and reports mismatches in the log and the job summary. If a
file mismatches repeatedly, the source region is unreadable rather than merely
slow — enable **Salvage unreadable blocks** so XactCopy recovers what it can and
fills the rest with a known pattern (**Settings → Copy Defaults → Salvage fill
pattern**) instead of failing the file outright.

### "Access is denied" on lots of files

XactCopy runs as the invoking user and does not request elevation. Copying a
whole system drive will hit protected directories. Either run XactCopy as an
administrator, or point it at the specific folders you need.

If the denials are intermittent rather than constant, they are probably
contention, not permissions: turn on **Wait for lock/contention release** and
**Treat Access Denied as contention** under **Settings → Copy Defaults**.

### The destination fills up

XactCopy checks free space before a copy and warns in the log when the source is
larger than the space available. It does not stop you — a partial rescue is often
what you want. Clear space and re-run; the journal resumes where it stopped.

### The source or destination disappears mid-run

Enable **Wait for media**. Instead of failing, XactCopy pauses and waits for the
volume to return. It also captures a media-identity baseline at the start of a
run and re-checks it, so swapping in a *different* disc or stick is detected
rather than silently copied into the same journal.

### Files change while being copied

**Settings → Performance → Source mutation policy** decides what happens: fail
the file, skip it, or wait for it to reappear. For a live source, *skip* keeps
the job moving but makes the result incomplete; it is not a consistent
snapshot.

## Scan problems

### The scan is very slow

Set **Settings → Performance → Scan profile** to *Fast file assessment* and raise
**Scan workers**. The fast profile still falls back to precise reads around any
fault it finds, so the resulting map stays accurate.

### The raw-volume assessment backend falls back to standard reads

The raw backend is intentionally limited to allocated files on local NTFS
volumes. It also requires XactCopy to be running as Administrator because it
opens the source volume for read-only direct access. Network paths, other file
systems, compressed or encrypted files, reparse-point files, and changing file
layouts fall back to ordinary file reads automatically. Check the operations log
for the specific fallback reason.

Raw-volume scanning does not inspect free or unallocated sectors. Use a
dedicated disk-surface diagnostic tool when the goal is to survey the entire
physical medium.

### Scanning a whole drive misses areas

Scanning walks the file system, so it covers allocated files, not free space or
unallocated sectors. To survey the whole surface you need a disk-level tool;
XactCopy's job is to map the regions that matter to *your data* and skip them on
later copies.

## Resume and recovery

### It didn't offer to resume after a crash

Check **Settings → Recovery & Startup → Prompt to resume interrupted runs**. If
you dismissed the prompt with *Later*, **Jobs → Resume Interrupted Job** brings
it back. Note that discarding a run discards its resume point.

### A resumed job re-copies work it had already done

Up to one journal-flush interval of work can be repeated after a crash. On large
jobs that interval grows on purpose — the journal is rewritten in full on every
flush, so flushes are budgeted to stay near 5% of run time. Redoing a few seconds
of copying is cheaper than the I/O of flushing constantly. See
[ARCHITECTURE.md](ARCHITECTURE.md#durable-state-journals).

Current journals retain both the task settings and source-bound progress. Job
Manager's selected-run details show the checkpoint counts and bad/failed
filenames before you resume. If it warns that a run is legacy, that journal was
created before exact per-run settings were stored; review the restored mode and
controls because the old file cannot prove which backend or tuning was used.

### The worker keeps restarting

The supervisor restarts the worker when its heartbeat stops for 10 seconds, the
pipe is lost, or the worker exits. Lack of file progress by itself produces a
stall warning but does not kill a worker whose heartbeat is healthy; a Windows
storage call may legitimately remain blocked in a driver. Automatic crash
recovery is capped at three attempts, after which the run fails visibly instead
of looping forever. Check the red supervisor messages and Windows storage events
before raising **Operation timeout** or **Max retries**—those settings can make a
failing device do substantially more work, but they do not repair a lost
heartbeat.

## Application problems

### Journals are eating disk space

Under **Settings → Diagnostics → Journal Storage**: discard backups when a run
completes, set a retention period in days, and set how many journals are always
kept regardless. Snapshots over 64 KB are already compressed.

### The Explorer context menu didn't appear

Toggle **Settings → Explorer Integration** off and on. The entries are per-user
under `HKCU\Software\Classes`; Explorer sometimes needs a restart
(`taskkill /f /im explorer.exe & start explorer`) to pick them up.

### Launching from Explorer surfaces the existing window instead of copying

That is intended — only one instance runs at a time, and a second launch hands
its selection to the running window. If a job is already running, finish or
cancel it before starting another.

### The update check says a new version is available but the install fails

Download the installer from
[Releases](https://github.com/Wimukthi/XactCopy/releases/latest) and run it
manually. The in-app updater replaces files in the install directory, which needs
the same rights the installer had. Verify the download against its `.sha256` as
described in [SECURITY.md](../SECURITY.md).

### The UI looks wrong at high DPI or after moving monitors

XactCopy is Per-Monitor-V2 DPI aware and rebuilds its fonts and layout on monitor
changes. If something still looks off, try **Settings → Appearance → UI scale**
and **UI density**, and please report it with a screenshot and your display
scaling.

## Still stuck

Open an issue with the version, your Windows version, what the source media is,
and the relevant part of the operations log — see
[CONTRIBUTING.md](../CONTRIBUTING.md#reporting-bugs). For anything security
related, follow [SECURITY.md](../SECURITY.md) instead.
