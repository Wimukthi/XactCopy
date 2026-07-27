<div align="center">

# XactCopy

**A resilient file mover and bad-block scanner for unstable media.**

XactCopy copies files that ordinary tools give up on — failing drives, scratched
discs, flaky USB sticks, and network shares that drop out — and recovers as much
readable data as possible instead of aborting at the first error.

![XactCopy main window](docs/images/main-window.png)

</div>

## Why XactCopy

When a normal copy hits an unreadable sector it stops, and you lose the whole
transfer. XactCopy treats bad media as the expected case:

- **It keeps going.** Unreadable regions are retried, worked around, and logged —
  the rest of the file (and the rest of the job) still gets copied.
- **It resumes.** Every job is journaled, so an interrupted or crashed transfer
  picks up exactly where it left off instead of starting over.
- **It tells the truth.** Copied data can be verified with SHA‑256/512, and a
  bad-block scan maps out exactly which regions of a drive are failing.

## Features

- **Resilient copy engine** — resumable transfers, multi-strategy *rescue passes*
  over damaged regions, best-effort *salvage* of partially readable blocks,
  configurable retry/back-off, and a *fragile-media* mode that treats the source
  gently to avoid pushing a dying drive over the edge.
- **Bad-block scanning** — scan a drive or folder for unreadable sectors without
  copying anything, in a precise or a fast parallel profile, and save the result
  as a **bad-range map** so future copies skip known-bad regions instead of
  re-reading them.
- **Integrity verification** — verify copied data with SHA‑256/512, either fully
  or with sampling for speed.
- **Job Manager** — save copy/scan jobs, queue them, and review run history with
  per-run status and journal links.
- **Crash recovery** — if XactCopy or the machine goes down mid-transfer, it
  offers to resume the interrupted run on next launch.
- **Explorer integration** — right-click any file, folder, or drive to *Copy with
  XactCopy* or *Scan for Bad Blocks*, plus Open-with, Send-to, and Run-dialog
  entries.
- **In-app updates** — checks GitHub for new releases and can download, verify
  (SHA‑256), and install them in place.
- **Native and themed** — a fast C++/Win32 application with a proper dark mode,
  per-monitor DPI scaling, and taskbar progress. No runtime to install.

## Install

**Recommended:** download the latest `XactCopySetup-<version>-win-x64.exe` from the
[Releases](https://github.com/Wimukthi/XactCopy/releases) page and run it. The
installer registers XactCopy, adds Start-menu (and optional desktop) shortcuts,
and can be updated in place from within the app.

**Portable:** the release also works unzipped — keep `XactCopy.exe` and
`XactCopyExecutive.exe` together in the same folder and run `XactCopy.exe`.

Windows 10 or later, 64-bit.

## Quick start

1. Launch **XactCopy** and pick a **Source** and **Destination** folder.
2. Choose a **Mode**:
   - **Copy** — move data from source to destination.
   - **Scan Bad Blocks** — read the source and map unreadable regions (no
     destination needed).
3. Adjust options if you like (sensible defaults are pre-selected), then click
   **Start**. Progress, throughput, ETA, and a live log update as it runs; use
   **Pause**/**Cancel** at any time.

For a failing drive, a good recipe is: **Scan Bad Blocks** first to build a
bad-range map, then **Copy** with *Use bad-range map* and *Salvage* enabled so
the copy spends its time on recoverable data.

See the **[User Guide](docs/USER_GUIDE.md)** for what every option does.

## Options at a glance

| Control | Choices | Purpose |
| --- | --- | --- |
| **Mode** | Copy · Scan Bad Blocks | Move data, or just map bad regions |
| **Engine** | Auto · Managed Rescue · Native Fast | Trade raw speed for damaged-media resilience |
| **Overwrite** | Overwrite · Skip existing · Overwrite if newer · Ask | How existing destination files are handled |
| **Verify** | None · Sampled · Full | Post-copy SHA‑256/512 integrity check |
| **Salvage / Resume / Bad-range map / Adaptive buffer** | toggles | Core resilience behaviors |
| **Continue on error · Skip known-bad · Wait for media · Fragile mode** | toggles | Fine-tune error handling |
| **Buffer (MB) · Retries · Timeout (s)** | numbers | Performance and persistence tuning |

## Documentation

- **[User Guide](docs/USER_GUIDE.md)** — every mode, option, and workflow.
- **[Architecture](docs/ARCHITECTURE.md)** — how the supervised worker, journals,
  and recovery fit together.
- **[Contributing / Building from source](CONTRIBUTING.md)** — toolchains, build
  scripts, tests, and packaging.
- **[Changelog](CHANGELOG.md)** — version history.

## License

XactCopy is released under the **GNU General Public License v3.0**. See
[LICENSE](LICENSE). Component licensing and corresponding-source information
are recorded in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
