<div align="center">

# XactCopy

**A verified file copier and recovery tool for unstable media.**

[![Release](https://img.shields.io/github/v/release/Wimukthi/XactCopy?label=release)](https://github.com/Wimukthi/XactCopy/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/Wimukthi/XactCopy/total?label=downloads)](https://github.com/Wimukthi/XactCopy/releases)
[![License](https://img.shields.io/badge/license-GPL--3.0-blue)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2B%20x64-0078D4)](https://github.com/Wimukthi/XactCopy/releases/latest)

XactCopy copies files that ordinary tools give up on — failing drives, scratched
discs, flaky USB sticks, and network shares that drop out — and recovers as much
readable data as possible instead of aborting at the first error.

![XactCopy copying from a failing drive](docs/images/main-window.png)

</div>

## Why XactCopy

When a normal copy hits an unreadable region it stops, and you lose the whole
transfer. XactCopy treats bad media as the expected case:

- **It recovers deliberately.** Unreadable regions are retried, worked around,
  and logged. Recover Media can preserve the readable portions of a file, and
  Continue on error can move on to later files; either policy is reported as
  non-exact whenever data was lost or skipped.
- **It resumes.** Every job is journaled, so an interrupted or crashed transfer
  picks up where it left off instead of starting over.
- **It tells the truth.** Copied data can be verified with SHA-256/512, and a
  readability assessment records which allocated file ranges could not be read.
  It is not a whole-disk surface diagnostic. Recovered,
  skipped, or partially enumerated output is reported as incomplete rather than
  being presented as an exact copy.

## Features

- **Resilient copy engine** — resumable transfers, multi-strategy *rescue passes*
  over damaged regions, best-effort *salvage* of partially readable blocks,
  configurable retry/back-off, and a *fragile-media* mode that treats a dying
  drive gently rather than pushing it over the edge.
- **Readability assessment** — read a drive's allocated files or a selected
  folder without creating a destination, record unreadable file ranges, and save the
  result as a **bad-range map** so later copies skip known-bad regions instead of
  re-reading them. An optional elevated raw-volume backend reads allocated file
  extents directly on local NTFS volumes and falls back safely when unsupported.
- **Integrity verification** — full SHA-256 verification is the safe default;
  sampled and unverified modes remain available as attended-only choices.
- **Job Manager** — save copy/scan jobs in a signed, rotating catalog, queue
  integrity-safe jobs for unattended execution, and review run history with
  per-run status and journal links.
- **Crash recovery** — if XactCopy or the machine goes down mid-transfer, the
  next launch offers to resume the interrupted run.
- **Explorer integration** — right-click any file, folder, or drive to *Copy with
  XactCopy* or *Assess Readable Files*, plus Open-with, Send-to, and Run-dialog
  entries.
- **In-app updates** — checks GitHub for new releases, then downloads, verifies
  (SHA-256), and installs them in place.
- **Native and themed** — a C++/Win32 application with a real dark mode,
  per-monitor DPI scaling, and taskbar progress. No runtime to install.

## Install

**Installer (recommended)** — download the latest
`XactCopySetup-<version>-win-x64.exe` from the
[Releases](https://github.com/Wimukthi/XactCopy/releases/latest) page and run it.
It registers XactCopy, adds Start-menu and optional desktop shortcuts, and can be
updated in place from inside the app.

**Portable** — the `XactCopy-<version>-win-x64.zip` from the same page runs
unzipped. Keep `XactCopy.exe` and `XactCopyExecutive.exe` together in one folder
and run `XactCopy.exe`.

Requires **Windows 10 or later, 64-bit**. Each release also publishes a `.sha256`
file next to every asset so you can verify the download.

## Quick start

1. Launch **XactCopy** and pick a **Source** and a **Destination**.
2. Choose a **Mode**:
   - **Verified Copy / Recovery** — copy data from source to destination.
   - **Assess Readable Files** — read allocated source files and map unreadable regions (no
     destination needed).
3. Adjust options if you like — the defaults are sensible — then press **Start**.
   Progress, throughput, ETA, and a live log update as it runs; **Pause** and
   **Cancel** stay available throughout.

The main-window choices are per-run overrides. Use **Save defaults** when you
want the reusable options to seed future runs; use a saved job when the mode,
paths, and complete option set must be retained together.

For an actively failing drive, copy the irreplaceable data first with **Recover
Media** and **Fragile-media mode**. A full assessment reads the source once
without rescuing data and can consume part of a dying device's remaining life.
Use **Assess Readable Files** before repeated recovery passes only when the media
is stable enough and the saved map will reduce later reads. A skip hint is
trusted only after two matching observations; non-exact recovery output keeps
its original extension and receives an adjacent `.recovery.json` manifest.

## Options at a glance

| Control | Choices | Purpose |
| --- | --- | --- |
| **Mode** | Verified Copy / Recovery · Assess Readable Files | Copy data, or assess allocated source files without a destination |
| **Profile** | Verified Copy · Recover Media · Custom | Apply a coherent safe-copy or recovery preset, or use expert defaults |
| **Conflict** | Overwrite · Skip existing · Overwrite if newer · Stop on conflict | How existing destination files are handled |
| **Verify** | None · Sampled · Full | Post-copy SHA-256/512 integrity check; Full is the safe default |
| **Salvage · Resume · Bad-range map · Adaptive buffer** | toggles | Core resilience behaviours |
| **Continue on error · Skip known-bad · Wait for media · Fragile mode** | toggles | Error-handling behaviour |
| **Buffer (MB) · Retries · Timeout (s)** | numbers | Performance and persistence tuning |

Every option is explained in the [User Guide](docs/USER_GUIDE.md).

## Documentation

| Document | Contents |
| --- | --- |
| [User Guide](docs/USER_GUIDE.md) | Every mode, option, dialog, and workflow |
| [Troubleshooting](docs/TROUBLESHOOTING.md) | What to do when a copy or scan misbehaves |
| [Architecture](docs/ARCHITECTURE.md) | Supervised worker, IPC protocol, journals, on-disk formats |
| [Contributing](CONTRIBUTING.md) | Building from source, tests, packaging, conventions |
| [Changelog](CHANGELOG.md) | Version history |
| [Security policy](SECURITY.md) | Reporting a vulnerability |

## Building from source

XactCopy is a C++20 Win32 application built by a single PowerShell script — no
CMake, no package manager. With MSYS2 g++ (or MSVC) and a sibling
[Wimukthi.Win32Theme](https://github.com/Wimukthi/Wimukthi.Win32Theme) checkout:

```powershell
.\build.ps1 -RunTests
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for prerequisites, the MSVC path, and
packaging an installer.

## License

XactCopy is released under the **GNU General Public License v3.0** — see
[LICENSE](LICENSE). Component licensing and corresponding-source information are
recorded in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
