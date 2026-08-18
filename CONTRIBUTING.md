# Contributing

XactCopy is a native C++20 Win32 application built without CMake, targeting
either the MSYS2 **g++** toolchain or Visual Studio **MSVC**. Everything is
driven by a single PowerShell script.

## Prerequisites

- **Windows 10/11, 64-bit.**
- One C++20 toolchain:
  - **MSYS2 g++** (default) — `build.ps1` prepends `C:\msys64\mingw64\bin` to
    `PATH`, so a standard MSYS2 install needs no extra setup; **or**
  - **MSVC** — Visual Studio 2022 or newer with the Desktop C++ workload.
    `build.ps1` finds `vcvars64.bat` via `vswhere` (including preview and
    Insiders installs).
- A checkout of **[Wimukthi.Win32Theme](https://github.com/Wimukthi/Wimukthi.Win32Theme)**
  beside this repository — the default layout is `…\Wimukthi.Win32Theme` and
  `…\XactCopyNative`. Pass `-ThemeRoot <path>` for anything else.
- Optional, for packaging: **[Inno Setup 6](https://jrsoftware.org/isinfo.php)**
  (`winget install --id JRSoftware.InnoSetup --exact`).

## Build and test

```powershell
# g++ (default): build the worker, UI, and tests, then run the tests
.\build.ps1 -RunTests
```

```powershell
# MSVC
.\build.ps1 -Compiler msvc -RunTests
```

```powershell
# Theme framework checked out somewhere else
.\build.ps1 -ThemeRoot D:\Libraries\Wimukthi.Win32Theme -RunTests
```

Binaries are written in-tree to `build\`, which is git-ignored:

- `XactCopy.exe` — the UI / supervisor.
- `XactCopyExecutive.exe` — the worker.
- `xactcopy_{core,storage,supervisor,worker}_tests.exe` — the test suites.

The test run covers the wire-format goldens, storage round-trips, a supervised
copy job with kill-mid-job auto-recovery, and the integrity contract: staged
publication, conflict semantics, source-identity binding, resume coverage,
salvage visibility, metadata and EFS fidelity, and raw-volume extent mapping.
Run it before opening a pull request; both toolchains must keep building.

One raw-volume test needs a real NTFS volume and Administrator rights, so it is
opt-in and skipped by default:

```powershell
$env:XACTCOPY_RAW_DISK_TEST = "1"; .\build.ps1 -RunTests
```

Do not pipe `build.ps1` through `2>&1` in Windows PowerShell 5.1. Redirecting a
native command's stderr wraps each line in a `NativeCommandError`, which the
script's `$ErrorActionPreference = "Stop"` turns into a terminating error — an
ordinary compiler warning then looks like a failed build.

## Repository layout

| Path | Contents |
| --- | --- |
| `src/` | Application sources — see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md#source-layout) for the breakdown |
| `src/version.h` | The single source of truth for the product version |
| `tests/` | Unit-test suites and `tests/golden/` fixtures |
| `installer/` | Inno Setup script and its build wrapper |
| `tools/` | `New-AppIcon.ps1`, which regenerates `Icons\xactcopy.ico` from `Assets\xactcopy.png` |
| `docs/` | User guide, troubleshooting, architecture, and the integrity hardening contract |

## Releasing

Bump `src/version.h` — the resource metadata, the About window, the updater's
comparison, and the installer's file names all read from it, so nothing else
needs editing. Then:

```powershell
.\installer\build-installer.ps1
```

This builds the binaries (unless `-SkipBuild`), reads the version from
`src\version.h`, and runs Inno Setup to produce
`installer\output\XactCopySetup-<version>-win-x64.exe`.

To cut a release, add a `CHANGELOG.md` entry and attach two assets to a GitHub
release tagged `v<version>`:

- `XactCopySetup-v<version>-win-x64.exe` — the Inno Setup installer.
- `XactCopy-v<version>-win-x64.zip` — the portable build, flat at the archive
  root: `XactCopy.exe`, `XactCopyExecutive.exe`, `LICENSE`, `README.md`,
  `THIRD_PARTY_NOTICES.md`, and `licenses\`. The updater searches the extracted
  tree for `XactCopy.exe`, so keep the executables at the root.

Each asset needs a `.sha256` sidecar beside it containing **only** the bare
lowercase hash followed by CRLF — 66 bytes, no filename, no `*` marker.

The in-app updater reads `releases/latest` and **refuses to install a package it
cannot checksum**. It takes the hash from GitHub's asset `digest` field if the
API reports one, and otherwise from a sibling checksum asset — `<name>.sha256`,
`.sha256sum`, `.sha256.txt`, or a `checksums.txt`. Publishing the `.sha256` files
is what keeps that path working regardless of what the API returns, so treat them
as required release assets.

## Golden files

`tests/golden/` pins the wire format: the IPC envelopes, the indented JSON of
stored artifacts, and the escaping, number, and date formatting rules. The core
and storage suites byte-compare against it, so an accidental change to any
serializer fails the build rather than silently altering an on-disk format.

These files were originally emitted by the .NET implementation to prove the two
builds were byte-compatible. That implementation is retired and the generators
that produced them (`GoldenGen`, `StorageProbe`, `InteropProbe`) have been
removed, so the goldens are now **frozen fixtures** — there is no regeneration
tool, and that is deliberate. If you intentionally change a serialized shape,
update the affected golden by hand or emit it from the writer under test, and say
why in the commit.

Two of them play a specific role worth knowing:

- `golden_journal_payload.json` is in the **old .NET shape** and is not what the
  current writer emits. It is the fixture proving the reader still loads journals
  written by earlier builds, which is what keeps an upgrade from losing a user's
  resume state. Leave it alone.
- `golden_journal_payload_slim.json` is the **current writer's** output, which
  omits default-valued members. This is the one to regenerate when journal
  serialization changes.

## Conventions

- Match the surrounding code's style, naming, and comment density.
- Keep both toolchains building. The UI is a single translation unit, so a header
  change rebuilds all of it.
- The UI must render correctly in both light and dark themes and at non-100% DPI.
- Verify UI changes with a screenshot rather than assuming. The app never steals
  foreground focus, so a capture harness has to drive it deliberately.
- Anything that changes an on-disk or wire format needs a golden-file update and
  an explanation.
- The invariants in
  [docs/INTEGRITY_HARDENING_PLAN.md](docs/INTEGRITY_HARDENING_PLAN.md) are the
  integrity contract. A change may not weaken one without saying so explicitly:
  no in-place truncation of an existing destination, no partial or recovered
  result presented as an exact copy, and no journal or bad-range hint applied
  across a source whose identity has changed.

## Reporting bugs

Include the XactCopy version (**Help → About**), Windows version, what the source
media is, and the relevant part of the operations log. For a failed or stalled
run, the journal under `%LOCALAPPDATA%\XactCopy\journals\` is usually the most
useful attachment. Security issues go through [SECURITY.md](SECURITY.md) instead.

## License

By contributing you agree that your contributions are licensed under the
project's **GNU GPL v3.0**.
