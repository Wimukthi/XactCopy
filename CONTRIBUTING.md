# Contributing & Building from source

XactCopy is a native C++20 application built without CMake, targeting either the
MSYS2 **g++** toolchain or Visual Studio **MSVC**. Everything is driven by a
single PowerShell script.

## Prerequisites

- **Windows 10/11, 64-bit.**
- One C++20 toolchain:
  - **MSYS2 g++** (default) — the mingw64 `bin` directory must be on `PATH`
    (the build script prepends `C:\msys64\mingw64\bin`), **or**
  - **MSVC** (Visual Studio 2022/2026, Desktop C++ workload).
- A checkout of **Wimukthi.Win32Theme** beside this repository. The default
  layout is `Software\Wimukthi.Win32Theme` and `Software\XactCopyNative`.
  For another location, pass `-ThemeRoot <path>` to `build.ps1`.
- Optional, for packaging: **[Inno Setup 6](https://jrsoftware.org/isinfo.php)**
  (`winget install --id JRSoftware.InnoSetup --exact`).

## Build and test

```powershell
# g++ (default): builds the worker, UI, and tests, then runs the tests
.\build.ps1 -RunTests

# MSVC
.\build.ps1 -Compiler msvc -RunTests

# Theme framework checked out somewhere else
.\build.ps1 -ThemeRoot D:\Libraries\Wimukthi.Win32Theme -RunTests
```

Binaries are written in-tree to `build\`:

- `XactCopy.exe` — the UI / supervisor.
- `XactCopyExecutive.exe` — the worker.
- `xactcopy_*_tests.exe` — the test suites (core, storage, supervisor, worker).

The test run covers wire-format goldens, storage round-trips, a supervised copy
job, and kill-mid-job auto-recovery with a byte-exact resume.

## Packaging an installer

```powershell
.\installer\build-installer.ps1
```

This builds the binaries (unless `-SkipBuild`), reads the product version from
`src\ui\app.rc`, and runs Inno Setup to produce
`installer\output\XactCopySetup-<version>-win-x64.exe`. To cut a release, attach
that setup (and/or a `win-x64.zip`) to a GitHub release tagged at the new version
— the in-app updater looks at `releases/latest`.

## Repository layout

- `src/core/` — header-only core: JSON (System.Text.Json-compatible),
  date/time wire formats, models, the IPC envelope + messages, the framed
  named-pipe transport, and crypto (BCrypt SHA-256/HMAC, DPAPI, base64).
- `src/storage/` — the journal store, bad-range-map store, and job catalog.
- `src/worker/` — the `XactCopyExecutive` worker: the resilient copy engine,
  scanner, rescue/salvage pipeline, and the IPC host (`main.cpp`).
- `src/ui/` — the `XactCopy.exe` UI: supervisor, recovery, settings, theming,
  the main window and dialogs (Settings, Job Manager, About, Update), shell
  integration, the updater, and the app icon/resources.
- `tests/` — the unit-test suites plus golden files.
- `installer/` — the Inno Setup script and its build wrapper.
- `tools/` — `New-AppIcon.ps1`, which regenerates the multi-resolution app icon.

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for how the pieces fit together.

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
update the affected golden by hand or emit it from the writer under test, and
say why in the commit.

Two of them play a specific role worth knowing:

- `golden_journal_payload.json` is in the **old .NET shape** and is not what the
  current writer emits. It is the fixture proving the reader still loads
  journals written by earlier builds, which is what keeps an upgrade from losing
  a user's resume state. Leave it alone.
- `golden_journal_payload_slim.json` is the **current writer's** output, which
  omits default-valued members. This is the one to regenerate when journal
  serialization changes.

## Conventions

- Match the surrounding code's style, naming, and comment density.
- Keep the two toolchains (g++ and MSVC) both building; the UI is a single
  translation unit so a header change rebuilds it.
- The UI must render correctly in both light and dark themes and at non-100% DPI.
- Verify UI changes with a screenshot rather than assuming; the app never steals
  foreground focus.

## License

By contributing you agree that your contributions are licensed under the project's
**GNU GPL v3.0**.
