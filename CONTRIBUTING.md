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
- Optional, for packaging: **[Inno Setup 6](https://jrsoftware.org/isinfo.php)**
  (`winget install --id JRSoftware.InnoSetup --exact`).
- Optional, for the cross-compatibility tests: the **.NET SDK** and a checkout of
  the legacy VB.NET tree (see [below](#cross-compatibility-tests)).

## Build and test

```powershell
# g++ (default): builds the worker, UI, and tests, then runs the tests
.\build.ps1 -RunTests

# MSVC
.\build.ps1 -Compiler msvc -RunTests
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
- `tools/` — .NET helpers used for test-vector generation and
  cross-compatibility checks (see below).

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for how the pieces fit together.

## Cross-compatibility tests

XactCopy's storage and IPC formats are byte-compatible with the previous VB.NET
implementation, which is preserved on the **`vbnet-legacy`** branch and the
**`vbnet-final`** tag. The `-CrossTests` suite validates both directions using
the .NET stores, so it needs that tree checked out alongside this one:

```powershell
# check the legacy VB.NET tree out next to this repo, then:
.\build.ps1 -CrossTests
```

The `tools/StorageProbe`, `tools/InteropProbe`, and `tools/GoldenGen` .NET
projects drive those checks and regenerate `tests/golden/`. They are only needed
when changing serialized shapes or the IPC protocol.

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
