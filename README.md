# XactCopy

Resilient file mover and bad-block scanner for unstable media — a native C++20
Win32 application. A supervised named-pipe worker (`XactCopyExecutive.exe`) runs
the resumable copy engine, rescue passes, and bad-block scan behind a versioned
JSON IPC boundary; the themed dark UI (`XactCopy.exe`) drives it, with a Job
Manager, an in-app updater, Explorer shell integration, and an Inno Setup
installer.

> The previous VB.NET implementation is preserved on the **`vbnet-legacy`**
> branch and the **`vbnet-final`** tag. The `-CrossTests` suite validates the
> native storage/IPC against that .NET code when both trees are checked out
> side by side.

Build output is written in-tree to `build\`.

## Status

- [x] **Phase 0** — foundation: JSON module, .NET time formats, build scripts
      (g++ + MSVC, no CMake), golden-file harness.
- [x] **Phase 1** — Core + IPC: all models/enums/messages, framed pipe
      transport with current-user-only security, protocol worker
      (`XactCopyExecutive.exe`) that completes the full IPC conversation with
      the real .NET supervisor stack. Job execution reports "not implemented".
- [x] **Phase 2** — Storage: `JobJournalStore` + `BadRangeMapStore` with the
      full hardening pipeline (atomic write-through snapshots, rotating
      backups, mirror snapshots, HMAC-signed envelopes, hash-chained ledger +
      anchors, DPAPI-protected map key, legacy fallback). Bidirectional
      cross-compat proven: each side reads the other's artifacts, verifies the
      other's signatures/ledgers, and honors trust fallback after tampering
      (`build.ps1 -CrossTests`).
- [x] **Phase 3a** — Worker copy engine: DirectoryScanner, journal-merged
      resumable copies, the full rescue-pass pipeline (FastScan / TrimSweep /
      TrimSweepReverse / Scrape / RetryBad with failure splitting and density
      adaptation), salvage fill, retry/backoff with contention + availability +
      source-mutation policies, fragile-media mode, media identity guard,
      native CopyFileEx fast path, small-file fast path, throughput throttle,
      SHA-256/512 full + sampled verification, and the env-driven fault
      injector (`XACTCOPY_DEV_FAULT_RULES`). Proven end-to-end by InteropProbe
      scenarios A–D (clean native copy, managed+verified copy, fault-injected
      rescue/salvage with byte-exact destination + journal validation, and
      salvage-disabled failure).
- [x] **Phase 3b** — Worker remainder: ScanOnly mode (precise + fast profile
      with parallel scan workers and precise fallback), bad-range-map
      integration (load/freshness/source-match, KnownBad seeding that skips
      re-reading mapped ranges, per-file persistence + final flush), parallel
      small-file CopyFileEx phase, and adaptive buffer sizing (EWMA throughput
      + latency guards). InteropProbe scenarios E–I prove: precise scan
      journal states, fast-scan fault detection writing a map the real .NET
      `BadRangeMapStore` loads with the exact bad block, map-hinted copies
      that never read known-bad ranges (asserted via zero read-retries),
      parallel-phase counters, and adaptive-buffer byte-exact copies.
      Remaining niche: experimental raw-disk scan backend (degrades to
      standard file reads with an explicit log, matching the .NET
      non-elevated fallback).
- [x] **Phase 4a** — UI core: `WorkerSupervisor` port (spawn/connect, receive
      loop, 1 Hz heartbeat monitor with 10 s staleness + activity-stall
      formula, auto-recovery kill/restart/forced-journal-resume, channel-loss
      completion), `RecoveryService` + `RecoveryStateStore`
      (.NET-compatible `runtime\recovery-state.json`, RunOnce autostart),
      DOM-preserving `AppSettings` over the .NET `settings.json` (unknown keys
      survive native saves), and a functional dark-themed Win32 main window
      (`XactCopyNative.exe`): folder pickers, mode/engine/overwrite/verify
      selections, start/pause/resume/cancel, dual progress, live log,
      crash-resume prompt (the UI binary is `XactCopy.exe`). Headless tests
      (`xactcopy_supervisor_tests`) cover
      settings preservation, recovery state round-trip, a clean supervised
      job, and kill-mid-job auto-recovery with byte-exact resume.
- [x] **Phase 4b (theming)** — shared theme layer (`src/ui/theme.h`, adapted
      from the AxiomCompress gui components): dark/light palettes with accent
      blending and readable-selection math, system dark-mode + high-contrast
      detection with live `ImmersiveColorSet` tracking, dark title bar +
      `DarkMode_Explorer`/`DarkMode_CFD` control theming with the uxtheme
      ordinal-135 app opt-in, owner-drawn buttons/checkboxes/combo items with
      hot/pressed/focus states, accent-tinted progress bars, severity-colored
      log rendering (error/warning/success/supervisor classes honoring
      `UiColorizeLogBySeverity` + `LogFontFamily`/`LogFontSizePoints`), and a
      single-instance guard that surfaces the existing window. DPI: process is
      Per-Monitor-V2 aware; `WM_DPICHANGED` adopts the suggested bounds,
      rebuilds fonts for the new monitor's DPI, and relayouts (no MFC needed —
      MFC has no dark-mode support and notoriously weak PMv2 dialog scaling;
      plain Win32 with explicit DPI handling is the recommended path and keeps
      the dual g++/MSVC toolchain). Known cosmetic gap: classic combo dropdown
      arrows (same limitation AxiomCompress accepted; needs fully custom combo
      chrome).
- [x] **Phase 4c (manifest + dialogs)** — embedded application manifest
      (ComCtl32 v6 + PerMonitorV2, `src/ui/app.manifest` via windres/rc.exe on
      both toolchains) which also completes the dark combo chrome (arrows now
      themed via v6 + `DarkMode_CFD`); progress bars keep accent colors by
      stripping visual styles per-control. Themed modal **Settings dialog**
      (appearance: theme/accent/log colorize/log font; job defaults: engine,
      verification, overwrite, buffer, retries, salvage/resume/map/adaptive/
      timestamps) with DOM-preserving save and live theme re-apply, and a
      themed **About dialog** — both on a shared `ModalHost` (`src/ui/
      dialogs.h`) that disables the owner and never leaks `WM_QUIT`.
- [x] **Phase 4d (menus + full settings)** — dark **menu bar** with the
      complete .NET MainForm structure (File: start/pause/resume/cancel, open
      journal/crash folders, exit; Tools: scan, settings, updates; Jobs: save
      as job, job manager, resume interrupted, run queued; Help: about) using
      ForceDark popups + UAH menubar owner-painting, with job-state-aware
      enable/disable (unported commands are greyed). The **Settings dialog**
      now mirrors the full .NET SettingsForm: 8 navigation pages (Appearance,
      Copy Defaults, Performance, Diagnostics, Verification, Updates,
      Recovery & Startup, Explorer Integration) generated from a declarative
      ~65-field table with staged edits across pages and DOM-preserving save.
- [x] **Phase 4e** — UI parity remainder. `JobCatalogStore` + `ManagedJob*`
      models (`src/storage/job_catalog.h`) with plain-STJ indented persistence,
      normalization/legacy-queue migration/dedup, and bidirectional
      cross-compat (`-CrossTests` catalog round-trip both ways). `JobManagerService`
      (`src/ui/job_manager.h`): saved jobs, run queue (enqueue/remove/move/
      clear/dequeue), run lifecycle marks, and history — headless-tested. Job
      Manager dialog + `TextPromptDialog` (`src/ui/job_manager_dialog.h`): a
      dark ListView console (custom-draw state colors + dark header) with view/
      status/search filters, details pane, the full action set, and 3 s
      auto-refresh. Main-window wiring: Save-As-Job / Job-Manager /
      Run-Next-Queued menu items un-greyed and functional; every copy/scan
      records a managed run (running/paused/resumed/completed/interrupted);
      queue auto-drains after each run and on startup
      (`AutoRunQueuedJobsOnStartup`); leftover Running/Paused runs are marked
      interrupted at startup. `ExplorerIntegrationService`
      (`src/ui/explorer_integration.h`) registers the HKCU shell verbs
      byte-identically to .NET, synced from `EnableExplorerContextMenu`, with
      `--from-explorer[-folder]` launch args feeding source-folder or
      selected-items (`SelectedRelativePaths`) mode. Missing main-window
      controls added: continue-on-error / skip-known-bad / wait-for-media /
      fragile checks, buffer / retries / timeout numerics, and the telemetry
      strip (EWMA throughput + running avg, smoothed-speed ETA, buffer
      utilization, rescue pass/regions/remaining, job summary, journal path,
      and a `UiShowDiagnostics`-gated diagnostics line).
- [x] **Phase 5** — updater + packaging + polish. In-app updater
      (`update_service.h` + `update_dialog.h`): GitHub-release check, download,
      SHA-256 verify, and in-place apply (exe → run; zip → wait/backup/robocopy/
      relaunch script). Taskbar progress (ITaskbarList3). App icon on the exe +
      every title bar (`app.rc` IDI_XACTCOPY + VERSIONINFO, `app_icon.h`); the
      UI binary is `XactCopy.exe`. Axiom-style About dialog + beautified
      Settings nav rail. Comprehensive Explorer shell integration
      (`explorer_integration.h`): copy verbs, a "Scan for Bad Blocks" verb
      (`--scan-from-explorer`), Applications\XactCopy.exe (Open-with/drag-drop),
      App Paths, and a Send-To shortcut. Inno Setup release pipeline
      (`installer\XactCopy.iss` + `installer\build-installer.ps1`). Remaining
      cosmetic-only: scan marquee, themed recovery prompt, single-instance arg
      forwarding, a few stored-but-unapplied appearance settings. See
      `PARITY.md`.

## Layout

- `src/core/` — header-only core: `json.h` (STJ-compatible DOM/writer/parser,
  compact + indented), `dotnet_time.h` (DateTimeOffset/TimeSpan wire formats),
  `models.h` (XactCopy.Core models), `ipc.h` (envelope + messages), `pipe.h`
  (framed named-pipe transport), `crypto.h` (BCrypt SHA-256/HMAC, DPAPI,
  base64, constant-time compare).
- `src/storage/` — `storage_models.h` (journal/map models), `stores.h`
  (JobJournalStore + BadRangeMapStore), and `job_catalog.h`
  (JobCatalog models + JobCatalogStore).
- `src/worker/` — `XactCopyExecutive` native worker: `engine_support.h`
  (errors/cancellation, scanner, fault injector, timed positional I/O,
  transfer session), `engine.h` (ResilientCopyEngine), `main.cpp` (IPC host
  with telemetry throttling).
- `src/ui/` — native UI: `supervisor.h` (WorkerSupervisor), `recovery.h`
  (RecoveryService/StateStore + RunOnce + LaunchOptions), `settings.h`
  (DOM-preserving settings), `theme.h` (shared dark/light theme layer),
  `dialogs.h` (Settings + About on a shared ModalHost), `job_manager.h`
  (JobManagerService), `job_manager_dialog.h` (Job Manager + TextPromptDialog),
  `explorer_integration.h` (comprehensive shell integration), `update_service.h`
  + `update_dialog.h` (in-app updater), `taskbar_progress.h`, `app_icon.h`,
  `main_window.cpp` (`XactCopy.exe` Win32 shell).
- `installer/` — Inno Setup script (`XactCopy.iss`) + `build-installer.ps1`
  wrapper that builds the binaries and produces `XactCopySetup-<ver>-win-x64.exe`.
- `tests/` — unit tests + golden files generated by the .NET serializer.
- `tools/GoldenGen/` — .NET console app that regenerates `tests/golden/` with
  the real System.Text.Json output (run after changing serialized shapes).
- `tools/InteropProbe/` — .NET console app that drives the native worker
  through the genuine XactCopy.Core IPC stack and asserts the conversation.
- `tools/StorageProbe/` — .NET console app that writes/verifies storage
  artifacts with the real .NET stores for the `-CrossTests` suite.

## Build and test

```powershell
# g++ (default; MSYS2 mingw64) — builds worker + tests, runs tests incl. goldens
.\build.ps1 -RunTests

# MSVC (VS18 Insiders)
.\build.ps1 -Compiler msvc -RunTests

# Bidirectional storage compatibility (.NET stores <-> native stores)
.\build.ps1 -CrossTests
```

Binaries land in the project's own `build\` folder.

Regenerate goldens and run the end-to-end interop probe:

```powershell
cd tools\GoldenGen; dotnet run -c Release -- ..\..\tests\golden
cd ..\InteropProbe; dotnet run -c Release -- ..\..\build\XactCopyExecutive.exe
```

## Wire-format compatibility notes (validated by goldens)

- Envelope: `{ProtocolVersion, MessageType, CorrelationId, SentUtc, Payload}`,
  compact, PascalCase, property order = declaration order.
- Frames: 4-byte little-endian length prefix, 16 MB ceiling.
- Strings: System.Text.Json default encoder — short escapes for
  `\b \t \n \f \r` and `\\`; `\uXXXX` (uppercase hex) for `"` `<` `>` `&` `'`
  `+` `` ` ``, other control chars, and all non-ASCII (surrogate pairs for
  astral planes).
- `DateTimeOffset`: `yyyy-MM-ddTHH:mm:ss[.fffffff]±hh:mm`, fraction omitted
  when zero, trailing zeros trimmed, written through the STJ fast path (no
  encoder escaping of `+`).
- `TimeSpan`: constant "c" format `[-][d.]hh:mm:ss[.fffffff]`, fraction always
  7 digits when non-zero.
- Enums: member-name strings, case-insensitive on read, integers accepted.
- Worker pipe: server side, byte mode, single instance, overlapped, DACL =
  GENERIC_ALL for the current user SID only (`PipeOptions.CurrentUserOnly`
  equivalent); launched with `--pipe <name>`; no handshake message — connect,
  then log event + 1 Hz heartbeats.
- Storage serializer: `WriteIndented = true` → 2-space indent, `"key": value`,
  CRLF newlines, empty containers inline (`{}`/`[]`).
- Map envelope: `{SchemaVersion, SavedUtcTicks, PayloadSha256, Signature,
  Payload}`; `PayloadSha256` = SHA-256 of the *standalone* indented payload
  bytes (both sides re-serialize on load — this is why indented byte parity is
  mandatory); `Signature` = base64 HMAC-SHA256 of
  `"{payloadSha}|{savedTicks}|{pathFingerprint}"`.
- Journal trust: snapshot hash = SHA-256 of raw file bytes (formatting-free);
  ledger frames = `magic 0x58434A4C, version 1, len, indented-JSON record,
  FNV-1a-32 checksum` (all LE); record hash =
  `sha256hex("{Seq}|{Ticks}|{SnapshotHash}|{SnapshotLen}|{PrevHash}")`;
  record/anchor signatures = base64 HMAC over hash + path fingerprint.
- Path fingerprint = `sha256hex(UTF8(ToUpperInvariant(GetFullPath(path))))`.
- HMAC keys in `%LOCALAPPDATA%\XactCopy\security\`: `journal-hmac.key` raw
  32 bytes; `badmap-hmac.key` DPAPI-protected (CurrentUser), legacy raw
  accepted + migrated. Both sides share the same key files.
