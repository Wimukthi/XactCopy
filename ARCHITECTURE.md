# XactCopy Architecture

## Objective

XactCopy is a resilient Windows file mover/scanner designed for unstable sources and destinations (failing disks, intermittent USB/network paths, lock contention, and unclean exits). The design prioritizes:

- Maximum forward progress.
- Recoverability after interruption.
- Clear operator control of risk/performance tradeoffs.
- UI responsiveness while heavy I/O runs in a separate process.

## Runtime Topology

### `XactCopy.UI` (WinForms host)

- Owns user interaction, settings, theming, queue/jobs, and updater UX.
- Dispatches runs through `WorkerSupervisor`.
- Persists recovery metadata and can auto-prompt resume after unclean exits.
- Applies throttled progress/log updates to keep UI responsive under high event rates.

### `XactCopy.Worker` (`XactCopyExecutive`)

- Out-of-process execution engine connected via versioned named-pipe IPC.
- Runs copy/scan pipelines, retry/backoff, rescue passes, verification, and persistence flushes.
- Emits structured progress, logs, heartbeats, completion, and fatal events.

### Shared libraries

- `XactCopy.Core`: run options, contracts, enums, and shared models.
- `XactCopy.Storage`: durable stores for journal and bad-range map state.

## Core Data Stores

### Job Journal (`JobJournalStore`)

- Tracks per-file state and range coverage so runs can resume.
- Uses atomic write/replace semantics with integrity hardening:
  - Multi-generation backups.
  - Mirror snapshots.
  - Signed/hash-chained ledger metadata.
- Recovery loads prefer trusted/latest candidates and gracefully fall back across snapshot sets.

### Bad-Range Map (`BadRangeMapStore`)

- Source-scoped map of unreadable byte ranges discovered by scan/copy.
- Used to seed rescue state and optionally skip known-bad ranges on later runs.
- Persisted with:
  - Atomic writes.
  - Rotating backups.
  - Mirror snapshots.
  - Signed envelope validation with per-machine HMAC key.
- Legacy plain JSON maps remain readable for backward compatibility.

## Operation Modes

### `Copy`

- Source-to-destination transfer with overwrite policy, verification policy, salvage policy, and retry profile.
- Existing destination handling (`overwrite/skip/newer/ask`) is applied.
- Optional `Rescue Engine` passes recover difficult regions.

### `ScanOnly`

- Read-only analysis mode (no destination file writes).
- Enumerates and reads source content to detect bad ranges.
- Updates bad-range map for future copy optimization/hardening.
- Copy-only skip/overwrite rules are not applied.

## Worker Pipeline

1. Validate options and normalize roots (root-safe normalization preserves `D:\` semantics).
2. Resolve source file set (`DirectoryScanner` + optional Explorer selected-items subset).
3. Load/merge journal state.
4. Load bad-range map and seed known-bad ranges where enabled.
5. Execute file loop:
   - Availability checks (source/destination).
   - Media identity guard checks (serial/share identity).
   - Mode-specific behavior (`Copy` or `ScanOnly`).
   - Retry/backoff on transient failures.
   - Optional salvage fill, verification, and rescue passes.
6. Periodic + terminal persistence flushes (journal/map).
7. Emit completion result.

## Reliability and Hardening Features

- Supervisor-managed worker restart and run continuity.
- Pause/resume/cancel support via execution control.
- Policy-driven handling for:
  - Media disappearance (`WaitForMediaAvailability`).
  - File lock contention (`WaitForFileLockRelease`, `TreatAccessDeniedAsContention`).
  - Source mutation/disappearance (`SourceMutationPolicy`).
- Media identity capture/validation to reduce wrong-target writes after remounts.
- Journal path fallback when default storage is not writable.
- Settings/run cloning paths preserve complete option snapshots for historical accuracy.

## UI and Performance Strategy

- Heavy I/O and retry loops are isolated from UI thread by worker process boundary.
- Main form receives throttled progress/log events and uses virtualized log rendering.
- Job Manager uses virtual grid patterns for large run-history datasets.
- Diagnostics strip/telemetry provides rendering and event-pressure visibility.

## Known Limits

- No user-space copier can bypass kernel/storage stack hangs in all scenarios.
- Salvage mode preserves continuity, not original unreadable bytes.
- Verification behavior depends on selected mode and salvage outcomes.
