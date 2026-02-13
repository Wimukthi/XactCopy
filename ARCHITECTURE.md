# XactCopy Architecture

## Goal

Copy files from unstable media with maximal forward progress, predictable recovery, and minimal operator babysitting.

## Reliability Contract

- The app must stay responsive while copying.
- Single-file failures must not crash the process.
- Copy progress must survive app restarts through a durable journal.
- I/O stalls must be bounded by retry + timeout policy.
- Worker process failure/stall must be recoverable without restarting the UI.
- Operators must be able to choose strict mode or salvage mode.

## Main Components

### `XactCopy.UI` (`MainForm` + `WorkerSupervisor`)

- GUI for configuring source/destination and reliability options.
- Starts and cancels copy jobs asynchronously.
- Displays live progress, per-file status, and event logs.
- Supervises the worker with heartbeat timeout detection.
- Restarts worker automatically and reissues the active job command when safe.

### `XactCopy.Worker` (`Program` + `ResilientCopyService`)

- Hosts a named-pipe server and executes copy jobs out-of-process.
- Emits heartbeat, progress, log, result, and fatal messages over IPC.
- Runs block-based copy with retry/timeout/salvage semantics.

### `XactCopy.Core`

- Shared copy models.
- Versioned IPC envelope/message contract.
- Length-prefixed JSON message framing utilities.

### `XactCopy.Storage` (`JobJournalStore`)

- Atomic journal persistence (temp-write + replace).
- Deterministic job IDs from source/destination pair.
- Resume state for file bytes copied and recovered ranges.

### `DirectoryScanner` (worker)

- Enumerates the source tree safely.
- Catches per-directory/per-file enumeration failures and logs them.

## Copy Pipeline

1. Validate options and normalize paths.
2. Scan source files and build/merge journal.
3. For each file:
   - Resume from saved offset when possible.
   - Read block with timeout and retries.
   - If read still fails and salvage is enabled, zero-fill block and continue.
   - Write block with timeout and retries.
   - Persist journal periodically and on important transitions.
4. Optionally verify with SHA-256 (only when no salvage blocks were used).
5. Emit final result summary.

## Failure Behavior

- Read timeout or transient `IOException`: retry with exponential backoff.
- Persistent read failure:
  - Salvage enabled: fill block with zeros and continue.
  - Salvage disabled: fail file.
- Write timeout/failure: retry, then fail file if retries are exhausted.
- File failure:
  - Continue-on-error enabled: proceed to next file.
  - Disabled: stop job immediately.
- Worker process stall/crash: supervisor restarts worker and resumes from journal.
- Unexpected UI/process exception: captured to crash log path.

## Known Limits

- If the OS kernel blocks file open/read/write at a level where cancellation cannot complete, the operation can still stall longer than expected.
- Salvage mode preserves job continuity, not original bytes in unreadable sectors.
- Verification requires readable source data and is skipped for recovered files.

## Future Hardening

- Add manifest export for forensic reporting of recovered/failed byte ranges.
- Add volume-health telemetry and adaptive timeout profiles.
- Add multi-worker scheduling for controlled parallel copy on healthy media.
