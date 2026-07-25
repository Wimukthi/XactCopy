# Architecture

XactCopy is two cooperating processes joined by a small, versioned message
protocol. This split is what makes it resilient: the part that talks to failing
hardware is isolated from the part you interact with, and either can recover if
the other misbehaves.

## Two processes

| Process | Binary | Role |
| --- | --- | --- |
| **UI / supervisor** | `XactCopy.exe` | The window you see. Configures jobs, launches and monitors the worker, renders progress and logs, and owns recovery. |
| **Worker** | `XactCopyExecutive.exe` | Does the actual reading and writing: the copy engine, rescue passes, salvage, verification, and bad-block scanning. |

The worker is a child process the UI starts on demand. Keeping the risky I/O in a
separate process means a hang or crash while wrestling a dead drive can't take the
UI down with it — the supervisor simply restarts the worker and resumes.

## Talking between them

The UI and worker communicate over a **named pipe** using a compact,
versioned **JSON** protocol:

- The worker is the pipe **server** (`--pipe <name>`), created in byte mode as a
  single instance, secured so **only the current user** can connect.
- Messages are length-prefixed frames (4-byte little-endian length, with a
  ceiling to reject oversized frames) wrapping a JSON envelope:
  `{ProtocolVersion, MessageType, CorrelationId, SentUtc, Payload}`.
- Once connected, the worker emits a connect log event and a **1 Hz heartbeat**;
  job progress and telemetry stream back as messages while a job runs.

Because the boundary is an explicit, versioned protocol rather than in-process
calls, each side can be developed and tested independently.

## Supervision and auto-recovery

The supervisor watches the worker's heartbeat and job activity. If the worker
stops heart-beating or stalls beyond a threshold, the supervisor **kills and
restarts** it and re-issues the job with *resume-from-journal* forced, so the
transfer continues from where it stopped rather than starting over. Loss of the
pipe is treated as a completed/failed channel and handled the same way.

## Durable state: journals

Every job is backed by a **journal** that records what has been copied and
verified. Journals are written through atomically (temp file + rename), kept with
rotating backups and mirror copies, and made tamper-evident:

- Each snapshot is content-hashed (SHA‑256 of the raw file bytes).
- An append-only **ledger** hash-chains snapshots together, with periodic signed
  **anchors**.
- Ledger records and anchors are **HMAC-signed** so the app can detect tampering
  and fall back to a trusted state.

This is what powers **resume** and **crash recovery**: on restart the journal is
merged with the current source to pick up precisely where the last run left off.

## Bad-range maps

A bad-range map records the unreadable byte ranges of a specific source. Maps are
stored in a signed envelope — `{SchemaVersion, SavedUtcTicks, PayloadSha256,
Signature, Payload}` — where the payload hash and an HMAC signature (bound to a
fingerprint of the source path) let XactCopy trust a map only if it matches the
media it was made for. Enabling *Use bad-range map* / *Skip known-bad ranges*
lets a later copy avoid re-reading regions already known to be dead.

## Recovery state

A small recovery-state file tracks the currently active run and whether the last
shutdown was clean. On launch, an unclean shutdown with an active run triggers the
crash-resume prompt; leftover runs still marked *running* are marked
*interrupted*.

## Security and storage locations

Everything lives under `%LOCALAPPDATA%\XactCopy\`:

- **Journals**, **bad-range maps**, the **job catalog**, and **recovery state**.
- **HMAC keys** under `security\`: the journal key as raw bytes, the map key
  **DPAPI-protected** for the current user.
- **Settings** as a JSON document whose unknown keys are preserved across
  upgrades, so newer and older builds can share one file.

## The application layer

The UI is a native **C++/Win32** application (no managed runtime):

- A shared theming layer provides dark/light palettes, accent blending, dark
  title bars and menus, and owner-drawn controls, tracking the system theme live.
- The process is **Per-Monitor-V2 DPI aware**; fonts and layout rebuild on
  monitor changes.
- Long-running work is kept off the UI thread; progress messages are coalesced so
  a job over millions of small files can't flood the message queue.

## On-disk and wire formats

XactCopy's journals, maps, and IPC messages use fixed, documented encodings so the
formats are stable and inspectable:

- **JSON** is emitted compactly for IPC and indented (2-space, CRLF) for stored
  artifacts, with a serializer that matches .NET's System.Text.Json escaping and
  number/date formatting.
- **Timestamps** use ISO‑8601 `DateTimeOffset` and the constant `TimeSpan`
  ("c") format.
- **Path fingerprints** are `SHA‑256(UTF‑8(ToUpperInvariant(fullPath)))`.
- **Signatures** are base64 HMAC‑SHA256 over the relevant hash plus the path
  fingerprint.

These formats are exercised by the golden-file and cross-compatibility tests
described in [CONTRIBUTING](../CONTRIBUTING.md).
