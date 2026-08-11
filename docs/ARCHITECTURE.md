# Architecture

XactCopy is two cooperating processes joined by a small, versioned message
protocol. That split is what makes it resilient: the part that talks to failing
hardware is isolated from the part you interact with, and either can recover if
the other misbehaves.

## Two processes

| Process | Binary | Role |
| --- | --- | --- |
| **UI / supervisor** | `XactCopy.exe` | The window you see. Configures jobs, launches and monitors the worker, renders progress and logs, and owns recovery. |
| **Worker** | `XactCopyExecutive.exe` | Does the actual reading and writing: the copy engine, rescue passes, salvage, verification, and read-only file-readability assessment. |

The worker is a child process the UI starts on demand. Keeping the risky I/O in a
separate process means a hang or crash while wrestling a dead drive can't take
the UI down with it — the supervisor restarts the worker and resumes.

## The IPC contract

The two processes talk over a **named pipe** using a versioned **JSON** protocol.

- The worker is the pipe **server** (`--pipe <name>`), created in byte mode as a
  single instance, with a security descriptor that lets **only the current user**
  connect.
- Messages are length-prefixed frames — a 4-byte little-endian length, capped so
  oversized frames are rejected — wrapping a JSON envelope:
  `{ProtocolVersion, MessageType, CorrelationId, SentUtc, Payload}`.
- `ProtocolVersion` is `1`. A frame carrying any other version is rejected
  rather than guessed at.
- Once connected, the worker emits a connect log event and a **1 Hz heartbeat**.
  Job progress and telemetry stream back as messages while a job runs.

Because the boundary is an explicit, versioned protocol rather than in-process
calls, each side can be developed and tested independently. `tests/golden/`
pins the envelope encodings byte-for-byte.

## Supervision and auto-recovery

The supervisor watches two things:

- **The heartbeat.** No heartbeat for **10 seconds** means the worker is gone or
  wedged.
- **Job activity.** A running, unpaused job that makes no progress for longer
  than its activity budget is reported as potentially stalled. That budget is
  derived from the job's own settings — roughly *4 × operation timeout + 2 ×
  maximum retry delay + 10 s*, clamped to a sane range.

A healthy heartbeat is proof that the worker process and IPC loop are alive, so
an activity timeout emits a warning but does **not** kill a possibly valid,
blocking Windows I/O request. A stale heartbeat, a lost pipe, or a worker exit
does restart the worker and re-issue the job with *resume-from-journal* forced.
Automatic recovery is capped at three attempts per job; a repeatedly crashing
worker is then failed visibly instead of being relaunched forever. Recovery
threads are owned and joined by the supervisor so they cannot outlive the UI.

## Durable state: journals

Every job is backed by a **journal** recording what has been copied and verified.
Journals are written through atomically (temp file, then rename), kept with
rotating backups and mirror copies, and made tamper-evident:

- Each snapshot is content-hashed (SHA-256 over the raw file bytes).
- An append-only **ledger** hash-chains snapshots together, with periodic signed
  **anchors**.
- Ledger records and anchors are **HMAC-signed**, so tampering is detectable and
  the app can fall back to a trusted state.

Primary and mirror snapshots have independent ledgers. A save is accepted only
when at least one complete snapshot/ledger chain is durable; reduced redundancy
is reported and the surviving chain is used to repair a missing, stale, or
trailing-torn peer. Once a signed ledger or anchor exists, a parseable unsigned
JSON file cannot be used as a downgrade fallback.

This is what powers **resume** and **crash recovery**: on restart the journal is
merged with the current source. Partial destination coverage is reconstructed;
a journal's completed marker permits reuse only after a fresh full source/
destination hash comparison.

Because a snapshot is rewritten in full on every flush, its cost scales with the
journal — a whole-drive job can produce tens of megabytes of JSON. Flush
frequency is
therefore *budgeted* rather than fixed: after a save costing *N* ms, the next
flush waits at least *19N*, holding journal I/O near 5% of run time at any
journal size. Small journals save in about a millisecond and keep their original
sub-second cadence. The trade is that a large job may redo up to one interval of
work after a crash, which is cheap next to rewriting the snapshot twice a second.

## Bad-range maps

A bad-range map records the unreadable byte ranges of a specific source. Maps are
stored in a signed envelope — `{SchemaVersion, SavedUtcTicks, PayloadSha256,
Signature, Payload}` — where the payload hash and an HMAC signature bound to a
fingerprint of the source path let XactCopy trust a map only if it matches the
media it was made for. *Use bad-range map* and *Skip known-bad ranges* let a
later copy avoid re-reading regions known to be dead. A read hint additionally
requires the same range to have been observed twice against the same file
identity/fingerprint. That fingerprint includes file length, last-write time,
and NTFS change time, so restoring an old user-visible timestamp does not make a
mutated file inherit stale physical-range hints. Synthetic salvage-filled ranges
are never promoted to a future skip hint.

## Destination publication and fidelity

Copy data is written to a random same-directory staging file. The worker copies
metadata, performs the selected verification, flushes the stage, and publishes
it with `MoveFileExW(MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)`. A
pre-publication error therefore cannot truncate a previous destination. If a
post-publication attribute restore fails, the already committed bytes remain a
successful exact result and the metadata failure is reported separately.

The transaction boundary is one file. Windows does not provide an atomic rename
for an arbitrary set of files spread through an existing directory, so a job
that fails after earlier files commit can leave a mixed generation. A future
whole-set transaction mode would need a versioned destination root/snapshot and
an explicit final switch; it cannot be emulated by adding more `CopyFileExW`
flags.

Managed Rescue preserves file timestamps, basic attributes, DACLs, file ADS,
and EFS encryption; full verification hashes copied ADS as well as the default
stream. Child-directory timestamps/basic attributes/DACLs are applied deepest
first after file publication. Privilege-limited owner/group/SACL fidelity,
directory ADS, hard-link topology, and unsupported sparse/compressed/reparse
semantics are surfaced as warnings instead of being silently claimed.

## Raw-volume scan backend

Scan-only jobs can optionally open a local NTFS volume read-only and resolve
each selected file's allocated extents with `FSCTL_GET_RETRIEVAL_POINTERS`.
Volume-aligned reads are translated back into the existing file-relative rescue
ranges, so journals, bad-range maps, and the IPC contract remain unchanged.
Raw access requires elevation and falls back to standard file reads for
unsupported layouts; it does not inspect free or unallocated space. Cached
extent layouts are discarded whenever file identity, length, last-write time,
or NTFS change time no longer matches the open file.

## Recovery state

A small recovery-state file tracks the currently active run and whether the last
shutdown was clean. On launch, an unclean shutdown with an active run triggers
the crash-resume prompt, and any run still marked *running* is re-marked
*interrupted*. Automatic resume/queue execution requires confirmed media
identities, full verification, Continue on error disabled, and no expert
recovered-overwrite or plaintext-EFS override. Otherwise an attended review is
mandatory.

## Saved-job catalog

Saved jobs can launch unattended work, so `jobs\catalog.json` is not trusted as
plain configuration. It is stored in an HMAC-SHA256 envelope bound to its path,
with two rotating generations and an independently committed mirror. Reads walk
the trusted candidate set and fall back after corruption/tampering. Legacy plain
catalogs remain readable and are upgraded on the next save. A dequeue is not
allowed to proceed in memory if its durable catalog update fails.

## Storage and security

Everything lives under `%LOCALAPPDATA%\XactCopy\`: journals, bad-range maps, the
job catalog, recovery state, settings, and HMAC keys under `security\`. The
journal key retains its legacy raw 32-byte representation for compatibility;
map and catalog keys are **DPAPI-protected** for the current user and migrate
legacy raw values atomically. An unreadable existing key is never silently
replaced because that would orphan every artifact signed by it.
Settings are a JSON document whose unknown keys are preserved across upgrades, so
newer and older builds can share one file. The
[User Guide](USER_GUIDE.md#where-xactcopy-stores-its-data) has the full layout.

## The application layer

The UI is a native **C++/Win32** application with no managed runtime:

- `Wimukthi.Win32Theme` provides the reusable Windows integration — process and
  window dark-mode opt-in, title bars, menus, common controls, High Contrast, and
  live system-theme changes. XactCopy keeps only what is specific to it: its
  palette, accent blending, icons, progress bars, and other owner drawing.
- The process is **Per-Monitor-V2 DPI aware**; fonts and layout rebuild on
  monitor changes.
- Long-running work stays off the UI thread, and progress messages are coalesced
  so a job over millions of small files can't flood the message queue.

The UI is a single translation unit (`src/ui/main_window.cpp` plus header-only
modules), so any header change rebuilds it.

## On-disk and wire formats

Journals, maps, and IPC messages use fixed, documented encodings, so the formats
are stable and inspectable.

- **JSON** is emitted compactly for IPC and indented (2-space, CRLF) for stored
  artifacts, by a serializer that matches .NET's `System.Text.Json` escaping and
  number/date formatting.
- **Journal file entries omit default-valued members**, and omit `RelativePath`
  when it only repeats the key the entry is stored under. Readers treat an absent
  member as its default, so this is an encoding difference rather than a schema
  one. It is worth roughly 40% of a large journal, which is otherwise dominated
  by empty strings, `false`, and `[]`.
- **Journal snapshots over 64 KB are compressed** into a container: `"XCJZ"`, a
  version and algorithm byte, and the uncompressed length, followed by an
  XPRESS_HUFF stream. Smaller journals stay plain JSON, so the common case is
  still readable in a text editor. Reads sniff the magic, so plain-JSON snapshots
  — written by older builds, or by the retired .NET build — keep loading
  unchanged. The ledger hashes the bytes as written, compressed or not.
  Job Manager's **Open Journal** action validates the trusted snapshot, exports
  a readable JSON view under `%LOCALAPPDATA%\XactCopy\journal-views\`, and opens
  that view; the resumable XCJZ snapshot is never modified.
  Bad-range maps, the job catalog, ledgers, and anchors are all small and stay
  uncompressed and inspectable.
- **Timestamps** use ISO-8601 `DateTimeOffset` and the constant `TimeSpan` ("c")
  format.
- **Path fingerprints** are `SHA-256(UTF-8(ToUpperInvariant(fullPath)))`.
- **Signatures** are base64 HMAC-SHA256 over the relevant hash, save time, and
  path fingerprint. Bad-range maps and saved-job catalogs use signed envelopes;
  catalog payloads written by older builds are accepted only through the legacy
  migration path.

These formats are pinned by the golden files described in
[CONTRIBUTING](../CONTRIBUTING.md), which the test suites byte-compare against.

## Source layout

| Path | Contents |
| --- | --- |
| `src/core/` | Header-only core: JSON, date/time wire formats, models, the IPC envelope and messages, the framed named-pipe transport, and crypto (BCrypt SHA-256/HMAC, DPAPI, base64) |
| `src/storage/` | Journal store, bad-range-map store, and job catalog |
| `src/worker/` | The `XactCopyExecutive` worker: copy engine, scanner, rescue/salvage pipeline, and the IPC host |
| `src/ui/` | The `XactCopy.exe` UI: supervisor, recovery, settings, theming, main window and dialogs, shell integration, updater, and resources |
| `tests/` | Unit-test suites plus the golden fixtures |
