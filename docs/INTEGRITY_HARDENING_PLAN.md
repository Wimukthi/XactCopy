# XactCopy integrity hardening plan

All thirteen delivery phases below are complete as of 2.0.0.8. The **objective**
and **non-negotiable invariants** are the durable part: they are the contract
new work must not weaken. The phase list is kept as the record of how the
contract was reached, and **remaining architectural boundaries** states what
XactCopy deliberately does not promise.

## Objective

Make the result of a copy unambiguous and keep a failed, interrupted, stale,
or partially recovered operation from replacing or being mistaken for an exact
backup. Performance features remain available, but they must not bypass the
integrity contract.

## Non-negotiable invariants

1. A copy is a clean success only when every enumerated source file is present,
   no file was skipped or salvaged, and the destination was published after
   the selected verification policy passed.
2. Existing destination paths are never truncated in place. Work is performed
   in a same-directory staging path, flushed, and atomically replaced only
   after the file is ready.
3. A journal range is not evidence that destination bytes are correct. Partial
   coverage is discarded after an interrupted copy; a completed entry is
   reusable only after a full source/destination validation.
4. A source replacement is detected by file identity as well as size and
   timestamp. A changed source cannot inherit old journal or bad-range hints.
5. Native `CopyFileExW` is an explicit fast path. Auto mode uses the managed
   retry/rescue path whenever the configured policy requires it; media waits,
   fragile mode, lock waits, access-denied contention, and source-mutation
   handling always force managed I/O.
6. Raw extent mappings and bad-range-map hints are usable only when their
   source identity is still valid. Unsupported raw layouts fall back to
   standard reads without silently becoming a different scan result.
7. Enumeration errors are part of the result. Continue-on-error may continue
   processing a partial listing, but it cannot produce a clean success.

## Delivery phases

### Phase 1: result and conflict semantics — complete

- Refuse `OverwritePolicy::Ask` conflicts until an actual interactive prompt
  protocol exists; never silently treat Ask as overwrite.
- Make copy results non-clean when files are skipped, recovered, or failed.
- Show failed, recovered, and skipped counts in the main window and Job
  Manager summaries.
- Treat source-mutation skips as failed/incomplete journal entries.
- Reject `AllowJournalRootRemap` and the unsafe combination of
  `TreatAccessDeniedAsContention` with salvage.
- Persist and surface a distinct `Incomplete` run status when a partial
  operation made progress but could not claim an exact success.

### Phase 2: destination transaction — complete

- Stage every single-file copy beside its final path.
- Clone an existing destination into the stage only when resumable coverage is
  needed; otherwise create an empty stage.
- Flush staged bytes and publish with same-volume `MoveFileExW` replacement.
- Apply the same transaction rule to the parallel native small-file path.
- Reject destination-inside-source jobs before enumeration, so output and
  abandoned stage files cannot become new source input.
- Clean stale, generated stage files only in the destination tree, using an
  active lock probe so a live job is never reclaimed.
- Extend the transaction to timestamps, basic file attributes, alternate data
  streams, and security descriptors. Read-only publication temporarily clears
  the bit across durable flush/replace and restores the exact source mask.

### Phase 3: journal and source identity — complete

- Persist source file index and volume serial in journal entries.
- Reset stale partial/rescued coverage after interruption.
- Reuse a completed entry only after a full hash check; otherwise recopy it.
  Verified reuse counts as a completed file, while policy skips remain
  explicitly incomplete.
- Reject same-size/same-timestamp source replacements using file identity.
- Destination identity/hash metadata remains an optional future optimization;
  completed reuse currently performs content validation instead.
- Keep legacy journals readable; missing identity fields mean “unknown” and
  must not be treated as a match.

### Phase 4: native and retry policy — complete

- Keep Auto on managed rescue when `MaxRetries` or salvage behavior is active.
- Keep NativeFast explicit, staged, identity-checked, and clearly labelled as
  one `CopyFileExW` attempt per file before fallback.
- Require full verification, source/media identity checks, and staging before
  publishing parallel NativeFast small files; otherwise the phase is disabled.
- Add timeout/retry diagnostics that distinguish API failure, fallback, salvage,
  and indefinite media/lock waiting.
- Treat XactCopy retries as additive to filesystem, storage-class-driver,
  controller, and device retries. Default to 2, cap all user/pass retry counts
  at 32, warn above 3, and reject non-deterministic random salvage fill.

### Phase 5: raw scan and bad-range map — complete

- Persist and compare file identity in map entries.
- Treat legacy map entries with no identity as readable history, never as
  authoritative skip hints.
- Invalidate cached raw layouts when file index or volume serial changes.
- Keep map age finite by default; make “never expire” an explicitly warned
  advanced mode.
- Do not persist map changes when `UpdateBadRangeMapFromRun` is false,
  including ScanOnly.
- Require two matching bad-range observations before using a read-skip hint,
  and never persist synthetic recovered ranges as future hints.
- Add elevated raw integration coverage for mutable/replaced files and sparse
  layouts.

The elevated raw test is opt-in because it requires administrator access and a
volume that supports the relevant sparse-file and raw-volume controls. Set
`XACTCOPY_RAW_DISK_TEST=1` to run it.

### Phase 6: enumeration and metadata completeness — complete for the defined
contract

- Return a structured incomplete-scan result for `FindFirstFileExW` and
  `FindNextFileW` failures.
- Add source/destination identity checks around enumeration and commit.
- Define and test the metadata contract for timestamps, ACLs, ADS, and empty
  directories. Managed Rescue preserves file ADS and full-verifies their
  content; child-directory timestamps/basic attributes/DACLs are applied
  deepest first. Owner/group and SACL limitations are reported as warnings.
- Warn when Managed Rescue cannot preserve native-copy metadata. EFS is
  preserved or fails safely; extended attributes, compression state, sparse
  layout, and reparse-point semantics remain filesystem-specific and are
  deliberately not recreated by the managed path.

### Phase 7: defaults, UI, and compatibility — complete for current settings

- Add an integrity-oriented default profile: salvage off, continue off, full
  verification on, no bad-range skipping, and conservative worker selection.
  Missing-key fallbacks use this profile; existing explicit keys remain intact.
- Review integrity-reducing saved values at execution time. Manual runs show a
  consolidated warning; automatic queue/recovery refuses sampled/no
  verification, Continue on error, recovered overwrite, plaintext EFS, or
  unbound media identities.
- Label recovery and incomplete runs separately from exact completion.
- Keep IPC and journal readers backward compatible; new fields are optional.

Existing saved keys are intentionally preserved verbatim. Safety is enforced at
the point of use, so upgrades do not silently rewrite expert preferences and
stale profiles cannot silently become unattended jobs.

### Phase 8: recovery output, source stability, and EFS — complete

- Publish every non-exact salvage result as a visibly named recovery sidecar,
  including when the ordinary destination did not previously exist. Replacing
  an existing destination with synthetic bytes requires an explicit expert
  override.
- Exclude synthetic recovered ranges from future bad-range-map hints.
- Hold managed source handles without write sharing and compare file identity,
  length, last-write time, and NTFS change time before publication. This catches
  same-file, same-size mutations even when an application restores the old
  last-write timestamp.
- Preserve EFS encryption on managed copies. Plaintext publication requires a
  separate expert option and is reported in the result.
- Distinguish a failure after atomic publication from a pre-commit failure: the
  exact published bytes remain committed and an attribute-restore failure is a
  metadata warning, not a fictional rollback.

### Phase 9: media binding and unattended persistence — complete

- Bind saved jobs to source/destination volume GUID plus serial; retain legacy
  serial-only identity compatibility without treating two different modern
  identities as equivalent.
- Persist worker-confirmed identities into crash-recovery state immediately and
  carry them in progress IPC.
- Resolve junctions, mount points, symbolic links, and nearest existing
  ancestors before accepting source/destination separation.
- Store saved jobs/queue/history in a path-bound HMAC-SHA256 envelope with two
  rotations and a mirror. Legacy plain catalogs migrate on save; tampered
  primary content falls back to an authenticated candidate.
- Roll back in-memory catalog mutations when durable persistence fails, so a
  failed dequeue cannot execute and then reappear after restart.

### Phase 10: filesystem fidelity reporting — complete for the defined contract

- Preserve and full-verify file alternate streams, preserve EFS or fail, and
  apply child-directory timestamps/basic attributes/DACLs after file work.
- Detect source hard links without adding a second enumeration pass. Current
  output uses independent exact files and carries an explicit fidelity warning.
- Report directory ADS, hard-link topology, privilege-limited owner/group/SACL,
  and unsupported sparse/compressed/reparse semantics instead of silently
  claiming a metadata-perfect clone.

### Phase 11: regression expansion and validation — complete

- Cover safe missing-field defaults, attended/unattended policy, signed catalog
  migration/tamper fallback, confirmed bad-map hints, synthetic-range exclusion,
  same-size/timestamp mutation, physical path aliases, random fill/retry limits,
  EFS preservation/downgrade, ADS verification, directory metadata, hard-link
  reporting, and post-publication metadata faults.
- Keep the elevated raw-disk integration opt-in; all non-elevated MSVC suites
  run in the normal validation pass.

### Phase 12: assurance, cancellation, and product semantics — complete

- Separate hard failures, integrity-assurance notices, and metadata-fidelity
  notices in result IPC/history instead of appending every limitation to
  `ErrorMessage`.
- Report actual read, write, verification, skipped, and journal-reused bytes;
  skipped work no longer inflates transfer speed.
- Reuse the incremental source hash from healthy managed copies during full
  verification and apply timeout per hash I/O operation rather than imposing an
  arbitrary whole-file deadline.
- Make cancellation two-stage: cooperative on the first request, then forced
  after a bounded grace period or an explicit second request.
- Keep attended-only queued jobs in the queue with a concrete reason and allow
  later safe jobs to run. Persistence failures in the job catalog, recovery
  state, and RunOnce registration are surfaced to the operations log.
- Preserve the original extension on non-exact recovery output and write an
  adjacent readable manifest containing source/media identity, fill policy, and
  every synthetic range.
- Replace the misleading Ask/scan/engine presentation with Stop on conflict,
  Verified Copy/Recover Media profiles, and Assess Readable Files. Assessment
  documentation now warns against spending a dying device's remaining reads on
  diagnosis before recovering high-value data.
- Restrict updater feeds/assets to HTTPS, exact eligible package names, and
  complete expected byte counts. Portable updates validate archive paths and
  package version, back up only affected files, validate after replacement, and
  roll back on failure.
- Verify both ends of worker IPC against the spawned supervisor/worker PIDs,
  discover the worker only as an exact sibling executable, and use exact system
  utility paths for forced shutdown/update helpers.
- Apply monitor DPI plus application scale/density consistently to layout,
  owner-drawn glyphs, combo padding, and dialog controls across monitor moves.

### Phase 13: durable trust roots and final UX closure -- complete

- Treat a settings, journal, map, catalog, or key write as durable only after
  the temporary file flush and atomic replacement both succeed; preserve the
  original Win32 error through cleanup.
- Never replace an unreadable existing HMAC key. Create new keys with a
  no-replace race and atomically migrate legacy raw map/catalog keys to
  current-user DPAPI without changing their value. Keep the journal key's raw
  encoding for compatibility with already-installed 2.x builds.
- Require at least one complete journal snapshot/ledger chain per save, surface
  reduced redundancy, repair a stale or trailing-torn peer from the trusted
  chain, and refuse unsigned JSON downgrade once a ledger or anchor exists.
- Bind bad-range hints and cached raw-volume extents to NTFS change time as well
  as file identity, length, and last-write time.
- Bind completed scan-journal entries to the same file identity and NTFS change
  time before reusing prior scan coverage, even if length and last-write time
  were restored after an in-place mutation.
- Bound updater packages at 2 GiB, reject Windows archive aliases and duplicate
  extraction paths, refuse target-tree reparse points, hash-verify every copied
  package file, bound synchronous cancellation latency, and clean random private
  update directories without following reparse points.
- Normalize unsupported settings values before saving an unvisited page; expose
  every supported theme/scale choice, make progress-bar styles visually
  distinct, and give main, settings, Job Manager, About, and prompt controls
  usage-oriented tooltips.

## Remaining architectural boundaries

These are explicit product boundaries, not silent correctness claims:

1. **Whole-job atomicity.** File replacement is atomic, but an arbitrary folder
   update is not. Earlier verified files remain committed if a later file fails.
   A true all-or-nothing mode needs a versioned destination root or filesystem
   snapshot plus an explicit final switch; it cannot be built from another
   `CopyFileExW` retry flag.
2. **Hard-link topology.** Linked names currently become independent exact
   files. A future opt-in topology-preservation mode needs a source identity
   graph, destination-filesystem capability checks, overwrite conflict rules,
   and journal representation before it can safely call `CreateHardLinkW`.
3. **Directory alternate streams and filesystem-specific layouts.** These are
   reported but not recreated by Managed Rescue. Native Fast may preserve more
   filesystem-native state, but only content and the documented metadata
   contract are asserted by XactCopy.
4. **Adversary model.** DPAPI/HMAC detects corruption and unauthenticated edits;
   it is not a security boundary against arbitrary code already executing as
   the same Windows user.

## Implementation notes for the scan-start regression

Large Fast scans no longer block the UI while writing their initial journal
checkpoint. The checkpoint is written on a background thread, the worker pool
starts immediately, and the checkpoint is joined before final journal
publication. Resume preparation now logs and cancellation-checks source
identity, resume coverage, source-index, and journal-entry phases, so a large
scan cannot appear frozen between `Loaded journal` and the first file event.
Fresh journal entries do not perform a second per-file handle lookup: source
identity comparison is limited to existing journal entries that already carry
an identity, while copy paths persist the identity they capture at copy time.

## Required regression matrix

- Ask conflict, overwrite failure, and atomic replacement with a sentinel target.
- Read/write/timeout/offline fault injection with and without salvage.
- Fragile skip, source disappearance, access denied, and continue-on-error.
- Same-length destination corruption and same-size/same-timestamp source
  replacement during resume.
- Partial journal coverage after crash and completed journal validation.
- Root remap rejection and legacy journal compatibility.
- Bad-range-map identity, age, recovered-range, and ScanOnly update policy.
- Native/parallel media remount and cancellation behavior.
- Enumeration start/end failures and partial-listing result semantics.
- Raw extent cache invalidation, elevated reads, sparse files, and fallback.
- Durable flush/commit ordering and abandoned-stage cleanup.
- Post-publish attribute failure, EFS preservation/plaintext override, ADS full
  verification, child-directory metadata, and hard-link reporting.
- Signed catalog legacy migration, tamper fallback, media-bound auto resume,
  and rejection of attended-only policies during unattended execution.

The current sequential MSVC validation passes the core, storage, supervisor,
and worker suites. It covers the safety and fidelity paths listed above,
including live EFS and NTFS hard-link tests on a capable test volume. The
elevated raw integration remains opt-in as described above.
