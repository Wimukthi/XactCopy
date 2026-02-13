# XactCopy

XactCopy is a VB.NET WinForms copier built for resilient transfers on unstable storage media.

## License

GNU GPL v3.0. See `LICENSE`.

## Highlights

- Out-of-process worker process with supervisor restart handling.
- Journal-based resume and recovery after unexpected exits.
- Pause/resume/cancel controls with run-state persistence.
- Salvage mode for unreadable sectors with configurable fill pattern.
- Retry/backoff timeout controls for degraded disks.
- Adaptive buffer mode with live speed, ETA, and buffer telemetry.
- Job manager (saved jobs, queue, history).
- Explorer context menu integration (`Copy with XactCopy`).
- Dark/system/classic theme support.
- Built-in update checker.

## Repository Layout

- `src/XactCopy.UI` WinForms GUI, theme, settings, shell integration.
- `src/XactCopy.Worker` Copy execution worker process.
- `src/XactCopy.Core` Shared models, options, protocol contracts.
- `src/XactCopy.Storage` Journals and persistent state storage.
- `tests/XactCopy.Tests` Unit tests.

## Build

Requires .NET 10 SDK.

```powershell
dotnet build XactCopy.slnx
dotnet test XactCopy.slnx
dotnet run --project src/XactCopy.UI/XactCopy.UI.vbproj
```

## Versioning

- A monotonic build counter is stored in `src/XactCopy.UI/BuildVersion.txt`.
- Each build increments the counter.
- Version is computed as `Major.Minor.Patch.Revision`:
- `Major` increments every 1000 builds (base major starts at `1`).
- `Minor` increments every 100 builds.
- `Patch` increments every 10 builds.
- `Revision` is the last digit of the build counter.

## Release

- Changelog entries: `CHANGELOG.md`
- Release process notes: `docs/RELEASE_PROCESS.md`

## Notes

- Salvage mode can keep a copy moving, but unreadable source bytes are replaced by the selected fill pattern.
- Verification is skipped for salvaged file regions by design.
- No copier can guarantee recovery from every hardware or OS failure mode.
