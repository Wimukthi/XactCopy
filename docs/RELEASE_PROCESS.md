# Release Process

## Versioning Model

XactCopy uses auto-incremented four-part versions from `src/XactCopy.UI/BuildVersion.txt`:

- `Major = 1 + (build / 1000)`
- `Minor = (build % 1000) / 100`
- `Patch = (build % 100) / 10`
- `Revision = build % 10`

The counter increments on each UI project build.

## Release Steps

1. Build and test:
   - `dotnet build XactCopy.slnx`
   - `dotnet test XactCopy.slnx --no-build`
2. Confirm changelog updates in `CHANGELOG.md`.
3. Commit changes and push to `main`.
4. Create a git tag from the computed version (for example `v1.0.0.1`).
5. Create a GitHub release using the same tag and include changelog notes.
