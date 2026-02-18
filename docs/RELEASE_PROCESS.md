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
   - `dotnet publish src/XactCopy.UI/XactCopy.UI.vbproj -c Release -o artifacts/publish/<version>/win-x64`
2. Update release documentation (required gate):
   - Add a new version section in `CHANGELOG.md`.
   - Add a concise one-line entry for the same version in `README.md` under `Brief Version History`.
   - Keep both files in sync with the release tag.
3. Package release artifact:
   - Zip `artifacts/publish/<version>/win-x64` to `artifacts/releases/XactCopy-v<version>-win-x64.zip`.
   - Publish matching checksum file `XactCopy-v<version>-win-x64.zip.sha256`.
4. Commit changes and push to `main`.
5. Create a git tag from the computed version (for example `v1.0.0.1`).
6. Create a GitHub release using the same tag and include notes aligned with `CHANGELOG.md`.
7. Upload both assets (`.zip` and `.sha256`) to the release.
