# Contributing

## Prerequisites

- .NET 10 SDK
- Windows environment for WinForms development

## Local Build

```powershell
dotnet build XactCopy.slnx
dotnet test XactCopy.slnx
```

## Coding Guidelines

- Language: VB.NET
- Keep `Option Strict`, `Option Infer`, and `Option Explicit` enabled.
- Prefer small, focused changes.
- Keep UI behavior consistent between Settings defaults and Main form runtime controls.
- Add or update tests when changing core copy, journal, or protocol behavior.

## Pull Request Checklist

- Build succeeds with no errors.
- Tests pass.
- User-facing behavior changes are documented in `CHANGELOG.md`.
- GPLv3 license headers and terms are preserved.
