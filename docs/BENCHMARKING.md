# Benchmarking

`XactCopy.Benchmarks` is a repeatable local harness for comparing copy-engine policies, scan-only reads, and small-file worker settings.

## Quick Smoke Run

```powershell
dotnet run --project tools/XactCopy.Benchmarks -- --mode copy --scenario small --policies auto,managed,native --iterations 1 --small-files 100 --small-size-kb 8
```

The runner creates deterministic source data, copies or scans it depending on `--mode`, and writes:

- `benchmark-results.json`
- `benchmark-results.md`

By default reports go under `artifacts/benchmarks/<timestamp>` and generated data is removed after the run.

## Useful Options

```powershell
dotnet run --project tools/XactCopy.Benchmarks -- `
  --scenario small,mixed,large `
  --mode copy,scan `
  --policies auto,managed,native `
  --scan-profile fast `
  --scan-workers 8 `
  --iterations 3 `
  --workers 8 `
  --threshold-kb 256 `
  --small-files 5000 `
  --small-size-kb 8 `
  --large-files 4 `
  --large-size-mb 256 `
  --output artifacts/benchmarks/manual-run `
  --keep
```

- `auto` uses XactCopy's default native/managed selection.
- `managed` forces the managed rescue engine.
- `native` forces native-fast preference while still respecting safety gates.
- `--mode scan` runs scan-only bad-block detection without writing destination files.
- `--scan-profile auto|fast|precise` chooses the scan-only engine.
- `--workers` controls the parallel small-file phase.
- `--scan-workers` controls parallel fast-scan workers.
- `--keep` leaves generated source and destination data in place for inspection.

For meaningful numbers, close other heavy disk users and run on the target storage type. Use at least three iterations when comparing tuning changes.
