# AGENTS.md - pub/reader-data-viewer

Entry-point notes. For details see [README.md](README.md).

## Role

- **Reader Data Viewer** — the app this repository exists to build. Not yet implemented.
- Reading detection (Notepad / UI Automation) → join three 1,000,000-row × 10-column CSVs
  → search by key 1 → display
- A-B joins on key 1, B-C joins on key 2, both one-to-one
- Three implementations to be compared, each instrumented per stage:
  VBA/Excel, C#/WinForms, C# + Excel
- `benchmarks/` holds method-selection evidence only. It is not the product.

## runtime / connections

- runtime: local Windows + Excel only
- no network at run time; the benchmark fetches its source data at build time only
- transport (benchmark): UI Automation for control, one notification cell +
  `Worksheet_Change` for completion

## Where to look first

- `README.md` — what the app will do, and what the benchmark already settled
- `benchmarks/excel-background-bench/docs/results.md` — the measurements
- `benchmarks/excel-background-bench/README.md` — how to build and reproduce them
- `benchmarks/excel-background-bench/vba/modZipRule.bas` — the shared conversion rule
- `benchmarks/excel-background-bench/worker/ZipWorker.cs` — the out-of-process worker
- `benchmarks/excel-background-bench/xll/ZbXll.cs` — the in-process (XLL) variant

## Dev commands

```powershell
# benchmarks/excel-background-bench
pwsh -File worker\build_worker.ps1 -Out <repo>\prebuilt\ZipWorker.exe
pwsh -File xll\fetch_exceldna.ps1
pwsh -File xll\build_xll.ps1
pwsh -File build\build_workbooks.ps1 -Root .
pwsh -File build\run_bench.ps1 -Root . -Count 2000 -Methods 1,2,4,5,6,7,9,10,13,14,15,16
```

## Guardrails

- Never touch an Excel instance or process this code did not start.
  No `GetActiveObject`, no running-object-table lookups — bind from a window handle only.
- Do not commit generated artifacts: workbooks, `.exe`, `.xll`, the 1,000,000-row data,
  Japan Post data, or the Excel-DNA download. All are reproducible from the scripts.
- Do not update `Application.StatusBar` or write to sheets while a measurement is running.
  Stamp the end time, finalise the numbers, then display once.
- When a method fails, record the failure. Never fall back to a different method.
- Keep the conversion rule identical across VBA / C# / PowerShell. The PowerShell one
  is the independent oracle; changing only one side silently invalidates every result.
- `.bas` files are imported in the system ANSI code page (CP932 on Japanese Windows).
  Source must survive a CP932 round trip.
