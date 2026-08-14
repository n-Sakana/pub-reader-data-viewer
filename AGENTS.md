# AGENTS.md - pub/reader-data-viewer

Entry-point notes. For details see [README.md](README.md).

## Role

- **Reader Data Viewer** — the app this repository exists to build. Implemented.
- Reading detection (Notepad / UI Automation) → join three 1,000,000-row × 10-column CSVs
  → search by key 1 → display, with every stage timed
- A-B joins on key 1, B-C joins on key 2, both one-to-one, over all rows every time
- Three implementations, same workload and same measurement boundaries:
  VBA/Excel, C#/WinForms, C# + Excel
- `benchmarks/` holds method-selection evidence only. It is not the product and is
  read-only: do not change anything under it.

## runtime / connections

- runtime: local Windows + Excel only, no network at run time
- detection: UI Automation polling of a Notepad window (ValuePattern on the text host)
- the app writes nothing except the display surface; workbooks are opened read-only

## Where to look first

- `README.md` — what the app does, how to build and run it, known limits
- `docs/results.md` — the 1:1 measurements and everything that had to be discovered
- `docs/results2.md` — the one-to-many measurements: what an index costs, and the traps
- `docs/app.md` — the practical build: ledger model, FE/BE architecture, measurements

Two frozen comparisons and one practical build live here, with separate sources,
distributables and measurements:

| | sources | build | data | measured |
|---|---|---|---|---|
| 1:1, 1,000,000 rows, 3 methods | `src/csharp`, `src/vba`, `src/cmd` | `build_all.ps1` | `data/` | `docs/results.md` |
| one-to-many, 100,000 rows, 4 methods | `src/v2/**` | `build_all2.ps1` | `data-100k/`, `data-tiny/` | `docs/results2.md` |
| practical, standard index only | `src/app/**` | `build_app.ps1` | copies of `data-100k/` | `docs/app.md` |

The practical build persists the merged ledger (identity = key2, processed flag
carried across rebuilds) and checks for updates at startup by comparing CONTENT
against the saved state. The old per-trigger re-read contract applies to the
frozen comparisons only. The VBA practical build (v2) follows a proven Excel
FE/BE architecture with the ledger SPLIT OUT of the parent book: the FE is a small
UI-only workbook; the ledger workbook (`ReaderDataViewer-Ledger.xlsx`) and its
sidecar mirror (`.state`, UTF-16LE) are read, compared, carried, written and saved
ONLY by an invisible BE Excel, which also owns every heavy stage and the Notepad
watch, publishes atomically to a file channel, and the FE pulls with a short
bounded OnTime pump. The BE never calls into the FE by COM, the FE holds no
COM reference to the BE (it drops them as soon as the BE's bootstrap returns),
and the FE never materializes ledger rows or saves itself (the reasoning, and the
measurement it rests on, is summarised in `docs/app.md`; v1 — ledger inside the
FE book — is archived outside the repository and documented at the end of
`docs/app.md`).

**The practical VBA build (`src/app/vba`) contains no `Declare` and no `Shell`.**
It runs on VBA + COM only: the BE is started with `CreateObject("Excel.Application")`
and a bootstrap that arms `Application.OnTime` and returns; the clock is `Timer`;
sub-second waits come from a WMI event source's `NextEvent` timeout; the Notepad
window is found with UIA `GetFocusedElement` plus a parent walk; liveness is a
lease-file lock in both directions. Keep it that way — the reasons, and the
measurements behind each choice, are in `docs/app.md` ("Win32 / Shell を使わない実装").
The frozen comparison builds keep their Declare blocks and are not touched.

The 1:1 sources are **frozen**: they still have to reproduce the four distributables that
were measured for `docs/results.md`. Work on the one-to-many side goes in `src/v2`, even
when that means a deliberate copy (`Rdv2Watch.cs`, `modRdv2Uia.bas`).

- `src/csharp/RdvCore.cs` — the 1:1 engine (both .cmd builds share it verbatim)
- `src/vba/modRdvEngine.bas` — the same algorithm inside Excel
- `src/v2/csharp/Rdv2Core.cs` — the one-to-many engine; the index is NOT in it
- `src/v2/csharp/Rdv2IdxHash.cs` / `Rdv2IdxDict.cs` — two classes with the same name and
  the same members; the packer puts exactly one in each .cmd, so there is no dispatch
- `src/v2/vba/clsRdv2IdxHash.cls` / `clsRdv2IdxDict.cls` — the same trick in VBA

## Dev commands

```
build.bat            everything in dist\, from source (double-clickable, no admin)
```

```powershell
powershell -ExecutionPolicy Bypass -File build\build_dist.ps1     # what build.bat runs
powershell -ExecutionPolicy Bypass -File build\build_all.ps1
powershell -ExecutionPolicy Bypass -File build\run_bench.ps1 -Method csharp -Repeat 7
powershell -ExecutionPolicy Bypass -File build\run_bench.ps1 -Method hybrid -Repeat 7
powershell -ExecutionPolicy Bypass -File build\run_bench.ps1 -Method vba    -Repeat 7
powershell -ExecutionPolicy Bypass -File build\summarize.ps1 -Log work\bench-vba-<stamp>.tsv

powershell -ExecutionPolicy Bypass -File build\build_all2.ps1
powershell -ExecutionPolicy Bypass -File build\run_bench2.ps1 -Method chash -Repeat 8
powershell -ExecutionPolicy Bypass -File build\run_bench2.ps1 -Method cdict -Repeat 8
powershell -ExecutionPolicy Bypass -File build\run_bench2.ps1 -Method vhash -Repeat 8
powershell -ExecutionPolicy Bypass -File build\run_bench2.ps1 -Method vdict -Repeat 8

powershell -ExecutionPolicy Bypass -File build\build_app.ps1
powershell -ExecutionPolicy Bypass -File build\bench_app.ps1 -Method csharp -Launches 7
powershell -ExecutionPolicy Bypass -File build\bench_app.ps1 -Method vba    -Launches 7
```

## Guardrails

- Never touch an Excel instance or process this code did not start. No `GetActiveObject`,
  no running-object-table lookups — bind from a window handle only.
- Never start, close or kill Notepad. The app attaches to a window that is already there
  and only reads from it.
- Do not commit generated artifacts: `data\` (241 MB), `dist\` (the four distributables),
  `work\` (logs). All are reproducible from `build\`.
- Four methods on the one-to-many side must stay comparable too: same files, same key
  sequence, same measurement boundary. The boundary ends when one candidate is on screen
  or the candidate list is built — never after a person has chosen. The post-choice
  display is a separate figure.
- An index must answer with a SET of rows. `key -> single row` loses rows that legitimately
  share a key 1, and no build here may do that.
- The three methods must stay comparable. Same input files, same join rule, same
  measurement boundaries, same screen. Every run reports a join checksum, a probe count
  and two matched counts; all three methods must agree, and the checksum must match
  `data\expected.txt`, which the generator computed independently.
- Never cache between triggers. Each merge-select re-reads all three CSVs, rebuilds all
  three indexes and redoes both joins over all 1,000,000 rows.
- Do not update the status bar or write timings while a measurement is running. Stamp the
  end time, finalise the numbers, then display once.
- When a method fails, record the failure. Never fall back to a different method.
- The polynomial hash is inlined into three loops in `modRdvEngine.bas`. `build\build_workbooks.ps1`
  compares every copy against `modRdvSpec.RdvHashBytes` and stops the build if they differ.
- `.bas` files are imported in the system ANSI code page (CP932 on Japanese Windows).
  Source must survive a CP932 round trip.
- `.ps1` files that contain non-ASCII need a UTF-8 BOM: Windows PowerShell 5.1 reads a
  BOM-less script in the ANSI code page and the parser breaks on the mangled text.
- C# embedded in the `.cmd` files is compiled by the in-box .NET Framework `csc`:
  **C# 5 only**, and no verbatim strings (the packer escapes non-ASCII to `\uXXXX`).
