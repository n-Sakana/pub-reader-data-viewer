# ============================================================================
# bench_save.ps1 -- the three ways one processed record can reach the ledger,
# scored the way a per-record save has to be scored.
#
#   powershell -File build\bench_save.ps1
#   powershell -File build\bench_save.ps1 -Method zip -Items 10
#
# WHAT IS MEASURED
#   One record, from the save request being accepted to that record being on
#   disk and confirmed. Nothing method-specific is moved outside it: the
#   current method's ledger-workbook OPEN, the ADO provider's first connection
#   and the binary method's walk of the deflate stream all land on the FIRST
#   record, because that is what a normal session pays. The UI's one-second
#   pump is NOT in this figure -- it is reported separately by bench_app.ps1.
#
# THE SCORE IS THE MAXIMUM OF THE TEN, first record included. All ten values
# are printed and written to the TSV. A method that is only fast after an
# expensive warm-up scores as its warm-up.
#
# WHAT IS CHECKED, every record
#   1. the record is re-read straight out of the package, by a reader that is
#      neither the writer nor Excel, before the next record is asked for
#   2. the whole package is diffed against what it was one record ago: the
#      decompressed target sheet must differ in exactly the ONE byte of this
#      record's flag, and no other part may differ at all
#   A method that cannot pass 2 is reported as failing it, with the parts and
#   byte counts that changed -- not quietly excused.
#
# Each method gets a new session on its own fresh copy of the ledger. Excel is
# started by this script and only what it started is ever closed.
# ============================================================================
[CmdletBinding()]
param(
  [ValidateSet('all', 'book', 'book-flagonly', 'ado', 'zip')] [string] $Method = 'all',
  [int] $Items = 10,
  [string] $Root = "",
  [switch] $KeepScratch
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_own.ps1')   # exact Excel ownership, never a pid diff
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type @"
using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;

// Comparing 128 MB of XML byte by byte, twice per record, is the only way to
// prove "one byte and nothing else" -- and it has to be compiled to be usable.
public static class RdvPkg {
  public static Dictionary<string, byte[]> Parts(string path) {
    var d = new Dictionary<string, byte[]>();
    using (var z = ZipFile.OpenRead(path)) {
      foreach (var e in z.Entries) {
        using (var s = e.Open())
        using (var ms = new MemoryStream()) { s.CopyTo(ms); d[e.FullName] = ms.ToArray(); }
      }
    }
    return d;
  }
  public static string BiggestSheet(Dictionary<string, byte[]> parts) {
    string best = null; long len = -1;
    foreach (var kv in parts)
      if (kv.Key.StartsWith("xl/worksheets/") && kv.Value.LongLength > len) { len = kv.Value.LongLength; best = kv.Key; }
    return best;
  }
  // returns "" when the only difference is one byte of the target sheet
  public static string DiffReport(Dictionary<string, byte[]> a, Dictionary<string, byte[]> b, string sheet, out int sheetBytes) {
    string report = ""; sheetBytes = -1;
    foreach (var kv in a) {
      byte[] x = kv.Value, y;
      if (!b.TryGetValue(kv.Key, out y)) { report += "part removed: " + kv.Key + "; "; continue; }
      int n = 0; string first = "";
      int len = Math.Min(x.Length, y.Length);
      for (int i = 0; i < len; i++)
        if (x[i] != y[i]) { n++; if (n <= 2) first += i + ":" + (char)x[i] + "->" + (char)y[i] + " "; }
      if (x.Length != y.Length) { n += Math.Abs(x.Length - y.Length); first += "LEN " + x.Length + "->" + y.Length; }
      if (kv.Key == sheet) { sheetBytes = n; if (n != 1) report += "sheet differs in " + n + " bytes (" + first + "); "; }
      else if (n != 0) report += kv.Key + " differs in " + n + " bytes (" + first + "); ";
    }
    foreach (var kv in b) if (!a.ContainsKey(kv.Key)) report += "part added: " + kv.Key + "; ";
    return report;
  }
  // the raw bytes and every header field of every entry, so "untouched" can be
  // checked at the container level too
  public static string RawReport(string p1, string p2, string sheet) {
    var a = Raw(p1); var b = Raw(p2); string r = "";
    foreach (var kv in a) {
      if (kv.Key == sheet) continue;
      string y;
      if (!b.TryGetValue(kv.Key, out y)) { r += "entry gone: " + kv.Key + "; "; continue; }
      if (kv.Value != y) r += "entry changed: " + kv.Key + "; ";
    }
    return r;
  }
  static Dictionary<string, string> Raw(string path) {
    var d = new Dictionary<string, string>();
    byte[] b = File.ReadAllBytes(path);
    int eocd = -1;
    for (int i = b.Length - 22; i >= 0; i--)
      if (b[i] == 0x50 && b[i+1] == 0x4B && b[i+2] == 0x05 && b[i+3] == 0x06) { eocd = i; break; }
    int cnt = BitConverter.ToUInt16(b, eocd + 10);
    long q = BitConverter.ToUInt32(b, eocd + 16);
    for (int e = 0; e < cnt; e++) {
      int nlen = BitConverter.ToUInt16(b, (int)q + 28), xlen = BitConverter.ToUInt16(b, (int)q + 30), clen = BitConverter.ToUInt16(b, (int)q + 32);
      long lo = BitConverter.ToUInt32(b, (int)q + 42);
      string nm = System.Text.Encoding.ASCII.GetString(b, (int)q + 46, nlen);
      int lnlen = BitConverter.ToUInt16(b, (int)lo + 26), lxlen = BitConverter.ToUInt16(b, (int)lo + 28);
      long data = lo + 30 + lnlen + lxlen;
      long cs = BitConverter.ToUInt32(b, (int)q + 20);
      var h = new System.Security.Cryptography.SHA256Managed();
      byte[] raw = new byte[cs];
      Array.Copy(b, data, raw, 0, cs);
      string sig = BitConverter.ToString(h.ComputeHash(raw)) + "|"
        + BitConverter.ToUInt16(b, (int)q + 8) + "|" + BitConverter.ToUInt16(b, (int)q + 10) + "|"
        + BitConverter.ToUInt16(b, (int)q + 12) + "|" + BitConverter.ToUInt16(b, (int)q + 14) + "|"
        + BitConverter.ToUInt32(b, (int)q + 16) + "|" + cs + "|" + BitConverter.ToUInt32(b, (int)q + 24) + "|"
        + xlen + "|" + clen;
      d[nm] = sig;
      q = q + 46 + nlen + xlen + clen;
    }
    return d;
  }
  // the flag of one row, straight out of the package
  public static string Flag(string path, int sheetRow) {
    using (var z = ZipFile.OpenRead(path)) {
      ZipArchiveEntry sheet = null;
      foreach (var e in z.Entries)
        if (e.FullName.StartsWith("xl/worksheets/") && (sheet == null || e.Length > sheet.Length)) sheet = e;
      string needle = "<c r=\"A" + sheetRow + "\"";
      using (var sr = new StreamReader(sheet.Open())) {
        char[] buf = new char[1 << 20]; string carry = ""; int n;
        while ((n = sr.Read(buf, 0, buf.Length)) > 0) {
          string s = carry + new string(buf, 0, n);
          int i = s.IndexOf(needle);
          if (i >= 0) return s.Substring(i, Math.Min(60, s.Length - i));
          carry = s.Length > 200 ? s.Substring(s.Length - 200) : s;
        }
      }
    }
    return "(not found)";
  }
}
"@ -ReferencedAssemblies System.IO.Compression, System.IO.Compression.FileSystem

$distV = Join-Path $Root 'dist\app-vba'
$srcLedger = Join-Path $distV 'ReaderDataViewer-Ledger.xlsx'
$srcState = Join-Path $distV 'ReaderDataViewer-Ledger.state'
if (-not (Test-Path -LiteralPath $srcLedger)) { throw "not built: run build\build_app.ps1 first" }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$scratch = Join-Path $Root ("work\bench-save-c\{0}" -f $stamp)
New-Item -ItemType Directory -Path $scratch -Force | Out-Null
$tsv = Join-Path $Root ("work\bench-save-contract-{0}.tsv" -f $stamp)
Set-Content -LiteralPath $tsv -Value "method`titem`tsheet_row`tsave_ms`treread`tsheet_diff_bytes`tother_diffs`tnote" -Encoding utf8
$rows = New-Object System.Collections.ArrayList
function Rec([string] $m, [int] $i, [int] $r, [string] $ms, [string] $reread, [string] $sd, [string] $od, [string] $note) {
  Add-Content -LiteralPath $tsv -Value ("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}" -f $m, $i, $r, $ms, $reread, $sd, $od, ($note -replace "`t", ' ')) -Encoding utf8
  [void]$rows.Add([pscustomobject]@{ Method = $m; Item = $i; Row = $r; Ms = $ms; Reread = $reread; SheetDiff = $sd; Other = $od; Note = $note })
}
function Say([string] $s) { Write-Output ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $s) }

# the same ten rows for every method, spread over the whole ledger so the score
# cannot be flattered by picking shallow ones
$targets = @()
for ($i = 0; $i -lt $Items; $i++) { $targets += (2 + [int]([Math]::Floor($i * (99999.0 / [Math]::Max(1, $Items - 1))))) }
$targets = @($targets | Sort-Object -Unique)
Say ("target sheet rows: " + ($targets -join ','))

$cp932 = [Text.Encoding]::GetEncoding(932)
$modules = @('modRdv3Spec', 'modRdv3Chan', 'modRdv3Zip', 'modRdv3Save', 'modRdv3Bench')
$srcDir = Join-Path $Root 'src\app\vba'
$stage = Join-Path $env:TEMP ("rdv3bc_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
foreach ($m in $modules) {
  # a module-level declaration after a procedure does not compile, and an
  # invisible automation Excel shows that as a modal nobody can see
  $seenProc = $false; $ln = 0
  foreach ($line in [IO.File]::ReadAllLines((Join-Path $srcDir "$m.bas"))) {
    $ln++
    if ($line -match "^\s*'") { continue }
    if ($line -match '(?i)^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property)\s') { $seenProc = $true; continue }
    if (-not $seenProc) { continue }
    if ($line -match '(?i)^\s*(Public|Private|Global)\s+(Const|Type|Declare|WithEvents)\s' -or
        $line -match '(?i)^\s*(Public|Private|Global)\s+[A-Za-z_][A-Za-z0-9_]*\s*(\([^)]*\))?\s+As\s') {
      throw "$m line ${ln}: module-level declaration after a procedure: $($line.Trim())"
    }
  }
  $t = [IO.File]::ReadAllText((Join-Path $srcDir "$m.bas"))
  if ($t -ne $cp932.GetString($cp932.GetBytes($t))) { throw "$m does not survive CP932" }
  $t = ($t -replace "`r`n", "`n") -replace "`n", "`r`n"
  [IO.File]::WriteAllText((Join-Path $stage "$m.bas"), $t, $cp932)
}
Say 'preflight: ok'

$stale = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like '*ReaderDataViewer*' })
if ($stale.Count -gt 0) { throw "a Reader Data Viewer workbook is open in Excel. Close it first." }

$methods = if ($Method -eq 'all') { @('book', 'ado', 'zip') } else { @($Method) }
foreach ($mth in $methods) {
  Say ("=== {0}: new session, fresh copy" -f $mth)
  $dir = Join-Path $scratch $mth
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
  $live = Join-Path $dir 'ReaderDataViewer-Ledger.xlsx'
  Copy-Item $srcLedger $live -Force
  Copy-Item $srcState (Join-Path $dir 'ReaderDataViewer-Ledger.state') -Force
  $itemTsv = Join-Path $dir 'items.tsv'

  # identity is settled before anything is done to it (excel_own.ps1):
  # a reused or unidentifiable instance throws here and is never driven
  $rdvOwn = New-OwnedExcel
  $xl = $rdvOwn.App
  $mine = @($rdvOwn.Pid)
  $xl.Visible = $false; $xl.DisplayAlerts = $false
  try {
    $wb = $xl.Workbooks.Add(-4167)
    $wb.SaveAs((Join-Path $stage ($mth + '.xlsm')), 52)
    foreach ($m in $modules) { [void]$wb.VBProject.VBComponents.Import((Join-Path $stage "$m.bas")) }
    $wb.Save()
    $st = $xl.Run("'" + $wb.Name + "'!modRdv3Save.Rdv3SaveSelfTest")
    if ($st -notmatch 'crc\(123456789\)=CBF43926' -or $st -notmatch 'delta_ok=True') { throw "self-test failed: $st" }

    $opened = $xl.Run("'" + $wb.Name + "'!modRdv3Bench.Rdv3BenchOpen", $mth, $live, $itemTsv)
    if ($opened -notlike 'ok*') { Say "  open failed: $opened"; continue }

    # The current method keeps the ledger workbook OPEN, so the file is locked
    # against other readers for the whole session. The checks are done on a
    # copy so that every method is checked the same way; the lock itself is
    # reported rather than worked around silently.
    $snap = Join-Path $dir 'snapshot.xlsx'
    function Snap { Copy-Item $live $snap -Force; return $snap }
    $prev = [RdvPkg]::Parts((Snap))
    $sheet = [RdvPkg]::BiggestSheet($prev)
    for ($i = 0; $i -lt $targets.Count; $i++) {
      $row = $targets[$i]
      $ledgerRow = $row - 2
      $r = $xl.Run("'" + $wb.Name + "'!modRdv3Bench.Rdv3BenchStep", $i, $ledgerRow)
      if ($r -notlike 'ok*') {
        Say ("  item {0} row {1}: FAILED {2}" -f $i, $row, $r)
        Rec $mth $i $row '' 'FAIL' '' '' $r
        break
      }
      $ms = ($r -replace '.*\bms=([0-9.]+).*', '$1')
      # 1. the record itself, read back out of the package right now
      [void](Snap)
      $flag = [RdvPkg]::Flag($snap, $row)
      $ok1 = ($flag -match '<v>1</v>')
      # 2. this record and nothing else
      $now = [RdvPkg]::Parts($snap)
      $sd = 0
      $rep = [RdvPkg]::DiffReport($prev, $now, $sheet, [ref]$sd)
      $prev = $now
      Rec $mth $i $row $ms $(if ($ok1) { 'ok' } else { 'FAIL' }) $sd $(if ($rep -eq '') { 'none' } else { $rep }) $r
      Say ("  item {0,2} row {1,6}  {2,9} ms   reread {3}   sheet_diff {4}   other {5}" -f `
        $i, $row, $ms, $(if ($ok1) { 'ok' } else { 'FAIL' }), $sd, $(if ($rep -eq '') { 'none' } else { $rep.Substring(0, [Math]::Min(70, $rep.Length)) }))
    }
    $closed = $xl.Run("'" + $wb.Name + "'!modRdv3Bench.Rdv3BenchClose")
    Say ("  closed: {0}" -f $closed)
  } finally {
    try { $xl.Quit() } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 800
    foreach ($p in @($mine | Where-Object { $_ })) {
      $q = Get-Process -Id $p -ErrorAction SilentlyContinue
      if ($q) { Say "  closing my excel $p"; $q.Kill() }
    }
  }
  if (-not $KeepScratch) { Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
}
Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue

Write-Output ''
Write-Output ("=== summary -> {0}" -f $tsv)
Write-Output ("{0,-6} {1,10} {2,10} {3,12} {4,10} {5,12} {6}" -f 'method', 'records', 'max ms', 'first ms', 'min ms', 'rereads ok', 'one byte and nothing else')
foreach ($mth in @('book', 'book-flagonly', 'ado', 'zip')) {
  $set = @($rows | Where-Object { $_.Method -eq $mth -and $_.Ms -ne '' })
  if ($set.Count -eq 0) { continue }
  $v = @($set | ForEach-Object { [double]$_.Ms })
  $okReads = @($set | Where-Object { $_.Reread -eq 'ok' }).Count
  $clean = @($set | Where-Object { $_.Other -eq 'none' -and $_.SheetDiff -eq '1' }).Count
  Write-Output ("{0,-6} {1,10} {2,10:N1} {3,12:N1} {4,10:N1} {5,10} {6}" -f `
    $mth, $set.Count, (($v | Measure-Object -Maximum).Maximum), $v[0], (($v | Measure-Object -Minimum).Minimum),
    ("$okReads/$($set.Count)"), ("$clean/$($set.Count)"))
  Write-Output ("       all: " + (($v | ForEach-Object { '{0:N1}' -f $_ }) -join '  '))
}
Write-Output ("scratch: {0}" -f $scratch)
