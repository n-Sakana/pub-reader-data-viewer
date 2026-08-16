# ============================================================================
# bench_e2e.ps1 -- the three save methods measured on the REAL user path, with
# the real distributable, the real FE/BE processes and the real file channel.
#
#   powershell -File build\bench_e2e.ps1                         (everything)
#   powershell -File build\bench_e2e.ps1 -Method zip -Mode cold -Trials 1
#
# WHAT IS MEASURED. Both units are stamped from OUTSIDE the app, by watching
# the FE<->BE file channel in %TEMP%\rdv3 every 20 ms; nothing in the product
# is changed for the measurement.
#
#   A. the update-check job
#      start  the FE's request to the BE materialises: rdv3_<sid>_fe_lease.lock
#             appears (written at the top of Rdv3HostSpawn). Everything after
#             it is inside: worker book extraction, CreateObject (a second
#             Excel PROCESS), Workbooks.Open of the worker, the bootstrap, the
#             3-CSV merge, the saved-state load, the full content comparison.
#      end    CHECK visible in rdv3_<sid>_agg.dat = the BE returned its result.
#      apply  a second interval for the rebuild: the FE's decision request ->
#             READY, holding the carry, the 100k x 29 write, Workbook.Save and
#             the sidecar write. The operator's answer to the MsgBox sits
#             BETWEEN the two intervals and is in neither.
#
#   B. the processed-registration job
#      start  rdv3_<sid>_req_mark.dat appears (the FE issued the save request,
#             right after the operator answered the confirm dialog)
#      end    res=marked visible in the aggregate = the BE returned the
#             success result to the FE. (The kind is MARK since the
#             confirmation got its own slot; RESULT is still accepted so this
#             harness also runs against the pre-fix build.)
#      Everything the method needs is inside: the book method's 24 MB
#      Workbooks.Open, ADO's provider discovery + connect + ACE write-back +
#      Close, the ZIP method's read + deflate walk + splice + atomic replace,
#      and the sidecar write all three share.
#
#   The FE's one-second pump is NOT in either figure (it is common to all three
#   methods and outside the job); the FE-side e2e is reported separately, as is
#   the FE launch, which every method pays before any job starts.
#
# CONDITIONS
#   cold       a fresh copy of the whole distribution, a new Excel process and
#              a new session per trial; trial i uses target i
#   coldapply  the same, but the saved state no longer matches the CSVs, so the
#              app asks and rebuilds: job A then persists the whole ledger
#   cont       one session, ten consecutive registrations, each persisted and
#              verified before the next
#
# EVERY SAVE IS PROVEN, not believed: the flag is re-read out of the package by
# a reader that is neither the app nor Excel, and the package is diffed against
# what it was immediately before the save -- the target sheet must differ in
# exactly the one flag byte and no other part may differ at all.
#
# Only processes this script started are ever closed, and the close is a COM
# Quit or a WM_CLOSE, never a kill except as a reported last resort.
# ============================================================================
[CmdletBinding()]
param(
  [ValidateSet('vba', 'csharp')] [string] $Build = 'vba',
  [ValidateSet('all', 'book', 'ado', 'zip')] [string] $Method = 'all',
  [ValidateSet('all', 'cold', 'coldapply', 'cont', 'race')] [string] $Mode = 'all',
  [int] $Trials = 10,
  [int] $Rounds = 10,
  [string] $Root = "",
  [int] $ReadyTimeoutSec = 300,
  [int] $SaveTimeoutSec = 300,
  [switch] $KeepScratch
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_own.ps1')   # exact Excel ownership, never a pid diff
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class RdvE2EWin {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr root, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  // A dialog answerer that runs on its own thread: a modal raised BY a
  // blocking COM call (Workbook.Close -> Workbook_BeforeClose -> MsgBox)
  // cannot be answered from the thread that is inside that call.
  public static void AnswerLater(string title, int seconds) {
    System.Threading.Thread t = new System.Threading.Thread(delegate() {
      DateTime end = DateTime.Now.AddSeconds(seconds);
      while (DateTime.Now < end) {
        IntPtr d = FindWindowW("#32770", title);
        if (d != IntPtr.Zero) {
          System.Threading.Thread.Sleep(150);
          foreach (Kid k in Kids(d))
            if (k.Cls == "Button") PostMessage((IntPtr)k.Hwnd, 0x00F5, IntPtr.Zero, IntPtr.Zero);
        }
        System.Threading.Thread.Sleep(200);
      }
    });
    t.IsBackground = true;
    t.Start();
  }
  public class Kid { public long Hwnd; public string Cls; public string Text; }
  public static List<Kid> Kids(IntPtr root) {
    List<Kid> outp = new List<Kid>();
    EnumChildWindows(root, delegate(IntPtr h, IntPtr l) {
      StringBuilder cb = new StringBuilder(256); GetClassNameW(h, cb, 256);
      StringBuilder tb = new StringBuilder(512); GetWindowTextW(h, tb, 512);
      Kid k = new Kid(); k.Hwnd = h.ToInt64(); k.Cls = cb.ToString(); k.Text = tb.ToString();
      outp.Add(k);
      return true;
    }, IntPtr.Zero);
    return outp;
  }
}
"@

# The C# build is driven entirely by posted messages (WM_CHAR to the edit,
# BM_CLICK to the buttons), which do not need the keyboard focus. So the window
# is pushed to the BACK of the z-order as soon as it appears and whatever the
# operator was using keeps the foreground: a test run no longer interrupts them.
function Park-Window([IntPtr] $hwnd) {
  $keep = [RdvE2EWin]::GetForegroundWindow()
  # HWND_BOTTOM = 1, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE = 0x0002|0x0001|0x0010
  [void][RdvE2EWin]::SetWindowPos($hwnd, [IntPtr]1, 0, 0, 0, 0, 0x0013)
  if ($keep -ne [IntPtr]::Zero -and $keep -ne $hwnd) {
    [void][RdvE2EWin]::SetForegroundWindow($keep)
  }
}
[void][RdvE2EWin]::SetProcessDPIAware()

Add-Type @"
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Globalization;
using System.Collections.Generic;

// The channel between the FE and the BE is a handful of small files. Watching
// them from outside puts a timestamp on both ends of a job without changing a
// line of the product: the FE writing a request, and the BE's answer becoming
// visible in the aggregate. One clock (a Stopwatch), one polling loop, so both
// ends of every interval carry the same systematic error (<= one poll).
public class RdvChan {
  string dir;
  Thread th;
  volatile bool run;
  Stopwatch sw = new Stopwatch();
  List<string> ev = new List<string>();
  HashSet<string> files = new HashSet<string>();
  Dictionary<string, long> vers = new Dictionary<string, long>();
  string sid = "";
  Encoding enc = Encoding.GetEncoding(932);
  public int PollMs = 20;
  public string Sid { get { return sid; } }
  public double Now { get { return sw.Elapsed.TotalMilliseconds; } }
  public DateTime Anchor;
  public RdvChan(string d) { dir = d; }

  public void Start() {
    if (Directory.Exists(dir))
      foreach (string f in Directory.GetFiles(dir, "rdv3_*")) files.Add(Path.GetFileName(f));
    Anchor = DateTime.Now;
    sw.Start();
    run = true;
    th = new Thread(new ThreadStart(Loop));
    th.IsBackground = true;
    th.Start();
  }
  public void Stop() { run = false; try { th.Join(500); } catch { } }
  public string[] Events() { lock (ev) { return ev.ToArray(); } }
  void Add(string type, string key, string detail) {
    string t = sw.Elapsed.TotalMilliseconds.ToString("F1", CultureInfo.InvariantCulture);
    lock (ev) { ev.Add(t + "|" + type + "|" + key + "|" + detail); }
  }
  string ReadAll(string p) {
    for (int i = 0; i < 3; i++) {
      try {
        using (FileStream fs = new FileStream(p, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (StreamReader sr = new StreamReader(fs, enc)) return sr.ReadToEnd();
      } catch { }
    }
    return null;
  }
  void Loop() {
    while (run) {
      try {
        if (Directory.Exists(dir)) {
          foreach (string f in Directory.GetFiles(dir, "rdv3_*")) {
            string n = Path.GetFileName(f);
            if (!files.Contains(n)) {
              files.Add(n);
              Add("file", n, "");
              if (n.EndsWith("_fe_lease.lock") && n.StartsWith("rdv3_") && sid.Length == 0)
                sid = n.Substring(5, n.Length - 5 - "_fe_lease.lock".Length);
            }
          }
        }
        if (sid.Length > 0) {
          Agg(Path.Combine(dir, "rdv3_" + sid + "_agg.dat"));
          string[] kinds = new string[] { "decision", "search", "pick", "mark", "watch" };
          for (int i = 0; i < kinds.Length; i++)
            Req(Path.Combine(dir, "rdv3_" + sid + "_req_" + kinds[i] + ".dat"), kinds[i]);
        }
      } catch { }
      Thread.Sleep(PollMs);
    }
  }
  void Agg(string p) {
    if (!File.Exists(p)) return;
    string t = ReadAll(p);
    if (t == null) return;
    foreach (string raw in t.Split('\n')) {
      string s = raw.TrimEnd('\r');
      if (s.Length < 6 || s[0] != 'r') continue;
      string[] parts = s.Split('\t');
      if (parts.Length < 4 || parts[0] != "r") continue;
      long v;
      if (!long.TryParse(parts[2], out v)) continue;
      string key = "agg:" + parts[1];
      long prev;
      if (vers.TryGetValue(key, out prev) && v <= prev) continue;
      vers[key] = v;
      Add("agg", parts[1], "ver=" + v + " " + parts[3]);
    }
  }
  void Req(string p, string kind) {
    if (!File.Exists(p)) return;
    string t = ReadAll(p);
    if (t == null) return;
    string s = t.Split('\n')[0].TrimEnd('\r');
    string[] parts = s.Split('\t');
    if (parts.Length < 3 || parts[0] != "q") return;
    long v;
    if (!long.TryParse(parts[1], out v)) return;
    string key = "req:" + kind;
    long prev;
    if (vers.TryGetValue(key, out prev) && v <= prev) return;
    vers[key] = v;
    Add("req", kind, "ver=" + v + " " + parts[2]);
  }
}

public static class RdvE2E {
  // Ledger rows spread over the whole file whose key 1 is UNIQUE, so a search
  // through the real UI puts exactly one record on screen (which is what the
  // processed button needs). The nearest unique row to each position is taken,
  // and the same series is then used by every method.
  public static string[] PickTargets(string statePath, int[] wanted) {
    List<string> keys = new List<string>(120000);
    using (FileStream fs = File.OpenRead(statePath))
    using (StreamReader sr = new StreamReader(fs, Encoding.Unicode)) {
      sr.ReadLine();
      string line;
      while ((line = sr.ReadLine()) != null) {
        int t1 = line.IndexOf('\t');
        if (t1 < 0) { keys.Add(""); continue; }
        int t2 = line.IndexOf('\t', t1 + 1);
        keys.Add(t2 > t1 ? line.Substring(t1 + 1, t2 - t1 - 1) : "");
      }
    }
    Dictionary<string, int> cnt = new Dictionary<string, int>();
    foreach (string k in keys) { int c; cnt.TryGetValue(k, out c); cnt[k] = c + 1; }
    List<string> res = new List<string>();
    HashSet<int> used = new HashSet<int>();
    foreach (int w in wanted) {
      int best = -1;
      for (int d = 0; d < keys.Count && best < 0; d++) {
        int a = w - d, b = w + d;
        if (a >= 0 && a < keys.Count && cnt[keys[a]] == 1 && !used.Contains(a)) best = a;
        else if (b >= 0 && b < keys.Count && cnt[keys[b]] == 1 && !used.Contains(b)) best = b;
      }
      if (best < 0) throw new Exception("no unique key 1 anywhere near row " + w);
      used.Add(best);
      res.Add(best + "\t" + keys[best] + "\t" + keys.Count);
    }
    return res.ToArray();
  }
  // A saved state that no longer matches the CSVs, made by changing ONE
  // character of the LAST row's content. The app compares CONTENT, so this is
  // the same "difference found" it reports when a CSV changes -- and the
  // rebuild job it triggers is the same work either way.
  public static string PatchState(string src, string dst) {
    byte[] b = File.ReadAllBytes(src);
    int i = b.Length - 2;
    while (i >= 0) {
      char c = (char)(b[i] | (b[i + 1] << 8));
      if (c != '\r' && c != '\n' && c != '\t') break;
      i -= 2;
    }
    if (i < 0) throw new Exception("the saved state has no content");
    char old = (char)(b[i] | (b[i + 1] << 8));
    char neu = (old == 'X') ? 'Y' : 'X';
    b[i] = (byte)neu; b[i + 1] = 0;
    File.WriteAllBytes(dst, b);
    return "byte " + i + ": '" + old + "' -> '" + neu + "'";
  }
}
"@

Add-Type @"
using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;

// 128 MB of sheet XML, compared byte by byte before and after every single
// save. Compiled, because that is the only way it is usable.
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
  // Every row whose FIRST cell is TRUE, with the two key columns next to it.
  // The C# build writes inline strings and no r= attributes, so its ledger has
  // to be re-read this way -- and this proves WHICH record carries the flag,
  // not just how many do.
  public static string[] TrueRows(string path) {
    var res = new List<string>();
    using (var z = ZipFile.OpenRead(path)) {
      ZipArchiveEntry sheet = null;
      foreach (var e in z.Entries)
        if (e.FullName.StartsWith("xl/worksheets/") && (sheet == null || e.Length > sheet.Length)) sheet = e;
      using (var sr = new StreamReader(sheet.Open())) {
        char[] buf = new char[1 << 20];
        string carry = "";
        int n, rowNo = 0;
        while ((n = sr.Read(buf, 0, buf.Length)) > 0) {
          string s = carry + new string(buf, 0, n);
          int idx = 0;
          while (true) {
            int a = s.IndexOf("<row", idx);
            if (a < 0) { idx = Math.Max(idx, s.Length - 8); break; }
            int b = s.IndexOf("</row>", a);
            if (b < 0) { idx = a; break; }
            rowNo++;
            string row = s.Substring(a, b - a);
            var vals = new List<string>();
            int p = 0;
            while (vals.Count < 3) {
              int t1 = row.IndexOf("<t>", p);
              if (t1 < 0) break;
              int t2 = row.IndexOf("</t>", t1);
              if (t2 < 0) break;
              vals.Add(row.Substring(t1 + 3, t2 - t1 - 3));
              p = t2 + 4;
            }
            if (vals.Count > 0 && (vals[0] == "TRUE" || vals[0] == "1"))
              res.Add(rowNo + ":" + string.Join(":", vals.ToArray()));
            idx = b + 6;
          }
          carry = idx < s.Length ? s.Substring(idx) : "";
        }
      }
    }
    return res.ToArray();
  }
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

# ---------------------------------------------------------------------------
# bookkeeping
# ---------------------------------------------------------------------------
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$scratch = Join-Path $Root ("work\bench-e2e-run\{0}" -f $stamp)
New-Item -ItemType Directory -Path $scratch -Force | Out-Null
$tsv = Join-Path $Root ("work\bench-e2e-{0}.tsv" -f $stamp)
$cols = @('method', 'variant', 'trial', 'ledger_row', 'sheet_row', 'key1', 'xl_start_ms', 'fe_launch_ms',
          'a_check_ms', 'a_ready_ms', 'apply_ms', 'spawn_ms', 'fe_boot_ready_ms', 'b_ms', 'b_fe_e2e_ms',
          'b_save_ms', 'search_ms', 'reread', 'sheet_diff', 'other_diffs', 'raw_entries', 'note')
Set-Content -LiteralPath $tsv -Value ($cols -join "`t") -Encoding utf8
$rows = New-Object System.Collections.ArrayList
function Say([string] $s) { Write-Output ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $s) }
function Rec([hashtable] $h) {
  $line = ($cols | ForEach-Object { $v = $h[$_]; if ($null -eq $v) { '' } else { ("$v" -replace "`t", ' ') } }) -join "`t"
  Add-Content -LiteralPath $tsv -Value $line -Encoding utf8
  [void]$rows.Add([pscustomobject]$h)
}

$distV = Join-Path $Root 'dist\app-vba'
$srcLedger = Join-Path $distV 'ReaderDataViewer-Ledger.xlsx'
$srcState = Join-Path $distV 'ReaderDataViewer-Ledger.state'
if (-not (Test-Path -LiteralPath $srcLedger)) { throw "not built: run build\build_app.ps1 first" }
$chanDir = Join-Path $env:TEMP 'rdv3'
if (-not (Test-Path -LiteralPath $chanDir)) { New-Item -ItemType Directory -Path $chanDir -Force | Out-Null }

$others = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($others.Count -gt 0) {
  Say ("NOTE: {0} Excel process(es) were already running and are not touched: {1}" -f $others.Count, (($others | ForEach-Object { $_.Id }) -join ','))
}
if (@(Get-Process notepad -ErrorAction SilentlyContinue).Count -gt 0) {
  Say 'NOTE: Notepad is open (never started or closed here); the BE only binds while its text field has focus'
}

# ---------------------------------------------------------------------------
# the fixed target series, and the stale saved state for the rebuild variant
# ---------------------------------------------------------------------------
$srcDir = Join-Path $Root 'work\e2e-src'
New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
$targetsFile = Join-Path $srcDir 'targets.tsv'
$staleState = Join-Path $srcDir 'stale.state'
if (-not (Test-Path -LiteralPath $targetsFile)) {
  $wanted = @()
  for ($i = 0; $i -lt 10; $i++) { $wanted += [int]([Math]::Floor($i * (99999.0 / 9.0))) }
  $wanted += 30000
  $wanted += 70000
  Say ("picking targets near ledger rows " + ($wanted -join ','))
  $picked = [RdvE2E]::PickTargets($srcState, [int[]]$wanted)
  Set-Content -LiteralPath $targetsFile -Value $picked -Encoding utf8
}
$targets = @()
foreach ($line in (Get-Content -LiteralPath $targetsFile)) {
  if ($line.Trim().Length -eq 0) { continue }
  $p = $line.Split("`t")
  $targets += [pscustomobject]@{ Row = [int]$p[0]; Key1 = $p[1]; SheetRow = ([int]$p[0]) + 2 }
}
Say ("targets (sheet row/key1): " + (($targets | ForEach-Object { "{0}/{1}" -f $_.SheetRow, $_.Key1 }) -join '  '))
if (-not (Test-Path -LiteralPath $staleState)) {
  Say ("stale saved state built: " + [RdvE2E]::PatchState($srcState, $staleState))
}

# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------
function Answer-Dialog([string] $title, [string] $buttonPrefix, [int] $sec) {
  $cmdId = 1
  if ($buttonPrefix -like 'はい*') { $cmdId = 6 } elseif ($buttonPrefix -like 'いいえ*') { $cmdId = 7 }
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $dlg = [RdvE2EWin]::FindWindowW('#32770', $title)
    if ($dlg -ne [IntPtr]::Zero) {
      Start-Sleep -Milliseconds 150
      $btn = [IntPtr]::Zero
      foreach ($k in [RdvE2EWin]::Kids($dlg)) {
        if ($k.Cls -eq 'Button' -and ($k.Text.Replace('&', '') -like ($buttonPrefix + '*'))) { $btn = [IntPtr]$k.Hwnd; break }
      }
      if ($btn -ne [IntPtr]::Zero) { [void][RdvE2EWin]::PostMessage($btn, 0x00F5, [IntPtr]0, [IntPtr]0) }
      else { [void][RdvE2EWin]::PostMessage($dlg, 0x0111, [IntPtr]$cmdId, [IntPtr]0) }
      return $true
    }
    Start-Sleep -Milliseconds 60
  }
  return $false
}
function EvMs([string] $e) { return [double]$e.Substring(0, $e.IndexOf('|')) }
# the first channel event at or after $after that matches, or $null. Every
# event is examined once, so a long wait costs nothing.
function Wait-Ev($w, [string] $pattern, [double] $after, [int] $sec) {
  $t0 = Get-Date
  $seen = 0
  while ($true) {
    $all = $w.Events()
    for ($i = $seen; $i -lt $all.Count; $i++) {
      if ($all[$i] -match $pattern -and (EvMs $all[$i]) -ge $after) { return $all[$i] }
    }
    $seen = $all.Count
    if (((Get-Date) - $t0).TotalSeconds -ge $sec) { return $null }
    Start-Sleep -Milliseconds 25
  }
}
function Read-Log([string] $p) {
  if (-not (Test-Path -LiteralPath $p)) { return @() }
  $enc = [Text.Encoding]::GetEncoding(932)
  for ($i = 0; $i -lt 20; $i++) { try { return @([IO.File]::ReadAllLines($p, $enc)) } catch { Start-Sleep -Milliseconds 80 } }
  return @()
}
# wait for a new FE log line (the FE only renders on its own pump, so this is
# how the harness knows the screen has caught up)
function Wait-Log([string] $p, [string] $pattern, [int] $from, [int] $sec) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $all = Read-Log $p
    for ($i = $from; $i -lt $all.Count; $i++) { if ($all[$i] -match $pattern) { return $all[$i] } }
    Start-Sleep -Milliseconds 120
  }
  return $null
}

# ---------------------------------------------------------------------------
# a real session on a fresh copy of the distribution
# ---------------------------------------------------------------------------
function New-Session([string] $method, [string] $variant, [string] $tag) {
  $dir = Join-Path $scratch $tag
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
  Copy-Item -Path (Join-Path $distV '*') -Destination $dir -Recurse -Force
  if ($variant -eq 'coldapply') { Copy-Item $staleState (Join-Path $dir 'ReaderDataViewer-Ledger.state') -Force }
  if ($method -ne 'book') {
    Set-Content -LiteralPath (Join-Path $dir 'ReaderDataViewer-Ledger.xlsx.savemethod') -Value $method -Encoding ascii
  }
  $log = Join-Path $dir 'ReaderDataViewer.log'
  if (Test-Path -LiteralPath $log) { Remove-Item -LiteralPath $log -Force }

  $w = New-Object -TypeName RdvChan -ArgumentList $chanDir
  $w.Start()
  $t0 = $w.Now
  # identity is settled before anything is done to it (excel_own.ps1): the pid
  # used to be taken only after the workbook was open, so a reused instance had
  # already been driven by the time anyone asked whose it was
  $rdvOwn = New-OwnedExcel
  $xl = $rdvOwn.App
  $xlMs = $w.Now - $t0
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  for ($i = 0; $i -lt 10; $i++) { try { $xl.AutomationSecurity = 1; break } catch { Start-Sleep -Milliseconds 500 } }
  $t1 = $w.Now
  $wb = $xl.Workbooks.Open((Join-Path $dir 'ReaderDataViewer.xlsm'))
  $launchMs = $w.Now - $t1
  $hwnd = [IntPtr]$xl.Hwnd
  try { $xl.WindowState = -4137 } catch { }
  $ws = $wb.Worksheets.Item('UI')
  $ws.Activate()
  [void][RdvE2EWin]::SetForegroundWindow($hwnd)
  Start-Sleep -Milliseconds 500
  $win = $xl.ActiveWindow
  $dpi = [RdvE2EWin]::GetDpiForWindow($hwnd)
  if ($dpi -eq 0) { $dpi = 96 }
  return [pscustomobject]@{
    Dir = $dir; Log = $log; Ledger = (Join-Path $dir 'ReaderDataViewer-Ledger.xlsx')
    Snap = (Join-Path $dir 'snap.xlsx'); Before1 = (Join-Path $dir 'before.xlsx')
    W = $w; Xl = $xl; Wb = $wb; Ws = $ws; Hwnd = $hwnd; XlMs = $xlMs; LaunchMs = $launchMs
    X0 = $win.PointsToScreenPixelsX(0); Y0 = $win.PointsToScreenPixelsY(0); Scale = ($dpi / 72.0)
    Mine = [System.Collections.ArrayList]@($rdvOwn.Pid)
    BeSeen = [System.Collections.ArrayList]@()
  }
}

# the key goes into the input the way a person's typing does: the top-left cell
# of the merged block the builder named rdvInput, which is where modRdv3Ui reads
# it back from (a merged Range answers .Value2 with an array, not a string).
function Set-Key($s, [string] $key) {
  $s.Ws.Range('rdvInput').Cells(1, 1).Value2 = $key
}

# WHERE THE POINTER ACTUALLY IS before pressing the button.
#
# Range.Left/Top are measured from A1; PointsToScreenPixelsX(0) is the left edge
# of the VISIBLE area, so the two only agree while the sheet is scrolled to A1 --
# and the screen was redrawn onto a pseudo-pixel grid, which moved every control
# far from the origin. A click that misses now hits an empty grid cell and the
# run simply proves nothing, which is exactly how a whole E2E passed while
# testing nothing. Ask Excel what is under the point and refuse to click if it
# is not the control that was asked for.
function Point-Of($s, [string] $addr) {
  $r = $s.Ws.Range($addr)
  return [pscustomobject]@{
    R = $r
    X = [int]($s.X0 + ($r.Left + $r.Width / 2) * $s.Scale)
    Y = [int]($s.Y0 + ($r.Top + $r.Height / 2) * $s.Scale)
  }
}

function Assert-OnTarget($s, [string] $addr, [int] $x, [int] $y) {
  $hit = $null
  try { $hit = $s.Xl.ActiveWindow.RangeFromPoint($x, $y) } catch { }
  if ($null -eq $hit) {
    throw ("click target '{0}' is not on screen: ({1},{2}) is over nothing" -f $addr, $x, $y)
  }
  $hitAddr = ''
  try { $hitAddr = $hit.Address(0, 0) } catch { $hitAddr = '<not a range>' }
  # by row/column bounds, not Application.Intersect: Intersect answers Nothing
  # through COM in a way PowerShell cannot tell from an error, and it called a
  # cell that WAS inside the button a miss
  $t = $s.Ws.Range($addr)
  $r1 = [int]$t.Row; $c1 = [int]$t.Column
  $r2 = $r1 + [int]$t.Rows.Count - 1; $c2 = $c1 + [int]$t.Columns.Count - 1
  $inside = $false
  try {
    $inside = ([int]$hit.Row -ge $r1 -and [int]$hit.Row -le $r2 -and
               [int]$hit.Column -ge $c1 -and [int]$hit.Column -le $c2)
  } catch { $inside = $false }
  if (-not $inside) {
    throw ("click target '{0}' ({1}) is not under ({2},{3}) -- {4} is" -f
           $addr, $s.Ws.Range($addr).Address(0, 0), $x, $y, $hitAddr)
  }
}

function Click-Cell($s, [string] $addr) {
  $p = Point-Of $s $addr
  $x = $p.X; $y = $p.Y
  Assert-OnTarget $s $addr $x $y
  [void][RdvE2EWin]::SetCursorPos($x, $y)
  Start-Sleep -Milliseconds 120
  [RdvE2EWin]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero)
  Start-Sleep -Milliseconds 50
  [RdvE2EWin]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)
}

# WHICH processes this session owns, still running.
#
# The FE is the instance this script created and it was asked for its own pid.
# The BE is a SECOND Excel that the FE starts for itself, so this script never
# saw it created -- but the FE writes its window into its own log
# ("spawn ok be=hwnd:<n>"), and that window names its process the same way. A
# pid this cannot identify is left alone; it is not ours to end.
# A WINDOW HANDLE ONLY NAMES A PROCESS WHILE THE PROCESS IS ALIVE. This used to
# be called only from Live-Owned, which runs at teardown -- by then the BE has
# quit itself, FromHandle fails on a dead handle every time, and the run ends
# with no record of which process the BE was. It is called at READY instead,
# while the BE is demonstrably up, and what it resolves is kept.
function Add-BePid($s) {
  if (-not (Test-Path -LiteralPath $s.Log)) { return }
  $txt = [IO.File]::ReadAllText($s.Log, [Text.Encoding]::GetEncoding(932))
  foreach ($m in [regex]::Matches($txt, 'be=hwnd:(\d+)')) {
    $h = [long]$m.Groups[1].Value
    if ($s.BeSeen -contains $h) { continue }        # resolved once already
    $bp = Get-PidFromHwnd $h
    if ($bp -gt 0) {
      [void]$s.BeSeen.Add($h)
      if ($s.Mine -notcontains $bp) { [void]$s.Mine.Add($bp) }
    }
  }
}

# what this session owns and can say so about, for the record
function Owned-Note($s) {
  return ('owned excel: ' + (@($s.Mine) -join ',') +
          $(if (@($s.BeSeen).Count -eq 0) { ' (BE not resolved)' } else { '' }))
}

function Live-Owned($s) {
  # a late attempt costs nothing and is silent when the handle is already gone
  Add-BePid $s | Out-Null
  return @($s.Mine | Where-Object {
    $q = Get-Process -Id $_ -ErrorAction SilentlyContinue
    $null -ne $q -and $q.ProcessName -eq 'EXCEL'
  })
}

function Close-Session($s) {
  # every COM reference this script holds has to go first: Excel does not quit
  # while an outside client still holds one (measured: the FE process outlived
  # Quit by more than 20 s and had to be killed, which is exactly what this
  # harness must not do to a workbook path)
  try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($s.Ws) } catch { }
  $s.Ws = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  # If the FE is still holding an undecided mark it will REFUSE this close and
  # say so in a modal -- which would park this thread inside the COM call, so
  # the answerer has to be armed before it, on another thread.
  [RdvE2EWin]::AnswerLater('保存中のため終了できません', 260)
  try { if ($null -ne $s.Wb) { $s.Wb.Close($false) } } catch { }
  # the FE may defer its own close by one pump tick; let it finish
  for ($i = 0; $i -lt 20; $i++) {
    $cnt = -1
    try { $cnt = $s.Xl.Workbooks.Count } catch { $cnt = -1 }
    if ($cnt -le 0) { break }
    Start-Sleep -Milliseconds 500
  }
  # still open = the exit guard is holding it. Wait for the app's own decision
  # (its ceiling is 180 s) and then close again; never force it.
  $cnt = -1
  try { $cnt = $s.Xl.Workbooks.Count } catch { $cnt = -1 }
  if ($cnt -gt 0) {
    Say '  the exit guard is holding the close; waiting for the app to decide the save'
    $rel = Wait-Log $s.Log '(exit released|stage=processed after_s=)' 0 240
    Say ("  " + $(if ($null -eq $rel) { 'no decision line appeared' } else { $rel }))
    try { $s.Wb.Close($false) } catch { }
    for ($i = 0; $i -lt 20; $i++) {
      try { $cnt = $s.Xl.Workbooks.Count } catch { $cnt = -1 }
      if ($cnt -le 0) { break }
      Start-Sleep -Milliseconds 500
    }
  }
  try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($s.Wb) } catch { }
  if ($null -ne $s.Xl) {
    # Quit is ignored while a macro is still running (the FE finishes its own
    # close from a pump tick), and it does not fail when it is ignored -- so
    # ask more than once, and only stop when the process is actually gone
    for ($i = 0; $i -lt 8; $i++) {
      try { $s.Xl.Quit() } catch { }
      Start-Sleep -Milliseconds 1500
      $live = @(Live-Owned $s)
      if ($live.Count -eq 0) { break }
    }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($s.Xl) } catch { }
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  # the BE quits itself when the FE drops its lease; give it the same grace the
  # app gives it, then close by window, and only then report a last-resort kill.
  # After a REBUILD the BE still holds the 24 MB ledger workbook open and its
  # own close takes longer than the app's 12 s stop-wait (measured: one BE
  # needed more than 21 s), so the grace here is 42 s.
  for ($i = 0; $i -lt 60; $i++) {
    $live = @(Live-Owned $s)
    if ($live.Count -eq 0) { break }
    Start-Sleep -Milliseconds 700
  }
  $live = @(Live-Owned $s)
  foreach ($p in $live) {
    $q = Get-Process -Id $p -ErrorAction SilentlyContinue
    if ($null -eq $q) { continue }
    if ($q.MainWindowHandle -ne [IntPtr]::Zero) {
      Say ("  closing my excel {0} by window" -f $p)
      [void][RdvE2EWin]::PostMessage($q.MainWindowHandle, 0x0010, [IntPtr]0, [IntPtr]0)
      Start-Sleep -Seconds 3
    }
    $q = Get-Process -Id $p -ErrorAction SilentlyContinue
    if ($null -ne $q) { Say ("  LAST RESORT: killing my excel {0}" -f $p); $q.Kill(); Start-Sleep -Milliseconds 500 }
  }
  $s.W.Stop()
  Set-Content -LiteralPath (Join-Path $s.Dir 'channel.txt') -Value $s.W.Events() -Encoding utf8
}

# search for one key through the real screen, and wait until the FE has
# actually rendered the record (the processed button needs it on screen)
function Do-Search($s, [string] $key, [int] $sec) {
  Set-Key $s $key
  Start-Sleep -Milliseconds 150
  $from = (Read-Log $s.Log).Count
  $t0 = $s.W.Now
  Click-Cell $s 'rdvBtnSearch'
  $req = Wait-Ev $s.W ('\|req\|search\|') $t0 30
  $res = $null
  if ($null -ne $req) { $res = Wait-Ev $s.W ('\|agg\|RESULT\|.*res=single') (EvMs $req) $sec }
  if ($null -eq $res) {
    # Write-Host, not Say: Say writes to the OUTPUT stream, and anything a
    # function writes there becomes part of its return value. The retry note
    # made a failed search come back non-null, so the caller believed it had a
    # record on screen and blamed the missing confirm dialog instead.
    Write-Host '    the search did not answer; clicking once more'
    $from = (Read-Log $s.Log).Count
    $t0 = $s.W.Now
    Click-Cell $s 'rdvBtnSearch'
    $req = Wait-Ev $s.W ('\|req\|search\|') $t0 30
    if ($null -ne $req) { $res = Wait-Ev $s.W ('\|agg\|RESULT\|.*res=single') (EvMs $req) $sec }
    if ($null -eq $res) { return $null }
  }
  $ln = Wait-Log $s.Log ("`tsearch`tkey=" + $key + " ") $from 60
  if ($null -eq $ln) { return $null }
  return [pscustomobject]@{ ReqMs = (EvMs $req); Ms = ((EvMs $res) - (EvMs $req)); Line = $ln }
}

# [処理済み] -> はい, and the job B interval around it
function Do-Mark($s, [int] $sec) {
  $from = (Read-Log $s.Log).Count
  $t0 = $s.W.Now
  Click-Cell $s 'rdvBtnProcessed'
  if (-not (Answer-Dialog '処理済みの確認' 'はい' 30)) {
    return [pscustomobject]@{ Ok = $false; Note = 'the confirm dialog never appeared' }
  }
  $req = Wait-Ev $s.W ('\|req\|mark\|') $t0 30
  if ($null -eq $req) { return [pscustomobject]@{ Ok = $false; Note = 'the FE never issued the save request' } }
  $res = Wait-Ev $s.W ('\|agg\|(RESULT|MARK)\|.*res=marked') (EvMs $req) $sec
  if ($null -eq $res) {
    $err = Wait-Ev $s.W ('\|agg\|(RESULT|MARK)\|.*res=markerr') (EvMs $req) 1
    return [pscustomobject]@{ Ok = $false; Note = $(if ($null -ne $err) { "markerr $err" } else { 'no result within the timeout' }) }
  }
  # and wait for the FE to say it is decided, so the next round starts from a
  # released exit guard exactly as a person's next click would
  $ln = Wait-Log $s.Log 'processed\tkey2=' $from 60
  return [pscustomobject]@{
    Ok = $true; ReqMs = (EvMs $req); Ms = ((EvMs $res) - (EvMs $req))
    Line = $(if ($null -eq $ln) { '' } else { $ln }); FeSaw = ($null -ne $ln)
  }
}

# the save, proven: re-read + the package diff against what it was before it
function Verify-Mark($s, $prevParts, [string] $sheetName, [int] $sheetRow) {
  Copy-Item $s.Ledger $s.Snap -Force
  $flag = [RdvPkg]::Flag($s.Snap, $sheetRow)
  $ok = ($flag -match '<v>1</v>') -or ($flag -match '<is><t>TRUE')
  $now = [RdvPkg]::Parts($s.Snap)
  $sd = 0
  $rep = [RdvPkg]::DiffReport($prevParts, $now, $sheetName, [ref]$sd)
  $raw = [RdvPkg]::RawReport($s.Before1, $s.Snap, $sheetName)
  return [pscustomobject]@{
    Reread = $(if ($ok) { 'ok' } else { 'FAIL' }); Flag = $flag; SheetDiff = $sd
    Other = $(if ($rep -eq '') { 'none' } else { $rep }); Raw = $(if ($raw -eq '') { 'none' } else { $raw })
  }
}

function Snapshot-Before($s) {
  Copy-Item $s.Ledger $s.Before1 -Force
  return [RdvPkg]::Parts($s.Before1)
}

function Fill-SaveFigures([hashtable] $h, [string] $line) {
  if ($line -match 'save_ms=([0-9.]+)') { $h.b_save_ms = $Matches[1] }
  if ($line -match 'e2e_ms=([0-9.]+)') { $h.b_fe_e2e_ms = $Matches[1] }
}
# what the app itself reports about its own startup, for reference next to the
# figures measured from outside: the FE occupancy of starting the BE, and the
# FE's own boot -> operable (a superset of job A: it adds the FE's paint and
# the pump tick that picks READY up)
function Fill-BootFigures([hashtable] $h, [string] $log) {
  $all = Read-Log $log
  foreach ($ln in $all) {
    if ($ln -match 'spawn_ms=([0-9.]+)') { $h.spawn_ms = $Matches[1] }
    if ($ln -match 'boot_to_ready_ms=([0-9.]+)') { $h.fe_boot_ready_ms = $Matches[1] }
  }
}

# ---------------------------------------------------------------------------
# one cold trial: a new session, job A, then the first registration of it
# ---------------------------------------------------------------------------
function Run-Cold([string] $method, [string] $variant, [int] $i) {
  $t = $targets[$i % 10]
  $tag = "{0}-{1}-{2:d2}" -f $method, $variant, $i
  Say ("=== {0}  target sheet row {1} (key1 {2})" -f $tag, $t.SheetRow, $t.Key1)
  $s = $null
  $h = @{ method = $method; variant = $variant; trial = $i; ledger_row = $t.Row; sheet_row = $t.SheetRow; key1 = $t.Key1 }
  $done = $false
  try {
    $s = New-Session $method $variant $tag
    $h.xl_start_ms = '{0:F1}' -f $s.XlMs
    $h.fe_launch_ms = '{0:F1}' -f $s.LaunchMs
    $lease = Wait-Ev $s.W ('\|file\|rdv3_.*_fe_lease\.lock\|') 0 90
    if ($null -eq $lease) { $h.note = 'the FE never asked the BE for anything'; return }
    $t0 = EvMs $lease
    $chk = Wait-Ev $s.W ('\|agg\|CHECK\|') $t0 $ReadyTimeoutSec
    if ($null -eq $chk) { $h.note = 'no CHECK'; return }
    $h.a_check_ms = '{0:F1}' -f ((EvMs $chk) - $t0)
    if ($variant -eq 'coldapply') {
      if ($chk -notmatch 'res=diff') { $h.note = "the app did not see a difference: $chk"; return }
      if (-not (Answer-Dialog '更新の確認' 'はい' 90)) { $h.note = 'the update dialog never appeared'; return }
      $dec = Wait-Ev $s.W ('\|req\|decision\|.*apply') $t0 60
      if ($null -eq $dec) { $h.note = 'no decision request'; return }
      $rdy = Wait-Ev $s.W ('\|agg\|READY\|') (EvMs $dec) $ReadyTimeoutSec
      if ($null -eq $rdy) { $h.note = 'no READY after the rebuild'; return }
      # the BE is up now (it just answered READY); name its process while a
      # window handle still resolves to one, and say what this session owns
      Add-BePid $s
      Say ('    ' + (Owned-Note $s))
      $h.apply_ms = '{0:F1}' -f ((EvMs $rdy) - (EvMs $dec))
      $h.a_ready_ms = '{0:F1}' -f ((EvMs $rdy) - $t0)
    }
    else {
      if ($chk -notmatch 'res=same') { $h.note = "unexpected check result: $chk" }
      $rdy = Wait-Ev $s.W ('\|agg\|READY\|') $t0 $ReadyTimeoutSec
      if ($null -eq $rdy) { $h.note = 'no READY'; return }
      # the BE is up now (it just answered READY); name its process while a
      # window handle still resolves to one, and say what this session owns
      Add-BePid $s
      Say ('    ' + (Owned-Note $s))
      $h.a_ready_ms = '{0:F1}' -f ((EvMs $rdy) - $t0)
    }
    Start-Sleep -Milliseconds 700
    Fill-BootFigures $h $s.Log
    $srch = Do-Search $s $t.Key1 90
    if ($null -eq $srch) { $h.note = 'the search never put a record on screen'; return }
    $h.search_ms = '{0:F1}' -f $srch.Ms

    $prev = Snapshot-Before $s
    $sheet = [RdvPkg]::BiggestSheet($prev)
    $mk = Do-Mark $s $SaveTimeoutSec
    if (-not $mk.Ok) { $h.note = $mk.Note; return }
    $h.b_ms = '{0:F1}' -f $mk.Ms
    Fill-SaveFigures $h $mk.Line
    $v = Verify-Mark $s $prev $sheet $t.SheetRow
    $h.reread = $v.Reread; $h.sheet_diff = $v.SheetDiff; $h.other_diffs = $v.Other; $h.raw_entries = $v.Raw
    $prev = $null
    if ($null -eq $h.note) { $h.note = $(if ($mk.FeSaw) { 'ok' } else { 'ok (the BE answered, but the FE never rendered the confirmation)' }) }
    Say ("    A check {0} ms   ready {1} ms{2}   B {3} ms   reread {4}  sheet_diff {5}  other {6}" -f `
      $h.a_check_ms, $h.a_ready_ms, $(if ($h.apply_ms) { "  (apply $($h.apply_ms) ms)" } else { '' }),
      $h.b_ms, $h.reread, $h.sheet_diff, $(if ("$($h.other_diffs)".Length -gt 60) { "$($h.other_diffs)".Substring(0, 60) } else { $h.other_diffs }))
  }
  catch { $h.note = "harness error: " + $_.Exception.Message + " @" + $_.InvocationInfo.ScriptLineNumber }
  finally {
    if ($null -ne $s) {
      Copy-Item $s.Log (Join-Path $scratch ("{0}.log" -f $tag)) -Force -ErrorAction SilentlyContinue
      Close-Session $s
      Copy-Item (Join-Path $s.Dir 'channel.txt') (Join-Path $scratch ("{0}.channel.txt" -f $tag)) -Force -ErrorAction SilentlyContinue
      if (-not $KeepScratch) { Remove-Item $s.Dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    if ("$($h.note)" -ne 'ok' -and "$($h.note)" -ne '') { Say ("    NOTE: " + $h.note) }
    Rec $h
    $done = $true
  }
}

# ---------------------------------------------------------------------------
# one continuous session: ten registrations, one after the other
# ---------------------------------------------------------------------------
function Run-Cont([string] $method) {
  $tag = "{0}-cont" -f $method
  Say ("=== {0}: one session, {1} consecutive registrations" -f $tag, $Rounds)
  $s = $null
  try {
    $s = New-Session $method 'cont' $tag
    $lease = Wait-Ev $s.W ('\|file\|rdv3_.*_fe_lease\.lock\|') 0 90
    if ($null -eq $lease) { Say '    the FE never asked the BE for anything'; return }
    $t0 = EvMs $lease
    $rdy = Wait-Ev $s.W ('\|agg\|READY\|') $t0 $ReadyTimeoutSec
    if ($null -eq $rdy) { Say '    never reached READY'; return }
    # the BE is up now (it just answered READY); name its process while a
    # window handle still resolves to one, and say what this session owns
    Add-BePid $s
    Say ('    ' + (Owned-Note $s))
    Start-Sleep -Milliseconds 700
    for ($i = 0; $i -lt $Rounds; $i++) {
      $t = $targets[$i % 10]
      $h = @{ method = $method; variant = 'cont'; trial = $i; ledger_row = $t.Row; sheet_row = $t.SheetRow; key1 = $t.Key1 }
      if ($i -eq 0) {
        $h.xl_start_ms = '{0:F1}' -f $s.XlMs
        $h.fe_launch_ms = '{0:F1}' -f $s.LaunchMs
        $h.a_ready_ms = '{0:F1}' -f ((EvMs $rdy) - $t0)
        Fill-BootFigures $h $s.Log
      }
      $srch = Do-Search $s $t.Key1 90
      if ($null -eq $srch) { $h.note = 'the search never put a record on screen'; Rec $h; continue }
      $h.search_ms = '{0:F1}' -f $srch.Ms
      $prev = Snapshot-Before $s
      $sheet = [RdvPkg]::BiggestSheet($prev)
      $mk = Do-Mark $s $SaveTimeoutSec
      if (-not $mk.Ok) { $h.note = $mk.Note; Rec $h; continue }
      $h.b_ms = '{0:F1}' -f $mk.Ms
      Fill-SaveFigures $h $mk.Line
      $v = Verify-Mark $s $prev $sheet $t.SheetRow
      $h.reread = $v.Reread; $h.sheet_diff = $v.SheetDiff; $h.other_diffs = $v.Other; $h.raw_entries = $v.Raw
      $prev = $null
      $h.note = $(if ($mk.FeSaw) { 'ok' } else { 'ok (the BE answered, but the FE never rendered the confirmation)' })
      Rec $h
      # a mark the FE never saw leaves its exit guard held, and the next one
      # would be refused -- wait for the app's own decision before going on
      if (-not $mk.FeSaw) {
        Say '    the FE missed the confirmation; waiting for its own decision before the next round'
        [void](Wait-Log $s.Log '(exit released|stage=processed after_s=)' 0 240)
      }
      Say ("    round {0,2}  row {1,6}  B {2,9} ms   reread {3}  sheet_diff {4}  other {5}" -f `
        $i, $t.SheetRow, $h.b_ms, $h.reread, $h.sheet_diff, $(if ("$($h.other_diffs)".Length -gt 50) { "$($h.other_diffs)".Substring(0, 50) } else { $h.other_diffs }))
    }
    # the one concurrency figure, after the ten so it cannot disturb them: the
    # same search, once on an idle app and once 200 ms into a save
    $searchKey = $targets[11].Key1
    $victim = $targets[10]
    $idle = Do-Search $s $searchKey 90
    if ($null -eq $idle) { Say '    idle reference search failed'; return }
    $pick = Do-Search $s $victim.Key1 90
    if ($null -eq $pick) { Say '    could not put the concurrency target on screen'; return }
    $h = @{ method = $method; variant = 'cont-busy'; trial = 0; ledger_row = $victim.Row; sheet_row = $victim.SheetRow; key1 = $victim.Key1 }
    $from = (Read-Log $s.Log).Count
    $tm0 = $s.W.Now
    Click-Cell $s 'rdvBtnProcessed'
    if (-not (Answer-Dialog '処理済みの確認' 'はい' 30)) { Say '    concurrency: no confirm dialog'; return }
    $req = Wait-Ev $s.W ('\|req\|mark\|') $tm0 30
    if ($null -eq $req) { Say '    concurrency: no save request'; return }
    Start-Sleep -Milliseconds 200
    Set-Key $s $searchKey
    $ts0 = $s.W.Now
    Click-Cell $s 'rdvBtnSearch'
    $sreq = Wait-Ev $s.W ('\|req\|search\|') $ts0 30
    $mres = Wait-Ev $s.W ('\|agg\|(RESULT|MARK)\|.*res=marked') (EvMs $req) $SaveTimeoutSec
    $sres = $null
    if ($null -ne $sreq) { $sres = Wait-Ev $s.W ('\|agg\|RESULT\|.*res=single') (EvMs $sreq) $SaveTimeoutSec }
    if ($null -ne $mres) { $h.b_ms = '{0:F1}' -f ((EvMs $mres) - (EvMs $req)) }
    if ($null -ne $sres -and $null -ne $sreq) {
      $busy = (EvMs $sres) - (EvMs $sreq)
      $h.search_ms = '{0:F1}' -f $busy
      $h.note = ("search issued 200 ms into a save: {0:N1} ms, idle {1:N1} ms, increment {2:N1} ms" -f $busy, $idle.Ms, ($busy - $idle.Ms))
    }
    else { $h.note = 'the concurrent search never answered' }
    # and what the SCREEN made of it: the aggregate keeps one record per kind,
    # so a search answered right after a mark can replace the mark's own
    # confirmation before the FE's pump reads it
    $seen = Wait-Log $s.Log 'processed\tkey2=' $from 20
    if ($null -eq $seen) {
      $h.note = $h.note + '; the FE never rendered the save confirmation (the search RESULT replaced it in the aggregate) -- the record IS on disk, but the app holds the exit until its 180 s ceiling and then calls the save undecided'
      Say '    the FE never rendered the confirmation of the concurrent save (see the note)'
    }
    Rec $h
    Say ("    " + $h.note)
  }
  catch { Say ("    harness error: " + $_.Exception.Message + " @" + $_.InvocationInfo.ScriptLineNumber) }
  finally {
    if ($null -ne $s) {
      Copy-Item $s.Log (Join-Path $scratch ("{0}.log" -f $tag)) -Force -ErrorAction SilentlyContinue
      Close-Session $s
      Copy-Item (Join-Path $s.Dir 'channel.txt') (Join-Path $scratch ("{0}.channel.txt" -f $tag)) -Force -ErrorAction SilentlyContinue
      if (-not $KeepScratch) { Remove-Item $s.Dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
  }
}

# ---------------------------------------------------------------------------
# -Mode race: the save-confirmation / search race, tried for real.
#
# A search answered around the same moment as a save used to share ONE channel
# slot with the save's confirmation, so whichever was published second erased
# the other (measured: work\race-evidence-before\). The confirmation now has its
# own slot (MARK) and names the request it answers. This mode drives the real
# app through the three orderings and checks, every round:
#
#   (a) the search answer is rendered, for the key that was searched
#   (b) the save confirmation is rendered
#   (c) the exit guard is released right after the confirmation (never the
#       180 s ceiling, never "undecided")
#   (d) an independent reader finds the flag on the record that was marked
#   (e) exactly one confirmation, for this request, and none ignored as foreign
#
# It also records what the channel actually did (which record was published
# second), so a round that did NOT exercise the race is visible as such.
# ---------------------------------------------------------------------------
$raceTsv = Join-Path $Root ("work\race-{0}.tsv" -f $stamp)
$raceCols = @('variant', 'session', 'round', 'mark_row', 'mark_key1', 'search_key1', 'b_ms',
              'release_ms', 'published_second', 'chk_search', 'chk_confirm', 'chk_release',
              'chk_persist', 'chk_unique', 'note')
$raceRows = New-Object System.Collections.ArrayList
function RaceRec([hashtable] $h) {
  if (-not (Test-Path -LiteralPath $raceTsv)) {
    Set-Content -LiteralPath $raceTsv -Value ($raceCols -join "`t") -Encoding utf8
  }
  $line = ($raceCols | ForEach-Object { $v = $h[$_]; if ($null -eq $v) { '' } else { ("$v" -replace "`t", ' ') } }) -join "`t"
  Add-Content -LiteralPath $raceTsv -Value $line -Encoding utf8
  [void]$raceRows.Add([pscustomobject]$h)
}
# the same click, with the delays cut down: the "search just before the save"
# ordering only exists if the second click lands before the FE's fast follow-up
# ticks have rendered the first answer
function Click-CellFast($s, [string] $addr) {
  $p = Point-Of $s $addr
  $x = $p.X; $y = $p.Y
  Assert-OnTarget $s $addr $x $y
  [void][RdvE2EWin]::SetCursorPos($x, $y)
  Start-Sleep -Milliseconds 25
  [RdvE2EWin]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero)
  Start-Sleep -Milliseconds 20
  [RdvE2EWin]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)
}
# issue a search and DO NOT wait for it
function Issue-Search($s, [string] $key) {
  Set-Key $s $key
  Click-CellFast $s 'rdvBtnSearch'
}
# [処理済み], and make sure the click actually landed: a real mouse click goes
# wherever the foreground is, so one lost click must not be read as a product
# failure. Returns $true once the confirm dialog is up.
function Click-Processed($s, [bool] $fast) {
  for ($try = 0; $try -lt 2; $try++) {
    if ($try -gt 0) {
      Write-Host '    the confirm dialog did not appear; clicking [処理済み] once more'
      [void][RdvE2EWin]::SetForegroundWindow($s.Hwnd)
      Start-Sleep -Milliseconds 300
    }
    if ($fast) { Click-CellFast $s 'rdvBtnProcessed' } else { Click-Cell $s 'rdvBtnProcessed' }
    for ($i = 0; $i -lt 60; $i++) {
      if ([RdvE2EWin]::FindWindowW('#32770', '処理済みの確認') -ne [IntPtr]::Zero) { return $true }
      Start-Sleep -Milliseconds 100
    }
  }
  return $false
}

function Run-Race([string] $variant, [int] $session, [int] $rounds) {
  $tag = "race-{0}-{1:d2}" -f $variant, $session
  Say ("=== {0}: {1} round(s)" -f $tag, $rounds)
  $s = $null
  try {
    $s = New-Session 'book' 'race' $tag
    $lease = Wait-Ev $s.W ('\|file\|rdv3_.*_fe_lease\.lock\|') 0 90
    if ($null -eq $lease) { RaceRec @{ variant = $variant; session = $session; round = -1; note = 'the FE never asked the BE for anything' }; return }
    $rdy = Wait-Ev $s.W ('\|agg\|READY\|') (EvMs $lease) $ReadyTimeoutSec
    if ($null -eq $rdy) { RaceRec @{ variant = $variant; session = $session; round = -1; note = 'never reached READY' }; return }
    # the BE is up now (it just answered READY); name its process while a
    # window handle still resolves to one, and say what this session owns
    Add-BePid $s
    Say ('    ' + (Owned-Note $s))
    Start-Sleep -Milliseconds 700

    for ($r = 0; $r -lt $rounds; $r++) {
      $tA = $targets[(($session * $rounds) + $r) % 10]
      $tB = $targets[10 + ($r % 2)]
      $h = @{ variant = $variant; session = $session; round = $r; mark_row = $tA.SheetRow
              mark_key1 = $tA.Key1; search_key1 = $tB.Key1 }
      [void][RdvE2EWin]::SetForegroundWindow($s.Hwnd)
      Start-Sleep -Milliseconds 200

      # put the record to be marked on screen (this search is not part of the
      # race; it is how a person selects the record)
      if ($variant -ne 'searchonly') {
        $sel = Do-Search $s $tA.Key1 90
        if ($null -eq $sel) { $h.note = 'could not put the record on screen'; RaceRec $h; continue }
      }
      $from = (Read-Log $s.Log).Count
      $t0 = $s.W.Now
      $markReq = $null
      $searchReq = $null

      # NOTE: `continue` inside a switch belongs to the SWITCH in PowerShell,
      # not to the enclosing loop, so failures are carried out in a flag.
      $fail = ''
      switch ($variant) {
        'during' {
          if (Click-Processed $s $false) {
            [void](Answer-Dialog '処理済みの確認' 'はい' 30)
            $markReq = Wait-Ev $s.W ('\|req\|mark\|') $t0 30
            if ($null -ne $markReq) {
              Start-Sleep -Milliseconds 200
              Issue-Search $s $tB.Key1
            }
            else { $fail = 'the FE never issued the save request' }
          }
          else { $fail = 'the confirm dialog never appeared' }
        }
        'before' {
          # search first, then the save, fast enough that the answer to the
          # search is still unread when the save finishes
          Issue-Search $s $tB.Key1
          if (Click-Processed $s $true) {
            [void](Answer-Dialog '処理済みの確認' 'はい' 30)
            $markReq = Wait-Ev $s.W ('\|req\|mark\|') $t0 30
            if ($null -eq $markReq) { $fail = 'the FE never issued the save request' }
          }
          else { $fail = 'the confirm dialog never appeared' }
        }
        'after' {
          if (Click-Processed $s $false) {
            [void](Answer-Dialog '処理済みの確認' 'はい' 30)
            $markReq = Wait-Ev $s.W ('\|req\|mark\|') $t0 30
            if ($null -ne $markReq) {
              # the instant the BE's answer lands in the channel -- before the
              # FE's next pump tick can have taken it
              [void](Wait-Ev $s.W ('\|agg\|(RESULT|MARK)\|.*res=marked') (EvMs $markReq) $SaveTimeoutSec)
              Issue-Search $s $tB.Key1
            }
            else { $fail = 'the FE never issued the save request' }
          }
          else { $fail = 'the confirm dialog never appeared' }
        }
        'solo' {
          if (Click-Processed $s $false) {
            [void](Answer-Dialog '処理済みの確認' 'はい' 30)
            $markReq = Wait-Ev $s.W ('\|req\|mark\|') $t0 30
            if ($null -eq $markReq) { $fail = 'the FE never issued the save request' }
          }
          else { $fail = 'the confirm dialog never appeared' }
        }
        'searchonly' {
          Issue-Search $s $tB.Key1
        }
      }
      if ($fail -ne '') { $h.note = $fail; RaceRec $h; Say ("    NOTE: " + $fail); continue }

      # --- what the channel carried ---
      $markEv = $null
      if ($variant -ne 'searchonly') {
        $markEv = Wait-Ev $s.W ('\|agg\|(RESULT|MARK)\|.*res=marked') (EvMs $markReq) $SaveTimeoutSec
        if ($null -eq $markEv) { $h.note = 'the BE never confirmed the save'; RaceRec $h; continue }
        $h.b_ms = '{0:F1}' -f ((EvMs $markEv) - (EvMs $markReq))
      }
      $searchEv = $null
      if ($variant -ne 'solo') {
        $searchEv = Wait-Ev $s.W ('\|agg\|RESULT\|.*key=' + $tB.Key1) $t0 90
      }
      if ($null -ne $markEv -and $null -ne $searchEv) {
        $h.published_second = $(if ((EvMs $searchEv) -gt (EvMs $markEv)) { 'search' } else { 'confirmation' })
      }

      # --- (a) the search answer reached the screen, for the right key ---
      if ($variant -ne 'solo') {
        $ln = Wait-Log $s.Log ("`tsearch`tkey=" + $tB.Key1 + " .*hits=1") $from 60
        $h.chk_search = $(if ($null -ne $ln) { 'PASS' } else { 'FAIL' })
      }
      else { $h.chk_search = '-' }

      # --- (b) the confirmation reached the screen, for the right record ---
      if ($variant -ne 'searchonly') {
        $done = Wait-Log $s.Log ('processed\tkey2=.* row=' + ($tA.Row + 1) + ' value=TRUE') $from 60
        $h.chk_confirm = $(if ($null -ne $done) { 'PASS' } else { 'FAIL' })
        $seenAt = $s.W.Now
        # --- (c) released right away, not by the ceiling ---
        $rel = Wait-Log $s.Log 'processed save decided \(saved\); exit released' $from 60
        $all = Read-Log $s.Log
        $late = @($all[$from..([Math]::Max($from, $all.Count - 1))] | Where-Object { $_ -match 'after_s=|保存未確定|確定しません' }).Count
        $h.release_ms = '{0:F0}' -f ($seenAt - (EvMs $markEv))
        $h.chk_release = $(if ($null -ne $rel -and $late -eq 0 -and ($seenAt - (EvMs $markEv)) -lt 5000) { 'PASS' } else { 'FAIL' })
      }
      else { $h.chk_confirm = '-'; $h.chk_release = '-' }

      # --- (d) an independent reader finds it on disk ---
      if ($variant -ne 'searchonly') {
        Copy-Item $s.Ledger $s.Snap -Force
        $flag = [RdvPkg]::Flag($s.Snap, $tA.SheetRow)
        $h.chk_persist = $(if ($flag -match '<v>1</v>') { 'PASS' } else { 'FAIL' })
      }
      else { $h.chk_persist = '-' }

      # --- (e) exactly one confirmation, none ignored as foreign ---
      $all = Read-Log $s.Log
      $tail = @($all[$from..([Math]::Max($from, $all.Count - 1))])
      $nConf = @($tail | Where-Object { $_ -match 'processed\tkey2=' }).Count
      $nIgn = @($tail | Where-Object { $_ -match 'confirmation ignored' }).Count
      if ($variant -eq 'searchonly') { $h.chk_unique = $(if ($nConf -eq 0 -and $nIgn -eq 0) { 'PASS' } else { 'FAIL' }) }
      else { $h.chk_unique = $(if ($nConf -eq 1 -and $nIgn -eq 0) { 'PASS' } else { 'FAIL' }) }
      $h.note = "confirmations=$nConf ignored=$nIgn"
      RaceRec $h
      Say ("    {0} r{1}  row {2}  search {3}  confirm {4}  release {5} ({6} ms)  persist {7}  unique {8}  2nd={9}" -f `
        $variant, $r, $tA.SheetRow, $h.chk_search, $h.chk_confirm, $h.chk_release, $h.release_ms, $h.chk_persist, $h.chk_unique, $h.published_second)
    }
  }
  catch { Say ("    harness error: " + $_.Exception.Message + " @" + $_.InvocationInfo.ScriptLineNumber) }
  finally {
    if ($null -ne $s) {
      Copy-Item $s.Log (Join-Path $scratch ("{0}.log" -f $tag)) -Force -ErrorAction SilentlyContinue
      Close-Session $s
      Copy-Item (Join-Path $s.Dir 'channel.txt') (Join-Path $scratch ("{0}.channel.txt" -f $tag)) -Force -ErrorAction SilentlyContinue
      if (-not $KeepScratch) { Remove-Item $s.Dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
  }
}

# ---------------------------------------------------------------------------
# the C# build, as a REFERENCE series next to the three VBA methods. It has one
# save path and it cannot be swapped, so it is not one of the three; it is here
# to answer "how long does the user wait" for that build on the same two jobs.
# Its own log carries millisecond timestamps and it is a single process, so the
# same two boundaries are read straight out of it:
#   A  "decision check started" (written immediately before the job is posted
#      to the worker thread) -> "decision ready rows=..."
#   B  "processed save started ... (exit held until it is decided)" ->
#      "processed key2=... value=TRUE" (written when the worker finished and
#      handed the result to the UI thread)
# ---------------------------------------------------------------------------
function LogStamp([string] $line) {
  $t = $line.Substring(0, 23)
  return [datetime]::ParseExact($t, 'yyyy-MM-dd HH:mm:ss.fff', [Globalization.CultureInfo]::InvariantCulture)
}
function Run-ColdCSharp([int] $i) {
  $t = $targets[$i % 10]
  $tag = "csharp-cold-{0:d2}" -f $i
  Say ("=== {0}  target ledger row {1} (key1 {2})" -f $tag, $t.Row, $t.Key1)
  $h = @{ method = 'csharp'; variant = 'cold'; trial = $i; ledger_row = $t.Row; sheet_row = ($t.Row + 2); key1 = $t.Key1 }
  $dir = Join-Path $scratch $tag
  $proc = $null
  $main = [IntPtr]::Zero
  try {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    Copy-Item -Path (Join-Path $Root 'dist\app-csharp\*') -Destination $dir -Recurse -Force
    $log = Join-Path $dir 'ReaderDataViewer.log'
    $ledger = Join-Path $dir 'ReaderDataViewer-Ledger.xlsx'
    if (Test-Path -LiteralPath $log) { Remove-Item -LiteralPath $log -Force }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $proc = Start-Process -FilePath (Join-Path $dir 'ReaderDataViewer.cmd') -PassThru
    $ready = Wait-Log $log "`tdecision`tready " 0 $ReadyTimeoutSec
    if ($null -eq $ready) { $h.note = 'never reached READY'; return }
    $h.xl_start_ms = '{0:F1}' -f $sw.Elapsed.TotalMilliseconds
    $all = Read-Log $log
    $started = @($all | Where-Object { $_ -match "`tdecision`tcheck started" })[0]
    $h.a_ready_ms = '{0:F1}' -f ((LogStamp $ready) - (LogStamp $started)).TotalMilliseconds
    $h.a_check_ms = $h.a_ready_ms
    foreach ($ln in $all) { if ($ln -match 'boot_to_ready_ms=([0-9.]+)') { $h.fe_boot_ready_ms = $Matches[1] } }

    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $w = $walker.GetFirstChild([System.Windows.Automation.AutomationElement]::RootElement)
    while ($null -ne $w) {
      if ($w.Current.Name -like 'Reader Data Viewer*' -and $w.Current.ProcessId -ne $PID) { $main = [IntPtr]$w.Current.NativeWindowHandle; break }
      $w = $walker.GetNextSibling($w)
    }
    if ($main -eq [IntPtr]::Zero) { $h.note = 'main window not found'; return }
    Park-Window $main
    Start-Sleep -Milliseconds 400
    $editH = [IntPtr]::Zero
    foreach ($k in [RdvE2EWin]::Kids($main)) { if ($k.Cls -like '*.EDIT.*') { $editH = [IntPtr]$k.Hwnd; break } }
    if ($editH -eq [IntPtr]::Zero) { $h.note = 'search box not found'; return }
    foreach ($ch in $t.Key1.ToCharArray()) {
      [void][RdvE2EWin]::PostMessage($editH, 0x0102, [IntPtr][int]$ch, [IntPtr]0)
      Start-Sleep -Milliseconds 20
    }
    Start-Sleep -Milliseconds 250
    $from = (Read-Log $log).Count
    foreach ($k in [RdvE2EWin]::Kids($main)) { if ($k.Text -eq '検索') { [void][RdvE2EWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break } }
    $hit = Wait-Log $log ("`tsearch`tkey=" + $t.Key1 + " ") $from 60
    if ($null -eq $hit) { $h.note = 'no search line'; return }
    if ($hit -match 'hits=1 ms=([0-9.]+)') { $h.search_ms = $Matches[1] }

    $beforeRows = [RdvPkg]::TrueRows($ledger)
    $before = Join-Path $dir 'before.xlsx'
    Copy-Item $ledger $before -Force
    $prev = [RdvPkg]::Parts($before)
    $sheet = [RdvPkg]::BiggestSheet($prev)

    $from = (Read-Log $log).Count
    foreach ($k in [RdvE2EWin]::Kids($main)) { if ($k.Text -eq '処理済み') { [void][RdvE2EWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break } }
    if (-not (Answer-Dialog '処理済みの確認' 'はい' 30)) { $h.note = 'the confirm dialog never appeared'; return }
    $st = Wait-Log $log 'save started .*exit held' $from 30
    $done = Wait-Log $log "`tprocessed`tkey2=.*value=TRUE" $from $SaveTimeoutSec
    if ($null -eq $st -or $null -eq $done) { $h.note = 'no save start/finish line'; return }
    $h.b_ms = '{0:F1}' -f ((LogStamp $done) - (LogStamp $st)).TotalMilliseconds
    if ($done -match 'persist_ms=([0-9.]+)') { $h.b_save_ms = $Matches[1] }
    if ($done -match 'e2e_ms=([0-9.]+)') { $h.b_fe_e2e_ms = $Matches[1] }

    $afterRows = [RdvPkg]::TrueRows($ledger)
    $added = @($afterRows | Where-Object { $beforeRows -notcontains $_ })
    $h.reread = $(if ($added.Count -eq 1 -and $added[0] -match (":" + $t.Key1 + ":")) { 'ok' } else { 'FAIL' })
    $now = [RdvPkg]::Parts($ledger)
    $sd = 0
    $rep = [RdvPkg]::DiffReport($prev, $now, $sheet, [ref]$sd)
    $h.sheet_diff = $sd
    $h.other_diffs = $(if ($rep -eq '') { 'none' } else { $rep })
    $h.raw_entries = $(if (([RdvPkg]::RawReport($before, $ledger, $sheet)) -eq '') { 'none' } else { [RdvPkg]::RawReport($before, $ledger, $sheet) })
    $prev = $null; $now = $null
    $h.note = ("ok; flagged row " + ($added -join ','))
    Say ("    A {0} ms   B {1} ms   reread {2}  sheet_diff {3}" -f $h.a_ready_ms, $h.b_ms, $h.reread, $h.sheet_diff)
  }
  catch { $h.note = "harness error: " + $_.Exception.Message + " @" + $_.InvocationInfo.ScriptLineNumber }
  finally {
    if ($main -ne [IntPtr]::Zero -and [RdvE2EWin]::IsWindow($main)) {
      [void][RdvE2EWin]::PostMessage($main, 0x0010, [IntPtr]0, [IntPtr]0)
      Start-Sleep -Seconds 3
    }
    foreach ($ph in @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object { $_.CommandLine -like '*RDV_SELF*' })) {
      $q = Get-Process -Id $ph.ProcessId -ErrorAction SilentlyContinue
      if ($q) { Say ("  closing my app host " + $ph.ProcessId); $q.Kill() }
    }
    if ($null -ne $proc -and -not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    Copy-Item (Join-Path $dir 'ReaderDataViewer.log') (Join-Path $scratch ("{0}.log" -f $tag)) -Force -ErrorAction SilentlyContinue
    if (-not $KeepScratch) { Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
    if ("$($h.note)" -notlike 'ok*') { Say ("    NOTE: " + $h.note) }
    Rec $h
  }
}

# ---------------------------------------------------------------------------
# run
# ---------------------------------------------------------------------------
if ($Build -eq 'csharp') {
  Add-Type -AssemblyName UIAutomationClient
  Add-Type -AssemblyName UIAutomationTypes
  for ($i = 0; $i -lt $Trials; $i++) { Run-ColdCSharp $i }
}
$methods = if ($Method -eq 'all') { @('book', 'ado', 'zip') } else { @($Method) }
$modes = if ($Mode -eq 'all') { @('cold', 'coldapply', 'cont') } else { @($Mode) }
if ($Build -eq 'vba') {
  foreach ($md in $modes) {
    if ($md -eq 'race') {
      # 3 orderings x 2 sessions x 2 rounds = 12 raced saves, plus the two
      # controls (a save with no search, a search with no save)
      foreach ($v in @('during', 'before', 'after')) {
        for ($i = 0; $i -lt 2; $i++) { Run-Race $v $i 2 }
      }
      Run-Race 'solo' 0 2
      Run-Race 'searchonly' 0 2
      continue
    }
    foreach ($mth in $methods) {
      if ($md -eq 'cont') { Run-Cont $mth; continue }
      for ($i = 0; $i -lt $Trials; $i++) { Run-Cold $mth $md $i }
    }
  }
}

# ---------------------------------------------------------------------------
# summary
# ---------------------------------------------------------------------------
function Series([string] $mth, [string] $variant, [string] $field) {
  $set = @($rows | Where-Object { $_.method -eq $mth -and $_.variant -eq $variant })
  $out = @()
  foreach ($r in $set) { $v = $r.$field; if ($null -ne $v -and "$v" -ne '') { $out += [double]$v } }
  return $out
}
Write-Output ''
Write-Output ("=== summary -> {0}" -f $tsv)
foreach ($variant in @('cold', 'coldapply', 'cont')) {
  foreach ($mth in @('book', 'ado', 'zip', 'csharp')) {
    $b = Series $mth $variant 'b_ms'
    if ($b.Count -eq 0) { continue }
    $set = @($rows | Where-Object { $_.method -eq $mth -and $_.variant -eq $variant -and $_.b_ms })
    $okReads = @($set | Where-Object { $_.reread -eq 'ok' }).Count
    $clean = @($set | Where-Object { $_.other_diffs -eq 'none' -and "$($_.sheet_diff)" -eq '1' -and $_.raw_entries -eq 'none' }).Count
    Write-Output ("{0,-10} {1,-5} B: n={2} max={3:N1} first={4:N1} min={5:N1} ms   reread {6}/{2}   one byte and nothing else {7}/{2}" -f `
      $variant, $mth, $b.Count, ($b | Measure-Object -Maximum).Maximum, $b[0], ($b | Measure-Object -Minimum).Minimum, $okReads, $clean)
    Write-Output ("             all B: " + (($b | ForEach-Object { '{0:N1}' -f $_ }) -join '  '))
    foreach ($f in @('a_check_ms', 'a_ready_ms', 'apply_ms', 'spawn_ms', 'fe_boot_ready_ms',
                     'xl_start_ms', 'fe_launch_ms', 'b_save_ms', 'b_fe_e2e_ms', 'search_ms')) {
      $v = Series $mth $variant $f
      if ($v.Count -eq 0) { continue }
      Write-Output ("             {0,-13} n={1} max={2:N1} min={3:N1} median={4:N1} ms" -f `
        $f, $v.Count, ($v | Measure-Object -Maximum).Maximum, ($v | Measure-Object -Minimum).Minimum, (($v | Sort-Object)[[int]($v.Count / 2)]))
    }
  }
}
foreach ($r in @($rows | Where-Object { $_.variant -eq 'cont-busy' })) { Write-Output ("cont-busy  {0,-5} {1}" -f $r.method, $r.note) }
if ($raceRows.Count -gt 0) {
  Write-Output ''
  Write-Output ("=== race checks -> {0}" -f $raceTsv)
  $names = @('chk_search', 'chk_confirm', 'chk_release', 'chk_persist', 'chk_unique')
  $pass = 0; $fail = 0
  foreach ($r in $raceRows) {
    foreach ($n in $names) {
      $v = $r.$n
      if ($v -eq 'PASS') { $pass++ } elseif ($v -eq 'FAIL') { $fail++ }
    }
  }
  foreach ($v in @('during', 'before', 'after', 'solo', 'searchonly')) {
    $set = @($raceRows | Where-Object { $_.variant -eq $v })
    if ($set.Count -eq 0) { continue }
    $bad = @($set | Where-Object { $_.chk_search -eq 'FAIL' -or $_.chk_confirm -eq 'FAIL' -or $_.chk_release -eq 'FAIL' -or $_.chk_persist -eq 'FAIL' -or $_.chk_unique -eq 'FAIL' -or ("$($_.note)" -notlike 'confirmations=*') }).Count
    $second = (@($set | ForEach-Object { $_.published_second }) | Where-Object { $_ }) -join ','
    $rel = @($set | Where-Object { $_.release_ms } | ForEach-Object { [double]$_.release_ms })
    Write-Output ("{0,-11} rounds={1} bad={2}  release_ms max={3}  published second: {4}" -f `
      $v, $set.Count, $bad, $(if ($rel.Count) { '{0:N0}' -f ($rel | Measure-Object -Maximum).Maximum } else { '-' }), $second)
  }
  Write-Output ("checks: {0} PASS, {1} FAIL" -f $pass, $fail)
  foreach ($r in $raceRows) {
    foreach ($n in $names) { if ($r.$n -eq 'FAIL') { Write-Output ("  FAIL {0}/{1} r{2}: {3} ({4})" -f $r.variant, $r.session, $r.round, $n, $r.note) } }
  }
}
Write-Output ("scratch: {0}" -f $scratch)
