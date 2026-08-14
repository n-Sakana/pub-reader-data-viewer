# ============================================================================
# gen_data2.ps1 -- the one-to-many dataset the four index methods are compared on.
#
#   data-100k\tableA.csv    100,000 rows x 10 fields, key1 unique
#   data-100k\tableB.csv    100,000 rows x 10 fields, key1 REPEATS, key2 unique
#   data-100k\tableC.csv    100,000 rows x 10 fields, key2 unique
#   data-100k\expected.txt  the oracle: row counts, join checksum, candidate counts
#
# The difference from gen_data.ps1 is the shape of B, and it is the whole point
# of this dataset: key 1 is legitimately one-to-many. B is the middle table -- a
# slip with several lines - so one key 1 can own 1, 2, 3 or 5 rows of B, each
# with its own key 2. Searching by key 1 therefore returns a SET of candidates,
# and an index that keeps only the last row for a key would silently lose them.
#
#   key1 block                     rows per key1   B rows
#   1        .. N/100                    5          N * 0.05
#   N/100+1  .. N/100 + N/20             3          N * 0.15
#   next N * 0.15 keys                   2          N * 0.30
#   next N * 0.50 keys                   1          N * 0.50
#   the rest                             0          -
#                                                   ------- N
#
# The blocks are contiguous and in key order, so any key's candidate count is
# known without reading the file: key1 00000001 has 5, and the small -Rows 200
# set has the same shape in miniature for testing.
#
# A stays sorted by key1 and C sorted by key2; B is written in a seeded shuffle.
#
#   pwsh -File build\gen_data2.ps1
#   pwsh -File build\gen_data2.ps1 -Rows 200 -OutDir data-tiny
# ============================================================================
[CmdletBinding()]
param(
  [int]    $Rows   = 100000,
  [string] $OutDir = "",
  [switch] $Force
)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if ([string]::IsNullOrEmpty($OutDir)) { $OutDir = Join-Path $root 'data-100k' }
if (-not [IO.Path]::IsPathRooted($OutDir)) { $OutDir = Join-Path $root $OutDir }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$marker = Join-Path $OutDir 'expected.txt'
if ((Test-Path -LiteralPath $marker) -and -not $Force) {
  $have = Select-String -LiteralPath $marker -Pattern '^rows=(\d+)$' | Select-Object -First 1
  if ($have -and [int]$have.Matches[0].Groups[1].Value -eq $Rows) {
    Write-Output "already generated: $OutDir (rows=$Rows). -Force to rebuild."
    Get-Content -LiteralPath $marker | Select-Object -First 8 | ForEach-Object { Write-Output "  $_" }
    return
  }
}

$cs = @'
using System;
using System.Globalization;
using System.IO;
using System.Text;

public static class RdvGen2
{
    private const int  PermMul = 524287;
    private const int  PermAdd = 12345;
    private const long Mod     = 1000000007L;

    private static int N;

    // key2 of the i-th B row: a bijection of 1..N onto itself, so C is a plain
    // 1:1 partner table and every B row points at a different C row
    private static int Perm(int n)
    {
        long v = ((long)(n - 1) * PermMul + PermAdd) % N;
        return (int)v + 1;
    }

    private static uint Mix(uint table, uint n, uint slot)
    {
        uint h = 2166136261u;
        h = (h ^ table) * 16777619u;
        h = (h ^ n) * 16777619u;
        h = (h ^ slot) * 16777619u;
        h ^= h >> 13;
        return h;
    }

    private static readonly string[] Grades = { "A1", "A2", "B1", "B2", "C1" };
    private static readonly string[] Status = { "OPEN", "HOLD", "DONE", "VOID" };
    private static readonly string[] Cats   = { "ELEC", "MECH", "CHEM", "FOOD", "PAPR", "MISC" };

    private static string Pad(int v, int w)
    {
        return v.ToString(CultureInfo.InvariantCulture).PadLeft(w, '0');
    }

    private static string Date(uint r)
    {
        int y = 2023 + (int)(r % 3u);
        int m = 1 + (int)((r / 3u) % 12u);
        int d = 1 + (int)((r / 36u) % 28u);
        return Pad(y, 4) + Pad(m, 2) + Pad(d, 2);
    }

    // how many B rows key1 n owns
    public static int Multiplicity(int n)
    {
        int b1 = N / 100;                 // 5 rows each
        int b2 = b1 + N / 20;             // 3 rows each
        int b3 = b2 + (N * 15) / 100;     // 2 rows each
        int b4 = b3 + N / 2;              // 1 row each
        if (n <= b1) { return 5; }
        if (n <= b2) { return 3; }
        if (n <= b3) { return 2; }
        if (n <= b4) { return 1; }
        return 0;
    }

    public static string RowA(int n)
    {
        StringBuilder b = new StringBuilder(96);
        b.Append(Pad(n, 8)).Append(',');
        b.Append("A").Append(Pad((int)(Mix(1, (uint)n, 1) % 99999u) + 1, 5)).Append(',');
        b.Append("CUSTOMER-").Append(Pad(n, 7)).Append(',');
        b.Append(Grades[Mix(1, (uint)n, 2) % 5u]).Append(',');
        b.Append(Date(Mix(1, (uint)n, 3))).Append(',');
        b.Append(Pad((int)(Mix(1, (uint)n, 4) % 9999999u) + 1, 7)).Append(',');
        b.Append("0.").Append(Pad((int)(Mix(1, (uint)n, 5) % 10000u), 4)).Append(',');
        b.Append((Mix(1, (uint)n, 6) % 2u) == 0u ? "Y" : "N").Append(',');
        b.Append("D").Append(Pad((int)(Mix(1, (uint)n, 7) % 999u) + 1, 3)).Append(',');
        b.Append("NOTE-").Append(Pad((int)(Mix(1, (uint)n, 8) % 99999u), 5));
        return b.ToString();
    }

    // one line of a slip: key1 is the slip, line is which line of it, key2 is
    // the item that line points at
    public static string RowB(int n, int line, int k2)
    {
        StringBuilder b = new StringBuilder(96);
        uint seed = (uint)(n * 8 + line);
        int qty  = (int)(Mix(2, seed, 1) % 999u) + 1;
        int unit = (int)(Mix(2, seed, 2) % 99989u) + 10;
        b.Append(Pad(n, 8)).Append(',');
        b.Append(Pad(k2, 8)).Append(',');
        b.Append("SL").Append(Pad(n, 8)).Append(',');
        b.Append(Date(Mix(2, seed, 3))).Append(',');
        b.Append(Pad(qty, 3)).Append(',');
        b.Append(Pad(unit, 5)).Append(',');
        b.Append(Pad(qty * unit, 8)).Append(',');
        b.Append(Status[Mix(2, seed, 4) % 4u]).Append(',');
        b.Append(Pad(line, 3)).Append(',');
        b.Append("MEMO-").Append(Pad((int)(Mix(2, seed, 6) % 99999u), 5));
        return b.ToString();
    }

    public static string RowC(int k2)
    {
        StringBuilder b = new StringBuilder(96);
        b.Append(Pad(k2, 8)).Append(',');
        b.Append("IT").Append(Pad((int)(Mix(3, (uint)k2, 1) % 999999u) + 1, 6)).Append(',');
        b.Append("MAKER-").Append(Pad((int)(Mix(3, (uint)k2, 2) % 9999u) + 1, 4)).Append(',');
        b.Append(Cats[Mix(3, (uint)k2, 3) % 6u]).Append(',');
        b.Append(Pad((int)(Mix(3, (uint)k2, 4) % 999999u) + 1, 6)).Append(',');
        b.Append(Pad((int)(Mix(3, (uint)k2, 5) % 99999u), 5)).Append(',');
        b.Append("L").Append(Pad((int)(Mix(3, (uint)k2, 6) % 99999u) + 1, 5)).Append(',');
        b.Append("LOT").Append(Pad((int)(Mix(3, (uint)k2, 7) % 99999u) + 1, 5)).Append(',');
        b.Append(Date(Mix(3, (uint)k2, 8))).Append(',');
        b.Append("RMK-").Append(Pad((int)(Mix(3, (uint)k2, 9) % 999999u), 6));
        return b.ToString();
    }

    public const string HeadA = "key1,a_code,a_name,a_grade,a_date,a_amount,a_rate,a_flag,a_dept,a_note";
    public const string HeadB = "key1,key2,b_slip,b_date,b_qty,b_unit,b_total,b_status,b_line,b_memo";
    public const string HeadC = "key2,c_item,c_maker,c_cat,c_price,c_stock,c_loc,c_lot,c_exp,c_remark";

    private static StreamWriter Open(string path)
    {
        FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 20);
        StreamWriter w = new StreamWriter(fs, new UTF8Encoding(false), 1 << 20);
        w.NewLine = "\r\n";
        return w;
    }

    public static string Run(string dir, int rows)
    {
        N = rows;

        // lay out every B row as (key1, line), then shuffle the order they are
        // written in. key2 is handed out in that shuffled order, so a B row's
        // key2 tells you nothing about its key1.
        int total = 0;
        for (int n = 1; n <= N; n++) { total += Multiplicity(n); }
        if (total != N)
        {
            throw new InvalidOperationException("B row count " + total.ToString(CultureInfo.InvariantCulture)
                + " != " + N.ToString(CultureInfo.InvariantCulture));
        }

        int[] rowKey = new int[N];
        int[] rowLine = new int[N];
        int at = 0;
        for (int n = 1; n <= N; n++)
        {
            int m = Multiplicity(n);
            for (int line = 1; line <= m; line++)
            {
                rowKey[at] = n;
                rowLine[at] = line;
                at++;
            }
        }

        ulong s = 88172645463325252UL;
        for (int i = N - 1; i > 0; i--)
        {
            s ^= s << 13; s ^= s >> 7; s ^= s << 17;
            int j = (int)(s % (ulong)(i + 1));
            int t = rowKey[i]; rowKey[i] = rowKey[j]; rowKey[j] = t;
            t = rowLine[i]; rowLine[i] = rowLine[j]; rowLine[j] = t;
        }

        StreamWriter w = Open(Path.Combine(dir, "tableA.csv"));
        w.WriteLine(HeadA);
        for (int n = 1; n <= N; n++) { w.WriteLine(RowA(n)); }
        w.Flush(); w.Close();

        // A is sorted by key1  -> key1 n is data row n-1
        // C is sorted by key2  -> key2 k is data row k-1
        long chk = 0;
        w = Open(Path.Combine(dir, "tableB.csv"));
        w.WriteLine(HeadB);
        for (int i = 0; i < N; i++)
        {
            int k2 = Perm(i + 1);
            w.WriteLine(RowB(rowKey[i], rowLine[i], k2));
            chk = (chk * 31 + (rowKey[i] - 1) + (k2 - 1)) % Mod;
        }
        w.Flush(); w.Close();

        w = Open(Path.Combine(dir, "tableC.csv"));
        w.WriteLine(HeadC);
        for (int k = 1; k <= N; k++) { w.WriteLine(RowC(k)); }
        w.Flush(); w.Close();

        StringBuilder exp = new StringBuilder();
        exp.Append("rows=").Append(N.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
        exp.Append("fields=10\r\n");
        exp.Append("keylen=8\r\n");
        exp.Append("shape=one-to-many\r\n");
        exp.Append("joinchecksum=").Append(chk.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
        exp.Append("headA=").Append(HeadA).Append("\r\n");
        exp.Append("headB=").Append(HeadB).Append("\r\n");
        exp.Append("headC=").Append(HeadC).Append("\r\n");

        // how many key1 values own each candidate count, and one worked example
        // of each so a test can name a key and know what it should get
        int[] want = new int[] { 5, 3, 2, 1, 0 };
        for (int i = 0; i < want.Length; i++)
        {
            int keys = 0;
            int first = 0;
            for (int n = 1; n <= N; n++)
            {
                if (Multiplicity(n) == want[i])
                {
                    keys++;
                    if (first == 0) { first = n; }
                }
            }
            exp.Append("cand").Append(want[i].ToString(CultureInfo.InvariantCulture))
               .Append(".keys=").Append(keys.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
            exp.Append("cand").Append(want[i].ToString(CultureInfo.InvariantCulture))
               .Append(".firstkey=").Append(first > 0 ? Pad(first, 8) : "").Append("\r\n");
        }
        File.WriteAllText(Path.Combine(dir, "expected.txt"), exp.ToString(), new UTF8Encoding(false));
        return chk.ToString(CultureInfo.InvariantCulture);
    }
}
'@

Write-Output "compiling generator..."
Add-Type -TypeDefinition $cs -Language CSharp

Write-Output "generating $Rows rows x 3 tables into $OutDir ..."
$sw = [Diagnostics.Stopwatch]::StartNew()
$chk = [RdvGen2]::Run($OutDir, $Rows)
$sw.Stop()

Write-Output ("done in {0:N1} s   joinchecksum={1}" -f $sw.Elapsed.TotalSeconds, $chk)
foreach ($f in 'tableA.csv', 'tableB.csv', 'tableC.csv', 'expected.txt') {
  $fi = Get-Item -LiteralPath (Join-Path $OutDir $f)
  Write-Output ("  {0,-14} {1,14:N0} bytes" -f $fi.Name, $fi.Length)
}
Get-Content -LiteralPath (Join-Path $OutDir 'expected.txt') | Select-String -Pattern '^cand' | ForEach-Object { Write-Output ("  " + $_) }
