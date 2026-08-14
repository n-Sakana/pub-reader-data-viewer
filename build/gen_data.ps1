# ============================================================================
# gen_data.ps1 -- build the three synthetic CSVs every method reads.
#
#   data\tableA.csv   1,000,000 data rows x 10 fields, keyed by key1
#   data\tableB.csv   1,000,000 data rows x 10 fields, keyed by key1 AND key2
#   data\tableC.csv   1,000,000 data rows x 10 fields, keyed by key2
#   data\expected.txt the oracle: row count, join checksum, sample records
#
# The three implementations (VBA / C# / hybrid) read these exact bytes. Nothing
# generates its own data, so "same input" is a property of the files, not of
# three separate generators agreeing.
#
# Everything here is deterministic: same script, same bytes, every time.
#   key1  "00000001".."01000000"   8 digits, leading zeros are significant
#   key2  a fixed permutation of the same space, so B->C is not the same order
#         as A->B and the joins have to actually probe a hash map
#
# A is written sorted by key1, C sorted by key2, B in a seeded shuffle. That is
# what a master export plus a transaction export looks like, and it stops any
# implementation from getting the right answer by walking the files in step.
#
#   pwsh -File build\gen_data.ps1            (or Windows PowerShell 5.1)
#   pwsh -File build\gen_data.ps1 -Rows 1000 -OutDir data-small
# ============================================================================
[CmdletBinding()]
param(
  [int]    $Rows   = 1000000,
  [string] $OutDir = "",
  [switch] $Force
)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if ([string]::IsNullOrEmpty($OutDir)) { $OutDir = Join-Path $root 'data' }
if (-not [IO.Path]::IsPathRooted($OutDir)) { $OutDir = Join-Path $root $OutDir }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$marker = Join-Path $OutDir 'expected.txt'
if ((Test-Path -LiteralPath $marker) -and -not $Force) {
  $have = Select-String -LiteralPath $marker -Pattern '^rows=(\d+)$' | Select-Object -First 1
  if ($have -and [int]$have.Matches[0].Groups[1].Value -eq $Rows) {
    Write-Output "already generated: $OutDir (rows=$Rows). -Force to rebuild."
    Get-Content -LiteralPath $marker | Select-Object -First 6 | ForEach-Object { Write-Output "  $_" }
    return
  }
}

$cs = @'
using System;
using System.Globalization;
using System.IO;
using System.Text;

public static class RdvGen
{
    // key2 = a bijection of 1..N onto itself. 524287 is prime and does not
    // divide N, so n -> (n-1)*524287 + 12345 (mod N) + 1 is one-to-one.
    private const int  PermMul = 524287;
    private const int  PermAdd = 12345;
    private const long Mod     = 1000000007L;

    private static int N;

    private static int Perm(int n)
    {
        long v = ((long)(n - 1) * PermMul + PermAdd) % N;
        return (int)v + 1;
    }

    // one small deterministic value per (table, row, slot); order-independent,
    // so a row looks the same no matter where it lands in the file
    private static uint Mix(uint table, uint n, uint slot)
    {
        uint h = 2166136261u;
        h = (h ^ table) * 16777619u;
        h = (h ^ n) * 16777619u;
        h = (h ^ slot) * 16777619u;
        h ^= h >> 13;
        return h;
    }

    private static readonly string[] Grades  = { "A1", "A2", "B1", "B2", "C1" };
    private static readonly string[] Status  = { "OPEN", "HOLD", "DONE", "VOID" };
    private static readonly string[] Cats    = { "ELEC", "MECH", "CHEM", "FOOD", "PAPR", "MISC" };

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

    public static string RowB(int n)
    {
        StringBuilder b = new StringBuilder(96);
        int qty  = (int)(Mix(2, (uint)n, 1) % 999u) + 1;
        int unit = (int)(Mix(2, (uint)n, 2) % 99989u) + 10;
        b.Append(Pad(n, 8)).Append(',');
        b.Append(Pad(Perm(n), 8)).Append(',');
        b.Append("SL").Append(Pad(n, 8)).Append(',');
        b.Append(Date(Mix(2, (uint)n, 3))).Append(',');
        b.Append(Pad(qty, 3)).Append(',');
        b.Append(Pad(unit, 5)).Append(',');
        b.Append(Pad(qty * unit, 8)).Append(',');
        b.Append(Status[Mix(2, (uint)n, 4) % 4u]).Append(',');
        b.Append(Pad((int)(Mix(2, (uint)n, 5) % 999u) + 1, 3)).Append(',');
        b.Append("MEMO-").Append(Pad((int)(Mix(2, (uint)n, 6) % 99999u), 5));
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
        // UTF-8 without BOM, CRLF, ASCII payload. The BOM matters: a leading
        // BOM would land inside the first key of the header line.
        FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 20);
        StreamWriter w = new StreamWriter(fs, new UTF8Encoding(false), 1 << 20);
        w.NewLine = "\r\n";
        return w;
    }

    public static string Run(string dir, int rows)
    {
        N = rows;
        int[] order = new int[N];
        for (int i = 0; i < N; i++) { order[i] = i + 1; }

        // seeded Fisher-Yates: B is a transaction export, not a sorted master
        ulong s = 88172645463325252UL;
        for (int i = N - 1; i > 0; i--)
        {
            s ^= s << 13; s ^= s >> 7; s ^= s << 17;
            int j = (int)(s % (ulong)(i + 1));
            int t = order[i]; order[i] = order[j]; order[j] = t;
        }

        int[] samples = new int[] { 1, 42, N / 2, N };
        string[] sampleA = new string[samples.Length];
        string[] sampleB = new string[samples.Length];
        string[] sampleC = new string[samples.Length];

        StreamWriter w = Open(Path.Combine(dir, "tableA.csv"));
        w.WriteLine(HeadA);
        for (int n = 1; n <= N; n++) { w.WriteLine(RowA(n)); }
        w.Flush(); w.Close();

        // A is sorted by key1, so key1 n sits at data-row index n-1.
        // C is sorted by key2, so key2 k sits at data-row index k-1.
        // The checksum walks B in file order and folds in both partner indices,
        // which is only reproducible if every row of both joins matched.
        long chk = 0;
        w = Open(Path.Combine(dir, "tableB.csv"));
        w.WriteLine(HeadB);
        for (int i = 0; i < N; i++)
        {
            int n = order[i];
            w.WriteLine(RowB(n));
            long ai = n - 1;
            long ci = Perm(n) - 1;
            chk = (chk * 31 + ai + ci) % Mod;
        }
        w.Flush(); w.Close();

        w = Open(Path.Combine(dir, "tableC.csv"));
        w.WriteLine(HeadC);
        for (int k = 1; k <= N; k++) { w.WriteLine(RowC(k)); }
        w.Flush(); w.Close();

        for (int i = 0; i < samples.Length; i++)
        {
            int n = samples[i];
            sampleA[i] = RowA(n);
            sampleB[i] = RowB(n);
            sampleC[i] = RowC(Perm(n));
        }

        StringBuilder exp = new StringBuilder();
        exp.Append("rows=").Append(N.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
        exp.Append("fields=10\r\n");
        exp.Append("keylen=8\r\n");
        exp.Append("joinchecksum=").Append(chk.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
        exp.Append("headA=").Append(HeadA).Append("\r\n");
        exp.Append("headB=").Append(HeadB).Append("\r\n");
        exp.Append("headC=").Append(HeadC).Append("\r\n");
        for (int i = 0; i < samples.Length; i++)
        {
            string tag = "sample" + (i + 1).ToString(CultureInfo.InvariantCulture);
            exp.Append(tag).Append(".key1=").Append(Pad(samples[i], 8)).Append("\r\n");
            exp.Append(tag).Append(".key2=").Append(Pad(Perm(samples[i]), 8)).Append("\r\n");
            exp.Append(tag).Append(".a=").Append(sampleA[i]).Append("\r\n");
            exp.Append(tag).Append(".b=").Append(sampleB[i]).Append("\r\n");
            exp.Append(tag).Append(".c=").Append(sampleC[i]).Append("\r\n");
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
$chk = [RdvGen]::Run($OutDir, $Rows)
$sw.Stop()

Write-Output ("done in {0:N1} s   joinchecksum={1}" -f $sw.Elapsed.TotalSeconds, $chk)
foreach ($f in 'tableA.csv', 'tableB.csv', 'tableC.csv', 'expected.txt') {
  $fi = Get-Item -LiteralPath (Join-Path $OutDir $f)
  Write-Output ("  {0,-14} {1,14:N0} bytes" -f $fi.Name, $fi.Length)
}
