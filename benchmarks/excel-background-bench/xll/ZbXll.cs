// ============================================================================
// ZbXll -- Excel-DNA / XLL による「Excel の中で走る」変換。
//
// 他の C# 方式との違いは 1 点だけ。プロセス境界が無いこと。
// 方式 3-6, 8-10, 13-15 は別プロセスの ZipWorker.exe が COM 越しに Excel を
// 触るが、この方式は Excel 自身のプロセスに .NET が読み込まれ、XLL の C API で
// 直接セル配列をやり取りする。COM のマーシャリングも RPC も通らない。
//
// 変換規則は ZipWorker.cs / modZipRule.bas / make_dict_reference.ps1 と同一。
// 郵便番号から数字以外を落として 7 桁に揃え、辞書を引き、無ければ「該当なし」。
// ============================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelDna.Integration;

public static class ZbXll
{
    // CP932 に無い文字を混ぜないため、ソース上は \uXXXX で書く。
    // "該当なし" = 該当なし
    private const string NotFound = "該当なし";

    // ------------------------------------------------------------------
    // Excel から Application.Run("ZbXllConvert", ...) で呼ばれる入口。
    // IsMacroType = true にしないと、シートを読み書きする C API を呼べない。
    // 戻り値は 1 本の文字列。VBA 側はこれを分解して計測欄へ入れる。
    //   OK|rows|hash|master|input|dict|convert|write|notFound
    //   NG|どの段階で|実際のエラー
    // 失敗しても別経路へ切り替えず、起きたことをそのまま返す。
    // ------------------------------------------------------------------
    [ExcelFunction(Name = "ZbXllConvert", IsMacroType = true)]
    public static object ZbXllConvert(
        string wbName, string masterSheet, string inputSheet, string outputSheet,
        double masterRowsD, double rowsD, double firstRowD)
    {
        string stage = "start";
        try
        {
            int masterRows = (int)masterRowsD;
            int rows = (int)rowsD;
            int firstRow = (int)firstRowD;
            // C API の行・列は 0 起点。シート上の 1 行目は 0。
            int r0 = firstRow - 1;

            var sw = new Stopwatch();

            stage = "master-read";
            sw.Restart();
            ExcelReference mRef = MakeRef(wbName, masterSheet, r0, r0 + masterRows - 1, 0, 2);
            object[,] master = (object[,])mRef.GetValue();
            sw.Stop(); long msMaster = sw.ElapsedMilliseconds;

            stage = "input-read";
            sw.Restart();
            ExcelReference iRef = MakeRef(wbName, inputSheet, r0, r0 + rows - 1, 0, 0);
            object[,] input = (object[,])iRef.GetValue();
            sw.Stop(); long msInput = sw.ElapsedMilliseconds;

            stage = "dict";
            sw.Restart();
            var dict = new Dictionary<string, string>(masterRows * 2);
            for (int i = 0; i < masterRows; i++)
            {
                string key = NormalizeZip(AsText(master[i, 0]));
                if (key.Length == 0) continue;
                if (!dict.ContainsKey(key)) dict[key] = AsText(master[i, 1]);
            }
            sw.Stop(); long msDict = sw.ElapsedMilliseconds;

            stage = "convert";
            sw.Restart();
            var output = new object[rows, 1];
            int notFound = 0;
            string hit;
            for (int i = 0; i < rows; i++)
            {
                string key = NormalizeZip(AsText(input[i, 0]));
                if (key.Length != 0 && dict.TryGetValue(key, out hit)) output[i, 0] = hit;
                else { output[i, 0] = NotFound; notFound++; }
            }
            sw.Stop(); long msConvert = sw.ElapsedMilliseconds;

            uint hash = FnvOf(output, rows);

            stage = "output-write";
            sw.Restart();
            ExcelReference oRef = MakeRef(wbName, outputSheet, r0, r0 + rows - 1, 0, 0);
            oRef.SetValue(output);
            sw.Stop(); long msWrite = sw.ElapsedMilliseconds;

            return string.Format("OK|{0}|{1:X8}|{2}|{3}|{4}|{5}|{6}|{7}",
                rows, hash, msMaster, msInput, msDict, msConvert, msWrite, notFound);
        }
        catch (Exception ex)
        {
            // 段階と実エラーを残す。ここを握りつぶすと「使えなかった理由」が消える。
            return "NG|" + stage + "|" + ex.GetType().Name + ": " + ex.Message;
        }
    }

    // 対象ブックの対象シートの矩形を指す参照を作る。
    // xlSheetId にブック名付きのシート名を渡して、そのシートの ID を得る。
    private static ExcelReference MakeRef(string wbName, string sheet,
                                          int rowFirst, int rowLast, int colFirst, int colLast)
    {
        object id = XlCall.Excel(XlCall.xlSheetId, "[" + wbName + "]" + sheet);
        var sheetRef = id as ExcelReference;
        if (sheetRef == null)
            throw new InvalidOperationException("シートが見つからない: [" + wbName + "]" + sheet);
        return new ExcelReference(rowFirst, rowLast, colFirst, colLast, sheetRef.SheetId);
    }

    private static string AsText(object v)
    {
        if (v == null) return "";
        string s = v as string;
        if (s != null) return s;
        if (v is double)
        {
            double d = (double)v;
            if (d == Math.Floor(d) && Math.Abs(d) < 1e15)
                return ((long)d).ToString(System.Globalization.CultureInfo.InvariantCulture);
            return d.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (v is ExcelEmpty || v is ExcelMissing || v is ExcelError) return "";
        return v.ToString();
    }

    // 数字以外を落として 7 桁に揃える。全角数字も拾う。
    private static string NormalizeZip(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var buf = new char[7];
        int n = 0;
        for (int i = 0; i < s.Length && n < 7; i++)
        {
            char c = s[i];
            if (c >= '0' && c <= '9') buf[n++] = c;
            else if (c >= '０' && c <= '９') buf[n++] = (char)(c - '０' + '0');
        }
        if (n != 7) return "";
        return new string(buf, 0, 7);
    }

    // ZipWorker.cs / PowerShell 側と同じ FNV-1a 32bit。
    // UTF-16 コード単位を 1 つずつ、行末に '\n' を混ぜる。
    private static uint FnvOf(object[,] a, int rows)
    {
        uint h = 2166136261;
        for (int i = 0; i < rows; i++)
        {
            string s = (string)a[i, 0];
            for (int j = 0; j < s.Length; j++)
            {
                h ^= s[j];
                h *= 16777619;
            }
            h ^= '\n';
            h *= 16777619;
        }
        return h;
    }
}
