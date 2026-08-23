// ============================================================================
// Rdv3Data.cs -- the "data" member of settings.json: what the CSVs are and how
// they become the ledger.
//
//   "data": {
//     "encoding": "utf-8",                       // or "shift_jis", any .NET name
//     "tables": { "A": { "file": "tableA.csv", "key": "key1" }, ... },
//     "spine": "B",                              // one ledger row per row of B
//     "joins": [ { "table": "A", "on": "key1" }, // B.key1 -> A.key
//                { "table": "C", "on": "key2" } ],
//     "ledger": { "search": "B.key1",            // what the search box looks up
//                 "columns": [ "B.key1", "B.key2", "A.a_code", ... ] }
//   }
//
// A table's "key" is the column that identifies one row of it (unique). The
// spine's key is the ledger row's identity (carry-over and the work state key
// on it). Every ledger column is "<table>.<column>" and the xlsx header is the
// column names in that order. There is no expression here: names only, and
// every name is checked -- against the definition when the file is read, and
// against the CSV headers before the window opens.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public sealed class Rdv3TableDef
{
    public string Id = "";
    public string File = "";
    public string Key = "";
    public int Ord;                          // position in Rdv3Data.Tables
    public string[] Head;                    // the CSV header, once read
}

public sealed class Rdv3JoinDef
{
    public string Table = "";                // the table joined to the spine
    public string On = "";                   // the spine column that holds its key
    public int TableOrd;
    public int OnField = -1;                 // index of On in the spine's header
}

public sealed class Rdv3ColumnRef
{
    public string Ref = "";                  // "A.a_code"
    public string Table = "";
    public string Column = "";
    public int TableOrd;
    public int Field = -1;                   // index of Column in the table's header
}

public sealed class Rdv3Data
{
    public string EncodingName = "utf-8";
    public Encoding Enc = new UTF8Encoding(false);
    public List<Rdv3TableDef> Tables = new List<Rdv3TableDef>();
    public string Spine = "";
    public int SpineOrd;
    public List<Rdv3JoinDef> Joins = new List<Rdv3JoinDef>();
    public string SearchRef = "";
    public int SearchCol = -1;               // index in Columns
    public int IdentityCol = -1;             // index in Columns of <spine>.<spine key>
    public List<Rdv3ColumnRef> Columns = new List<Rdv3ColumnRef>();

    public Rdv3TableDef SpineTable { get { return Tables[SpineOrd]; } }

    public Rdv3TableDef TableOf(string id)
    {
        for (int i = 0; i < Tables.Count; i++) { if (Tables[i].Id == id) { return Tables[i]; } }
        return null;
    }

    // the ledger column index of "A.a_code", or -1
    public int IndexOf(string reference)
    {
        for (int i = 0; i < Columns.Count; i++) { if (Columns[i].Ref == reference) { return i; } }
        return -1;
    }

    public string[] ColumnRefs
    {
        get
        {
            string[] r = new string[Columns.Count];
            for (int i = 0; i < r.Length; i++) { r[i] = Columns[i].Ref; }
            return r;
        }
    }

    // the xlsx header: the CSV column names, in ledger order
    public string[] Head
    {
        get
        {
            string[] h = new string[Columns.Count];
            for (int i = 0; i < h.Length; i++) { h[i] = Columns[i].Column; }
            return h;
        }
    }

    // ---- reading the definition -------------------------------------------------
    public static Rdv3Data Read(Rdv3Json o)
    {
        Rdv3Data d = new Rdv3Data();
        o.Only("encoding", "tables", "spine", "joins", "ledger");

        d.EncodingName = o.StrOr("encoding", "utf-8");
        try
        {
            d.Enc = (string.Equals(d.EncodingName, "utf-8", StringComparison.OrdinalIgnoreCase))
                ? (Encoding)new UTF8Encoding(false) : Encoding.GetEncoding(d.EncodingName);
        }
        catch (Exception ex) { throw o.Member("encoding").Fail("is not an encoding this machine knows (" + ex.Message + ")"); }

        Rdv3Json tables = o.Obj("tables", true);
        if (tables.Order.Count == 0) { throw tables.Fail("names no table"); }
        for (int i = 0; i < tables.Order.Count; i++)
        {
            string id = tables.Order[i];
            Rdv3Json to = tables.Member(id);
            if (id.Trim().Length == 0 || id.IndexOf('.') >= 0 || id.IndexOf(' ') >= 0)
            {
                throw to.Fail("a table id must be a word without . or spaces");
            }
            if (to.Kind != Rdv3Json.TObject) { throw to.Fail("must be an object { file, key }"); }
            to.Only("file", "key");
            Rdv3TableDef t = new Rdv3TableDef();
            t.Id = id;
            t.File = to.Need("file");
            t.Key = to.Need("key");
            t.Ord = d.Tables.Count;
            d.Tables.Add(t);
        }

        d.Spine = o.Need("spine");
        Rdv3TableDef spine = d.TableOf(d.Spine);
        if (spine == null) { throw o.Member("spine").Fail(d.Spine + " is not one of the tables"); }
        d.SpineOrd = spine.Ord;

        List<Rdv3Json> joins = o.Objs("joins", false);
        for (int i = 0; i < joins.Count; i++)
        {
            Rdv3Json jo = joins[i];
            jo.Only("table", "on");
            Rdv3JoinDef j = new Rdv3JoinDef();
            j.Table = jo.Need("table");
            j.On = jo.Need("on");
            Rdv3TableDef t = d.TableOf(j.Table);
            if (t == null) { throw jo.Member("table").Fail(j.Table + " is not one of the tables"); }
            if (t.Ord == d.SpineOrd) { throw jo.Member("table").Fail("the spine cannot be joined to itself"); }
            for (int k = 0; k < d.Joins.Count; k++)
            {
                if (d.Joins[k].Table == j.Table) { throw jo.Fail(j.Table + " is joined twice"); }
            }
            j.TableOrd = t.Ord;
            d.Joins.Add(j);
        }
        // a table that is neither the spine nor joined would never contribute a
        // value, so naming it is a mistake
        for (int i = 0; i < d.Tables.Count; i++)
        {
            if (i == d.SpineOrd) { continue; }
            bool joined = false;
            for (int k = 0; k < d.Joins.Count; k++) { if (d.Joins[k].TableOrd == i) { joined = true; } }
            if (!joined) { throw tables.Member(d.Tables[i].Id).Fail("is neither the spine nor joined to it"); }
        }

        Rdv3Json ledger = o.Obj("ledger", true);
        ledger.Only("search", "columns");
        string[] cols = ledger.Strs("columns", true);
        if (cols.Length == 0) { throw ledger.Member("columns").Fail("names no column"); }
        Rdv3Json colsNode = ledger.Member("columns");
        for (int i = 0; i < cols.Length; i++)
        {
            Rdv3ColumnRef c = ParseRef(d, cols[i], colsNode.At(i));
            if (d.IndexOf(c.Ref) >= 0) { throw colsNode.At(i).Fail(c.Ref + " is listed twice"); }
            d.Columns.Add(c);
        }
        string identity = spine.Id + "." + spine.Key;
        d.IdentityCol = d.IndexOf(identity);
        if (d.IdentityCol < 0) { throw colsNode.Fail("must include the spine's key " + identity + " (the row identity)"); }
        d.SearchRef = ledger.Need("search");
        ParseRef(d, d.SearchRef, ledger.Member("search"));
        d.SearchCol = d.IndexOf(d.SearchRef);
        if (d.SearchCol < 0) { throw ledger.Member("search").Fail(d.SearchRef + " must be one of the ledger columns"); }
        return d;
    }

    private static Rdv3ColumnRef ParseRef(Rdv3Data d, string s, Rdv3Json at)
    {
        Rdv3ColumnRef c = new Rdv3ColumnRef();
        c.Ref = (s == null) ? "" : s.Trim();
        int dot = c.Ref.IndexOf('.');
        if (dot <= 0 || dot == c.Ref.Length - 1) { throw at.Fail("a column is written <table>.<column>, not " + c.Ref); }
        c.Table = c.Ref.Substring(0, dot);
        c.Column = c.Ref.Substring(dot + 1);
        Rdv3TableDef t = d.TableOf(c.Table);
        if (t == null) { throw at.Fail(c.Table + " is not one of the tables"); }
        c.TableOrd = t.Ord;
        return c;
    }

    // ---- binding to the CSV headers ------------------------------------------------
    // heads[t] is the header row of Tables[t]. Every name the definition uses
    // must be there; the first that is not stops the app with the file and
    // the name.
    public void Bind(string[][] heads)
    {
        for (int t = 0; t < Tables.Count; t++)
        {
            Tables[t].Head = heads[t];
            if (FieldOf(heads[t], Tables[t].Key) < 0) { throw Missing(Tables[t], Tables[t].Key, "tables." + Tables[t].Id + ".key"); }
        }
        string[] spineHead = heads[SpineOrd];
        for (int i = 0; i < Joins.Count; i++)
        {
            Joins[i].OnField = FieldOf(spineHead, Joins[i].On);
            if (Joins[i].OnField < 0) { throw Missing(SpineTable, Joins[i].On, "joins[" + i.ToString(CultureInfo.InvariantCulture) + "].on"); }
        }
        for (int i = 0; i < Columns.Count; i++)
        {
            Rdv3ColumnRef c = Columns[i];
            c.Field = FieldOf(heads[c.TableOrd], c.Column);
            if (c.Field < 0) { throw Missing(Tables[c.TableOrd], c.Column, "ledger.columns"); }
        }
    }

    private static Rdv3DataError Missing(Rdv3TableDef t, string column, string where)
    {
        return new Rdv3DataError(Rdv3Text.DataNoColumn.Replace("{file}", t.File).Replace("{row}", "1")
            .Replace("{name}", column) + " (data." + where + ")");
    }

    private static int FieldOf(string[] head, string name)
    {
        if (head == null) { return -1; }
        for (int i = 0; i < head.Length; i++) { if (head[i] == name) { return i; } }
        return -1;
    }

    // one line for the log
    public string Describe()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("tables=").Append(Tables.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" spine=").Append(Spine);
        sb.Append(" joins=");
        for (int i = 0; i < Joins.Count; i++) { if (i > 0) { sb.Append(','); } sb.Append(Joins[i].Table).Append("(on ").Append(Joins[i].On).Append(')'); }
        sb.Append(" columns=").Append(Columns.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" identity=").Append(Columns[IdentityCol].Ref);
        sb.Append(" search=").Append(SearchRef);
        sb.Append(" encoding=").Append(EncodingName);
        return sb.ToString();
    }
}
