using System;
using System.Collections.Generic;

// The HTML owns the layout. This reader accepts only values and behavior;
// it never reconstructs a GUI tree from the operator's settings.
public static class Rdv3FixedScreen
{
    public static readonly string[] MainFields = { "userId", "userName", "userCategory",
        "applicationDate", "applicationNumber", "cardNumber", "applicantName",
        "birthDate", "qualification", "office", "remarks", "plan" };
    public static readonly string[] CandidateFields = { "applicationNumber", "cardNumber",
        "applicant", "applicationDate", "payment" };

    public static Rdv3Screen Read(Rdv3Json root)
    {
        int before = root.ErrorCount;
        root.Only("bindings", "judgments", "workState", "export", "candidates", "actions");
        Rdv3Screen screen = new Rdv3Screen();
        screen.Bindings = new Dictionary<string, Rdv3Bind>(StringComparer.Ordinal);
        screen.StartWidth = 750;
        screen.StartHeight = 690;
        root.Check(delegate {
            Rdv3Json bindings = root.Obj("bindings", true);
            bindings.Only(MainFields);
            foreach (string name in MainFields)
            {
                bindings.Check(delegate { screen.Bindings.Add(name, Rdv3Bind.Read(bindings.Obj(name, true))); });
            }
        });
        foreach (string name in new string[] { "searchKey", "pendingCount", "appState", "ledgerRows", "clock" })
        {
            screen.Bindings.Add(name, new Rdv3Bind { State = name, Empty = "" });
        }
        root.Check(delegate {
            Rdv3Json judgments = root.Obj("judgments", true);
            judgments.Only("paymentStatus");
            screen.Judgments.Add("paymentStatus", Rdv3Judgment.Read("paymentStatus", judgments.Obj("paymentStatus", true)));
        });
        root.Check(delegate { screen.Work = Rdv3WorkState.Read(root.Obj("workState", true)); });
        root.Check(delegate {
            Rdv3Json export = root.Obj("export", true);
            export.Only("defaultFields");
            screen.ExportDefaultFields = export.Strs("defaultFields", true);
            if (screen.ExportDefaultFields.Length == 0) { throw export.Fail("names no field"); }
        });
        root.Check(delegate {
            Rdv3Json actions = root.Obj("actions", true);
            actions.Only("updateRecords", "deleteRecords");
            foreach (string action in new string[] { "updateRecords", "deleteRecords" })
            {
                screen.FixedActions.Add(action, actions.Need(action));
            }
        });
        root.Check(delegate {
            Rdv3Json candidates = root.Obj("candidates", true);
            candidates.Only(CandidateFields);
            screen.Candidates = new Rdv3CandidatesDef();
            screen.Candidates.Columns.Add(new Rdv3ColumnDef { Value = new Rdv3Bind { State = "rowNumber" } });
            foreach (string name in CandidateFields)
            {
                Rdv3Json entry = candidates.Obj(name, true);
                entry.Only("value", "looks");
                Rdv3ColumnDef column = new Rdv3ColumnDef { Value = Rdv3Bind.Read(entry.Obj("value", true)) };
                if (name == "payment") { column.Render = "tag"; }
                Rdv3Json looks = entry.Obj("looks", false);
                if (looks != null)
                {
                    if (name != "payment") { throw entry.Fail("only payment accepts looks"); }
                    foreach (string key in looks.Order)
                    {
                        column.Looks.Add(key, looks.Word(key, "neutral", Rdv3WorkState.Looks));
                    }
                }
                screen.Candidates.Columns.Add(column);
            }
            screen.Candidates.Columns.Add(new Rdv3ColumnDef {
                Value = new Rdv3Bind { State = "workStateShort" }, Render = "tag" });
        });
        root.Guard(before, "fixed screen values (incomplete definition)");
        return screen;
    }
}
