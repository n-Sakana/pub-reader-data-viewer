using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public static class Rdv3ArchiveForm
{
    public static List<Rdv3DeletedRecord> Pick(Rdv3Form owner, Rdv3Data data, Rdv3WorkState work, IList<Rdv3DeletedRecord> records)
    {
        const int pageSize = 100;
        int page = 0;
        HashSet<int> selected = new HashSet<int>();
        string[] labels = new string[data.Columns.Count + 1]; labels[0] = work.Column;
        for (int i = 0; i < data.Columns.Count; i++)
        { string label = data.LabelOf(data.Columns[i].Ref); labels[i + 1] = label.Length == 0 ? data.Columns[i].Ref : label; }
        while (true)
        {
            StringBuilder json = new StringBuilder("{\"title\":").Append(Rdv3WebJson.Q(Rdv3Text.ArchiveTitle));
            json.Append(",\"hint\":").Append(Rdv3WebJson.Q(Rdv3Text.ArchiveHint));
            json.Append(",\"page\":").Append(page).Append(",\"pageSize\":").Append(pageSize).Append(",\"total\":").Append(records.Count);
            json.Append(",\"labels\":").Append(Rdv3WebJson.S(labels)).Append(",\"selected\":[");
            int n = 0; foreach (int index in selected) { if (n++ > 0) { json.Append(','); } json.Append(index); }
            json.Append("],\"rows\":[");
            for (int i = page * pageSize; i < Math.Min(records.Count, (page + 1) * pageSize); i++)
            {
                if (i > page * pageSize) { json.Append(','); }
                Rdv3DeletedRecord record = records[i];
                string[] cells = record.Line.Split('\t');
                string[] identity = new string[data.IdentityCols.Length];
                for (int k = 0; k < identity.Length; k++) { identity[k] = cells[data.IdentityCols[k]]; }
                string state = work.ByStored(record.State).Text;
                List<string> detail = new List<string> { state }; detail.AddRange(cells);
                json.Append("{\"index\":").Append(i).Append(",\"identity\":").Append(Rdv3WebJson.Q(string.Join(" / ", identity)));
                json.Append(",\"state\":").Append(Rdv3WebJson.Q(state)).Append(",\"deletedAt\":").Append(Rdv3WebJson.Q(record.DeletedAt));
                json.Append(",\"values\":").Append(Rdv3WebJson.S(detail.ToArray())).Append('}');
            }
            json.Append("]}");
            Rdv3Json result = owner.ShowModal("archive", json.ToString());
            Rdv3Json indices = result.Member("selected");
            if (indices != null && indices.IsArray)
            {
                selected.Clear();
                foreach (Rdv3Json index in indices.Items)
                { if (index.Kind == Rdv3Json.TNumber && index.Num >= 0 && index.Num < records.Count && index.Num == Math.Floor(index.Num)) { selected.Add((int)index.Num); } }
            }
            if (result.Has("page"))
            { page = Math.Max(0, Math.Min(Math.Max(0, (records.Count - 1) / pageSize), Rdv3Form.Number(result, "page", 0))); continue; }
            List<Rdv3DeletedRecord> requested = new List<Rdv3DeletedRecord>();
            if (Rdv3Form.Flag(result, "ok", false)) { foreach (int index in selected) { requested.Add(records[index]); } }
            return requested;
        }
    }
}
