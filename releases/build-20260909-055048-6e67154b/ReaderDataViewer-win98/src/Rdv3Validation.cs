using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

// Collect only across independent checks. A failed prerequisite stops its
// dependent work; no partly validated configuration may reach the application.
public sealed class Rdv3ValidationError : Exception
{
    public readonly string[] Errors;
    public readonly string[] Unchecked;
    public readonly string Stage;

    internal Rdv3ValidationError(List<string> errors, List<string> skipped, string stage, string next)
        : base("FAIL " + errors.Count.ToString(CultureInfo.InvariantCulture) + " errors; stopped at " + stage)
    {
        Errors = errors.ToArray(); Stage = stage;
        List<string> remaining = new List<string>(skipped);
        if (!string.IsNullOrEmpty(next)) { remaining.Add(next); }
        Unchecked = remaining.ToArray();
    }
}

internal sealed class Rdv3ValidationStop : Exception
{
    internal Rdv3ValidationStop(string remaining) : base(remaining) { }
}

public sealed class Rdv3Validation
{
    private readonly List<string> errors = new List<string>();
    private readonly List<string> skipped = new List<string>();
    public int Count { get { return errors.Count; } }

    public void Add(Exception error) { errors.Add(error.Message); }
    public void Skip(string scope) { if (!skipped.Contains(scope)) { skipped.Add(scope); } }

    public bool Check(string scope, Action check)
    {
        try { check(); return true; }
        catch (Rdv3ValidationStop stop) { Skip(stop.Message); return false; }
        catch (Rdv3LoadError error) { Add(error); }
        catch (Rdv3DataError error) { Add(error); }
        catch (IOException error) { Add(error); }
        catch (UnauthorizedAccessException error) { Add(error); }
        Skip("remaining checks within " + scope + " after the reported error");
        return false;
    }

    public void Guard(int before, string remaining)
    { if (Count > before) { throw new Rdv3ValidationStop(remaining); } }

    public void Finish(string stage, string next)
    { if (Count > 0) { throw new Rdv3ValidationError(errors, skipped, stage, next); } }
}
