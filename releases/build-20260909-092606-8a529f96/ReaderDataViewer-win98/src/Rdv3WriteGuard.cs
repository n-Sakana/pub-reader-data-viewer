using System;

internal sealed class Rdv3WriteGuard
{
    private enum Phase { Idle, WaitingForLock, Writing }
    private readonly object gate = new object();
    private Phase phase;
    private bool closing;
    private bool closeRequested;

    public bool Pending { get { lock (gate) { return phase != Phase.Idle; } } }
    public bool Closing { get { lock (gate) { return closing; } } }

    public void Begin(bool shared)
    {
        lock (gate)
        {
            if (closing) { throw new OperationCanceledException("application is closing"); }
            if (phase != Phase.Idle) { throw new InvalidOperationException("a write is already pending"); }
            phase = shared ? Phase.WaitingForLock : Phase.Writing;
            closeRequested = false;
        }
    }

    public bool TryEnterSharedWrite()
    {
        lock (gate)
        {
            // Closing and acquiring a lease must agree on who won before any
            // shared read/write starts. A lease acquired after close is released.
            if (closing) { return false; }
            if (phase != Phase.WaitingForLock) { throw new InvalidOperationException("no shared write is waiting"); }
            phase = Phase.Writing;
            return true;
        }
    }

    public bool TryClose()
    {
        lock (gate)
        {
            // A managed write cannot be stopped safely. Its timeout is only a
            // warning; only Complete may release this hold after the job returns.
            if (phase == Phase.Writing) { closeRequested = true; return false; }
            closing = true;
            return true;
        }
    }

    public bool Complete(out bool requestedClose)
    {
        lock (gate)
        {
            requestedClose = closeRequested;
            if (phase == Phase.Idle) { return false; }
            phase = Phase.Idle;
            closeRequested = false;
            return true;
        }
    }
}
