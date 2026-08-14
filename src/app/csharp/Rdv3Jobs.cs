// ============================================================================
// Rdv3Jobs.cs -- the worker: one background thread, a job queue, run IDs.
//
// Everything heavy (CSV read, indexing, joins, ledger comparison, carry-over,
// xlsx persistence, search) runs here. The UI thread only ever posts jobs and
// receives results marshalled back to it. There is no child process in the C#
// build -- the worker is a thread, and its owning PID is this process, which
// is logged at boot so the ownership is explicit.
//
// Every job carries a run ID. The app tracks which run ID it is currently
// waiting for; a result that arrives with any other ID is stale (a timed-out
// job that finished late, a superseded search) and is discarded with a log
// line, never applied. A managed thread cannot be killed safely, so a timeout
// does not abort the job -- it reports the timeout and abandons the ID.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Threading;

public sealed class Rdv3Job
{
    public string RunId = "";
    public string Kind = "";
    public Action Work;
    public double TimeoutMs;
    public long StartedAt;
    public bool TimeoutReported;
}

public sealed class Rdv3Worker
{
    private readonly Queue<Rdv3Job> queue = new Queue<Rdv3Job>();
    private readonly object gate = new object();
    private Thread thread;
    private volatile bool stopping;
    private Rdv3Job current;

    // called on the worker thread when a job throws; the app routes it to the
    // UI and the log. Never swallowed here.
    public Action<Rdv3Job, Exception> OnError;

    public void Start()
    {
        if (thread != null && thread.IsAlive) { return; }
        stopping = false;
        thread = new Thread(Loop);
        thread.IsBackground = true;
        thread.Name = "rdv-worker";
        thread.Start();
    }

    public void Stop()
    {
        stopping = true;
        lock (gate) { Monitor.PulseAll(gate); }
        Thread t = thread;
        if (t != null) { t.Join(2000); }
    }

    public void Post(Rdv3Job job)
    {
        lock (gate)
        {
            queue.Enqueue(job);
            Monitor.Pulse(gate);
        }
    }

    public int Pending()
    {
        lock (gate) { return queue.Count; }
    }

    // polled by a UI timer: the job that has been running past its deadline,
    // reported once. The job itself keeps running; its run ID is abandoned by
    // the caller so a late result is discarded.
    public Rdv3Job TakeOverdue()
    {
        lock (gate)
        {
            Rdv3Job c = current;
            if (c == null || c.TimeoutMs <= 0 || c.TimeoutReported) { return null; }
            if (Rdv3Clock.MsSince(c.StartedAt) < c.TimeoutMs) { return null; }
            c.TimeoutReported = true;
            return c;
        }
    }

    private void Loop()
    {
        while (!stopping)
        {
            Rdv3Job job = null;
            lock (gate)
            {
                while (queue.Count == 0 && !stopping) { Monitor.Wait(gate, 500); }
                if (stopping) { return; }
                job = queue.Dequeue();
                job.StartedAt = Rdv3Clock.Now();
                current = job;
            }
            try
            {
                job.Work();
            }
            catch (Exception ex)
            {
                Action<Rdv3Job, Exception> h = OnError;
                if (h != null)
                {
                    try { h(job, ex); }
                    catch (Exception) { }
                }
            }
            finally
            {
                lock (gate) { current = null; }
            }
        }
    }
}
