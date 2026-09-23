#nullable enable
using System;
using System.Diagnostics;
using System.Threading;

internal sealed class ProductionOperationProbe : IDisposable
{
    private readonly Process _process = Process.GetCurrentProcess();
    private readonly ManualResetEvent _stop = new(false);
    private readonly Stopwatch _clock = new();
    private Thread? _sampler;
    private TimeSpan _cpu;
    private long _allocated;
    private int _gen0;
    private int _gen1;
    private int _gen2;
    private long _peakWorkingSet;
    private long _peakPrivateBytes;

    internal void Start()
    {
        _stop.Reset();
        _process.Refresh();
        _peakWorkingSet = _process.WorkingSet64;
        _peakPrivateBytes = _process.PrivateMemorySize64;
        _cpu = _process.TotalProcessorTime;
        _gen0 = GC.CollectionCount(0);
        _gen1 = GC.CollectionCount(1);
        _gen2 = GC.CollectionCount(2);
        _allocated = GC.GetTotalAllocatedBytes(true);
        _clock.Restart();
        _sampler = new Thread(Sample) { IsBackground = true, Name = "Beacon benchmark memory sampler" };
        _sampler.Start();
    }

    internal Metrics Stop()
    {
        double elapsed = _clock.Elapsed.TotalMilliseconds;
        long allocated = GC.GetTotalAllocatedBytes(true) - _allocated;
        _stop.Set();
        _sampler?.Join();
        _process.Refresh();
        _peakWorkingSet = Math.Max(_peakWorkingSet, _process.WorkingSet64);
        _peakPrivateBytes = Math.Max(_peakPrivateBytes, _process.PrivateMemorySize64);
        return new Metrics
        {
            ElapsedMs = elapsed,
            CpuMs = (_process.TotalProcessorTime - _cpu).TotalMilliseconds,
            AllocatedBytes = allocated,
            Gen0 = GC.CollectionCount(0) - _gen0,
            Gen1 = GC.CollectionCount(1) - _gen1,
            Gen2 = GC.CollectionCount(2) - _gen2,
            PeakWorkingSetBytes = _peakWorkingSet,
            PeakPrivateBytes = _peakPrivateBytes
        };
    }

    private void Sample()
    {
        while (!_stop.WaitOne(20))
        {
            _process.Refresh();
            _peakWorkingSet = Math.Max(_peakWorkingSet, _process.WorkingSet64);
            _peakPrivateBytes = Math.Max(_peakPrivateBytes, _process.PrivateMemorySize64);
        }
    }

    public void Dispose()
    {
        _stop.Set();
        _sampler?.Join();
        _stop.Dispose();
        _process.Dispose();
    }

    internal sealed class Metrics
    {
        public double ElapsedMs { get; set; }
        public double CpuMs { get; set; }
        public long AllocatedBytes { get; set; }
        public int Gen0 { get; set; }
        public int Gen1 { get; set; }
        public int Gen2 { get; set; }
        public long PeakWorkingSetBytes { get; set; }
        public long PeakPrivateBytes { get; set; }
    }
}
