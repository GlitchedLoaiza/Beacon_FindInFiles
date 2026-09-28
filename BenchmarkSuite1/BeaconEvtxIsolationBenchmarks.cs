#nullable enable
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Beacon.Beacon;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;

// Diagnostic, single-invocation probes, not throughput benchmarks. Each case runs in
// its own BenchmarkDotNet process because a native message-format call cannot be
// interrupted by a CancellationToken. A timeout is a lower bound, not a scan time.
[SimpleJob(RunStrategy.ColdStart, launchCount: 1, warmupCount: 0, iterationCount: 1, invocationCount: 1)]
public class BeaconEvtxIsolationBenchmarks
{
    private const int DeadlineSeconds = 10;
    private static readonly string Root = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads",
        "TSS_PSVMJMODONNELL_260924-113511_Log-ITN_Report");
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private static string EvidenceDirectory
    {
        get
        {
            for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory != null; directory = directory.Parent)
            {
                if (File.Exists(Path.Combine(directory.FullName, "Beacon_FindInFiles.slnx")))
                    return Path.Combine(directory.FullName, "BenchmarkDotNet.Artifacts", "evtx-isolation-Report20260924-1344");
            }
            throw new DirectoryNotFoundException("Cannot locate the Beacon workspace for diagnostic evidence.");
        }
    }

    public IEnumerable<object[]> FileCases()
    {
        foreach (var path in Directory.EnumerateFiles(Root, "*.evtx", SearchOption.AllDirectories).OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            yield return new object[] { Path.GetRelativePath(Root, path) };
    }

    public IEnumerable<object[]> ProviderCases()
    {
        var directory = EvidenceDirectory;
        var reports = Directory.Exists(directory)
            ? Directory.EnumerateFiles(directory, "file-*.json")
                .Select(path => JsonSerializer.Deserialize<ProbeEvidence>(File.ReadAllText(path))!)
                .Where(report => report.Root == Root && report.Status != "Started" &&
                    File.Exists(Path.Combine(Root, report.RelativePath)) &&
                    new FileInfo(Path.Combine(Root, report.RelativePath)).Length == report.FileBytes)
                .GroupBy(report => report.RelativePath, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.OrderByDescending(report => report.StartedUtc).First()).ToArray()
            : Array.Empty<ProbeEvidence>();
        var slow = reports.Where(report => report.Status == "TimedOut").ToArray();
        var candidates = (slow.Length > 0 ? slow : reports.OrderByDescending(report => report.ElapsedMs).Take(1))
            .Select(report => report.RelativePath).ToArray();
        if (candidates.Length == 0)
            candidates = FileCases().Take(1).Select(item => (string)item[0]).ToArray();
        foreach (var relativePath in candidates.OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
        {
            var providers = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var reader = new EventLogReader(Path.Combine(Root, relativePath), PathType.FilePath))
            {
                while (true)
                {
                    using var record = reader.ReadEvent();
                    if (record == null) break;
                    if (!string.IsNullOrWhiteSpace(record.ProviderName)) providers.Add(record.ProviderName);
                }
            }
            foreach (var provider in providers)
                yield return new object[] { relativePath, provider };
        }
    }

    [Benchmark]
    [ArgumentsSource(nameof(FileCases))]
    public Task<int> ProbeFile(string relativePath) => Probe(relativePath, "", "file");

    [Benchmark]
    [ArgumentsSource(nameof(ProviderCases))]
    public Task<int> ProbeProvider(string relativePath, string provider) => Probe(relativePath, provider, "provider");

    private static async Task<int> Probe(string relativePath, string provider, string phase)
    {
        var path = Path.Combine(Root, relativePath);
        var settings = new BeaconSettings { EvtxProvider = provider };
        var evidence = new ProbeEvidence
        {
            Root = Root, RelativePath = relativePath, ProviderFilter = provider,
            FileBytes = new FileInfo(path).Length, StartedUtc = DateTimeOffset.UtcNow,
            Phase = phase, ProcessId = Environment.ProcessId, Status = "Started"
        };
        var output = EvidenceDirectory;
        Directory.CreateDirectory(output);
        var outputPath = Path.Combine(output, $"{phase}-{Environment.ProcessId}-{Guid.NewGuid():N}.json");
        File.WriteAllText(outputPath, JsonSerializer.Serialize(evidence, JsonOptions));
        Console.WriteLine($"EVTX_PROBE_START file={relativePath}; provider={provider}; deadline={DeadlineSeconds}s");
        var issues = new ConcurrentQueue<string>();
        int issueCount = 0;
        using var cancellation = new CancellationTokenSource();
        var token = cancellation.Token;
        using var process = Process.GetCurrentProcess();
        var cpuBefore = process.TotalProcessorTime;
        var watch = Stopwatch.StartNew();
        var work = Task.Factory.StartNew(() =>
        {
            var result = new StructuredSearchResult<EventRecordSummary>();
            var service = new EvtxSearchService(new SearchQuery("enroll", SearchMode.PlainText, false), settings);
            service.Collect(path, result, token, (error, stage, severity) =>
            {
                if (Interlocked.Increment(ref issueCount) <= 8)
                    issues.Enqueue($"{stage}: {error.GetType().Name}: {error.Message}");
            });
            return result;
        }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
        try
        {
            var result = await work.WaitAsync(TimeSpan.FromSeconds(DeadlineSeconds)).ConfigureAwait(false);
            evidence.Status = "Completed";
            evidence.Matches = result.Records.Count;
            evidence.PartialReason = result.PartialReason;
        }
        catch (TimeoutException)
        {
            evidence.Status = "TimedOut";
            cancellation.Cancel();
            await Task.WhenAny(work, Task.Delay(250)).ConfigureAwait(false);
            evidence.ReturnedWithinCancellationGrace = work.IsCompleted;
            _ = work.ContinueWith(completed => { if (completed.IsFaulted) _ = completed.Exception; }, TaskScheduler.Default);
        }
        catch (Exception error)
        {
            evidence.Status = "Error";
            evidence.Error = error.GetType().Name + ": " + error.Message;
        }
        finally
        {
            watch.Stop();
            process.Refresh();
            evidence.ElapsedMs = watch.Elapsed.TotalMilliseconds;
            evidence.ProcessCpuMs = (process.TotalProcessorTime - cpuBefore).TotalMilliseconds;
            evidence.DiagnosticCount = Volatile.Read(ref issueCount);
            evidence.Diagnostics = issues.ToArray();
            File.WriteAllText(outputPath, JsonSerializer.Serialize(evidence, JsonOptions));
            Console.WriteLine("EVTX_PROBE_RESULT " + JsonSerializer.Serialize(evidence));
        }
        return evidence.Status == "Completed" ? evidence.Matches : -1;
    }

    public sealed class ProbeEvidence
    {
        public string Root { get; set; } = "";
        public string RelativePath { get; set; } = "";
        public string ProviderFilter { get; set; } = "";
        public long FileBytes { get; set; }
        public DateTimeOffset StartedUtc { get; set; }
        public string Phase { get; set; } = "";
        public int ProcessId { get; set; }
        public string Query { get; set; } = "enroll";
        public int DeadlineSeconds { get; set; } = BeaconEvtxIsolationBenchmarks.DeadlineSeconds;
        public string Status { get; set; } = "";
        public double ElapsedMs { get; set; }
        public double ProcessCpuMs { get; set; }
        public int Matches { get; set; }
        public string PartialReason { get; set; } = "";
        public bool? ReturnedWithinCancellationGrace { get; set; }
        public int DiagnosticCount { get; set; }
        public string[] Diagnostics { get; set; } = Array.Empty<string>();
        public string? Error { get; set; }
    }
}
