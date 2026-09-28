#nullable enable
using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Beacon.Beacon;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;

// Compare verified copies only. Never move or remove the source report's metadata.
// The bounded wait does not interrupt native formatting; separate benchmark processes
// contain any worker that is still inside Windows after cancellation is requested.
[SimpleJob(RunStrategy.ColdStart, launchCount: 1, warmupCount: 0, iterationCount: 1, invocationCount: 1)]
public class BeaconEvtxLocaleControlBenchmarks
{
    private const int DeadlineSeconds = 30;
    private string _evidence = "";
    private string _copyDirectory = "";
    private string _sourcePath = "";
    private string _copyPath = "";
    private string _metadataPath = "";
    private string _provider = "";
    private string _hash = "";
    private Task? _worker;

    [Params("ShellCore", "ApplicationCritical")]
    public string Specimen { get; set; } = "";

    [Params(false, true)]
    public bool WithLocaleMetadata { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var workspace = new DirectoryInfo(AppContext.BaseDirectory);
        while (workspace != null && !File.Exists(Path.Combine(workspace.FullName, "Beacon_FindInFiles.slnx")))
            workspace = workspace.Parent;
        if (workspace == null) throw new DirectoryNotFoundException("Cannot locate the Beacon workspace.");
        _evidence = Path.Combine(workspace.FullName, "BenchmarkDotNet.Artifacts", "evtx-isolation-Report20260924-1344");
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads",
            "TSS_PSVMJMODONNELL_260924-113511_Log-ITN_Report");
        if (Specimen == "ShellCore")
        {
            _sourcePath = Path.Combine(root, "Intune_Report-2026-09-24.1135.11", "09_EventLogs",
                "PSVMJMODONNELL_evt_Microsoft-Windows-Shell-Core-Operational.evtx");
            _provider = "Microsoft-Windows-Shell-Core";
        }
        else
        {
            _sourcePath = Path.Combine(root, "PSVMJMODONNELL_evt_Application.evtx");
            _provider = "Critical";
        }
        _metadataPath = Path.Combine(Path.GetDirectoryName(_sourcePath)!, "LocaleMetaData",
            Path.GetFileNameWithoutExtension(_sourcePath) + "_1033.MTA");
        if (!File.Exists(_metadataPath)) throw new FileNotFoundException("Missing control metadata.", _metadataPath);
        _copyDirectory = Path.Combine(_evidence, "copies", $"{Specimen}-{WithLocaleMetadata}-{Environment.ProcessId}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_copyDirectory);
        _copyPath = Path.Combine(_copyDirectory, Path.GetFileName(_sourcePath));
        File.Copy(_sourcePath, _copyPath);
        using (var original = File.OpenRead(_sourcePath))
        using (var copy = File.OpenRead(_copyPath))
        {
            var hash = SHA256.HashData(original);
            if (!hash.SequenceEqual(SHA256.HashData(copy))) throw new InvalidDataException("The EVTX copy differs from its source.");
            _hash = Convert.ToHexString(hash);
        }
        if (WithLocaleMetadata)
        {
            var metadataDirectory = Path.Combine(_copyDirectory, "LocaleMetaData");
            Directory.CreateDirectory(metadataDirectory);
            File.Copy(_metadataPath, Path.Combine(metadataDirectory, Path.GetFileName(_metadataPath)));
        }
    }

    [Benchmark]
    public async Task<int> SearchWithControlledMetadata()
    {
        var startedUtc = DateTimeOffset.UtcNow;
        var issues = new ConcurrentQueue<string>();
        using var cancellation = new CancellationTokenSource();
        var token = cancellation.Token;
        using var process = Process.GetCurrentProcess();
        var cpuBefore = process.TotalProcessorTime;
        var watch = Stopwatch.StartNew();
        var work = Task.Factory.StartNew(() =>
        {
            var settings = new BeaconSettings { EvtxProvider = _provider };
            var result = new StructuredSearchResult<EventRecordSummary>();
            new EvtxSearchService(new SearchQuery("enroll", SearchMode.PlainText, false), settings)
                .Collect(_copyPath, result, token, (error, stage, severity) => issues.Enqueue(stage + ": " + error.Message));
            return result;
        }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
        _worker = work;
        string status;
        int? matches = null;
        string? partialReason = null;
        string? failure = null;
        bool? returnedWithinCancellationGrace = null;
        try
        {
            var result = await work.WaitAsync(TimeSpan.FromSeconds(DeadlineSeconds)).ConfigureAwait(false);
            status = "Completed";
            matches = result.Records.Count;
            partialReason = result.PartialReason;
        }
        catch (TimeoutException)
        {
            status = "TimedOut";
            cancellation.Cancel();
            await Task.WhenAny(work, Task.Delay(250)).ConfigureAwait(false);
            returnedWithinCancellationGrace = work.IsCompleted;
            _ = work.ContinueWith(completed => { if (completed.IsFaulted) _ = completed.Exception; }, TaskScheduler.Default);
        }
        catch (Exception error)
        {
            status = "Error";
            failure = error.GetType().Name + ": " + error.Message;
        }
        watch.Stop();
        process.Refresh();
        var evidence = new
        {
            Specimen, WithLocaleMetadata, SourcePath = _sourcePath, CopyPath = _copyPath,
            MetadataPath = _metadataPath, Provider = _provider, Sha256 = _hash,
            Query = "enroll", StartedUtc = startedUtc, ProcessId = Environment.ProcessId,
            DeadlineSeconds, Status = status, ElapsedMs = watch.Elapsed.TotalMilliseconds,
            ProcessCpuMs = (process.TotalProcessorTime - cpuBefore).TotalMilliseconds,
            Matches = matches, PartialReason = partialReason,
            ReturnedWithinCancellationGrace = returnedWithinCancellationGrace,
            Error = failure, Diagnostics = issues.ToArray()
        };
        File.WriteAllText(Path.Combine(_evidence, $"locale-{Specimen}-{WithLocaleMetadata}-{Environment.ProcessId}-{Guid.NewGuid():N}.json"),
            JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine("EVTX_LOCALE_CONTROL " + JsonSerializer.Serialize(evidence));
        return matches ?? -1;
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if ((_worker == null || _worker.IsCompleted) && Directory.Exists(_copyDirectory))
        {
            try
            {
                Directory.Delete(_copyDirectory, true);
            }
            catch (IOException error)
            {
                // Windows can retain a locale-metadata handle until the probe process exits.
                Console.WriteLine($"EVTX_COPY_RETAINED {_copyDirectory}: {error.Message}");
            }
        }
    }
}
