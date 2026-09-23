#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Beacon.Beacon;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;

[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 5, invocationCount: 1)]
public class BeaconProductionScalingBenchmarks
{
    private static readonly string[] ArchiveExtensions = { ".zip" };
    private readonly List<SourceSearchResult> _results = new();
    private readonly List<string> _issues = new();
    private ProductionSearchCorpus _corpus = null!;
    private ProductionOperationProbe _probe = null!;
    private SearchQuery _query = null!;
    private BeaconSettings _settings = null!;
    private string _source = "";
    private string _expectedDigest = "";
    private string _evidence = "";
    private int _processed;
    private int _matching;
    private bool _limit;
    private int _activeJobs;
    private int _operation;

    [Params(ProductionSearchWorkload.PlainFiles, ProductionSearchWorkload.CompressedZips,
        ProductionSearchWorkload.SmallFiles, ProductionSearchWorkload.DenseResults,
        ProductionSearchWorkload.WholeDocuments, ProductionSearchWorkload.SingleArchive)]
    public ProductionSearchWorkload Workload { get; set; }

    // -1 measures the original sequential service; 0 exercises normal automatic policy.
    [ParamsSource(nameof(WorkerCounts))]
    public int Workers { get; set; }

    public IEnumerable<int> WorkerCounts => new[] { -1, 0, 1, 2, 4, 8, 14, Environment.ProcessorCount }
        .Where(count => count <= Environment.ProcessorCount).Distinct().OrderBy(count => count);

    [GlobalSetup]
    public void Setup()
    {
        _corpus = ProductionSearchCorpus.Open();
        _source = _corpus.Source(Workload);
        _query = new SearchQuery(ProductionSearchCorpus.Marker, SearchMode.PlainText, false);
        _settings = new BeaconSettings
        {
            IncludedExtensions = ".log;.json",
            SearchFileContents = true,
            SearchFileNames = false,
            SearchFullPaths = false,
            StopAfterFirstMatchPerFile = false,
            MaximumStructuredMatches = 100000,
            MaximumTotalResults = 100000,
            MaximumFileSizeMb = 1024,
            MaximumArchiveEntrySizeMb = 1024,
            MaximumArchiveExpandedSizeMb = 2048,
            MaximumCompressionRatio = 100
        };
        _evidence = Environment.GetEnvironmentVariable("BEACON_SCALING_EVIDENCE") ??
            Path.Combine(_corpus.Root, "measurements");
        Directory.CreateDirectory(_evidence);
        _probe = new ProductionOperationProbe();
        using (var baseline = new SourceSearchService(_query, _settings, ArchiveExtensions,
            publish: Capture, report: Report))
        {
            baseline.RunAsync(_source, false, CancellationToken.None).GetAwaiter().GetResult();
            _processed = baseline.FilesProcessed;
            _matching = baseline.MatchingFiles;
            _limit = baseline.ResultLimitReached;
        }
        ValidateCounts();
        _expectedDigest = Digest();
        _results.Clear();
        _issues.Clear();
        Console.WriteLine($"PRODUCTION_CORPUS workload={Workload}; workers={Workers}; logicalProcessors={Environment.ProcessorCount}; " +
            $"files={_corpus.ExpectedFiles(Workload)}; expandedBytes={_corpus.ExpandedBytes(Workload)}; " +
            $"compressedBytesPerArchive={_corpus.Manifest.CompressedBytesPerArchive}; repeated/cache-influenced reads, not a cold-cache test.");
    }

    [IterationSetup]
    public void BeginOperation()
    {
        _results.Clear();
        _issues.Clear();
        _processed = 0;
        _matching = 0;
        _limit = false;
        _activeJobs = 0;
        _probe.Start();
    }

    [Benchmark]
    public async Task<int> SearchCorpus()
    {
        if (Workers == -1)
        {
            using var service = new SourceSearchService(_query, _settings, ArchiveExtensions,
                publish: Capture, report: Report);
            await service.RunAsync(_source, false, CancellationToken.None).ConfigureAwait(false);
            _processed = service.FilesProcessed;
            _matching = service.MatchingFiles;
            _limit = service.ResultLimitReached;
            _activeJobs = 1;
        }
        else
        {
            using var run = Workers == 0
                ? new SearchRunCoordinator(_query, _settings, ArchiveExtensions, publish: Capture, report: Report)
                : new SearchRunCoordinator(_query, _settings, ArchiveExtensions,
                    new SearchConcurrencyPolicy(Workers), publish: Capture, report: Report);
            await run.RunAsync(_source, CancellationToken.None).ConfigureAwait(false);
            _processed = run.FilesProcessed;
            _matching = run.MatchingFiles;
            _limit = run.ResultLimitReached;
            _activeJobs = run.PeakWorkers;
        }
        return _matching;
    }

    [IterationCleanup]
    public void EndOperation()
    {
        var metrics = _probe.Stop();
        ValidateCounts();
        if (Digest() != _expectedDigest)
            throw new InvalidOperationException("Ordered paths, records, context or highlights differ from the sequential baseline.");
        var record = new
        {
            Workload = Workload.ToString(),
            Workers,
            Mode = Workers == -1 ? "OriginalSequential" : Workers == 0 ? "Automatic" : "ForcedWorkers",
            Operation = ++_operation,
            LogicalProcessors = Environment.ProcessorCount,
            Files = _processed,
            ExpandedBytes = _corpus.ExpandedBytes(Workload),
            PeakActiveJobs = _activeJobs,
            Metrics = metrics
        };
        File.AppendAllText(Path.Combine(_evidence, $"{Workload}-{Workers}-{Environment.ProcessId}.jsonl"), JsonSerializer.Serialize(record) + Environment.NewLine);
    }

    private void Capture(SourceSearchResult result) => _results.Add(result);

    private void Report(string source, Exception error, string stage, string severity) =>
        _issues.Add($"{source}: {stage}: {severity}: {error.Message}");

    private void ValidateCounts()
    {
        int expected = _corpus.ExpectedFiles(Workload);
        if (_issues.Count != 0 || _limit || _processed != expected || _matching != expected || _results.Count != expected ||
            _results.Any(result => !string.IsNullOrEmpty(result.PartialReason)))
            throw new InvalidOperationException($"Incomplete scan: files={_processed}, matching={_matching}, results={_results.Count}, expected={expected}, limit={_limit}; {string.Join("; ", _issues)}");
    }

    private string Digest()
    {
        var snapshot = _results.Select(result => new
        {
            result.LogicalPath,
            result.Kind,
            result.PartialReason,
            result.Details,
            result.Events,
            result.Requests
        });
        return Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(snapshot)));
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        _probe?.Dispose();
        _results.Clear();
    }
}
