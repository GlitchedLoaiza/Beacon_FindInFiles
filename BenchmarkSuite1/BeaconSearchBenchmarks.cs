using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using Beacon.Beacon;
using Microsoft.VSDiagnostics;

[CPUUsageDiagnoser]
[MemoryDiagnoser]
public class BeaconSearchBenchmarks
{
    private byte[] _text = Array.Empty<byte>();
    private string _zip = "";
    private string _root = "";
    private SearchQuery _query = null!;
    private BeaconSettings _settings = null!;
    [Params(0, 1000, 1)]
    public int MatchEvery { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var text = new StringBuilder();
        for (int i = 1; i <= 100000; i++)
            text.Append(i.ToString("D6")).Append(MatchEvery > 0 && i % MatchEvery == 0 ? " error code=503 request failed" : " info code=200 request complete").Append('\n');
        _text = Encoding.UTF8.GetBytes(text.ToString());
        _query = new SearchQuery("error", SearchMode.PlainText, false);
        _settings = new BeaconSettings
        {
            StopAfterFirstMatchPerFile = false,
            MaximumStructuredMatches = 10000,
            MaximumCompressionRatio = 10000
        };
        _root = Path.Combine(Path.GetTempPath(), "BeaconBenchmark-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_root);
        _zip = Path.Combine(_root, "input.zip");
        using var zip = ZipFile.Open(_zip, ZipArchiveMode.Create);
        using var output = zip.CreateEntry("input.log", CompressionLevel.NoCompression).Open();
        output.Write(_text);
    }

    [Benchmark]
    public async Task<int> PlainText()
    {
        using var stream = new MemoryStream(_text, false);
        var result = await SearchTextCollector.CollectAsync(stream, _query, 10000, false, false, CancellationToken.None);
        return result.Details.Count;
    }

    [Benchmark]
    public async Task<int> StoredZipSearch()
    {
        using var service = new SourceSearchService(_query, _settings, new[] { ".zip" });
        await service.RunAsync(_zip, false, CancellationToken.None);
        return service.MatchingFiles;
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (Directory.Exists(_root))
            Directory.Delete(_root, true);
    }
}