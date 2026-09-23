#nullable enable
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;

public enum ProductionSearchWorkload
{
    PlainFiles,
    CompressedZips,
    SmallFiles,
    DenseResults,
    WholeDocuments,
    SingleArchive
}

internal sealed class ProductionSearchCorpus
{
    internal const string Marker = "BEACON_PRODUCTION_SCALING_MATCH";
    private const string ManifestName = "beacon-production-corpus.json";
    private const int SchemaVersion = 2;
    private const long TargetCompressedBytes = 200L * 1024 * 1024;
    internal string Root { get; }
    internal CorpusManifest Manifest { get; }

    private ProductionSearchCorpus(string root, CorpusManifest manifest)
    {
        Root = root;
        Manifest = manifest;
    }

    internal static ProductionSearchCorpus Open()
    {
        string root = Path.GetFullPath(Environment.GetEnvironmentVariable("BEACON_SCALING_CORPUS") ??
            Path.Combine(Path.GetTempPath(), "BeaconProductionScaling-v2-" + Environment.ProcessorCount));
        string path = Path.Combine(root, ManifestName);
        if (!File.Exists(path))
            throw new InvalidOperationException("Prepare the fixed corpus first: --prepare-production-scaling. Corpus path: " + root);
        var manifest = JsonSerializer.Deserialize<CorpusManifest>(File.ReadAllText(path)) ?? throw new InvalidDataException("Missing corpus manifest.");
        if (manifest.Version != SchemaVersion || manifest.Files < Environment.ProcessorCount || manifest.Files < 8)
            throw new InvalidDataException("The corpus must contain enough independent jobs for this machine; prepare a new corpus path.");
        return new ProductionSearchCorpus(root, manifest);
    }

    internal static ProductionSearchCorpus Prepare()
    {
        string root = Path.GetFullPath(Environment.GetEnvironmentVariable("BEACON_SCALING_CORPUS") ??
            Path.Combine(Path.GetTempPath(), "BeaconProductionScaling-v2-" + Environment.ProcessorCount));
        if (File.Exists(Path.Combine(root, ManifestName))) return Open();
        if (Directory.Exists(root) && Directory.EnumerateFileSystemEntries(root).Any())
            throw new IOException("Refusing to overwrite a nonempty corpus directory without its completed manifest: " + root);
        int files = Math.Max(8, Environment.ProcessorCount);
        long needed = files * 650L * 1024 * 1024 + 2L * 1024 * 1024 * 1024;
        var drive = new DriveInfo(Path.GetPathRoot(root)!);
        if (drive.AvailableFreeSpace < needed)
            throw new IOException($"Corpus needs approximately {needed:N0} free bytes including a reserve; available {drive.AvailableFreeSpace:N0}.");
        Directory.CreateDirectory(root);
        foreach (string folder in new[] { "plain", "zip", "small", "dense", "documents" })
            Directory.CreateDirectory(Path.Combine(root, folder));
        string textPath = Path.Combine(root, "plain", "00000.log");
        string zipPath = Path.Combine(root, "zip", "00000.zip");
        var random = new Random(20260401);
        byte[] alphabet = Encoding.ASCII.GetBytes("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789+/");
        byte[] block = new byte[1024 * 1024];
        using (var text = new FileStream(textPath, FileMode.CreateNew, FileAccess.Write))
        using (var file = new FileStream(zipPath, FileMode.CreateNew, FileAccess.Write))
        {
            using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: true);
            using var entry = archive.CreateEntry("input.log", CompressionLevel.Optimal).Open();
            while (file.Length < TargetCompressedBytes)
            {
                random.NextBytes(block);
                for (int i = 0; i < block.Length; i++)
                    block[i] = (i & 127) == 127 ? (byte)'\n' : alphabet[block[i] & 63];
                text.Write(block);
                entry.Write(block);
            }
            byte[] marker = Encoding.ASCII.GetBytes(Marker + "\n");
            text.Write(marker);
            entry.Write(marker);
        }
        long compressed = new FileInfo(zipPath).Length;
        long expanded = new FileInfo(textPath).Length;
        if (compressed < TargetCompressedBytes || compressed > TargetCompressedBytes + 4 * 1024 * 1024)
            throw new InvalidDataException("The generated archive is not approximately 200 MiB compressed.");
        for (int i = 1; i < files; i++)
        {
            File.Copy(textPath, Path.Combine(root, "plain", $"{i:D5}.log"));
            File.Copy(zipPath, Path.Combine(root, "zip", $"{i:D5}.zip"));
        }
        string normal = "info request completed " + new string('x', 80) + "\n";
        string small = string.Concat(Enumerable.Repeat(normal, 127)) + Marker + "\n";
        for (int i = 0; i < 1024; i++)
            File.WriteAllText(Path.Combine(root, "small", $"{i:D5}.log"), small, new UTF8Encoding(false));
        var dense = new StringBuilder();
        for (int i = 0; i < 4096; i++)
            dense.Append(i % 4 == 3 ? Marker : "info request completed").Append(" record=").Append(i).Append('\n');
        string document = JsonSerializer.Serialize(new { entries = Enumerable.Repeat(new string('x', 240), 4096).Append(Marker).ToArray() });
        for (int i = 0; i < files; i++)
        {
            File.WriteAllText(Path.Combine(root, "dense", $"{i:D5}.log"), dense.ToString(), new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(root, "documents", $"{i:D5}.json"), document, new UTF8Encoding(false));
        }
        var manifest = new CorpusManifest
        {
            Version = SchemaVersion,
            Files = files,
            LogicalProcessorsAtCreation = Environment.ProcessorCount,
            CompressedBytesPerArchive = compressed,
            ExpandedBytesPerFile = expanded,
            Compression = "ZIP Deflate, CompressionLevel.Optimal",
            CreatedUtc = DateTimeOffset.UtcNow
        };
        File.WriteAllText(Path.Combine(root, ManifestName), JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
        return new ProductionSearchCorpus(root, manifest);
    }

    internal string Source(ProductionSearchWorkload workload) => workload switch
    {
        ProductionSearchWorkload.PlainFiles => Path.Combine(Root, "plain"),
        ProductionSearchWorkload.CompressedZips => Path.Combine(Root, "zip"),
        ProductionSearchWorkload.SmallFiles => Path.Combine(Root, "small"),
        ProductionSearchWorkload.DenseResults => Path.Combine(Root, "dense"),
        ProductionSearchWorkload.WholeDocuments => Path.Combine(Root, "documents"),
        ProductionSearchWorkload.SingleArchive => Path.Combine(Root, "zip", "00000.zip"),
        _ => throw new ArgumentOutOfRangeException(nameof(workload))
    };

    internal int ExpectedFiles(ProductionSearchWorkload workload) => workload switch
    {
        ProductionSearchWorkload.SmallFiles => 1024,
        ProductionSearchWorkload.SingleArchive => 1,
        _ => Manifest.Files
    };

    internal long ExpandedBytes(ProductionSearchWorkload workload)
    {
        if (workload == ProductionSearchWorkload.SingleArchive) return Manifest.ExpandedBytesPerFile;
        if (workload == ProductionSearchWorkload.CompressedZips) return Manifest.ExpandedBytesPerFile * Manifest.Files;
        return Directory.EnumerateFiles(Source(workload)).Sum(path => new FileInfo(path).Length);
    }

    internal sealed class CorpusManifest
    {
        public CorpusManifest() { }
        public int Version { get; set; }
        public int Files { get; set; }
        public int LogicalProcessorsAtCreation { get; set; }
        public long CompressedBytesPerArchive { get; set; }
        public long ExpandedBytesPerFile { get; set; }
        public string Compression { get; set; } = "";
        public DateTimeOffset CreatedUtc { get; set; }
    }
}
