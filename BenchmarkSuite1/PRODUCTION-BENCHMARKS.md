# Production multicore measurements

`BeaconProductionScalingBenchmarks` invokes the existing sequential `SourceSearchService` and the actual `SearchRunCoordinator`. It does not implement a second search scheduler. The original `BeaconSearchBenchmarks.cs` is unchanged. The earlier experimental four-file multicore class is not present on the Beta branch; its historical timings are not a like-for-like baseline for this suite.

## Fixed corpus

Run from the solution directory with the .NET 10 Windows Desktop runtime installed:

```powershell
$env:BEACON_SCALING_CORPUS = Join-Path $env:TEMP ('BeaconProductionScaling-' + [Guid]::NewGuid().ToString('N'))
$env:BEACON_SCALING_EVIDENCE = Join-Path (Get-Location) 'BenchmarkDotNet.Artifacts\production-telemetry'
dotnet run --project BenchmarkSuite1 -c Release -f net10.0-windows -- --prepare-production-scaling
dotnet run --project BenchmarkSuite1 -c Release -f net10.0-windows -- --filter '*BeaconProductionScalingBenchmarks*' --artifacts 'BenchmarkDotNet.Artifacts\production'
```

Preparation is separate from measurement. It preflights disk space and refuses to overwrite a nonempty directory without a valid completed manifest. The large corpus has `max(8, Environment.ProcessorCount)` independent files and ZIPs. Every worker setting searches the same corpus. Each ZIP is approximately 200 MiB **compressed on disk**, using standard Deflate/Optimal compression; the corresponding plain file contains identical expanded text. A unique marker at EOF forces complete large-file reads. The manifest records actual sizes and the corpus version.

Additional fixed controls cover 1,024 small files, dense matching records, whole JSON documents, and one large archive. Ordinary capture/security limits remain enabled but are chosen not to truncate these fixtures. Single-archive entry processing remains sequential. The synthetic text is deterministic but is not a representative customer-log corpus.

Worker parameters:

- `-1`: original sequential service.
- `0`: production automatic policy, including resource safeguards.
- `1`, `2`, `4`, `8`, `14`, and the available logical-processor count: internal forced-worker policy; duplicates and counts above available capacity are excluded.

There is no new user-facing performance setting. The internal override is for tests/benchmarks only. `PeakActiveJobs` counts admitted source jobs, including jobs waiting for heavy-work admission or ordered output; it does not count simultaneously executing CPU cores.

## Timing, memory, and correctness

BenchmarkDotNet uses one launch, one warmup, and five measured operations per case. Report confidence/error and variation, not just best times. Setup establishes a sequential ordered-result digest; iteration cleanup checks the same paths, records, context, highlights, counts, and absence of partial results or diagnostics. Generation and validation are excluded from benchmark timings.

A separate process-wide probe records CPU time, `GC.GetTotalAllocatedBytes`, GC collection deltas, and working-set/private-byte peaks sampled every 20 ms. These are diagnostic observations and include small monitoring/harness costs. Sampled memory peaks may miss short spikes; they are not RAM-capacity requirements or allocation totals. The JSONL telemetry includes warmup operations; BenchmarkDotNet's report is authoritative for measured timing statistics. Do not confuse allocations with peak/live memory or active jobs with processor utilization.

The configured run emits one warmup, five actual operations, and a final memory-diagnostic operation. `Summarize-ProductionScaling.ps1` uses probe operations 2–6 with the BenchmarkDotNet timing report and rejects an incomplete matrix. Probe GC deltas can include forced collections made by the harness outside the timed body; use BenchmarkDotNet's GC columns for workload collection counts. For example:

```powershell
pwsh -File BenchmarkSuite1/Summarize-ProductionScaling.ps1 -EvidencePath tests/Beacon.SafetyChecks/PerformanceEvidence/multicore-production
```

Setup's sequential validation and repeated operations warm the filesystem cache. These runs are **not cold-cache storage tests**. No cache flushing, affinity changes, thread-pool tuning, or priority changes are performed by the fixture. BenchmarkDotNet's own execution environment should be recorded from its logs. Different CPU topologies, power policies, storage, other running applications, and memory pressure can change results.

BenchmarkDotNet can select a temporary High performance power plan and a different GC configuration from the WPF application. These are benchmark-harness choices, not application settings changes. Record the actual GC mode and the power-plan restoration from the console log. Compare execution modes within the same harness; do not interpret the difference between headless and WPF timings as pure UI overhead when their runtime environments differ.

## Complete application path

The STA safety executable has a separate measurement entry point:

```powershell
dotnet run --project tests/Beacon.SafetyChecks -c Release -- --measure-production-app $env:BEACON_SCALING_CORPUS $env:BEACON_SCALING_EVIDENCE
```

This measures `StartScan` through counting, automatic/forced-one searching, ordered result publication, and final completion. Each case has one warmup and three measured repetitions, alternating mode order. A visible but off-screen WPF window uses the real controls. Startup network/update checks and the welcome offer are disabled by the existing test seams; no saved preferences or welcome marker are changed. Full text preview and Summary switching are measured separately after scanning. Native rendered-document startup is not included in preview interaction measurements. Off-screen dispatcher observations do not replace interactive responsiveness or display-hardware testing.

## Artifacts and cleanup

Keep environment/corpus manifests, BenchmarkDotNet reports/logs, process telemetry, application JSONL, and the final comparison report together. Record unavailable hardware coverage and any failed cases. The shared corpus is deliberately retained across benchmark processes; after preserving evidence, remove only the dedicated generated corpus directory whose path is recorded in its manifest/environment file. Do not remove source trees or arbitrary paths supplied by other users.

The run-specific report under `tests/Beacon.SafetyChecks/PerformanceEvidence/multicore-production` describes the measured implementation, results, and limitations.
