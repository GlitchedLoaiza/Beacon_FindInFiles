# Beacon production multicore and preview measurements

Measured September 22–23, 2026, on the Beta workspace based on commit `20edcaa4ee10ca0997023744d712aa73564624a6` with the implementation changes applied. Reproduction instructions are in [PRODUCTION-BENCHMARKS.md](../../../../BenchmarkSuite1/PRODUCTION-BENCHMARKS.md).

## Environment and workload

- Intel Core i7-12800H: 14 physical cores, 20 logical processors available to the process.
- Approximately 64 GiB installed memory; Kingston NV2 NVMe SSD; Windows 11 Enterprise 10.0.26200.
- SDK 10.0.401; benchmark runtime .NET 10.0.12, x64 RyuJIT/AVX2.
- One fixed 20-source large corpus at every concurrency. Each ZIP contains **209,918,357 compressed bytes (~200.19 MiB)** and expands to **277,872,672 bytes (~265 MiB)**. The equivalent plain files contain identical text. Aggregate expanded work is 5,557,453,440 bytes (~5.18 GiB), not 200 MiB total.
- ZIP Deflate with `CompressionLevel.Optimal`; deterministic synthetic ASCII lines and a unique marker at EOF. The marker requires a complete large-file scan.
- Separate controls: 1,024 small files, dense matching records, whole JSON documents, and one large archive.
- Repeated, cache-influenced reads. Sequential validation warms each benchmark workload. **No cold-cache, HDD, or network-share claims are made.**

All 48 headless cases completed. Each used one warmup, five measured operations, and a memory-diagnostic operation. Ordered result digests compared paths, kinds, details, context, highlights, and structured records against the sequential service; counts, partial-result flags, and diagnostics were also checked. There were no validation mismatches.

The sequential comparison is the current standalone `SourceSearchService` path, not a separately installed historical application binary. The former experimental four-file multicore benchmark is absent from Beta; the existing `BeaconSearchBenchmarks.cs` was preserved unchanged.

## Headless production path

Durations cover the entire fixed workload. Values below are means; full errors, standard deviations, and telemetry are in `comparison.csv`, `comparison.json`, and the BenchmarkDotNet reports.

| Workload | Sequential service | Production automatic | Speedup | Automatic admitted jobs |
|---|---:|---:|---:|---:|
| 20 large plain files | 14.839 s | 3.019 s | 4.92x | 20 |
| 20 compressed ZIPs | 42.404 s | 5.648 s | 7.51x | 20 |
| 1,024 small files | 357.63 ms | 104.44 ms | 3.42x | 20 |
| Dense captured records | 81.28 ms | 17.53 ms | 4.64x | 20 |
| Whole JSON documents | 119.49 ms | 92.98 ms | 1.28x | 20 |
| One large archive | 2.288 s | 2.243 s | effectively unchanged | 1 |

### Worker scaling on the large workloads

| Requested workers | Plain files | ZIPs |
|---|---:|---:|
| 1 | 15.336 s | 42.440 s |
| 2 | 7.207 s | 22.357 s |
| 4 | 4.371 s | 12.929 s |
| 8 | 3.536 s | 9.527 s |
| 14 | 3.202 s | 7.806 s |
| 20 | 2.978 s | 5.666 s |
| Automatic | 3.019 s | 5.648 s |

The automatic path demonstrably admits more than eight jobs on this machine. `PeakActiveJobs` includes jobs waiting for resource admission or ordered output; it is not a count of CPU cores simultaneously executing. Whole-document materialization is more restricted than streaming text/ZIP processing. Individual archive trees remain sequential.

### Overhead, memory, and variance

- Forced-one ZIP time is essentially the sequential service's time. Large-text forced-one time was about 3.35% higher, with substantial run variance.
- Forced-one small-file time was 540 ms versus 358 ms sequential; dense results were 106 ms versus 81 ms. Per-file service snapshots and bounded ordered handoffs are additional work compared with one sequential service over the directory. These safety costs are most visible in short/low-concurrency workloads. They were not hidden by dropping results or weakening resource/ordering protections. Automatic mode was faster in both controls on this machine.
- Short control cases have broad confidence intervals. The automatic large-text run also contained a slower sample: mean 3.019 s, standard deviation 0.687 s. ZIP automatic time was more stable: mean 5.648 s, standard deviation 0.246 s. These are preliminary workload-specific measurements, not universal or exact speedup guarantees.
- Sampled headless peak working set was about **61→66 MiB** for sequential→automatic text, and **66→74 MiB** for ZIPs. Working set excludes the operating system's multi-gigabyte filesystem cache.
- Large scans allocated approximately **19,300 MiB in total over the operation**, mostly short-lived data. This is not resident RAM consumption. Allocation totals stayed essentially level across final worker counts.
- Faster wall time used more total CPU: ZIP process CPU time was approximately 51.7 CPU-seconds sequential versus 84.1 CPU-seconds automatic. Reduced latency does not imply reduced energy use; battery behavior was not measured.
- Probe peaks are sampled every 20 ms and can miss brief spikes. Probe GC deltas include some harness collections; BenchmarkDotNet GC columns describe workload collections per 1,000 operations.

## Measurement-driven correction

The exploratory run shared one `SearchQuery` across workers. Parallel text operations allocated approximately 33–35 GB versus approximately 20.23 GB sequentially. The .NET 10 [Regex implementation](https://github.com/dotnet/runtime/blob/v10.0.0/src/libraries/System.Text.RegularExpressions/src/System/Text/RegularExpressions/Regex.cs) has a single cached runner per regex instance; concurrent sharing can require repeated runner allocation.

Each sequential worker now reuses its own equivalent query. Matching behavior and safety limits are unchanged. The automatic run's allocation total returned to approximately 20.23 GB, and ordered-result checks continued to pass. The stopped exploratory logs are retained under `pre-query-isolation-*`; they are not mixed into the final comparison.

## Complete application path

A separate STA harness used a real, visible-but-off-screen WPF window. It measured `StartScan` through file counting, searching, result publication, and completion. There was one warmup and three measured repetitions per mode/workload, with mode order alternating. Every operation checked counted/scanned/results totals, completion state, and diagnostics. All 48 application operations completed successfully; 36 were measured repetitions.

| Workload | Application forced-one | Application automatic | Speedup |
|---|---:|---:|---:|
| Large plain files | 16.909 s | 2.904 s | 5.82x |
| Compressed ZIPs | 45.235 s | 6.045 s | 7.48x |
| Small files | 576.25 ms | 251.13 ms | 2.29x |
| Dense captured records | 128.68 ms | 101.66 ms | 1.27x |
| Whole JSON documents | 233.18 ms | 186.41 ms | 1.25x |
| Single archive | 2.351 s | 2.355 s | unchanged |

The WPF timings include counting and publication, unlike the headless cases. BenchmarkDotNet used non-concurrent workstation GC and temporarily selected its High performance power plan. Both saved pre/post power-plan records are **Balanced**. The separate application measurements ran under the restored environment. Therefore differences between the two harnesses are not a clean measurement of UI overhead alone.

### Preview interactions

Full preview remained at its configured **25 MiB** limit; it was not shortened to obtain passing measurements. In automatic-mode application repetitions:

- Large plain-file Full preview readiness averaged approximately **834 ms**; switching to captured Summary averaged **67 ms**.
- ZIP Full preview readiness averaged approximately **1,032 ms**; switching to captured Summary averaged **66 ms**.
- The captured EOF match is available in Summary even when it lies beyond the Full prefix.
- Separate functional/presentation checks exercised 10,000 captured records in Light, Dark, and Beacon themes. Only two or three summary block views were realized in the observed viewport, with all records reachable.

The first measurement helper waited at `DispatcherPriority.ContextIdle`, below WPF's continuing off-screen background formatting. That wait was incorrectly broader than preview readiness. It was corrected to pump at Background priority while checking the tracked read/index task and viewport layout. **No production paragraph rewrite or preview-limit reduction was needed.** The investigation log is retained as `application-readiness-investigation.log`.

The persistent WPF measurement process reached a sampled working set of approximately **1.34 GiB** during the large ZIP sequence, which includes previous Full preview allocations and normal WPF/GC retention. This must not be confused with the much smaller fresh headless-process figures or attributed solely to the worker count. Working set subsequently decreased on later workloads. Native rendered-document startup was not included in preview interaction timings and is recorded as unmeasured, not zero cost. Off-screen testing does not replace interactive display/battery/hardware testing.

## Accepted boundaries and remaining coverage gaps

- No fixed eight- or fourteen-worker ceiling. Processor availability, work availability, memory admission, and bounded output control execution automatically.
- Stable traversal-order publication and global capped prefixes are preserved. A slow earliest source can limit throughput because ordering/backpressure are deliberate safety requirements.
- More workers do not guarantee proportional gains. Small workloads at one worker can be slower than the standalone sequential service.
- One archive tree remains sequential; the multi-archive speedup does not imply intra-archive parallelism.
- No physical dual-/quad-core machine, HDD, network share, cold-cache run, or battery-energy test was available. Injected policy tests cover processor counts and low/unknown-memory fallback, not hardware speedups.
- Continuous throughput tuning, core pinning, and matching-engine rewrites were not introduced.

## Evidence

- `environment.json`, `corpus.json`: machine and immutable workload description.
- `bdn/results/BeaconProductionScalingBenchmarks-report.*`: BenchmarkDotNet timing/allocation reports.
- `benchmark-console.log`: complete final benchmark log and runtime details.
- `telemetry/*.jsonl`: process-wide diagnostic observations.
- `comparison.csv` / `comparison.json`: validated headless summary.
- `application/application.jsonl`, `application-console.log`: complete application operations.
- `application-comparison.csv` / `.json`: measured application summaries.
- `power-plan-before.txt`, `power-plan-after.txt`: environment restoration.
- `safety-after-query-isolation.log`: 89 passing safety checks after the measured backend correction.

Run `BenchmarkSuite1/Summarize-ProductionScaling.ps1` to regenerate comparison artifacts. The large generated corpus is disposable after evidence is retained; its recorded path is not an application input or a required repository artifact.
