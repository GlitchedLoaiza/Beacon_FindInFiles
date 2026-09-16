# Measured performance audit — Beacon 2.1

## Scope and rollback baseline

Production baseline: commit `120daa4a5888aea05d63a93b169d0d74abcbd413` on `experiments-latest`. Production was unchanged when the baseline below was measured. The solution and `BenchmarkSuite1` were the only pending measurement-host changes.

This is an incremental, measurement-driven audit, not a claim that every workload is optimized or that all release gates are complete. Do not remove security limits, checksum validation, redaction, cancellation or context fidelity to improve a timing.

## Measurement host repair (H01)

`BenchmarkSuite1` targets .NET 10 Windows with WPF/WinForms references. It references the real Beacon project; no search implementation is copied into benchmarks. The SDK's `ValidateExecutableReferencesMatchSelfContained=false` is scoped to this API-calling measurement host. Beacon's self-contained, single-file and ReadyToRun settings are unchanged. `TargetFrameworks` retains the generated host's multi-target property form with one target.

The generated CPU diagnoser is retained. Both benchmark methods and their setup are frozen after the first successful run. They call `SearchTextCollector.CollectAsync` and `SourceSearchService.RunAsync`, respectively. A 100,000-line UTF-8 input is prepared outside measurement; ZIP entries are stored/uncompressed to expose streaming/checksum overhead. MatchEvery=0 has no matches, 1000 has 100 matches, and 1 stops at the configured 10,000-record limit. Dense runs must not be compared to sparse runs as if they read the same number of lines.

## Baselines

- Initial interactive app CPU trace: `aa655bd1-180d-4228-bf96-667b1af63a71`. Text collection total CPU 15.84%; ZIP read total 9.48%; CRC32 self CPU 5.11%. Interactive timing/workload is not controlled. Unresolved runtime frames limit attribution.
- Repeatable CPU benchmark trace: `e88862ef-272e-4e1c-add0-71e10fbc46c7`. CPU summary identifies `MatchContext.FromText` (7.77% total, 2.89% self); most runtime samples are unresolved. Use BenchmarkDotNet times, not raw sample totals, for before/after claims.

| Method | MatchEvery | Mean ms | Error ms | StdDev ms | Allocated MB | Gen0 / 1000 ops | Gen1 / 1000 ops | Gen2 / 1000 ops |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| PlainText | 0 | 15.00 | 0.179 | 0.168 | 32.81 | 2734.3750 | 46.8750 | 0 |
| StoredZipSearch | 0 | 29.83 | 0.128 | 0.120 | 33.19 | 2750.0000 | 93.7500 | 0 |
| PlainText | 1 | 38.73 | 0.727 | 1.065 | 47.71 | 4230.7692 | 2076.9231 | 1076.9231 |
| StoredZipSearch | 1 | 39.03 | 0.635 | 0.594 | 47.91 | 3923.0769 | 1923.0769 | 923.0769 |
| PlainText | 1000 | 14.64 | 0.093 | 0.087 | 33.24 | 2765.6250 | 31.2500 | 0 |
| StoredZipSearch | 1000 | 31.06 | 0.241 | 0.225 | 33.61 | 2781.2500 | 62.5000 | 0 |

## Final decision — 2026-09-15

**Retain the combined P02 + P04 candidate.** A counterbalanced A1–B1–B2–A2 sequence showed both candidate runs faster than both original-code runs in every measured scenario. Allocated bytes were identical between repeats of each variant and decreased in every scenario. Some higher-generation GC counts increased; those trade-offs are reported below.

Only `Beacon_FindInFiles/SearchResults.vb` differs in production from the baseline: 44 added / 9 removed lines. `ArchiveCompatibility.vb`, `ArchiveSafety.vb`, `MatchContext.vb`, `SearchQuery.vb`, project dependencies, UI and publishing settings are unchanged. P03 was fully reverted. These are text-collection improvements, **not** a blanket application speedup, a leak audit, a certification of every format, or release approval.

## Change and experiment ledger

| ID | Decision | Files / scope | Purpose, risk and evidence |
| --- | --- | --- | --- |
| H01 | Keep | `BenchmarkSuite1/BenchmarkSuite1.csproj`, solution registration; generated benchmark/program source retained | Repair .NET 10/WPF/WinForms/project references; scope the SDK self-contained-reference exception to the API-calling benchmark host. Frozen benchmark source exercises the real application methods. No copy of production logic is benchmarked. |
| P01 | Superseded, not retained as a separate patch | Earlier `SearchResults.vb` bookkeeping experiment | Deferred line-label formatting; initially reused a transient line and removed a pending-array copy. Allocation improved, but early dense timings regressed. Label-only isolation did not establish consistent timing; original-code reruns also slowed. Do not cite P01 as an independently proven speedup. |
| P02 | Keep with P04 as one measured change | Line-oriented branch of `SearchTextCollector.CollectAsync` | Format labels only for matches; build a matching line's detail/capture once instead of invoking the multiline helper and discarding its context; remove finished pending contexts by index. Risk: excerpt/cropping, focus, order and aliasing regressions; parity and existing context tests cover them. |
| P03 | Rejected / reverted | Experimental async overrides in `BoundedReadStream` and `VerifiedZipStream` | Attempted to remove base-stream async scheduling. Increased ZIP allocations without a consistent elapsed-time benefit. Neither async override nor its experiment-only test remains. Existing CRC/cancellation tests remain intact. |
| P04 | Keep with P02 as one measured change | Private `CopyHistory` / `CopyContextLine` in `SearchResults.vb` | Copy bounded file-context windows directly with known capacity; avoid layered LINQ and transient collections. Text strings can be shared because they are immutable; published `ContextLine` and `SearchSpan` objects stay independent. P04 was present but unvalidated at the previous handoff; this A/B validates the combined P02+P04 state, not P04 independently. |
| T01 | Keep | `tests/Beacon.SafetyChecks/OptimizationChecks.vb`, one registration in `Program.vb` | Adds all-mode dense context/order/EOF/focus checks, mutation isolation, long Unicode/tabbed/empty lines and zero-width/first-only behavior. Compares line details against the unchanged general `FromRecord` helper. Passes on both A and B; no existing tests were weakened. |
| E01 | Keep | This log, `PerformanceEvidence`, README/release-validation links | Preserves original/exploratory/controlled evidence, exact identities, reversible production patch and remaining audit limitations. No benchmark binaries, customer logs, production settings or tour markers are added to test evidence. |

## Controlled A/B protocol

- Baseline A: the exact `SearchResults.vb` at `120daa4a5888aea05d63a93b169d0d74abcbd413`. A1 and A2 production trees were checked against that commit, not approximated with a substitute implementation.
- Candidate B: the existing P02+P04 source saved before switching. Source and benchmark hashes were checked before measurement and after restoring the final state. A patch and backup were saved before the first revert. No other production file changed between runs.
- Sequence: **A1, B1, B2, A2**, using the same CPU benchmark collection facility and the same fully qualified targets: `BeaconSearchBenchmarks.PlainText` and `BeaconSearchBenchmarks.StoredZipSearch`. Each collection ran all six frozen cases. No benchmark logic/setup/attributes were changed after their first successful run.
- Machine: Intel Core i7-12800H, 14 cores / 20 logical processors; Windows 11 build 26200; .NET SDK 10.0.401; BenchmarkDotNet 0.15.2; .NET 10.0.12 x64 RyuJIT AVX2, concurrent workstation GC.
- The OS reported Balanced before and after. BenchmarkDotNet's own logs report its High performance power-plan setup during each benchmark; the agent did not override affinity, priority or power settings and did not kill unrelated workloads. This is a development machine, not a thermally controlled laboratory.
- Mean comparisons below are the arithmetic average of the two BDN run means for each variant, calculated from the saved CSV's printed precision. They are not a pooled confidence interval or an independent statistical significance test. `Error`, `StdDev`, iteration samples and BDN outlier diagnostics are preserved in the raw reports/logs; no new manual outlier filtering was applied. Dense B runs have noticeable iteration variation but remain faster than both A runs.
- BenchmarkDotNet memory/GC numbers are per measured operation (GC columns per 1,000 operations). They are not heap-retention/leak measurements. The stored ZIP fixture intentionally does not measure decompressor performance, CAB, RAR, EVTX, HAR or GUI latency. Dense runs stop at the 10,000-record limit rather than consume all 100,000 lines.

### Run identities

| Run | Variant | CPU trace session |
| --- | --- | --- |
| A1 | Original | `0ec898f6-1f52-427b-9a90-2df376f8e6af` |
| B1 | P02+P04 | `362f1331-f269-4281-84e1-3fe202c9ab66` |
| B2 | Same P02+P04, no intervening edits | `2b3889a1-fe22-47bb-a90a-701a7f5ff4e0` |
| A2 | Original restored and identity-checked | `f7d3da09-95b8-4faf-87c6-f8889fc8e4a1` |

| Identity | SHA-256 |
| --- | --- |
| Baseline `SearchResults.vb` | `5B20F243A9691E00DC314FFAEDE9233DF892E092EC177C8AE754E55E2E1D8027` |
| Candidate `SearchResults.vb` | `50DB802123AD30C9A1385A4805B9DA9064D9F70E5070A539BB20195BEDD4E708` |
| Frozen `BeaconSearchBenchmarks.cs` | `B4F38E8B56726DF3FC122623E557E7C2D2A0B94A8A33F537F9E710805452B8AD` |
| Frozen benchmark `.csproj` | `9C544A3B4AB8EF1D7379BC16AA63394CDA50A5BBFCB788EB5625E1319B423C46` |

These are exact file-byte hashes from this Windows workspace; checkout newline conversions can alter a file-byte hash. The patch also records Git content identities for review.

### All four elapsed-time results (milliseconds)

| Method | MatchEvery | A1 | B1 | B2 | A2 |
| --- | ---: | ---: | ---: | ---: | ---: |
| PlainText | 0 | 23.19 | 15.13 | 14.78 | 23.73 |
| StoredZipSearch | 0 | 42.22 | 35.36 | 34.37 | 44.79 |
| PlainText | 1 | 58.06 | 42.68 | 41.66 | 58.27 |
| StoredZipSearch | 1 | 57.15 | 45.86 | 43.90 | 59.36 |
| PlainText | 1000 | 24.33 | 15.14 | 14.95 | 24.10 |
| StoredZipSearch | 1000 | 45.03 | 36.21 | 34.57 | 44.54 |

### Before / after / delta

Lower time and allocated bytes are better. Rounded presentation uses the saved CSV precision; full comparison is in `AB-comparison.csv`.

| Method | MatchEvery | A mean ms | B mean ms | Time delta | A allocated MB | B allocated MB | Allocation delta |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| PlainText | 0 | 23.460 | 14.955 | -36.25% | 32.81 | 25.95 | -20.91% |
| StoredZipSearch | 0 | 43.505 | 34.865 | -19.86% | 33.19 | 26.33 | -20.67% |
| PlainText | 1 | 58.165 | 42.170 | -27.50% | 47.71 | 22.23 | -53.41% |
| StoredZipSearch | 1 | 58.255 | 44.880 | -22.96% | 47.91 | 22.43 | -53.18% |
| PlainText | 1000 | 24.215 | 15.045 | -37.87% | 33.24 | 26.12 | -21.42% |
| StoredZipSearch | 1000 | 44.785 | 35.390 | -20.98% | 33.61 | 26.49 | -21.18% |

### GC trade-offs (mean counts per 1,000 operations)

| Method | MatchEvery | Generation | A | B | Count delta | Percent delta |
| --- | ---: | --- | ---: | ---: | ---: | ---: |
| PlainText | 0 | Gen0 | 2726.56 | 2156.25 | -570.31 | -20.92% |
| PlainText | 0 | Gen1 | 39.06 | 46.88 | +7.81 | +20.00% |
| PlainText | 0 | Gen2 | 0 | 0 | 0 | 0% |
| StoredZipSearch | 0 | Gen0 | 2738.64 | 2171.43 | -567.21 | -20.71% |
| StoredZipSearch | 0 | Gen1 | 87.12 | 69.05 | -18.07 | -20.75% |
| StoredZipSearch | 0 | Gen2 | 0 | 0 | 0 | 0% |
| PlainText | 1 | Gen0 | 4166.67 | 2214.29 | -1952.38 | -46.86% |
| PlainText | 1 | Gen1 | 2083.33 | 1285.71 | -797.62 | -38.29% |
| PlainText | 1 | Gen2 | 1083.33 | 428.57 | -654.76 | -60.44% |
| StoredZipSearch | 1 | Gen0 | 3900 | 2250 | -1650 | -42.31% |
| StoredZipSearch | 1 | Gen1 | 1900 | 1500 | -400 | -21.05% |
| StoredZipSearch | 1 | Gen2 | 900 | 583.33 | -316.67 | -35.19% |
| PlainText | 1000 | Gen0 | 2750 | 2171.88 | -578.12 | -21.02% |
| PlainText | 1000 | Gen1 | 62.50 | 328.12 | +265.62 | +425.00% |
| PlainText | 1000 | Gen2 | 0 | 46.88 | +46.88 | n/a (zero baseline) |
| StoredZipSearch | 1000 | Gen0 | 2763.64 | 2207.14 | -556.49 | -20.14% |
| StoredZipSearch | 1000 | Gen1 | 190.91 | 207.14 | +16.23 | +8.50% |
| StoredZipSearch | 1000 | Gen2 | 0 | 0 | 0 | 0% |

All six scenarios are faster in this sequence, but not every GC metric improves. In particular sparse plain text promotes/collects more at higher generations. This result must not be advertised as universally reducing GC pauses or memory retention. Representative long-running scans still need observation.

## Prior exploratory runs — not discarded

The earlier measurements drifted considerably even with original code restored. Their raw logs and session index are preserved under [PerformanceEvidence/20260915-exploratory](PerformanceEvidence/20260915-exploratory), not used as the sole acceptance baseline. The performance statements from that phase were provisional.

| Trace session | Variant / finding |
| --- | --- |
| `e88862ef-272e-4e1c-add0-71e10fbc46c7` | Initial original benchmark baseline, shown above |
| `436d26ea-0e44-43ad-8f2c-5049620b530a` | P01: sparse/no-match faster, dense slower; not accepted |
| `a8488a21-3914-4af4-8848-ffd6e2abf917` | Label-only experiment; unresolved timing regression |
| `9d84a059-15c8-46cf-bed8-5cdc3632e591` | Restored original also slowed; confirmed cross-run drift |
| `7a64cbe3-f020-42f6-974a-4e0739a3ff3d` | P02 candidate; lower allocations, mixed timing versus initial baseline |
| `63f666e1-7a45-405b-80c0-c791d5a693bc` | P02 confirmation; allocations reproduced, ZIP timing unstable |
| `59df3b70-48ab-48bd-ad4b-6be3c1c280e6` | P03 async-wrapper experiment increased ZIP allocations; rejected |
| `181101da-700a-460f-b0dc-df9222807a12` | P02 with P03 reverted; still variable timing; P04 not yet applied |

The official BDN log for `59df3b70...` records PlainText/no-match median **9.058 ms**. An earlier chat table accidentally copied 0.0862 (its standard deviation) into that median cell; the preserved machine output is authoritative.

## Evidence locations and validation

[PerformanceEvidence/20260915-abba](PerformanceEvidence/20260915-abba) contains all four BDN Markdown/CSV reports, complete benchmark logs as `.txt`, source/host manifests, environment, `AB-comparison.csv`, the production patch, regression outputs, publish log and artifact manifest. No reports are reconstructed in place of machine output. Original trace documents are referenced by session ID; they are local profiler artifacts rather than portable source files.

- New parity tests passed on the unmodified baseline and the measured candidate before A/B.
- Final retained candidate: **71 groups passed in Release and 71 in Debug**, including line-context parity, existing first-match, document, cancellation, read-error, Unicode, ZIP CRC/limit, nested archive and UI tests.
- Dependency audit: no known vulnerable packages reported for application, regression project or benchmark project using the configured feeds (see `dependency-audit.txt`). This is not a complete security audit.
- Local production-only Release publish succeeded with existing .NET 10 / win-x64 / self-contained / single-file / ReadyToRun / no-trimming settings. No production project property or dependency changed. Benchmark tooling is not a production project reference and no benchmark/diagnoser files were emitted in the publish folder.
- Published executable: `Beacon_FindInFiles/bin/Release/performance-abba-20260915-140418/Beacon.exe`, **85,809,150 bytes**, SHA-256 `3685916E8524B4DC34EE5F8C78B69A238022D423CCB42F0FB57CFCE2B2DBFF5E`, file version `2.1.0.0`, **unsigned**. Three existing WebView2 documentation XML files were also emitted. No executable was uploaded or signed.
- Product version embeds the baseline commit because these changes are not committed. The commit alone does not identify the changed candidate: use the source/patch and executable hashes above.
- The published executable was not interactively launched in this validation pass; real user settings/onboarding state were not touched. Repeat packaged-app/customer-file smoke tests on the exact final build before shipping. Existing Step 12 signing, notices, versioning and manual gates remain open.

## Rollback — targeted and tested

The file [candidate-P02-P04.patch](PerformanceEvidence/20260915-abba/candidate-P02-P04.patch) is a normal Git patch from the baseline to the retained production candidate, for `SearchResults.vb` only. Its reverse application was checked against the final source. A1/A2 restored the baseline successfully and passed the same parity tests; B1/B2 and the final state matched the saved candidate hash. No whole-tree reset, checkout or commit was performed.

Run from the repository root. Stop if `--check` fails; later source edits may need a manual hunk-by-hunk merge. Do not use `--force`, discard unrelated changes, or restore the entire repository.

```powershell
$patch = 'tests/Beacon.SafetyChecks/PerformanceEvidence/20260915-abba/candidate-P02-P04.patch'
git status --short
git diff -- Beacon_FindInFiles/SearchResults.vb
git apply --reverse --check $patch
if ($LASTEXITCODE -ne 0) { throw 'Rollback conflicts with current source; review manually' }
git apply --reverse $patch
if ($LASTEXITCODE -ne 0) { throw 'Rollback failed' }

dotnet build Beacon_FindInFiles.slnx -c Release
if ($LASTEXITCODE -ne 0) { throw 'Build failed after rollback' }
dotnet run --project tests/Beacon.SafetyChecks/Beacon.SafetyChecks.vbproj -c Release -- "$PWD/Beacon_FindInFiles/7za.exe"
if ($LASTEXITCODE -ne 0) { throw 'Regression failed after rollback' }
```

To reapply after a successful rollback, run `git apply --check` and then `git apply` with the same patch (without `--reverse`), check exit codes, rebuild and retest. Keep H01, T01 and the evidence while bisecting; they call the real implementation and pass with both variants. P02/P04 were accepted as a combined change; partial rollback requires new measurements. To stop shipping the optimization, republish after rollback and produce a new hash; the old executable does not change when source is reverted. Remove the benchmark project/solution entry only if intentionally abandoning the measurement tooling, and review separately from production rollback.

## Broader audit follow-ups — identified, not optimized in this pass

| Area inspected | Observation / next evidence needed | Status |
| --- | --- | --- |
| UI progress | `UpdateScanProgress` schedules a dispatcher callback per processed file; `FileLabelTimer_Tick` is currently empty despite throttling comments. Measure many-small-file UI latency and stale-update behavior before coalescing. | Unchanged; targeted benchmark/trace pending |
| ZIP streams | Initial interactive trace charged CRC work; raw reads must retain validation. The async-wrapper experiment did not justify itself. Do not replace CRC or bypass byte limits without its own correctness/baseline cycle. | P03 rejected; original safety code retained |
| Compressed TAR/CAB | User-requested full pre-count and search can repeat bounded decompression/extraction. Caching changes temporary-disk/security/lifetime behavior, not a trivial speed switch. | Unchanged by design |
| HAR | JSON DOM, joined search strings, decoding and redaction can allocate heavily on large captures. Current benchmarks do not cover it. | Separate representative memory/CPU measurement pending |
| EVTX | Native message formatting and display-name lookup had measurable interactive CPU; caching/matching changes may alter provider/culture behavior. | Native-file measurement pending; no change |
| Preview/WebView2 | Bounded prefix reading and fallback remain; source opening/reading in UI adapters can affect responsiveness. | Long-file/cancellation/UI latency measurements pending |
| Diagnostics/export | Detached snapshots and deduplicated bounded diagnostics are preserved. Large exports and repeated snapshots need their own measurements. | Unchanged |
| Tour, update checks, single instance | Lifecycle/persistence covered by regression tests; these fixtures do not measure startup or long-lived idle behavior. | No performance changes; manual smoke gates remain |

This closes the requested controlled A/B decision and change/rollback log. It does **not** close the entire application-wide performance audit or grant release sign-off. Follow-up changes require their own fixed-workload baseline, correctness checks and before/after validation.
