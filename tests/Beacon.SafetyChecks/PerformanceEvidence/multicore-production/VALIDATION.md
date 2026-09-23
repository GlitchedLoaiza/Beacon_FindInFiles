# Final implementation validation

Status: completed in the Beta workspace on September 23, 2026. Changes remain uncommitted; unrelated workspace content was preserved.

## Builds and safety checks

- Visual Studio workspace build: successful.
- Release solution build: successful, **0 warnings and 0 errors** (`final-release-build.log`).
- Debug STA safety executable: **89 passed; 0 failed** (`final-debug-safety.log`).
- Release STA safety executable: **89 passed; 0 failed** (`final-release-safety.log`).
- `git diff --check`: successful.
- `code-snapshot.json` records SHA-256 hashes of source/project/script files at final validation.

The safety executable is a linked-source STA harness, not a Test Explorer test project. No skipped Test Explorer run is represented as a passing unit-test run.

## Acceptance coverage

- Automatic processor-aware selection has no fixed eight- or fourteen-worker ceiling. Policy tests cover 1, 2, 4, 8, 14, 20, 32, and higher reported capacities plus unknown/low-memory fallback. Legacy saved worker values do not cap normal automatic scheduling or get rewritten.
- Ordered results, captured details/context, global capped prefixes, diagnostics, completed-source counters, cancellation, bounded backpressure, and retained nested-preview sources are checked against the sequential path. Archive containers are not counted a second time.
- Stress cases include a slow first source, later archive-result floods, speculative faults beyond a cap, callback failures, cancellation while waiting for output/resources, mixed file types, solid archives, corrupt sources, HAR redaction, and existing EVTX behavior.
- The real application still counts before searching and awaits the final result-publication drain. Integration checks include replay and Summary while a controlled multicore search remains active.
- Text Previous/Next navigation covers variable-length and adjacent matches, direction changes, wraparound, Unicode/newline forms, query modes, zero-width handling, and stale preview reads.
- Summary uses captured export contexts without rereading deleted disk/ZIP/nested/CAB sources. Separate overlapping five-line windows, original line identity, clipping notices, metadata handling, zero-width record anchors, bounded Full-preview targets, and session-only mode behavior are tested.
- A 10,000-record Summary remains virtualized in Light, Dark, and Beacon themes at minimum window dimensions, with font/wrapping and accessibility-name checks.
- Tour replay is separate from startup eligibility. Tests cover existing, missing, and unwritable marker locations; unchanged marker contents/last-write metadata; repeated replay without duplicates; completion/dismissal; active-search preservation; and owner shutdown. Tests use isolated marker paths, not the user's real marker.

## Performance evidence

- All **48 headless production benchmark cases** completed with matching ordered-result digests, expected counts, no unintended partial results, and no diagnostics.
- All **48 application measurement operations** completed with expected counted/scanned/result totals and successful final state. This includes 12 warmups and 36 measured repetitions.
- The measured regex runner-cache allocation issue was addressed by reusing an equivalent query per worker; matching algorithms and safety limits were not rewritten.
- The Full-preview measurement delay was an idle-priority harness wait for off-screen WPF formatting. The helper now checks tracked preview readiness with Background-priority pumping. Production paragraph layout and the configured 25 MiB preview limit were not changed to obtain results.
- `REPORT.md` records timing, variance, allocations, sampled memory, CPU tradeoffs, forced-one overhead on small workloads, and unavailable hardware/storage coverage. No universal speedup or bounded total application-RAM guarantee is claimed.
- Before/after benchmark power-plan records both show Balanced.

## Cleanup and intentionally deferred work

The exact generated large corpus was deleted only after checking its manifest against the retained evidence copy; `cleanup.json` records the removal. Benchmark reports, raw telemetry, application observations, exploratory evidence, and final build/test logs remain available.

Individual archive trees are still sequential. Continuous throughput tuning, custom P/E-core pinning, unrelated matcher rewrites, cold-cache/HDD/network measurements, battery testing, and physical dual-/quad-core hardware validation remain outside this implementation. These are documented scope boundaries, not unfinished implementation steps.

## Later feedback and 2.2.0 preparation

After the original 89-check milestone above, follow-up changes made the Beacon text-selection overlay translucent, added a current/total text-match bar, and cleared all text/EVTX/HAR navigation on Reset, new scans, and deselection. Summary-selector visibility was verified for eligible plain-text results. These follow-ups passed **93 checks in both Debug and Release**, with a successful workspace build and whitespace check.

The application version is subsequently being prepared as **2.2.0**. Fresh version-specific build/test/package evidence and the Beta → Release handoff are recorded in [the 2.2.0 preparation](../../../../docs/releases/2.2.0/PREPARATION.md). Original benchmark logs, 89-check logs, and `code-snapshot.json` retain their original meaning; they have not been relabeled as a validation snapshot or binary checksum for 2.2.0. No performance rerun is implied by a version/documentation change.
