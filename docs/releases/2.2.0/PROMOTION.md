# Beacon 2.2.0 — Release promotion and upload handoff

## Status

**Validated Release source and shipping package, ready for the maintainer's GitHub Release upload.** The authorized Beta/Release history reconciliation is committed as `06b8a89c26a6723e2625d2e66a8bdd58e6d5920d`. The later documentation-only commit records the package without changing its application inputs. Beta and Release are published together after these records are committed; the final handoff verifies their remote tips. No GitHub Release asset upload or tag creation is performed by this task.

## Preserved inputs and history

- Beta source before reconciliation: `56e07390240620a758f8fe7dd79c9f252a5e5be1`.
- Release before promotion: `9b3932298096d74a1a530833b9a94bca9fa476d8`.
- Local and remote `Beacon2.1` preserve that exact prior Release commit and must not move.
- The merge retains Beta's complete 2.2 application/tests/benchmarks and Release's older documentation and `docs/releases/2.1.0/packages/7dbb342/` records.
- The four documentation conflicts are reconciled without force-resetting or replacing either history. The [original Beta candidate](PREPARATION.md) and its hashes remain archived; they are not the final shipping package.

## Package policy

A fresh package is built from a clean, committed **Release** source revision using `scripts/Publish-Beacon.ps1` and its existing default settings: version 2.2.0 / file version 2.2.0.0, Release configuration, .NET 10 Windows x64, self-contained single-file, trimming off, ReadyToRun off. Package/dependency versions and unsigned status are unchanged.

Upload files will be placed in the workspace-local directory:

`artifacts/Beacon-2.2.0-Release/`

The deliverables are `Beacon-2.2.0-win-x64.zip`, `Beacon.exe`, `SHA256SUMS.txt`, and `release-manifest.json`. The ZIP includes the required legal notices. The [shipping manifest](packages/06b8a89/release-manifest.json) records clean Release source, build time **2026-09-23T22:23:21.5559591Z**, bundled .NET/Windows Desktop **10.0.12**, and unchanged package versions.

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| `Beacon.exe` | 77,655,857 | `376D76BD7C6B488A148728319DF450B450125A69C841A767EEC0281BACB99935` |
| `Beacon-2.2.0-win-x64.zip` | 72,102,813 | `F909198059D60551333218E2232EF92DE485BA3A7DD813B39A374B857B7A62CD` |

The [checksum file](packages/06b8a89/SHA256SUMS.txt) is an exact copy of the upload companion. Raw publishing contains only `Beacon.exe`; the ZIP's **12 entries** were independently hashed against staging and include Beacon/component legal notices, not tests, benchmarks, private files, or loose API-documentation XML. See [package verification](packages/06b8a89/package-verification.json) and [entry hashes](packages/06b8a89/package-contents.json).

## Completed validation

| Check | Result | Evidence |
| --- | --- | --- |
| Workspace build | Passed | Visual Studio build result |
| Debug solution build | Passed; zero warnings/errors | [debug-build.txt](packages/06b8a89/debug-build.txt) |
| Release solution build | Passed; zero warnings/errors | [release-build.txt](packages/06b8a89/release-build.txt) |
| Debug safety suite | **93 passed; 0 failed** | [debug-regressions.txt](packages/06b8a89/debug-regressions.txt) |
| Release safety suite | **93 passed; 0 failed** | [release-regressions.txt](packages/06b8a89/release-regressions.txt) |
| Repeated Release suite | **93 passed; 0 failed** | [release-repeat-regressions.txt](packages/06b8a89/release-repeat-regressions.txt) |
| Direct/transitive package audit | No known vulnerabilities reported using configured feeds | [dependency-audit.json](packages/06b8a89/dependency-audit.json) |
| Exact ZIP executable smoke | Startup, duplicate instance, normal close passed | [published-smoke.json](packages/06b8a89/published-smoke.json), [log](packages/06b8a89/published-smoke.txt) |

The main window opened in **5,475.4 ms** in this single smoke observation, not a performance benchmark. The second process exited with code 0 while the original remained running, and normal shutdown exited 0 without forced cleanup. Existing settings and onboarding marker contents, sizes, and last-write timestamps were unchanged. No private-file scan, settings save, ownership recovery, or marker reset was performed. Startup may query public release metadata. The [retained smoke script](packages/06b8a89/published-smoke-script.txt) documents the exact scope.

## Validation boundary

The recorded checks apply to this source and package. The dependency audit is point-in-time evidence from configured feeds, not a comprehensive security assessment. The application remains **unsigned**; follow organizational policy and do not disable Windows security protections.

The previous non-reproduced Help minimum-size assertion remains part of the historical candidate record; it passed in all three fresh safety runs. Automated fixtures and this limited startup smoke do not replace every representative-data, provider-dependent EVTX, physical display, battery, or target-machine check. No signing credentials or Windows security policy were changed.

## Commit and upload provenance

The package is built from the committed shipping source before its checksums can be committed. A later documentation-only commit records the manifest/checksums/README, with a source-equivalence check proving that application, project, packaging and license inputs did not change after the build. The manifest's `SourceCommit` is the authoritative build revision; a subsequent documentation commit must not be mistaken for a different application build.

The [clean shipping source record](packages/06b8a89/shipping-source.json) and [validated input hashes](packages/06b8a89/validated-source-files.json) identify the build. The subsequent documentation-only commit has identical application/project/packaging/license inputs. Do not rebuild merely to attach those documentation changes to the executable's informational version; the recorded SourceCommit and artifact hashes remain authoritative.

After final validation and record updates, Beta and Release are pushed normally, atomically when supported. Their remote tips and the unchanged Beacon2.1 backup are verified. No force-push is used.

## GitHub Release upload

1. Create or update the stable **Beacon v2.2.0** release, using a numeric tag such as **`v2.2.0`** and the promoted Release source. No tag has been created by this task.
2. Upload **`artifacts/Beacon-2.2.0-Release/Beacon-2.2.0-win-x64.zip`**. This is the end-user download and includes the required notices.
3. Attach **`SHA256SUMS.txt`** and **`release-manifest.json`** from that same directory. Do not mix files from the older `artifacts/Beacon-2.2.0` Beta candidate.
4. Publish only after the maintainer's final release review. Source-branch promotion alone does not make the binary downloadable or change the stable-release API response.

The standalone `Beacon.exe` is in the same directory if needed for local use. When redistributing it separately, retain the notices supplied by the ZIP. Rebuilding, repackaging, or signing changes bytes and requires fresh verification values; never reuse an older checksum for changed artifacts.
