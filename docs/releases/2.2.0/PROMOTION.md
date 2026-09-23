# Beacon 2.2.0 — Release promotion and upload handoff

## Status

The maintainer has authorized promoting Beta to the remote Release branch and preparing the end-user package. History reconciliation is in progress; validation, final package provenance, smoke results, and remote verification will be recorded here before completion. No GitHub Release asset upload or tag creation is performed by this task.

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

The expected deliverables are `Beacon-2.2.0-win-x64.zip`, `Beacon.exe`, `SHA256SUMS.txt`, and `release-manifest.json`. The ZIP includes the required legal notices. Exact sizes, hashes, source commit, and validation evidence are populated only after packaging and verification.

## Validation boundary

The reconciled source must pass workspace/Debug/Release builds, the complete STA safety executable, and a point-in-time dependency audit. Package verification checks versions, raw publish contents, notices, helper bytes, archive inventory and hashes. A limited exact-binary smoke checks startup, second-instance behavior, and normal close while preserving existing settings/onboarding state and leaving unrelated processes alone.

The previous non-reproduced Help minimum-size assertion remains part of the historical candidate record. Do not imply that automated fixtures or a limited startup smoke replace all representative-data, provider-dependent EVTX, display, battery, or target-machine testing. No signing credentials or Windows security policy are changed.

## Commit and upload provenance

The package is built from the committed shipping source before its checksums can be committed. A later documentation-only commit records the manifest/checksums/README, with a source-equivalence check proving that application, project, packaging and license inputs did not change after the build. The manifest's `SourceCommit` is the authoritative build revision; a subsequent documentation commit must not be mistaken for a different application build.

After final validation and record updates, Beta and Release are pushed normally, atomically when supported. Their remote tips and the unchanged Beacon2.1 backup are verified. The maintainer can then create/update the GitHub release (for example with tag `v2.2.0`) and attach the exact ZIP plus checksum/manifest files. Rebuilding, repackaging, or signing changes bytes and requires fresh verification values.
