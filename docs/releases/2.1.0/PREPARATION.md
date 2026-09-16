# Beacon 2.1.0 — preparation for later master promotion

## Current status

**Prepared for review; no merge, checkout of master, commit, push, tag, or GitHub release upload was performed.** Application code and dependency versions were not changed during this preparation. The new README/changelog, packaging script, license notices, and verification records are intentionally left uncommitted on `experiments-latest`.

This record supersedes the old README download hash and the previous local EXE-only 2.1.0 ZIP. Historical performance/Step 12 records are preserved and still describe their own artifacts, not this package. A valid hash and passing tests are not release approval or a verified publisher identity.

## Branch boundaries recorded

| Reference | Commit at start and verification |
| --- | --- |
| `experiments-latest` | `90f960dd88e71679c31c2da1c4e0579ebae60dac` |
| `master` | `2115a9eea2624a0380c92c1e0a368a5b0c989345` |
| `Beacon2.0.1` | `2115a9eea2624a0380c92c1e0a368a5b0c989345` |

Local `master...experiments-latest` shows **1 master-only / 5 experiments-only commits** at this snapshot. Do not assume a fast-forward or overwrite master blindly. `Beacon2.0.1` preserves the original master's committed contents and was previously pushed; no remote refs were modified or fetched by this preparation.

## Cleanup and changes

- Removed an ignored accidental root file containing ANSI-colored output from an old Git branch-list command, plus its one-off ignore entry. It was not application source.
- Confirmed no generated build outputs or user-specific files are tracked. Required `7za.exe`, source assets, pinned RAR fixtures, regressions, `BenchmarkSuite1`, and all performance/rollback evidence remain.
- Left `.vs`, bin/obj caches, user-specific `.vbproj.user` settings, and the ignored OneDrive publish profile untouched. Broad `git clean`, resets, and deletion of evidence were not used.
- Rewrote the root README for actual 2.1 controls, features, defaults, limits, unsigned SmartScreen behavior, build/package instructions and verification. Added the v2.1.0 changelog; retained v2.0.1 historical notes and original contributor credits.
- Moved Beacon's existing MIT terms into a standalone `LICENSE` without changing attribution/year. Added third-party notice/source information and exact upstream license text. Original dependency credits remain intact.
- Added `scripts/Publish-Beacon.ps1` to publish to fresh staging, gather matching runtime/dependency notices, verify ZIP contents, back up old artifacts, and generate checksums/manifest. No signing credentials, production settings, customer files, or benchmark/test payloads are included in the release.
- Added narrowly scoped `.gitignore` exceptions for these release records. Build output and raw profiler artifacts remain ignored.

## Validation for this preparation

- Full Visual Studio solution build: **passed**.
- Release regression runner: **71 passed; 0 failed**. See [validation/release-regressions.txt](validation/release-regressions.txt).
- Direct/transitive NuGet vulnerability audit: **no known vulnerable packages** reported for application, safety runner, and benchmark host using configured feeds. See [validation/dependency-audit.txt](validation/dependency-audit.txt). This is point-in-time feed evidence, not a comprehensive security audit.
- Packaging script syntax and publish-profile discovery passed.
- Fresh production publish: **passed**, raw output contains only `Beacon.exe`. The three WebView2 API-documentation XML files are excluded.
- The ZIP contains **12 files**, each independently hashed and compared with the staging bundle; no traversal/duplicate entries or unrelated artifacts were found.
- Bundled 7za bytes match the pinned package's x64 helper.
- Publisher/company branding is **GlitchedLoaiza**, file version **2.1.0.0**, product version **2.1.0+90f960dd88e71679c31c2da1c4e0579ebae60dac**. Authenticode status is **NotSigned**.
- No smoke test of this exact published build was performed during this preparation. The earlier smoke results do not automatically apply to it. Real customer logs, signing and final distribution approval remain owner checks.

The first packaging validation caught both checksum records on one line. The script was corrected to write and validate exactly one line per artifact, then publish/package was rerun successfully. That first attempt is not the final ZIP described here; the final hashes below and the final two-line checksum file are authoritative.

## Exact artifacts

Published to the previously requested local folder:

`C:\Users\luislo\OneDrive - Microsoft\Documents 1\Tools\Beacon`

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| `Beacon.exe` | 77,624,725 | `C35F908A02A52E361B58E82E31E00B531864C31CD7F8CA420D5A3CE68E97A686` |
| `Beacon-2.1.0-win-x64.zip` | 72,071,577 | `EDDF545A27265041DAD7862DFDF2C3A0F069324C424AB417DBE5CC8F0E786BAB` |

- [SHA256SUMS.txt](SHA256SUMS.txt) and [release-manifest.json](release-manifest.json) are exact copies of the final local output, not hashes of an older candidate.
- Build time: **2026-09-16T01:14:29.9313109Z** (2026-09-15 evening on the development machine).
- Configuration: Release, .NET 10 Windows x64, self-contained, single-file, trimming disabled, **ReadyToRun disabled** by the saved Folder profile. .NET and Windows Desktop runtime packs: **10.0.12**.
- Packages: SharpCompress **0.50.4**, 7-Zip.CommandLine **25.1.0**, WebView2 SDK **1.0.2792.45**. The installed WebView2 browser runtime is not distributed in the ZIP.
- ZIP contains the EXE, Beacon license, notice index, and component license files under `licenses/`. Those small legal files are intentional, unlike the unnecessary WebView2 API documentation removed from raw publishing.
- Source was unchanged under `Beacon_FindInFiles/`; the manifest flags uncommitted changes because documentation, notices, and the packaging script were being prepared. Commit metadata alone does not identify ZIP bytes.

Staging/logs and backups are retained under ignored build output:

- Final stage: `Beacon_FindInFiles/bin/Release/package-20260915-191414-3d28935f/`.
- Pre-preparation EXE/ZIP backup: `Beacon_FindInFiles/bin/Release/package-20260915-191058-03c7bf3d/previous-release/`.
- The final stage also backs up the immediately preceding packaging attempt. Do not confuse it with the original published pair.
- Final publish logs/property evidence: [validation/publish.txt](validation/publish.txt), [validation/publish-properties.json](validation/publish-properties.json).

The previous user-provided ZIP and EXE were backed up rather than discarded. Do not delete those backups until the refreshed artifacts have been reviewed. The archived application remains on `Beacon2.0.1`; no binary rollback or branch rollback was performed.

## Later promotion checklist — instructions only

1. Review the pending changes with `git status --short` and `git diff`. Review the license/source-information obligations and current known limitations; verify no customer data, signing material, or generated binaries are staged. The release EXE/ZIP stay outside Git; only checksums, manifests, code and documentation belong in this change.
2. Commit the reviewed preparation on `experiments-latest` when authorized. This pass intentionally did **not** stage, commit or push anything. A future clean build from the intended release commit is preferable for final provenance.
3. Inspect the master-only change and branch history before choosing the final merge strategy. Do not reset or force-push master to make the difference disappear. Preserve the `Beacon2.0.1` archive branch.
4. Open a reviewed pull request from `experiments-latest` to `master` when authorized. Resolve any conflicts by preserving the current intended 2.1 implementation and correct release-verification information. This document does not authorize executing the move now.
5. On the exact source/build to distribute, rerun the solution build, Release regressions, dependency audit and package script. A different SDK, commit, publish profile, runtime pack, signing operation or ZIP timestamp can change the hashes. **After any rebuild/repackage, replace the README hash table, SHA256SUMS, manifest and release notes together.** The hashes here are not a promise of byte-reproducible builds.
6. Smoke-test the final packaged app on intended Windows machines: startup/tour opt-out and second launch, native window behavior, WebView2 fallback, normal/large/nested archives, real EVTX/provider metadata, HAR redaction both ways, export/save/clipboard, cancellation/reset/close and temporary cleanup. Use the existing [manual release gates](../../../tests/Beacon.SafetyChecks/RELEASE-VALIDATION.md) and record actual results rather than marking unknown cases passed.
7. If the maintainer later chooses signing, sign/timestamp the final EXE **before** ZIP creation and regenerate all hashes. The script currently packages the unsigned publish output; it is not a signing pipeline. Do not describe the branding metadata as verified publisher identity.
8. Upload the exact final ZIP with its legal notices, `SHA256SUMS.txt`, and `release-manifest.json` to a versioned GitHub release only after approval. The update checker accepts a numeric tag such as `v2.1.0`, or a title such as `Beacon v2.1.0` for the existing `PublicRelease` tag. Do not quietly replace an older version's bytes or reuse obsolete checksums.
9. After the approved promotion/publication, remove the provisional “prepared for release” wording from README only when it truthfully describes the public state.

## Rollback and boundaries

No application behavior changed in this preparation. For a documentation/packaging rollback, selectively revert the new README/.gitignore hunks and the new release-support files; do not reset the whole repository or remove regression/performance evidence. If replacing local artifacts must be undone, restore the **original EXE and ZIP together** from the pre-preparation backup and remove/regenerate the corresponding checksum/manifest rather than leaving mismatched verification information.

The measured optimization rollback remains separately documented in [PERFORMANCE-AUDIT.md](../../../tests/Beacon.SafetyChecks/PERFORMANCE-AUDIT.md). Do not use that production patch to undo this documentation-only preparation.

**Master promotion: NOT PERFORMED. GitHub upload: NOT PERFORMED. Signing: NOT PERFORMED. Manual/distribution sign-off: PENDING.**
