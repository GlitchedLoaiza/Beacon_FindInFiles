# Beacon 2.1.0 — source promotion and release preparation

## Current status

**Source promotion is separate from binary publication.** The current 2.1.0 source includes Beacon Theme. This review covers the documentation and a history-preserving promotion from `experiments-latest` to **local `master` only**, retaining `Beacon2.0.1`. A remote push, tag, signing operation, or GitHub release upload is not included.

**Archived package:** the EXE/ZIP hashes and 71-group package validation below identify the exact pre-theme artifact, not a build of current source. Rebuild, rerun release checks, smoke-test the final package, and regenerate the checksum/manifest/README table before distributing the theme-enabled application. Promoting source does not replace those artifacts or approve distribution.

**Published-release snapshot (2026-09-16 UTC):** GitHub's latest stable release is `PublicRelease`, titled `Beacon v2.0.1`, with `Beacon.v2.0.1.zip` as its only asset. It does not include a separate `SHA256SUMS.txt` asset. Recheck the release page before distribution; the archived 2.1.0 checksums below do not verify that v2.0.1 ZIP.

## Source promotion validation

- Application/test source reviewed: `20edcaa4ee10ca0997023744d712aa73564624a6` (`experiments-latest`). This review changes documentation only, not application code or dependencies.
- Local `master` and `Beacon2.0.1` at review start: `2115a9eea2624a0380c92c1e0a368a5b0c989345`.
- At review start, `master...experiments-latest` contains **1 master-only / 8 experiments-only commits**. The only master-only change is the v2.0.1 README checksum, preserved in `Beacon2.0.1`. Use a normal merge and retain the reviewed 2.1 README if that historical checksum conflicts.

**Source validation (2026-09-16 UTC): PASS, with the benchmark-only warning below.** Environment: Windows 11 (`10.0.26200.0`), .NET SDK `10.0.401`, Visual Studio Professional 2026. Both regression runs include the Beacon Theme checks.

| Check | Result |
| --- | --- |
| Full Visual Studio solution build, Debug/Any CPU | 3 projects succeeded, 0 failed; `MSB3270` warning in `BenchmarkSuite1` |
| Debug safety runner | 72 passed, 0 failed; exit code 0 |
| Release safety runner | 72 passed, 0 failed; exit code 0 |
| Documentation consistency | 11 tracked Markdown files and 31 local file/directory/heading links checked; README structure preserved |
| Archived artifact references | EXE/ZIP hashes and sizes match the unchanged checksum file and manifest |
| Packaging script syntax | Passed; no new package was built or distributed during this review |

`MSB3270` reports that the development-only `BenchmarkSuite1` MSIL/Any CPU host references AMD64 `Beacon.dll`. This existing warning is accepted for the documentation-only source promotion, not as validation of benchmark execution. Application code, dependencies, and benchmark settings remain unchanged by this review.

Raw regression logs are retained locally under the ignored directory `artifacts/source-promotion-20260916-164624-44792878/` as `debug-regressions.txt` and `release-regressions.txt`. They are not repository release assets and may be removed by workspace cleanup. Re-run the checks from the repository root when validating a different source snapshot:

```powershell
dotnet run --project tests/Beacon.SafetyChecks/Beacon.SafetyChecks.vbproj -c Debug -- "$PWD/Beacon_FindInFiles/7za.exe"
if ($LASTEXITCODE -ne 0) { throw 'Debug safety checks failed' }
dotnet run --project tests/Beacon.SafetyChecks/Beacon.SafetyChecks.vbproj -c Release -- "$PWD/Beacon_FindInFiles/7za.exe"
if ($LASTEXITCODE -ne 0) { throw 'Release safety checks failed' }
```

No new dependency audit, published-binary smoke test, signing, or distribution approval is claimed by these source checks. The historical package logs below do not substitute for release validation of the exact theme-enabled shipping artifact.

## Historical preparation status

The original preparation on **2026-09-16 UTC** did not merge, check out master, commit, push, tag, or upload a GitHub release. Its README/changelog, packaging script, license notices, and verification records were initially left uncommitted and have since been committed. Application code and dependency versions were not changed during that preparation.

That package superseded an earlier local EXE-only 2.1.0 ZIP, not the public v2.0.1 download. Historical performance/Step 12 records remain evidence for their own artifacts, not this package or the current source. A valid hash and passing tests are not release approval or a verified publisher identity.

## Branch boundaries at the original preparation

| Reference | Commit at the original preparation |
| --- | --- |
| `experiments-latest` | `90f960dd88e71679c31c2da1c4e0579ebae60dac` |
| `master` | `2115a9eea2624a0380c92c1e0a368a5b0c989345` |
| `Beacon2.0.1` | `2115a9eea2624a0380c92c1e0a368a5b0c989345` |

At that snapshot, local `master...experiments-latest` showed **1 master-only / 5 experiments-only commits**. `Beacon2.0.1` preserved the original master's committed contents and had already been pushed; no remote refs were modified or fetched by that preparation. Use the current branch checks above rather than this historical divergence count for promotion.

## Historical preparation changes

- Removed an ignored accidental root file containing ANSI-colored output from an old Git branch-list command, plus its one-off ignore entry. It was not application source.
- Confirmed no generated build outputs or user-specific files are tracked. Required `7za.exe`, source assets, pinned RAR fixtures, regressions, `BenchmarkSuite1`, and all performance/rollback evidence remain.
- Left `.vs`, bin/obj caches, user-specific `.vbproj.user` settings, and the ignored OneDrive publish profile untouched. Broad `git clean`, resets, and deletion of evidence were not used.
- Rewrote the root README for actual 2.1 controls, features, defaults, limits, unsigned SmartScreen behavior, build/package instructions and verification. Added the v2.1.0 changelog; retained v2.0.1 historical notes and original contributor credits.
- Moved Beacon's existing MIT terms into a standalone `LICENSE` without changing attribution/year. Added third-party notice/source information and exact upstream license text. Original dependency credits remain intact.
- Added `scripts/Publish-Beacon.ps1` to publish to fresh staging, gather matching runtime/dependency notices, verify ZIP contents, back up old artifacts, and generate checksums/manifest. No signing credentials, production settings, customer files, or benchmark/test payloads are included in the release.
- Added narrowly scoped `.gitignore` exceptions for these release records. Build output and raw profiler artifacts remain ignored.

## Historical package validation

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

## Archived pre-theme artifacts

Prepared in a maintainer-local release directory. That machine-specific path is not required for building or packaging; use a dedicated output directory with `scripts/Publish-Beacon.ps1 -Destination`.

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

## Promotion and publication checklist

1. Review the pending changes with `git status --short` and `git diff`. Review the license/source-information obligations and current known limitations; verify no customer data, signing material, or generated binaries are staged. The release EXE/ZIP stay outside Git; only checksums, manifests, code and documentation belong in this change.
2. Commit the reviewed documentation on `experiments-latest` after the source-validation checks pass. Build the exact intended release commit again before packaging for final binary provenance.
3. Inspect the master-only change and branch history before merging. Use a normal, history-preserving merge; do not reset or force-push master to make the difference disappear. Preserve the `Beacon2.0.1` archive branch.
4. For this source promotion, merge into **local `master` only**. Verify that `git diff master experiments-latest` is empty, the working tree is clean, and the backup and remote-tracking refs are unchanged. Do not push, tag, or publish without separate authorization.
5. On the exact source/build to distribute, rerun the solution build, Release regressions, dependency audit and package script. A different SDK, commit, publish profile, runtime pack, signing operation or ZIP timestamp can change the hashes. **After any rebuild/repackage, replace the README hash table, SHA256SUMS, manifest and release notes together.** The hashes here are not a promise of byte-reproducible builds.
6. Smoke-test the final packaged app on intended Windows machines: startup/tour opt-out and second launch, all four themes (including Beacon Theme in Windows light and dark modes), native window behavior, WebView2 fallback, normal/large/nested archives, real EVTX/provider metadata, HAR redaction both ways, export/save/clipboard, cancellation/reset/close and temporary cleanup. Use the existing [manual release gates](../../../tests/Beacon.SafetyChecks/RELEASE-VALIDATION.md) and record actual results rather than marking unknown cases passed.
7. If the maintainer later chooses signing, sign/timestamp the final EXE **before** ZIP creation and regenerate all hashes. The script currently packages the unsigned publish output; it is not a signing pipeline. Do not describe the branding metadata as verified publisher identity.
8. Upload the exact final ZIP with its legal notices, `SHA256SUMS.txt`, and `release-manifest.json` to a versioned GitHub release only after approval. The update checker accepts a numeric tag such as `v2.1.0`, or a title such as `Beacon v2.1.0` for the existing `PublicRelease` tag. Do not quietly replace an older version's bytes or reuse obsolete checksums.
9. Update the README's release status and verification table to match the exact published artifacts. A source promotion alone must not be described as a new downloadable release.

## Rollback and boundaries

This documentation review does not change application behavior; the promoted branch intentionally retains the existing 2.1 implementation. For a documentation rollback, selectively revert the documentation commit. The original packaging preparation can be reverted separately through its README/.gitignore hunks and release-support files; do not reset the whole repository or remove regression/performance evidence. If replacing local artifacts must be undone, restore the **original EXE and ZIP together** from the pre-preparation backup and remove/regenerate the corresponding checksum/manifest rather than leaving mismatched verification information.

The measured optimization rollback remains separately documented in [PERFORMANCE-AUDIT.md](../../../tests/Beacon.SafetyChecks/PERFORMANCE-AUDIT.md). Do not use that production patch to undo this documentation-only preparation.

**Remote push, GitHub upload, and signing: NOT PERFORMED by this source-promotion review. Manual/binary-distribution sign-off: PENDING.**
