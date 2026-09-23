# Beacon 2.2.0 — Beta to Release preparation

> **Historical Beta-candidate preparation.** This record and its original hashes are retained for provenance. See [PROMOTION.md](PROMOTION.md) for the later authorized Release promotion and shipping package; do not substitute these candidate bytes for the final upload.

## Current status

**Locally packaged and automatically validated; ready for owner review and branch reconciliation, not yet promoted or published.** Preparation remains on Beta. No commit, merge, branch switch, push, tag, signing, or GitHub release upload was performed. Existing implementation changes and historical artifacts are preserved.

The application version is **2.2.0**, with assembly/file version **2.2.0.0**. The splash and update checker use assembly-derived version information. The .NET 10 target and package versions are unchanged. This release includes automatic multicore searching, plain-text Previous/Next navigation and match-position counts, captured-context Summary, marker-preserving tour replay, readable Beacon highlights, and complete navigation cleanup after Reset.

The packaging/validation sections below describe the new candidate's actual results. Old 2.1.0 checksums, the earlier 89-check benchmark milestone, and its source snapshot do not identify or validate this 2.2.0 binary.

## Branch boundaries

At the start of preparation, local refs and a read-only `git ls-remote` check agreed:

| Reference | Commit |
| --- | --- |
| `Beta` / `origin/Beta` | `20edcaa4ee10ca0997023744d712aa73564624a6` |
| `Release` / `origin/Release` | `9b3932298096d74a1a530833b9a94bca9fa476d8` |

`git rev-list --left-right --count Release...Beta` returned **5 Release-only / 0 Beta-only commits**. The new implementation is currently in the Beta working tree, including required untracked production and test files; the committed Beta tip alone does not contain those features.

The Release-only differences concern README/release documentation, the 2.1.0 package-history records, and their `.gitignore` exceptions. No committed application, packaging-script, license, or dependency-source differences were found between these branch tips. Preserve `docs/releases/2.1.0/packages/7dbb342/` and the corresponding history on Release. Do not replace Release wholesale with the older committed Beta tree.

Once the pending 2.2 work is committed on Beta, history will diverge unless Release has first been reconciled. README and maintainer release-guide edits may conflict. Preserve the 2.2 user guide and new artifact records while retaining historical 2.1 package evidence. No fetch or remote-ref update was needed for the read-only inspection; repeat the remote comparison immediately before promotion.

### Non-destructive promotion preflight

A three-way **file-level preview** against the common base and Release was run on normalized temporary copies with `git merge-file --stdout --diff3`. It did not edit the real working files, index, or refs. The [preflight record](validation/promotion-preflight.json) identifies these manual reconciliation points:

| File | Conflict regions in preview | Resolution to preserve |
| --- | ---: | --- |
| `.gitignore` | 1 | Keep both Release's 2.1.0 package-history exceptions and the new narrow 2.2.0 record exceptions. |
| `README.md` | 4 | Keep the user-friendly 2.2 guide and exact 2.2 candidate hashes; retain earlier release notes and download-verification guidance. |
| `tests/Beacon.SafetyChecks/README.md` | 2 | Keep current 2.2 validation pointers and Release's historical package records. |
| `tests/Beacon.SafetyChecks/RELEASE-VALIDATION.md` | 2 | Keep the 2.2 preparation pointer while preserving the earlier release and smoke evidence with its original scope. |

The Release-only `docs/releases/2.1.0/packages/7dbb342/` records must remain available after promotion. This preview is not a merge approval or a substitute for repeating the real merge after committing and refreshing remote refs. Raw previews are retained under ignored package staging, not inserted into documentation as conflict markers.

## Local candidate

The packaging destination is the workspace-local, ignored directory:

`artifacts/Beacon-2.2.0/`

Do not use or overwrite the personal OneDrive publish destination. Reuse `scripts/Publish-Beacon.ps1` with its default **ReadyToRun off**, Release, win-x64, self-contained, single-file, trimming-disabled configuration. The ZIP must include Beacon's license and matching third-party notices. The executable remains unsigned unless a later, separately approved signing step is performed.

Built **2026-09-23T17:19:36.6256147Z** from the Beta working tree based on `20edcaa4ee10ca0997023744d712aa73564624a6`. The manifest explicitly records **uncommitted source changes**. Company/publisher branding remains GlitchedLoaiza; Authenticode status is **NotSigned**.

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| `Beacon.exe` | 77,655,849 | `CC7BFFDFAB568F10888129740B9639626081CB6D6BF13BD82EB6CBEC3D098E60` |
| `Beacon-2.2.0-win-x64.zip` | 72,102,795 | `D6E9E8FA02A608B8CA28DE423389F7737DB290F6E60259BDBB7C061F05AF277A` |

Exact copies of the candidate's [SHA256SUMS.txt](SHA256SUMS.txt) and [release-manifest.json](release-manifest.json) are retained here. The raw publish contains only `Beacon.exe`; the ZIP contains **12 verified files**: the EXE, Beacon's license, the notice index, and component notices under `licenses/`. Every ZIP entry was independently hashed against the staging bundle; see [package-contents.json](validation/package-contents.json).

The included .NET and Windows Desktop runtimes are **10.0.12**. Packages remain SharpCompress **0.50.4**, 7-Zip.CommandLine **25.1.0**, and WebView2 SDK **1.0.2792.45**. The installed WebView2 browser runtime is not included in the ZIP. [Package verification](validation/package-validation.txt), [publish properties](validation/publish-properties.json), [publish output](validation/publish.txt), and [staging locations](validation/package-location.json) record the exact build.

The manifest must disclose uncommitted source changes. Commit metadata alone does not identify the candidate: exact hashes identify the packaged bytes. Rebuilding, repackaging, signing, or rebuilding after a merge requires new checksums and validation. A locally prepared candidate is not evidence that a 2.2 download is already published.

## Automated validation

| Check | Result | Evidence |
| --- | --- | --- |
| Workspace build | Passed | Visual Studio build result |
| Debug solution build | Passed; 0 warnings, 0 errors | [debug-build.txt](validation/debug-build.txt) |
| Release solution build | Passed; 0 warnings, 0 errors | [release-build.txt](validation/release-build.txt) |
| Debug STA safety suite | **93 passed; 0 failed** | [debug-regressions.txt](validation/debug-regressions.txt) |
| Release STA safety suite | **93 passed; 0 failed** | [release-regressions.txt](validation/release-regressions.txt) |
| Direct/transitive dependency audit | No known vulnerable packages reported for all three projects using configured feeds | [dependency-audit.json](validation/dependency-audit.json) |
| Version metadata | Package 2.2.0; assembly/file 2.2.0.0; packaged product 2.2.0 plus recorded source revision | [source-version-properties.json](validation/source-version-properties.json), [package-validation.txt](validation/package-validation.txt) |
| Publish/package integrity | Passed; 12 ZIP entries, matching notices/helper, exact EXE/ZIP hashes | [package-contents.json](validation/package-contents.json), [publish.txt](validation/publish.txt) |
| Documentation and source consistency | README artifact values and local links verified; packaged source hashes still match; whitespace check passed | [documentation-links.json](validation/documentation-links.json), [source-state-after.json](validation/source-state-after.json) |
| Promotion boundary | Beta/Release and tracking refs unchanged; four documentation overlaps require reconciliation | [promotion-preflight.json](validation/promotion-preflight.json), [source-state-after.json](validation/source-state-after.json) |

The audit is point-in-time evidence from the configured Microsoft package feed proxy and local SDK feed, not a comprehensive security review. The benchmark host's processor target was aligned to x64 to resolve MSB3270 against the x64 application; no warning suppression, dependency upgrade, or application behavior change was used.

The first Debug run reported **92 passed / 1 failed** at the existing Help minimum-size/regex-link assertion. Its [original log](validation/initial-debug-regressions.txt) is retained. The assertion was given dimensions/visibility diagnostics without changing its acceptance threshold, and the failure did not recur in the final Debug or Release runs. No Help layout was changed during this preparation. Treat this as a non-reproduced observation, not a proven layout fix; explicitly review Help at minimum size in the exact-candidate manual smoke test.

[Source-state evidence](validation/source-state-before.json), [source-file hashes](validation/source-files.json), and the [final state check](validation/source-state-after.json) identify the reviewed working tree. No historical artifact records were overwritten. The index and branch refs were not changed by the merge preview or packaging.

Validation output is retained under `validation/`. Previous performance findings remain in [the production benchmark report](../../../tests/Beacon.SafetyChecks/PerformanceEvidence/multicore-production/REPORT.md); those timings apply to their recorded environment and workload, not a new performance run of the version-bumped executable.

## Owner-controlled promotion checklist

These are **future actions**, not commands already executed:

1. Review `git status --short` and the complete implementation diff on Beta. Include all required new production files, tests, benchmark helpers, release documents, and intended notices. Do not commit bin/obj output, private samples, local profiles, generated application binaries, or unrelated files merely because they appear in the workspace.
2. Stage only reviewed files, inspect `git diff --cached` and `git diff --cached --check`, then commit the intended 2.2 work on Beta.
3. Refresh remote refs with `git fetch origin`. Reconcile `origin/Release` into Beta without force-resetting or overwriting either branch. If there are documentation conflicts, retain the 2.2 instructions plus the existing Release-side 2.1 package history. Keep the versioned ignore exceptions for both releases.
4. Rebuild, rerun the complete safety suite, and package from the reconciled source. Regenerate the checksums, manifest, README table, and release evidence if any shipping bytes changed. Check that the source/commit provenance accurately describes that package.
5. Complete the manual gates below and obtain release approval. Prefer a reviewed pull request **Beta → Release**. For an approved local merge instead, first synchronize Release with `origin/Release`, then merge the reviewed Beta commit normally; do not force-push or discard Release history.
6. Push only the reviewed commits/merge when authorized. Verify the remote Release tip and CI/manual validation outcomes.
7. Use a stable numeric tag such as **`v2.2.0`** for the approved Release commit, after checking that the tag does not already exist. Publish the exact approved ZIP, `SHA256SUMS.txt`, and manifest as release assets. Do not reuse or silently replace the bytes/hashes of an older release.
8. Update the README's preparation banner only when the corresponding release is actually published. If signing is added, verify the signing identity and regenerate hashes from the signed artifacts.

## Manual gates before distribution

- [ ] Smoke-test the **exact packaged candidate** on a target Windows x64 machine: startup, displayed version, duplicate-launch behavior, and normal close.
- [ ] Use authorized representative files to exercise scanning, count/progress, cancellation/reset, nested archive previews, Previous/Next, match counters, Full/Summary, and all themes.
- [ ] Verify Help's minimum-size reading area and regex documentation link in light/dark themes, including the non-reproduced assertion noted above.
- [ ] Replay the tour from Help and confirm the once-only startup state and current search remain intact; do not delete the user's real marker just to test this.
- [ ] Check WebView2 availability/fallback, real EVTX provider behavior, HAR privacy choices, and applicable workplace policy.
- [ ] Review the unsigned SmartScreen behavior, license/notice redistribution, download verification instructions, and final release assets.
- [ ] Record sign-off and the final shipping commit/hash after reconciliation. Passing automated tests and valid checksums are not release approval or a guarantee of a bug-free application.

No interactive packaged-application smoke test is assumed or inherited from an earlier release. Preparation must not modify the real user's settings, onboarding marker, credentials, security policy, or running Beacon instance.
