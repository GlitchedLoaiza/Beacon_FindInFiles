# Beacon release validation — Step 12

## Status

**Beacon Theme follow-up:** the current source includes a fourth, system-aware theme with muted-crimson/off-white buttons and coordinated red focus/selection accents. It is not in the previously published artifacts recorded here. Their checksums are preserved as historical evidence only. A theme-enabled release requires a fresh publish, hashes and release smoke checks; this change does not upload or replace an existing release.

**Current release preparation:** see [Beacon 2.1.0 preparation](../../docs/releases/2.1.0/PREPARATION.md) for the refreshed local artifacts, exact EXE/ZIP hashes, included notices, 71 passing Release regression groups, and later-promotion checklist. That package uses the existing Folder profile with ReadyToRun disabled and remains unsigned. No master promotion or GitHub release upload was performed; the historical candidate hashes and smoke results below must not be reused as verification for it.

**Superseded candidate:** the source now targets **2.1.0** and includes first-launch onboarding plus a splash version label. The 2.0.1 binary, checksum and published smoke results below are historical evidence only and do not validate the new 2.1 build. Rebuild/publish 2.1, regenerate its hash, repeat startup checks (first launch, skipped/exited tour and second launch), and complete the manual gates before shipping. Do not distribute the old candidate as 2.1.

**Later optimization validation:** [PERFORMANCE-AUDIT.md](PERFORMANCE-AUDIT.md) records the 2026-09-15 controlled A/B result, reversible production patch, 71 passing groups in Debug and Release, and a successful local 2.1 single-file publish. That optimized candidate has its own executable hash and is unsigned; it was not interactively smoke-tested in that pass. It does not inherit the historical 2.0.1 smoke-test sign-off below. The manual and distribution gates still apply to the exact final shipping binary.

**Automated validation and a limited published-executable smoke test passed. Distribution approval is pending the manual gates below.** This is not a claim that every archive, Windows installation, or provider configuration is supported without further testing. No release was uploaded, no app version was changed, and no signing operation was performed.

## Validated baseline

- Validation date: 2026-09-15 (development machine local date).
- Branch: `experiments-latest`.
- Source commit: `896fcb5021e2f6d12f8975dc4880dfb6ac5e74df`.
- Source tree was clean when validation began. This report and its README link were added afterward; they do not change the validated executable.
- Windows: `10.0.26200.0` (Windows 11).
- .NET SDK: `10.0.401`.
- Target: `net10.0-windows`, `win-x64`.
- App/file version: `2.0.1` / `2.0.1.0` (unchanged).
- Packages: SharpCompress `0.50.4`, 7-Zip.CommandLine `25.1.0`, WebView2 `1.0.2792.45`.

## Evidence and results

Local evidence directory, relative to the repository root:

`Beacon_FindInFiles/bin/Release/validation-20260915-104635/`

Build artifacts are local and may be removed by cleaning the workspace; the commands below allow validation to be repeated.

| Check | Result | Evidence |
| --- | --- | --- |
| Visual Studio solution build | Passed | Build tool reported success |
| Full regression runner in Release | 69 passed, 0 failed; exit 0 | `release-tests.log` |
| Direct/transitive vulnerability audit | No vulnerable packages reported for either project using configured feeds | `dependency-audit.log` |
| Resolved dependency inventory | Recorded | `resolved-packages.log` |
| Bundled CAB helper consistency | SHA-256 matches the pinned NuGet package's x64 helper | Hash below |
| Fresh Release rebuild and publish | Exit 0; no warning/error lines in captured log | `publish.log` |
| Candidate inventory, version, signing and checksum | Recorded | `artifact-manifest.json` |
| Published startup / duplicate launch / normal exit | Passed | `published-smoke.log` |

The vulnerability audit queried `https://packagefeedproxy.microsoft.io/nuget/v3/index.json` and the configured local SDK feed. A clean audit is a point-in-time result from those sources, not a guarantee against undisclosed vulnerabilities or a security audit of all application code.

### Candidate configuration

Explicitly validated: Release, self-contained, single-file, ReadyToRun enabled, trimming disabled. Existing project native-library extraction and bundle compression settings were retained. The saved `FolderProfile.pubxml` was **not** used: it sets ReadyToRun to false and publishes to a user-specific OneDrive directory. That profile remains unchanged; decide the shipping configuration before final sign-off and rerun validation if it differs.

### Candidate artifact

- File: `publish/Beacon.exe`.
- Size: **85,802,608 bytes**.
- Product version: `2.0.1+896fcb5021e2f6d12f8975dc4880dfb6ac5e74df`.
- SHA-256: `60944829B909B76784CADB5395F4D9C64763BEE8C9040668958E9041B38D042D`.
- Authenticode status: **NotSigned**.
- No loose SharpCompress DLL or runtime configuration was emitted.
- Three WebView2 XML documentation files were also emitted: `Microsoft.Web.WebView2.Core.xml`, `Microsoft.Web.WebView2.WinForms.xml`, `Microsoft.Web.WebView2.Wpf.xml`. They are not proof that managed runtime dependencies are unbundled. Review what is included in the shipping download rather than publishing the entire evidence directory.
- Bundled `7za.exe` and package helper SHA-256: `574BB90D17732F3CC4145FD4BA12D8F29B9D63400881C0E6FE3110C88B0485DE`.

Do not reuse this checksum after another build, version change, signing operation or binary modification. Generate the final release checksum from the exact distributed bytes.

### Published process smoke test

No Beacon process was running before the test. Only the newly launched candidate processes were controlled.

1. Launched the published `Beacon.exe`.
2. Observed the main title `Beacon: Find in Files` within **4.96 seconds**, confirming startup progressed beyond the splash. This is one observation, not a performance benchmark.
3. Launched a second copy: it exited with code **0** and the original stayed running.
4. Requested normal close through `CloseMainWindow`: the original exited with code **0**.
5. Compared persisted settings hashes before and after: unchanged.

This did not scan private files, invoke ownership recovery, exercise native save/clipboard dialogs or verify visual focus transfer. The normal startup update check may run; its live outcome was not used as a release test. Fake-response regression tests cover update outcomes independently.

## Repeatable commands

Run from the repository root in PowerShell. Check each native command's exit code; do not treat a subsequent successful command as proof that an earlier one passed.

```powershell
dotnet run --project tests/Beacon.SafetyChecks/Beacon.SafetyChecks.vbproj -c Release -- "$PWD/Beacon_FindInFiles/7za.exe"
if ($LASTEXITCODE -ne 0) { throw 'Regression tests failed' }

dotnet list Beacon_FindInFiles.slnx package --vulnerable --include-transitive
if ($LASTEXITCODE -ne 0) { throw 'Dependency audit failed' }

dotnet publish Beacon_FindInFiles/Beacon_FindInFiles.vbproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishReadyToRun=true -p:PublishTrimmed=false -t:Rebuild,Publish -o Beacon_FindInFiles/bin/Release/release-validation/publish
if ($LASTEXITCODE -ne 0) { throw 'Publish failed' }

Get-FileHash Beacon_FindInFiles/bin/Release/release-validation/publish/Beacon.exe -Algorithm SHA256
Get-AuthenticodeSignature Beacon_FindInFiles/bin/Release/release-validation/publish/Beacon.exe
```

The regression runner is a console executable, not a Test Explorer test assembly. Its 69 groups include many assertions and cover real archive fixtures, temporary pipelines, WPF templates, WebView2 rendering and headless service behavior. They do not substitute for the following checks on the exact shipping executable.

## Manual release gates — not yet signed off

Record tester, machine/Windows build, candidate hash, actual result and date for each item. Mark a failure as blocking until resolved and retested. Use authorized/sanitized fixtures; never alter original evidence or production permissions for testing.

- [ ] **Representative content:** Search real text, JSON, XML and HTML from folders and archives. Confirm previews/highlights, logical paths, count/search totals, partial notices, export contents and Unicode names. Include large files and unreadable/corrupt sources.
- [ ] **Archive variants:** Test customer ZIP64 and multipart RAR/7z sets, large solid archives, compressed TAR/PAX and nested CAB combinations. Compare expected names/counts/content. Reopen selected previews after scan completion. Do not disable checksums or byte/path limits to make a failure disappear.
- [ ] **Cancellation and lifetime:** Cancel during counting and scanning, reset and rescan, then close during extraction/export. Confirm no stale results, deleted-in-use preview files or leftover active helper processes. Verify retained temporary sources disappear after normal cleanup.
- [ ] **Native EVTX:** Use authorized real EVTX files with multiple events, missing providers and portable LocaleMetaData. Verify filters, XML, multi-record collection and warning explanations. Synthetic event tests do not certify source-machine provider rendering.
- [ ] **HAR and privacy:** Test real sanitized HAR variants, decoded bodies, invalid encodings, redaction on/off and sensitive-only matches. Inspect both preview and exported report. Redaction is not anonymization; query text, paths and other fields can remain sensitive.
- [ ] **Windows UX:** Check Light/Dark/System, native caption buttons, snapping, dragging, maximized/minimized restoration, high contrast, keyboard navigation, Help and regex link, date pickers, side panes and popup placement at different DPI scales and across monitors. Reduced-motion behavior must remain usable.
- [ ] **Native dialogs and clipboard:** Folder/archive selection, HTML Save/overwrite/cancel, copy paths/XML/diagnostics, opening the report in an external browser. Verify cancellation preserves existing destination files and sources cannot be overwritten.
- [ ] **Single instance:** Rapid double launch during splash, during a scan and with Settings/Help open. Confirm no extra main window and acceptable activation behavior. Foreground focus is still subject to Windows rules.
- [ ] **Target-machine runtime:** Test on intended Windows x64 machines, including one without the .NET SDK. Test with WebView2 present and absent and under relevant endpoint-security/AppLocker restrictions. A self-contained .NET executable does not bundle the WebView2 browser runtime. Check older supported Windows behavior and user-writable temporary-directory requirements.
- [ ] **Access-denied prompts:** Validate skip/report and optional ownership confirmation only on a disposable test folder. Startup remains `asInvoker`; administrator approval should be requested only for the explicit operation, not silently at app launch.
- [ ] **Version/update contract:** Assign the intended release version consistently in Version, AssemblyVersion and FileVersion. The current candidate remains 2.0.1 and should not masquerade as a new version. Confirm the GitHub release tag/title is parseable by the checker. Version changes require rebuilding and regenerating evidence/hash.
- [ ] **Signing and reputation:** Decide whether to Authenticode-sign the final executable and timestamp it. No certificate was selected and this candidate is unsigned; do not promise that SmartScreen will accept it without prompts. Hash after signing.
- [ ] **Redistribution/legal review:** Include appropriate notices/source-information obligations for the bundled 7-Zip components and SharpCompress/WebView2 licenses. The raw publish output contains no third-party notice file; the RAR fixture notice under tests is not an application distribution notice. The pinned 7-Zip package includes `tools/License.txt` and `tools/readme.txt`; review official terms before packaging. Do not ship test archives or evidence logs as application content.
- [ ] **Documentation and release assets:** Have a first-time user follow Help, review known limitations, choose the correct publish configuration, and prepare release notes and exact download checksum. Review repository license/ownership requirements with the release owner. No commit, push, upload, tag or release creation is authorized by this validation pass.

## Sign-off

Automated evidence: **PASS** for the baseline and configuration above.

Manual functional/environment sign-off: **PENDING**.

Version, signing, notices and distribution approval: **PENDING**.

Release publication: **NOT PERFORMED**.
