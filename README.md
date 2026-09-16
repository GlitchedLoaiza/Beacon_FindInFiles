# Beacon: Find in Files 2.1

![Version](https://img.shields.io/badge/version-2.1.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20x64-lightgrey)
![Framework](https://img.shields.io/badge/.NET-10-purple)
![License](https://img.shields.io/badge/license-MIT-green)

**Find information in logs without opening every file yourself.** Beacon searches text files, Windows event logs (EVTX), saved web requests (HAR), and supported archives. Choose a source, search for words or a pattern, inspect matches, and export a readable report.

Maintained by **GlitchedLoaiza**. This branch describes **2.1.0 prepared for release**; it does not mean the branch has been merged into `master` or that the 2.1.0 download has already been uploaded. The previous `master` contents are preserved on [`Beacon2.0.1`](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/tree/Beacon2.0.1).

## Download and run

1. Open the official [GitHub Releases](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/releases) page and select the version you want.
2. Download its ZIP and matching `SHA256SUMS.txt`. For the prepared 2.1.0 package, the ZIP is named `Beacon-2.1.0-win-x64.zip`.
3. Verify the download as described below, then extract the ZIP to a folder you can write to.
4. Run **`Beacon.exe`**. No installer or separate .NET installation is required for the published self-contained build.

The release ZIP contains `Beacon.exe`, Beacon's license, and third-party license notices. Keep those notices when redistributing it. The benchmark runner, regression fixtures, and WebView2 API-documentation XML files are not application dependencies and are not included in the release ZIP.

### Requirements

- **Windows 11 x64** is the primary validated environment. Other Windows versions and managed-device policies require their own testing; native caption coloring falls back where unsupported.
- **.NET 10 is bundled** in the published EXE. Building from source requires the .NET 10 SDK.
- **Microsoft Edge WebView2 Runtime** is used for complete HTML/XML/JSON previews. It is often already installed on Windows but is not bundled with Beacon. A text fallback is available if initialization fails.
- A writable user profile and temporary directory are needed for settings and bounded archive extraction. Security policies must allow the bundled CAB helper to run.

### Verify the exact downloaded file

The prepared 2.1.0 [SHA-256 list](docs/releases/2.1.0/SHA256SUMS.txt) and [artifact manifest](docs/releases/2.1.0/release-manifest.json) identify the precise EXE and ZIP built for this preparation. **Use checksums for the same version and artifact; an EXE checksum is not a ZIP checksum.** Different rebuilds, packaging changes, or signing produce different hashes even if the version number stays the same.

<!-- RELEASE-HASHES: copied from the verified 2.1.0 package; regenerate after rebuilding or repackaging. -->
Prepared locally on **2026-09-16 UTC** from `experiments-latest` at `90f960dd`, using Release / win-x64 / self-contained / single-file with ReadyToRun **off**:

| File | Size (bytes) | SHA-256 |
| --- | ---: | --- |
| `Beacon.exe` | 77,624,725 | `C35F908A02A52E361B58E82E31E00B531864C31CD7F8CA420D5A3CE68E97A686` |
| `Beacon-2.1.0-win-x64.zip` | 72,071,577 | `EDDF545A27265041DAD7862DFDF2C3A0F069324C424AB417DBE5CC8F0E786BAB` |

These hashes supersede the earlier local 2.1.0 EXE-only ZIP, not every file with the same version name. No new GitHub release upload or branch promotion was performed as part of this preparation.
<!-- /RELEASE-HASHES -->

In PowerShell, run this in the download folder:

```powershell
Get-FileHash -LiteralPath '.\Beacon-2.1.0-win-x64.zip' -Algorithm SHA256
# After extraction:
Get-FileHash -LiteralPath '.\Beacon.exe' -Algorithm SHA256
```

Compare the full `Hash` value with the corresponding file in the version's official `SHA256SUMS.txt`. If it differs, do not run the file: confirm the version and redownload from the official release. Hashes detect changed bytes; they **do not provide a verified publisher, prove a file is harmless, or bypass Windows protection**. Self-built copies are not expected to match a published binary automatically.

### Windows SmartScreen and signing

The prepared executable is **unsigned**, not self-signed. Windows may show **“Windows protected your PC”** or **“Unknown publisher”** for an unrecognized download. A reputation warning is different from Defender naming a specific malware threat; neither should be ignored without review. The `GlitchedLoaiza` company/copyright metadata is branding, not an Authenticode signature.

Use the official source/download, verify the matching checksum, and follow your organization's software policy. Do not disable SmartScreen or Defender to use Beacon. No certificate purchase, signing service, or guaranteed warning-free launch is promised by this release.

## Your first search

1. Click **Folder** to choose a folder, or **Archive** to choose a compressed archive. **Path** displays the selected source.
2. Choose **Literal text** in **Search mode** and enter a term such as **"error"** in **Search for**. Leave out the surrounding example quotes.
3. Leave **Case sensitive** unchecked if **"error"** and **"ERROR"** should both match, then click **Scan**.
4. Wait while Beacon counts eligible files and then searches them. Select a file in **Matched Files** to read its preview.
5. Expand **Match details** to select a matching line, event, or request. Use the preview's navigation buttons to explore additional matches.

The optional welcome tour explains these steps and exporting. You can decline or exit it at any point; it is offered only once per user. The **? Help** button immediately left of **Settings** opens a searchable offline guide, including five beginner regex lessons and examples checked against Beacon's search engine. No tour step searches files or exports data automatically.

### Search modes

| Mode | When to use it | Example input and meaning |
| --- | --- | --- |
| **Literal text** | Exact words, phrases, or punctuation | **"connection failed"** finds that phrase. |
| **Whole word** | Avoid parts of longer words | **"cat"** matches the word, not **"catalog"**. |
| **Regular expression** | Find text with a changing value or shape | **"code=\d+"** finds **"code=503"** or **"code=7"**. Uses .NET regex. |
| **Any term** | At least one search term is enough | **"error timeout"** finds either word. |
| **All terms** | Every term must occur in the same record | **"error timeout"** requires both words, in either order. |

Double quotes above label examples; do not type them unless grouping a phrase in **Any term** or **All terms**. For a phrase plus another term, the actual field can contain `"connection failed" timeout`. Ordinary text is searched line by line; EVTX and HAR matching works within one event/request. Regex operations have a timeout; inputs are limited to 4,096 characters and Any/All searches to 32 terms. The in-app tutorial links to [Microsoft's .NET regex reference](https://learn.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference).

### Changing, cancelling, and resetting

- Change the search or source, then click **Scan** again. Editing the search box does not change an existing result set or its highlights.
- Click **Cancel** or press **Esc** to cancel an active scan. Cancellation is cooperative: an in-progress codec or native operation may take time to return. Results already collected may be partial.
- **Reset** clears the current work. Export any results you need before resetting. It is different from **Restore defaults** in Settings.

### Reading counts and limits

A *record* is a matching text line, EVTX event, or HAR request. Several highlighted words can belong to one record. A **+** beside a count indicates partial results; hover for the reason and review **Diagnostics**. The counted/scanned file total includes files with no match and is not the matching-file count.

New settings collect multiple records and leave HAR redaction off. Explicit saved preferences remain intact. If only one event/request appears, check **Settings → Search → Stop after the first matching line, event, or request per file**, clear preview filters if needed, save, and rescan.

Selected defaults (configurable, not unlimited guarantees):

| Setting | Default |
| --- | ---: |
| Maximum total matching files | 10,000 |
| Maximum detailed records per file | 300 |
| Maximum matching EVTX events / HAR requests per file | 300 each; the smaller applicable record limit wins |
| Maximum input file size | 500 MB |
| Maximum preview size | 25 MB |
| Maximum request/response body size | 25 MB per body |
| Archive nesting depth | 1; configurable from 0 to 5 |

Archive entry counts, expanded bytes, compression ratios, and helper timeouts have additional safeguards. Counting uses the same eligible-file traversal as scanning, including nested archives, so it can add startup work. Files changing between passes, read errors, cancellation, or limits can prevent scan totals from reaching the pre-count.

## Investigating events and web requests

### EVTX: Windows event logs

- Read event ID, provider, severity, timestamp, and message in the event preview.
- Use **Settings → EVTX and HAR** for pre-scan event ID/provider/severity/UTC filters; save and rescan.
- Open the collapsible **Event tools** pane on the right to narrow already-collected events. Clearing it restores captured events; it does not recover data omitted by the original scan.
- **Raw event XML** stays beneath the message. View it or copy it when the captured document is complete; shortened XML is labeled and cannot be copied as complete data.
- Missing messages can require provider resources from the source computer. XML stays searchable, and Event tools explains the supported offline export/LocaleMetaData workflow. Beacon does **not** download or register provider DLLs.

### HAR: saved HTTP traffic

- Inspect method, URL, status, UTC time, duration, headers, and available request/response bodies. Beacon reads the saved capture; it does not replay web requests.
- Filter before scanning in Settings or narrow the captured view using the collapsible **HAR tools** pane: method, exact host, status codes, MIME text, minimum duration, and UTC range.
- Optional Base64 decoding supports bounded UTF-8 text. Oversized, invalid, binary, or unsupported encodings produce omission notices rather than unbounded reads.
- **HAR redaction is opt-in.** Unchecked, original matching data remains visible in new previews and exports. When enabled, searches still use original bounded data, but header values, URL query/user information, recognized JSON/form secrets, and unstructured bodies may be hidden afterward. A sensitive-only match can therefore have no visible highlight.
- Changing redaction requires saving and rescanning; a previously redacted result cannot reveal discarded values. The Settings On/Off explanation follows the checkbox.

**Redaction is not anonymization.** Search terms, file/URL paths, server addresses, unrecognized fields, diagnostics, and non-HAR results can remain sensitive. Review everything before sharing, regardless of redaction state.

Both sets of filters use editable UTC date/time pickers with arrows, direct text input, Apply/Cancel, and **Any time**. EVTX severity supports multiple named selections; provider dropdowns suggest captured providers and accept manual text. All filled-in filter fields must match. Dates are inclusive and use UTC, not automatically the local incident time.

## Exporting and diagnosing problems

- **Export…** saves a self-contained HTML report grouped by file, with source paths, jump links, highlighted snippets, and up to five captured context lines. Logos are embedded. Long lines can be shortened.
- Reports retain the scan's query, settings, filter scope, partial-coverage warnings, and HAR redaction state. **Preview-only filters do not remove results from exports.** A running/cancelled scan exports only what has been captured so far.
- **Copy paths** copies result source locations, including logical paths inside archives.
- **Diagnostics** opens the separate problem list. Filter it, refresh its snapshot, or copy diagnostic information. It is not a second export screen.
- Writes use a temporary destination and replace the report only on success. Cancelling export preserves an existing destination report. Source-overwrite safeguards remain in place.

## Files and archives

Default text extensions: `.txt`, `.log`, `.json`, `.xml`, `.csv`, `.html`, `.reg`, `.ini`, `.cfg`, `.config`, `.nfo`. Specialized readers handle `.evtx` and `.har`. Text extensions and excluded folders are configurable under **Files and access**.

| Container | Engine / scope |
| --- | --- |
| ZIP, 7z, RAR/RAR5, TAR | SharpCompress 0.50.4; common and solid fixtures are regression-tested. |
| GZIP text; compressed TAR (GZIP, BZIP2, XZ) and aliases such as `.tgz`, `.tbz2`, `.txz` | Bounded wrapper handling preserves nested paths and preview reopening. Raw non-TAR BZIP2/XZ payloads are not newly guaranteed as supported. |
| CAB | Bundled 7-Zip Extra `7za.exe` from 7-Zip.CommandLine 25.1.0. |

An archive inside another archive counts toward nesting depth; depth 0 disables entering nested archives. Encrypted entries, unsafe paths, and reported links are rejected. Byte, entry, ratio, and time limits can stop processing early. A partial read is not a whole-archive integrity certification. Customer multipart/ZIP64 edge cases and very large solid archives still require representative testing.

## Appearance, Help, and updates

- **Settings → Preview and diagnostics → Appearance** offers **Light**, **Dark**, and **System theme**. There is no separate toolbar theme-toggle button.
- Native Windows title bars and controls remain; supported caption colors follow the app. Fonts, wrapping, formatting, preview size, and diagnostic detail are configurable.
- **Help** is searchable offline. The optional borderless first-launch tour explains simple searching and exporting and will not be offered again on later launches, including after skipping it.
- Startup checks notify about newer stable GitHub releases without blocking the app. The taskbar-centered notice has a five-second countdown; **View release** opens GitHub only on a click. No updates are automatically downloaded or installed.
- **Settings → Preview and diagnostics → Updates → Check for updates** shows a persistent inline result. Failed checks are distinct from being up to date.
- Beacon is single-instance per Windows user/session. A second launch requests activation of the existing window; Windows still controls foreground focus.

### Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| `Enter` | Start a scan when Scan is available. |
| `Esc` | Cancel an active scan. |
| `F3` | Next match/event/request for the active supported preview. |
| `Shift+F3` | Previous event/request; plain-text preview currently still moves forward. |
| `Ctrl+Down` | Next matching file. |
| `Ctrl+R` | Reset when not scanning. |
| `Tab` / `Shift+Tab` | Move between controls. |

Use the visible navigation buttons when a shortcut does not apply to the active preview. Native window buttons, system menu, dragging, and snapping are retained.

## Known limitations and safe use

- Searches and previews operate within configured limits; no “unlimited file size,” “instant cancellation,” or “cannot crash” guarantee is made.
- HAR input is parsed as a bounded JSON document; large captures can still use substantial memory. Unstructured bodies may be withheld when redaction is enabled.
- HTML/XML/JSON text search uses extracted text, not a full browser search model. Source matches and displayed highlights can differ. Truncated documents are shown as text rather than rendered as incomplete HTML.
- EVTX message rendering depends on native Windows/provider resources. Portable locale metadata should stay next to its EVTX on disk; individual archive-entry extraction does not bring sibling metadata automatically.
- The exposed scan worker count is reserved for future work and is not a parallelism or speed guarantee.
- The full pre-count can repeat extraction work; temporary nested files are retained for previews until Reset/close. Do not delete them while the app uses them.
- Access-denied ownership recovery is optional, confirmed, and nonrecursive; it can request administrator approval. Only test it on disposable, authorized paths. Ownership does not necessarily grant read access.
- Network locations and security-policy restrictions may add delays or prevent file/helper access. Record failures in Diagnostics rather than assuming the source contains no match.
- Signing, distribution notices, representative-file checks, and target-machine smoke tests remain release-owner responsibilities. See [release preparation](docs/releases/2.1.0/PREPARATION.md) for current evidence and outstanding sign-off.

## Build, test, and package from source

Use the **.NET 10 SDK** on Windows and Visual Studio with .NET 10/WPF support (the current validation environment uses Visual Studio 2026). The solution file is **`Beacon_FindInFiles.slnx`**.

```powershell
git clone https://github.com/GlitchedLoaiza/Beacon_FindInFiles.git
Set-Location Beacon_FindInFiles
dotnet build Beacon_FindInFiles.slnx -c Release
dotnet run --project tests/Beacon.SafetyChecks/Beacon.SafetyChecks.vbproj -c Release -- "$PWD/Beacon_FindInFiles/7za.exe"
```

The regression runner is a console application, not a Test Explorer assembly. It tests headless services, actual WPF controls, real WebView2 rendering, archive fixtures, lifecycle/cleanup, redaction, help, and onboarding without using customer logs. Full run instructions and limitations are in [tests/Beacon.SafetyChecks/README.md](tests/Beacon.SafetyChecks/README.md).

Create a self-contained release package and fresh checksums with PowerShell 7:

```powershell
.\scripts\Publish-Beacon.ps1 -Destination "$PWD\artifacts\Beacon-2.1.0"
```

The script defaults to ReadyToRun off, matching the current Folder profile; `-ReadyToRun` opts in when no profile is supplied. `-PublishProfilePath` accepts a full `.pubxml` path. User-specific profiles are ignored by Git, so no maintainer's OneDrive path is required to build the project. Raw `dotnet publish` defaults to the project's ReadyToRun setting (on); different configurations produce different hashes. The script stages output, includes legal notices, checks ZIP entries, preserves previous destination artifacts, and writes `SHA256SUMS.txt` plus `release-manifest.json`. It does not merge, commit, sign, or upload anything.

The [`BenchmarkSuite1`](BenchmarkSuite1/BenchmarkSuite1.csproj) measurement project is development-only. Its recorded [A–B–B–A results and rollback patch](tests/Beacon.SafetyChecks/PERFORMANCE-AUDIT.md) cover specific text/stored-ZIP workloads, not an application-wide speed guarantee. Regression tests, benchmarks, and historical evidence are intentionally retained for future diagnosis.

## Contributing and feedback

- **GlitchedLoaiza** — lead developer and maintainer.
- **sgtxjosue** — HAR module and debugging contributions.

Report bugs through [GitHub Issues](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/issues) or discuss features in [GitHub Discussions](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/discussions). Include the version, relevant settings, expected/actual behavior, and a sanitized sample where possible. Do not upload credentials, private logs, or unreviewed HAR captures.

## Changelog

### v2.1.0 — prepared for release

#### Added

- Dedicated Settings window with validated file/archive/record limits, Light/Dark/System themes, and installed-font selection.
- Five search modes: Literal text, Whole word, Regular expression, Any term, and All terms, plus file-name/full-path targets.
- Detailed result locations and navigation, compact file counts, a collapsible Match details pane, and event/HAR tools panes.
- EVTX pre-scan and collected-event filters, explicit raw XML inspection/copying, missing-provider diagnostics, and offline export guidance.
- HAR method/host/status/MIME/duration/UTC filters, bounded body handling, optional Base64 decoding, and opt-in presentation redaction.
- Editable UTC pickers with nonanimated arrows/direct input, named multi-select severity choices, and discovered-provider suggestions.
- Direct HTML export with captured five-line context, safe text encoding, jump links, embedded Beacon logos, and a background watermark.
- Separate searchable Diagnostics and a beginner-focused offline Help window, including five regex lessons and 12 validated recipes.
- Optional, once-per-user, borderless welcome tour; version shown at the splash's lower-right.
- Startup update notifications and Settings-only manual update checks; no automatic update downloads or installation.
- Single-instance startup with existing-window activation and crash recovery, plus theme-aware native title bars.

#### Fixed and changed

- Full nested-file pre-count shares the same eligibility/depth rules as searching; CABs are counted by their contents rather than as one file.
- First-match-per-file and HAR redaction now default to unchecked. Explicit saved preferences remain intact; changed scan settings require rescanning.
- HAR Settings guidance now accurately reports redaction On/Off instead of always claiming it is enabled. Opted-out previews/exports retain matched data.
- Improved nested archive preview lifetime, selected-archive display paths, Unicode-safe bounded previews, and incomplete-document text fallback.
- Corrected stale WebView2 highlight completions and shared initialization; preserved author styling in HTML previews.
- Replaced redundant toolbar controls and moved theme selection into Settings; aligned bottom actions with the preview and extended the left pane downward.
- Standardized first-party company, author, copyright, and publisher branding to GlitchedLoaiza. This does not create a verified signing identity.
- Excluded the three WebView2 API-documentation XML files from publish output; release packaging now includes the required legal notices and artifact checksums.

#### Security, compatibility, and maintenance

- Upgraded SharpCompress from 0.38.0 to **0.50.4**, addressing the reported traversal advisory and adding the required factory/stream compatibility changes. Retained **7-Zip.CommandLine 25.1.0** for CAB.
- Added bounded compressed-TAR handling, explicit ZIP size/CRC validation for complete entry reads, and traversal/link/encrypted-entry protections. Limits and CRC checks were not disabled to accept malformed inputs.
- Added controlled line/context allocation optimizations. The documented six-case A–B–B–A test saw roughly 20–38% lower mean time and 21–53% lower allocated bytes, with some higher-generation GC increases. These are workload-specific measurements, not universal app claims.
- Extracted active source traversal, HAR/EVTX collection, and preview I/O into UI-independent services, with captured settings and explicit temporary-source ownership.
- Expanded automated regression coverage, preserved archive fixture provenance, and retained measured-change/rollback evidence. See the current preparation record for the exact validation run and remaining manual gates.

### v2.0.1 — historical release notes

#### 🐛 Bug Fixes
- **Focus restoration on scan cancel**: Keyboard focus now correctly returns to the search field when a scan is cancelled via the Cancel button or Esc key, allowing the user to immediately start a new search without manually clicking.
- **Focus restoration on scan completion**: Keyboard focus now correctly returns to the search field upon scan completion, consistent with the cancel behaviour.
- **Non-EVTX files in nested archives not loading**: Text and HTML/XML/JSON files contained within nested archives (e.g., a ZIP inside a CAB) were incorrectly reported as corrupted or failed to open in the preview pane. The archive entry routing logic has been corrected to handle all supported file types at any nesting depth.
- **Event ID search not returning results**: Searching for a term such as `Event ID 813` or a bare numeric Event ID failed to match events even when the ID was present. The EVTX scanner now evaluates the synthesised `Event ID {n}` label alongside the raw XML and formatted message, ensuring all matching events are surfaced in results.

#### ✨ Improvements
- **EVTX raw XML formatted as indented markup**: When a Windows Event Log message cannot be rendered via its provider DLL, or when the search term is found only in the raw event data, the fallback display now presents the XML in properly indented, human-readable form rather than a single-line string.
- **HTML preview rendered in light mode when OS dark mode is active**: The WebView2 preview pane now instructs the Chromium engine to apply a light colour scheme, preventing pages that do not define their own dark-mode CSS from having their colours inverted when the operating system is in dark mode. Pages that ship their own dark-mode styles are unaffected.

---

## License and acknowledgments

Beacon is available under the [MIT License](LICENSE), copyright (c) 2025 **GlitchedLoaiza**. The existing copyright year and contributor credits are preserved.

Thanks to **SharpCompress**, **7-Zip**, **Microsoft WebView2**, and **.NET/WPF**. Their original license attributions remain separate from Beacon's branding. See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) and the packaged `licenses` directory for the bundled components.

---

**Made with ❤️ for log analysis workflows**
