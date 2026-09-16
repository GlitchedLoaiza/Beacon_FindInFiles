# Beacon: Find in Files v2.1 — Your logs, easier to explore

![Version](https://img.shields.io/badge/version-2.1.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20x64-lightgrey)
![Framework](https://img.shields.io/badge/.NET-10-purple)
![License](https://img.shields.io/badge/license-MIT-green)

**Less time opening files. More time finding answers.** Beacon is a Windows desktop tool for finding words, error codes, and patterns across logs and archives. Choose a folder or archive, enter your search, and explore the matching files—all in one place.

Built with ❤️ by **GlitchedLoaiza** for troubleshooting, log analysis, and anyone tired of searching files one by one.

> **Getting ready for 2.1!** This README describes the prepared 2.1.0 build. Check [GitHub Releases](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/releases) for available downloads. Release-preparation details live [here](docs/releases/2.1.0/PREPARATION.md); the previous version is preserved on [`Beacon2.0.1`](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/tree/Beacon2.0.1).

## 🎯 Features at a glance

- **Search across file types:** text logs, Windows events (EVTX), saved web traffic (HAR), and supported archives.
- **Search your way:** literal text, whole words, regex patterns, or combinations of terms, with optional case sensitivity.
- **Explore without losing your place:** highlighted previews, matching locations, and navigation between files, events, and requests.
- **Look inside nested archives:** configurable depth and limits, with matching files shown under their original archive paths.
- **Investigate events and requests:** collapsible EVTX/HAR tools, editable date pickers, and filters that help narrow the view.
- **Share a readable report:** export matching files and nearby context to a self-contained HTML file.
- **Make it comfortable:** Light, Dark, System, and Beacon Theme, adjustable previews, and native Windows window controls.
- **Get help as you go:** an optional welcome tour and searchable offline Help, including a beginner regex guide.

Whether you're tracking an error across diagnostic bundles, finding a configuration value, or reviewing a saved web request, Beacon helps you get to the relevant text.

## 📦 Download and run

1. Visit the official [Releases page](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/releases).
2. Download the ZIP for your chosen version and its `SHA256SUMS.txt` verification file.
3. Verify the download using the instructions below, then extract it to a folder you can write to.
4. Run **`Beacon.exe`**. That's it—no installer or separate .NET installation is needed.

Keep the included `LICENSE` and `licenses` folder with the app when sharing it. You don't need the source code, benchmarks, or test files to use Beacon.

### 🖥️ What you'll need

- **Windows 11 x64** is the primary tested environment. Other Windows versions need separate testing.
- **.NET 10 is included** in the published executable.
- **Microsoft Edge WebView2 Runtime** provides HTML/XML/JSON previews. It's often already on Windows; Beacon offers a text fallback if it can't initialize WebView2.
- Writable settings/temp folders and permission to run the bundled CAB helper. Workplace device policies may restrict these.

### 🔐 Verify your download

Use the checksum for the **same file and version** you downloaded. The [2.1.0 checksum list](docs/releases/2.1.0/SHA256SUMS.txt) includes separate values for the ZIP and EXE.

<details>
<summary>Show the prepared 2.1.0 hashes and PowerShell verification steps</summary>

These values identify the exact prepared package. The [artifact manifest](docs/releases/2.1.0/release-manifest.json) contains the full build details.

**Source update:** the package recorded below predates Beacon Theme. Its hashes still identify that earlier download; a rebuilt theme-enabled release needs fresh packaging and checksums.

<!-- RELEASE-HASHES: copied from the verified 2.1.0 package; regenerate after rebuilding or repackaging. -->
Prepared locally on **2026-09-16 UTC** from `experiments-latest` at `90f960dd` (Release, win-x64, self-contained single-file, ReadyToRun off):

| File | Size (bytes) | SHA-256 |
| --- | ---: | --- |
| `Beacon.exe` | 77,624,725 | `C35F908A02A52E361B58E82E31E00B531864C31CD7F8CA420D5A3CE68E97A686` |
| `Beacon-2.1.0-win-x64.zip` | 72,071,577 | `EDDF545A27265041DAD7862DFDF2C3A0F069324C424AB417DBE5CC8F0E786BAB` |

These hashes replace the earlier local EXE-only package's hashes. Check the release notes before comparing a different 2.1.0 download; this preparation has not uploaded a new GitHub release.
<!-- /RELEASE-HASHES -->

In PowerShell, run this in the download folder:

```powershell
Get-FileHash -LiteralPath '.\Beacon-2.1.0-win-x64.zip' -Algorithm SHA256
# After extraction:
Get-FileHash -LiteralPath '.\Beacon.exe' -Algorithm SHA256
```

Compare the full `Hash` value with that file's entry in `SHA256SUMS.txt`. If they differ, don't run the file: confirm the version and download it again from the official release.

Rebuilding, repackaging, or signing changes the checksum. A self-built copy may differ from a published one. Hashes verify matching bytes—not publisher identity or whether an application is harmless.

</details>

### ⚠️ A note about Windows SmartScreen

Beacon is currently **unsigned**, not self-signed. Windows may show **“Windows protected your PC”** or **“Unknown publisher”** for an unfamiliar download. This is different from an antivirus detection naming a specific threat.

Download from the official release, verify its checksum, and follow your organization's software policy. Please don't disable SmartScreen or Defender. The GlitchedLoaiza branding is not a verified signing certificate.

## 🚀 Your first search

1. Click **Folder** to choose a folder, or **Archive** to choose a compressed archive. **Path** displays the selected source.
2. Choose **Literal text** in **Search mode** and enter a term such as **"error"** in **Search for**. Leave out the surrounding example quotes.
3. Leave **Case sensitive** unchecked if **"error"** and **"ERROR"** should both match, then click **Scan**.
4. Wait while Beacon counts eligible files and then searches them. Select a file in **Matched Files** to read its preview.
5. Expand **Match details** to select a matching line, event, or request. Use the preview's navigation buttons to explore additional matches.

💡 **New to Beacon?** Try the optional welcome tour, or click **? Help** beside Settings whenever you need a hand. You can skip or exit the tour at any point; it never runs a search or export for you. Help works offline and includes five beginner regex lessons.

### Search modes

| Mode | When to use it | Example input and meaning |
| --- | --- | --- |
| **Literal text** | Exact words, phrases, or punctuation | **"connection failed"** finds that phrase. |
| **Whole word** | Avoid parts of longer words | **"cat"** matches the word, not **"catalog"**. |
| **Regular expression** | Find text with a changing value or shape | **"code=\d+"** finds **"code=503"** or **"code=7"**. Uses .NET regex. |
| **Any term** | At least one search term is enough | **"error timeout"** finds either word. |
| **All terms** | Every term must occur in the same record | **"error timeout"** requires both words, in either order. |

The double quotes above label examples—leave them out when typing. In **Any term** or **All terms**, quotes can instead keep a phrase together: the actual field can contain `"connection failed" timeout`.

Not sure about regex? Start with **Literal text**, then explore the in-app lessons or [Microsoft's .NET regex reference](https://learn.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference). Regex searches have a time limit; Any/All accept up to 32 terms, with 4,096 characters allowed in the search field.

### Changing, cancelling, and resetting

- Change the search or source, then click **Scan** again. Existing results keep their original search until you rescan.
- Click **Cancel** or press **Esc** to stop an active scan. Some file-reading operations take a moment to finish; results collected so far may be partial.
- **Reset** clears the current work. Export any results you need before resetting. It is different from **Restore defaults** in Settings.

### Reading counts and limits

A *record* is a text line, Windows event, or web request. Several highlighted words can belong to one record. A **+** beside a count means there may be more results; hover over it for the reason or check **Diagnostics**. The scanned-file count also includes files with no match.

Seeing only one event or request? Check the first-match option under **Settings → Search**, clear any Event/HAR tools filters, and rescan. Saved preferences remain in effect when you update Beacon.

<details>
<summary>Default limits and how to collect more results</summary>

Adjust these in Settings, save, and run the search again. Higher limits can use more memory, disk space, and time.

| Setting | Default |
| --- | ---: |
| Maximum total matching files | 10,000 |
| Maximum detailed records per file | 300 |
| Maximum matching EVTX events / HAR requests per file | 300 each; the smaller applicable record limit wins |
| Maximum input file size | 500 MB |
| Maximum preview size | 25 MB |
| Maximum request/response body size | 25 MB per body |
| Archive nesting depth | 1; configurable from 0 to 5 |

Archive entry counts, expanded bytes, compression ratios, and helper timeouts have additional safeguards. Beacon counts before searching, including inside nested archives. That can take time, and changing files, read errors, or cancellation can leave the final scanned count below the original total.

New settings leave first-match mode and HAR redaction unchecked. Existing explicitly saved choices are preserved.

</details>

## 🔎 Explore events and web requests

### EVTX: Windows event logs

- Select an EVTX result to read its event ID, provider, severity, time, and message.
- Open **Event tools** on the right to narrow the events you see—for example, choose **Error** and click **Apply filters**. **Clear filters** brings the other collected events back.
- To filter the next scan instead, use **Settings → EVTX and HAR**, save, and rescan.
- Expand **Raw event XML** beneath the message to inspect or copy the event's structured data. Incomplete XML is clearly labeled.
- Missing a readable message? Event tools explains how to get a portable log from the source computer. XML remains searchable; Beacon doesn't download provider DLLs.

### HAR: saved HTTP traffic

- Read request methods, URLs, statuses, timings, headers, and available bodies. Beacon examines the saved HAR file; it doesn't replay requests.
- Use **HAR tools** to filter the view by method, host, status, content type, duration, or time. The same filters in Settings apply to your next scan.
- Optional Base64 decoding reveals supported encoded text. Body-size and decoding notices explain when content can't be shown.
- **Redaction is your choice.** Leave it unchecked to see original matching data. Enabling it hides recognized sensitive values and withholds unstructured bodies, including some search matches. Save and rescan after changing it.

🕒 Date filters use **UTC** and include both endpoints. Pick a date with the arrows, type it directly, or choose **Any time** to remove the limit. Provider suggestions and named severity choices make event filtering easier, too.

🔒 **Review before sharing.** HAR redaction isn't full anonymization: paths, search terms, and other fields can still contain private data. An earlier redacted result needs a new scan to reveal values again.

## 📤 Save results and investigate problems

- **Export…** creates a readable HTML report you can open in a browser, with matching files, highlighted snippets, and up to five nearby lines. The logo and styling travel with the file.
- **Copy paths** copies the matching files' locations, including paths inside archives.
- **Diagnostics** explains skipped sources, errors, and limits. Use it when you expected more results; you can filter, refresh, and copy the problem list.

Reports keep the scan's settings and privacy state. **Filters in Event tools or HAR tools only change the preview—not which results are exported.** Exporting during or after a cancelled scan includes only the results collected so far. Cancelling an export keeps an existing destination report intact.

## 📄 Supported files and archives

Default text extensions: `.txt`, `.log`, `.json`, `.xml`, `.csv`, `.html`, `.reg`, `.ini`, `.cfg`, `.config`, `.nfo`. Specialized readers handle `.evtx` and `.har`. Text extensions and excluded folders are configurable under **Files and access**.

| Container | Support |
| --- | --- |
| ZIP, 7z, RAR/RAR5, TAR | Search entries and reopen their previews; common and solid-archive fixtures are tested. |
| GZIP text; compressed TAR (GZIP, BZIP2, XZ) | Includes aliases such as `.tgz`, `.tbz2`, and `.txz`. Raw non-TAR BZIP2/XZ payloads are not guaranteed. |
| CAB | Read using the bundled 7-Zip helper—no separate installation needed. |

Set nesting depth in **Settings → Archives**: the default is 1, with values from 0 to 5. Depth 0 skips archives inside another archive. Encrypted entries, unsafe paths, and archive links are rejected, and size/expansion limits can stop processing early. Check Diagnostics for the reason.

## 🎨 Make Beacon yours

- Choose **Light**, **Dark**, **System theme**, or **Beacon Theme** under **Settings → Preview and diagnostics → Appearance**, then click **Save**.
- **Beacon Theme** uses a softer, logo-inspired crimson with off-white button text. Input/preview outlines and selection highlights use coordinated red tones while backgrounds follow Windows' light/dark app preference. Disabled buttons stay neutral. Switch back to another theme whenever you like.
- Adjust fonts, wrapping, formatting, and preview size to suit your reading style.
- Expand the side panes when you need extra tools and collapse them when you want more room.
- Keep familiar Windows title-bar buttons, dragging, and snapping. Supported title-bar colors follow the app's theme.
- Launch Beacon again and it asks Windows to bring your existing window forward rather than opening a second copy.

### 🔄 Stay up to date

Beacon checks for a newer stable release at startup. A brief notification offers **View release** and disappears after five seconds. Prefer to check yourself? Open **Settings → Preview and diagnostics → Updates → Check for updates** for a message that stays in Settings.

**You're in control:** updates are never downloaded or installed automatically. An unsuccessful check is reported separately from being up to date.

## ⌨️ Keyboard shortcuts

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

## ❓ Helpful tips and known limitations

- **No matches?** Check the source, search mode, capitalization, and filters. Diagnostics may explain why files were skipped.
- **A match isn't highlighted?** It may be in a name/path, hidden by HAR redaction, or displayed differently in a formatted document. Regex can also match a position without selecting characters; Help explains this.
- **A preview is shortened?** Raise the preview limit in Settings if appropriate. Incomplete HTML/XML/JSON is shown as text so you can still read the available content.
- **Large or unusual files?** HAR documents can use considerable memory. Multipart/ZIP64 edge cases and large solid archives need testing with your representative files; a partial read isn't a full archive integrity check.
- **An EVTX message is missing?** Check Raw event XML and the offline guidance in Event tools. Logs brought from another machine may need its exported locale metadata alongside the EVTX.
- **Network or access problems?** Network latency and workplace policies can limit access. Ownership recovery is optional and asks for confirmation; don't try it casually on system or evidence files.
- **Looking for more speed?** The scan worker count is reserved for future work. Counting can repeat archive extraction, and temporary nested files remain available for previews until Reset/close. Leave those files alone while Beacon uses them.

The in-app **Help** guide has step-by-step troubleshooting. Maintainer validation and release sign-off are tracked separately in the [release preparation checklist](docs/releases/2.1.0/PREPARATION.md).

## 🔧 For developers

Want to build Beacon or help improve it? The source, regression tests, benchmarks, and rollback evidence are all included.

<details>
<summary>Build, test, and package from source</summary>

Use the **.NET 10 SDK** on Windows and Visual Studio with .NET 10/WPF support (the current validation environment uses Visual Studio 2026). The solution file is **`Beacon_FindInFiles.slnx`**.

```powershell
git clone https://github.com/GlitchedLoaiza/Beacon_FindInFiles.git
Set-Location Beacon_FindInFiles
dotnet build Beacon_FindInFiles.slnx -c Release
dotnet run --project tests/Beacon.SafetyChecks/Beacon.SafetyChecks.vbproj -c Release -- "$PWD/Beacon_FindInFiles/7za.exe"
```

The regression runner is a console application, not a Test Explorer assembly. It exercises services, WPF controls, WebView2, archive safety, and workflows using test fixtures. See the [test guide](tests/Beacon.SafetyChecks/README.md) for full instructions and limitations.

Create a self-contained release package and fresh checksums with PowerShell 7:

```powershell
.\scripts\Publish-Beacon.ps1 -Destination "$PWD\artifacts\Beacon-2.1.0"
```

The script creates the EXE/ZIP, includes legal notices, verifies the contents, preserves old artifacts, and writes checksums and a manifest. It does not commit, sign, or upload anything. It defaults to ReadyToRun off; `-ReadyToRun` opts in without a profile, and `-PublishProfilePath` accepts a `.pubxml` path. Raw `dotnet publish` uses the project's ReadyToRun default (on), so its hash may differ. User-specific publish paths are not required.

The development-only [`BenchmarkSuite1`](BenchmarkSuite1/BenchmarkSuite1.csproj) has [measured results and a rollback log](tests/Beacon.SafetyChecks/PERFORMANCE-AUDIT.md). Its findings apply to the tested text/stored-ZIP workloads, not every scan.

For final packaging and branch-promotion checks, follow [release preparation](docs/releases/2.1.0/PREPARATION.md). Keep the original component notices in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

</details>

## 🤝 Contributing and feedback

Ideas, bug reports, and suggestions are welcome—especially if something could be clearer or easier to use.

- **GlitchedLoaiza** — lead developer and maintainer.
- **sgtxjosue** — HAR module and debugging contributions.

Share bugs through [GitHub Issues](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/issues) or talk about features in [GitHub Discussions](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/discussions). Please include your version, what you expected, and what happened. If you attach a sample, remove passwords, credentials, and private data first.

## 📋 Changelog

### v2.1.0 — prepared for release

#### ✨ Added

- Dedicated Settings window with validated file/archive/record limits, Light/Dark/System themes, and installed-font selection.
- **Beacon Theme:** system-dependent light/dark backgrounds with muted-crimson action buttons, off-white text, coordinated red focus/selection accents, and themed hover/pressed/disabled states. Existing themes remain available.
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

#### 🛠️ Fixed and changed

- Full nested-file pre-count shares the same eligibility/depth rules as searching; CABs are counted by their contents rather than as one file.
- First-match-per-file and HAR redaction now default to unchecked. Explicit saved preferences remain intact; changed scan settings require rescanning.
- HAR Settings guidance now accurately reports redaction On/Off instead of always claiming it is enabled. Opted-out previews/exports retain matched data.
- Improved nested archive preview lifetime, selected-archive display paths, Unicode-safe bounded previews, and incomplete-document text fallback.
- Corrected stale WebView2 highlight completions and shared initialization; preserved author styling in HTML previews.
- Replaced redundant toolbar controls and moved theme selection into Settings; aligned bottom actions with the preview and extended the left pane downward.
- Standardized first-party company, author, copyright, and publisher branding to GlitchedLoaiza. This does not create a verified signing identity.
- Excluded the three WebView2 API-documentation XML files from publish output; release packaging now includes the required legal notices and artifact checksums.

#### 🛡️ Security, compatibility, and maintenance

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

## 📜 License and acknowledgments

Beacon is available under the [MIT License](LICENSE), copyright (c) 2025 **GlitchedLoaiza**.

Special thanks to **SharpCompress**, **7-Zip**, **Microsoft WebView2**, and **.NET/WPF** for helping make Beacon possible. Their original license attributions are included in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) and the release's `licenses` folder.

---

**Made with ❤️ for log analysis workflows**
