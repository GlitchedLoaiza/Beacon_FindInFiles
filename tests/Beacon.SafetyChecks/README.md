# Beacon safeguard regression checks

This Windows/.NET 10 console runner compiles production code and WPF views as linked source files, using the same WebView2 and SharpCompress versions as Beacon. Checks create isolated off-screen WPF windows on an STA thread. Source/navigation tests instantiate MainWindow with its application startup/shutdown handlers detached and use temporary fixtures, not customer logs. The runner does not alter persisted settings, ownership, or permissions. Browser checks require the installed WebView2 Runtime and use a separate temporary profile.

From the repository root:

```powershell
dotnet run --project tests/Beacon.SafetyChecks/Beacon.SafetyChecks.vbproj -- "$PWD/Beacon_FindInFiles/7za.exe"
```

A nonzero exit code indicates a failure. This is an executable regression runner, not a Test Explorer/MSTest project. Solution builds compile it; run the command above to execute the checks.

## Measured optimization and rollback

See [PERFORMANCE-AUDIT.md](PERFORMANCE-AUDIT.md) for the original source baseline, repaired benchmark host, retained line/context optimization, rejected experiments and exact rollback commands. The 2026-09-15 A–B–B–A validation retained the same frozen benchmarks and source hashes across four runs; machine-generated reports and full logs are preserved in [PerformanceEvidence/20260915-abba](PerformanceEvidence/20260915-abba). Mean times improved in the tested text/stored-ZIP scenarios and allocations decreased, but some higher-generation GC counts increased. This is not an application-wide speed guarantee or release approval. That audit's final Debug and Release regression runs each passed 71 groups; counts in the step-by-step sections below likewise describe historical snapshots, not the current runner's expected total.

## Step 12 release validation

The [source promotion validation record](../../docs/releases/2.1.0/PREPARATION.md#source-promotion-validation) tracks the current source review, build/regression results, and local-only move to master. The same [preparation record](../../docs/releases/2.1.0/PREPARATION.md) preserves the archived pre-theme EXE/ZIP checksums, manifest, legal-notice inventory, and 71-group package-validation run. Those artifacts do not validate the theme-enabled source, and source promotion does not publish a binary release.

See [RELEASE-VALIDATION.md](RELEASE-VALIDATION.md) for the historical Step 12 source commit, Release test/build/audit evidence, candidate checksum, published-executable smoke results and outstanding manual release gates. The 2026-09-15 pass completed 69 Release regression groups and a local published startup/duplicate-launch/normal-exit check. It does not authorize distribution: representative-file/target-machine testing, final version/signing choices, and notice/distribution review still apply to the exact shipping binary.

## Beginner help and multi-record defaults

Beacon 2.1 adds a once-only optional welcome tour after the main window opens. Start tour follows eight coaching steps beside source/search/mode/scan/results/details/export/diagnostics controls; Not now or Exit tour dismisses it at any stage. Back and Next/Finish are always user-driven. The tour never selects sources, changes search input, starts scans or saves exports automatically. It can explain disabled controls before results exist. The owned coaching window follows owner movement/resize and closes with the owner.

The offer is remembered per user in `%LOCALAPPDATA%/Beacon/welcome-tour.offered`, separately from Settings. Existing users also receive one offer when first running this implementation. Skipping or closing counts as offered, so the second and later launches do not show it again. Restore defaults does not reset onboarding. An unwritable marker location skips the offer safely. Tests use temporary marker paths, not the user's real state. The splash displays the running assembly's version at the lower-right; release metadata is 2.1.0 / 2.1.0.0, displayed as 2.1.

The accessible ? Help button immediately left of Settings opens one owned, nonmodal Help window. Its offline guide includes quick-start instructions, all search modes/targets, counts and partial results, navigation, EVTX/HAR tools, date/provider/severity selectors, archive safeguards, redaction, exports, Settings, updates, shortcuts and troubleshooting. Search filters titles and descriptions; text can be selected/copied. Help follows the theme at opening and keeps native window controls.

The guide uses task instructions and expected outcomes, with advanced provider recovery separated from basic EVTX use. Five regex lessons progress from first input and symbol meanings through building a pattern, 12 common recipes and troubleshooting. Example-label quotes are distinguished from quotes users must type in Any/All phrase searches. Each recipe is checked through the real SearchQuery engine for matching, nonmatching and highlighted text; additional assertions cover multiline/CRLF, case sensitivity, dot/newline and zero-width behavior.

Regex topics show an optional Microsoft .NET regex quick-reference link. It opens only on a click, requires internet access, and reports browser failures inline. Link tests inject a browser action rather than launching an external page. The help itself stays offline. Tests establish content accuracy and UI behavior, not beginner comprehension; a first-time-user walkthrough is still recommended.

`StopAfterFirstMatchPerFile` defaults to false for new settings and Restore defaults. Existing explicitly saved values are preserved. Users with the earlier true preference must uncheck it in Settings → Search, save and rescan to collect more events. Regression checks cover the default/missing-property behavior, saved true preservation, Restore defaults, Help search/selection/empty state, minimum window size and both themes.

## Step 11 — active search and preview service boundaries

The active HAR and EVTX collectors live in `HarSearchService` and `EvtxSearchService`, independent of WPF. Each service clones its settings at construction and uses the captured query. `StructuredSearchResult` holds aligned records/details plus partial coverage; `EventRecordSummary` is shared with the preview. `SourceSearchService` orchestrates collection and merges accumulated structured records in `Finally` if a later read fails.

HAR input streams remain caller-owned. EVTX native records/readers are disposed in the service. No UI calls, clipboard operations, downloads or theme changes occur in either collector. Direct service tests exercise settings isolation, record limits/indexes, malformed HAR diagnostics, redacted sensitive-only matches, XML fallback, cancellation and retention of earlier records after a later failure. The existing UI/archive integration suite is retained.

`PreviewContentService` now owns disk, SharpCompress entry and CAB preview reads. It clones settings, returns `PreviewText`, enforces the existing hard budgets and preserves readable soft truncation/Unicode decoding. It releases file/archive handles before UI rendering and cleans its CAB extraction directory in `Finally`; CAB errors retain logical archive/entry context. Missing entries return no preview, while cancellation, integrity and hard-limit failures propagate to the UI adapter. Both text and WebView2 paths use this service; rendering and incomplete-document text fallback remain in MainWindow. Direct tests cover disk/ZIP/CAB reads, missing entries, settings isolation, hard/soft limits, cancellation and cleanup.

`SourceSearchService` owns the active folder/disk/SharpCompress/CAB traversal, metadata matching, text/structured collection and total-result limit. Count and search share that traversal and cloned options. It publishes plain `SourceSearchResult` records and uses callbacks for progress, diagnostics, source mapping and access-denied prompts. It is single-use and not intended for concurrent calls. Dispose it only after its run completes; count services are disposed immediately, while search services retain temporary nested sources until Reset/close so previews can reopen them. `MainWindow.Search` now handles UI conversion and lifetime, not recursive reading. Headless tests cover count parity, settings isolation, result caps, metadata-only EVTX, nested previews, cancellation and temporary cleanup.

The active-service separation and direct regression coverage planned for Step 11 are complete. WPF rendering, dialogs, UI-bound selection and application lifecycle remain UI responsibilities. Older private compatibility scanner helpers remain in MainWindow; their broader deletion is not required to route the active pipeline through services and is not claimed here. Step 12 still requires native EVTX/provider and packaged-app/customer-archive smoke tests, plus final release validation.

## SharpCompress security upgrade

Both projects use SharpCompress **0.50.4**; CAB support continues through the unchanged bundled 7-Zip helper. The signed NuGet package identifies upstream commit `c083c6efd843a844b0c8f7878787360e815be781`, matching the official 0.50.4 tag. This version is outside GHSA-6c8g-7p36-r338's affected range and adds TAR symlink protections beyond 0.50.3.

`ArchiveCompatibility` preserves compressed TAR entry semantics using bounded, delete-on-close temporary TARs. The supported System GZip provider avoids partial-read disposal errors; no integrity option is disabled. ZIP raw entry streams receive explicit streaming CRC32/size validation at EOF because SharpCompress's extraction-helper checksum option does not validate raw `OpenEntryStream` reads. Early first-match/cancel/limit exits do not certify an unread remainder and must not drain it beyond configured limits. Both text and WebView2 archive previews use the bounded entry helper.

Regression coverage includes ZIP, solid 7z, TAR/PAX, standalone GZip, TAR.GZ/BZ2/XZ and aliases, nested compressed TAR, Unicode names, count/search parity, archive reopens, encrypted entries, path/link rejection, malicious directory entries, checksum failures, preview limits, cancellation and cleanup. Pinned RAR/RAR5 normal/solid fixtures live under `Fixtures/Archives` with upstream license/provenance. Payloads are hashed in memory, never executed. The archived text payload hash was independently checked with installed full 7-Zip because the loose upstream text file differs from the archived version; the test runner itself does not depend on that installed tool.

Still smoke-test representative customer archives, multipart archives, ZIP64 edge cases, large solid-archive workloads and native packaged-app interactions before shipping. No finite fixture suite guarantees all archives are compatible. Do not disable CRC or byte/path limits to make a failing archive appear successful.

## Single-instance startup

Beacon allows one instance per Windows user and logon session, across executable locations/versions using the same application identity. The guard runs before the splash: subsequent launches signal the existing instance and exit without another splash, scan or update check. The existing UI restores a minimized window and requests activation of its current window or owned dialog. Windows foreground restrictions still apply; no focus-stealing workaround is used. Command-line paths are not forwarded.

Named mutex ownership is released at application exit and can be recovered after a terminated owner. Separate users/sessions remain independent. Regression tests use unique names and real child processes for secondary/concurrent launches, activation and normal/crash relaunch. Manually smoke-test the packaged app with rapid double launch, a visible splash, minimized/maximized main window, an open Settings dialog, and normal shutdown/relaunch.

## Native title-bar themes

MainWindow, Settings and Diagnostics retain native Windows frames, title/icon and system buttons. On Windows 11 build 22000 or newer, documented DWM attributes color the caption and text to match Beacon's theme. The main window reapplies them when Light/Dark/System/Beacon Theme changes; dialogs use their opening theme. Older Windows versions or rejected attributes leave the native frame available. High contrast resets caption colors to system defaults; the borderless splash is unchanged.

### Beacon Theme

Settings → Preview and diagnostics → Appearance offers **Beacon Theme** as a fourth saved option. It follows Windows app light/dark colors for the main window's background and native caption. Its original pure logo-red/black buttons were refined to logo-inspired muted crimson: `#B83243` normal, `#C43D4D` hover and `#9E2939` pressed, with `#FFF5F5` off-white text. The measured text contrast is 5.49:1, 4.76:1 and 6.94:1 respectively. The logo asset itself remains unchanged. Disabled buttons use neutral backgrounds and readable secondary text instead of faded red. Settings, Help, Diagnostics, the tour and shared picker buttons use the same button palette when opened from Beacon Theme.

Unselected tabs, dropdown surfaces, result-row hover backgrounds and event/HAR counter strips remain neutral. Focus/link accents use `#E58E9A` on dark backgrounds and `#B83243` on light backgrounds; selected rows/tabs and text selections use subtle `#482B31` / `#F7E6E9` backgrounds. Both primary and secondary selected-row text retain at least 4.5:1 contrast. The Beacon-only rounded RichTextBox template retains `PART_ContentHost`, selection and scrolling while replacing native blue hover/focus outlines. Changing back to Light, Dark or System restores the standard preview style and accents. The System default and numeric values of the existing theme choices are unchanged; Restore defaults resets the selected preference to System, and unsaved Settings choices do not recolor the application. Open dialogs retain their opening palette, matching the existing theme behavior.

`BeaconThemeChecks` verifies the unchanged PNG color, enum/serialization compatibility, both simulated Windows background modes, main-window theme switches and system notifications, actual normal/hover/pressed/focus/disabled button template states, selected filename/count contrast, input and preview focus outlines, RichTextBox selection/scrolling and owner/picker accent inheritance. Readonly interaction-state keys are set only by the test harness and restored afterward; tests do not move the user's pointer, change the Windows color preference or save user settings. Previously published EXE/ZIP hashes predate this feature and must be regenerated for a new release. Native caption buttons and HTML report colors are not recolored by this theme.

Regression checks cover native dark-mode readback, acceptance of caption/text colors, repeated theme changes and preservation of window styles, title, icon and resize mode. Windows does not support caption-color readback through DwmGetWindowAttribute on the tested system. Manually verify caption appearance in both themes, active/inactive states, high contrast, minimize/maximize/restore/close, dragging, double-click maximize, Alt+Space and Windows Snap Layouts. No custom frame hit testing or replacement caption buttons are introduced.

## Startup release notification

Settings → Preview and diagnostics → Updates includes **Check for updates**. It displays a persistent inline result: a newer version is available, Beacon is up to date with no updates pending, or the check could not be completed. It never triggers the floating notification. Duplicate clicks are disabled while checking, and closing Settings cancels the request. Manual checks use the same stable-release rules and request limits as startup checks.

Beacon makes one asynchronous request per main-window session to this repository's GitHub latest-release API. The request has an eight-second timeout and a 256 KiB response limit. Stable releases only are considered; offline, rate-limit and invalid responses produce no popup or scan diagnostic. Closing cancels the request. No logs, paths or search contents are sent, and no update assets are downloaded or installed.

A newer version shows a themed floating notification centered just above the main Windows desktop taskbar, independent of Beacon's location and size. Shell taskbar bounds and the popup's actual native dimensions determine its screen position; the primary work area is the fallback when taskbar geometry is unavailable. Non-bottom taskbars use the bottom of the corresponding work area. It slides up into view, displays “A newer version of Beacon is available (5s)” with a decreasing countdown, then slides down after five seconds. Windows reduced-motion preferences disable sliding. Closing or minimizing stops the notification. View release opens the fixed repository release page only after a click. Version comparison uses the application's assembly version against a numeric release tag (such as `v2.1.0`), or the existing `Beacon v2.1.0` title when the tag is `PublicRelease`. Bump the project Version, AssemblyVersion and FileVersion consistently for each release. Dismissal lasts for the current session; the next launch checks again. The application has no dummy notification button or sample-version handler; automated checks exercise the notification directly from the separate test project.

## Coverage

- Startup release checks use fake HTTP responses: stable version comparisons, PublicRelease/title fallback, request headers, malformed/oversized responses, offline/HTTP failures and cancellation. Actual-window checks cover notification display, action layout and dismissal without launching a browser.
- XML indentation, preserved declaration, disabled formatting, malformed input, DTD rejection, and parser size bounds.
- Low-priority light HTML canvas CSS (without overriding author styles).
- Unsafe archive paths, reserved device names, encrypted entries, links, entry counts, sizes, and compression ratios.
- Actual streamed byte budgets, exact-limit EOF, asynchronous reads, and cancellation.
- Filesystem exclusions, hidden-file filtering, file-size reporting, and cancellation.
- Junction-cycle prevention using an isolated temporary junction.
- A real CAB produced by Windows makecab, including a filename containing spaces.
- Corrupt CAB input, unsafe metadata, excess process output, stalled readers, and cancellation of an active reader.
- Light/dark Settings templates: every tab and dropdown, selected/normal option colors, heading and button text, checkbox marks, focus cues, disabled appearance, and minimum-window-size layout. Text color pairs must meet a 4.5:1 contrast ratio; hover/pressed button palette colors are checked too.
- TextBox glyph bounds at normal and larger font sizes, with a negative control recreating the original duplicate-padding clipping.
- Installed-font dropdown membership and sorting, saved-font selection, missing-font fallback, Restore Defaults, and type-to-select configuration.
- Splash scene: resolved animation targets, requested 60 FPS timeline setting, 3px/4.5s logo float, 3s scanning beam, continuous ring configuration, parallax timing, stable visual count, reduced-motion state, loading-only accessibility, work-area scaling and cleanup after close.
- Search modes, quoted phrases, Unicode word boundaries, regex syntax/timeouts, zero-width matches, cancellation, record limits, and explicit partial counts.
- Actual disk/HAR/nested-ZIP/CAB search paths, name/path-only matching, query snapshots, text/HAR detail navigation and synthetic event forward/backward regex navigation.
- Full pre-scan counts share the search traversal: standalone files, nested ZIP/CAB contents, multi-file CABs, depth limits, directory/excluded entries, count cancellation and temporary extraction cleanup. Counting must not publish hits or increment scanned files.
- Real WebView2 highlighting: phrases spanning inline elements, UTF-16 offsets, .NET lookbehind ranges, script/hidden-text exclusion, bounded capture, document preservation and rejection of stale or changed-DOM snapshots.
- Structured diagnostics: concurrent deduplication, repeat counts, capacity/truncation reporting and detached snapshots.
- CSV/JSON/HTML reports: original-query preservation, optional excerpt/technical-data omission, Unicode/quoting, CSV formula protection, HTML encoding, source overwrite prevention and cancellation/locked-destination preservation.
- Diagnostics UI: light/dark contrast, filtering, Settings-controlled technical details, refresh, minimum-size layout and absence of redundant export/options controls. Tests do not use the system clipboard or native save dialog.
- Captured five-line context: file boundaries, first-match lookahead, record-local event/request context, long-line shortening, read-limit recovery, detached copies and export after the original file is deleted. The generated highlighted report is also loaded in WebView2.
- HTML branding: exact embedded Beacon PNG bytes, successful browser image decoding, logo alignment to the left of BEACON, and decorative watermark opacity, placement and noninteractive layering in light/dark themes. Untrusted image/script markup remains escaped.

## Step 9 EVTX investigation — offline increment

- Settings → EVTX and HAR provides pre-scan IDs, exact case-insensitive provider name, severity levels and inclusive UTC start/end filters. Blank fields mean unrestricted. Event IDs remain comma-separated numbers, not ranges. Filters combine with AND and apply to EVTX contents before message formatting; name/path-only matches remain independent. Invalid settings are rejected before a scan clears existing results.
- Both Settings and Event tools use editable provider dropdowns populated from captured EVTX events where available. Suggestions are sorted and deduplicated; custom provider text remains supported before scanning. The severity dropdown offers named, checkable levels (0–5), preserves multiple selections, and uses All levels for an unrestricted filter. Saved filter strings retain their existing format.
- UTC date/time fields support direct typing (`yyyy-MM-dd HH:mm:ss` or `yyyy-MM-ddTHH:mm:ssZ`) and a nonanimated picker with editable year/month/day/hour/minute/second components. Click the arrows or use Up/Down in a component, then Apply or Enter to commit. Cancel or Escape leaves the original value unchanged; Any time clears the bound. Month/year arrows handle month ends and leap years; invalid input and date overflow are reported instead of silently clearing the filter.
- Event preview → Filter collected events uses the same semantics but only narrows captured matches. It does not rescan files, change exported results or bypass first-match/record limits. Counters distinguish visible and captured events. Selecting an explicit Match details entry clears preview filters so the jump remains reliable.
- Raw event XML is captured separately (up to 65,536 characters), displayed as text with safe formatting and copyable when complete. Shortened XML is clearly labeled and cannot be copied as a complete document. XML may contain sensitive data; clipboard interaction remains a manual smoke check.
- Missing message text produces a warning once per provider per file. The automatic raw-XML fallback setting is honored; XML searching and the explicit XML tools remain available. The UI offers only the implemented offline resource policy; provider acquisition/download modes remain deferred.
- On the source computer, export with display information or use `wevtutil al` on a COPY in a trusted working directory. Preserve the EVTX and LocaleMetaData together and open directly from disk. Archive-entry extraction does not carry sibling metadata automatically. Never modify original evidence or use untrusted symlink/junction-containing destinations. Beacon neither downloads provider binaries nor registers/loads arbitrary DLLs.
- Report metadata includes the captured pre-scan EVTX filter scope. Tests cover filter validation, inclusive UTC bounds, null metadata, persistence, preview navigation, XML state and themed expanders. Native provider-dependent EVTX rendering and portable-metadata behavior still require authorized real-file smoke tests.

## Step 10 HAR investigation

- HAR redaction is opt-in: new settings, missing saved values and Restore defaults leave `RedactSensitiveHarData` unchecked. Explicit saved true/false choices are preserved. Default searches retain original matching text and highlights in previews and exports; review sensitive data before sharing. To reveal values from an earlier redacted search, uncheck the option in Settings → EVTX and HAR, save and rescan. Tests verify the default, saved choices, Restore defaults, visible default matches and explicitly enabled redaction.
- Settings → EVTX and HAR provides pre-scan method, exact host, comma-separated status codes, response MIME substring, minimum duration in milliseconds and inclusive UTC bounds. Blank means unrestricted. The HAR preview has a collapsible right-side tools pane with the same filters; it narrows only captured requests. Navigation/counters distinguish visible and captured requests, while reports retain all captured matches.
- Request and response bodies have the configured HAR byte limit in addition to file/archive limits. Oversized bodies are omitted from search/preview and reported as partial coverage. Optional Base64 decoding checks encoded and decoded sizes, uses strict UTF-8 and rejects invalid or binary content. Plain response text is already decoded according to HAR conventions. Unknown encodings/character sets are not converted by guessing. The HAR JSON document itself still uses the existing bounded-input DOM parser, not a streaming JSON reader.
- With redaction enabled, matching uses original bounded request data, then captures only presentation-safe requests and context. All header values, URL query values/user information/fragments and recognized secret fields in nested JSON/form bodies are hidden. Unstructured or unsupported bodies are withheld. Sensitive-only matches remain as explicit hidden-match records with no misleading visible highlight. Disable redaction and rescan if raw inspection is required; changing Settings never reveals data in an already-redacted snapshot.
- This is conservative presentation redaction, not anonymization: field names, URL paths, search terms, file paths, server addresses, arbitrary JSON values under unrecognized keys and non-HAR results can still contain sensitive data. Review every export. Reports indicate per-file redaction and preserve that state through detached copies; HTML/CSV/JSON use the same captured context. Source files are never rewritten.
- Fake HAR tests cover combined filters/UTC boundaries, bounded and invalid Base64, JSON/form/header/query redaction, sensitive-only active-pipeline matches, themed preview filtering and exports. Existing disk/archive HAR pipeline checks continue to run. Real customer-format HAR variants and native save/clipboard workflows still warrant manual smoke testing.

## Step 7 search behavior

- When a scan starts from a selected archive, match labels are relative to that archive's contents: `outer.zip | logs/app.log` becomes `logs/app.log`, while `outer.zip | inner.zip | logs/app.log` becomes `inner.zip | logs/app.log`. Only the selected outer prefix is removed, including when absolute-path display is enabled. Folder-started scans retain archive names to distinguish sources. Full logical paths, full-path matching, copied source paths and preview/navigation paths remain unchanged. Regression checks exercise both display modes, multiple nesting levels and preview reopening. The Debug snapshot for this step passed 64 checks.

- Before searching, **Counting files…** traverses eligible folders and archive contents using the captured settings and configured archive nesting depth. Non-archive files are counted once; archive containers are not added on top of their contents. Counting does not parse regular-file contents or evaluate the query. Nested archives and CABs can require extraction in both passes, adding startup time; normal safety limits and cancellation still apply, and count-pass temporary files are removed before searching.
- Search progress uses the pre-counted total, not the former shallow estimate. This total is distinct from matching files. Files changing between passes, content-read failures, cancellation or result/safety limits can prevent the final scanned count from reaching the pre-counted total. Existing report fields retain the `EstimatedTotalFiles` name for compatibility.

- **Literal text** treats regex punctuation literally. **Whole word** applies Unicode word boundaries to the literal term; there is no duplicate Exact match checkbox.
- **Regular expression** uses .NET syntax, multiline anchors and a 100 ms timeout per matching operation. Invalid patterns are rejected before scanning. Queries are limited to 4,096 characters.
- **Any term / All terms** split on whitespace and support double-quoted phrases, for example `"connection failed" timeout`. All terms must occur in the same text line, event, HAR request or single name/path target. Up to 32 terms are accepted.
- HTML/XML/JSON are searched as normalized extracted document text (one record), consistent with the existing document search behavior. Script/style text is removed by the existing lightweight extraction approach; this is not a full HTML parser. Rendered/pretty-printed preview text can differ from the searchable representation, so source counts and visible highlight counts need not be identical.
- Selected contents, filename and full-path targets are alternatives. Metadata-only matches do not require parsing an EVTX or unsupported binary file. Nested paths retain original archive names rather than temporary filenames.
- **First matching record per file** remains in Settings. It stops after the first matching text line/event/request. Per-file detail limits and EVTX/HAR limits still apply. A compact count with `+` means results may be incomplete; its tooltip states why.
- **Match details** starts collapsed. Expand it to navigate to a source line, event or request. Visible highlight ranges are bounded at 10,000; zero-width regex matches can select a record without highlighting characters.
- A scan snapshots its query. Editing the next query does not reinterpret existing results. WebView2 consumes .NET-generated ranges rather than converting .NET regex into JavaScript regex. Stale browser work cannot replace a newer highlight count.
- Archive attribute filtering is applied when exposed by the reader. SharpCompress 0.50.4 exposes ZIP attributes; formats without attributes retain the specific unsupported-metadata fallback. Path, encryption, link-target and byte-budget checks remain in place.

Step 7 covers search; reporting is described below. EVTX provider acquisition and project-wide service extraction remain later roadmap work. Native EVTX reading/provider-DLL behavior still needs authorized real-file smoke tests; automated event navigation uses synthetic records.

## Step 8 diagnostics and reports

- **Export…** opens an HTML-only save dialog directly. It saves a readable report grouped by filename with source paths, match locations, jump links, highlighted snippets and a compact search summary. There is no intermediate export-options screen. Bottom actions align with the preview pane and follow its divider when resized.
- Every HTML report embeds the existing Beacon logo beside its heading and as a subtle lower-right background watermark. Images use PNG data URIs, so moved/shared reports need no companion files or network access. The content security policy permits only data images and inline styling; other resources remain blocked.
- **Diagnostics** opens a separate problems list, with source/message filtering, Refresh and Copy all diagnostics. It explains skipped or incomplete reads and contains no export controls. Opening it does not stop a scan.
- Context is captured during scanning, not reread during export. Text snippets normally show two lines before and after the matched line, shifting at file boundaries to retain up to five available lines. First-match mode reads only the limited extra context needed, without collecting additional matches. EVTX/HAR context stays within the matched event/request; document context is labeled as extracted or normalized text.
- Each context line is limited to 500 characters, with explicit shortening around the match. Single-line records, short files and read limits may produce fewer than five lines. Older results without captured context show an excerpt and a request to rerun the search rather than inventing surrounding lines.
- Running, cancelled, failed and result-limit-stopped scans are explicitly incomplete. Per-file truncation, warnings/errors and omitted diagnostics also affect the coverage label. A completed report is complete only within its configured search scope.
- The diagnostic store keeps up to 10,000 distinct entries. Repeated identical notifications increment a repeat count; new notifications beyond capacity are counted as omitted. Source/message/technical fields are bounded and marked when truncated. Intentional extension/folder filters are not individually logged as errors.
- The existing Summary/Detailed setting controls stack-trace capture and technical diagnostic display. Basic issue messages remain available in either mode. There is no automatic on-disk diagnostic log.
- Diagnostic filters affect the view only. Copy all diagnostics includes the unfiltered list.
- **Copy paths** copies distinct logical paths for current results, including archive paths, without copying log contents. The redundant Copy result action has been removed.
- HTML export includes **unredacted matching text and context** unless HAR redaction was enabled for that scan. HAR sections reflect their captured redaction state; paths, queries, error messages and non-HAR content can still be sensitive. The button tooltip, save dialog and report warn about sharing this data. Review the complete report before sharing, even when it reports redaction.
- JSON/CSV writers are retained internally for compatibility/tests, but are not offered in the primary UI. Their excerpt-free projections also remove all captured context. JSON retains structured metadata; CSV retains its consistent 16-column schema.
- CSV fields are quoted; formula-like cells are prefixed with an apostrophe. This intentionally changes affected cell text; use JSON for lossless interchange. Spreadsheet re-saving/import behavior can change escaping, so review untrusted reports before sharing or opening them in another tool.
- HTML is self-contained, encodes untrusted values and includes a restrictive content-security policy. It does not embed original files, executable scripts or external resources.
- Exports are written to a unique temporary file in the chosen directory and moved into place only after a successful write. Normal cancellation/failure removes the temporary file and preserves the existing report. Known scanned-root/result source paths are protected. The export button becomes Cancel export while writing, and application shutdown cancels/awaits the writer.

Manual release checks still include native save/overwrite dialogs, clipboard interaction, real EVTX/provider behavior and normal application startup/scan-cancel flows. The automated report tests use isolated fixtures and do not alter user data.

## Splash preview and startup checks

The splash test renders a preview to `bin/Debug/net10.0-windows/artifacts/splash-preview.png` under this test project. Its test-only constructor disables automatic main-window transition; production retains the existing 3.2-second delay and 0.35-second fade. Animation cycles do not artificially prolong startup.

The log text is illustrative. The actual status reads “Preparing your workspace…” and then “Opening Beacon…”. Windows client-area animation preferences disable ambient motion and the fade; on software rendering tiers the blurred layer effects are removed.

The checks validate animation configuration and sampled state, not measured frame rate. Manually review the moving splash and normal startup handoff, early close, multiple monitors/scaling and the Windows reduced-animation setting. A sustained 60 FPS claim requires profiling on the target hardware.

The test runner also serves as a controlled child-process fixture for the last group. It uses isolated temporary directories and removes them after each check. Tests require Windows makecab.exe, cmd.exe, and the bundled 7za.exe. No takeown operation is tested or performed.

## Phase 6 behavior and limitations

- Settings limits now apply to scan inputs and streamed archive data. Archive limits are per archive; nested archives also have the configured nesting limit.
- Encrypted, linked, unsafe, or over-limit archive content is rejected and reported. Password prompts are not implemented.
- Total-result and structured-match limits are reported as potentially incomplete results.
- Progress now uses a full pre-scan traversal, including bounded nested-archive extraction; see the counting behavior above.
- CAB output is copied through bounded streams rather than letting the reader write arbitrary paths. The configured timeout includes listing and entry extraction.
- Reset and shutdown await the active scan before cleaning up its temporary files.
- Other settings/features from later roadmap phases are not completed by this work.

## Manual smoke checks still required

1. Display the reported HTML and XML samples in both light and dark modes, from disk and archives. Confirm document colors and search highlighting remain readable.
2. Confirm complete JSON/HTML/XML previews still use WebView2. Oversized disk/archive/CAB previews must retain the bounded readable prefix and append a clear truncation notice, not replace content with a raw limit error. Truncated web documents must appear as plain text, not partially rendered HTML.
3. Set small total/EVTX/HAR match limits; verify caps and partial-result status text.
4. Scan a nested CAB/ZIP, then select its results after scanning. Confirm retained files support preview navigation.
5. Cancel a scan, reset, and start another. Close Beacon during extraction. Confirm no stale results or deleted-in-use temporary files.
6. Exercise access-denied skip/prompt behavior manually on an authorized test directory. Do not change ownership on system or production data.

The checks instantiate actual WPF templates, selected source pipelines and a real WebView2 control, but do not assess screenshots or simulate physical keyboard/mouse input. Native EVTX/provider behavior, complete application startup/scan-cancel flows, elevated ownership recovery, representative large-document rendering and a manual visual/keyboard review still require release smoke testing.

## SharpCompress 0.50.4 upgrade

The application and this runner both pin SharpCompress to **0.50.4**, upgraded from 0.38.0. **7-Zip.CommandLine 25.1.0 and bundled 7za.exe are retained unchanged**; CAB processing still uses the existing bounded 7-Zip path. No framework or visual-style change was part of that dependency upgrade. Existing uncommitted work was preserved at that snapshot, including the combined archive edits approved for review.

Historical validation for that dependency upgrade on the development machine: the original 0.38.0 baseline passed 54 checks. After correcting the error-only oversized-preview regression, the 0.50.4 runner passed **63 checks, 0 failures** in both Debug and Release. The IDE Debug build succeeded; a clean Visual Studio MSBuild Release rebuild completed with **0 warnings and 0 errors**. All runner/build exit codes for those final runs were zero. Preview regressions require readable content, not merely a limit-error message, and cover Unicode boundaries, exact-limit EOF, short reads, disk/archive consistency and safe partial-HTML fallback.

- The archive factory uses the new `OpenArchive` API. Compressed TAR wrappers (GZIP, BZIP2 and XZ) are decompressed to a size/ratio-bounded, delete-on-close temporary stream before SharpCompress opens the TAR. This preserves PAX metadata, logical entry paths, nesting depth and preview reopening without unbounded memory buffering. `.tar.xz` is recognized through its actual `.xz` final extension.
- SharpCompress's documented system GZIP provider is used to avoid its managed GZIP partial-read disposal exceptions masking cancellation or quota failures. Standalone GZIP text remains supported. Decompression checks cancellation between bounded reads; individual codec calls are not forcibly interruptible.
- ZIP entry reads validate declared size and CRC32 when read through EOF, including asynchronous reads. This is explicit because SharpCompress's raw `OpenEntryStream` does not provide that check. Partial searches/previews intentionally do not drain the remaining entry merely to validate CRC; they are not whole-archive integrity checks.
- Text/web archive previews retain strict entry integrity and expansion budgets, but the display-size cap is a soft truncation boundary. Oversized previews show their readable prefix with a notice explaining that more content exists and the preview setting can be increased. Complete exact-limit files are not labeled truncated; BOM-aware decoding avoids broken Unicode characters at the cut. Partial HTML/XML/JSON falls back to plain text. The reader consumes at most the display cap plus one lookahead byte. Encrypted entries and exposed link targets remain rejected; arbitrary archive paths are never passed to filesystem extraction.
- Added regression groups cover ZIP, solid 7z, TAR/PAX, compressed TAR aliases, headerless and standalone GZIP, Unicode names/content, count/search agreement, nested archives, preview reopening, hidden/system ZIP flags, malicious directory entries, encrypted ZIP/7z, TAR links, corrupted GZIP/ZIP checksums, expansion/entry/preview limits, cancellation and temporary-file cleanup.
- Pinned upstream RAR4/RAR5 and solid-RAR fixtures are read and hashed in memory; embedded executable data is never executed. See `Fixtures/Archives/NOTICE.txt` for provenance/license. A separate comparison of five upstream RAR fixtures (15 entries, including RAR5 BLAKE2) produced identical names, sizes and SHA-256 hashes with 0.38.0 and 0.50.4 in sequential sync/async passes.

Compressed-wrapper detection now spools the complete bounded wrapper before search, counting or preview, including standalone GZIP probing. This adds temporary-disk I/O and can reject a wrapper whose expanded TAR headers/padding exceed the configured total budget. It does not cache decompressed evidence between opens.

Before release, manually scan representative large archives and multipart RAR/7z sets, select previews after scanning, cancel/reset/close during extraction, and verify the packaged single-file application on the target machine. Automated fixtures reduce regression risk; they do not guarantee every archive variant or a bug-free release. Raw standalone BZIP2/XZ payloads (not TAR wrappers) are not newly claimed as supported.

## API references used during implementation

- Microsoft Learn: WebView2 DefaultBackgroundColor; XNode.ToString; XmlReaderSettings.DtdProcessing; FileSystemInfo.ResolveLinkTarget.
- MDN: CSS @layer precedence (ordinary author rules override layered defaults).
- SharpCompress 0.50.4 tagged source, package metadata/signature, upstream archive fixtures and the bundled 7-Zip help/upstream console listing implementation.
- Microsoft WPF ComboBox/TextBox template documentation, TabControl and TextBoxBase source for template/padding behavior, and Fonts.SystemFontFamilies for installed-font enumeration.
- Microsoft Learn: Timeline.DesiredFrameRate, SystemParameters.ClientAreaAnimation and Storyboard.Remove; Microsoft WPF UIElementAutomationPeer source for accessibility child enumeration.
