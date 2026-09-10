# Beacon safeguard regression checks

This Windows/.NET 10 console runner compiles production code and WPF views as linked source files, using the same WebView2 and SharpCompress versions as Beacon. Checks create isolated off-screen WPF windows on an STA thread. Source/navigation tests instantiate MainWindow with its application startup/shutdown handlers detached and use temporary fixtures, not customer logs. The runner does not alter persisted settings, ownership, or permissions. Browser checks require the installed WebView2 Runtime and use a separate temporary profile.

From the repository root:

```powershell
dotnet run --project tests/Beacon.SafetyChecks/Beacon.SafetyChecks.vbproj -- "$PWD/Beacon_FindInFiles/7za.exe"
```

A nonzero exit code indicates a failure. This is an executable regression runner, not a Test Explorer/MSTest project. Solution builds compile it; run the command above to execute the checks.

## Coverage

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
- Real WebView2 highlighting: phrases spanning inline elements, UTF-16 offsets, .NET lookbehind ranges, script/hidden-text exclusion, bounded capture, document preservation and rejection of stale or changed-DOM snapshots.
- Structured diagnostics: concurrent deduplication, repeat counts, capacity/truncation reporting and detached snapshots.
- CSV/JSON/HTML reports: original-query preservation, optional excerpt/technical-data omission, Unicode/quoting, CSV formula protection, HTML encoding, source overwrite prevention and cancellation/locked-destination preservation.
- Diagnostics UI: light/dark contrast, filtering, Settings-controlled technical details, refresh, minimum-size layout and absence of redundant export/options controls. Tests do not use the system clipboard or native save dialog.
- Captured five-line context: file boundaries, first-match lookahead, record-local event/request context, long-line shortening, read-limit recovery, detached copies and export after the original file is deleted. The generated highlighted report is also loaded in WebView2.

## Step 7 search behavior

- **Literal text** treats regex punctuation literally. **Whole word** applies Unicode word boundaries to the literal term; there is no duplicate Exact match checkbox.
- **Regular expression** uses .NET syntax, multiline anchors and a 100 ms timeout per matching operation. Invalid patterns are rejected before scanning. Queries are limited to 4,096 characters.
- **Any term / All terms** split on whitespace and support double-quoted phrases, for example `"connection failed" timeout`. All terms must occur in the same text line, event, HAR request or single name/path target. Up to 32 terms are accepted.
- HTML/XML/JSON are searched as normalized extracted document text (one record), consistent with the existing document search behavior. Script/style text is removed by the existing lightweight extraction approach; this is not a full HTML parser. Rendered/pretty-printed preview text can differ from the searchable representation, so source counts and visible highlight counts need not be identical.
- Selected contents, filename and full-path targets are alternatives. Metadata-only matches do not require parsing an EVTX or unsupported binary file. Nested paths retain original archive names rather than temporary filenames.
- **First matching record per file** remains in Settings. It stops after the first matching text line/event/request. Per-file detail limits and EVTX/HAR limits still apply. A compact count with `+` means results may be incomplete; its tooltip states why.
- **Match details** starts collapsed. Expand it to navigate to a source line, event or request. Visible highlight ranges are bounded at 10,000; zero-width regex matches can select a record without highlighting characters.
- A scan snapshots its query. Editing the next query does not reinterpret existing results. WebView2 consumes .NET-generated ranges rather than converting .NET regex into JavaScript regex. Stale browser work cannot replace a newer highlight count.
- Archive attribute filtering is applied when exposed by the reader. SharpCompress 0.38 does not implement `Attrib` for some formats (including ZIP); unavailable attributes are not treated as an archive-read failure. Path, encryption, link-target and byte-budget checks remain in place.

Step 7 covers search; reporting is described below. EVTX provider acquisition, HAR-specific payload/redaction features and project-wide service extraction remain later roadmap work. Native EVTX reading/provider-DLL behavior still needs authorized real-file smoke tests; automated event navigation uses synthetic records.

## Step 8 diagnostics and reports

- **Export HTML report** opens an HTML-only save dialog directly. It saves a readable report grouped by filename with source paths, match locations, jump links, highlighted snippets and a compact search summary. There is no intermediate export-options screen.
- **Diagnostics** opens a separate problems list, with source/message filtering, Refresh and Copy all diagnostics. It explains skipped or incomplete reads and contains no export controls. Opening it does not stop a scan.
- Context is captured during scanning, not reread during export. Text snippets normally show two lines before and after the matched line, shifting at file boundaries to retain up to five available lines. First-match mode reads only the limited extra context needed, without collecting additional matches. EVTX/HAR context stays within the matched event/request; document context is labeled as extracted or normalized text.
- Each context line is limited to 500 characters, with explicit shortening around the match. Single-line records, short files and read limits may produce fewer than five lines. Older results without captured context show an excerpt and a request to rerun the search rather than inventing surrounding lines.
- Running, cancelled, failed and result-limit-stopped scans are explicitly incomplete. Per-file truncation, warnings/errors and omitted diagnostics also affect the coverage label. A completed report is complete only within its configured search scope.
- The diagnostic store keeps up to 10,000 distinct entries. Repeated identical notifications increment a repeat count; new notifications beyond capacity are counted as omitted. Source/message/technical fields are bounded and marked when truncated. Intentional extension/folder filters are not individually logged as errors.
- The existing Summary/Detailed setting controls stack-trace capture and technical diagnostic display. Basic issue messages remain available in either mode. There is no automatic on-disk diagnostic log.
- Diagnostic filters affect the view only. Copy all diagnostics includes the unfiltered list.
- **Copy result** copies the selected logical path and stored match locations/counts without excerpts. **Copy paths** copies distinct logical paths for current results, including archive paths. These commands do not copy full log contents.
- HTML export intentionally includes **unredacted matching text and context**. The button tooltip, save dialog and report warn about sharing this data. Reports state `RedactionApplied = false`; later HAR redaction settings are not implemented by this feature. Paths, queries and error messages can also be sensitive.
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
- Progress is an estimate; counting does not extract nested archives.
- CAB output is copied through bounded streams rather than letting the reader write arbitrary paths. The configured timeout includes listing and entry extraction.
- Reset and shutdown await the active scan before cleaning up its temporary files.
- Other settings/features from later roadmap phases are not completed by this work.

## Manual smoke checks still required

1. Display the reported HTML and XML samples in both light and dark modes, from disk and archives. Confirm document colors and search highlighting remain readable.
2. Confirm normal JSON and large HTML/XML previews still use WebView2; verify the preview-size warning with an oversized file.
3. Set small total/EVTX/HAR match limits; verify caps and partial-result status text.
4. Scan a nested CAB/ZIP, then select its results after scanning. Confirm retained files support preview navigation.
5. Cancel a scan, reset, and start another. Close Beacon during extraction. Confirm no stale results or deleted-in-use temporary files.
6. Exercise access-denied skip/prompt behavior manually on an authorized test directory. Do not change ownership on system or production data.

The checks instantiate actual WPF templates, selected source pipelines and a real WebView2 control, but do not assess screenshots or simulate physical keyboard/mouse input. Native EVTX/provider behavior, complete application startup/scan-cancel flows, elevated ownership recovery, representative large-document rendering and a manual visual/keyboard review still require release smoke testing.

## API references used during implementation

- Microsoft Learn: WebView2 DefaultBackgroundColor; XNode.ToString; XmlReaderSettings.DtdProcessing; FileSystemInfo.ResolveLinkTarget.
- MDN: CSS @layer precedence (ordinary author rules override layered defaults).
- SharpCompress 0.38.0 upstream IEntry interface and the bundled 7-Zip help/upstream console listing implementation.
- Microsoft WPF ComboBox/TextBox template documentation, TabControl and TextBoxBase source for template/padding behavior, and Fonts.SystemFontFamilies for installed-font enumeration.
- Microsoft Learn: Timeline.DesiredFrameRate, SystemParameters.ClientAreaAnimation and Storyboard.Remove; Microsoft WPF UIElementAutomationPeer source for accessibility child enumeration.
