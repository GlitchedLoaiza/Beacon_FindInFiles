# Beacon: Find in Files v2.0 - Advanced Log Search Utility

![Version](https://img.shields.io/badge/version-2.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Framework](https://img.shields.io/badge/.NET-10-purple)
![License](https://img.shields.io/badge/license-MIT-green)

A powerful Windows desktop application for searching text patterns across multiple file types including plain text files, Windows Event Logs (EVTX), HTTP Archive (HAR) files, and compressed archives.

## 🎯 Features

### Core Functionality
- **Multi-Format Search**: Search across text files, EVTX logs, HAR files, and archives in a single scan
- **Nested Archive Support**: Recursively search inside ZIP, 7Z, RAR, TAR, GZ, BZ2, and CAB files (up to 1 level deep)
- **Real-Time Highlighting**: Automatic search term highlighting in all preview modes
- **Smart Preview Modes**:
  - **Text Preview**: RichTextBox with inline highlighting for plain text files
  - **WebView2 Preview**: Rendered HTML/XML/JSON with syntax highlighting
  - **EVTX Preview**: Formatted Windows Event Log entries with full metadata
  - **HAR Preview**: HTTP request/response viewer with headers and payloads

### Advanced Features
- **Exact Match Support**: Word boundary matching for precise searches
- **Case-Sensitive Search**: Toggle case sensitivity on/off
- **Light/Dark Theme**: Automatic Windows theme detection with manual toggle
- **Async Scanning**: Non-blocking UI with real-time progress tracking
- **Cancellable Operations**: Stop long-running scans instantly (Esc key)
- **Keyboard Shortcuts**: Full keyboard navigation support
- **EVTX Resilience**: Graceful handling of missing message DLLs with raw XML fallback
- **Memory Optimized**: Intelligent throttling and match limits to prevent memory exhaustion

## 🖥️ System Requirements

- **OS**: Windows 11 (64-bit recommended)
- **Runtime**: None required - self-contained deployment includes .NET 10 runtime
- **WebView2 Runtime**: Optional - for HTML/XML/JSON preview (pre-installed on Windows 11, falls back to text mode if unavailable)

## 📦 Installation

### Option 1: Download Release (Recommended)
1. Download the latest release from the [Releases](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/releases) page
2. Extract the ZIP file to your preferred location
3. Run `Beacon_FindInFiles.exe`

**No installation required** - the application is self-contained with all dependencies bundled (including .NET 10 runtime and 7za.exe for CAB support).

## ⚠️ Windows SmartScreen Warning

Due to code signing costs for open-source projects, this release uses a self-signed certificate. Windows may show a "Windows protected your PC" warning.

**To run the application:**
1. Click "More info"
2. Click "Run anyway"

**Verify authenticity:**
- **SHA-256**: 308F7EA6464224D95A3F2A6570100A0CE25B79A23199A5861BF78A4B7FEF4AD0
- Compare with official release hash on GitHub

We're working on obtaining a trusted certificate for future releases.

### Option 2: Build from Source
```sh
# Clone the repository
git clone https://github.com/GlitchedLoaiza/Beacon_FindInFiles.git

# Open in Visual Studio 2022+
cd Beacon_FindInFiles
start Beacon_FindInFiles.sln

# Build solution (Ctrl+Shift+B)
# Run (F5)
```

## 🚀 Usage

### Basic Search
1. Click **Browse Folder** to select a directory, or **Browse Archive** to select an archive file
2. Enter your **search term** in the Search box
3. Click **Scan (Enter)** or press `Enter`
4. Select results from the list to preview matches
5. Use **F3** to navigate between matches

### Search Options
- **Case Sensitive**: Match exact letter casing
- **Exact Match**: Use word boundaries (whole word matching)

### Keyboard Shortcuts
| Shortcut | Action |
|----------|--------|
| `Enter` | Start scan |
| `Esc` | Cancel active scan |
| `F3` | Find next match in preview |
| `Shift+F3` | Find previous match (EVTX/HAR only) |
| `Ctrl+Down` | Select next file in results |
| `Ctrl+R` | Reset all fields |
| `🌙/☀️ Button` | Toggle dark/light theme |

## 📄 Supported File Formats

### Text Files
`.txt`, `.log`, `.json`, `.xml`, `.csv`, `.html`, `.reg`, `.ini`, `.cfg`, `.config`, `.nfo`

### Log Files
- **EVTX**: Windows Event Log files (`.evtx`)
- **HAR**: HTTP Archive files (`.har`)

### Archives
`.zip`, `.7z`, `.rar`, `.tar`, `.gz`, `.bz2`, `.tgz`, `.tbz`, `.tbz2`, `.tar.gz`, `.tar.bz2`, `.tar.xz`, `.txz`, `.cab`

**Note**: Archives can be nested up to 1 level deep (e.g., `diagnostic.zip » logs.cab » app.log`)

## 🔧 Technical Details

### Architecture
- **Language**: Visual Basic .NET
- **Framework**: WPF (Windows Presentation Foundation) targeting .NET 10
- **UI Engine**: WebView2 for HTML/XML/JSON rendering (optional)
- **Archive Library**: SharpCompress (multi-format support) + bundled 7za.exe for CAB files
- **EVTX Reader**: System.Diagnostics.Eventing.Reader

### Performance Optimizations
- **Async/Await**: Non-blocking file I/O and UI updates
- **UI Throttling**: Limits label updates to 120ms intervals during scans
- **Match Limiting**: Max 300 events/requests per EVTX/HAR file
- **Pre-Filtering**: XML-based search before expensive EVTX formatting (2-10x performance improvement)
- **Streaming**: Line-by-line text processing to minimize memory usage

### EVTX Handling
The application gracefully handles missing Windows Event Log message DLLs:
- Attempts to use `FormatDescription()` for human-readable messages
- Falls back to raw XML when DLLs are unavailable
- Displays warning messages to inform users about missing DLLs
- Never crashes due to missing event log resources

## 🎨 Theme Support

Beacon automatically detects your Windows theme preference and applies matching colors:
- **Dark Mode**: VS Code-inspired dark palette (#1E1E1E background)
- **Light Mode**: Clean white palette with high contrast

Manually toggle themes using the `🌙`/`☀️` button in the top-right corner.

## 📝 Use Cases

### IT Support & Diagnostics
- Search error codes across multiple log files
- Analyze Windows Event Logs for specific Event IDs
- Extract HTTP requests from HAR files captured in browser DevTools

### Development & Debugging
- Find configuration values across multiple `.config` files
- Search JSON/XML data files for specific keys
- Locate stack traces in application logs

### Security & Compliance
- Search for IP addresses or user IDs across log archives
- Audit event logs for specific security events
- Review HTTP traffic captured in HAR files

## 🐛 Known Limitations

- **Archive Depth**: Maximum 1 level of nested archives
- **Match Limits**: 300 matches per EVTX/HAR file to prevent memory issues

## ❓ FAQ

**Q: Why does Windows SmartScreen/Defender block the app?**  
A: Self-signed executable triggers SmartScreen warnings. Verify SHA-256 hash matches the official release and click "Run anyway". We're working on obtaining a trusted code signing certificate.

**Q: Does this require .NET 10 installation?**  
A: No - self-contained deployment includes the .NET 10 runtime.

**Q: Can I scan network drives?**  
A: Yes, but performance depends on network speed and latency.

**Q: What is the maximum file size supported?**  
A: No hard limit, but very large files (>1GB) may slow down scanning. Archives are processed efficiently with streaming.

**Q: Why can't I find matches in EVTX files?**  
A: Ensure you're searching for exact event ID numbers or message content. Some events may require message DLLs to display properly (fallback to raw XML is automatic).

**Q: Does this work on Windows 10?**  
A: The application is optimized for Windows 11. Windows 10 may work but requires separate WebView2 Runtime installation for HTML/XML preview.

## 🤝 Contributing & Feedback

This project is developed and maintained by:
- **GlitchedLoaiza** - Lead Developer
- **sgtxjosue** - HAR Module & Debugging

For bug reports, feature requests, or general feedback:
- Open an issue on [GitHub Issues](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/issues)
- Start a discussion on [GitHub Discussions](https://github.com/GlitchedLoaiza/Beacon_FindInFiles/discussions)

## 📜 License

MIT License

Copyright (c) 2025 GlitchedLoaiza

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## 🙏 Acknowledgments

- **SharpCompress** - Multi-format archive support
- **7-Zip** (7za.exe) - CAB file extraction (bundled)
- **WebView2** - Modern web rendering engine

---

**Made with ❤️ for log analysis workflows**
