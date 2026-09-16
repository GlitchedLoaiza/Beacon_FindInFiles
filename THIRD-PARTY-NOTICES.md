# Beacon third-party components

Beacon's own source is licensed under [MIT](LICENSE), copyright GlitchedLoaiza. That license does not replace the licenses of bundled third-party software. Preserve the component license files when redistributing Beacon. No listed third party endorses this application.

The release packaging script copies the license/notice files for the exact restored packages and runtime packs into the ZIP's `licenses` directory. They are legal notices, not the WebView2 API documentation XML files excluded from the publish output. The WebView2 browser runtime is not bundled; the .NET runtime and WebView2 SDK assemblies/native loader are bundled in the self-contained application.

| Component | Version in this source | License / included material | Source |
| --- | --- | --- | --- |
| SharpCompress | 0.50.4 | MIT; `licenses/SharpCompress-0.50.4-LICENSE.txt` | https://github.com/adamhathcock/sharpcompress/tree/c083c6efd843a844b0c8f7878787360e815be781 |
| 7-Zip Extra (`7za.exe`, x64) | 25.01, distributed by `7-Zip.CommandLine` 25.1.0 | LGPL 2.1 or later and BSD components; package `License.txt`, `readme.txt`, and the full LGPL 2.1 text | https://www.7-zip.org/ ; corresponding source https://www.7-zip.org/a/7z2501-src.7z |
| Microsoft.Web.WebView2 SDK | 1.0.2792.45 | Package `LICENSE.txt` and `NOTICE.txt` | https://www.nuget.org/packages/Microsoft.Web.WebView2/1.0.2792.45 |
| .NET runtime and Windows Desktop runtime | Recorded from the published runtime configuration in `release-manifest.json` | Runtime pack license and third-party notices where supplied | https://github.com/dotnet/runtime ; https://github.com/dotnet/wpf ; https://github.com/dotnet/winforms |

## 7-Zip distribution

Beacon embeds the unmodified x64 `7za.exe` from the pinned package, extracts it to a Beacon-owned versioned temporary location when CAB handling is needed, and invokes it as a separate process. The packaging script compares its SHA-256 with that package's binary. The package's notice credits Igor Pavlov, Facebook, and Yann Collet for the relevant LGPL/BSD code. The standalone `7za.exe` is not the full `7z.exe`/`7z.dll` distribution; retain its actual package license information rather than substituting terms for a different binary.

The corresponding 25.01 source is available at the official URL above. If redistributing, preserve that source-access information and the license materials and review the applicable LGPL source-availability obligations. This file is a notice, not a new restriction on rights granted by those licenses or a legal certification.

## Maintainer procedure

- Run `scripts/Publish-Beacon.ps1` to assemble the application ZIP and notice files from the restored dependency versions.
- The script stops if it cannot find the required license materials; do not substitute an EXE-only ZIP for a public release without handling notices separately.
- Keep dependency versions, notices, source links, artifact hashes, and release notes in sync when updating components.
- Benchmark/test dependencies and pinned RAR fixtures are development-only and must not be included in the application ZIP. The fixture-specific notice remains under `tests/Beacon.SafetyChecks/Fixtures/Archives/NOTICE.txt`.
- If signing is introduced later, sign and timestamp the final executable before packaging and regenerate every executable/ZIP checksum. Beacon is currently unsigned; publisher branding is not an Authenticode identity.
