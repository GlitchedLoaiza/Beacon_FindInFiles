#requires -Version 7.0
[CmdletBinding()]
param(
	[Parameter(Mandatory)]
	[string] $Destination,
	[string] $PublishProfilePath,
	[switch] $ReadyToRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$project = Join-Path $repo 'Beacon_FindInFiles/Beacon_FindInFiles.vbproj'
$destinationPath = [IO.Path]::GetFullPath($Destination)
[xml] $projectXml = Get-Content -LiteralPath $project -Raw
$version = [string] $projectXml.Project.PropertyGroup.Version
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw 'Expected a stable three-part Beacon version.' }
if ($PublishProfilePath) {
	$PublishProfilePath = (Resolve-Path -LiteralPath $PublishProfilePath).Path
}
$destinationRoot = [IO.Path]::GetPathRoot($destinationPath)
if ($destinationPath.TrimEnd([IO.Path]::DirectorySeparatorChar) -eq $destinationRoot.TrimEnd([IO.Path]::DirectorySeparatorChar)) {
	throw 'Choose a dedicated application release directory, not a drive root.'
}
$staging = Join-Path $repo ('Beacon_FindInFiles/bin/Release/package-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
$publish = Join-Path $staging 'publish'
$bundle = Join-Path $staging 'bundle'
$noticeDirectory = Join-Path $bundle 'licenses'
New-Item -ItemType Directory -Path $noticeDirectory -Force | Out-Null

function Invoke-DotNetCapture {
	param([string[]] $Arguments, [string] $Log)
	$output = @(& dotnet @Arguments 2>&1 | ForEach-Object { $_.ToString() })
	$code = $LASTEXITCODE
	$output | Set-Content -LiteralPath $Log -Encoding utf8
	$output | ForEach-Object { Write-Host $_ }
	if ($code -ne 0) { throw "dotnet exited with $code; see $Log" }
	return ($output -join [Environment]::NewLine)
}

function Copy-RequiredFile {
	param([string] $Source, [string] $Target)
	if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) { throw "Required release file missing: $Source" }
	Copy-Item -LiteralPath $Source -Destination $Target
}

try {
	$arguments = @('publish', $project, '-c', 'Release', '-r', 'win-x64', '--self-contained', 'true',
		'-p:PublishSingleFile=true', '-p:PublishTrimmed=false', '-o', $publish, '--nologo', '-v:minimal')
	if ($PublishProfilePath) {
		$arguments += '-p:PublishProfileFullPath=' + $PublishProfilePath
	} else {
		$arguments += '-p:PublishReadyToRun=' + $ReadyToRun.IsPresent.ToString().ToLowerInvariant()
	}
	$null = Invoke-DotNetCapture $arguments (Join-Path $staging 'publish.log')
	$publishFiles = @(Get-ChildItem -LiteralPath $publish -File -Recurse)
	if ($publishFiles.Count -ne 1 -or $publishFiles[0].Name -ne 'Beacon.exe') {
		throw 'Raw publish must contain only Beacon.exe. Review unexpected dependencies before packaging.'
	}
	$exe = Join-Path $publish 'Beacon.exe'
	$fileInfo = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
	if ($fileInfo.FileVersion -ne "$version.0" -or $fileInfo.CompanyName -ne 'GlitchedLoaiza') {
		throw 'Published version/publisher branding does not match source metadata.'
	}

	$propertiesArgs = @('msbuild', $project, '-p:Configuration=Release', '-p:RuntimeIdentifier=win-x64', '-p:SelfContained=true',
		'-getProperty:TargetDir,IntermediateOutputPath,MSBuildProjectDirectory,PublishReadyToRun,Version', '-verbosity:quiet', '-nologo')
	if ($PublishProfilePath) { $propertiesArgs += '-p:PublishProfileFullPath=' + $PublishProfilePath }
	else { $propertiesArgs += '-p:PublishReadyToRun=' + $ReadyToRun.IsPresent.ToString().ToLowerInvariant() }
	$properties = (Invoke-DotNetCapture $propertiesArgs (Join-Path $staging 'publish-properties.json') | ConvertFrom-Json).Properties
	$runtimeConfig = Get-Content -LiteralPath (Join-Path $properties.TargetDir 'Beacon.runtimeconfig.json') -Raw | ConvertFrom-Json
	$frameworks = @($runtimeConfig.runtimeOptions.includedFrameworks)
	$assetsPath = Join-Path ([IO.Path]::GetFullPath((Join-Path $properties.MSBuildProjectDirectory $properties.IntermediateOutputPath))) 'project.assets.json'
	if (-not (Test-Path -LiteralPath $assetsPath)) { $assetsPath = Join-Path $properties.MSBuildProjectDirectory 'obj/project.assets.json' }
	$assets = Get-Content -LiteralPath $assetsPath -Raw | ConvertFrom-Json
	$packageRoots = @($assets.packageFolders.PSObject.Properties.Name)

	function Find-PackageDirectory {
		param([string] $Package, [string] $PackageVersion)
		$relative = $Package.ToLowerInvariant() + '/' + $PackageVersion.ToLowerInvariant()
		foreach ($packageRoot in $packageRoots) {
			$candidate = Join-Path $packageRoot $relative
			if (Test-Path -LiteralPath $candidate -PathType Container) { return $candidate }
		}
		throw "Restored package directory not found: $Package/$PackageVersion"
	}
	function Get-ResolvedVersion {
		param([string] $Package)
		$entry = @($assets.libraries.PSObject.Properties | Where-Object { $_.Name.StartsWith($Package + '/', [StringComparison]::OrdinalIgnoreCase) })
		if ($entry.Count -ne 1) { throw "Expected one resolved version of $Package" }
		return $entry[0].Name.Split('/')[1]
	}

	Copy-RequiredFile $exe (Join-Path $bundle 'Beacon.exe')
	Copy-RequiredFile (Join-Path $repo 'LICENSE') (Join-Path $bundle 'LICENSE')
	Copy-RequiredFile (Join-Path $repo 'THIRD-PARTY-NOTICES.md') (Join-Path $bundle 'THIRD-PARTY-NOTICES.md')
	Copy-RequiredFile (Join-Path $repo 'licenses/LGPL-2.1.txt') (Join-Path $noticeDirectory 'LGPL-2.1.txt')
	$sharpVersion = Get-ResolvedVersion 'SharpCompress'
	Copy-RequiredFile (Join-Path $repo "licenses/SharpCompress-$sharpVersion-LICENSE.txt") (Join-Path $noticeDirectory "SharpCompress-$sharpVersion-LICENSE.txt")
	$sevenVersion = Get-ResolvedVersion '7-Zip.CommandLine'
	$seven = Find-PackageDirectory '7-Zip.CommandLine' $sevenVersion
	$bundledHelper = Join-Path $properties.MSBuildProjectDirectory '7za.exe'
	if ((Get-FileHash -LiteralPath $bundledHelper).Hash -ne (Get-FileHash -LiteralPath (Join-Path $seven 'tools/x64/7za.exe')).Hash) {
		throw 'Bundled 7za.exe differs from the pinned NuGet helper.'
	}
	Copy-RequiredFile (Join-Path $seven 'tools/License.txt') (Join-Path $noticeDirectory '7-Zip-LICENSE.txt')
	Copy-RequiredFile (Join-Path $seven 'tools/readme.txt') (Join-Path $noticeDirectory '7-Zip-README.txt')
	$webVersion = Get-ResolvedVersion 'Microsoft.Web.WebView2'
	$web = Find-PackageDirectory 'Microsoft.Web.WebView2' $webVersion
	Copy-RequiredFile (Join-Path $web 'LICENSE.txt') (Join-Path $noticeDirectory 'WebView2-SDK-LICENSE.txt')
	Copy-RequiredFile (Join-Path $web 'NOTICE.txt') (Join-Path $noticeDirectory 'WebView2-SDK-NOTICE.txt')
	foreach ($framework in $frameworks) {
		$pack = Find-PackageDirectory ($framework.name + '.Runtime.win-x64') $framework.version
		$licenseName = if ($framework.name -eq 'Microsoft.NETCore.App') { 'LICENSE.TXT' } else { 'LICENSE' }
		Copy-RequiredFile (Join-Path $pack $licenseName) (Join-Path $noticeDirectory ($framework.name + '-LICENSE.txt'))
		$thirdParty = Join-Path $pack 'THIRD-PARTY-NOTICES.TXT'
		if (Test-Path -LiteralPath $thirdParty) {
			Copy-RequiredFile $thirdParty (Join-Path $noticeDirectory ($framework.name + '-THIRD-PARTY-NOTICES.txt'))
		}
	}

	$zipName = "Beacon-$version-win-x64.zip"
	$zipPath = Join-Path $staging $zipName
	[IO.Compression.ZipFile]::CreateFromDirectory($bundle, $zipPath, [IO.Compression.CompressionLevel]::Optimal, $false)
	$expectedEntries = @{}
	foreach ($file in Get-ChildItem -LiteralPath $bundle -File -Recurse) {
		$relative = [IO.Path]::GetRelativePath($bundle, $file.FullName).Replace('\', '/')
		$expectedEntries[$relative] = (Get-FileHash -LiteralPath $file.FullName).Hash
	}
	$archive = [IO.Compression.ZipFile]::OpenRead($zipPath)
	try {
		if ($archive.Entries.Count -ne $expectedEntries.Count) { throw 'ZIP entry count does not match staged bundle.' }
		$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
		foreach ($entry in $archive.Entries) {
			if (-not $seen.Add($entry.FullName) -or -not $expectedEntries.ContainsKey($entry.FullName)) { throw 'Unexpected or duplicate ZIP entry.' }
			$stream = $entry.Open()
			try { $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream)) }
			finally { $stream.Dispose() }
			if ($hash -ne $expectedEntries[$entry.FullName]) { throw "ZIP content hash mismatch: $($entry.FullName)" }
		}
	} finally { $archive.Dispose() }

	$commit = & git -C $repo rev-parse HEAD
	if ($LASTEXITCODE -ne 0) { throw 'Could not record source commit.' }
	$branch = & git -C $repo symbolic-ref --quiet --short HEAD
	if ($LASTEXITCODE -ne 0) { throw 'Could not record source branch.' }
	$status = @(& git -C $repo status --porcelain=v1)
	if ($LASTEXITCODE -ne 0) { throw 'Could not record source state.' }
	$zipHash = (Get-FileHash -LiteralPath $zipPath).Hash
	$signature = (Get-AuthenticodeSignature -LiteralPath $exe).Status.ToString()
	$manifest = [ordered]@{
		Version = $version; FileVersion = $fileInfo.FileVersion; ProductVersion = $fileInfo.ProductVersion
		Company = $fileInfo.CompanyName; Signature = $signature; BuiltUtc = [DateTimeOffset]::UtcNow.ToString('O')
		SourceCommit = $commit; SourceBranch = $branch; SourceHasUncommittedChanges = ($status.Count -gt 0)
		Configuration = 'Release'; RuntimeIdentifier = 'win-x64'; SelfContained = $true; PublishSingleFile = $true
		PublishReadyToRun = [bool]::Parse($properties.PublishReadyToRun); RuntimeFrameworks = $frameworks
		Packages = [ordered]@{ SharpCompress = $sharpVersion; '7-Zip.CommandLine' = $sevenVersion; 'Microsoft.Web.WebView2' = $webVersion }
		Artifacts = @(
			[ordered]@{ File = 'Beacon.exe'; Bytes = (Get-Item -LiteralPath $exe).Length; SHA256 = $expectedEntries['Beacon.exe'] },
			[ordered]@{ File = $zipName; Bytes = (Get-Item -LiteralPath $zipPath).Length; SHA256 = $zipHash })
		ZipEntries = @($expectedEntries.Keys | Sort-Object)
	}
	$manifestPath = Join-Path $staging 'release-manifest.json'
	$manifest | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $manifestPath -Encoding utf8
	$checksums = Join-Path $staging 'SHA256SUMS.txt'
	$checksumLines = @(
		"$($expectedEntries['Beacon.exe'])  Beacon.exe"
		"$zipHash  $zipName"
	)
	$checksumLines | Set-Content -LiteralPath $checksums -Encoding ascii
	$writtenChecksums = @(Get-Content -LiteralPath $checksums)
	if ($writtenChecksums.Count -ne 2) { throw 'Expected one checksum line per release artifact.' }
	for ($index = 0; $index -lt $checksumLines.Count; $index++) {
		if ($writtenChecksums[$index] -cne $checksumLines[$index] -or $writtenChecksums[$index] -notmatch '^[0-9A-F]{64}  [^\r\n]+$') {
			throw 'Written checksum file does not match the verified artifacts.'
		}
	}

	$running = @(Get-Process -Name Beacon -ErrorAction SilentlyContinue | Where-Object {
		$_.Path -and [StringComparer]::OrdinalIgnoreCase.Equals($_.Path, (Join-Path $destinationPath 'Beacon.exe')) })
	if ($running.Count -gt 0) { throw 'The destination Beacon.exe is running. Close it before replacing release files.' }
	New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
	$replacements = @{'Beacon.exe' = $exe; $zipName = $zipPath; 'SHA256SUMS.txt' = $checksums; 'release-manifest.json' = $manifestPath}
	foreach ($name in $replacements.Keys) {
		$existing = Join-Path $destinationPath $name
		if (Test-Path -LiteralPath $existing) {
			$backup = Join-Path $staging 'previous-release'
			New-Item -ItemType Directory -Path $backup -Force | Out-Null
			Copy-Item -LiteralPath $existing -Destination (Join-Path $backup $name)
		}
	}
	foreach ($name in $replacements.Keys) {
		$temporary = Join-Path $destinationPath ($name + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
		try {
			Copy-Item -LiteralPath $replacements[$name] -Destination $temporary
			if ((Get-FileHash -LiteralPath $temporary).Hash -ne (Get-FileHash -LiteralPath $replacements[$name]).Hash) { throw 'Release copy verification failed.' }
			[IO.File]::Move($temporary, (Join-Path $destinationPath $name), $true)
		} finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary } }
	}
	foreach ($name in $replacements.Keys) {
		if ((Get-FileHash -LiteralPath (Join-Path $destinationPath $name)).Hash -ne (Get-FileHash -LiteralPath $replacements[$name]).Hash) { throw 'Destination validation failed.' }
	}
	Write-Host "Verified release package: $(Join-Path $destinationPath $zipName)"
	Write-Host "Staging, logs and backups retained at: $staging"
	[pscustomobject]@{ Destination = $destinationPath; Staging = $staging; ExeSHA256 = $expectedEntries['Beacon.exe']; ZipSHA256 = $zipHash }
} catch {
	Write-Warning "Packaging stopped. Staging/evidence was retained at $staging. Existing files, if replaced, are backed up under previous-release."
	throw
}
