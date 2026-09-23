param([Parameter(Mandatory = $true)][string]$EvidencePath)

$ErrorActionPreference = 'Stop'
$EvidencePath = [IO.Path]::GetFullPath($EvidencePath)
$environment = Get-Content -Raw (Join-Path $EvidencePath 'environment.json') | ConvertFrom-Json
$report = Join-Path $EvidencePath 'bdn/results/BeaconProductionScalingBenchmarks-report.csv'
$rows = @(Import-Csv $report)
$workerCounts = @(-1, 0, 1, 2, 4, 8, 14, [int]$environment.LogicalProcessors) | Where-Object { $_ -le $environment.LogicalProcessors } | Sort-Object -Unique
if ($rows.Count -ne 6 * $workerCounts.Count) { throw "Expected the complete six-workload matrix, got $($rows.Count) rows." }

function Milliseconds([string]$value) {
	$match = [regex]::Match($value, '^\s*([\d.,]+)\s*(ns|us|µs|μs|ms|s)\s*$')
	if (-not $match.Success) { throw "Missing or invalid benchmark duration: '$value'." }
	$number = [double]::Parse($match.Groups[1].Value, [Globalization.NumberStyles]::Float -bor [Globalization.NumberStyles]::AllowThousands, [Globalization.CultureInfo]::InvariantCulture)
	$factor = switch ($match.Groups[2].Value) { 'ns' { 0.000001 }; 'us' { 0.001 }; 'µs' { 0.001 }; 'μs' { 0.001 }; 'ms' { 1.0 }; 's' { 1000.0 } }
	return $number * $factor
}

function Average($items, [string]$property) {
	return ($items | Measure-Object -Property $property -Average).Average
}

$telemetry = @(Get-ChildItem (Join-Path $EvidencePath 'telemetry') -Filter '*.jsonl' -File | ForEach-Object {
	Get-Content $_.FullName | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_ | ConvertFrom-Json }
})
$baselines = @{}
$forcedOne = @{}
foreach ($row in $rows) {
	if ([int]$row.Workers -eq -1) { $baselines[$row.Workload] = Milliseconds $row.Mean }
	if ([int]$row.Workers -eq 1) { $forcedOne[$row.Workload] = Milliseconds $row.Mean }
}
$comparison = @($rows | ForEach-Object {
	$row = $_
	$workers = [int]$row.Workers
	$samples = @($telemetry | Where-Object { $_.Workload -eq $row.Workload -and $_.Workers -eq $workers -and $_.Operation -ge 2 -and $_.Operation -le 6 })
	if ($samples.Count -ne 5) { throw "Expected five measured probe operations for $($row.Workload)/$workers; got $($samples.Count)." }
	$mean = Milliseconds $row.Mean
	$probe = @($samples | ForEach-Object { $_.Metrics })
	[pscustomobject][ordered]@{
		Workload = $row.Workload
		Workers = $workers
		Mode = if ($workers -eq -1) { 'OriginalSequential' } elseif ($workers -eq 0) { 'Automatic' } else { 'ForcedWorkers' }
		MeanMs = [Math]::Round($mean, 3)
		StdDevMs = [Math]::Round((Milliseconds $row.StdDev), 3)
		ErrorMs = [Math]::Round((Milliseconds $row.Error), 3)
		SpeedupVsSequential = [Math]::Round($baselines[$row.Workload] / $mean, 3)
		SpeedupVsForcedOne = [Math]::Round($forcedOne[$row.Workload] / $mean, 3)
		TimeReductionPercent = [Math]::Round(100 * (1 - $mean / $baselines[$row.Workload]), 2)
		ExpandedMiBPerSecond = [Math]::Round(($samples[0].ExpandedBytes / 1MB) / ($mean / 1000), 2)
		Files = $samples[0].Files
		PeakActiveJobs = ($samples | Measure-Object PeakActiveJobs -Maximum).Maximum
		ProbeAllocatedMiB = [Math]::Round((Average $probe 'AllocatedBytes') / 1MB, 2)
		SampledPeakWorkingSetMiB = [Math]::Round(($probe | Measure-Object PeakWorkingSetBytes -Maximum).Maximum / 1MB, 2)
		SampledPeakPrivateMiB = [Math]::Round(($probe | Measure-Object PeakPrivateBytes -Maximum).Maximum / 1MB, 2)
		CpuMs = [Math]::Round((Average $probe 'CpuMs'), 2)
		AverageBusyLogicalProcessors = [Math]::Round((Average $probe 'CpuMs') / (Average $probe 'ElapsedMs'), 2)
		BenchmarkAllocated = $row.Allocated
		BenchmarkGen0 = $row.Gen0
		BenchmarkGen1 = $row.Gen1
		BenchmarkGen2 = $row.Gen2
	}
})
$comparison | Export-Csv -NoTypeInformation -Encoding utf8 (Join-Path $EvidencePath 'comparison.csv')
$comparison | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $EvidencePath 'comparison.json')
$comparison | Where-Object { $_.Workers -in @(-1, 0, 1) } | Format-Table Workload,Mode,MeanMs,SpeedupVsSequential,PeakActiveJobs,SampledPeakWorkingSetMiB -AutoSize

$applicationPath = Join-Path $EvidencePath 'application/application.jsonl'
if (Test-Path $applicationPath) {
	$application = @(Get-Content $applicationPath | ForEach-Object { $_ | ConvertFrom-Json } | Where-Object { -not $_.Warmup })
	$groups = @($application | Group-Object { $_.Workload + '|' + $_.Workers })
	if ($groups.Count -ne 12) { throw "Expected both application modes for six workloads; got $($groups.Count) groups." }
	$appRows = @($groups | ForEach-Object {
		$values = @($_.Group)
		if ($values.Count -ne 3) { throw "Expected three measured application repetitions for $($_.Name)." }
		$mean = Average $values 'ScanMs'
		$sumSquares = ($values | ForEach-Object { [Math]::Pow($_.ScanMs - $mean, 2) } | Measure-Object -Sum).Sum
		[pscustomobject][ordered]@{
			Workload = $values[0].Workload
			Workers = $values[0].Workers
			Mode = $values[0].Mode
			MeanScanMs = [Math]::Round($mean, 3)
			StdDevScanMs = [Math]::Round([Math]::Sqrt($sumSquares / ($values.Count - 1)), 3)
			PeakActiveJobs = ($values | Measure-Object PeakActiveJobs -Maximum).Maximum
			AllocatedMiB = [Math]::Round((Average $values 'AllocatedBytes') / 1MB, 2)
			SampledPeakWorkingSetMiB = [Math]::Round(($values | Measure-Object SampledPeakWorkingSetBytes -Maximum).Maximum / 1MB, 2)
			FullPreviewMs = if ($null -eq $values[0].FullPreviewMs) { $null } else { Average $values 'FullPreviewMs' }
			SummarySwitchMs = if ($null -eq $values[0].SummarySwitchMs) { $null } else { Average $values 'SummarySwitchMs' }
			MaxDispatcherRoundTripMs = ($values | Measure-Object MaxDispatcherRoundTripMs -Maximum).Maximum
		}
	})
	$appRows | Export-Csv -NoTypeInformation -Encoding utf8 (Join-Path $EvidencePath 'application-comparison.csv')
	$appRows | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $EvidencePath 'application-comparison.json')
	$appRows | Format-Table Workload,Mode,MeanScanMs,PeakActiveJobs,FullPreviewMs,SummarySwitchMs -AutoSize
}
