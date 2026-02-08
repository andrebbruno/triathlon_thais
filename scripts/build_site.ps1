param(
  [string]$ReportsDir,
  [string]$SiteDir,
  [switch]$RebuildAll
)

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $ReportsDir) { $ReportsDir = Join-Path $repoRoot "Relatorios_Intervals" }
if (-not $SiteDir) { $SiteDir = Join-Path $repoRoot "docs" }

if (-not (Test-Path $ReportsDir)) {
  Write-Host "Reports directory not found: $ReportsDir"
  exit 1
}

$siteReports = Join-Path $SiteDir "reports"
New-Item -ItemType Directory -Path $siteReports -Force | Out-Null

Copy-Item -Path (Join-Path $ReportsDir "*.json") -Destination $siteReports -Force -ErrorAction SilentlyContinue
Copy-Item -Path (Join-Path $ReportsDir "*.md") -Destination $siteReports -Force -ErrorAction SilentlyContinue

$reportFiles = Get-ChildItem $ReportsDir -Filter "report_*.json" | Sort-Object Name -Descending
$analysisFiles = Get-ChildItem $ReportsDir -Filter "analysis_*.md" -ErrorAction SilentlyContinue | Sort-Object Name -Descending
$plannedFiles = Get-ChildItem $ReportsDir -Filter "planned_*.md" | Sort-Object Name -Descending
$longtermFiles = Get-ChildItem $ReportsDir -Filter "intervals_longterm_*coach_edition.json" | Sort-Object Name -Descending

# Notas semanais são usadas para a análise interna, não devem aparecer no relatório público.
$IncludeWeekNotesInReport = $false

function Html-Escape {
  param([string]$Text)
  if ($null -eq $Text) { return "" }
  $t = Fix-TextEncoding -Text $Text
  return [System.Net.WebUtility]::HtmlEncode($t)
}

function Should-RebuildReportHtml {
  param(
    [System.IO.FileInfo]$ReportFile,
    [string]$OutputPath,
    [bool]$Force
  )

  if ($Force) { return $true }
  if (-not $ReportFile) { return $false }
  if (-not (Test-Path $OutputPath)) { return $true }

  try {
    $outItem = Get-Item $OutputPath -ErrorAction Stop
    return ($ReportFile.LastWriteTimeUtc -gt $outItem.LastWriteTimeUtc)
  } catch {
    return $true
  }
}

function Fix-TextEncoding {
  param([string]$Text)
  if (-not $Text) { return "" }
  # Heuristic: most mojibake we see from API responses includes "Ã" / "Â" sequences.
  if ($Text -match "[ÃÂ]") {
    try {
      $fixed = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("Windows-1252").GetBytes($Text))
      if ($fixed -match "[ÃÂ]") {
        $fixed = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($Text))
      }
      return $fixed
    } catch { return $Text }
  }
  return $Text
}

function Is-OffPlannedEvent {
  param([object]$Event)

  if (-not $Event) { return $false }
  $name = Fix-TextEncoding -Text ([string]$Event.name)
  if ($name -and $name.Trim().ToUpperInvariant() -eq "OFF") { return $true }

  $desc = Fix-TextEncoding -Text ([string]$Event.description)
  if ($desc -and $desc -match "(?i)descanso") { return $true }

  return $false
}

function Get-PlanAdherenceSummary {
  param(
    [object[]]$PlannedEvents,
    [object[]]$Activities
  )

  $planned = if ($PlannedEvents) { @($PlannedEvents) } else { @() }
  $acts = if ($Activities) { @($Activities) } else { @() }

  $plannedOff = @($planned | Where-Object { Is-OffPlannedEvent -Event $_ })
  $plannedWorkouts = @($planned | Where-Object { -not (Is-OffPlannedEvent -Event $_) })

  $activityIds = New-Object System.Collections.Generic.HashSet[string]
  foreach ($a in $acts) {
    if ($a.id) { [void]$activityIds.Add([string]$a.id) }
  }

  $matchedEventIds = New-Object System.Collections.Generic.HashSet[string]
  foreach ($a in $acts) {
    if ($a.planejado -and $a.planejado.event_id) {
      [void]$matchedEventIds.Add([string]$a.planejado.event_id)
    }
  }

  $doneWorkouts = 0
  $missedWorkouts = @()
  foreach ($p in $plannedWorkouts) {
    $eventId = [string]$p.event_id
    $pairedId = [string]$p.paired_activity_id
    $matched = $false

    if ($pairedId -and $activityIds.Contains($pairedId)) { $matched = $true }
    elseif ($eventId -and $matchedEventIds.Contains($eventId)) { $matched = $true }

    if ($matched) { $doneWorkouts += 1 } else { $missedWorkouts += $p }
  }

  $offRespected = 0
  $offBroken = 0
  $offBrokenDates = @()
  foreach ($p in $plannedOff) {
    $d = [string]$p.start_date
    $hasActivity = $false
    if ($d) {
      $hasActivity = (@($acts | Where-Object { $_.start_date_local -eq $d } | Select-Object -First 1) -ne $null)
    }
    if (-not $hasActivity) { $offRespected += 1 } else { $offBroken += 1; if ($d) { $offBrokenDates += $d } }
  }

  $extras = @($acts | Where-Object { $_.planejado -eq $null })

  $overallValue = $null
  $workoutValue = $null
  if ($planned.Count -gt 0) {
    $overallValue = [math]::Round((($doneWorkouts + $offRespected) / $planned.Count) * 100, 1)
  }
  if ($plannedWorkouts.Count -gt 0) {
    $workoutValue = [math]::Round(($doneWorkouts / $plannedWorkouts.Count) * 100, 1)
  }

  return [PSCustomObject]@{
    planned_total     = $planned.Count
    planned_workouts  = $plannedWorkouts.Count
    planned_off       = $plannedOff.Count
    done_workouts     = $doneWorkouts
    missed_workouts   = $missedWorkouts
    off_respected     = $offRespected
    off_broken        = $offBroken
    off_broken_dates  = $offBrokenDates
    extras            = $extras
    adherence_overall = $overallValue
    adherence_workouts = $workoutValue
  }
}

function Format-Duration-Short {
  param([double]$Minutes)
  if ($Minutes -eq $null) { return "n/a" }
  $total = [math]::Round($Minutes, 0)
  if ($total -lt 60) { return "$total" + "min" }
  $h = [math]::Floor($total / 60)
  $m = $total % 60
  return "{0}h{1:00}min" -f $h, $m
}

function Pace-To-Secs {
  param([string]$Pace)
  if (-not $Pace) { return $null }
  $parts = $Pace -replace "/km","" -split ":"
  if ($parts.Length -ne 2) { return $null }
  return ([int]$parts[0] * 60) + [int]$parts[1]
}

function Read-AnalysisForWeek {
  param(
    [object[]]$Files,
    [string]$WeekStart,
    [string]$WeekEnd
  )
  if (-not $Files -or -not $WeekStart -or -not $WeekEnd) { return "" }
  $match = $Files | Where-Object { $_.Name -eq ("analysis_{0}_{1}.md" -f $WeekStart, $WeekEnd) } | Select-Object -First 1
  if (-not $match) { return "" }
  return (Get-Content $match.FullName -Raw)
}

function Get-MemorySectionLines {
  param(
    [string]$Text,
    [string]$HeaderPattern
  )
  if (-not $Text) { return @() }
  $match = [regex]::Match($Text, "$HeaderPattern\s*([\s\S]*?)(?:\n---|\n## )")
  if (-not $match.Success) { return @() }
  $block = $match.Groups[1].Value
  return @($block -split "`n" | Where-Object { $_ -match "^\s*-\s+" } | ForEach-Object { ($_ -replace "^\s*-\s*", "").Trim() } | Where-Object { $_ })
}

function Get-MemoryCalendar {
  param([string]$Text)
  if (-not $Text) { return @() }
  $match = [regex]::Match($Text, "## 2\. CALENDARIO 2026([\s\S]*?)(?:\n---|\n## )|## 2\. CALENDÁRIO 2026([\s\S]*?)(?:\n---|\n## )")
  if (-not $match.Success) { return @() }
  $block = if ($match.Groups[1].Value) { $match.Groups[1].Value } else { $match.Groups[2].Value }
  $lines = $block -split "`n" | Where-Object { $_ -match "^\|" }
  $rows = @()
  foreach ($line in $lines) {
    if ($line -match "^\|\s*Data\s*\|") { continue }
    if ($line -match "^\|\s*-+") { continue }
    $cols = $line.Trim("|") -split "\|"
    if ($cols.Count -lt 5) { continue }
    $dateRaw = (Fix-TextEncoding $cols[0].Trim())
    $name = (Fix-TextEncoding $cols[1].Trim())
    $type = (Fix-TextEncoding $cols[2].Trim())
    $priority = $cols[3].Trim()
    $status = $cols[4].Trim()
    $dateObj = $null
    if ($dateRaw -match "\d{2}/\d{2}/\d{4}") {
      try { $dateObj = [DateTime]::ParseExact($dateRaw, "dd/MM/yyyy", $null) } catch { $dateObj = $null }
    } elseif ($dateRaw -match "\d{2}/\d{2}") {
      try { $dateObj = [DateTime]::ParseExact("$dateRaw/2026", "dd/MM/yyyy", $null) } catch { $dateObj = $null }
    }
    $rows += [PSCustomObject]@{
      date_raw = $dateRaw
      date = $dateObj
      name = $name
      type = $type
      priority = $priority
      status = $status
    }
  }
  return $rows
}

function Get-WellnessForDate {
  param(
    [object[]]$Wellness,
    [string]$Date
  )

  if (-not $Date) { return $null }
  return $Wellness | Where-Object { $_.data -eq $Date } | Select-Object -First 1
}

function Classify-Quality {
  param(
    [object]$Plan
  )

  if (-not $Plan) {
    return @{ label = "Sem plano"; level = "neutral" }
  }

  $dt = $Plan.delta_time_min
  $dd = $Plan.delta_distance_km
  $dtAbs = if ($dt -ne $null) { [math]::Abs([double]$dt) } else { $null }
  $ddAbs = if ($dd -ne $null) { [math]::Abs([double]$dd) } else { $null }

  if (($dtAbs -ne $null -and $dtAbs -le 5) -or ($ddAbs -ne $null -and $ddAbs -le 1)) {
    return @{ label = "No alvo"; level = "good" }
  }

  if ($dt -ne $null -and $dt -gt 5) {
    return @{ label = "Acima"; level = "warn" }
  }

  if ($dt -ne $null -and $dt -lt -5) {
    return @{ label = "Abaixo"; level = "bad" }
  }

  return @{ label = "Parcial"; level = "neutral" }
}

function Build-Insight {
  param(
    [object]$Activity,
    [object]$WellnessDay,
    [double]$AvgSleep,
    [double]$AvgHrv,
    [double]$AvgRhr
  )

  $notes = @()

  if ($WellnessDay) {
    if ($WellnessDay.sono_h -lt 6.5) { $notes += "Sono baixo pode elevar FC e reduzir qualidade." }
    elseif ($WellnessDay.sono_h -ge 7.5) { $notes += "Sono bom favorece execução." }

    if ($WellnessDay.hrv -lt ($AvgHrv - 3)) { $notes += "HRV abaixo da média: atenção a fadiga." }
    if ($WellnessDay.fc_reposo -gt ($AvgRhr + 3)) { $notes += "FC repouso acima da média: sinal de stress." }
  }

  if ($Activity.type -eq "Run" -and $Activity.notas -match "joelho") {
    $notes += "Joelho citado: manter volume protegido."
  }

  if ($notes.Count -eq 0) { return "Execução sem alertas claros." }
  return ($notes -join " ")
}

function Build-ReportHtml {
  param(
    [string]$ReportPath,
    [string]$OutputPath
  )

  $report = Get-Content $ReportPath -Raw | ConvertFrom-Json
  $range = "$($report.semana.inicio) a $($report.semana.fim)"

  $activities = @($report.atividades)
  $wellness = @($report.bem_estar)
  $planned = @()
  if ($report.PSObject.Properties.Name -contains "treinos_planejados") {
    $planned = @($report.treinos_planejados)
  }

  $totalTime = $report.semana.tempo_total_horas
  $totalDist = $report.semana.distancia_total_km
  $totalTss = $report.semana.carga_total_tss
  $ctl = $report.metricas.CTL
  $atl = $report.metricas.ATL
  $tsb = $report.metricas.TSB
  $ramp = $report.metricas.RampRate
  $peso = $report.metricas.peso_atual

  $distGroups = $activities | Group-Object type | ForEach-Object {
    $timeMin = ($_.Group | Measure-Object moving_time_min -Sum).Sum
    $distKm = ($_.Group | Measure-Object distance_km -Sum).Sum
    [PSCustomObject]@{
      type = $_.Name
      time_h = if ($timeMin) { [math]::Round($timeMin / 60, 2) } else { 0 }
      dist_km = if ($distKm) { [math]::Round($distKm, 1) } else { 0 }
    }
  } | Sort-Object type

  $distLabels = @($distGroups | ForEach-Object { $_.type })
  $distValues = @($distGroups | ForEach-Object { $_.time_h })

  $wellDates = @($wellness | ForEach-Object { $_.data })
  $ctlVals = @($wellness | ForEach-Object { $_.ctl })
  $atlVals = @($wellness | ForEach-Object { $_.atl })
  $sleepVals = @($wellness | ForEach-Object { $_.sono_h })
  $hrvVals = @($wellness | ForEach-Object { $_.hrv })
  $rhrVals = @($wellness | ForEach-Object { $_.fc_reposo })

  $avgSleep = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object sono_h -Average).Average), 2) } else { 0 }
  $avgHrv = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object hrv -Average).Average), 1) } else { 0 }
  $avgRhr = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object fc_reposo -Average).Average), 1) } else { 0 }
  $sleepDelta = if ($avgSleep -gt 0) { [math]::Round(($avgSleep - $idealSleep), 1) } else { 0 }
  $hrvDelta = if ($avgHrv -gt 0) { [math]::Round(($avgHrv - $baselineHrv), 1) } else { 0 }
  $rhrDelta = if ($avgRhr -gt 0) { [math]::Round(($avgRhr - $baselineRhr), 1) } else { 0 }

  $notesWeek = @()
  if ($report.PSObject.Properties.Name -contains "notas_semana") {
    $notesWeek = @($report.notas_semana)
  }

  $distLabelsJson = ConvertTo-Json $distLabels -Compress
  $distValuesJson = ConvertTo-Json $distValues -Compress
  $wellDatesJson = ConvertTo-Json $wellDates -Compress
  $ctlJson = ConvertTo-Json $ctlVals -Compress
  $atlJson = ConvertTo-Json $atlVals -Compress
  $sleepJson = ConvertTo-Json $sleepVals -Compress
  $hrvJson = ConvertTo-Json $hrvVals -Compress
  $rhrJson = ConvertTo-Json $rhrVals -Compress

  function Parse-PlanTarget {
    param(
      [string]$Description,
      [string]$Type
    )

    if (-not $Description) { return $null }

    if ($Type -eq "Ride") {
      $m = [regex]::Match($Description, "(\d+)\s*-\s*(\d+)\s*W", "IgnoreCase")
      if ($m.Success) { return @{ min = [int]$m.Groups[1].Value; max = [int]$m.Groups[2].Value; unit = "W" } }
      $m2 = [regex]::Match($Description, "(\d+)\s*W", "IgnoreCase")
      if ($m2.Success) { $v = [int]$m2.Groups[1].Value; return @{ min = $v; max = $v; unit = "W" } }
    }

    if ($Type -eq "Run") {
      $m = [regex]::Match($Description, "(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/km", "IgnoreCase")
      if ($m.Success) { return @{ min = $m.Groups[1].Value; max = $m.Groups[2].Value; unit = "pace" } }
      $m2 = [regex]::Match($Description, "(\d{1,2}:\d{2})/km", "IgnoreCase")
      if ($m2.Success) { $v = $m2.Groups[1].Value; return @{ min = $v; max = $v; unit = "pace" } }
    }

    return $null
  }

  function Pace-From-Activity {
    param(
      [double]$Minutes,
      [double]$DistanceKm
    )

    if (-not $DistanceKm -or $DistanceKm -le 0) { return $null }
    $pace = $Minutes / $DistanceKm
    $min = [math]::Floor($pace)
    $sec = [math]::Round(($pace - $min) * 60)
    if ($sec -eq 60) { $min += 1; $sec = 0 }
    return "{0}:{1:00}/km" -f $min, $sec
  }

  $activityRows = @()
  $activityCards = @()
  foreach ($a in ($activities | Sort-Object start_date_local)) {
    $plan = $a.planejado
    $planText = "Sem planejado"
    if ($plan) {
      $pt = if ($plan.moving_time_min -ne $null) { "$($plan.moving_time_min) min" } else { "n/a" }
      $pd = if ($plan.distance_km -ne $null) { "$($plan.distance_km) km" } else { "n/a" }
      $dt = if ($plan.delta_time_min -ne $null) { "$($plan.delta_time_min) min" } else { "n/a" }
      $dd = if ($plan.delta_distance_km -ne $null) { "$($plan.delta_distance_km) km" } else { "n/a" }
      $planText = "Plan: $pt | $pd | delta $dt | delta $dd"
    }

    $notes = ""
    if ($a.notas) { $notes = $a.notas }

    $wellDay = Get-WellnessForDate -Wellness $wellness -Date $a.start_date_local
    $sleepText = if ($wellDay) { "$($wellDay.sono_h)h" } else { "n/a" }
    $hrvText = if ($wellDay) { "$($wellDay.hrv)" } else { "n/a" }
    $rhrText = if ($wellDay) { "$($wellDay.fc_reposo)" } else { "n/a" }
    $quality = Classify-Quality -Plan $plan
    $insight = Build-Insight -Activity $a -WellnessDay $wellDay -AvgSleep $avgSleep -AvgHrv $avgHrv -AvgRhr $avgRhr

    $planTarget = Parse-PlanTarget -Description $plan.description -Type $a.type
    $actualTarget = $null
    if ($a.type -eq "Ride" -and $a.average_watts) {
      $actualTarget = "$($a.average_watts) W"
    } elseif ($a.type -eq "Run") {
      $paceActual = Pace-From-Activity -Minutes $a.moving_time_min -DistanceKm $a.distance_km
      if ($paceActual) { $actualTarget = $paceActual }
    }

    $planTargetText = "n/a"
    if ($planTarget) {
      if ($planTarget.unit -eq "W") { $planTargetText = "$($planTarget.min)-$($planTarget.max) W" }
      if ($planTarget.unit -eq "pace") { $planTargetText = "$($planTarget.min)-$($planTarget.max)" }
    }

    $activityCards += @"
<div class=""activity-card"">
  <div class=""activity-head"">
    <div>
      <div class=""activity-name"">$(Html-Escape $a.name)</div>
      <div class=""activity-date"">$(Html-Escape $a.start_date_local) · $(Html-Escape $a.type)</div>
    </div>
    <span class=""badge badge-$($quality.level)"">$($quality.label)</span>
  </div>
  <div class=""chips"">
    <span class=""chip"">Tempo: $($a.moving_time_min) min</span>
    <span class=""chip"">Dist: $($a.distance_km) km</span>
    <span class=""chip"">Sono: $sleepText</span>
    <span class=""chip"">HRV: $hrvText</span>
    <span class=""chip"">FC Rep: $rhrText</span>
  </div>
  <div class=""activity-plan"">$(Html-Escape $planText)</div>
  <div class=""activity-target""><strong>Alvo:</strong> $(Html-Escape $planTargetText) · <strong>Executado:</strong> $(Html-Escape $actualTarget)</div>
  <div class=""activity-insight"">$(Html-Escape $insight)</div>
  <div class=""activity-notes"">$(Html-Escape $notes)</div>
</div>
"@

    $activityRows += @"
<tr>
  <td>$(Html-Escape $a.start_date_local)</td>
  <td>$(Html-Escape $a.type)</td>
  <td>$(Html-Escape $a.name)</td>
  <td>$($a.moving_time_min) min</td>
  <td>$($a.distance_km) km</td>
  <td>$(Html-Escape $planText)</td>
  <td>$(Html-Escape $notes)</td>
</tr>
"@
  }

  $notesBlock = ""
  if ($IncludeWeekNotesInReport -and $notesWeek.Count -gt 0) {
    $lines = @()
    foreach ($n in $notesWeek) {
      $lines += "<div class=""note-item""><strong>$(Html-Escape $n.name)</strong><div>$(Html-Escape $n.description)</div></div>"
    }
    $notesHtml = $lines -join ""
    $notesBlock = "<section class=""card notes""><h2>Notas da Semana</h2>$notesHtml</section>"
  }

  $html = @"
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Relatorio Semanal - $range</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Sans:wght@300;400;600&display=swap" rel="stylesheet">
  <style>
    :root{
      --bg:#f4f3ef;
      --ink:#0f172a;
      --muted:#6b7280;
      --card:#ffffff;
      --accent:#1f8ef1;
      --accent-2:#16a34a;
      --accent-3:#f59e0b;
      --accent-4:#111827;
      --shadow:0 16px 45px rgba(15,23,42,0.12);
    }
    *{box-sizing:border-box}
    body{
      margin:0;
      font-family:"IBM Plex Sans",sans-serif;
      background:radial-gradient(circle at 20% 20%, #ffffff 0%, #f4f3ef 55%, #ede9e3 100%);
      color:var(--ink);
    }
    .wrap{max-width:1200px;margin:24px auto;padding:0 24px}
    .hero{
      background:linear-gradient(135deg,#0f172a 0%,#1f2937 100%);
      color:#f8fafc;
      padding:28px;
      border-radius:22px;
      position:relative;
      overflow:hidden;
      box-shadow:var(--shadow);
    }
    .hero:after{
      content:"";
      position:absolute;
      width:320px;height:320px;
      right:-80px;top:-140px;
      background:radial-gradient(circle,#1f8ef1 0%,rgba(31,142,241,0) 70%);
      opacity:.6;
    }
    .hero h1{font-family:"Space Grotesk",sans-serif;margin:0 0 6px 0;font-size:26px}
    .hero p{margin:0;color:#cbd5f5}
    .grid{display:grid;gap:16px}
    .grid-4{grid-template-columns:repeat(auto-fit,minmax(200px,1fr))}
    .section{margin-top:20px}
    .card{
      background:var(--card);
      border-radius:16px;
      padding:18px;
      box-shadow:var(--shadow);
    }
    .kpi{font-size:28px;font-weight:700;font-family:"Space Grotesk",sans-serif}
    .label{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.12em}
    .charts{display:grid;gap:16px;grid-template-columns:repeat(auto-fit,minmax(280px,1fr))}
    .note-item{padding:8px 0;border-bottom:1px dashed #e5e7eb}
    .activity-card{border:1px solid #e5e7eb;border-radius:14px;padding:14px;margin-top:12px;background:#fff}
    .activity-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}
    .activity-name{font-weight:600;font-family:"Space Grotesk",sans-serif}
    .activity-date{color:var(--muted);font-size:12px}
    .chips{display:flex;flex-wrap:wrap;gap:8px;margin:8px 0}
    .chip{background:#f8fafc;border-radius:999px;padding:4px 10px;font-size:11px;color:#334155}
    .badge{padding:4px 10px;border-radius:999px;font-size:11px;font-weight:600}
    .badge-good{background:#dcfce7;color:#166534}
    .badge-warn{background:#fef3c7;color:#92400e}
    .badge-bad{background:#fee2e2;color:#991b1b}
    .badge-neutral{background:#e2e8f0;color:#334155}
    .activity-plan{font-size:12px;color:#475569;margin-bottom:6px}
    .activity-target{font-size:12px;color:#111827;margin-bottom:6px}
    .activity-insight{font-size:12px;color:#0f172a;background:#f1f5f9;border-left:3px solid #1f8ef1;padding:8px;border-radius:8px}
    .activity-notes{font-size:12px;color:#64748b;margin-top:6px}
    table{width:100%;border-collapse:collapse;font-size:12px}
    th,td{padding:8px;border-bottom:1px solid #e5e7eb;text-align:left}
    th{color:var(--muted);font-weight:600}
  </style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <h1>Relatorio Semanal</h1>
    <p>$range</p>
  </div>

  <div class="grid grid-4 section">
    <div class="card"><div class="label">Tempo total</div><div class="kpi">$totalTime h</div></div>
    <div class="card"><div class="label">Distancia</div><div class="kpi">$totalDist km</div></div>
    <div class="card"><div class="label">Carga</div><div class="kpi">$totalTss</div></div>
    <div class="card"><div class="label">TSB</div><div class="kpi">$tsb</div></div>
  </div>
  <div class="grid grid-4 section">
    <div class="card"><div class="label">CTL</div><div class="kpi">$ctl</div></div>
    <div class="card"><div class="label">ATL</div><div class="kpi">$atl</div></div>
    <div class="card"><div class="label">RampRate</div><div class="kpi">$ramp</div></div>
    <div class="card"><div class="label">Peso</div><div class="kpi">$peso</div></div>
  </div>

  <div class="section charts">
    <div class="card"><h2>Distribuicao por modalidade</h2><canvas id="dist-chart"></canvas></div>
    <div class="card"><h2>CTL / ATL</h2><canvas id="pmc-chart"></canvas></div>
    <div class="card"><h2>Bem-estar diario</h2><canvas id="well-chart"></canvas></div>
  </div>

  $notesBlock

  <section class="card section">
    <h2>Qualidade por Sessao (planejado vs executado + wellness)</h2>
    $($activityCards -join "`n")
  </section>

  <section class="card section">
    <h2>Atividades (planejado vs executado)</h2>
    <table>
      <thead>
        <tr><th>Data</th><th>Tipo</th><th>Nome</th><th>Tempo</th><th>Distancia</th><th>Planejado</th><th>Notas</th></tr>
      </thead>
      <tbody>
        $($activityRows -join "`n")
      </tbody>
    </table>
  </section>
</div>
<script>
  new Chart(document.getElementById('dist-chart'),{
    type:'doughnut',
    data:{labels:$distLabelsJson,datasets:[{data:$distValuesJson,backgroundColor:['#1f8ef1','#16a34a','#f59e0b','#111827','#ef4444']}]},
    options:{plugins:{legend:{position:'bottom'}},cutout:'60%'}
  });
  new Chart(document.getElementById('pmc-chart'),{
    type:'line',
    data:{labels:$wellDatesJson,datasets:[
      {label:'CTL',data:$ctlJson,borderColor:'#1f8ef1',tension:.3},
      {label:'ATL',data:$atlJson,borderColor:'#f59e0b',tension:.3}
    ]},
    options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false}}}
  });
  new Chart(document.getElementById('well-chart'),{
    type:'line',
    data:{labels:$wellDatesJson,datasets:[
      {label:'Sono (h)',data:$sleepJson,borderColor:'#16a34a',tension:.3},
      {label:'HRV',data:$hrvJson,borderColor:'#1f8ef1',tension:.3},
      {label:'FC Repouso',data:$rhrJson,borderColor:'#ef4444',tension:.3}
    ]},
    options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false}}}
  });
</script>
</body>
</html>
"@

  Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}

function Build-ReportHtmlModern {
  param(
    [string]$ReportPath,
    [string]$OutputPath
  )

  function Format-Duration {
    param([double]$Minutes)
    if ($Minutes -eq $null) { return "n/a" }
    $total = [math]::Round($Minutes, 0)
    if ($total -lt 60) { return "$total" + "min" }
    $h = [math]::Floor($total / 60)
    $m = $total % 60
    return "{0}h{1:00}min" -f $h, $m
  }

  function Pace-To-Secs {
    param([string]$Pace)
    if (-not $Pace) { return $null }
    $parts = $Pace -replace "/km","" -split ":"
    if ($parts.Length -ne 2) { return $null }
    return ([int]$parts[0] * 60) + [int]$parts[1]
  }

  function Format-Pace {
    param([double]$MinutesPerKm)
    if ($MinutesPerKm -eq $null) { return "n/a" }
    $min = [math]::Floor($MinutesPerKm)
    $sec = [math]::Round(($MinutesPerKm - $min) * 60)
    if ($sec -eq 60) { $min += 1; $sec = 0 }
    return "{0}:{1:00}/km" -f $min, $sec
  }

  function Format-Pace100 {
    param([double]$MinutesPer100)
    if ($MinutesPer100 -eq $null) { return "n/a" }
    $min = [math]::Floor($MinutesPer100)
    $sec = [math]::Round(($MinutesPer100 - $min) * 60)
    if ($sec -eq 60) { $min += 1; $sec = 0 }
    return "{0}:{1:00}/100m" -f $min, $sec
  }

  function Parse-NumberFromText {
    param(
      [string]$Value,
      [double]$Default = 0
    )
    if (-not $Value) { return $Default }
    $clean = $Value -replace "[^0-9,\.]", ""
    if (-not $clean) { return $Default }
    $clean = $clean -replace ",", "."
    return [double]$clean
  }

  function Format-DeltaValue {
    param(
      [double]$Value,
      [string]$Unit
    )
    if ($Value -eq $null) { return "n/a" }
    $sign = if ($Value -gt 0) { "+" } elseif ($Value -lt 0) { "-" } else { "" }
    $abs = [math]::Abs([math]::Round($Value, 1))
    return "$sign$abs$Unit"
  }

  function Format-DeltaPace {
    param([double]$DeltaSeconds)
    if ($DeltaSeconds -eq $null) { return "n/a" }
    $sign = if ($DeltaSeconds -gt 0) { "+" } elseif ($DeltaSeconds -lt 0) { "-" } else { "" }
    $sec = [math]::Abs([math]::Round($DeltaSeconds, 0))
    $min = [math]::Floor($sec / 60)
    $s = $sec % 60
    return "{0}{1}:{2:00}/km" -f $sign, $min, $s
  }

  function Safe-Divide {
    param(
      [double]$Num,
      [double]$Den
    )
    if ($Den -and $Den -ne 0) { return $Num / $Den }
    return $null
  }

  function Format-Percent {
    param([double]$Value)
    if ($Value -eq $null) { return "n/a" }
    return "{0:0.0}%" -f $Value
  }

  function Build-MetricCard {
    param(
      [string]$Value,
      [string]$Label
    )
    return "<div class=""metric""><div class=""value"">$Value</div><div class=""label"">$Label</div></div>"
  }

  function Get-PerformanceScore {
    param([double]$CTL, [double]$TSB, [double]$RampRate)
    if ($CTL -eq $null -or $TSB -eq $null -or $RampRate -eq $null) { return $null }
    $ctlScore = [math]::Min($CTL / 100 * 40, 40)
    $tsbScore = if ($TSB -ge -5 -and $TSB -le 10) { 30 }
                elseif ($TSB -ge -15 -and $TSB -lt -5) { 20 }
                elseif ($TSB -gt 10 -and $TSB -le 20) { 20 }
                else { 10 }
    $rampScore = if ($RampRate -ge 2 -and $RampRate -le 6) { 30 }
                 elseif ($RampRate -ge 0 -and $RampRate -lt 2) { 20 }
                 elseif ($RampRate -gt 6 -and $RampRate -le 10) { 15 }
                 else { 5 }
    return [math]::Round($ctlScore + $tsbScore + $rampScore, 0)
  }

  function Get-RecoveryStatus {
    param([double]$TSB, [double]$HRV, [double]$RestingHR, [double]$Sleep)
    if ($TSB -eq $null -or $HRV -eq $null -or $RestingHR -eq $null -or $Sleep -eq $null) { return $null }
    $score = 0; $factors = @()
    if ($TSB -ge -5) { $score += 25; $factors += "TSB otimo" }
    elseif ($TSB -ge -15) { $score += 15; $factors += "TSB aceitavel" }
    else { $score += 5; $factors += "TSB critico" }

    if ($HRV -ge 45) { $score += 25; $factors += "HRV bom" }
    elseif ($HRV -ge 38) { $score += 15; $factors += "HRV medio" }
    else { $score += 5; $factors += "HRV baixo" }

    if ($RestingHR -le 52) { $score += 25; $factors += "FC repouso normal" }
    elseif ($RestingHR -le 58) { $score += 15; $factors += "FC repouso elevada" }
    else { $score += 5; $factors += "FC repouso alta" }

    if ($Sleep -ge 7.5) { $score += 25; $factors += "Sono adequado" }
    elseif ($Sleep -ge 6.5) { $score += 15; $factors += "Sono razoavel" }
    else { $score += 5; $factors += "Sono insuficiente" }

    $status = if ($score -ge 80) { "EXCELENTE" }
              elseif ($score -ge 60) { "BOM" }
              elseif ($score -ge 40) { "MODERADO" }
              else { "CRITICO" }
    return @{ score = $score; status = $status; factors = $factors }
  }

  function Get-WellnessFlags {
    param(
      [object]$WellDay,
      [double]$BaselineRhr,
      [double]$BaselineHrv,
      [double]$IdealSleep
    )
    $flags = @()
    if ($WellDay -and $WellDay.sono_h -ne $null -and $IdealSleep -gt 0 -and $WellDay.sono_h -lt ($IdealSleep - 0.5)) { $flags += "sono baixo" }
    if ($WellDay -and $WellDay.fc_reposo -ne $null -and $BaselineRhr -gt 0 -and $WellDay.fc_reposo -gt ($BaselineRhr + 5)) { $flags += "FC repouso alta" }
    if ($WellDay -and $WellDay.hrv -ne $null -and $BaselineHrv -gt 0 -and $WellDay.hrv -lt ($BaselineHrv - 5)) { $flags += "HRV baixa" }
    return $flags
  }

  function Parse-PlanTarget {
    param(
      [string]$Description,
      [string]$Type
    )

    if (-not $Description) { return $null }

    if ($Type -eq "Ride") {
      $m = [regex]::Match($Description, "(\d+)\s*-\s*(\d+)\s*W", "IgnoreCase")
      if ($m.Success) { return @{ min = [int]$m.Groups[1].Value; max = [int]$m.Groups[2].Value; unit = "W" } }
      $m2 = [regex]::Match($Description, "(\d+)\s*W", "IgnoreCase")
      if ($m2.Success) { $v = [int]$m2.Groups[1].Value; return @{ min = $v; max = $v; unit = "W" } }
    }

    if ($Type -eq "Run") {
      $m = [regex]::Match($Description, "(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/km", "IgnoreCase")
      if ($m.Success) { return @{ min = $m.Groups[1].Value; max = $m.Groups[2].Value; unit = "pace" } }
      $m2 = [regex]::Match($Description, "(\d{1,2}:\d{2})/km", "IgnoreCase")
      if ($m2.Success) { $v = $m2.Groups[1].Value; return @{ min = $v; max = $v; unit = "pace" } }
    }

    return $null
  }

  function Compare-Target {
    param(
      [object]$PlanTarget,
      [string]$Type,
      [double]$AvgWatts,
      [string]$ActualPace
    )

    if (-not $PlanTarget) { return "Sem alvo detalhado no plano." }
    if ($PlanTarget.unit -eq "W" -and $AvgWatts) {
      if ($AvgWatts -lt $PlanTarget.min) { return "Abaixo do alvo (" + [math]::Round($PlanTarget.min - $AvgWatts,0) + "W)" }
      if ($AvgWatts -gt $PlanTarget.max) { return "Acima do alvo (" + [math]::Round($AvgWatts - $PlanTarget.max,0) + "W)" }
      return "Dentro do alvo"
    }
    if ($PlanTarget.unit -eq "pace" -and $ActualPace) {
      $actualSecs = Pace-To-Secs -Pace $ActualPace
      $minSecs = Pace-To-Secs -Pace $PlanTarget.min
      $maxSecs = Pace-To-Secs -Pace $PlanTarget.max
      if ($actualSecs -gt $maxSecs) { return "Mais lento que o alvo" }
      if ($actualSecs -lt $minSecs) { return "Mais rápido que o alvo" }
      return "Dentro do alvo"
    }
    return "Sem dado executado comparavel."
  }

  function Estimate-Race {
    param(
      [string]$Type,
      [string]$Name,
      [double]$RunPace,
      [double]$BikeSpeed,
      [double]$SwimPace100
    )

    if (-not $RunPace -or $RunPace -le 0) { $RunPace = 6.5 }
    if (-not $BikeSpeed -or $BikeSpeed -le 0) { $BikeSpeed = 28 }
    if (-not $SwimPace100 -or $SwimPace100 -le 0) { $SwimPace100 = 2.2 }

    if ($Type -match "Corrida") {
      $dist = 10
      if ($Name -match "Meia") { $dist = 21.1 }
      $mins = $RunPace * $dist
      return Format-Duration-Short -Minutes $mins
    }
    if ($Type -match "Triathlon Sprint" -or $Name -match "Triathlon Sprint") {
      $swim = ($SwimPace100 * 7.5)
      $bike = if ($BikeSpeed -gt 0) { (20 / $BikeSpeed) * 60 } else { 50 }
      $run = $RunPace * 5
      return Format-Duration-Short -Minutes ($swim + $bike + $run + 4)
    }

    if ($Type -match "Ol" -or $Name -match "Olimp") {
      $swim = ($SwimPace100 * 15) # 1500m
      $bike = if ($BikeSpeed -gt 0) { (40 / $BikeSpeed) * 60 } else { 90 }
      $run = $RunPace * 10
      return Format-Duration-Short -Minutes ($swim + $bike + $run + 6)
    }

    if ($Name -match "70\.3" -or $Type -match "70\.3") {
      $swim = ($SwimPace100 * 19) # 1900m
      $bike = if ($BikeSpeed -gt 0) { (90 / $BikeSpeed) * 60 } else { 210 }
      $run = $RunPace * 21.1
      return Format-Duration-Short -Minutes ($swim + $bike + $run + 8)
    }

    return "n/a"
  }

  function Get-Objective {
    param(
      [object]$Plan,
      [object]$Activity
    )

    $text = ""
    if ($Plan) { $text += "$($Plan.name) $($Plan.description) " }
    $text += "$($Activity.name)"
    $text = $text.ToLower()

    if ($text -match "sweet spot|ss|tempo|threshold|limiar") {
      return "Sweet Spot para elevar limiar e sustentar potencia com controle."
    }
    if ($text -match "tiro|interval|vo2") {
      return "Intervalos para estimular VO2 e velocidade."
    }
    if ($text -match "endurance|z2|base|longao|longão|longo") {
      return "Base aerobica para eficiencia metabolica e resistencia."
    }
    # "Leve" costuma aparecer em Z2 (base) e em recuperacao. Se chegou ate aqui, nao caiu em Z2/base.
    if ($text -match "recuper|recovery|regener|solto|leve") {
      return "Recuperacao ativa para reduzir fadiga e manter circulacao."
    }
    if ($text -match "tecnica") {
      if ($Activity.type -eq "Swim") { return "Tecnica de nado para eficiencia e economia." }
      return "Tecnica para eficiencia de movimento."
    }
    if ($Activity.type -eq "Swim") { return "Natacao aerobica com foco tecnico." }
    if ($Activity.type -eq "Ride") { return "Sessao de ciclismo para evolucao aerobica." }
    if ($Activity.type -eq "Run") { return "Corrida protegida visando manter base sem sobrecarga." }
    if ($Activity.type -eq "Strength" -or $Activity.type -eq "WeightTraining") { return "Forca para estabilidade, prevencao de lesao e suporte ao triathlon." }
    return "Treino geral com foco em consistencia."
  }

  $report = Get-Content $ReportPath -Raw | ConvertFrom-Json
  $range = "$($report.semana.inicio) a $($report.semana.fim)"
  $referenceDate = [DateTime]::Parse($report.semana.fim)
  $analysisMd = Read-AnalysisForWeek -Files $analysisFiles -WeekStart $report.semana.inicio -WeekEnd $report.semana.fim
  $analysisBlock = ""
  if ($analysisMd) {
    $lines = $analysisMd -split "`n"
    $section = @()
    foreach ($line in $lines) {
      $trim = $line.Trim()
      if ($trim -match "^#") { continue }
      if ($trim -match "^\-\s+") {
        $item = $trim -replace '^\-\s*',''
        $itemEscaped = Html-Escape $item
        if ($itemEscaped -match "\*\*") {
          $itemEscaped = $itemEscaped -replace "\*\*(.+?)\*\*", '<strong>$1</strong>'
        }
        $section += "<li>$itemEscaped</li>"
      }
    }
    if ($section.Count -gt 0) {
      $analysisBlock = "<section class=""card section""><h2>Analise Semanal (resumo)</h2><ul>$($section -join '')</ul></section>"
    }
  }

  $activities = @($report.atividades)
  $wellness = @($report.bem_estar)
  $activityCount = $activities.Count

  $memoryPath = Join-Path $repoRoot "COACHING_MEMORY.md"
  $memoryText = if (Test-Path $memoryPath) { Get-Content $memoryPath -Raw } else { "" }
  $athleteName = if ($memoryText -match "\*\*Nome:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "Atleta" }
  $phaseTitle = if ($memoryText -match "\*\*Fase:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "Base Geral" }
  $phaseFocus = if ($memoryText -match "\*\*Foco principal:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $phaseRun = if ($memoryText -match "\*\*Run:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $phaseSwim = if ($memoryText -match "\*\*Swim:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $phaseForce = if ($memoryText -match "\*\*Força:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }

  function Get-MemoryNumber {
    param(
      [string]$Pattern,
      [double]$Default
    )
    if ($memoryText -match $Pattern) {
      return Parse-NumberFromText -Value $matches[1] -Default $Default
    }
    return $Default
  }

  $ftpBike = Get-MemoryNumber "\*\*FTP:\*\*\s*([0-9,\.]+)" 200
  $ftpRun = Get-MemoryNumber "\*\*FTP Run:\*\*\s*([0-9,\.]+)" 300
  $lthr = Get-MemoryNumber "\*\*LTHR:\*\*\s*([0-9,\.]+)" 165
  $baselineRhr = Get-MemoryNumber "\*\*FC Repouso baseline:\*\*\s*~?([0-9,\.]+)" 48
  $baselineHrv = Get-MemoryNumber "\*\*HRV baseline:\*\*\s*~?([0-9,\.]+)" 45
  $idealSleep = Get-MemoryNumber "\*\*Sono ideal:\*\*\s*([0-9,\.]+)" 7.5

  $thresholdPaceSecs = $null
  if ($memoryText -match "Threshold Pace:\s*~?([0-9:]+)/km") {
    $thresholdPaceSecs = Pace-To-Secs -Pace $matches[1]
  }
  if (-not $thresholdPaceSecs) { $thresholdPaceSecs = 330 }

  $phaseObjectives = @()
  if ($phaseFocus) { $phaseObjectives += $phaseFocus }
  if ($phaseRun) { $phaseObjectives += "Run: $phaseRun" }
  if ($phaseSwim) { $phaseObjectives += "Swim: $phaseSwim" }
  if ($phaseForce) { $phaseObjectives += "Força: $phaseForce" }

  $phaseTransition = @(
    "Introduzir intervalos de Threshold (bike)",
    "Tiros curtos na corrida (se joelho permitir)",
    "Natação: aumentar volume + séries de ritmo"
  )

  $futurePhases = Get-MemorySectionLines -Text $memoryText -HeaderPattern "### Pr[oó]ximas Fases \(planejado\)"
  $calendarEvents = Get-MemoryCalendar -Text $memoryText
  $mainEvent = $calendarEvents | Where-Object { $_.priority -match "A" -or $_.status -match "PROVA PRINCIPAL" } | Select-Object -First 1
  $mainEventName = if ($mainEvent) { $mainEvent.name } else { "Prova A" }
  $mainEventDateText = if ($mainEvent -and $mainEvent.date) { $mainEvent.date.ToString("dd/MM/yyyy") } elseif ($mainEvent) { $mainEvent.date_raw } else { "" }
  $daysToMain = if ($mainEvent -and $mainEvent.date) { ($mainEvent.date.Date - $referenceDate.Date).Days } else { $null }
  $daysToMainText = if ($daysToMain -ne $null) { "$daysToMain dias" } else { "n/a" }

  $phaseList = if ($futurePhases.Count -gt 0) { ($futurePhases | ForEach-Object { "<li>$_</li>" }) -join "" } else { "<li>Sem fases futuras cadastradas.</li>" }
  $longTermPlanBlock = ""
  if ($futurePhases.Count -gt 0 -or $mainEvent) {
    $longTermPlanBlock = @"
  <section class="card section">
    <h2>Plano de Longo Prazo</h2>
    <div class="plan-grid">
      <div class="plan-card">
        <h3>Rota até Prova A</h3>
        <div><strong>$mainEventName</strong> ${mainEventDateText}</div>
        <div style="margin-top:6px">Faltam <strong>$daysToMainText</strong> (referência: semana atual)</div>
      </div>
      <div class="plan-card">
        <h3>Calendário de fases</h3>
        <ul>$phaseList</ul>
      </div>
    </div>
    <div class="plan-note">As transições tendem a ocorrer nos períodos acima e podem ser ajustadas conforme resposta do corpo, joelho e proximidade da prova principal.</div>
  </section>
"@
  }

  $totalTime = $report.semana.tempo_total_horas
  $totalDist = $report.semana.distancia_total_km
  $totalTss = $report.semana.carga_total_tss
  $ctl = $report.metricas.CTL
  $atl = $report.metricas.ATL
  $tsb = $report.metricas.TSB
  $ramp = $report.metricas.RampRate
  $peso = $report.metricas.peso_atual
  $pesoText = if ($peso) { [math]::Round($peso, 1) } else { "n/a" }
  $pesoDisplay = if ($pesoText -eq "n/a") { "n/a" } else { "$pesoText kg" }

  $ctlText = if ($ctl -ne $null) { [math]::Round($ctl, 1) } else { "n/a" }
  $atlText = if ($atl -ne $null) { [math]::Round($atl, 1) } else { "n/a" }
  $tsbText = if ($tsb -ne $null) { [math]::Round($tsb, 1) } else { "n/a" }

  $ctlStart = ($wellness | Where-Object { $_.ctl -ne $null } | Select-Object -First 1).ctl
  $ctlEnd = ($wellness | Where-Object { $_.ctl -ne $null } | Select-Object -Last 1).ctl
  $atlStart = ($wellness | Where-Object { $_.atl -ne $null } | Select-Object -First 1).atl
  $atlEnd = ($wellness | Where-Object { $_.atl -ne $null } | Select-Object -Last 1).atl
  $ctlDelta = if ($ctlStart -ne $null -and $ctlEnd -ne $null) { [math]::Round(($ctlEnd - $ctlStart), 1) } else { $null }
  $atlDelta = if ($atlStart -ne $null -and $atlEnd -ne $null) { [math]::Round(($atlEnd - $atlStart), 1) } else { $null }

  $tsbStatus = "Equilibrado"
  $tsbClass = "tsb-ok"
  if ($tsb -le -25) { $tsbStatus = "Fadigado"; $tsbClass = "tsb-high" }
  elseif ($tsb -le -10) { $tsbStatus = "Cansado"; $tsbClass = "tsb-med" }
  elseif ($tsb -le 0) { $tsbStatus = "Controlado"; $tsbClass = "tsb-low" }
  elseif ($tsb -gt 0) { $tsbStatus = "Descansado"; $tsbClass = "tsb-good" }

  $plannedEvents = @()
  if ($report.PSObject.Properties.Name -contains "treinos_planejados") {
    $plannedEvents = @($report.treinos_planejados)
  }

  $adherence = Get-PlanAdherenceSummary -PlannedEvents $plannedEvents -Activities $activities
  $plannedCount = $adherence.planned_total
  $plannedWorkoutCount = $adherence.planned_workouts
  $plannedOffCount = $adherence.planned_off
  $doneWorkoutCount = $adherence.done_workouts
  $offRespectedCount = $adherence.off_respected
  $extraCount = @($adherence.extras).Count
  $missedWorkouts = @($adherence.missed_workouts)

  $complianceValue = $adherence.adherence_overall
  $workoutComplianceValue = $adherence.adherence_workouts
  $complianceText = if ($complianceValue -ne $null) { Format-Percent -Value $complianceValue } else { "n/a" }
  $workoutComplianceText = if ($workoutComplianceValue -ne $null) { Format-Percent -Value $workoutComplianceValue } else { "n/a" }
  $adherenceForLabel = if ($workoutComplianceValue -ne $null) { $workoutComplianceValue } else { $complianceValue }
  $complianceLabel = "Sem dados"
  $complianceClass = "comp-neutral"
  if ($adherenceForLabel -ne $null) {
    if ($adherenceForLabel -ge 85) { $complianceLabel = "Excelente"; $complianceClass = "comp-good" }
    elseif ($adherenceForLabel -ge 70) { $complianceLabel = "Bom"; $complianceClass = "comp-mid" }
    else { $complianceLabel = "Baixo"; $complianceClass = "comp-low" }
  }

  $plannedTimeMin = if ($plannedEvents.Count -gt 0) { ($plannedEvents | Measure-Object moving_time_min -Sum).Sum } else { $null }
  $plannedDistKm = if ($plannedEvents.Count -gt 0) { ($plannedEvents | Measure-Object distance_km -Sum).Sum } else { $null }
  $plannedTimeText = if ($plannedTimeMin -ne $null) { "{0:0.1}h" -f ($plannedTimeMin / 60) } else { "n/a" }
  $plannedDistText = if ($plannedDistKm -ne $null) { "{0:0.1}km" -f $plannedDistKm } else { "n/a" }
  $executedTimeText = if ($totalTime -ne $null) { "{0:0.1}h" -f $totalTime } else { "n/a" }
  $executedDistText = if ($totalDist -ne $null) { "{0:0.1}km" -f $totalDist } else { "n/a" }
  $complianceDetail = if ($plannedCount -gt 0) {
    $workoutPart = if ($plannedWorkoutCount -gt 0) { "Treinos: $doneWorkoutCount/$plannedWorkoutCount ($workoutComplianceText)" } else { "Treinos: n/a" }
    $offPart = if ($plannedOffCount -gt 0) { "Descanso: $offRespectedCount/$plannedOffCount" } else { "Descanso: n/a" }
    $extraPart = "Extras: $extraCount"
    "$workoutPart · $offPart · $extraPart. Planejado: $plannedCount sess | $plannedTimeText | $plannedDistText. Executado: $activityCount sess | $executedTimeText | $executedDistText."
  } else {
    "Sem planejamento importado nesta semana."
  }

  $ctlDeltaText = if ($ctlDelta -ne $null) { Format-DeltaValue -Value $ctlDelta -Unit "" } else { "n/a" }
  $atlDeltaText = if ($atlDelta -ne $null) { Format-DeltaValue -Value $atlDelta -Unit "" } else { "n/a" }

  $classification = "HOLD"
  $classCss = "hold"
  if ($tsb -le -25) { $classification = "STEP BACK"; $classCss = "stepback" }
  elseif ($tsb -le -10) { $classification = "HOLD"; $classCss = "hold" }
  else { $classification = "PUSH"; $classCss = "push" }

  $avgSleep = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object sono_h -Average).Average), 2) } else { 0 }
  $avgHrv = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object hrv -Average).Average), 1) } else { 0 }
  $avgRhr = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object fc_reposo -Average).Average), 1) } else { 0 }

  $dailyLoads = $activities | Group-Object start_date_local | ForEach-Object { ($_.Group | Measure-Object suffer_score -Sum).Sum }
  $meanLoad = if ($dailyLoads.Count -gt 0) { ($dailyLoads | Measure-Object -Average).Average } else { 0 }
  $sdLoad = 0
  if ($dailyLoads.Count -gt 1) {
    $variance = ($dailyLoads | ForEach-Object { [math]::Pow($_ - $meanLoad, 2) } | Measure-Object -Average).Average
    $sdLoad = [math]::Sqrt($variance)
  }
  $monotony = if ($sdLoad -gt 0) { [math]::Round($meanLoad / $sdLoad, 2) } else { 0 }
  $strain = $totalTss

  $deltaSleepText = if ($avgSleep -gt 0) { Format-DeltaValue -Value $sleepDelta -Unit "h" } else { "n/a" }
  $deltaHrvText = if ($avgHrv -gt 0) { Format-DeltaValue -Value $hrvDelta -Unit " ms" } else { "n/a" }
  $deltaRhrText = if ($avgRhr -gt 0) { Format-DeltaValue -Value $rhrDelta -Unit " bpm" } else { "n/a" }

  $deltaSleepClass = if ($sleepDelta -lt -0.5) { "neg" } elseif ($sleepDelta -gt 0.5) { "pos" } else { "neu" }
  $deltaHrvClass = if ($hrvDelta -lt -3) { "neg" } elseif ($hrvDelta -gt 3) { "pos" } else { "neu" }
  $deltaRhrClass = if ($rhrDelta -gt 3) { "neg" } elseif ($rhrDelta -lt -3) { "pos" } else { "neu" }

  $runDist = ($activities | Where-Object type -eq "Run" | Measure-Object distance_km -Sum).Sum
  $runTime = ($activities | Where-Object type -eq "Run" | Measure-Object moving_time_min -Sum).Sum
  $runPace = if ($runDist -gt 0) { $runTime / $runDist } else { 6.5 }

  if ($memoryText -match "Threshold Pace:\s*~?([0-9:]+)/km") {
    $p = $matches[1]
    $secs = Pace-To-Secs -Pace $p
    if ($secs) { $runPace = $secs / 60 * 1.15 }
  }

  $bikeDist = ($activities | Where-Object type -eq "Ride" | Measure-Object distance_km -Sum).Sum
  $bikeTime = ($activities | Where-Object type -eq "Ride" | Measure-Object moving_time_min -Sum).Sum
  $bikeSpeed = if ($bikeTime -gt 0) { $bikeDist / ($bikeTime / 60) } else { 28 }

  $swimDist = ($activities | Where-Object type -eq "Swim" | Measure-Object distance_km -Sum).Sum
  $swimTime = ($activities | Where-Object type -eq "Swim" | Measure-Object moving_time_min -Sum).Sum
  $swimPace100 = if ($swimDist -gt 0) { $swimTime / ($swimDist * 10) } else { 2.2 }

  $strengthTime = ($activities | Where-Object { $_.type -eq "Strength" -or $_.type -eq "WeightTraining" } | Measure-Object moving_time_min -Sum).Sum
  $totalTimeMin = ($activities | Measure-Object moving_time_min -Sum).Sum
  $swimPct = if ($totalTimeMin -gt 0) { [math]::Round(($swimTime / $totalTimeMin) * 100, 1) } else { 0 }
  $bikePct = if ($totalTimeMin -gt 0) { [math]::Round(($bikeTime / $totalTimeMin) * 100, 1) } else { 0 }
  $runPct = if ($totalTimeMin -gt 0) { [math]::Round(($runTime / $totalTimeMin) * 100, 1) } else { 0 }
  $strengthPct = if ($totalTimeMin -gt 0) { [math]::Round(($strengthTime / $totalTimeMin) * 100, 1) } else { 0 }

  $iconRun = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 12 6 12 9 6 13 18 16 12 21 12"/></svg>'
  $iconRide = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="5" cy="17" r="3"/><circle cx="19" cy="17" r="3"/><path d="M5 17l5-8h4l4 8M10 9h4"/></svg>'
  $iconSwim = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 17c2 1.5 4 1.5 6 0s4-1.5 6 0 4 1.5 6 0"/><path d="M3 21c2 1.5 4 1.5 6 0s4-1.5 6 0 4 1.5 6 0"/><circle cx="7" cy="7" r="2"/><path d="M9 9l4 2"/></svg>'
  $iconStrength = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 10v4M7 8v8M17 8v8M21 10v4"/><path d="M7 12h10"/></svg>'
  $iconInfo = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="10" x2="12" y2="16"/><line x1="12" y1="7" x2="12.01" y2="7"/></svg>'
  $iconHeart = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20.8 4.6a5.5 5.5 0 0 0-7.8 0L12 5.6l-1-1a5.5 5.5 0 1 0-7.8 7.8l1 1L12 21l7.8-7.6 1-1a5.5 5.5 0 0 0 0-7.8z"/></svg>'
  $iconWave = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12c2-4 4 4 6 0s4 4 6 0 4 4 6 0"/></svg>'
  $iconMoon = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3 7 7 0 0 0 21 12.8z"/></svg>'
  $iconScale = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M12 8v4"/><path d="M9 8h6"/></svg>'
  $iconBars = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="4" y1="20" x2="4" y2="10"/><line x1="10" y1="20" x2="10" y2="6"/><line x1="16" y1="20" x2="16" y2="14"/><line x1="22" y1="20" x2="22" y2="4"/></svg>'
  $iconBolt = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>'

  $raceCards = @()
  $nextRaceName = ""
  $nextRaceDays = $null
  $nextRacePriority = ""
  $nextRaceAName = ""
  $nextRaceADays = $null
  $otherRaces = @()
  $calendarEvents = Get-MemoryCalendar -Text $memoryText | Sort-Object date
  foreach ($ev in $calendarEvents) {
      $raceName = $ev.name
      $raceType = $ev.type
      $priority = if ($ev.priority) { ($ev.priority -replace "\*","") } else { "C" }
      $status = if ($ev.status -and $ev.status -ne "-") { $ev.status } else { "" }
      $raceDate = $ev.date
      if (-not $raceDate) { continue }
      $days = ($raceDate.Date - $referenceDate.Date).Days
      $days = [math]::Max(0, $days)
      $weeks = if ($days -gt 0) { [math]::Floor($days / 7) } else { 0 }
      $estimate = Estimate-Race -Type $raceType -Name $raceName -RunPace $runPace -BikeSpeed $bikeSpeed -SwimPace100 $swimPace100

      if ($days -ge 0) {
        if ($nextRaceDays -eq $null -or $days -lt $nextRaceDays) {
          $nextRaceDays = $days
          $nextRaceName = $raceName
          $nextRacePriority = $priority
        }
        if ($priority -eq "A") {
          if ($nextRaceADays -eq $null -or $days -lt $nextRaceADays) {
            $nextRaceADays = $days
            $nextRaceAName = $raceName
          }
        } else {
          $otherRaces += $raceName
        }
      }

      $stage = "Build"
      if ($priority -eq "A") { $stage = "Peak" }
      elseif ($priority -eq "B") { $stage = if ($raceType -match "Corrida") { "Especifico Run" } else { "Especifico" } }

      $raceCards += @"
<div class="race-card priority-$priority">
  <div class="race-days">
    <div class="days">$days</div>
    <div class="label">DIAS</div>
    <div class="sub">$weeks sem</div>
  </div>
  <div class="race-info">
    <div class="race-name">$raceName</div>
    <div class="race-meta">$raceType | Prioridade $priority</div>
    <div class="race-estimate">Estimativa hoje: $estimate</div>
    $(if ($status) { "<div class=""race-status"">$status</div>" } else { "" })
  </div>
  <div class="race-badge">$stage</div>
</div>
"@
  }
  if ($raceCards.Count -eq 0) {
    $raceCards = @("<div class=""card section"">Sem provas cadastradas.</div>")
  }

  $performanceCards = @"
<div class="performance-grid">
  <div class="stat-card">
    <div class="stat-value">$ctlText</div>
    <div class="stat-label">CTL (Fitness)</div>
    <div class="stat-pill">$ctlDeltaText</div>
    <div class="stat-note">Forma fisica acumulada nos ultimos 42 dias.</div>
  </div>
  <div class="stat-card">
    <div class="stat-value">$atlText</div>
    <div class="stat-label">ATL (Fadiga)</div>
    <div class="stat-pill">$atlDeltaText</div>
    <div class="stat-note">Carga de treino dos ultimos 7 dias.</div>
  </div>
  <div class="stat-card">
    <div class="stat-value">$tsbText</div>
    <div class="stat-label">TSB (Forma)</div>
    <div class="stat-pill $tsbClass">$tsbStatus</div>
    <div class="stat-note">CTL - ATL. Negativo = cansado. Ideal +5 a +15.</div>
  </div>
  <div class="stat-card">
    <div class="stat-value">$complianceText</div>
    <div class="stat-label">Aderência ao plano</div>
    <div class="stat-pill $complianceClass">$complianceLabel</div>
    <div class="stat-note">Quanto do plano foi cumprido (treinos + descanso). $complianceDetail</div>
  </div>
</div>
"@

  $wellnessCards = @"
<div class="wellness-grid">
  <div class="wellness-card">
    <div class="value">$avgRhr</div>
    <div class="label"><span class="well-icon">$iconHeart</span>FC repouso <a class="info-icon" href="#glossario-wellness" title="Batimentos em repouso; aumento pode indicar fadiga.">$iconInfo</a></div>
    <div class="delta $deltaRhrClass">$deltaRhrText vs base</div>
  </div>
  <div class="wellness-card">
    <div class="value">$avgHrv</div>
    <div class="label"><span class="well-icon">$iconWave</span>HRV <a class="info-icon" href="#glossario-wellness" title="Variabilidade da frequencia cardiaca; baixo sinaliza estresse.">$iconInfo</a></div>
    <div class="delta $deltaHrvClass">$deltaHrvText vs base</div>
  </div>
  <div class="wellness-card">
    <div class="value">$avgSleep h</div>
    <div class="label"><span class="well-icon">$iconMoon</span>Sono medio <a class="info-icon" href="#glossario-wellness" title="Horas dormidas; abaixo do ideal reduz recuperacao.">$iconInfo</a></div>
    <div class="delta $deltaSleepClass">$deltaSleepText vs ideal</div>
  </div>
  <div class="wellness-card">
    <div class="value">$pesoDisplay</div>
    <div class="label"><span class="well-icon">$iconScale</span>Peso <a class="info-icon" href="#glossario-wellness" title="Acompanhar tendencia semanal, nao dia isolado.">$iconInfo</a></div>
    <div class="delta neu">última medida</div>
  </div>
  <div class="wellness-card">
    <div class="value">$monotony</div>
    <div class="label"><span class="well-icon">$iconBars</span>Monotonia <a class="info-icon" href="#glossario-wellness" title="Variabilidade da carga diaria; alto = risco de estresse.">$iconInfo</a></div>
    <div class="delta neu">carga diária</div>
  </div>
  <div class="wellness-card">
    <div class="value">$strain</div>
    <div class="label"><span class="well-icon">$iconBolt</span>Strain <a class="info-icon" href="#glossario-wellness" title="Carga semanal total; alto com sono baixo = alerta.">$iconInfo</a></div>
    <div class="delta neu">carga semanal</div>
  </div>
</div>
"@

  $phaseObjectiveItems = ($phaseObjectives | ForEach-Object { "<li>$_</li>" }) -join ""
  $phaseTransitionItems = ($phaseTransition | ForEach-Object { "<li>$_</li>" }) -join ""

  $notesWeek = @()
  if ($report.PSObject.Properties.Name -contains "notas_semana") {
    $notesWeek = @($report.notas_semana)
  }
  $notesBlock = ""
  if ($IncludeWeekNotesInReport -and $notesWeek.Count -gt 0) {
    $noteLines = @()
    foreach ($n in $notesWeek) {
      $noteLines += "<div class=""note-item""><strong>$(Html-Escape $n.name)</strong><div>$(Html-Escape $n.description)</div></div>"
    }
    $notesBlock = "<section class=""card section""><h2>Notas da Semana</h2>$($noteLines -join '')</section>"
  }

  $totalTimeText = if ($totalTime -ne $null) { "{0:0.0}h" -f $totalTime } else { "n/a" }
  $tsbStateText = "controlada"
  $tsbAdvice = "Boa semana para manter a consistência."
  if ($tsb -le -25) { $tsbStateText = "pesada"; $tsbAdvice = "Vamos priorizar recuperação (sono e um pouco menos intensidade)." }
  elseif ($tsb -le -10) { $tsbStateText = "moderada"; $tsbAdvice = "Segurar a intensidade e priorizar sono." }
  elseif ($tsb -le 0) { $tsbStateText = "leve"; $tsbAdvice = "Só monitorar sinais de fadiga e manter o plano." }
  elseif ($tsb -gt 0) { $tsbStateText = "boa"; $tsbAdvice = "Boa janela para colocar qualidade com calma." }

  $adherenceLine = ""
  if ($plannedCount -gt 0) {
    $parts = @()
    if ($plannedWorkoutCount -gt 0) { $parts += "Treinos: $doneWorkoutCount/$plannedWorkoutCount ($workoutComplianceText)" }
    if ($plannedOffCount -gt 0) { $parts += "Descanso: $offRespectedCount/$plannedOffCount" }
    $parts += "Extras: $extraCount"
    $adherenceLine = "Aderência ao plano: $complianceText. " + ($parts -join " · ") + "."
  } else {
    $adherenceLine = "Aderência ao plano: n/a (sem planejamento importado nesta semana)."
  }

  $sleepLine = ""
  if ($avgSleep -gt 0 -and $idealSleep -gt 0 -and $avgSleep -lt ($idealSleep - 0.5)) {
    $sleepLine = "Seu sono médio foi $avgSleep h (ideal ~${idealSleep}h). Se der, tente ganhar +30-45 min em 2 noites."
  }

  $summaryText = "Boa semana de consistência: $activityCount sessões, $totalTimeText de treino e $totalTss TSS. A fadiga ficou $tsbStateText (TSB $tsbText). $tsbAdvice"

  $wins = @()
  if ($plannedWorkoutCount -gt 0) {
    $wins += "Você concluiu $doneWorkoutCount de $plannedWorkoutCount treinos planejados."
  }
  if ($plannedOffCount -gt 0 -and $offRespectedCount -gt 0) {
    $wins += "Descanso planejado respeitado ($offRespectedCount/$plannedOffCount)."
  }
  $longRun = $activities | Where-Object type -eq "Run" | Sort-Object moving_time_min -Descending | Select-Object -First 1
  if ($longRun -and $longRun.moving_time_min -ge 60) {
    $mins = [math]::Round($longRun.moving_time_min, 0)
    $dist = if ($longRun.distance_km -ne $null) { [math]::Round($longRun.distance_km, 1) } else { $null }
    $distText = if ($dist -ne $null) { " · $dist km" } else { "" }
    $wins += "Longão feito: $(Fix-TextEncoding $longRun.name) ($mins min$distText)."
  }
  $qualityRun = $activities | Where-Object { $_.type -eq "Run" -and (Fix-TextEncoding $_.name) -match "(?i)interval|tiro|tempo|limiar|threshold" } | Sort-Object moving_time_min -Descending | Select-Object -First 1
  if ($qualityRun) {
    $wins += "Sessão de qualidade na corrida concluída: $(Fix-TextEncoding $qualityRun.name)."
  }
  $strengthSessions = @($activities | Where-Object { $_.type -eq "Strength" -or $_.type -eq "WeightTraining" })
  if ($strengthSessions.Count -gt 0) {
    $wins += "Força presente ($($strengthSessions.Count)x) para sustentar bike/corrida e proteger o joelho."
  }
  if ($extraCount -gt 0) {
    $wins += "Você ainda fez $extraCount sessão(ões) extra. Boa energia, só vamos encaixar sem atrapalhar recuperação."
  }
  if ($sleepLine) { $wins += $sleepLine }

  $pending = @()
  foreach ($p in $missedWorkouts) {
    $pending += "$($p.start_date): $(Fix-TextEncoding $p.name)"
  }
  if ($plannedOffCount -gt 0 -and $offRespectedCount -lt $plannedOffCount) {
    $pending += "Descanso não respeitado em: $(@($adherence.off_broken_dates) -join ', ')"
  }
  if ($pending.Count -eq 0) {
    $pending += "Nada importante ficou pendente do plano nesta semana."
  }

  $focusLines = @()
  if ((@($missedWorkouts | Where-Object { $_.type -eq "Swim" })).Count -gt 0) {
    $focusLines += "Priorizar natação: pelo menos 1 sessão técnica na semana (mesmo curta)."
  }
  if ((@($missedWorkouts | Where-Object { $_.type -eq "Ride" })).Count -gt 0) {
    $focusLines += "Garantir 1 bike endurance (Z2) para sustentar o triathlon sprint."
  }
  if ($extraCount -gt 0) {
    $focusLines += "Quando bater vontade de fazer extra, tente substituir (não somar) ou me avise para ajustar a semana."
  }
  $focusLines += "Manter os 2 pilares da corrida: 1 sessão de qualidade + 1 longão (Z2), com progressão segura."
  if ($avgSleep -gt 0 -and $idealSleep -gt 0 -and $avgSleep -lt ($idealSleep - 0.5)) {
    $focusLines += "Sono: proteger recuperação (qualquer +30 min por noite já ajuda)."
  }

  $nextRaceLine = if ($nextRaceAName -and $nextRaceADays -ne $null) {
    "Faltam $nextRaceADays dias para $nextRaceAName (Prova A)."
  } elseif ($nextRaceName -and $nextRaceDays -ne $null) {
    "Faltam $nextRaceDays dias para $nextRaceName."
  } else {
    "Próximas provas monitoradas no calendário."
  }
  $otherRacesLine = if ($otherRaces.Count -gt 0) { "Outras provas no radar: " + ($otherRaces -join ", ") + "." } else { "" }

  $winsHtml = ($wins | ForEach-Object { "<li>$(Html-Escape $_)</li>" }) -join ""
  $pendingHtml = ($pending | ForEach-Object { "<li>$(Html-Escape $_)</li>" }) -join ""
  $focusHtml = ($focusLines | ForEach-Object { "<li>$(Html-Escape $_)</li>" }) -join ""

  $feedbackBlock = @"
<section class="card section coach-card">
  <div class="coach-header">
    <div class="coach-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 11a8 8 0 0 1-8 8H7l-4 3 1-5a8 8 0 1 1 17-6z"/></svg></div>
    <div>
      <h2>Feedback do Coach</h2>
      <div class="coach-subtitle">O que foi bem, o que ajustar e o foco da próxima semana</div>
    </div>
  </div>
  <div class="coach-block">
    <div class="coach-title">Resumo rápido</div>
    <p>$(Html-Escape $summaryText)</p>
    <p class="muted" style="margin:8px 0 0 0">$(Html-Escape $adherenceLine)</p>
  </div>
  <div class="coach-block">
    <div class="coach-title">O que você fez bem</div>
    <ul>$winsHtml</ul>
  </div>
  <div class="coach-block">
    <div class="coach-title">O que ficou pendente</div>
    <ul>$pendingHtml</ul>
  </div>
  <div class="coach-block">
    <div class="coach-title">Foco para a próxima semana</div>
    <p>$(Html-Escape $nextRaceLine) $(Html-Escape $otherRacesLine)</p>
    <ol>$focusHtml</ol>
  </div>
</section>
"@

  $recommendationsBlock = @"
<section class="card section">
  <h2>Recomendacoes Tecnicas</h2>
  <div class="rec-grid">
    <div class="rec-card">
      <div class="rec-icon strength">$iconStrength</div>
      <div class="rec-title">Forca</div>
      <ul>
        <li>Gluteo medio/maximo (estabilidade no pedal e corrida)</li>
        <li>Core anti-rotacao (natacao e bike aero)</li>
        <li>Panturrilha excêntrica (protecao tendao Aquiles)</li>
        <li>Mobilidade de quadril (posicao aero)</li>
      </ul>
    </div>
    <div class="rec-card">
      <div class="rec-icon ride">$iconRide</div>
      <div class="rec-title">Ciclismo</div>
      <ul>
        <li>Cadencia: manter 85-95 rpm em Z2</li>
        <li>Pedalada redonda: foco no &quot;puxar&quot; (11h-2h)</li>
        <li>Posicao aero: aumentar tempo gradualmente</li>
        <li>Sweet Spot: progressao para Build</li>
      </ul>
    </div>
    <div class="rec-card">
      <div class="rec-icon run">$iconRun</div>
      <div class="rec-title">Corrida</div>
      <ul>
        <li>Cadencia alta (175-180 spm) para reduzir impacto</li>
        <li>Aterrissagem medio-pe</li>
        <li>Volume protegido ate joelho &lt; 3/10</li>
        <li>Fortalecimento: step-ups, agachamento unilateral</li>
      </ul>
    </div>
    <div class="rec-card">
      <div class="rec-icon swim">$iconSwim</div>
      <div class="rec-title">Natacao</div>
      <ul>
        <li>Rotacao de quadril (nao so ombros)</li>
        <li>Cotovelo alto na puxada</li>
        <li>Respiracao bilateral</li>
        <li>Aumentar frequencia: 3x/semana minimo</li>
      </ul>
    </div>
  </div>
</section>
"@

  $wellnessGlossary = @"
<section class="card section" id="glossario-wellness">
  <h2>Glossario Wellness</h2>
  <a class="back-link" href="#wellness">Voltar ao Wellness</a>
  <div class="glossary-grid">
    <div><strong>FC repouso:</strong> batimentos em repouso; aumento pode indicar fadiga.</div>
    <div><strong>HRV:</strong> variabilidade da frequencia cardiaca; baixo sinaliza estresse.</div>
    <div><strong>Sono medio:</strong> horas dormidas; abaixo do ideal reduz recuperacao.</div>
    <div><strong>Peso:</strong> acompanhamento semanal; use tendencia, nao dia isolado.</div>
    <div><strong>Monotonia:</strong> variabilidade da carga diaria; alto = risco de estresse.</div>
    <div><strong>Strain:</strong> carga semanal; alto com sono baixo = alerta.</div>
  </div>
</section>
"@

  $longterm = $null
  $longtermFile = Get-ChildItem $ReportsDir -Filter "intervals_longterm_*coach_edition.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($longtermFile) {
    try {
      $longterm = Get-Content -Raw $longtermFile.FullName | ConvertFrom-Json
    } catch {}
  }

  $trendReports = @()
  if ($longterm -and $longterm.analise_semanal) {
    foreach ($w in $longterm.analise_semanal) {
      $endDate = [DateTime]::Parse($w.fim)
      if ($endDate -le $referenceDate) {
        $trendReports += [PSCustomObject]@{
          end = $endDate
          label = $w.semana
          ctl = $w.ctl
          atl = $w.atl
          tsb = $w.tsb
          ramp = $w.rampRate
          tss = $w.total_tss
          hours = $w.tempo_horas
          weight = $w.peso_medio
          sleep = $w.sono_medio
          hrv = $w.hrv_medio
          rhr = $w.fc_repouso_media
          bike_pct = $w.bike_pct
          run_pct = $w.run_pct
          swim_pct = $w.swim_pct
          strength_pct = $w.strength_pct
          bike_if = $w.bike_analise.avg_if
          bike_vi = $w.bike_analise.avg_vi
          bike_dec = $w.bike_analise.avg_decoupling
          run_cad = $w.run_analise.avg_cadence
          run_dec = $w.run_analise.avg_decoupling
          swim_swolf = $w.swim_analise.avg_swolf
          bike_z1 = if ($w.bike_zones) { $w.bike_zones.z1_pct } else { $null }
          bike_z2 = if ($w.bike_zones) { $w.bike_zones.z2_pct } else { $null }
          bike_z3 = if ($w.bike_zones) { $w.bike_zones.z3_pct } else { $null }
          run_z1 = if ($w.run_zones) { $w.run_zones.z1_pct } else { $null }
          run_z2 = if ($w.run_zones) { $w.run_zones.z2_pct } else { $null }
          run_z3 = if ($w.run_zones) { $w.run_zones.z3_pct } else { $null }
        }
      }
    }
  } else {
    foreach ($file in $reportFiles) {
      try {
        $r = Get-Content -Raw $file.FullName | ConvertFrom-Json
        $endDate = [DateTime]::Parse($r.semana.fim)
        if ($endDate -le $referenceDate) {
          $sleepAvg = if ($r.bem_estar) { [math]::Round((($r.bem_estar | Measure-Object sono_h -Average).Average), 2) } else { $null }
          $hrvAvg = if ($r.bem_estar) { [math]::Round((($r.bem_estar | Measure-Object hrv -Average).Average), 1) } else { $null }
          $rhrAvg = if ($r.bem_estar) { [math]::Round((($r.bem_estar | Measure-Object fc_reposo -Average).Average), 1) } else { $null }
          $trendReports += [PSCustomObject]@{
            end = $endDate
            label = $endDate.ToString("dd/MM")
            ctl = $r.metricas.CTL
            atl = $r.metricas.ATL
            tsb = $r.metricas.TSB
            ramp = $r.metricas.RampRate
            tss = $r.semana.carga_total_tss
            hours = $r.semana.tempo_total_horas
            weight = $r.metricas.peso_atual
            sleep = $sleepAvg
            hrv = $hrvAvg
            rhr = $rhrAvg
          }
        }
      } catch {}
    }
  }
  $trendReports = @($trendReports | Sort-Object end)
  if ($trendReports.Count -gt 12) { $trendReports = $trendReports[-12..-1] }

  $trendLabels = ConvertTo-Json ($trendReports | ForEach-Object { $_.label }) -Compress
  $trendCtl = ConvertTo-Json ($trendReports | ForEach-Object { $_.ctl }) -Compress
  $trendAtl = ConvertTo-Json ($trendReports | ForEach-Object { $_.atl }) -Compress
  $trendTsb = ConvertTo-Json ($trendReports | ForEach-Object { $_.tsb }) -Compress
  $trendTss = ConvertTo-Json ($trendReports | ForEach-Object { $_.tss }) -Compress
  $trendHours = ConvertTo-Json ($trendReports | ForEach-Object { $_.hours }) -Compress
  $trendSleep = ConvertTo-Json ($trendReports | ForEach-Object { $_.sleep }) -Compress
  $trendHrv = ConvertTo-Json ($trendReports | ForEach-Object { $_.hrv }) -Compress
  $trendRhr = ConvertTo-Json ($trendReports | ForEach-Object { $_.rhr }) -Compress
  $trendWeight = ConvertTo-Json ($trendReports | ForEach-Object { $_.weight }) -Compress
  $trendBikePct = ConvertTo-Json ($trendReports | ForEach-Object { $_.bike_pct }) -Compress
  $trendRunPct = ConvertTo-Json ($trendReports | ForEach-Object { $_.run_pct }) -Compress
  $trendSwimPct = ConvertTo-Json ($trendReports | ForEach-Object { $_.swim_pct }) -Compress
  $trendStrengthPct = ConvertTo-Json ($trendReports | ForEach-Object { $_.strength_pct }) -Compress
  $trendBikeIf = ConvertTo-Json ($trendReports | ForEach-Object { $_.bike_if }) -Compress
  $trendBikeVi = ConvertTo-Json ($trendReports | ForEach-Object { $_.bike_vi }) -Compress
  $trendBikeDec = ConvertTo-Json ($trendReports | ForEach-Object { $_.bike_dec }) -Compress
  $trendRunCad = ConvertTo-Json ($trendReports | ForEach-Object { $_.run_cad }) -Compress
  $trendRunDec = ConvertTo-Json ($trendReports | ForEach-Object { $_.run_dec }) -Compress
  $trendSwimSwolf = ConvertTo-Json ($trendReports | ForEach-Object { $_.swim_swolf }) -Compress
  $trendBikeZ1 = ConvertTo-Json ($trendReports | ForEach-Object { $_.bike_z1 }) -Compress
  $trendBikeZ2 = ConvertTo-Json ($trendReports | ForEach-Object { $_.bike_z2 }) -Compress
  $trendBikeZ3 = ConvertTo-Json ($trendReports | ForEach-Object { $_.bike_z3 }) -Compress
  $trendRunZ1 = ConvertTo-Json ($trendReports | ForEach-Object { $_.run_z1 }) -Compress
  $trendRunZ2 = ConvertTo-Json ($trendReports | ForEach-Object { $_.run_z2 }) -Compress
  $trendRunZ3 = ConvertTo-Json ($trendReports | ForEach-Object { $_.run_z3 }) -Compress

  $trendWeeks = $trendReports.Count
  $trendWeeksRaw = if ($longterm -and $longterm.analise_semanal) { $longterm.analise_semanal.Count } else { $trendWeeks }
  $trendWeeksLabel = if ($trendWeeks -eq 1) { "1 semana" } else { "$trendWeeks semanas" }
  if ($trendWeeksRaw -gt $trendWeeks) { $trendWeeksLabel = "$trendWeeks semanas completas (de $trendWeeksRaw)" }
  $trendTitle = if ($longterm) { "Tendencias Longo Prazo ($trendWeeksLabel)" } else { "Tendencias ($trendWeeksLabel)" }
  $hasBikeMetrics = ($trendReports | Where-Object { $_.bike_if -gt 0 -or $_.bike_vi -gt 0 -or $_.bike_dec -ne $null }).Count -gt 0
  $hasRunMetrics = ($trendReports | Where-Object { $_.run_cad -gt 0 -or $_.run_dec -ne $null }).Count -gt 0
  $hasSwimMetrics = ($trendReports | Where-Object { $_.swim_swolf -gt 0 }).Count -gt 0
  $hasBikeZones = ($trendReports | Where-Object { $_.bike_z1 -gt 0 -or $_.bike_z2 -gt 0 -or $_.bike_z3 -gt 0 }).Count -gt 0
  $hasRunZones = ($trendReports | Where-Object { $_.run_z1 -gt 0 -or $_.run_z2 -gt 0 -or $_.run_z3 -gt 0 }).Count -gt 0

  $trendChartCards = @(
    '      <div class="card"><canvas id="trend-load-chart"></canvas></div>',
    '      <div class="card"><canvas id="trend-tss-chart"></canvas></div>',
    '      <div class="card"><canvas id="trend-well-chart"></canvas></div>',
    '      <div class="card"><canvas id="trend-modality-chart"></canvas></div>'
  )
  if ($hasBikeMetrics) { $trendChartCards += '      <div class="card"><canvas id="trend-bike-chart"></canvas></div>' }
  if ($hasRunMetrics) { $trendChartCards += '      <div class="card"><canvas id="trend-run-chart"></canvas></div>' }
  if ($hasSwimMetrics) { $trendChartCards += '      <div class="card"><canvas id="trend-swim-chart"></canvas></div>' }
  if ($hasBikeZones) { $trendChartCards += '      <div class="card"><canvas id="trend-zone-bike"></canvas></div>' }
  if ($hasRunZones) { $trendChartCards += '      <div class="card"><canvas id="trend-zone-run"></canvas></div>' }

  $trendChartScripts = @(
    "  new Chart(document.getElementById('trend-load-chart'),{type:'line',data:{labels:$trendLabels,datasets:[{label:'CTL',data:$trendCtl,borderColor:'#60a5fa',tension:.3},{label:'ATL',data:$trendAtl,borderColor:'#f59e0b',tension:.3},{label:'TSB',data:$trendTsb,borderColor:'#22c55e',tension:.3}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false}}}});",
    "  new Chart(document.getElementById('trend-tss-chart'),{data:{labels:$trendLabels,datasets:[{type:'bar',label:'TSS',data:$trendTss,backgroundColor:'rgba(94,163,255,0.4)'},{type:'line',label:'Horas',data:$trendHours,borderColor:'#f59e0b',tension:.3}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false}}}});",
    "  new Chart(document.getElementById('trend-well-chart'),{type:'line',data:{labels:$trendLabels,datasets:[{label:'Sono',data:$trendSleep,borderColor:'#22c55e',tension:.3},{label:'HRV',data:$trendHrv,borderColor:'#60a5fa',tension:.3},{label:'FC Repouso',data:$trendRhr,borderColor:'#ef4444',tension:.3},{label:'Peso',data:$trendWeight,borderColor:'#a855f7',tension:.3}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false}}}});",
    "  new Chart(document.getElementById('trend-modality-chart'),{type:'bar',data:{labels:$trendLabels,datasets:[{label:'Bike %',data:$trendBikePct,backgroundColor:'rgba(245,158,11,0.35)'},{label:'Run %',data:$trendRunPct,backgroundColor:'rgba(34,197,94,0.35)'},{label:'Swim %',data:$trendSwimPct,backgroundColor:'rgba(56,189,248,0.35)'},{label:'Forca %',data:$trendStrengthPct,backgroundColor:'rgba(168,85,247,0.35)'}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false},y:{beginAtZero:true,max:100}}}});"
  )
  if ($hasBikeMetrics) { $trendChartScripts += "  new Chart(document.getElementById('trend-bike-chart'),{type:'line',data:{labels:$trendLabels,datasets:[{label:'Bike IF',data:$trendBikeIf,borderColor:'#f59e0b',tension:.3},{label:'Bike VI',data:$trendBikeVi,borderColor:'#94a3b8',tension:.3},{label:'Bike Decoupling %',data:$trendBikeDec,borderColor:'#ef4444',tension:.3}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false}}}});" }
  if ($hasRunMetrics) { $trendChartScripts += "  new Chart(document.getElementById('trend-run-chart'),{type:'line',data:{labels:$trendLabels,datasets:[{label:'Cadencia Run',data:$trendRunCad,borderColor:'#22c55e',tension:.3},{label:'Decoupling Run %',data:$trendRunDec,borderColor:'#ef4444',tension:.3}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false}}}});" }
  if ($hasSwimMetrics) { $trendChartScripts += "  new Chart(document.getElementById('trend-swim-chart'),{type:'line',data:{labels:$trendLabels,datasets:[{label:'Swolf',data:$trendSwimSwolf,borderColor:'#38bdf8',tension:.3}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false}}}});" }
  if ($hasBikeZones) { $trendChartScripts += "  new Chart(document.getElementById('trend-zone-bike'),{type:'bar',data:{labels:$trendLabels,datasets:[{label:'Z1',data:$trendBikeZ1,backgroundColor:'rgba(34,197,94,0.35)'},{label:'Z2',data:$trendBikeZ2,backgroundColor:'rgba(245,158,11,0.35)'},{label:'Z3',data:$trendBikeZ3,backgroundColor:'rgba(239,68,68,0.35)'}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false},y:{beginAtZero:true,max:100}}}});" }
  if ($hasRunZones) { $trendChartScripts += "  new Chart(document.getElementById('trend-zone-run'),{type:'bar',data:{labels:$trendLabels,datasets:[{label:'Z1',data:$trendRunZ1,backgroundColor:'rgba(34,197,94,0.35)'},{label:'Z2',data:$trendRunZ2,backgroundColor:'rgba(245,158,11,0.35)'},{label:'Z3',data:$trendRunZ3,backgroundColor:'rgba(239,68,68,0.35)'}]},options:{plugins:{legend:{position:'bottom'}},scales:{x:{display:false},y:{beginAtZero:true,max:100}}}});" }

  $trendCtlDelta = if ($trendWeeks -gt 1) { [math]::Round(($trendReports[-1].ctl - $trendReports[0].ctl), 1) } else { 0 }
  $trendWeightDelta = if ($trendWeeks -gt 1 -and $trendReports[-1].weight -ne $null -and $trendReports[0].weight -ne $null) { [math]::Round(($trendReports[-1].weight - $trendReports[0].weight), 1) } else { $null }

  $last4 = if ($trendWeeks -gt 0) { $trendReports | Select-Object -Last ([Math]::Min(4, $trendWeeks)) } else { @() }
  $prev4 = if ($trendWeeks -gt 4) { $trendReports | Select-Object -Skip ([Math]::Max(0, $trendWeeks - 8)) -First ([Math]::Min(4, $trendWeeks - 4)) } else { @() }
  $last4Tss = if ($last4.Count -gt 0) { [math]::Round((($last4 | Measure-Object tss -Average).Average), 0) } else { $null }
  $prev4Tss = if ($prev4.Count -gt 0) { [math]::Round((($prev4 | Measure-Object tss -Average).Average), 0) } else { $null }
  $tssDelta = if ($last4Tss -ne $null -and $prev4Tss -ne $null) { [math]::Round(($last4Tss - $prev4Tss), 0) } else { $null }
  $last4Sleep = if ($last4.Count -gt 0) { [math]::Round((($last4 | Measure-Object sleep -Average).Average), 2) } else { $null }
  $prev4Sleep = if ($prev4.Count -gt 0) { [math]::Round((($prev4 | Measure-Object sleep -Average).Average), 2) } else { $null }
  $sleepTrend = if ($last4Sleep -ne $null -and $prev4Sleep -ne $null) { [math]::Round(($last4Sleep - $prev4Sleep), 2) } else { $null }

  $perfScore = Get-PerformanceScore -CTL $ctl -TSB $tsb -RampRate $ramp
  $recovery = Get-RecoveryStatus -TSB $tsb -HRV $avgHrv -RestingHR $avgRhr -Sleep $avgSleep
  $recoveryStatus = if ($recovery) { $recovery.status } else { "n/a" }
  $recoveryScore = if ($recovery) { $recovery.score } else { "n/a" }

  $longtermInsightsBlock = ""
  $predictionBlock = ""
  if ($longterm) {
    $insights = @($longterm.insights)
    $alerts = @($longterm.alertas)
    $insightItems = if ($insights.Count -gt 0) { ($insights | ForEach-Object { "<li>$_</li>" }) -join "" } else { "" }
    $alertItems = if ($alerts.Count -gt 0) { ($alerts | ForEach-Object { "<li>$_</li>" }) -join "" } else { "" }
    $blocks = @($longterm.analise_blocos)
    $blockCards = @()
    if ($blocks.Count -gt 0) {
      $lastBlocks = $blocks | Select-Object -Last ([Math]::Min(3, $blocks.Count))
      foreach ($b in $lastBlocks) {
        $blockCards += "<div class=""lt-block""><strong>$($b.bloco)</strong><div>$($b.semanas)</div><div>TSS: $($b.media_tss_semanal)</div><div>CTL: $($b.ctl_final) | TSB: $($b.tsb_final)</div></div>"
      }
    }
    $blocksHtml = if ($blockCards.Count -gt 0) { ($blockCards -join "") } else { "<div class=""lt-empty"">Sem blocos suficientes.</div>" }
    $periodo = if ($longterm.relatorio) { "$($longterm.relatorio.inicio) a $($longterm.relatorio.fim)" } else { "" }
    if ($longterm.predicao_evento) {
      $p = $longterm.predicao_evento
      $predictionBlock = @"
<div class="trend-score">
  <div class="score-card">
    <div class="score-title">Predicao Prova A</div>
    <div class="score-value">$($p.status_preparacao)</div>
    <div class="score-note">Dias: $($p.dias_restantes) | CTL atual: $($p.ctl_atual) | CTL proj: $($p.ctl_projetado)</div>
  </div>
  <div class="score-card">
    <div class="score-title">Taper sugerido</div>
    <div class="score-value">$($p.tsb_ideal_prova)</div>
    <div class="score-note">$($p.recomendacao_taper)</div>
  </div>
</div>
"@
    }
    $longtermInsightsBlock = @"
<div class="lt-meta">Longo prazo: $periodo</div>
<div class="lt-grid">
  <div class="lt-panel">
    <h3>Insights</h3>
    $(if ($insightItems) { "<ul class=""lt-list"">$insightItems</ul>" } else { "<div class=""lt-empty"">Sem insights.</div>" })
  </div>
  <div class="lt-panel">
    <h3>Alertas</h3>
    $(if ($alertItems) { "<ul class=""lt-list lt-alert"">$alertItems</ul>" } else { "<div class=""lt-empty"">Sem alertas.</div>" })
  </div>
  <div class="lt-panel">
    <h3>Blocos 4 semanas</h3>
    <div class="lt-blocks">$blocksHtml</div>
  </div>
</div>
"@
  }

  $trendCards = @"
<div class="performance-grid">
  <div class="stat-card">
    <div class="stat-value">$ctlText</div>
    <div class="stat-label">CTL atual</div>
    <div class="stat-pill">$(Format-DeltaValue -Value $trendCtlDelta -Unit "")</div>
    <div class="stat-note">Evolucao nas ultimas $trendWeeks semanas.</div>
  </div>
  <div class="stat-card">
    <div class="stat-value">$last4Tss</div>
    <div class="stat-label">TSS medio 4s</div>
    <div class="stat-pill">$(Format-DeltaValue -Value $tssDelta -Unit "")</div>
    <div class="stat-note">Comparado com as 4 semanas anteriores.</div>
  </div>
  <div class="stat-card">
    <div class="stat-value">$last4Sleep h</div>
    <div class="stat-label">Sono medio 4s</div>
    <div class="stat-pill">$(Format-DeltaValue -Value $sleepTrend -Unit "h")</div>
    <div class="stat-note">Tendencia recente de recuperacao.</div>
  </div>
  <div class="stat-card">
    <div class="stat-value">$(if ($trendReports.Count -gt 0) { $trendReports[-1].weight } else { "n/a" })</div>
    <div class="stat-label">Peso atual</div>
    <div class="stat-pill">$(if ($trendWeightDelta -ne $null) { Format-DeltaValue -Value $trendWeightDelta -Unit "kg" } else { "n/a" })</div>
    <div class="stat-note">Variacao desde o inicio do periodo.</div>
  </div>
</div>
<div class="trend-score">
  <div class="score-card">
    <div class="score-title">Performance Score</div>
    <div class="score-value">$perfScore</div>
    <div class="score-note">CTL + TSB + RampRate</div>
  </div>
  <div class="score-card">
    <div class="score-title">Recovery Status</div>
    <div class="score-value">$recoveryStatus</div>
    <div class="score-note">Sono, HRV, FC repouso e TSB</div>
  </div>
</div>
$predictionBlock
$longtermInsightsBlock
"@

  $activityCards = @()
  foreach ($a in ($activities | Sort-Object start_date_local)) {
    $plan = $a.planejado
    $planTarget = Parse-PlanTarget -Description $plan.description -Type $a.type

    $actualPaceMin = if ($a.type -eq "Run" -and $a.distance_km -gt 0) { $a.moving_time_min / $a.distance_km } else { $null }
    $actualPace = if ($actualPaceMin) { Format-Pace -MinutesPerKm $actualPaceMin } else { $null }

    $wellDay = $wellness | Where-Object { $_.data -eq $a.start_date_local } | Select-Object -First 1
    $wellFlags = Get-WellnessFlags -WellDay $wellDay -BaselineRhr $baselineRhr -BaselineHrv $baselineHrv -IdealSleep $idealSleep
    if ($wellDay) {
      $sleepDeltaDay = if ($wellDay.sono_h -ne $null) { $wellDay.sono_h - $idealSleep } else { $null }
      $hrvDeltaDay = if ($wellDay.hrv -ne $null) { $wellDay.hrv - $baselineHrv } else { $null }
      $rhrDeltaDay = if ($wellDay.fc_reposo -ne $null) { $wellDay.fc_reposo - $baselineRhr } else { $null }
      $sleepValue = if ($wellDay.sono_h -ne $null) { "$($wellDay.sono_h)h" } else { "n/a" }
      $hrvValue = if ($wellDay.hrv -ne $null) { $wellDay.hrv } else { "n/a" }
      $rhrValue = if ($wellDay.fc_reposo -ne $null) { $wellDay.fc_reposo } else { "n/a" }
      $wellText = "Sono $sleepValue ($(Format-DeltaValue -Value $sleepDeltaDay -Unit 'h')) | HRV $hrvValue ($(Format-DeltaValue -Value $hrvDeltaDay -Unit ' ms')) | FC $rhrValue ($(Format-DeltaValue -Value $rhrDeltaDay -Unit ' bpm'))"
      if ($wellFlags.Count -gt 0) { $wellText += " | Alerta: " + ($wellFlags -join ", ") }
    } else {
      $wellText = "Sem wellness do dia"
    }

    $comparison = Compare-Target -PlanTarget $planTarget -Type $a.type -AvgWatts $a.average_watts -ActualPace $actualPace
    if ($comparison -match "Abaixo" -and $wellDay -and $wellDay.sono_h -lt ($idealSleep - 0.5)) { $comparison += " (sono baixo pode ter influenciado)" }

    $planSummary = "Sem planejado encontrado."
    if ($plan) {
      $planParts = @()
      if ($plan.moving_time_min -ne $null) { $planParts += "$($plan.moving_time_min)min" }
      if ($plan.distance_km -ne $null) {
        if ($a.type -eq "Swim") { $planParts += ("{0}m" -f [math]::Round(($plan.distance_km * 1000), 0)) }
        else { $planParts += "$($plan.distance_km)km" }
      }
      if ($planTarget) {
        if ($planTarget.unit -eq "W") { $planParts += "$($planTarget.min)-$($planTarget.max)W" }
        elseif ($planTarget.unit -eq "pace") { $planParts += "$($planTarget.min)-$($planTarget.max)/km" }
      }
      if ($planParts.Count -eq 0) { $planParts = @("sem detalhes") }

      $actualParts = @()
      if ($a.moving_time_min -ne $null) { $actualParts += "{0}min" -f ([math]::Round($a.moving_time_min, 0)) }
      if ($a.distance_km -ne $null) {
        if ($a.type -eq "Swim") { $actualParts += ("{0}m" -f [math]::Round(($a.distance_km * 1000), 0)) }
        else { $actualParts += "{0}km" -f ([math]::Round($a.distance_km, 1)) }
      }
      if ($a.average_watts -ne $null -and $a.type -ne "Swim") { $actualParts += "{0}W" -f ([math]::Round($a.average_watts, 0)) }
      if ($actualPace -and $a.type -eq "Run") { $actualParts += $actualPace }
      if ($actualParts.Count -eq 0) { $actualParts = @("sem dados") }

      $deltaParts = @()
      if ($plan.delta_time_min -ne $null) { $deltaParts += "Delta tempo $(Format-DeltaValue -Value $plan.delta_time_min -Unit 'min')" }
      if ($plan.delta_distance_km -ne $null) {
        if ($a.type -eq "Swim") { $deltaParts += "Delta dist $(Format-DeltaValue -Value ($plan.delta_distance_km * 1000) -Unit 'm')" }
        else { $deltaParts += "Delta dist $(Format-DeltaValue -Value $plan.delta_distance_km -Unit 'km')" }
      }
      if ($comparison) { $deltaParts += "Intensidade: $comparison" }
      if ($deltaParts.Count -eq 0) { $deltaParts = @("sem comparação") }

      $planSummary = "Planejado: " + ($planParts -join " • ") + " | Executado: " + ($actualParts -join " • ") + " | " + ($deltaParts -join " • ")
    }

    $objective = Get-Objective -Plan $plan -Activity $a

    $durationText = "{0}min" -f ([math]::Round($a.moving_time_min,0))
    if ($a.type -eq "Swim" -and $a.distance_km -gt 0) {
      $distanceText = "{0}m" -f ([math]::Round($a.distance_km * 1000, 0))
    } else {
      $distanceText = "{0}km" -f ([math]::Round($a.distance_km,1))
    }
    $tssValue = $null
    if ($a.suffer_score -ne $null) { $tssValue = $a.suffer_score }
    elseif ($a.training_load -ne $null) { $tssValue = $a.training_load }
    elseif ($a.power_load -ne $null) { $tssValue = $a.power_load }
    elseif ($a.hr_load -ne $null) { $tssValue = $a.hr_load }
    elseif ($a.pace_load -ne $null) { $tssValue = $a.pace_load }
    elseif ($a.strain_score -ne $null) { $tssValue = $a.strain_score }
    $tssText = if ($tssValue -ne $null) { [math]::Round($tssValue,0) } else { "n/a" }
    $powerText = if ($a.average_watts) { [math]::Round($a.average_watts,0).ToString() + "W" } else { "-" }
    $npText = if ($a.normalized_power) { [math]::Round($a.normalized_power,0).ToString() + "W" } else { "-" }
    $hrText = if ($a.average_hr) { [math]::Round($a.average_hr,0).ToString() + " bpm" } else { "-" }
    $viText = if ($a.variabilidade) { [math]::Round($a.variabilidade,2).ToString() } elseif ($a.normalized_power -and $a.average_watts) { [math]::Round(($a.normalized_power / $a.average_watts),2).ToString() } else { "-" }

    $intensityPower = if ($a.normalized_power) { $a.normalized_power } else { $a.average_watts }
    $ifValue = $null
    if ($a.type -eq "Ride") { $ifValue = Safe-Divide -Num $intensityPower -Den $ftpBike }
    elseif ($a.type -eq "Run") { $ifValue = Safe-Divide -Num $intensityPower -Den $ftpRun }
    $ifText = if ($ifValue) { [math]::Round($ifValue, 2) } else { "-" }

    $paceLabel = "Pace"
    $paceValue = "-"
    if ($a.type -eq "Run") {
      $paceValue = if ($actualPace) { $actualPace } else { "-" }
    } elseif ($a.type -eq "Swim" -and $a.distance_km -gt 0) {
      $paceValue = Format-Pace100 -MinutesPer100 ($a.moving_time_min / ($a.distance_km * 10))
      $paceLabel = "Pace 100m"
    } elseif ($a.type -eq "Ride" -and $a.moving_time_min -gt 0) {
      $paceLabel = "Velocidade"
      $paceValue = "{0} km/h" -f ([math]::Round(($a.distance_km / ($a.moving_time_min / 60)),1))
    }

    $paceDeltaText = $null
    if ($a.type -eq "Run" -and $actualPaceMin) {
      $paceDeltaText = Format-DeltaPace -DeltaSeconds (($actualPaceMin * 60) - $thresholdPaceSecs)
    }

    $qualityLabel = "Dentro"
    $qualityClass = "quality-ok"
    if ($comparison -match "Sem alvo" -or $comparison -match "Sem dado") { $qualityLabel = "Sem alvo"; $qualityClass = "quality-low" }
    elseif ($comparison -match "Acima") { $qualityLabel = "Acima"; $qualityClass = "quality-high" }
    elseif ($comparison -match "Abaixo" -or $comparison -match "Mais lento") { $qualityLabel = "Abaixo"; $qualityClass = "quality-low" }
    if ($wellFlags.Count -gt 0 -and $qualityLabel -eq "Acima") { $qualityLabel = "Risco"; $qualityClass = "quality-risk" }

    $typeClass = ($a.type).ToLower()
    $typeIcon = $iconRun
    switch ($a.type) {
      "Run" { $typeClass = "run"; $typeIcon = $iconRun }
      "Ride" { $typeClass = "ride"; $typeIcon = $iconRide }
      "Swim" { $typeClass = "swim"; $typeIcon = $iconSwim }
      "Strength" { $typeClass = "strength"; $typeIcon = $iconStrength }
      "WeightTraining" { $typeClass = "strength"; $typeIcon = $iconStrength }
      default { $typeClass = "run"; $typeIcon = $iconRun }
    }

    $recommendation = "Manter consistência."
    if ($comparison -match "Acima" -and $wellFlags.Count -gt 0) { $recommendation = "Segurar intensidade e priorizar recuperação hoje." }
    elseif ($comparison -match "Acima") { $recommendation = "Sessão acima do alvo; monitorar fadiga." }
    elseif ($comparison -match "Abaixo" -and $wellFlags.Count -gt 0) { $recommendation = "Redução de carga ok dado wellness baixo; foco em sono." }
    elseif ($comparison -match "Abaixo") { $recommendation = "Se era sessão chave, revisar pacing/potência." }
    if ($a.type -eq "Run" -and $a.notas -match "joelho") { $recommendation = "Manter volume protegido até joelho < 4/10." }
    if ($a.average_hr -and $lthr -gt 0 -and $a.average_hr -gt ($lthr + 5)) { $recommendation += " FC média alta para o alvo." }

    $summaryBase = switch ($a.type) {
      "Run" { "Corrida" }
      "Ride" { "Ciclismo" }
      "Swim" { "Natacao" }
      "Strength" { "Forca" }
      "WeightTraining" { "Forca" }
      default { "Sessao" }
    }
    $targetText = if ($comparison -match "Dentro") { "dentro do alvo" } elseif ($comparison -match "Acima") { "acima do alvo" } elseif ($comparison -match "Abaixo" -or $comparison -match "Mais lento") { "abaixo do alvo" } else { "sem alvo definido" }
    $sentence1 = "$summaryBase $targetText."
    $sentence2Parts = @()
    if ($plan -and $plan.moving_time_min -ne $null -and $plan.moving_time_min -gt 0 -and $plan.delta_time_min -ne $null) {
      $pct = [math]::Abs(($plan.delta_time_min / $plan.moving_time_min) * 100)
      if ($pct -le 10) { $sentence2Parts += "Volume alinhado ao planejado." }
      elseif ($plan.delta_time_min -gt 0) { $sentence2Parts += "Volume acima do planejado." }
      else { $sentence2Parts += "Volume abaixo do planejado." }
    }
    if ($a.average_hr -and $lthr -gt 0 -and $a.average_hr -le ($lthr - 10)) { $sentence2Parts += "FC controlada." }
    elseif ($a.average_hr -and $lthr -gt 0 -and $a.average_hr -gt ($lthr + 5)) { $sentence2Parts += "FC alta para o objetivo." }
    if ($wellFlags.Count -gt 0) { $sentence2Parts += ("Wellness baixo (" + ($wellFlags -join ", ") + ") pode ter influenciado.") }
    $summaryText = if ($sentence2Parts.Count -gt 0) { "$sentence1 " + ($sentence2Parts -join " ") } else { $sentence1 }

    if ($a.type -eq "Swim") {
      $intensitySummary = "Pace $paceValue | Volume $distanceText"
    } elseif ($a.type -eq "Run") {
      $intensitySummary = "Potência $powerText | NP $npText | IF $ifText | HR $hrText"
    } else {
      $intensitySummary = "Potência $powerText | NP $npText | VI $viText | IF $ifText | HR $hrText"
    }

    $metricCards = @()
    switch ($a.type) {
      "Run" {
        $metricCards += Build-MetricCard -Value $durationText -Label "Duração"
        $metricCards += Build-MetricCard -Value $distanceText -Label "Distância"
        $metricCards += Build-MetricCard -Value $paceValue -Label "Pace"
        $metricCards += Build-MetricCard -Value $tssText -Label "TSS"
        $metricCards += Build-MetricCard -Value $hrText -Label "FC média"
        $metricCards += Build-MetricCard -Value $powerText -Label "Potência"
        $metricCards += Build-MetricCard -Value $npText -Label "NP"
        $metricCards += Build-MetricCard -Value $ifText -Label "IF"
      }
      "Ride" {
        $metricCards += Build-MetricCard -Value $durationText -Label "Duração"
        $metricCards += Build-MetricCard -Value $distanceText -Label "Distância"
        $metricCards += Build-MetricCard -Value $paceValue -Label "Velocidade"
        $metricCards += Build-MetricCard -Value $tssText -Label "TSS"
        $metricCards += Build-MetricCard -Value $hrText -Label "FC média"
        $metricCards += Build-MetricCard -Value $powerText -Label "Potência"
        $metricCards += Build-MetricCard -Value $npText -Label "NP"
        $metricCards += Build-MetricCard -Value $ifText -Label "IF"
      }
      "Swim" {
        $metricCards += Build-MetricCard -Value $durationText -Label "Duração"
        $metricCards += Build-MetricCard -Value $distanceText -Label "Distância"
        $metricCards += Build-MetricCard -Value $paceValue -Label "Pace 100m"
        $metricCards += Build-MetricCard -Value $tssText -Label "TSS"
        $metricCards += Build-MetricCard -Value $hrText -Label "FC média"
      }
      default {
        $metricCards += Build-MetricCard -Value $durationText -Label "Duração"
        $metricCards += Build-MetricCard -Value $tssText -Label "TSS"
        $metricCards += Build-MetricCard -Value $hrText -Label "FC média"
      }
    }
    $metricsHtml = $metricCards -join ""

    $analysisItems = @()
    $analysisSummary = "<p class=""analysis-summary"">$summaryText</p>"
    $analysisItems += "<li><strong>Objetivo:</strong> $(Html-Escape $objective)</li>"
    $analysisItems += "<li><strong>Planejado vs executado:</strong> $(Html-Escape $planSummary)</li>"
    if ($a.type -eq "Run" -and $paceDeltaText) { $analysisItems += "<li><strong>Pace vs threshold:</strong> $paceValue ($paceDeltaText)</li>" }
    $analysisItems += "<li><strong>Intensidade:</strong> $intensitySummary</li>"
    if ($a.type -eq "Run") { $analysisItems += "<li><strong>Joelho:</strong> manter protocolo e nao progredir se dor > 4/10.</li>" }
    $analysisItems += "<li><strong>Wellness do dia:</strong> $wellText</li>"
    if ($a.notas) { $analysisItems += "<li><strong>Sua nota:</strong> $(Html-Escape $a.notas)</li>" }
    $analysisItems += "<li><strong>Recomendação:</strong> $recommendation</li>"
    $analysisHtml = $analysisSummary + "<ul>" + ($analysisItems -join "") + "</ul>"

    $activityCards += @"
<div class="activity-card activity-$typeClass">
  <div class="activity-title">
    <div class="activity-icon $typeClass">$typeIcon</div>
    <div>
      <div class="activity-name">$(Html-Escape $a.name)</div>
      <div class="activity-date">$(Html-Escape $a.start_date_local)</div>
    </div>
    <div class="activity-badge $qualityClass">$qualityLabel</div>
  </div>
  <div class="activity-metrics">
    $metricsHtml
  </div>
  <div class="analysis-block">
    <div class="analysis-title">Análise Completa</div>
    $analysisHtml
  </div>
</div>
"@
  }

  if ($activityCards.Count -eq 0) {
    $activityCards = @("<div class=""card section"">Sem atividades registradas no período.</div>")
  }

  $distGroups = $activities | Group-Object type | ForEach-Object {
    $timeMin = ($_.Group | Measure-Object moving_time_min -Sum).Sum
    [PSCustomObject]@{ type = $_.Name; time_h = if ($timeMin) { [math]::Round($timeMin / 60, 2) } else { 0 } }
  } | Sort-Object type
  $distTotal = ($distGroups | Measure-Object time_h -Sum).Sum
  $distLabels = ConvertTo-Json ($distGroups | ForEach-Object {
      $pct = if ($distTotal -gt 0) { [math]::Round(($_.time_h / $distTotal) * 100, 0) } else { 0 }
      "$($_.type) · $pct%"
    }) -Compress
  $distValues = ConvertTo-Json ($distGroups | ForEach-Object { $_.time_h }) -Compress
  $wellDates = ConvertTo-Json ($wellness | ForEach-Object { $_.data }) -Compress
  $ctlVals = ConvertTo-Json ($wellness | ForEach-Object { $_.ctl }) -Compress
  $atlVals = ConvertTo-Json ($wellness | ForEach-Object { $_.atl }) -Compress
  $sleepVals = ConvertTo-Json ($wellness | ForEach-Object { $_.sono_h }) -Compress
  $hrvVals = ConvertTo-Json ($wellness | ForEach-Object { $_.hrv }) -Compress
  $rhrVals = ConvertTo-Json ($wellness | ForEach-Object { $_.fc_reposo }) -Compress

  $html = @"
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Relatório de Coaching</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Sora:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root{
      --bg:#0b1022;
      --bg-2:#10172b;
      --card:#121a33;
      --card-2:#10172b;
      --card-soft:#0f1a2f;
      --text:#e5e7eb;
      --muted:#94a3b8;
      --accent:#ff6f91;
      --accent-2:#4bb6c1;
      --accent-3:#f6c453;
      --accent-4:#a855f7;
      --accent-run:#22c55e;
      --accent-ride:#f6c453;
      --accent-swim:#4bb6c1;
      --accent-strength:#a855f7;
      --shadow:0 18px 40px rgba(8,15,30,0.45);
    }
    *{box-sizing:border-box}
    body{
      margin:0;
      font-family:"Sora",sans-serif;
      background:radial-gradient(circle at 15% 15%, rgba(255,111,145,0.25) 0%, #0b1022 50%, #090d1b 100%);
      color:var(--text)
    }
    body::before{
      content:"";
      position:fixed;
      inset:0;
      background:
        radial-gradient(circle at 18% 18%, rgba(255,111,145,0.18), transparent 45%),
        radial-gradient(circle at 82% 20%, rgba(94,163,255,0.14), transparent 45%),
        radial-gradient(circle at 50% 82%, rgba(244,114,182,0.18), transparent 50%);
      pointer-events:none;
      z-index:-1;
    }
    .wrap{max-width:1200px;margin:24px auto;padding:0 20px}
    .hero{background:linear-gradient(135deg,#5f6ddf 0%,#8f6ccf 60%,#d17ca8 100%);padding:24px;border-radius:20px;display:flex;justify-content:space-between;align-items:center;box-shadow:var(--shadow);position:relative;overflow:hidden}
    .hero::after{content:"";position:absolute;right:-80px;top:-80px;width:220px;height:220px;background:rgba(255,255,255,0.12);border-radius:50%}
    .hero h1{margin:0;font-family:"Space Grotesk",sans-serif;font-size:24px}
    .hero p{margin:6px 0 0 0;color:#e2e8f0}
    .hero-pill{background:rgba(255,255,255,0.2);padding:8px 16px;border-radius:999px;font-weight:600}
    .hero-status{padding:10px 20px;border-radius:14px;font-weight:700}
    .status-hold{background:#f59e0b;color:#111827}
    .status-push{background:#22c55e;color:#052e16}
    .status-stepback{background:#ef4444;color:#fff}
    .section{margin-top:22px}
    .card{background:var(--card);padding:18px;border-radius:16px;box-shadow:var(--shadow);border:1px solid rgba(148,163,184,0.12)}
    .grid{display:grid;gap:16px}
    .grid-4{grid-template-columns:repeat(auto-fit,minmax(200px,1fr))}
    .grid-2{grid-template-columns:repeat(auto-fit,minmax(280px,1fr))}
    .metric{background:var(--card-soft);padding:14px;border-radius:12px;text-align:center}
    .metric .value{font-size:20px;font-weight:700}
    .metric .label{font-size:11px;color:var(--muted)}
    .performance-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
    .stat-card{background:var(--card-soft);border-radius:14px;padding:16px;min-height:140px;display:flex;flex-direction:column;justify-content:space-between}
    .stat-value{font-size:28px;font-weight:700;color:var(--accent)}
    .stat-label{font-size:12px;color:var(--muted);margin-top:6px}
    .stat-pill{margin-top:10px;display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:999px;background:rgba(15,23,42,0.45);font-size:11px;width:max-content}
    .stat-note{font-size:11px;color:var(--muted);margin-top:10px;line-height:1.3}
    .tsb-high{background:rgba(239,68,68,0.2);color:#ef4444}
    .tsb-med{background:rgba(245,158,11,0.2);color:#f59e0b}
    .tsb-low{background:rgba(148,163,184,0.2);color:#cbd5f5}
    .tsb-good{background:rgba(34,197,94,0.2);color:#22c55e}
    .comp-good{background:rgba(34,197,94,0.2);color:#22c55e}
    .comp-mid{background:rgba(245,158,11,0.2);color:#f59e0b}
    .comp-low{background:rgba(239,68,68,0.2);color:#ef4444}
    .comp-neutral{background:rgba(148,163,184,0.2);color:#cbd5f5}
    .race-card{display:flex;align-items:center;gap:18px;background:var(--card-soft);border-radius:14px;padding:16px;margin-bottom:12px;border-left:4px solid transparent}
    .race-card.priority-A{border-left-color:var(--accent-3)}
    .race-card.priority-B{border-left-color:var(--accent)}
    .race-card.priority-C{border-left-color:var(--accent-2)}
    .race-days{width:70px;text-align:center}
    .race-days .days{font-size:26px;font-weight:700;color:#7dd3fc}
    .race-days .label{font-size:10px;color:var(--muted)}
    .race-days .sub{font-size:10px;color:var(--muted)}
    .race-name{font-weight:600}
    .race-meta{color:var(--muted);font-size:12px}
    .race-estimate{color:#34d399;font-size:12px;margin-top:4px}
    .race-status{margin-top:6px;font-size:11px;color:#facc15}
    .race-badge{margin-left:auto;background:rgba(15,23,42,0.45);padding:6px 12px;border-radius:999px;font-size:11px}
    .wellness-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px}
    .wellness-card{background:var(--card-soft);border-radius:12px;padding:14px;text-align:center}
    .wellness-card .value{font-size:22px;font-weight:700}
    .wellness-card .label{font-size:11px;color:var(--muted)}
    .delta{margin-top:6px;font-size:10px}
    .delta.pos{color:#22c55e}
    .delta.neg{color:#ef4444}
    .delta.neu{color:var(--muted)}
    .well-icon{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;margin-right:6px}
    .well-icon svg{width:14px;height:14px}
    .info-icon{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:50%;background:rgba(15,23,42,0.45);color:#cbd5f5;text-decoration:none;margin-left:6px}
    .info-icon svg{width:10px;height:10px}
    .phase-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
    .phase-grid ul{margin:10px 0 0 18px;color:var(--muted)}
    .timeline{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-top:12px}
    .timeline .step{background:var(--card-soft);padding:14px;border-radius:12px;text-align:center}
    .activity-card{background:var(--card);border-radius:16px;padding:18px;margin-top:16px;border:1px solid rgba(148,163,184,0.12)}
    .activity-card.activity-run{border-left:4px solid var(--accent-run)}
    .activity-card.activity-ride{border-left:4px solid var(--accent-ride)}
    .activity-card.activity-swim{border-left:4px solid var(--accent-swim)}
    .activity-card.activity-strength{border-left:4px solid var(--accent-strength)}
    .activity-title{display:flex;gap:12px;align-items:center;margin-bottom:12px}
    .activity-icon{width:42px;height:42px;border-radius:12px;background:rgba(15,23,42,0.45);display:flex;align-items:center;justify-content:center;font-weight:700;font-size:11px;letter-spacing:.6px}
    .activity-icon svg,.rec-icon svg{width:20px;height:20px}
    .activity-icon.run{background:rgba(34,197,94,0.2);color:#22c55e}
    .activity-icon.ride{background:rgba(245,158,11,0.2);color:#f59e0b}
    .activity-icon.swim{background:rgba(56,189,248,0.2);color:#38bdf8}
    .activity-icon.strength{background:rgba(168,85,247,0.2);color:#a855f7}
    .activity-badge{margin-left:auto;padding:6px 12px;border-radius:999px;font-size:11px;font-weight:600}
    .quality-ok{background:rgba(34,197,94,0.2);color:#22c55e}
    .quality-low{background:rgba(148,163,184,0.2);color:#cbd5f5}
    .quality-high{background:rgba(245,158,11,0.2);color:#f59e0b}
    .quality-risk{background:rgba(239,68,68,0.2);color:#ef4444}
    .activity-name{font-weight:600}
    .activity-date{font-size:12px;color:var(--muted)}
    .activity-metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:12px}
    .analysis-block{background:var(--card-soft);border-left:4px solid var(--accent-3);padding:12px;border-radius:12px}
    .analysis-title{font-weight:600;margin-bottom:6px}
    .analysis-summary{margin:0 0 10px 0;color:#e2e8f0;font-size:13px}
    .analysis-block ul{margin:0 0 0 18px;color:var(--muted)}
    .analysis-block li{margin-bottom:8px}
    .coach-card{background:linear-gradient(180deg,#0f274a 0%, #0f1f39 100%);border:1px solid rgba(94,163,255,0.25)}
    .coach-header{display:flex;align-items:center;gap:14px;margin-bottom:14px}
    .coach-icon{width:44px;height:44px;border-radius:50%;background:#38bdf8;display:flex;align-items:center;justify-content:center;color:#0b1020}
    .coach-icon svg{width:22px;height:22px}
    .coach-subtitle{font-size:12px;color:var(--muted)}
    .coach-block{margin-top:14px}
    .coach-title{font-weight:600;margin-bottom:6px}
    .coach-block ul,.coach-block ol{margin:6px 0 0 18px;color:var(--muted)}
    .rec-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
    .rec-card{background:var(--card-soft);border-radius:14px;padding:16px}
    .rec-icon{width:40px;height:40px;border-radius:12px;display:flex;align-items:center;justify-content:center;margin-bottom:10px}
    .rec-icon.strength{background:rgba(168,85,247,0.2);color:#e9d5ff}
    .rec-icon.ride{background:rgba(245,158,11,0.2);color:#fde68a}
    .rec-icon.run{background:rgba(34,197,94,0.2);color:#bbf7d0}
    .rec-icon.swim{background:rgba(56,189,248,0.2);color:#bae6fd}
    .back-link{display:inline-block;margin-bottom:10px;font-size:12px;color:#7dd3fc;text-decoration:none}
    .rec-title{font-weight:600;margin-bottom:8px}
    .rec-card ul{margin:0 0 0 16px;color:var(--muted)}
    .glossary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;color:var(--muted)}
    .trend-score{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-top:16px}
    .plan-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px;margin-top:12px}
    .plan-card{background:var(--card-soft);border-radius:14px;padding:14px;border:1px solid rgba(148,163,184,0.12)}
    .plan-card h3{margin:0 0 8px 0;font-size:16px}
    .plan-note{margin-top:10px;color:var(--muted);font-size:13px}
    .score-card{background:var(--card-soft);border-radius:14px;padding:16px}
    .score-title{font-size:12px;color:var(--muted)}
    .score-value{font-size:24px;font-weight:700;margin-top:6px;color:var(--accent)}
    .score-note{font-size:11px;color:var(--muted);margin-top:6px}
    .lt-meta{margin-top:12px;font-size:12px;color:var(--muted)}
    .lt-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;margin-top:12px}
    .lt-panel{background:var(--card-soft);border-radius:14px;padding:16px}
    .lt-panel h3{margin:0 0 8px 0;font-size:13px}
    .lt-list{margin:0 0 0 16px;color:var(--muted)}
    .lt-list li{margin-bottom:6px}
    .lt-alert li{color:#fca5a5}
    .lt-empty{font-size:12px;color:var(--muted)}
    .lt-blocks{display:grid;gap:8px}
    .lt-block{background:var(--card-2);border-radius:10px;padding:10px;font-size:12px;color:var(--muted)}
  </style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <div>
      <h1>Relatório de Coaching</h1>
      <p>$athleteName | Semana $range</p>
    </div>
    <div style="display:flex;gap:12px;align-items:center">
      <div class="hero-pill">$phaseTitle</div>
      <div class="hero-status status-$classCss">$classification</div>
    </div>
  </div>

  <section class="card section">
    <h2>Calendário de Provas</h2>
    $($raceCards -join "`n")
  </section>

  <section class="card section">
    <h2>Performance da Semana</h2>
    $performanceCards
  </section>

  <section class="card section" id="wellness">
    <h2>Wellness da Semana</h2>
    $wellnessCards
  </section>

  <section class="card section">
    <h2>Fase do Ciclo: $phaseTitle</h2>
    <div class="phase-grid">
      <div>
        <h3>Objetivo desta fase</h3>
        <ul>$phaseObjectiveItems</ul>
      </div>
      <div>
        <h3>Transição para Build</h3>
        <ul>$phaseTransitionItems</ul>
      </div>
    </div>
    <div class="timeline">
      <div class="step">Esta semana<br><strong>$classification</strong></div>
      <div class="step">Semana +1<br><strong>Normal</strong></div>
      <div class="step">Semana +2<br><strong>Volume</strong></div>
      <div class="step">Semana +3<br><strong>Deload</strong></div>
    </div>
  </section>

  <section class="card section">
    <div class="grid grid-4">
      <div class="metric"><div class="value">$totalTime h</div><div class="label">Tempo total</div></div>
      <div class="metric"><div class="value">$totalDist km</div><div class="label">Distância</div></div>
      <div class="metric"><div class="value">$totalTss</div><div class="label">Carga (TSS)</div></div>
      <div class="metric"><div class="value">$ramp</div><div class="label">RampRate</div></div>
    </div>
    <div class="grid grid-4" style="margin-top:12px">
      <div class="metric"><div class="value">$activityCount</div><div class="label">Atividades</div></div>
      <div class="metric"><div class="value">$pesoText</div><div class="label">Peso</div></div>
      <div class="metric"><div class="value">$monotony</div><div class="label">Monotonia</div></div>
      <div class="metric"><div class="value">$strain</div><div class="label">Strain</div></div>
    </div>
  </section>

  $notesBlock

  $analysisBlock

  $feedbackBlock

  <section class="card section">
    <h2>Gráficos</h2>
    <div class="grid grid-2">
      <div class="card"><canvas id="dist-chart"></canvas></div>
      <div class="card"><canvas id="pmc-chart"></canvas></div>
      <div class="card"><canvas id="well-chart"></canvas></div>
    </div>
  </section>

  <section class="section">
    <h2>Análise das Atividades</h2>
    $($activityCards -join "`n")
  </section>

  <section class="card section">
    <h2>$trendTitle</h2>
    $trendCards
    <div class="grid grid-2" style="margin-top:16px">
      $($trendChartCards -join "`n")
    </div>
  </section>

  $longTermPlanBlock

  $recommendationsBlock

  $wellnessGlossary
</div>
<script>
  new Chart(document.getElementById('dist-chart'),{
    type:'doughnut',
    data:{labels:$distLabels,datasets:[{data:$distValues,backgroundColor:['#7dd3fc','#22d3ee','#34d399','#fbbf24','#f472b6'],borderColor:'#0b1020',borderWidth:2}]},
    options:{plugins:{legend:{position:'bottom',labels:{color:'#e2e8f0',boxWidth:10}}},cutout:'60%'}
  });
  new Chart(document.getElementById('pmc-chart'),{
    type:'line',
    data:{labels:$wellDates,datasets:[
      {label:'CTL',data:$ctlVals,borderColor:'#7dd3fc',backgroundColor:'rgba(125,211,252,0.15)',tension:.3,borderWidth:2,pointRadius:2},
      {label:'ATL',data:$atlVals,borderColor:'#fbbf24',backgroundColor:'rgba(251,191,36,0.15)',tension:.3,borderWidth:2,pointRadius:2}
    ]},
    options:{plugins:{legend:{position:'bottom',labels:{color:'#e2e8f0'}}},scales:{x:{display:false},y:{ticks:{color:'#94a3b8'}}}}
  });
  new Chart(document.getElementById('well-chart'),{
    type:'line',
    data:{labels:$wellDates,datasets:[
      {label:'Sono (h)',data:$sleepVals,borderColor:'#34d399',backgroundColor:'rgba(52,211,153,0.15)',tension:.3,borderWidth:2,pointRadius:2},
      {label:'HRV',data:$hrvVals,borderColor:'#a78bfa',backgroundColor:'rgba(167,139,250,0.15)',tension:.3,borderWidth:2,pointRadius:2},
      {label:'FC Repouso',data:$rhrVals,borderColor:'#f87171',backgroundColor:'rgba(248,113,113,0.15)',tension:.3,borderWidth:2,pointRadius:2}
    ]},
    options:{plugins:{legend:{position:'bottom',labels:{color:'#e2e8f0'}}},scales:{x:{display:false},y:{ticks:{color:'#94a3b8'}}}}
  });
  $($trendChartScripts -join "`n")
</script>
</body>
</html>
"@

  Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}

foreach ($report in $reportFiles) {
  $outputName = $report.Name -replace "\.json$", ".html"
  $outputPath = Join-Path $siteReports $outputName
  # Incremental build: keep past report HTML stable unless the source JSON changed.
  if (Should-RebuildReportHtml -ReportFile $report -OutputPath $outputPath -Force:$RebuildAll) {
    Build-ReportHtmlModern -ReportPath $report.FullName -OutputPath $outputPath
  }
}

  $memoryPath = Join-Path $repoRoot "COACHING_MEMORY.md"
  $memoryText = if (Test-Path $memoryPath) { Get-Content $memoryPath -Raw } else { "" }
  $athleteName = if ($memoryText -match "\*\*Nome:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "Atleta" }
  $athleteAge = if ($memoryText -match "\*\*Idade:\*\*\s*([0-9]+)") { $matches[1].Trim() } else { "" }
  $athleteHeight = if ($memoryText -match "\*\*Altura:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $athleteWeightRaw = if ($memoryText -match "\*\*Peso atual:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $athleteWeight = if ($athleteWeightRaw) { ($athleteWeightRaw -replace "\s*\(.*\)\s*", "").Trim() } else { "" }
  $athleteExp = if ($memoryText -match "\*\*Experiência:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $athleteLevel = if ($memoryText -match "\*\*Nível:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $phaseTitle = if ($memoryText -match "\*\*Fase:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "Base Geral" }
  $phaseFocus = if ($memoryText -match "\*\*Foco principal:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $phaseBike = if ($memoryText -match "\*\*Bike:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $phaseRun = if ($memoryText -match "\*\*Run:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $phaseSwim = if ($memoryText -match "\*\*Swim:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }
  $phaseForce = if ($memoryText -match "\*\*For[cç]a:\*\*\s*([^\r\n]+)") { (Fix-TextEncoding $matches[1].Trim()) } else { "" }

  $latestReport = if ($reportFiles.Count -gt 0) { $reportFiles[0] } else { $null }
  $latestData = if ($latestReport) { Get-Content $latestReport.FullName -Raw | ConvertFrom-Json } else { $null }
  $latestHtmlName = if ($latestReport) { ($latestReport.Name -replace "\.json$", ".html") } else { "reports" }
  $latestWeight = if ($latestData -and $latestData.metricas.peso_atual -ne $null) { [math]::Round($latestData.metricas.peso_atual, 1) } else { $null }
  $athleteWeightDisplay = if ($latestWeight -ne $null) { "$latestWeight kg" } elseif ($athleteWeight) { $athleteWeight } else { "n/a" }

  $bikeFtpValue = $null
  if ($memoryText -match "\*\*FTP:\*\*\s*([0-9,\.]+)") {
    $bikeFtpValue = [double](($matches[1] -replace ",", "."))
  }
  $bikeFtpText = if ($bikeFtpValue -ne $null) { "{0:0} W" -f $bikeFtpValue } else { "n/a" }

  $runLthrValue = $null
  if ($memoryText -match "\*\*LTHR:\*\*\s*([0-9,\.]+)") {
    $runLthrValue = [double](($matches[1] -replace ",", "."))
  }
  $runThresholdPace = ""
  if ($memoryText -match "Threshold Pace:\*{0,2}\s*~?([0-9:]+)/km") {
    $runThresholdPace = $matches[1]
  }
  $runLactateText = if ($runLthrValue -ne $null -and $runThresholdPace) {
    "{0:0} bpm · {1}/km" -f $runLthrValue, $runThresholdPace
  } elseif ($runThresholdPace) {
    "$runThresholdPace/km"
  } elseif ($runLthrValue -ne $null) {
    "{0:0} bpm" -f $runLthrValue
  } else {
    "n/a"
  }

  $vo2RunValue = if ($latestData -and $latestData.metricas.vo2_run -ne $null) { $latestData.metricas.vo2_run } else { $null }
  $vo2BikeValue = if ($latestData -and $latestData.metricas.vo2_bike -ne $null) { $latestData.metricas.vo2_bike } else { $null }
  $vo2RunText = if ($vo2RunValue -ne $null) { "{0:0} ml/kg/min" -f $vo2RunValue } else { "n/a" }
  $vo2BikeText = if ($vo2BikeValue -ne $null) { "{0:0} ml/kg/min" -f $vo2BikeValue } else { "n/a" }
  if ($vo2RunText -eq "n/a" -and $memoryText -match "\*\*VO2 Run:\*\*\s*([^\r\n]+)") { $vo2RunText = (Fix-TextEncoding $matches[1].Trim()) }
  if ($vo2BikeText -eq "n/a" -and $memoryText -match "\*\*VO2 Bike:\*\*\s*([^\r\n]+)") { $vo2BikeText = (Fix-TextEncoding $matches[1].Trim()) }
  if ($vo2RunText -eq "n/a" -and $memoryText -match "\*\*VO2max Run:\*\*\s*([^\r\n]+)") { $vo2RunText = (Fix-TextEncoding $matches[1].Trim()) }
  if ($vo2BikeText -eq "n/a" -and $memoryText -match "\*\*VO2max Bike:\*\*\s*([^\r\n]+)") { $vo2BikeText = (Fix-TextEncoding $matches[1].Trim()) }
  if ($vo2RunText -eq "n/a" -and $runThresholdPace) {
    $secs = Pace-To-Secs -Pace $runThresholdPace
    if ($secs -gt 0) {
      $v = 60000 / $secs
      $vo2 = (-4.6 + (0.182258 * $v) + (0.000104 * $v * $v))
      $t = 60
      $pct = 0.8 + (0.1894393 * [math]::Exp(-0.012778 * $t)) + (0.2989558 * [math]::Exp(-0.1932605 * $t))
      $vdot = if ($pct -gt 0) { $vo2 / $pct } else { $vo2 }
      $vo2RunText = "{0:0} ml/kg/min (est.)" -f $vdot
    }
  }
  if ($vo2BikeText -eq "n/a" -and $bikeFtpValue -ne $null -and $latestWeight -ne $null -and $latestWeight -gt 0) {
    $map = $bikeFtpValue / 0.75
    $vo2b = (10.8 * ($map / $latestWeight)) + 7
    $vo2BikeText = "{0:0} ml/kg/min (est.)" -f $vo2b
  }

  $futurePhases = Get-MemorySectionLines -Text $memoryText -HeaderPattern "### Pr[oó]ximas Fases \(planejado\)"
  function Build-PhaseItems {
    param([string[]]$Lines)
    if (-not $Lines -or $Lines.Count -eq 0) { return "" }
    $items = @()
    foreach ($lineRaw in $Lines) {
      $line = (Fix-TextEncoding $lineRaw).Trim()
      if (-not $line) { continue }
      $range = ""
      $desc = $line
      if ($line -match "^([0-9]{2}/[0-9]{2}-[0-9]{2}/[0-9]{2})\s*:\s*(.+)$") {
        $range = $matches[1].Trim()
        $desc = $matches[2].Trim()
      }
      $objective = $desc
      $focus = ""
      $splitIndex = -1
      $depth = 0
      for ($i = 0; $i -lt $desc.Length; $i++) {
        $ch = $desc[$i]
        if ($ch -eq '(') { $depth++ }
        elseif ($ch -eq ')' -and $depth -gt 0) { $depth-- }
        if ($depth -eq 0 -and $ch -eq '+' -and $i -gt 0 -and $i -lt ($desc.Length - 1)) {
          if ($desc[$i - 1] -eq ' ' -and $desc[$i + 1] -eq ' ') {
            $splitIndex = $i
            break
          }
        }
      }
      if ($splitIndex -ge 0) {
        $objective = $desc.Substring(0, $splitIndex).Trim()
        $focus = $desc.Substring($splitIndex + 1).Trim().TrimStart('+').Trim()
      } elseif ($desc -match "(.+?)\s*\((.+)\)") {
        $objective = $matches[1].Trim()
        $focus = $matches[2].Trim()
      }
      if (-not $focus) { $focus = "Manter consistência e técnica." }
      $rangeHtml = if ($range) { "<div class=""phase-range"">$range</div>" } else { "<div class=""phase-range"">&nbsp;</div>" }
      $items += @"
<div class="phase-item">
  $rangeHtml
  <div class="phase-lines">
    <div class="phase-line"><strong>Objetivo:</strong> $objective</div>
    <div class="phase-line"><strong>Foco:</strong> $focus</div>
  </div>
</div>
"@
    }
    return ($items -join "`n")
  }
  $futurePhaseItems = Build-PhaseItems -Lines $futurePhases
  if (-not $futurePhaseItems) {
    $futurePhaseItems = '<div class="muted">Sem fases cadastradas.</div>'
  }
  $phaseNotes = @()
  if ($phaseFocus) { $phaseNotes += "Foco: $phaseFocus" }
  if ($phaseBike) { $phaseNotes += "Bike: $phaseBike" }
  if ($phaseRun) { $phaseNotes += "Run: $phaseRun" }
  if ($phaseSwim) { $phaseNotes += "Swim: $phaseSwim" }
  if ($phaseForce) { $phaseNotes += "Força: $phaseForce" }
  $phaseNotesItems = if ($phaseNotes.Count -gt 0) { ($phaseNotes | ForEach-Object { "<li>$_</li>" }) -join "" } else { "" }
  $phaseNotesHtml = if ($phaseNotes.Count -gt 0) { "<ul class=""phase-notes"">$phaseNotesItems</ul>" } else { "<div class=""muted"">Sem diretrizes adicionais.</div>" }
  $phaseVolumeTarget = "n/a"
  if ($latestData -and $latestData.semana.tempo_total_horas -ne $null) {
    $baseHours = [double]$latestData.semana.tempo_total_horas
    $low = [math]::Round([math]::Max(1, ($baseHours * 0.9)), 1)
    $high = [math]::Round(($baseHours * 1.1), 1)
    $phaseVolumeTarget = "$low–$high h/sem"
  } else {
    $phaseVolumeTarget = "7–9 h/sem"
  }
  $phaseIntensity = if ($phaseBike -match "Sweet Spot|FTP") { "1–2 sessões Sweet Spot/FTP" } else { "1–2 sessões de qualidade" }
  $phaseStrength = if ($phaseForce) { $phaseForce } else { "2x/semana" }

  $calendarEvents = Get-MemoryCalendar -Text $memoryText | Sort-Object date
  $today = (Get-Date).Date
  $calendarRunPace = 6.5
  $calendarBikeSpeed = 28
  $calendarSwimPace100 = 2.2
  if ($latestData) {
    $acts = @($latestData.atividades)
    $runDist = ($acts | Where-Object type -eq "Run" | Measure-Object distance_km -Sum).Sum
    $runTime = ($acts | Where-Object type -eq "Run" | Measure-Object moving_time_min -Sum).Sum
    if ($runDist -gt 0) { $calendarRunPace = $runTime / $runDist }
    if ($memoryText -match "Threshold Pace:\s*~?([0-9:]+)/km") {
      $p = $matches[1]
      $secs = Pace-To-Secs -Pace $p
      if ($secs) { $calendarRunPace = $secs / 60 * 1.15 }
    }
    $bikeDist = ($acts | Where-Object type -eq "Ride" | Measure-Object distance_km -Sum).Sum
    $bikeTime = ($acts | Where-Object type -eq "Ride" | Measure-Object moving_time_min -Sum).Sum
    if ($bikeTime -gt 0) { $calendarBikeSpeed = $bikeDist / ($bikeTime / 60) }
    $swimDist = ($acts | Where-Object type -eq "Swim" | Measure-Object distance_km -Sum).Sum
    $swimTime = ($acts | Where-Object type -eq "Swim" | Measure-Object moving_time_min -Sum).Sum
    if ($swimDist -gt 0) { $calendarSwimPace100 = $swimTime / ($swimDist * 10) }
  }
  function Estimate-Race-Index {
    param(
      [string]$Type,
      [string]$Name,
      [double]$RunPace,
      [double]$BikeSpeed,
      [double]$SwimPace100
    )
    if (-not $RunPace -or $RunPace -le 0) { $RunPace = 6.5 }
    if (-not $BikeSpeed -or $BikeSpeed -le 0) { $BikeSpeed = 28 }
    if (-not $SwimPace100 -or $SwimPace100 -le 0) { $SwimPace100 = 2.2 }
    if ($Type -match "Corrida") {
      $dist = 10
      if ($Name -match "Meia") { $dist = 21.1 }
      $mins = $RunPace * $dist
      return Format-Duration-Short -Minutes $mins
    }
    if ($Type -match "Triathlon Sprint" -or $Name -match "Triathlon Sprint") {
      $swim = ($SwimPace100 * 7.5)
      $bike = if ($BikeSpeed -gt 0) { (20 / $BikeSpeed) * 60 } else { 50 }
      $run = $RunPace * 5
      return Format-Duration-Short -Minutes ($swim + $bike + $run + 4)
    }
    if ($Type -match "Ol" -or $Name -match "Olimp") {
      $swim = ($SwimPace100 * 15)
      $bike = if ($BikeSpeed -gt 0) { (40 / $BikeSpeed) * 60 } else { 90 }
      $run = $RunPace * 10
      return Format-Duration-Short -Minutes ($swim + $bike + $run + 6)
    }
    if ($Name -match "70\.3" -or $Type -match "70\.3") {
      $swim = ($SwimPace100 * 19)
      $bike = if ($BikeSpeed -gt 0) { (90 / $BikeSpeed) * 60 } else { 210 }
      $run = $RunPace * 21.1
      return Format-Duration-Short -Minutes ($swim + $bike + $run + 8)
    }
    return "n/a"
  }
  $calendarCards = @()
  foreach ($ev in $calendarEvents) {
    $days = if ($ev.date) { ($ev.date.Date - $today).Days } else { $null }
    $daysText = if ($days -ne $null) { [math]::Max($days, 0) } else { "n/a" }
    $prioClass = if ($ev.priority -match "A") { "tag-a" } elseif ($ev.priority -match "B") { "tag-b" } else { "tag-c" }
    $estimate = Estimate-Race-Index -Type $ev.type -Name $ev.name -RunPace $calendarRunPace -BikeSpeed $calendarBikeSpeed -SwimPace100 $calendarSwimPace100
    $calendarCards += @"
      <div class="event-card $prioClass">
        <div class="event-days">$daysText<span>dias</span></div>
        <div class="event-info">
          <div class="event-name">$(Html-Escape $ev.name)</div>
          <div class="event-meta">$(Html-Escape $ev.type) | Prioridade $(Html-Escape $ev.priority)</div>
          <div class="event-meta">Estimativa hoje: $estimate</div>
          <div class="event-date">$(Html-Escape $ev.date_raw)</div>
        </div>
        <div class="event-tag">$([string]$ev.status)</div>
      </div>
"@
  }
  $calendarHtml = if ($calendarCards.Count -gt 0) { ($calendarCards -join "`n") } else { "<div class=""muted"">Sem provas cadastradas.</div>" }

  $ctl = if ($latestData) { $latestData.metricas.CTL } else { $null }
  $atl = if ($latestData) { $latestData.metricas.ATL } else { $null }
  $tsb = if ($latestData) { $latestData.metricas.TSB } else { $null }
  $ctlText = if ($ctl -ne $null) { [math]::Round($ctl,1) } else { "n/a" }
  $atlText = if ($atl -ne $null) { [math]::Round($atl,1) } else { "n/a" }
  $tsbText = if ($tsb -ne $null) { [math]::Round($tsb,1) } else { "n/a" }

  $plannedEvents = if ($latestData -and $latestData.PSObject.Properties.Name -contains "treinos_planejados") { @($latestData.treinos_planejados) } else { @() }
  $latestActivities = if ($latestData) { @($latestData.atividades) } else { @() }
  $adherence = Get-PlanAdherenceSummary -PlannedEvents $plannedEvents -Activities $latestActivities
  $complianceValue = $adherence.adherence_overall
  $complianceText = if ($complianceValue -ne $null) { "{0:0.0}%" -f $complianceValue } else { "n/a" }

  $reportsList = @()
  foreach ($file in $reportFiles) {
    $name = $file.Name
    $htmlName = ($file.Name -replace "\.json$", ".html")
    $label = $null
    if ($htmlName -match "report_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})") {
      $start = $null; $end = $null
      try { $start = [DateTime]::ParseExact($matches[1], "yyyy-MM-dd", $null) } catch { $start = $null }
      try { $end = [DateTime]::ParseExact($matches[2], "yyyy-MM-dd", $null) } catch { $end = $null }
      if ($start -and $end) { $label = "{0:dd/MM} a {1:dd/MM}" -f $start, $end }
    }
    if (-not $label) { $label = $htmlName }
    $reportsList += "<div class=""link-card""><a href=""reports/$htmlName"">$label</a></div>"
  }
  function Format-MonthLabel {
    param([string]$MonthKey)
    if (-not $MonthKey) { return "" }
    $dt = $null
    try { $dt = [DateTime]::ParseExact("$MonthKey-01", "yyyy-MM-dd", $null) } catch { $dt = $null }
    if ($dt) { return $dt.ToString("MMM yyyy").ToLower() }
    return $MonthKey
  }

  function Format-ReportLabel {
    param([string]$FileName)
    if (-not $FileName) { return "Relatorio" }
    if ($FileName -match "report_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})") {
      $start = $null; $end = $null
      try { $start = [DateTime]::ParseExact($matches[1], "yyyy-MM-dd", $null) } catch { $start = $null }
      try { $end = [DateTime]::ParseExact($matches[2], "yyyy-MM-dd", $null) } catch { $end = $null }
      if ($start -and $end) { return "{0:dd/MM} a {1:dd/MM}" -f $start, $end }
    }
    return ($FileName -replace "\.html$", "")
  }

  $reportGroups = @{}
  foreach ($file in $reportFiles) {
    $name = $file.Name
    $monthKey = ""
    if ($name -match "report_(\d{4}-\d{2})-\d{2}_") { $monthKey = $matches[1] }
    if (-not $monthKey) { $monthKey = "outros" }
    if (-not $reportGroups.ContainsKey($monthKey)) { $reportGroups[$monthKey] = @() }
    $reportGroups[$monthKey] += $file
  }
  $monthKeys = @($reportGroups.Keys | Sort-Object -Descending)
  $archiveButtons = @()
  foreach ($m in $monthKeys) {
    $label = Format-MonthLabel -MonthKey $m
    $archiveButtons += "<button class=""chip"" data-month-btn=""$m"">$label</button>"
  }
  $archiveGroups = @()
  foreach ($m in $monthKeys) {
    $items = @()
    foreach ($file in $reportGroups[$m]) {
      $htmlName = ($file.Name -replace "\.json$", ".html")
      $label = Format-ReportLabel -FileName $htmlName
      $items += "<div class=""link-card""><a href=""reports/$htmlName"">$label</a></div>"
    }
    $archiveGroups += "<div class=""archive-group"" data-month-group=""$m"">$($items -join "`n")</div>"
  }

  $enableThaisPage = $false
  if ($enableThaisPage) {
  $thaisDir = Join-Path $SiteDir "thais"
  $thaisAssetsDir = Join-Path $thaisDir "assets"
  New-Item -ItemType Directory -Path $thaisAssetsDir -Force | Out-Null

  $thaisAssetsSrc = Join-Path $repoRoot "assets\\thais"
  if (Test-Path $thaisAssetsSrc) {
    Copy-Item -Path (Join-Path $thaisAssetsSrc "*") -Destination $thaisAssetsDir -Force -ErrorAction SilentlyContinue
  }

  $thaisHeroImage1 = "assets/hero-1.svg"
  $thaisHeroImage2 = "assets/hero-2.svg"
  if (Test-Path $thaisAssetsSrc) {
    $thaisPhotos = Get-ChildItem -Path $thaisAssetsSrc -File | Where-Object { $_.Extension -match "\.jpe?g|\.png" } | Sort-Object Name
    if ($thaisPhotos.Count -ge 1) { $thaisHeroImage1 = "assets/$($thaisPhotos[0].Name)" }
    if ($thaisPhotos.Count -ge 2) { $thaisHeroImage2 = "assets/$($thaisPhotos[1].Name)" }
  }

  $thaisReportsDir = "C:\\Users\\Andre\\OneDrive\\Triathlon_Semanal_Andre_Codex\\Relatorios_Intervals"
  if (-not (Test-Path $thaisReportsDir)) { $thaisReportsDir = $ReportsDir }
  $thaisReportFiles = Get-ChildItem $thaisReportsDir -Filter "report_*.json" | Sort-Object Name -Descending
  $thaisAnalysisFiles = Get-ChildItem $thaisReportsDir -Filter "analysis_*.md" -ErrorAction SilentlyContinue | Sort-Object Name -Descending

  $thaisSiteReports = Join-Path $thaisDir "reports"
  New-Item -ItemType Directory -Path $thaisSiteReports -Force | Out-Null
  Copy-Item -Path (Join-Path $thaisReportsDir "*.json") -Destination $thaisSiteReports -Force -ErrorAction SilentlyContinue
  Copy-Item -Path (Join-Path $thaisReportsDir "*.md") -Destination $thaisSiteReports -Force -ErrorAction SilentlyContinue

  $mainReportsDir = $ReportsDir
  $mainAnalysisFiles = $analysisFiles
  $ReportsDir = $thaisReportsDir
  $analysisFiles = $thaisAnalysisFiles
   foreach ($report in $thaisReportFiles) {
     $outputName = $report.Name -replace "\.json$", ".html"
     $outputPath = Join-Path $thaisSiteReports $outputName
    # Incremental build: keep past report HTML stable unless the source JSON changed.
    if (Should-RebuildReportHtml -ReportFile $report -OutputPath $outputPath -Force:$RebuildAll) {
      Build-ReportHtmlModern -ReportPath $report.FullName -OutputPath $outputPath
    }
   }
  $ReportsDir = $mainReportsDir
  $analysisFiles = $mainAnalysisFiles

  $thaisLatestReport = if ($thaisReportFiles.Count -gt 0) { $thaisReportFiles[0] } else { $null }
  $thaisLatestData = if ($thaisLatestReport) { Get-Content $thaisLatestReport.FullName -Raw | ConvertFrom-Json } else { $null }
  $thaisLatestHtmlName = if ($thaisLatestReport) { ($thaisLatestReport.Name -replace "\.json$", ".html") } else { "reports" }
  $thaisCtl = if ($thaisLatestData) { $thaisLatestData.metricas.CTL } else { $null }
  $thaisAtl = if ($thaisLatestData) { $thaisLatestData.metricas.ATL } else { $null }
  $thaisTsb = if ($thaisLatestData) { $thaisLatestData.metricas.TSB } else { $null }
  $thaisCtlText = if ($thaisCtl -ne $null) { [math]::Round($thaisCtl,1) } else { "n/a" }
  $thaisAtlText = if ($thaisAtl -ne $null) { [math]::Round($thaisAtl,1) } else { "n/a" }
  $thaisTsbText = if ($thaisTsb -ne $null) { [math]::Round($thaisTsb,1) } else { "n/a" }

  $thaisPlannedEvents = if ($thaisLatestData -and $thaisLatestData.PSObject.Properties.Name -contains "treinos_planejados") { @($thaisLatestData.treinos_planejados) } else { @() }
  $thaisPlannedCount = $thaisPlannedEvents.Count
  $thaisMatchedCount = if ($thaisLatestData) { (@($thaisLatestData.atividades) | Where-Object { $_.planejado }).Count } else { 0 }
  $thaisComplianceValue = if ($thaisPlannedCount -gt 0) { [math]::Round((($thaisMatchedCount / $thaisPlannedCount) * 100), 1) } else { $null }
  $thaisComplianceText = if ($thaisComplianceValue -ne $null) { "{0:0.0}%" -f $thaisComplianceValue } else { "n/a" }

  $thaisReportsList = @()
  foreach ($file in $thaisReportFiles) {
    $htmlName = ($file.Name -replace "\.json$", ".html")
    $label = $null
    if ($htmlName -match "report_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})") {
      $start = $null; $end = $null
      try { $start = [DateTime]::ParseExact($matches[1], "yyyy-MM-dd", $null) } catch { $start = $null }
      try { $end = [DateTime]::ParseExact($matches[2], "yyyy-MM-dd", $null) } catch { $end = $null }
      if ($start -and $end) { $label = "{0:dd/MM} a {1:dd/MM}" -f $start, $end }
    }
    if (-not $label) { $label = $htmlName }
    $thaisReportsList += "<div class=""link-card""><a href=""reports/$htmlName"">$label</a></div>"
  }

  $memoryPath = Join-Path $repoRoot "COACHING_MEMORY.md"
  $memoryText = if (Test-Path $memoryPath) { Get-Content $memoryPath -Raw } else { "" }

  $thaisToday = (Get-Date).Date
  $thaisCalendarEvents = Get-MemoryCalendar -Text $memoryText | Sort-Object date
  $thaisRaceCards = @()
  $thaisRunPace = 6.5
  $thaisBikeSpeed = 28
  $thaisSwimPace100 = 2.2
  if ($thaisLatestData) {
    $thaisActs = @($thaisLatestData.atividades)
    $runDist = ($thaisActs | Where-Object type -eq "Run" | Measure-Object distance_km -Sum).Sum
    $runTime = ($thaisActs | Where-Object type -eq "Run" | Measure-Object moving_time_min -Sum).Sum
    if ($runDist -gt 0) { $thaisRunPace = $runTime / $runDist }
    if ($memoryText -match "Threshold Pace:\s*~?([0-9:]+)/km") {
      $p = $matches[1]
      $secs = Pace-To-Secs -Pace $p
      if ($secs) { $thaisRunPace = $secs / 60 * 1.15 }
    }
    $bikeDist = ($thaisActs | Where-Object type -eq "Ride" | Measure-Object distance_km -Sum).Sum
    $bikeTime = ($thaisActs | Where-Object type -eq "Ride" | Measure-Object moving_time_min -Sum).Sum
    if ($bikeTime -gt 0) { $thaisBikeSpeed = $bikeDist / ($bikeTime / 60) }
    $swimDist = ($thaisActs | Where-Object type -eq "Swim" | Measure-Object distance_km -Sum).Sum
    $swimTime = ($thaisActs | Where-Object type -eq "Swim" | Measure-Object moving_time_min -Sum).Sum
    if ($swimDist -gt 0) { $thaisSwimPace100 = $swimTime / ($swimDist * 10) }
  }
  function Estimate-Race-Thais {
    param(
      [string]$Type,
      [string]$Name,
      [double]$RunPace,
      [double]$BikeSpeed,
      [double]$SwimPace100
    )
    if (-not $RunPace -or $RunPace -le 0) { $RunPace = 6.5 }
    if (-not $BikeSpeed -or $BikeSpeed -le 0) { $BikeSpeed = 28 }
    if (-not $SwimPace100 -or $SwimPace100 -le 0) { $SwimPace100 = 2.2 }
    if ($Type -eq "Corrida") {
      $dist = 10
      if ($Name -match "Meia") { $dist = 21.1 }
      $mins = $RunPace * $dist
      return Format-Duration -Minutes $mins
    }
    if ($Type -match "Ol" -or $Name -match "Olimp") {
      $swim = ($SwimPace100 * 15)
      $bike = if ($BikeSpeed -gt 0) { (40 / $BikeSpeed) * 60 } else { 90 }
      $run = $RunPace * 10
      return Format-Duration -Minutes ($swim + $bike + $run + 6)
    }
    if ($Name -match "70\.3" -or $Type -match "70\.3") {
      $swim = ($SwimPace100 * 19)
      $bike = if ($BikeSpeed -gt 0) { (90 / $BikeSpeed) * 60 } else { 210 }
      $run = $RunPace * 21.1
      return Format-Duration -Minutes ($swim + $bike + $run + 8)
    }
    return "n/a"
  }
  foreach ($ev in $thaisCalendarEvents) {
    $days = if ($ev.date) { ($ev.date.Date - $thaisToday).Days } else { $null }
    $daysText = if ($days -ne $null) { [math]::Max($days, 0) } else { "n/a" }
    $weeksText = if ($days -ne $null) { [math]::Floor([math]::Max($days, 0) / 7) } else { "n/a" }
    $priority = if ($ev.priority) { $ev.priority } else { "C" }
    $stage = if ($priority -eq "A") { "Peak" } elseif ($priority -eq "B") { "Especifico" } else { "Build" }
    $estimate = Estimate-Race-Thais -Type $ev.type -Name $ev.name -RunPace $thaisRunPace -BikeSpeed $thaisBikeSpeed -SwimPace100 $thaisSwimPace100
    $thaisRaceCards += @"
<div class="race-card priority-$priority">
  <div class="race-days">
    <div class="days">$daysText</div>
    <div class="label">DIAS</div>
    <div class="sub">$weeksText sem</div>
  </div>
  <div class="race-info">
    <div class="race-name">$(Html-Escape $ev.name)</div>
    <div class="race-meta">$(Html-Escape $ev.type) | Prioridade $priority</div>
    <div class="race-estimate">Estimativa hoje: $estimate</div>
    <div class="race-date">$(Html-Escape $ev.date_raw)</div>
  </div>
  <div class="race-badge">$stage</div>
</div>
"@
  }
  $thaisCalendarHtml = if ($thaisRaceCards.Count -gt 0) { ($thaisRaceCards -join "`n") } else { "<div class=""muted"">Sem provas cadastradas.</div>" }

  $phaseObjectives = @()
  if ($phaseFocus) { $phaseObjectives += $phaseFocus }
  $phaseRun = if ($memoryText -match "\*\*Run:\*\*\s*([^\r\n]+)") { $matches[1].Trim() } else { "" }
  $phaseSwim = if ($memoryText -match "\*\*Swim:\*\*\s*([^\r\n]+)") { $matches[1].Trim() } else { "" }
  $phaseForce = if ($memoryText -match "\*\*Força:\*\*\s*([^\r\n]+)") { $matches[1].Trim() } else { "" }
  if ($phaseRun) { $phaseObjectives += "Run: $phaseRun" }
  if ($phaseSwim) { $phaseObjectives += "Swim: $phaseSwim" }
  if ($phaseForce) { $phaseObjectives += "Força: $phaseForce" }
  $phaseTransition = @(
    "Introduzir intervalos de Threshold (bike)",
    "Tiros curtos na corrida (se joelho permitir)",
    "Natação: aumentar volume + séries de ritmo"
  )
  $phaseObjectiveItems = if ($phaseObjectives.Count -gt 0) { ($phaseObjectives | ForEach-Object { "<li>$_</li>" }) -join "" } else { "<li>Sem objetivos cadastrados.</li>" }
  $phaseTransitionItems = ($phaseTransition | ForEach-Object { "<li>$_</li>" }) -join ""

  $thaisHtml = @"
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Thais Lourenço | Triathlon</title>
  <link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=Manrope:wght@300;400;600;700&display=swap" rel="stylesheet">
  <style>
    :root{
      --bg:#fff4f6;
      --ink:#1f1b24;
      --muted:#6b6074;
      --accent:#ff6f91;
      --accent-2:#4bb6c1;
      --accent-3:#f6c453;
      --card:#ffffff;
      --shadow:0 24px 55px rgba(31,27,36,0.14);
    }
    *{box-sizing:border-box}
    body{
      margin:0;
      font-family:"Manrope",sans-serif;
      color:var(--ink);
      background:
        radial-gradient(circle at 10% 10%, rgba(255,111,145,0.18), transparent 45%),
        radial-gradient(circle at 85% 15%, rgba(75,182,193,0.18), transparent 50%),
        radial-gradient(circle at 50% 80%, rgba(246,196,83,0.18), transparent 55%),
        var(--bg);
    }
    .wrap{max-width:1200px;margin:28px auto;padding:0 22px}
    .hero{
      background:linear-gradient(135deg,#ffffff 0%, #ffe6ee 55%, #fbe7ff 100%);
      border-radius:26px;
      padding:28px;
      display:grid;
      grid-template-columns:1.1fr 0.9fr;
      gap:24px;
      box-shadow:var(--shadow);
      align-items:center;
    }
    .hero h1{
      font-family:"DM Serif Display",serif;
      font-size:36px;
      margin:0 0 8px 0;
      letter-spacing:.2px;
    }
    .hero p{margin:6px 0;color:var(--muted)}
    .hero .pill{
      display:inline-flex;align-items:center;gap:10px;
      background:#fff;
      border:1px solid rgba(0,0,0,0.05);
      padding:8px 14px;
      border-radius:999px;
      font-weight:600;
      color:#d14b6a;
    }
    .hero-images{
      display:grid;
      grid-template-columns:1fr 1fr;
      gap:12px;
    }
    .hero-images img{
      width:100%;
      height:220px;
      object-fit:cover;
      border-radius:18px;
      box-shadow:var(--shadow);
      border:1px solid rgba(0,0,0,0.06);
    }
    .grid{display:grid;gap:16px}
    .grid-3{grid-template-columns:repeat(auto-fit,minmax(220px,1fr))}
    .grid-2{grid-template-columns:repeat(auto-fit,minmax(280px,1fr))}
    .grid-4{grid-template-columns:repeat(auto-fit,minmax(200px,1fr))}
    .card{
      background:var(--card);
      border-radius:18px;
      padding:18px;
      box-shadow:var(--shadow);
      border:1px solid rgba(0,0,0,0.04);
    }
    .kpi{
      font-size:26px;
      font-weight:700;
      color:#322a3a;
    }
    .label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.12em}
    h2{
      margin:26px 0 12px 0;
      font-size:18px;
      font-weight:700;
    }
    .accent-line{
      height:4px;border-radius:999px;background:linear-gradient(90deg,var(--accent),var(--accent-2));
      width:86px;margin:10px 0 16px 0;
    }
    .link-card{
      display:flex;justify-content:space-between;align-items:center;
      background:#fff7fb;border-radius:12px;padding:10px 12px;
      border:1px solid rgba(209,75,106,0.12);
    }
    .link-card a{color:#c43662;text-decoration:none;font-weight:600}
    .link-card a:hover{text-decoration:underline}
    .badge{
      display:inline-flex;align-items:center;gap:8px;
      font-size:14px;padding:10px 16px;border-radius:999px;
      background:#1f1b24;color:#fff;font-weight:600;
    }
    .race-card{
      display:flex;align-items:center;gap:18px;background:#fff;border-radius:16px;padding:16px;
      border:1px solid rgba(209,75,106,0.12);box-shadow:var(--shadow);
      margin-bottom:12px;
    }
    .race-card.priority-A{border-left:4px solid var(--accent-3)}
    .race-card.priority-B{border-left:4px solid var(--accent)}
    .race-card.priority-C{border-left:4px solid var(--accent-2)}
    .race-days{width:74px;text-align:center}
    .race-days .days{font-size:24px;font-weight:700;color:#c43662}
    .race-days .label{font-size:10px;color:var(--muted)}
    .race-days .sub{font-size:10px;color:var(--muted)}
    .race-name{font-weight:700}
    .race-meta,.race-date{font-size:12px;color:var(--muted)}
    .race-estimate{font-size:12px;color:#c43662;font-weight:600}
    .race-badge{margin-left:auto;background:#fff1f5;padding:6px 12px;border-radius:999px;font-size:11px;font-weight:700;color:#b82d55}
    .phase-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
    .timeline{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-top:12px}
    .timeline .step{background:#fff;border-radius:12px;padding:12px;text-align:center;border:1px solid rgba(209,75,106,0.12)}
    .split{
      display:grid;grid-template-columns:1.1fr 0.9fr;gap:16px;
    }
    .cta{
      display:inline-flex;align-items:center;gap:10px;
      background:#1f1b24;color:#fff;border-radius:999px;padding:10px 16px;text-decoration:none;font-weight:600;
    }
    .cta:hover{opacity:.92}
    @media (max-width: 900px){
      .hero{grid-template-columns:1fr}
      .hero-images img{height:180px}
      .split{grid-template-columns:1fr}
    }
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <div>
        <span class="pill">Triathlon Sprint • Temporada 2026</span>
        <h1>Thais Lourenço</h1>
        <p>Relatórios e acompanhamento semanal com foco em performance sustentável, técnica e consistência.</p>
        <div style="margin-top:16px;display:flex;gap:12px;flex-wrap:wrap">
          <a class="cta" href="reports/$thaisLatestHtmlName">Abrir relatório da semana</a>
          <span class="badge">Atualizado: $(Get-Date -Format "yyyy-MM-dd HH:mm")</span>
        </div>
      </div>
      <div class="hero-images">
        <img src="$thaisHeroImage1" alt="Thais no triathlon">
        <img src="$thaisHeroImage2" alt="Thais na natacao">
      </div>
    </section>

    <section style="margin-top:18px">
      <h2>Calendário de Provas</h2>
      <div class="accent-line"></div>
      $thaisCalendarHtml
    </section>

    <section>
      <h2>Fase do Ciclo: $phaseTitle</h2>
      <div class="accent-line"></div>
      <div class="card">
        <div class="phase-grid">
          <div>
            <h3>Objetivo desta fase</h3>
            <ul>$phaseObjectiveItems</ul>
          </div>
          <div>
            <h3>Transição para Build</h3>
            <ul>$phaseTransitionItems</ul>
          </div>
        </div>
        <div class="timeline">
          <div class="step">Esta semana<br><strong>HOLD</strong></div>
          <div class="step">Semana +1<br><strong>Normal</strong></div>
          <div class="step">Semana +2<br><strong>Volume</strong></div>
          <div class="step">Semana +3<br><strong>Deload</strong></div>
        </div>
      </div>
    </section>

    <section>
      <h2>Relatórios Semanais</h2>
      <div class="accent-line"></div>
      <div class="grid grid-2">$($thaisReportsList -join "`n")</div>
    </section>

    <section>
      <div class="split">
        <div class="card">
          <h2>Resumo da Atleta</h2>
          <div class="accent-line"></div>
          <p>Foco principal em triathlon sprint, com prova de 10 km como objetivo secundário. Natação em fase técnica, bike em desenvolvimento de FTP (Sweet Spot) e corrida com base aeróbica sólida.</p>
        </div>
        <div class="card">
          <h2>Fase Atual</h2>
          <div class="accent-line"></div>
          <p><strong>$phaseTitle</strong></p>
          <p class="muted">$phaseFocus</p>
        </div>
      </div>
    </section>
  </div>
</body>
</html>
"@

  Set-Content -Path (Join-Path $thaisDir "index.html") -Value $thaisHtml -Encoding UTF8
  }

  $thaisRedirectDir = Join-Path $SiteDir "thais"
  $thaisRedirectPath = Join-Path $thaisRedirectDir "index.html"
  New-Item -ItemType Directory -Path $thaisRedirectDir -Force | Out-Null
  $thaisRedirectHtml = @"
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta http-equiv="refresh" content="0; url=../index.html">
  <title>Redirecionando...</title>
</head>
<body>
  <p>Redirecionando para a página principal...</p>
</body>
</html>
"@
  Set-Content -Path $thaisRedirectPath -Value $thaisRedirectHtml -Encoding UTF8

  $assetsDir = Join-Path $SiteDir "assets\\thais"
  New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null
  $assetsSrc = Join-Path $repoRoot "assets\\thais"
  if (Test-Path $assetsSrc) {
    Copy-Item -Path (Join-Path $assetsSrc "*") -Destination $assetsDir -Force -ErrorAction SilentlyContinue
  }
  $heroImage1 = ""
  $heroImage2 = ""
  if (Test-Path $assetsSrc) {
    $photos = Get-ChildItem -Path $assetsSrc -File | Where-Object { $_.Extension -match "\.jpe?g|\.png" } | Sort-Object Name
    if ($photos.Count -ge 1) { $heroImage1 = "assets/thais/$($photos[0].Name)" }
    if ($photos.Count -ge 2) { $heroImage2 = "assets/thais/$($photos[1].Name)" }
  }

  $indexHtml = @"
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Relatórios Intervals</title>
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Sora:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    :root{--bg:#0b1022;--card:#121a33;--card-2:#10172b;--text:#e5e7eb;--muted:#94a3b8;--accent:#ff6f91;--accent-2:#4bb6c1;--accent-3:#f6c453;--accent-4:#a855f7}
    *{box-sizing:border-box}
    body{margin:0;background:radial-gradient(circle at 15% 15%, rgba(255,111,145,0.25) 0%, #0b1022 50%, #090d1b 100%);color:var(--text);font-family:"Sora",sans-serif}
    .wrap{max-width:1200px;margin:28px auto;padding:0 22px}
    .hero{background:linear-gradient(135deg,#5f6ddf 0%,#8f6ccf 60%,#d17ca8 100%);border-radius:22px;padding:26px;display:grid;grid-template-columns:1.1fr 0.9fr;gap:16px;align-items:center}
    .hero h1{margin:0;font-size:28px}
    .hero p{margin:6px 0 0 0;color:#eef2ff}
    .pill{background:rgba(15,23,42,0.3);padding:8px 14px;border-radius:999px;font-weight:600}
    .btn{display:inline-flex;align-items:center;gap:8px;background:#0b1022;color:#e5e7eb;border:1px solid rgba(148,163,184,0.3);padding:10px 14px;border-radius:999px;font-weight:600}
    .btn:hover{border-color:rgba(148,163,184,0.6)}
    .grid{display:grid;gap:14px}
    .grid-3{grid-template-columns:repeat(auto-fit,minmax(220px,1fr))}
    .grid-2{grid-template-columns:repeat(auto-fit,minmax(260px,1fr))}
    .card{background:var(--card);border-radius:16px;padding:16px;border:1px solid rgba(148,163,184,0.12)}
    .metric{background:var(--card-2);border-radius:14px;padding:14px}
    .metric .value{font-size:22px;font-weight:700}
    .metric .label{font-size:12px;color:var(--muted)}
    h2{margin:22px 0 12px 0;font-size:18px}
    .event-card{display:flex;gap:12px;align-items:center;background:var(--card-2);border-radius:14px;padding:12px;border-left:4px solid var(--accent)}
    .event-card.tag-b{border-left-color:var(--accent-3)}
    .event-card.tag-c{border-left-color:#38bdf8}
    .event-days{width:70px;text-align:center;font-weight:700;font-size:20px}
    .event-days span{display:block;font-size:11px;color:var(--muted)}
    .event-info{flex:1}
    .event-name{font-weight:700}
    .event-meta,.event-date{font-size:12px;color:var(--muted)}
    .event-tag{font-size:12px;background:#1f2a4a;padding:6px 10px;border-radius:999px}
    .link-card{display:flex;justify-content:space-between;align-items:center;background:var(--card-2);padding:10px 12px;border-radius:12px}
    .chip{background:rgba(96,165,250,0.15);border:1px solid rgba(96,165,250,0.35);color:#dbeafe;padding:6px 10px;border-radius:999px;font-size:12px;cursor:pointer}
    .chip.active{background:#60a5fa;color:#0b1022}
    .archive-group{display:none}
    .archive-group.active{display:grid;gap:10px}
    a{color:#93c5fd;text-decoration:none}
    a:hover{text-decoration:underline}
    .muted{color:var(--muted)}
    .hero-media{display:grid;grid-template-columns:1fr 1fr;gap:14px}
    .hero-media img{width:100%;height:230px;object-fit:cover;border-radius:18px;border:1px solid rgba(148,163,184,0.2)}
    .hero-left{display:flex;flex-direction:column;gap:10px;min-height:230px}
    .hero-actions{margin-top:auto;display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-start}
    .calendar-phase{align-items:stretch}
    .calendar-phase > div{display:flex;flex-direction:column}
    .calendar-phase .phase-card{flex:1;display:flex;flex-direction:column;gap:12px}
    .calendar-phase .phase-card ul{margin:0}
    .phase-list{display:grid;gap:10px}
    .phase-item{display:flex;gap:12px;padding-bottom:10px;border-bottom:1px dashed rgba(148,163,184,0.18)}
    .phase-item:last-child{border-bottom:none;padding-bottom:0}
    .phase-range{min-width:96px;font-weight:700;color:#93c5fd;font-size:12px}
    .phase-lines{display:grid;gap:4px}
    .phase-line{font-size:12px;color:#e2e8f0;line-height:1.4}
    .phase-summary{margin-top:auto;padding-top:12px;border-top:1px solid rgba(148,163,184,0.14)}
    .phase-caption{font-size:11px;text-transform:uppercase;letter-spacing:.12em;color:var(--muted)}
    .phase-notes{margin:8px 0 0 0;padding-left:18px;color:var(--muted);font-size:12px;line-height:1.5}
    .phase-hint{margin:8px 0 0 0;color:var(--muted);font-size:12px}
    .phase-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin-top:10px}
    .phase-kpi{background:var(--card-2);border-radius:12px;padding:10px;border:1px solid rgba(148,163,184,0.12)}
    .phase-kpi .label{font-size:10px;text-transform:uppercase;letter-spacing:.12em;color:var(--muted)}
    .phase-kpi .value{font-size:13px;font-weight:700;margin-top:4px}
    .kpi-strip{margin-top:14px}
    .kpi-bar{background:var(--card-2);border-radius:18px;padding:16px;border:1px solid rgba(148,163,184,0.12)}
    .kpi-bar-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
    .kpi-bar-title{font-size:14px;font-weight:700}
    .kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px}
    .kpi-pill{background:rgba(15,23,42,0.45);border:1px solid rgba(148,163,184,0.2);border-radius:12px;padding:10px 12px;display:flex;justify-content:space-between;align-items:center}
    .kpi-pill .label{font-size:10px;text-transform:uppercase;letter-spacing:.12em;color:var(--muted)}
    .kpi-pill .value{font-size:14px;font-weight:700}
    @media (max-width: 900px){
      .hero{grid-template-columns:1fr}
    }
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <div class="hero-left">
        <h1>Relatório de Coaching</h1>
        <p>$athleteName · Temporada 2026</p>
        <div class="hero-actions">
          <a class="btn" href="reports/$latestHtmlName">Abrir relatório da semana</a>
          <div class="pill">Atualizado: $(Get-Date -Format "yyyy-MM-dd HH:mm")</div>
        </div>
      </div>
      <div>
        <div class="hero-media">
          $(if ($heroImage1) { "<img src=""$heroImage1"" alt=""Thais em prova"">" } else { "" })
          $(if ($heroImage2) { "<img src=""$heroImage2"" alt=""Thais na natacao"">" } else { "" })
        </div>
      </div>
    </section>

    <section class="grid grid-3" style="margin-top:16px">
      <div class="card">
        <strong>Perfil do Atleta</strong>
        <div class="muted" style="margin-top:8px">$athleteAge anos · $athleteHeight · $athleteWeightDisplay</div>
        <div class="muted">$athleteExp · $athleteLevel</div>
      </div>
      <div class="card">
        <strong>Fase Atual</strong>
        <div style="margin-top:8px">$phaseTitle</div>
        <div class="muted">$phaseFocus</div>
      </div>
      <div class="card">
        <strong>Métricas (última semana)</strong>
        <div class="muted" style="margin-top:8px">CTL $ctlText · ATL $atlText · TSB $tsbText</div>
        <div class="muted">Aderência ao plano: $complianceText</div>
      </div>
    </section>

    <section class="kpi-strip">
      <div class="kpi-bar">
        <div class="kpi-bar-header">
          <div class="kpi-bar-title">KPIs Fisiológicos</div>
          <div class="muted">referência atual</div>
        </div>
        <div class="kpi-row">
          <div class="kpi-pill">
            <div class="label">VO2 Corrida</div>
            <div class="value">$vo2RunText</div>
          </div>
          <div class="kpi-pill">
            <div class="label">VO2 Ciclismo</div>
            <div class="value">$vo2BikeText</div>
          </div>
          <div class="kpi-pill">
            <div class="label">FTP Bike</div>
            <div class="value">$bikeFtpText</div>
          </div>
          <div class="kpi-pill">
            <div class="label">Limiar Corrida</div>
            <div class="value">$runLactateText</div>
          </div>
        </div>
      </div>
    </section>

    <section>
      <div class="grid grid-2 calendar-phase">
        <div>
          <h2>Calendário de Provas</h2>
          <div class="grid">$calendarHtml</div>
        </div>
        <div>
          <h2>Fases do Ano</h2>
          <div class="card phase-card">
            <div class="phase-list">$futurePhaseItems</div>
            <div class="phase-summary">
              <div class="phase-caption">Diretrizes do ciclo</div>
              $phaseNotesHtml
              <div class="phase-kpis">
                <div class="phase-kpi">
                  <div class="label">Volume alvo</div>
                  <div class="value">$phaseVolumeTarget</div>
                </div>
                <div class="phase-kpi">
                  <div class="label">Intensidade</div>
                  <div class="value">$phaseIntensity</div>
                </div>
                <div class="phase-kpi">
                  <div class="label">Força</div>
                  <div class="value">$phaseStrength</div>
                </div>
              </div>
              <p class="phase-hint">Revisões semanais consideram sono, HRV e carga (ATL/TSB).</p>
            </div>
          </div>
        </div>
      </div>
    </section>

    <section>
      <h2>Relatórios Semanais</h2>
      <div class="grid grid-2">$($reportsList -join "`n")</div>
    </section>

    <section>
      <h2>Relatórios Anteriores</h2>
      <div class="grid" style="margin-bottom:12px">$($archiveButtons -join "`n")</div>
      <div class="grid grid-2">$($archiveGroups -join "`n")</div>
    </section>
  </div>
  <script>
    const buttons = document.querySelectorAll('[data-month-btn]');
    const groups = document.querySelectorAll('[data-month-group]');
    function setMonth(month){
      buttons.forEach(b=>b.classList.toggle('active', b.dataset.monthBtn===month));
      groups.forEach(g=>g.classList.toggle('active', g.dataset.monthGroup===month));
    }
    if (buttons.length) {
      setMonth(buttons[0].dataset.monthBtn);
      buttons.forEach(b=>b.addEventListener('click',()=>setMonth(b.dataset.monthBtn)));
    }
  </script>
</body>
</html>
"@

Set-Content -Path (Join-Path $SiteDir "index.html") -Value $indexHtml -Encoding UTF8
Set-Content -Path (Join-Path $SiteDir ".nojekyll") -Value "" -Encoding ASCII

