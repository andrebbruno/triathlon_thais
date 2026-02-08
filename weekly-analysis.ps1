# weekly-analysis.ps1
# Analise semanal (planejado vs executado) + gera trainings.json para a proxima semana

param(
  [string]$ReportPath = "",
  [string]$OutputDir = "",
  [string]$TrainingsOut = "trainings.json",
  [int]$WeekShiftDays = 7
)

$repoRoot = $PSScriptRoot
if (-not $OutputDir) { $OutputDir = Join-Path $repoRoot "Relatorios_Intervals" }
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

function Read-Json {
  param([string]$Path)
  if (-not (Test-Path $Path)) { return $null }
  $bytes = [System.IO.File]::ReadAllBytes($Path)
  $text = [Text.Encoding]::UTF8.GetString($bytes)
  return ($text | ConvertFrom-Json)
}

function Get-LatestReport {
  param([string]$Dir)
  $file = Get-ChildItem $Dir -Filter "report_*.json" | Sort-Object Name -Descending | Select-Object -First 1
  if ($file) { return $file.FullName }
  return ""
}

function Parse-NumberFromText {
  param([string]$Value, [double]$Default)
  if (-not $Value) { return $Default }
  $clean = ($Value -replace "[^\d\.,]", "").Replace(",", ".")
  $num = 0.0
  if ([double]::TryParse($clean, [ref]$num)) { return $num }
  return $Default
}

function Normalize-Text {
  param([string]$Text)
  if (-not $Text) { return "" }
  $t = $Text.ToLowerInvariant()
  $t = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($t))
  $t = ($t -replace "[^a-z0-9\s-]", "").Trim()
  return $t
}

function Fix-TextEncoding {
  param([string]$Text)
  if (-not $Text) { return "" }
  # Keep this script ASCII-only (no BOM) for Windows PowerShell 5.1 compatibility.
  # Heuristic: typical mojibake includes U+00C3 / U+00C2 characters.
  $ch1 = [char]0x00C3
  $ch2 = [char]0x00C2
  if ($Text.IndexOf($ch1) -ge 0 -or $Text.IndexOf($ch2) -ge 0) {
    try {
      $fixed = [Text.Encoding]::UTF8.GetString([Text.Encoding]::GetEncoding("Windows-1252").GetBytes($Text))
      if ($fixed.IndexOf($ch1) -ge 0 -or $fixed.IndexOf($ch2) -ge 0) {
        $fixed = [Text.Encoding]::UTF8.GetString([Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($Text))
      }
      return $fixed
    } catch { return $Text }
  }
  return $Text
}

function To-Ascii {
  param([string]$Text)
  if (-not $Text) { return "" }
  $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $normalized.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$sb.Append($ch)
    }
  }
  return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Clean-DisplayText {
  param([string]$Text)
  if (-not $Text) { return "" }
  $fixed = Fix-TextEncoding -Text $Text
  $ascii = To-Ascii -Text $fixed
  $ascii = ($ascii -replace "[^A-Za-z0-9\s\-\(\)\/]", "")
  $ascii = ($ascii -replace "\s+", " ").Trim()
  return $ascii
}

function Display-Name {
  param([string]$Text)
  $fixed = Fix-TextEncoding -Text $Text
  return (($fixed -replace "\s+", " ").Trim())
}

function Score-Match {
  param([object]$Planned, [object]$Activity)
  $score = 0
  if ($Planned.type -and $Activity.type -and (Normalize-Text $Planned.type) -eq (Normalize-Text $Activity.type)) { $score += 2 }
  $pDate = $Planned.start_date
  $aDate = $Activity.start_date_local
  if ($pDate -and $aDate -and $pDate -eq $aDate) { $score += 2 }
  $pName = Normalize-Text $Planned.name
  $aName = Normalize-Text $Activity.name
  if ($pName -and $aName) {
    if ($aName -like "*$pName*") { $score += 2 }
    elseif ($pName -like "*$aName*") { $score += 1 }
    else {
      $first = ($pName -split "\s+")[0]
      if ($first -and $aName -like "*$first*") { $score += 1 }
    }
  }
  return $score
}

function Map-Type {
  param([string]$Type)
  if (-not $Type) { return "" }
  if ($Type -eq "Workout") { return "WeightTraining" }
  return $Type
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

function Format-Percent {
  param([double]$Value)
  if ($Value -eq $null) { return "n/a" }
  return "{0:0.0}%" -f $Value
}

function Shift-ExternalId {
  param([string]$ExternalId, [DateTime]$NewDate)
  if (-not $ExternalId) { return "" }
  $dateIso = $NewDate.ToString("yyyy-MM-dd")
  $dateCompact = $NewDate.ToString("yyyyMMdd")
  if ($ExternalId -match "^\d{4}-\d{2}-\d{2}") {
    return $ExternalId -replace "^\d{4}-\d{2}-\d{2}", $dateIso
  }
  if ($ExternalId -match "training_\d{8}") {
    return $ExternalId -replace "training_\d{8}", "training_$dateCompact"
  }
  return "$ExternalId-$dateIso"
}

if (-not $ReportPath) { $ReportPath = Get-LatestReport -Dir $OutputDir }
if (-not $ReportPath) { Write-Host "Nenhum report encontrado em $OutputDir"; exit 1 }

$report = Read-Json -Path $ReportPath
if (-not $report) { Write-Host "Falha ao ler report: $ReportPath"; exit 1 }

$memoryPath = Join-Path $repoRoot "COACHING_MEMORY.md"
$memoryText = if (Test-Path $memoryPath) { Get-Content $memoryPath -Raw } else { "" }
$baselineRhr = if ($memoryText -match "\*\*FC Repouso baseline:\*\*\s*~?([0-9,\.]+)") { Parse-NumberFromText $matches[1] 48 } else { 48 }
$baselineHrv = if ($memoryText -match "\*\*HRV baseline:\*\*\s*~?([0-9,\.]+)") { Parse-NumberFromText $matches[1] 45 } else { 45 }
$idealSleep = if ($memoryText -match "\*\*Sono ideal:\*\*\s*([0-9,\.]+)") { Parse-NumberFromText $matches[1] 7.5 } else { 7.5 }

$activities = @($report.atividades)
$planned = @($report.treinos_planejados)
$weekStart = $report.semana.inicio
$weekEnd = $report.semana.fim

# Planejado vs executado (1:1, evitando falso-positivo)
$plannedOff = @($planned | Where-Object { Is-OffPlannedEvent -Event $_ })
$plannedWorkouts = @($planned | Where-Object { -not (Is-OffPlannedEvent -Event $_) })

$activityIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($a in $activities) {
  if ($a.id) { [void]$activityIds.Add([string]$a.id) }
}

$matchedEventIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($a in $activities) {
  if ($a.planejado -and $a.planejado.event_id) {
    [void]$matchedEventIds.Add([string]$a.planejado.event_id)
  }
}

$doneWorkouts = @()
$missedWorkouts = @()
foreach ($p in $plannedWorkouts) {
  $p.type = Map-Type -Type $p.type
  $eventId = [string]$p.event_id
  $pairedId = [string]$p.paired_activity_id
  $matched = $false
  if ($pairedId -and $activityIds.Contains($pairedId)) { $matched = $true }
  elseif ($eventId -and $matchedEventIds.Contains($eventId)) { $matched = $true }

  if ($matched) { $doneWorkouts += $p } else { $missedWorkouts += $p }
}

$offRespected = 0
$offBroken = 0
foreach ($p in $plannedOff) {
  $d = [string]$p.start_date
  $hasActivity = $false
  if ($d) {
    $hasActivity = (@($activities | Where-Object { $_.start_date_local -eq $d } | Select-Object -First 1) -ne $null)
  }
  if (-not $hasActivity) { $offRespected += 1 } else { $offBroken += 1 }
}

$extraActivities = @($activities | Where-Object { $_.planejado -eq $null })
$extraCount = $extraActivities.Count

$plannedCount = $planned.Count
$plannedWorkoutCount = $plannedWorkouts.Count
$plannedOffCount = $plannedOff.Count
$doneWorkoutCount = $doneWorkouts.Count

$workoutCompliance = if ($plannedWorkoutCount -gt 0) { [math]::Round(($doneWorkoutCount / $plannedWorkoutCount) * 100, 1) } else { $null }
$overallAdherence = if ($plannedCount -gt 0) { [math]::Round((($doneWorkoutCount + $offRespected) / $plannedCount) * 100, 1) } else { $null }

# Bem-estar
$wellness = @($report.bem_estar)
$avgSleep = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object sono_h -Average).Average), 2) } else { $null }
$avgHrv = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object hrv -Average).Average), 1) } else { $null }
$avgRhr = if ($wellness.Count -gt 0) { [math]::Round((($wellness | Measure-Object fc_reposo -Average).Average), 1) } else { $null }

# Classificacao
$tsb = $report.metricas.TSB
$status = "HOLD"
if ($tsb -le -20 -or ($avgSleep -ne $null -and $avgSleep -lt ($idealSleep - 1)) -or ($avgHrv -ne $null -and $avgHrv -lt ($baselineHrv - 5)) -or ($avgRhr -ne $null -and $avgRhr -gt ($baselineRhr + 5))) {
  $status = "STEP BACK"
} elseif ($workoutCompliance -ne $null -and $workoutCompliance -ge 85 -and $tsb -ge -10 -and $tsb -le 10) {
  $status = "PUSH"
}

# Gerar markdown de analise
$analysisPath = Join-Path $OutputDir ("analysis_{0}_{1}.md" -f $weekStart, $weekEnd)
$lines = @()
$lines += "# Analise Semanal ($weekStart a $weekEnd)"
$lines += ""
$lines += "## Status da Semana"
$lines += "- Classificacao: **$status**"
$lines += "- Aderencia ao plano (geral): " + ($(if ($overallAdherence -ne $null) { (Format-Percent -Value $overallAdherence) } else { "n/a" })) + " | Treinos $doneWorkoutCount/$plannedWorkoutCount | Descanso $offRespected/$plannedOffCount | Extras $extraCount"
$lines += "- Aderencia (treinos): " + ($(if ($workoutCompliance -ne $null) { (Format-Percent -Value $workoutCompliance) } else { "n/a" }))
$lines += "- CTL/ATL/TSB: $($report.metricas.CTL) / $($report.metricas.ATL) / $($report.metricas.TSB)"
$lines += ""
$lines += "## Bem-estar (media)"
$lines += "- Sono: " + ($(if ($avgSleep -ne $null) { "$avgSleep h" } else { "n/a" }))
$lines += "- HRV: " + ($(if ($avgHrv -ne $null) { "$avgHrv" } else { "n/a" }))
$lines += "- FC repouso: " + ($(if ($avgRhr -ne $null) { "$avgRhr bpm" } else { "n/a" }))
$lines += ""
$missedSummary = if ($missedWorkouts.Count -gt 0) {
  ($missedWorkouts | ForEach-Object { "{0} ({1})" -f (Display-Name -Text $_.name), $_.start_date }) -join "; "
} else { "nenhum" }
$lines += "## Pendencias do plano"
$lines += "- Nao feitos: $missedSummary"
$extraSummary = if ($extraActivities.Count -gt 0) {
  ($extraActivities | Select-Object -First 4 | ForEach-Object { Display-Name -Text $_.name }) -join ", "
} else { "" }
if ($extraCount -gt 0) {
  $suffix = if ($extraActivities.Count -gt 4) { ", ..." } else { "" }
  $lines += "- Extras (fora do plano): $extraCount ($extraSummary$suffix)"
} else {
  $lines += "- Extras (fora do plano): 0"
}

($lines -join "`n") | Out-File -FilePath $analysisPath -Encoding utf8

# Gerar trainings.json para proxima semana (shift simples)
$nextStart = ([DateTime]::Parse($weekStart)).AddDays($WeekShiftDays)
$nextEnd = ([DateTime]::Parse($weekEnd)).AddDays($WeekShiftDays)
$trainings = @()
foreach ($p in $planned) {
  $start = [DateTime]::Parse($p.start_date_local)
  $newStart = $start.AddDays($WeekShiftDays)
  $newDate = $newStart.ToString("yyyy-MM-dd")
  $type = Map-Type -Type $p.type
    $desc = if ($p.description -and $p.description.Trim() -ne "") { $p.description } else { "Sessao planejada" }
    $trainings += [ordered]@{
      external_id = (Shift-ExternalId -ExternalId $p.external_id -NewDate $newStart)
      category = "WORKOUT"
      start_date_local = $newStart.ToString("yyyy-MM-ddTHH:mm:ss")
      type = $type
      name = $p.name
      description = $desc
    }
}

$trainingsPath = Join-Path $OutputDir ("trainings_{0}_{1}.json" -f $nextStart.ToString("yyyy-MM-dd"), $nextEnd.ToString("yyyy-MM-dd"))
$trainings | ConvertTo-Json -Depth 6 | Out-File -FilePath $trainingsPath -Encoding utf8
$trainingsRootPath = if ([System.IO.Path]::IsPathRooted($TrainingsOut)) { $TrainingsOut } else { (Join-Path $repoRoot $TrainingsOut) }
$trainings | ConvertTo-Json -Depth 6 | Out-File -FilePath $trainingsRootPath -Encoding utf8

Write-Host "Analise salva em: $analysisPath"
Write-Host "Trainings gerado em: $trainingsPath"
Write-Host "Trainings (root) atualizado: $trainingsRootPath"
