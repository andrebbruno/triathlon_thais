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
  if ($Text -match "Ã|Â|â") {
    try {
      $fixed = [Text.Encoding]::UTF8.GetString([Text.Encoding]::GetEncoding("Windows-1252").GetBytes($Text))
      if ($fixed -match "Ã|Â|â") {
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
  return (Clean-DisplayText -Text $Text)
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

# Planejado vs executado
$matches = @()
foreach ($p in $planned) {
  $p.name = Clean-DisplayText -Text $p.name
  $p.description = Fix-TextEncoding -Text $p.description
  $p.type = Map-Type -Type $p.type
  $best = $null; $bestScore = -1
  foreach ($a in $activities) {
    $score = Score-Match -Planned $p -Activity $a
    if ($score -gt $bestScore) { $bestScore = $score; $best = $a }
  }
  $matched = $bestScore -ge 3
  $matches += [PSCustomObject]@{
    planned = $p
    activity = if ($matched) { $best } else { $null }
    score = $bestScore
    matched = $matched
  }
}

$plannedCount = $planned.Count
$matchedCount = ($matches | Where-Object { $_.matched }).Count
$compliance = if ($plannedCount -gt 0) { [math]::Round(($matchedCount / $plannedCount) * 100, 1) } else { $null }

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
} elseif ($compliance -ne $null -and $compliance -ge 85 -and $tsb -ge -10 -and $tsb -le 10) {
  $status = "PUSH"
}

# Gerar markdown de analise
$analysisPath = Join-Path $OutputDir ("analysis_{0}_{1}.md" -f $weekStart, $weekEnd)
$lines = @()
$lines += "# Analise Semanal ($weekStart a $weekEnd)"
$lines += ""
$lines += "## Status da Semana"
$lines += "- Classificacao: **$status**"
$lines += "- Compliance: " + ($(if ($compliance -ne $null) { "$compliance%" } else { "n/a" }))
$lines += "- CTL/ATL/TSB: $($report.metricas.CTL) / $($report.metricas.ATL) / $($report.metricas.TSB)"
$lines += ""
$lines += "## Bem-estar (media)"
$lines += "- Sono: " + ($(if ($avgSleep -ne $null) { "$avgSleep h" } else { "n/a" }))
$lines += "- HRV: " + ($(if ($avgHrv -ne $null) { "$avgHrv" } else { "n/a" }))
$lines += "- FC repouso: " + ($(if ($avgRhr -ne $null) { "$avgRhr bpm" } else { "n/a" }))
$lines += ""
$lines += "## Planejado vs Executado"
foreach ($m in $matches) {
  $p = $m.planned
  $a = $m.activity
  if ($m.matched -and $a) {
    $pName = Display-Name -Text $p.name
    $aName = Display-Name -Text $a.name
    $pTime = if ($p.moving_time_min) { "$($p.moving_time_min)min" } else { "n/a" }
    $aTime = if ($a.moving_time_min) { "$([math]::Round($a.moving_time_min,1))min" } else { "n/a" }
    $pDist = if ($p.distance_km) { "$($p.distance_km)km" } else { "n/a" }
    $aDist = if ($a.distance_km) { "$([math]::Round($a.distance_km,1))km" } else { "n/a" }
    $lines += "- **$($p.start_date) - $pName** -> Executado: $aName | Tempo $pTime vs $aTime | Dist $pDist vs $aDist"
  } else {
    $pName = Display-Name -Text $p.name
    $lines += "- **$($p.start_date) - $pName** -> **Nao executado**"
  }
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
$trainings | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $repoRoot $TrainingsOut) -Encoding utf8

Write-Host "Analise salva em: $analysisPath"
Write-Host "Trainings gerado em: $trainingsPath"
Write-Host "Trainings (root) atualizado: $(Join-Path $repoRoot $TrainingsOut)"
