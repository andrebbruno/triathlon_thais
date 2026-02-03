# =============================================================================
# ANALISE DE LONGO PRAZO - INTERVALS.ICU - COACH EDITION
# =============================================================================
# Versao: 4.0 - Janeiro 2026
# Analise holistica avancada com metricas de triathlon
# =============================================================================

param(
    [int]$WeeksBack = 12,
    [string]$EventDate = "",
    [string]$AthleteId = "0",
    [string]$ApiKeyPath = "$PSScriptRoot\\api_key.txt",
    [string]$OutputDir = "",
    [switch]$EnableStreams
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# =============================================================================
# FUNCOES AUXILIARES
# =============================================================================

function Get-ApiKey {
    param ([string]$Path)

    if ($env:INTERVALS_API_KEY) { return $env:INTERVALS_API_KEY }
    if (Test-Path $Path) {
        return (Get-Content $Path -Raw).Trim()
    }

    $localPath = $null
    if ($env:USERPROFILE) { $localPath = Join-Path $env:USERPROFILE ".intervals\\api_key.txt" }
    elseif ($env:HOME) { $localPath = Join-Path $env:HOME ".intervals\\api_key.txt" }
    if ($localPath -and (Test-Path $localPath)) {
        return (Get-Content $localPath -Raw).Trim()
    }

    $secureKey = Read-Host "Enter your Intervals.icu API Key" -AsSecureString
    return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
    )
}

# =============================================================================
# CONFIGURACOES
# =============================================================================

$config = @{
    athlete_id = $AthleteId
    base_url   = "https://intervals.icu/api/v1"
}

$apiKey = Get-ApiKey -Path $ApiKeyPath
if (-not $apiKey) {
    Write-Host "API key not provided." -ForegroundColor Red
    exit 1
}

if ($config.athlete_id -eq "0" -and $env:INTERVALS_ATHLETE_ID) {
    $config.athlete_id = $env:INTERVALS_ATHLETE_ID
}

if ($config.athlete_id -eq "0") {
    $memoryPath = Join-Path -Path $PSScriptRoot -ChildPath "COACHING_MEMORY.md"
    if (Test-Path $memoryPath) {
        $memoryText = Get-Content -Raw $memoryPath
        if ($memoryText -match "Athlete ID:\s*([A-Za-z0-9_-]+)") {
            $config.athlete_id = $matches[1]
        }
    }
}

if (-not $EventDate) {
    $memoryPath = Join-Path -Path $PSScriptRoot -ChildPath "COACHING_MEMORY.md"
    if (Test-Path $memoryPath) {
        $lines = Get-Content $memoryPath
        $startIdx = ($lines | Select-String -Pattern "^\| Data \| Prova" | Select-Object -First 1).LineNumber
        if ($startIdx) {
            for ($i = $startIdx; $i -lt $lines.Count; $i++) {
                $line = $lines[$i].Trim()
                if (-not $line.StartsWith("|")) { break }
                if ($line -match "^\|\s*-") { continue }
                $cols = $line.Trim("|") -split "\|"
                if ($cols.Count -lt 5) { continue }
                $dateText = $cols[0].Trim()
                $priority = ($cols[3].Trim() -replace "\*","")
                if ($priority -eq "A") {
                    $parts = $dateText -split "/"
                    if ($parts.Count -eq 2) {
                        $year = (Get-Date).Year
                        $dt = Get-Date -Year $year -Month $parts[1] -Day $parts[0]
                        if ($dt -lt (Get-Date).AddDays(-1)) { $dt = $dt.AddYears(1) }
                        $EventDate = $dt.ToString("yyyy-MM-dd")
                    }
                    break
                }
            }
        }
    }
}

if (-not $OutputDir) {
    $OutputDir = Join-Path -Path $PSScriptRoot -ChildPath "Relatorios_Intervals"
}
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$useStreams = if ($PSBoundParameters.ContainsKey("EnableStreams")) { [bool]$EnableStreams } else { $true }

function Get-AuthHeader {
    $pair = "API_KEY:$apiKey"
    $base64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))
    return @{ "Authorization" = "Basic $base64"; "Accept" = "application/json" }
}

function Invoke-IntervalsAPI {
    param([string]$Endpoint, [hashtable]$QueryParams = @{})
    $headers = Get-AuthHeader
    $uri = "$($config.base_url)$Endpoint"
    if ($QueryParams.Count -gt 0) {
        $queryString = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
        $uri = "$uri`?$queryString"
    }
    try {
        # Usar WebRequest para controlar encoding UTF-8 corretamente
        $response = Invoke-WebRequest -Uri $uri -Headers $headers -Method Get -UseBasicParsing -ErrorAction Stop
        # Decodificar como UTF-8 usando RawContentStream ou Content
        if ($response.RawContentStream) {
            $bytes = New-Object byte[] $response.RawContentStream.Length
            $response.RawContentStream.Position = 0
            $response.RawContentStream.Read($bytes, 0, $bytes.Length) | Out-Null
            $utf8Content = [System.Text.Encoding]::UTF8.GetString($bytes)
        } else {
            # Fallback: re-encode o Content como UTF-8
            $bytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($response.Content)
            $utf8Content = [System.Text.Encoding]::UTF8.GetString($bytes)
        }
        Start-Sleep -Milliseconds 100
        return $utf8Content | ConvertFrom-Json
    }
    catch { Write-Host "  ERRO: $($_.Exception.Message)" -ForegroundColor Red; return $null }
}

function Convert-SecondsToMinutes { param([double]$Seconds); if ($null -eq $Seconds -or $Seconds -eq 0) { return 0 }; return [math]::Round($Seconds / 60, 1) }
function Convert-MetersToKm { param([double]$Meters); if ($null -eq $Meters -or $Meters -eq 0) { return 0 }; return [math]::Round($Meters / 1000, 2) }
function Convert-SpeedToKmh { param([double]$MetersPerSecond); if ($null -eq $MetersPerSecond -or $MetersPerSecond -eq 0) { return $null }; return [math]::Round($MetersPerSecond * 3.6, 1) }

function Convert-PaceToMinKm {
    param([double]$MetersPerSecond)
    if ($null -eq $MetersPerSecond -or $MetersPerSecond -eq 0) { return $null }
    $kmPerHour = $MetersPerSecond * 3.6
    $minPerKm = 60 / $kmPerHour
    $minutes = [int][math]::Floor($minPerKm)
    $seconds = [int][math]::Round(($minPerKm - $minutes) * 60)
    return "{0:D2}:{1:D2}/km" -f $minutes, $seconds
}

function Get-FieldValue {
    param([object]$Object, [string[]]$FieldNames, [object]$Default = $null)
    foreach ($field in $FieldNames) {
        $value = $Object.$field
        if ($null -ne $value -and $value -ne 0 -and $value -ne "") { return $value }
    }
    return $Default
}

function Sum-SeilerZones {
    param(
        [array]$Activities,
        [string]$PropertyName
    )

    if (-not $Activities -or $Activities.Count -eq 0) { return $null }

    $z1 = 0.0; $z2 = 0.0; $z3 = 0.0; $has = $false
    foreach ($act in $Activities) {
        $zones = $act.$PropertyName
        if ($zones) {
            if ($zones.Seiler_Z1_Easy_min -ne $null) { $z1 += [double]$zones.Seiler_Z1_Easy_min; $has = $true }
            if ($zones.Seiler_Z2_Moderate_min -ne $null) { $z2 += [double]$zones.Seiler_Z2_Moderate_min; $has = $true }
            if ($zones.Seiler_Z3_Hard_min -ne $null) { $z3 += [double]$zones.Seiler_Z3_Hard_min; $has = $true }
        }
    }

    if (-not $has) { return $null }

    $total = $z1 + $z2 + $z3
    $result = @{
        z1_min = [math]::Round($z1, 1)
        z2_min = [math]::Round($z2, 1)
        z3_min = [math]::Round($z3, 1)
        z1_pct = if ($total -gt 0) { [math]::Round(($z1 / $total) * 100, 1) } else { $null }
        z2_pct = if ($total -gt 0) { [math]::Round(($z2 / $total) * 100, 1) } else { $null }
        z3_pct = if ($total -gt 0) { [math]::Round(($z3 / $total) * 100, 1) } else { $null }
    }
    return $result
}

function Get-ZoneTimesFormatted {
    param([array]$ZoneTimes, [string]$Type)
    if ($null -eq $ZoneTimes -or $ZoneTimes.Count -eq 0) { return $null }

    # Converter para array de segundos (lidar com formato objeto ou numero)
    $secondsArray = @()
    foreach ($zone in $ZoneTimes) {
        if ($zone -is [PSCustomObject] -or $zone -is [hashtable]) {
            $secs = if ($zone.secs) { $zone.secs } elseif ($zone.seconds) { $zone.seconds } else { 0 }
            $secondsArray += [int]$secs
        } else {
            $secondsArray += [int]$zone
        }
    }

    $zoneNames = switch ($Type) {
        "power" { @("Z1_Recovery", "Z2_Endurance", "Z3_Tempo", "Z4_Threshold", "Z5_VO2Max", "Z6_Anaerobic", "Z7_Neuromuscular") }
        "hr"    { @("Z1_Recovery", "Z2_Aerobic", "Z3_Tempo", "Z4_SubThreshold", "Z5_Threshold", "Z6_Anaerobic", "Z7_Max") }
        "pace"  { @("Z1_Recovery", "Z2_Aerobic", "Z3_Tempo", "Z4_SubThreshold", "Z5_Threshold", "Z6_VO2Max", "Z7_Fast") }
        default { @("Z1", "Z2", "Z3", "Z4", "Z5", "Z6", "Z7") }
    }
    $result = @{}; $totalSeconds = 0
    for ($i = 0; $i -lt [Math]::Min($secondsArray.Count, $zoneNames.Count); $i++) {
        $seconds = $secondsArray[$i]; if ($seconds -gt 0) { $totalSeconds += $seconds }
        $result[$zoneNames[$i] + "_sec"] = $seconds; $result[$zoneNames[$i] + "_min"] = [math]::Round($seconds / 60, 1)
    }
    if ($totalSeconds -gt 0) {
        for ($i = 0; $i -lt [Math]::Min($secondsArray.Count, $zoneNames.Count); $i++) {
            $result[$zoneNames[$i] + "_pct"] = [math]::Round(($secondsArray[$i] / $totalSeconds) * 100, 1)
        }
    }
    if ($secondsArray.Count -ge 5) {
        $seilerZ1 = $secondsArray[0] + $secondsArray[1]; $seilerZ2 = $secondsArray[2] + $secondsArray[3]; $seilerZ3 = 0
        for ($i = 4; $i -lt $secondsArray.Count; $i++) { $seilerZ3 += $secondsArray[$i] }
        $result["Seiler_Z1_Easy_min"] = [math]::Round($seilerZ1 / 60, 1)
        $result["Seiler_Z2_Moderate_min"] = [math]::Round($seilerZ2 / 60, 1)
        $result["Seiler_Z3_Hard_min"] = [math]::Round($seilerZ3 / 60, 1)
        if ($totalSeconds -gt 0) {
            $result["Seiler_Z1_Easy_pct"] = [math]::Round(($seilerZ1 / $totalSeconds) * 100, 1)
            $result["Seiler_Z2_Moderate_pct"] = [math]::Round(($seilerZ2 / $totalSeconds) * 100, 1)
            $result["Seiler_Z3_Hard_pct"] = [math]::Round(($seilerZ3 / $totalSeconds) * 100, 1)
        }
    }
    return $result
}

function Get-StreamData {
    param([object]$Streams, [string[]]$Keys)
    if (-not $Streams) { return $null }

    if ($Streams -is [System.Collections.IEnumerable] -and -not ($Streams -is [string])) {
        foreach ($key in $Keys) {
            $item = $Streams | Where-Object { $_.type -eq $key -or $_.key -eq $key } | Select-Object -First 1
            if ($item) {
                if ($item.data) { return ,$item.data }
                if ($item.values) { return ,$item.values }
                if ($item.value) { return ,$item.value }
            }
        }
    }

    foreach ($key in $Keys) {
        if ($Streams.PSObject.Properties.Name -contains $key) {
            $val = $Streams.$key
            if ($val.data) { return ,$val.data }
            if ($val.values) { return ,$val.values }
            return ,$val
        }
    }

    return $null
}

function Get-SampleInterval {
    param([double[]]$Time)
    if (-not $Time -or $Time.Count -lt 2) { return 1 }
    $diffs = @()
    for ($i = 1; $i -lt $Time.Count; $i++) {
        $delta = $Time[$i] - $Time[$i - 1]
        if ($delta -gt 0) { $diffs += $delta }
    }
    if ($diffs.Count -eq 0) { return 1 }
    return [math]::Max(1, [math]::Round(($diffs | Measure-Object -Average).Average))
}

function Get-NormalizedPowerFromStream {
    param([double[]]$Power, [double[]]$Time)
    if (-not $Power -or $Power.Count -lt 30) { return $null }
    $sampleInterval = Get-SampleInterval -Time $Time
    $window = [math]::Max(1, [math]::Round(30 / $sampleInterval))
    if ($Power.Count -lt $window) { return $null }

    $sum = 0.0; $count = 0
    $rollingSum = 0.0
    for ($i = 0; $i -lt $Power.Count; $i++) {
        $rollingSum += [double]$Power[$i]
        if ($i -ge $window) { $rollingSum -= [double]$Power[$i - $window] }
        if ($i -ge ($window - 1)) {
            $avg = $rollingSum / $window
            $sum += [math]::Pow($avg, 4)
            $count++
        }
    }
    if ($count -eq 0) { return $null }
    return [math]::Round([math]::Pow(($sum / $count), 0.25), 0)
}

function Get-ZoneTimesFromStream {
    param([double[]]$Values, [double]$FTP, [double[]]$Time)
    if (-not $Values -or $Values.Count -eq 0 -or -not $FTP -or $FTP -le 0) { return $null }
    $dt = @()
    if ($Time -and $Time.Count -eq $Values.Count) {
        for ($i = 0; $i -lt $Values.Count; $i++) {
            if ($i -eq 0) { $dt += 1 } else { $dt += [math]::Max(1, [int]([double]$Time[$i] - [double]$Time[$i - 1])) }
        }
    } else {
        $dt = 1..$Values.Count | ForEach-Object { 1 }
    }

    $zones = @(0,0,0,0,0,0,0)
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $v = [double]$Values[$i]
        if ($v -le 0) { continue }
        $pct = ($v / $FTP) * 100
        $idx = if ($pct -lt 55) { 0 }
               elseif ($pct -lt 75) { 1 }
               elseif ($pct -lt 90) { 2 }
               elseif ($pct -lt 105) { 3 }
               elseif ($pct -lt 120) { 4 }
               elseif ($pct -lt 150) { 5 }
               else { 6 }
        $zones[$idx] += $dt[$i]
    }
    return $zones
}

function Get-DecouplingFromStreams {
    param([double[]]$Power, [double[]]$Hr)
    if (-not $Power -or -not $Hr) { return $null }
    $len = [math]::Min($Power.Count, $Hr.Count)
    if ($len -lt 60) { return $null }
    $half = [math]::Floor($len / 2)
    $p1 = ($Power[0..($half-1)] | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average
    $p2 = ($Power[$half..($len-1)] | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average
    $h1 = ($Hr[0..($half-1)] | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average
    $h2 = ($Hr[$half..($len-1)] | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average
    if (-not $p1 -or -not $p2 -or -not $h1 -or -not $h2) { return $null }
    return [math]::Round((($h2 / $p2) / ($h1 / $p1) - 1) * 100, 2)
}

function Calculate-Monotony {
    param([array]$DailyTSS)
    if ($null -eq $DailyTSS -or $DailyTSS.Count -lt 2) { return $null }
    $validTSS = $DailyTSS | Where-Object { $_ -gt 0 }
    if ($validTSS.Count -lt 2) { return $null }
    $avg = ($validTSS | Measure-Object -Average).Average
    $stdDev = [math]::Sqrt(($validTSS | ForEach-Object { [math]::Pow($_ - $avg, 2) } | Measure-Object -Sum).Sum / $validTSS.Count)
    if ($stdDev -eq 0) { return $null }
    return [math]::Round($avg / $stdDev, 2)
}

function Calculate-Strain { 
    param([double]$WeeklyTSS, [double]$Monotony)
    if ($null -eq $Monotony -or $Monotony -eq 0) { return $null }
    return [math]::Round($WeeklyTSS * $Monotony, 0)
}

function Get-PerformanceScore {
    param([double]$CTL, [double]$TSB, [double]$RampRate)
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
    $score = 0; $factors = @()
    if ($TSB -ge -5) { $score += 25; $factors += "TSB otimo ($TSB)" }
    elseif ($TSB -ge -15) { $score += 15; $factors += "TSB aceitavel ($TSB)" }
    else { $score += 5; $factors += "TSB critico ($TSB)" }
    
    if ($HRV -ge 45) { $score += 25; $factors += "HRV bom ($HRV)" }
    elseif ($HRV -ge 38) { $score += 15; $factors += "HRV medio ($HRV)" }
    else { $score += 5; $factors += "HRV baixo ($HRV)" }
    
    if ($RestingHR -le 52) { $score += 25; $factors += "FC repouso normal ($RestingHR)" }
    elseif ($RestingHR -le 58) { $score += 15; $factors += "FC repouso elevada ($RestingHR)" }
    else { $score += 5; $factors += "FC repouso muito alta ($RestingHR)" }
    
    if ($Sleep -ge 7.5) { $score += 25; $factors += "Sono adequado ($Sleep h)" }
    elseif ($Sleep -ge 6.5) { $score += 15; $factors += "Sono razoavel ($Sleep h)" }
    else { $score += 5; $factors += "Sono insuficiente ($Sleep h)" }
    
    $status = if ($score -ge 80) { "EXCELENTE" }
              elseif ($score -ge 60) { "BOM" }
              elseif ($score -ge 40) { "MODERADO" }
              else { "CRITICO" }
    
    return @{ score = $score; status = $status; factors = $factors }
}

# =============================================================================
# INICIO DO SCRIPT
# =============================================================================

Write-Host ""
Write-Host "========================================"
Write-Host "ANALISE DE LONGO PRAZO - $WeeksBack SEMANAS"
Write-Host "COACH EDITION - Metricas Avancadas"
Write-Host "========================================"
Write-Host ""

$endDate = Get-Date
$startDate = $endDate.AddDays(-($WeeksBack * 7))
$oldest = $startDate.ToString("yyyy-MM-dd")
$newest = $endDate.ToString("yyyy-MM-dd")

Write-Host "Periodo: $oldest ate $newest" -ForegroundColor Green
Write-Host ""

# =============================================================================
# COLETAR INFORMACOES DO ATLETA
# =============================================================================

Write-Host "Coletando informacoes do atleta..." -ForegroundColor Cyan
$athlete = Invoke-IntervalsAPI -Endpoint "/athlete/$($config.athlete_id)"

$ftp = 200; $ftpRun = 300; $lthr = 165; $maxHr = 185
if ($athlete) {
    if ($athlete.ftp) { $ftp = $athlete.ftp }
    if ($athlete.ftpRun) { $ftpRun = $athlete.ftpRun }
    if ($athlete.lthr) { $lthr = $athlete.lthr }
    if ($athlete.max_hr) { $maxHr = $athlete.max_hr }
    Write-Host "  OK FTP: Bike=$ftp W, Run=$ftpRun W | LTHR=$lthr bpm | MaxHR=$maxHr bpm" -ForegroundColor Green
}

# =============================================================================
# COLETAR DADOS DE WELLNESS
# =============================================================================

Write-Host "Coletando dados de wellness..." -ForegroundColor Cyan
$wellness = Invoke-IntervalsAPI -Endpoint "/athlete/$($config.athlete_id)/wellness" -QueryParams @{
    oldest = $oldest
    newest = $newest
}

if (-not $wellness) {
    Write-Host "ERRO: Falha ao coletar dados de wellness!" -ForegroundColor Red
    exit
}
Write-Host "  OK $($wellness.Count) dias de wellness" -ForegroundColor Green

# =============================================================================
# COLETAR ATIVIDADES
# =============================================================================

Write-Host "Coletando atividades..." -ForegroundColor Cyan
$activities = Invoke-IntervalsAPI -Endpoint "/athlete/$($config.athlete_id)/activities" -QueryParams @{
    oldest = $oldest
    newest = $newest
}

if (-not $activities) {
    Write-Host "ERRO: Falha ao coletar atividades!" -ForegroundColor Red
    exit
}
Write-Host "  OK $($activities.Count) atividades encontradas" -ForegroundColor Green
Write-Host ""

# =============================================================================
# PROCESSAR ATIVIDADES DETALHADAS
# =============================================================================

Write-Host "Processando atividades detalhadas..." -ForegroundColor Cyan

$activitiesDetailed = @()
$activityCount = 0

foreach ($activity in $activities) {
    $activityCount++
    Write-Host "  [$activityCount/$($activities.Count)] Processando $($activity.type) - $($activity.name)" -ForegroundColor DarkGray
    
    $activityDetail = Invoke-IntervalsAPI -Endpoint "/activity/$($activity.id)"
    if (-not $activityDetail) { continue }
    
    $type = $activity.type
    $activityDate = ([DateTime]$activity.start_date_local).ToString("yyyy-MM-dd")
    
    # Metricas basicas
    $distanceKm = Convert-MetersToKm $activity.distance
    $distanceM = $activity.distance
    $movingTimeMin = Convert-SecondsToMinutes $activity.moving_time
    $elapsedTimeMin = Convert-SecondsToMinutes $activity.elapsed_time
    
    # FC
    $avgHr = Get-FieldValue $activityDetail @("average_hr", "avg_hr")
    $maxHr = Get-FieldValue $activityDetail @("max_hr")
    if (-not $avgHr) { $avgHr = Get-FieldValue $activity @("average_hr", "avg_hr") }
    if (-not $maxHr) { $maxHr = Get-FieldValue $activity @("max_hr") }
    
    # TSS e Calorias
    $tss = Get-FieldValue $activityDetail @("icu_training_load", "tss")
    if (-not $tss) { $tss = Get-FieldValue $activity @("icu_training_load", "tss") }
    $calories = Get-FieldValue $activityDetail @("calories", "kilojoules")
    if (-not $calories) { $calories = Get-FieldValue $activity @("calories", "kilojoules") }
    
    # Elevacao
    $elevationGain = Get-FieldValue $activityDetail @("total_elevation_gain", "elevation_gain")
    if (-not $elevationGain) { $elevationGain = Get-FieldValue $activity @("total_elevation_gain", "elevation_gain") }
    
    # Metricas especificas por modalidade
    $avgPower = $null; $maxPower = $null; $normalizedPower = $null
    $intensityFactor = $null; $variabilityIndex = $null
    $avgPace = $null; $avgSpeed = $null; $avgCadence = $null
    $decouplingPct = $null; $avgSwolf = $null
    $powerZones = $null; $hrZones = $null; $paceZones = $null
    $streams = $null; $powerStream = $null; $hrStream = $null; $timeStream = $null; $cadenceStream = $null; $speedStream = $null
    
    if ($type -eq "Ride" -or $type -eq "VirtualRide") {
        # BIKE METRICS
        $avgPower = Get-FieldValue $activityDetail @("icu_average_watts", "average_watts", "avg_watts")
        if (-not $avgPower) { $avgPower = Get-FieldValue $activity @("icu_average_watts", "average_watts", "avg_watts") }
        $maxPower = Get-FieldValue $activityDetail @("max_watts")
        if (-not $maxPower) { $maxPower = Get-FieldValue $activity @("max_watts") }
        $normalizedPower = Get-FieldValue $activityDetail @("np", "normalized_power", "icu_weighted_avg_watts", "weighted_avg_watts")
        if (-not $normalizedPower) { $normalizedPower = Get-FieldValue $activity @("icu_weighted_avg_watts", "weighted_avg_watts", "np", "normalized_power") }
        $avgCadence = Get-FieldValue $activityDetail @("average_cadence", "avg_cadence")
        if (-not $avgCadence) { $avgCadence = Get-FieldValue $activity @("average_cadence", "avg_cadence") }
        
        if ($normalizedPower -and $avgPower -and $avgPower -gt 0) {
            $variabilityIndex = [math]::Round($normalizedPower / $avgPower, 3)
        }
        if ($normalizedPower -and $ftp -gt 0) {
            $intensityFactor = [math]::Round($normalizedPower / $ftp, 3)
        }
        
        # Decoupling
        $hrLoad = Get-FieldValue $activityDetail @("icu_hr_load")
        $tssLoad = Get-FieldValue $activityDetail @("icu_training_load")
        if (-not $hrLoad) { $hrLoad = Get-FieldValue $activity @("icu_hr_load") }
        if (-not $tssLoad) { $tssLoad = Get-FieldValue $activity @("icu_training_load") }
        if ($hrLoad -and $tssLoad -and $tssLoad -gt 0) {
            $decouplingPct = [math]::Round((($hrLoad / $tssLoad) - 1) * 100, 2)
        }
        
        # Zonas de power
        if ($activityDetail.icu_power_zone_times) {
            $powerZones = Get-ZoneTimesFormatted -ZoneTimes $activityDetail.icu_power_zone_times -Type "power"
        } elseif ($activity.icu_power_zone_times) {
            $powerZones = Get-ZoneTimesFormatted -ZoneTimes $activity.icu_power_zone_times -Type "power"
        }

        if ($useStreams -and (-not $avgPower -or -not $normalizedPower -or -not $powerZones -or -not $decouplingPct)) {
            $streams = Invoke-IntervalsAPI -Endpoint "/activity/$($activity.id)/streams"
        }
        if ($streams) {
            $powerStream = Get-StreamData -Streams $streams -Keys @("power", "watts", "pwr")
            $hrStream = Get-StreamData -Streams $streams -Keys @("hr", "heartrate", "heart_rate")
            $timeStream = Get-StreamData -Streams $streams -Keys @("time", "seconds", "timestamp")
            $cadenceStream = Get-StreamData -Streams $streams -Keys @("cadence", "cad")

            if (-not $avgPower -and $powerStream) {
                $avgPower = [math]::Round((($powerStream | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average), 0)
            }
            if (-not $normalizedPower -and $powerStream) {
                $normalizedPower = Get-NormalizedPowerFromStream -Power ([double[]]$powerStream) -Time ([double[]]$timeStream)
            }
            if (-not $variabilityIndex -and $normalizedPower -and $avgPower -and $avgPower -gt 0) {
                $variabilityIndex = [math]::Round($normalizedPower / $avgPower, 3)
            }
            if (-not $intensityFactor -and $normalizedPower -and $ftp -gt 0) {
                $intensityFactor = [math]::Round($normalizedPower / $ftp, 3)
            }
            if (-not $decouplingPct -and $powerStream -and $hrStream) {
                $decouplingPct = Get-DecouplingFromStreams -Power ([double[]]$powerStream) -Hr ([double[]]$hrStream)
            }
            if (-not $powerZones -and $powerStream -and $ftp -gt 0) {
                $zoneTimes = Get-ZoneTimesFromStream -Values ([double[]]$powerStream) -FTP $ftp -Time ([double[]]$timeStream)
                if ($zoneTimes) { $powerZones = Get-ZoneTimesFormatted -ZoneTimes $zoneTimes -Type "power" }
            }
            if (-not $avgCadence -and $cadenceStream) {
                $avgCadence = [math]::Round((($cadenceStream | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average), 0)
            }
        }
    }
    elseif ($type -eq "Run") {
        # RUN METRICS
        $avgPower = Get-FieldValue $activityDetail @("icu_average_watts", "average_watts")
        if (-not $avgPower) { $avgPower = Get-FieldValue $activity @("icu_average_watts", "average_watts") }
        $normalizedPower = Get-FieldValue $activityDetail @("np", "normalized_power", "icu_weighted_avg_watts", "weighted_avg_watts")
        if (-not $normalizedPower) { $normalizedPower = Get-FieldValue $activity @("icu_weighted_avg_watts", "weighted_avg_watts", "np", "normalized_power") }
        $avgCadence = Get-FieldValue $activityDetail @("average_cadence", "avg_cadence")
        if (-not $avgCadence) { $avgCadence = Get-FieldValue $activity @("average_cadence", "avg_cadence") }
        
        # Pace
        $paceMs = Get-FieldValue $activityDetail @("icu_pace", "pace", "average_pace")
        if (-not $paceMs) { $paceMs = Get-FieldValue $activity @("icu_pace", "pace", "average_pace") }
        if ($paceMs) { 
            $avgPace = Convert-PaceToMinKm -MetersPerSecond $paceMs
            $avgSpeed = Convert-SpeedToKmh -MetersPerSecond $paceMs
        }
        
        if ($normalizedPower -and $avgPower -and $avgPower -gt 0) {
            $variabilityIndex = [math]::Round($normalizedPower / $avgPower, 3)
        }
        if ($normalizedPower -and $ftpRun -gt 0) {
            $intensityFactor = [math]::Round($normalizedPower / $ftpRun, 3)
        }
        
        # Decoupling
        $hrLoad = Get-FieldValue $activityDetail @("icu_hr_load")
        $tssLoad = Get-FieldValue $activityDetail @("icu_training_load")
        if (-not $hrLoad) { $hrLoad = Get-FieldValue $activity @("icu_hr_load") }
        if (-not $tssLoad) { $tssLoad = Get-FieldValue $activity @("icu_training_load") }
        if ($hrLoad -and $tssLoad -and $tssLoad -gt 0) {
            $decouplingPct = [math]::Round((($hrLoad / $tssLoad) - 1) * 100, 2)
        }
        
        # Zonas de pace
        if ($activityDetail.icu_pace_zone_times) {
            $paceZones = Get-ZoneTimesFormatted -ZoneTimes $activityDetail.icu_pace_zone_times -Type "pace"
        } elseif ($activity.icu_pace_zone_times) {
            $paceZones = Get-ZoneTimesFormatted -ZoneTimes $activity.icu_pace_zone_times -Type "pace"
        }

        if ($useStreams -and (-not $avgPower -or -not $normalizedPower -or -not $paceZones -or -not $decouplingPct -or -not $avgPace)) {
            $streams = Invoke-IntervalsAPI -Endpoint "/activity/$($activity.id)/streams"
        }
        if ($streams) {
            $powerStream = Get-StreamData -Streams $streams -Keys @("power", "watts", "pwr")
            $hrStream = Get-StreamData -Streams $streams -Keys @("hr", "heartrate", "heart_rate")
            $timeStream = Get-StreamData -Streams $streams -Keys @("time", "seconds", "timestamp")
            $cadenceStream = Get-StreamData -Streams $streams -Keys @("cadence", "cad")
            $speedStream = Get-StreamData -Streams $streams -Keys @("speed", "velocity", "pace")

            if (-not $avgPower -and $powerStream) {
                $avgPower = [math]::Round((($powerStream | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average), 0)
            }
            if (-not $normalizedPower -and $powerStream) {
                $normalizedPower = Get-NormalizedPowerFromStream -Power ([double[]]$powerStream) -Time ([double[]]$timeStream)
            }
            if (-not $variabilityIndex -and $normalizedPower -and $avgPower -and $avgPower -gt 0) {
                $variabilityIndex = [math]::Round($normalizedPower / $avgPower, 3)
            }
            if (-not $intensityFactor -and $normalizedPower -and $ftpRun -gt 0) {
                $intensityFactor = [math]::Round($normalizedPower / $ftpRun, 3)
            }
            if (-not $decouplingPct -and $powerStream -and $hrStream) {
                $decouplingPct = Get-DecouplingFromStreams -Power ([double[]]$powerStream) -Hr ([double[]]$hrStream)
            }
            if (-not $paceZones -and $powerStream -and $ftpRun -gt 0) {
                $zoneTimes = Get-ZoneTimesFromStream -Values ([double[]]$powerStream) -FTP $ftpRun -Time ([double[]]$timeStream)
                if ($zoneTimes) { $paceZones = Get-ZoneTimesFormatted -ZoneTimes $zoneTimes -Type "power" }
            }
            if (-not $avgCadence -and $cadenceStream) {
                $avgCadence = [math]::Round((($cadenceStream | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average), 0)
            }
            if (-not $avgPace -and $speedStream) {
                $avgSpeedMs = [math]::Round((($speedStream | Where-Object { $_ -gt 0 } | Measure-Object -Average).Average), 3)
                if ($avgSpeedMs -and $avgSpeedMs -gt 0) {
                    $avgSpeed = Convert-SpeedToKmh -MetersPerSecond $avgSpeedMs
                    $avgPace = Convert-PaceToMinKm -MetersPerSecond $avgSpeedMs
                }
            }
        }
    }
    elseif ($type -eq "Swim") {
        # SWIM METRICS
        $paceMs = Get-FieldValue $activityDetail @("icu_pace", "pace", "average_pace")
        if (-not $paceMs) { $paceMs = Get-FieldValue $activity @("icu_pace", "pace", "average_pace") }
        if ($paceMs -and $paceMs -gt 0) {
            $pace100m = 100 / $paceMs
            $minutes = [int][math]::Floor($pace100m / 60)
            $seconds = [int][math]::Round($pace100m % 60)
            $avgPace = "{0:D2}:{1:D2}/100m" -f $minutes, $seconds
        }
        
        $avgSwolf = Get-FieldValue $activityDetail @("average_swolf", "avg_swolf")
        if (-not $avgSwolf) { $avgSwolf = Get-FieldValue $activity @("average_swolf", "avg_swolf") }
    }
    
    # Zonas de FC (para todos os tipos)
    if ($activityDetail.icu_hr_zone_times) {
        $hrZones = Get-ZoneTimesFormatted -ZoneTimes $activityDetail.icu_hr_zone_times -Type "hr"
    } elseif ($activity.icu_hr_zone_times) {
        $hrZones = Get-ZoneTimesFormatted -ZoneTimes $activity.icu_hr_zone_times -Type "hr"
    }
    
    # RPE e Feel
    $icuRpe = Get-FieldValue $activityDetail @("icu_rpe", "rpe")
    $feel = Get-FieldValue $activityDetail @("feel")
    
    # Notas
    $notes = if ($activity.description) { $activity.description } else { "" }
    
    $activitiesDetailed += [PSCustomObject]@{
        id = $activity.id
        name = $activity.name
        type = $type
        start_date_local = $activityDate
        distance_km = $distanceKm
        distance_m = $distanceM
        moving_time_min = $movingTimeMin
        elapsed_time_min = $elapsedTimeMin
        elevation_gain_m = $elevationGain
        average_hr = $avgHr
        max_hr = $maxHr
        average_power = $avgPower
        max_power = $maxPower
        normalized_power = $normalizedPower
        intensity_factor = $intensityFactor
        variability_index = $variabilityIndex
        tss = $tss
        calories = $calories
        average_pace = $avgPace
        average_speed_kmh = $avgSpeed
        average_cadence = $avgCadence
        decoupling_pct = $decouplingPct
        average_swolf = $avgSwolf
        icu_rpe = $icuRpe
        feel = $feel
        notes = $notes
        power_zones = $powerZones
        hr_zones = $hrZones
        pace_zones = $paceZones
    }
}

Write-Host "  OK $($activitiesDetailed.Count) atividades processadas" -ForegroundColor Green
Write-Host ""

# =============================================================================
# ANALISE POR SEMANA
# =============================================================================

Write-Host "Analisando por semana..." -ForegroundColor Cyan

$weeklyAnalysis = @()
$currentWeekStart = $startDate

while ($currentWeekStart -lt $endDate) {
    
    $weekEnd = $currentWeekStart.AddDays(6)
    $weekStartStr = $currentWeekStart.ToString("yyyy-MM-dd")
    $weekEndStr = $weekEnd.ToString("yyyy-MM-dd")
    
    # Wellness da semana
    $weekWellness = $wellness | Where-Object {
        $date = [DateTime]$_.id
        $date -ge $currentWeekStart -and $date -le $weekEnd
    }
    
    # Atividades da semana
    $weekActivities = $activitiesDetailed | Where-Object {
        $date = [DateTime]$_.start_date_local
        $date -ge $currentWeekStart -and $date -le $weekEnd
    }
    
    # Metricas de carga
    $endCTL = if ($weekWellness.Count -gt 0) { ($weekWellness | Select-Object -Last 1).ctl } else { 0 }
    $endATL = if ($weekWellness.Count -gt 0) { ($weekWellness | Select-Object -Last 1).atl } else { 0 }
    $endTSB = if ($weekWellness.Count -gt 0) { ($weekWellness | Select-Object -Last 1).tsb } else { 0 }
    $avgRampRate = if ($weekWellness.Count -gt 0) { ($weekWellness | Measure-Object -Property rampRate -Average).Average } else { 0 }
    
    # Totais da semana
    $totalTSS = ($weekActivities | Where-Object { $_.tss } | Measure-Object -Property tss -Sum).Sum
    $totalDistance = ($weekActivities | Where-Object { $_.distance_km } | Measure-Object -Property distance_km -Sum).Sum
    $totalTime = ($weekActivities | Measure-Object -Property moving_time_min -Sum).Sum
    $totalCalories = ($weekActivities | Where-Object { $_.calories } | Measure-Object -Property calories -Sum).Sum
    $totalElevation = ($weekActivities | Where-Object { $_.elevation_gain_m } | Measure-Object -Property elevation_gain_m -Sum).Sum
    
    # Wellness medios
    $avgWeight = ($weekWellness | Where-Object { $_.weight -gt 0 } | Measure-Object -Property weight -Average).Average
    $avgHRV = ($weekWellness | Where-Object { $_.hrv -gt 0 } | Measure-Object -Property hrv -Average).Average
    $avgRHR = ($weekWellness | Where-Object { $_.restingHR -gt 0 } | Measure-Object -Property restingHR -Average).Average
    $avgSleep = ($weekWellness | Where-Object { $_.sleepSecs -gt 0 } | ForEach-Object { $_.sleepSecs / 3600 } | Measure-Object -Average).Average
    
    # Monotonia e Strain
    $dailyTSSArray = @()
    for ($d = 0; $d -lt 7; $d++) {
        $dayDate = $currentWeekStart.AddDays($d)
        $dayActivities = $weekActivities | Where-Object { 
            ([DateTime]$_.start_date_local).Date -eq $dayDate.Date 
        }
        $dayTSS = ($dayActivities | Where-Object { $_.tss } | Measure-Object -Property tss -Sum).Sum
        $dailyTSSArray += if ($dayTSS) { $dayTSS } else { 0 }
    }
    
    $monotony = Calculate-Monotony -DailyTSS $dailyTSSArray
    $strain = Calculate-Strain -WeeklyTSS $totalTSS -Monotony $monotony
    
    # Analise por modalidade
    $bikeActivities = $weekActivities | Where-Object { $_.type -match "Ride" }
    $runActivities = $weekActivities | Where-Object { $_.type -eq "Run" }
    $swimActivities = $weekActivities | Where-Object { $_.type -eq "Swim" }
    $strengthActivities = $weekActivities | Where-Object { $_.type -match "Weight|Strength" }
    
    $bikeTime = ($bikeActivities | Measure-Object -Property moving_time_min -Sum).Sum
    $runTime = ($runActivities | Measure-Object -Property moving_time_min -Sum).Sum
    $swimTime = ($swimActivities | Measure-Object -Property moving_time_min -Sum).Sum
    $strengthTime = ($strengthActivities | Measure-Object -Property moving_time_min -Sum).Sum
    
    $totalTrainingTime = $bikeTime + $runTime + $swimTime + $strengthTime
    $bikePct = if ($totalTrainingTime -gt 0) { [math]::Round(($bikeTime / $totalTrainingTime) * 100, 1) } else { 0 }
    $runPct = if ($totalTrainingTime -gt 0) { [math]::Round(($runTime / $totalTrainingTime) * 100, 1) } else { 0 }
    $swimPct = if ($totalTrainingTime -gt 0) { [math]::Round(($swimTime / $totalTrainingTime) * 100, 1) } else { 0 }
    $strengthPct = if ($totalTrainingTime -gt 0) { [math]::Round(($strengthTime / $totalTrainingTime) * 100, 1) } else { 0 }
    
    # Scores
    $perfScore = Get-PerformanceScore -CTL $endCTL -TSB $endTSB -RampRate $avgRampRate
    $recovery = Get-RecoveryStatus -TSB $endTSB -HRV $avgHRV -RestingHR $avgRHR -Sleep $avgSleep
    
    # Metricas avancadas por modalidade
    $bikeAnalysis = @{
        total = $bikeActivities.Count
        tss = ($bikeActivities | Where-Object { $_.tss } | Measure-Object -Property tss -Sum).Sum
        distance_km = ($bikeActivities | Where-Object { $_.distance_km } | Measure-Object -Property distance_km -Sum).Sum
        time_min = $bikeTime
        avg_power = ($bikeActivities | Where-Object { $_.average_power } | Measure-Object -Property average_power -Average).Average
        avg_np = ($bikeActivities | Where-Object { $_.normalized_power } | Measure-Object -Property normalized_power -Average).Average
        avg_if = ($bikeActivities | Where-Object { $_.intensity_factor } | Measure-Object -Property intensity_factor -Average).Average
        avg_vi = ($bikeActivities | Where-Object { $_.variability_index } | Measure-Object -Property variability_index -Average).Average
        avg_hr = ($bikeActivities | Where-Object { $_.average_hr } | Measure-Object -Property average_hr -Average).Average
        avg_decoupling = ($bikeActivities | Where-Object { $_.decoupling_pct } | Measure-Object -Property decoupling_pct -Average).Average
    }

    $runAnalysis = @{
        total = $runActivities.Count
        tss = ($runActivities | Where-Object { $_.tss } | Measure-Object -Property tss -Sum).Sum
        distance_km = ($runActivities | Where-Object { $_.distance_km } | Measure-Object -Property distance_km -Sum).Sum
        time_min = $runTime
        avg_hr = ($runActivities | Where-Object { $_.average_hr } | Measure-Object -Property average_hr -Average).Average
        avg_cadence = ($runActivities | Where-Object { $_.average_cadence } | Measure-Object -Property average_cadence -Average).Average
        avg_decoupling = ($runActivities | Where-Object { $_.decoupling_pct } | Measure-Object -Property decoupling_pct -Average).Average
    }
    
    $swimAnalysis = @{
        total = $swimActivities.Count
        tss = ($swimActivities | Where-Object { $_.tss } | Measure-Object -Property tss -Sum).Sum
        distance_m = ($swimActivities | Where-Object { $_.distance_m } | Measure-Object -Property distance_m -Sum).Sum
        time_min = $swimTime
        avg_hr = ($swimActivities | Where-Object { $_.average_hr } | Measure-Object -Property average_hr -Average).Average
        avg_swolf = ($swimActivities | Where-Object { $_.average_swolf } | Measure-Object -Property average_swolf -Average).Average
    }

    $bikeZones = Sum-SeilerZones -Activities $bikeActivities -PropertyName "power_zones"
    $runZones = Sum-SeilerZones -Activities $runActivities -PropertyName "pace_zones"
    
    $weeklyAnalysis += [PSCustomObject]@{
        semana = $currentWeekStart.ToString("dd/MM")
        inicio = $weekStartStr
        fim = $weekEndStr
        total_atividades = $weekActivities.Count
        total_tss = if ($totalTSS) { [math]::Round($totalTSS, 0) } else { 0 }
        distancia_km = if ($totalDistance) { [math]::Round($totalDistance, 1) } else { 0 }
        tempo_horas = [math]::Round($totalTime / 60, 1)
        calorias = if ($totalCalories) { [math]::Round($totalCalories, 0) } else { 0 }
        elevacao_m = if ($totalElevation) { [math]::Round($totalElevation, 0) } else { 0 }
        ctl = [math]::Round($endCTL, 1)
        atl = [math]::Round($endATL, 1)
        tsb = [math]::Round($endTSB, 1)
        rampRate = [math]::Round($avgRampRate, 1)
        monotonia = $monotony
        strain = $strain
        peso_medio = if ($avgWeight) { [math]::Round($avgWeight, 1) } else { $null }
        hrv_medio = if ($avgHRV) { [math]::Round($avgHRV, 1) } else { $null }
        fc_repouso_media = if ($avgRHR) { [math]::Round($avgRHR, 1) } else { $null }
        sono_medio = if ($avgSleep) { [math]::Round($avgSleep, 1) } else { $null }
        performance_score = $perfScore
        recovery_score = $recovery.score
        recovery_status = $recovery.status
        bike_pct = $bikePct
        run_pct = $runPct
        swim_pct = $swimPct
        strength_pct = $strengthPct
        bike_analise = $bikeAnalysis
        run_analise = $runAnalysis
        swim_analise = $swimAnalysis
        bike_zones = $bikeZones
        run_zones = $runZones
    }
    
    $currentWeekStart = $currentWeekStart.AddDays(7)
}

Write-Host "  OK $($weeklyAnalysis.Count) semanas analisadas" -ForegroundColor Green
Write-Host ""

# =============================================================================
# ANALISE DE BLOCOS (4 semanas)
# =============================================================================

Write-Host "Analisando blocos de treinamento..." -ForegroundColor Cyan

$blocksCount = [math]::Floor($weeklyAnalysis.Count / 4)
$blockAnalysis = @()

for ($i = 0; $i -lt $blocksCount; $i++) {
    
    $blockWeeks = $weeklyAnalysis | Select-Object -Skip ($i * 4) -First 4
    
    $blockAnalysis += [PSCustomObject]@{
        bloco = "Bloco $($i + 1)"
        semanas = "$($blockWeeks[0].semana) - $($blockWeeks[-1].semana)"
        total_tss = ($blockWeeks | Measure-Object -Property total_tss -Sum).Sum
        media_tss_semanal = [math]::Round(($blockWeeks | Measure-Object -Property total_tss -Average).Average, 0)
        distancia_total = [math]::Round(($blockWeeks | Measure-Object -Property distancia_km -Sum).Sum, 1)
        tempo_total = [math]::Round(($blockWeeks | Measure-Object -Property tempo_horas -Sum).Sum, 1)
        calorias_total = [math]::Round(($blockWeeks | Measure-Object -Property calorias -Sum).Sum, 0)
        ctl_final = ($blockWeeks | Select-Object -Last 1).ctl
        tsb_final = ($blockWeeks | Select-Object -Last 1).tsb
        avg_monotonia = [math]::Round(($blockWeeks | Where-Object { $_.monotonia } | Measure-Object -Property monotonia -Average).Average, 2)
        avg_strain = [math]::Round(($blockWeeks | Where-Object { $_.strain } | Measure-Object -Property strain -Average).Average, 0)
        avg_performance_score = [math]::Round(($blockWeeks | Measure-Object -Property performance_score -Average).Average, 0)
        avg_recovery_score = [math]::Round(($blockWeeks | Measure-Object -Property recovery_score -Average).Average, 0)
        bike_pct = [math]::Round(($blockWeeks | Measure-Object -Property bike_pct -Average).Average, 1)
        run_pct = [math]::Round(($blockWeeks | Measure-Object -Property run_pct -Average).Average, 1)
        swim_pct = [math]::Round(($blockWeeks | Measure-Object -Property swim_pct -Average).Average, 1)
    }
}

Write-Host "  OK $blocksCount blocos analisados" -ForegroundColor Green
Write-Host ""

# =============================================================================
# PROGRESSAO DE METRICAS
# =============================================================================

Write-Host "Analisando progressoes..." -ForegroundColor Cyan

$progression = @{
    CTL = @{
        inicio = ($weeklyAnalysis | Select-Object -First 1).ctl
        fim = ($weeklyAnalysis | Select-Object -Last 1).ctl
        variacao = ($weeklyAnalysis | Select-Object -Last 1).ctl - ($weeklyAnalysis | Select-Object -First 1).ctl
        pct = if (($weeklyAnalysis | Select-Object -First 1).ctl -gt 0) {
            [math]::Round(((($weeklyAnalysis | Select-Object -Last 1).ctl - ($weeklyAnalysis | Select-Object -First 1).ctl) / ($weeklyAnalysis | Select-Object -First 1).ctl) * 100, 1)
        } else { 0 }
    }
    TSS_Semanal = @{
        media_primeiro_mes = [math]::Round(($weeklyAnalysis | Select-Object -First 4 | Measure-Object -Property total_tss -Average).Average, 0)
        media_ultimo_mes = [math]::Round(($weeklyAnalysis | Select-Object -Last 4 | Measure-Object -Property total_tss -Average).Average, 0)
        variacao = [math]::Round(($weeklyAnalysis | Select-Object -Last 4 | Measure-Object -Property total_tss -Average).Average - ($weeklyAnalysis | Select-Object -First 4 | Measure-Object -Property total_tss -Average).Average, 0)
    }
    Peso = @{
        inicio = ($weeklyAnalysis | Select-Object -First 1).peso_medio
        fim = ($weeklyAnalysis | Select-Object -Last 1).peso_medio
        variacao = if (($weeklyAnalysis | Select-Object -Last 1).peso_medio -and ($weeklyAnalysis | Select-Object -First 1).peso_medio) {
            ($weeklyAnalysis | Select-Object -Last 1).peso_medio - ($weeklyAnalysis | Select-Object -First 1).peso_medio
        } else { $null }
    }
    HRV = @{
        media_primeiro_mes = [math]::Round(($weeklyAnalysis | Select-Object -First 4 | Where-Object { $_.hrv_medio } | Measure-Object -Property hrv_medio -Average).Average, 1)
        media_ultimo_mes = [math]::Round(($weeklyAnalysis | Select-Object -Last 4 | Where-Object { $_.hrv_medio } | Measure-Object -Property hrv_medio -Average).Average, 1)
    }
    Monotonia = @{
        media_periodo = [math]::Round(($weeklyAnalysis | Where-Object { $_.monotonia } | Measure-Object -Property monotonia -Average).Average, 2)
        max = ($weeklyAnalysis | Where-Object { $_.monotonia } | Measure-Object -Property monotonia -Maximum).Maximum
    }
}

Write-Host "  OK Progressoes calculadas" -ForegroundColor Green
Write-Host ""

# =============================================================================
# PREDICAO PARA EVENTO
# =============================================================================

$eventPrediction = $null

if ($EventDate) {
    Write-Host "Calculando predicao para evento ($EventDate)..." -ForegroundColor Cyan
    
    try {
        $eventDateObj = [DateTime]::ParseExact($EventDate, "yyyy-MM-dd", $null)
        $daysToEvent = ($eventDateObj - $endDate).Days
        
        if ($daysToEvent -gt 0 -and $daysToEvent -le 90) {
            
            $avgRampRate = ($weeklyAnalysis | Select-Object -Last 4 | Measure-Object -Property rampRate -Average).Average
            $currentCTL = ($weeklyAnalysis | Select-Object -Last 1).ctl
            $weeksToEvent = [math]::Ceiling($daysToEvent / 7)
            
            $projectedCTL = $currentCTL + ($avgRampRate * $weeksToEvent)
            $optimalTaperCTL = $projectedCTL * 0.95
            $taperTSB = 5
            
            $eventPrediction = @{
                data_evento = $EventDate
                dias_restantes = $daysToEvent
                semanas_restantes = $weeksToEvent
                ctl_atual = [math]::Round($currentCTL, 1)
                ctl_projetado = [math]::Round($projectedCTL, 1)
                ctl_ideal_prova = [math]::Round($optimalTaperCTL, 1)
                tsb_ideal_prova = $taperTSB
                ramp_rate_atual = [math]::Round($avgRampRate, 1)
                recomendacao_taper = "Reduzir carga em 40% na ultima semana"
                status_preparacao = if ($projectedCTL -ge 60) { "EXCELENTE" }
                                    elseif ($projectedCTL -ge 45) { "BOA" }
                                    elseif ($projectedCTL -ge 35) { "MODERADA" }
                                    else { "INSUFICIENTE" }
            }
            
            Write-Host "  OK Predicao calculada: CTL projetado = $($eventPrediction.ctl_projetado)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  ERRO: Data do evento invalida (use formato yyyy-MM-dd)" -ForegroundColor Red
    }
}

# =============================================================================
# INSIGHTS E RECOMENDACOES
# =============================================================================

Write-Host "Gerando insights..." -ForegroundColor Cyan

$insights = @()
$warnings = @()

# Progressao de fitness
if ($progression.CTL.pct -gt 20) {
    $insights += "OK Excelente progressao de fitness: CTL aumentou $($progression.CTL.pct)%"
}
elseif ($progression.CTL.pct -lt 0) {
    $warnings += "[ALERTA] CTL em declinio: $($progression.CTL.pct)% - considere aumentar volume"
}
elseif ($progression.CTL.pct -lt 5) {
    $insights += "[INFO] Progressao lenta: CTL variou apenas $($progression.CTL.pct)%"
}

# TSS semanal
if ($progression.TSS_Semanal.variacao -gt 100) {
    $warnings += "[ALERTA] TSS semanal aumentou muito: +$($progression.TSS_Semanal.variacao) TSS/semana"
}
elseif ($progression.TSS_Semanal.variacao -lt -50) {
    $insights += "[INFO] TSS semanal diminuiu: $($progression.TSS_Semanal.variacao) TSS/semana"
}

# Consistencia
$totalWeeks = $weeklyAnalysis.Count
$activeWeeks = ($weeklyAnalysis | Where-Object { $_.total_atividades -gt 0 }).Count
$consistency = [math]::Round(($activeWeeks / $totalWeeks) * 100, 0)
$insights += "Consistencia: $consistency% ($activeWeeks/$totalWeeks semanas ativas)"

# Carga media
$avgWeeklyTSS = ($weeklyAnalysis | Measure-Object -Property total_tss -Average).Average
if ($avgWeeklyTSS -lt 200) {
    $insights += "[INFO] Carga semanal media baixa ($([math]::Round($avgWeeklyTSS, 0)) TSS) - ha espaco para crescimento"
}
elseif ($avgWeeklyTSS -gt 600) {
    $warnings += "[ALERTA] Carga semanal media muito alta ($([math]::Round($avgWeeklyTSS, 0)) TSS) - risco de overtraining"
}

# Monotonia
if ($progression.Monotonia.media_periodo -and $progression.Monotonia.media_periodo -gt 2.0) {
    $warnings += "[ALERTA] Monotonia media alta ($($progression.Monotonia.media_periodo)) - varie intensidades"
}

# Recuperacao
$avgRecoveryScore = ($weeklyAnalysis | Measure-Object -Property recovery_score -Average).Average
if ($avgRecoveryScore -lt 50) {
    $warnings += "[ALERTA] Score medio de recuperacao baixo ($([math]::Round($avgRecoveryScore, 0))) - priorize sono e descanso"
}

# HRV
if ($progression.HRV.media_ultimo_mes -and $progression.HRV.media_primeiro_mes) {
    if ($progression.HRV.media_ultimo_mes -lt $progression.HRV.media_primeiro_mes) {
        $hrvDrop = $progression.HRV.media_primeiro_mes - $progression.HRV.media_ultimo_mes
        $warnings += "[ALERTA] HRV medio caiu $([math]::Round($hrvDrop, 1)) pontos - possivel fadiga acumulada"
    }
}

# Peso
if ($progression.Peso.variacao) {
    if ([math]::Abs($progression.Peso.variacao) -gt 3) {
        $insights += "[INFO] Variacao de peso significativa: $([math]::Round($progression.Peso.variacao, 1)) kg"
    }
}

# Proporcao de modalidades
$avgBikePct = ($weeklyAnalysis | Measure-Object -Property bike_pct -Average).Average
$avgRunPct = ($weeklyAnalysis | Measure-Object -Property run_pct -Average).Average
$avgSwimPct = ($weeklyAnalysis | Measure-Object -Property swim_pct -Average).Average

$insights += "Proporcao media: Bike $([math]::Round($avgBikePct, 0))% | Run $([math]::Round($avgRunPct, 0))% | Swim $([math]::Round($avgSwimPct, 0))%"

if ($avgSwimPct -lt 10 -and $avgSwimPct -gt 0) {
    $insights += "[INFO] Natacao representa menos de 10% do treino - considere aumentar volume"
}

Write-Host "  OK $($insights.Count) insights e $($warnings.Count) alertas gerados" -ForegroundColor Green
Write-Host ""

# =============================================================================
# RELATORIO FINAL
# =============================================================================

$report = [PSCustomObject]@{
    relatorio = @{
        tipo = "Analise de Longo Prazo - Coach Edition"
        gerado_em = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        versao = "4.0"
        periodo_semanas = $WeeksBack
        inicio = $oldest
        fim = $newest
        athlete_id = $config.athlete_id
    }
    
    resumo_geral = @{
        total_atividades = $activitiesDetailed.Count
        total_semanas_ativas = $activeWeeks
        consistencia_pct = $consistency
        tss_medio_semanal = [math]::Round($avgWeeklyTSS, 0)
        distancia_total_km = [math]::Round(($weeklyAnalysis | Measure-Object -Property distancia_km -Sum).Sum, 1)
        tempo_total_horas = [math]::Round(($weeklyAnalysis | Measure-Object -Property tempo_horas -Sum).Sum, 1)
        calorias_total = [math]::Round(($weeklyAnalysis | Measure-Object -Property calorias -Sum).Sum, 0)
        elevacao_total_m = [math]::Round(($weeklyAnalysis | Measure-Object -Property elevacao_m -Sum).Sum, 0)
    }
    
    metricas_atleta = @{
        FTP_bike = $ftp
        FTP_run = $ftpRun
        LTHR = $lthr
        Max_HR = $maxHr
    }
    
    analise_semanal = $weeklyAnalysis
    analise_blocos = $blockAnalysis
    progressoes = $progression
    predicao_evento = $eventPrediction
    insights = $insights
    alertas = $warnings
}

# =============================================================================
# EXPORTAR
# =============================================================================

$fileName = "intervals_longterm_${WeeksBack}weeks_coach_edition.json"
$jsonPath = Join-Path -Path $OutputDir -ChildPath $fileName

$report | ConvertTo-Json -Depth 15 | Out-File -FilePath $jsonPath -Encoding UTF8

# =============================================================================
# EXIBIR RESUMO
# =============================================================================

Write-Host "========================================"
Write-Host "ANALISE CONCLUIDA"
Write-Host "========================================"
Write-Host ""

Write-Host "RESUMO - $WeeksBack SEMANAS" -ForegroundColor Yellow
Write-Host ""
Write-Host "Total de atividades: $($activitiesDetailed.Count)"
Write-Host "Consistencia: $consistency% ($activeWeeks/$totalWeeks semanas)"
Write-Host "TSS medio semanal: $([math]::Round($avgWeeklyTSS, 0))"
Write-Host "Distancia total: $([math]::Round(($weeklyAnalysis | Measure-Object -Property distancia_km -Sum).Sum, 1)) km"
Write-Host "Tempo total: $([math]::Round(($weeklyAnalysis | Measure-Object -Property tempo_horas -Sum).Sum, 1)) horas"
Write-Host "Progressao CTL: $($progression.CTL.inicio) -> $($progression.CTL.fim) ($($progression.CTL.pct)%)"
Write-Host ""

Write-Host "METRICAS AVANCADAS" -ForegroundColor Yellow
Write-Host "Monotonia media: $($progression.Monotonia.media_periodo)"
Write-Host "TSS variacao: $($progression.TSS_Semanal.variacao) TSS/semana"
if ($progression.Peso.variacao) {
    Write-Host "Peso variacao: $([math]::Round($progression.Peso.variacao, 1)) kg"
}
Write-Host ""

Write-Host "PROPORCAO DE MODALIDADES" -ForegroundColor Yellow
Write-Host "Bike: $([math]::Round($avgBikePct, 0))% | Run: $([math]::Round($avgRunPct, 0))% | Swim: $([math]::Round($avgSwimPct, 0))%"
Write-Host ""

if ($eventPrediction) {
    Write-Host "PREDICAO PARA EVENTO ($EventDate)" -ForegroundColor Yellow
    Write-Host "Dias restantes: $($eventPrediction.dias_restantes)"
    Write-Host "CTL atual: $($eventPrediction.ctl_atual)"
    Write-Host "CTL projetado: $($eventPrediction.ctl_projetado)"
    Write-Host "Status de preparacao: $($eventPrediction.status_preparacao)"
    Write-Host ""
}

if ($insights.Count -gt 0) {
    Write-Host "INSIGHTS" -ForegroundColor Yellow
    foreach ($insight in $insights) { Write-Host "  $insight" }
    Write-Host ""
}

if ($warnings.Count -gt 0) {
    Write-Host "ALERTAS" -ForegroundColor Yellow
    foreach ($warning in $warnings) { Write-Host "  $warning" }
    Write-Host ""
}

Write-Host "OK Relatorio completo salvo em: $jsonPath" -ForegroundColor Green
Write-Host ""
