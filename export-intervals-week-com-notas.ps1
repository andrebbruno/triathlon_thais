# export-intervals-week-com-notas.ps1
# Intervals.icu only: activities + wellness + notes (chat messages)

param (
    [string]$AthleteId = "0",
    [string]$ApiKeyPath = "$PSScriptRoot\api_key.txt",
    [string]$StartDate,
    [string]$EndDate
)

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

function Get-WeekRange {
    $today = Get-Date
    $dayOfWeek = [int]$today.DayOfWeek
    if ($dayOfWeek -eq 0) { $monday = $today.AddDays(-6) } else { $monday = $today.AddDays(-($dayOfWeek - 1)) }
    $sunday = $monday.AddDays(6)
    return @{
        Start = $monday.ToString("yyyy-MM-dd")
        End   = $sunday.ToString("yyyy-MM-dd")
    }
}

function Get-FirstValue {
    param ([object[]]$Values)
    foreach ($value in $Values) {
        if ($null -ne $value -and $value -ne "") { return $value }
    }
    return $null
}

if (-not $StartDate -or -not $EndDate) {
    $range = Get-WeekRange
    $StartDate = $range.Start
    $EndDate = $range.End
}

$apiKey = Get-ApiKey -Path $ApiKeyPath
if (-not $apiKey) {
    Write-Host "API key not provided."
    exit 1
}

$pair = "API_KEY:$apiKey"
$base64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
$headers = @{
    Authorization = "Basic $base64"
    Accept        = "application/json"
}

$activitiesUrl = "https://intervals.icu/api/v1/athlete/$AthleteId/activities?oldest=$StartDate&newest=$EndDate"
$wellnessUrl   = "https://intervals.icu/api/v1/athlete/$AthleteId/wellness?oldest=$StartDate&newest=$EndDate"
$eventsUrl     = "https://intervals.icu/api/v1/athlete/$AthleteId/events?oldest=$StartDate&newest=$EndDate"
function Get-ChatMessages {
    param (
        [int]$ChatId,
        [hashtable]$Headers,
        [hashtable]$Cache,
        [string]$OwnerId
    )

    if (-not $ChatId) { return @() }
    if ($Cache.ContainsKey($ChatId)) { return $Cache[$ChatId] }

    try {
        $uri = "https://intervals.icu/api/v1/chats/$ChatId/messages"
        $messages = Invoke-RestMethod -Uri $uri -Headers $Headers -Method Get -ErrorAction Stop
    } catch {
        $Cache[$ChatId] = @()
        return @()
    }

    $filtered = $messages | Where-Object {
        $_.type -eq "TEXT" -and $_.content -and ($null -eq $OwnerId -or $_.athlete_id -eq $OwnerId)
    } | Sort-Object created | ForEach-Object { $_.content.Trim() }

    $Cache[$ChatId] = @($filtered)
    return $Cache[$ChatId]
}

Write-Host "Fetching activities from Intervals.icu..."
try {
    $activities = Invoke-RestMethod -Uri $activitiesUrl -Headers $headers -Method Get -ErrorAction Stop
} catch {
    Write-Host "Failed to fetch activities: $($_.Exception.Message)"
    exit 1
}

if ($activities.Count -gt 0) {
    $ownerId = Get-FirstValue @($activities[0].icu_athlete_id, $activities[0].athlete_id)
} elseif ($AthleteId -ne "0") {
    $ownerId = $AthleteId
} else {
    $ownerId = $null
}

Write-Host "Fetching wellness data..."
try {
    $wellness = Invoke-RestMethod -Uri $wellnessUrl -Headers $headers -Method Get -ErrorAction Stop
} catch {
    Write-Host "Failed to fetch wellness: $($_.Exception.Message)"
    $wellness = @()
}

Write-Host "Fetching weekly notes..."
try {
    $events = Invoke-RestMethod -Uri $eventsUrl -Headers $headers -Method Get -ErrorAction Stop
} catch {
    Write-Host "Failed to fetch events: $($_.Exception.Message)"
    $events = @()
}

$plannedEvents = @()
foreach ($e in ($events | Where-Object { $_.category -eq "WORKOUT" })) {
    $eventDate = if ($e.start_date_local) { (Get-Date $e.start_date_local -Format "yyyy-MM-dd") } else { "" }
    $durationSec = Get-FirstValue @($e.moving_time, $e.workout_doc.duration)
    $distanceMeters = Get-FirstValue @($e.distance, $e.workout_doc.distance)

    $plannedEvents += [PSCustomObject]@{
        event_id          = "$($e.id)"
        external_id       = $e.external_id
        start_date_local  = $e.start_date_local
        end_date_local    = $e.end_date_local
        start_date        = $eventDate
        type              = $e.type
        name              = $e.name
        description       = $e.description
        moving_time_min   = if ($durationSec -ne $null) { [math]::Round($durationSec / 60, 1) } else { $null }
        distance_km       = if ($distanceMeters -ne $null) { [math]::Round($distanceMeters / 1000, 1) } else { $null }
        paired_activity_id = $e.paired_activity_id
    }
}

$activitiesProcessed = @()
$chatCache = @{}
$plannedUsed = New-Object System.Collections.Generic.HashSet[string]
foreach ($a in $activities) {
    $type = $a.type
    if ($type -eq "VirtualRide") { $type = "Ride" }
    if ($type -eq "VirtualRun") { $type = "Run" }

    $distance = Get-FirstValue @($a.distance, $a.icu_distance, 0)
    $movingTime = Get-FirstValue @($a.moving_time, $a.icu_recording_time, 0)
    $avgHr = Get-FirstValue @($a.average_heartrate, $a.average_hr)
    $avgWatts = Get-FirstValue @($a.icu_average_watts, $a.average_watts)
    $npWatts = Get-FirstValue @($a.icu_weighted_avg_watts, $a.weighted_avg_watts, $a.normalized_power)
    $load = Get-FirstValue @(
        $a.icu_training_load,
        $a.training_load,
        $a.power_load,
        $a.hr_load,
        $a.pace_load,
        $a.strain_score
    )
    $powerLoad = Get-FirstValue @($a.power_load, $a.icu_training_load)
    $hrLoad = Get-FirstValue @($a.hr_load)
    $paceLoad = Get-FirstValue @($a.pace_load)
    $strainScore = Get-FirstValue @($a.strain_score)

    $noteParts = @()
    if ($a.description) { $noteParts += $a.description.Trim() }
    $chatNotes = Get-ChatMessages -ChatId $a.icu_chat_id -Headers $headers -Cache $chatCache -OwnerId $ownerId
    if ($chatNotes.Count -gt 0) { $noteParts += $chatNotes }
    $notes = ($noteParts | Where-Object { $_ -and $_.Trim() -ne "" }) -join " | "

    $dateLocal = Get-FirstValue @($a.start_date_local, $a.start_date)
    $dateLocal = if ($dateLocal) { (Get-Date $dateLocal -Format "yyyy-MM-dd") } else { "" }

    $distanceKm = if ($distance) { [math]::Round($distance / 1000, 1) } else { 0 }
    $movingTimeMin = if ($movingTime) { [math]::Round($movingTime / 60, 1) } else { 0 }
    $vi = Get-FirstValue @($a.icu_variability_index, $(if ($avgWatts -and $npWatts) { $npWatts / $avgWatts }))

    $planMatch = $null
    $matchMethod = $null
    $directMatch = $plannedEvents | Where-Object { $_.paired_activity_id -eq $a.id } | Select-Object -First 1
    if ($directMatch) {
        $planMatch = $directMatch
        $matchMethod = "paired_activity_id"
        [void]$plannedUsed.Add($planMatch.event_id)
    } else {
        $fallbackMatch = $plannedEvents | Where-Object {
            $_.paired_activity_id -eq $null -and $_.start_date -eq $dateLocal -and $_.type -eq $type -and -not $plannedUsed.Contains($_.event_id)
        } | Select-Object -First 1
        if ($fallbackMatch) {
            $planMatch = $fallbackMatch
            $matchMethod = "date_type"
            [void]$plannedUsed.Add($planMatch.event_id)
        }
    }

    $planSummary = $null
    if ($planMatch) {
        $planSummary = [PSCustomObject]@{
            event_id          = $planMatch.event_id
            external_id       = $planMatch.external_id
            name              = $planMatch.name
            type              = $planMatch.type
            start_date_local  = $planMatch.start_date_local
            moving_time_min   = $planMatch.moving_time_min
            distance_km       = $planMatch.distance_km
            description       = $planMatch.description
            match_method      = $matchMethod
            delta_time_min    = if ($planMatch.moving_time_min -ne $null) { [math]::Round(($movingTimeMin - $planMatch.moving_time_min), 1) } else { $null }
            delta_distance_km = if ($planMatch.distance_km -ne $null) { [math]::Round(($distanceKm - $planMatch.distance_km), 1) } else { $null }
        }
    }

    $activitiesProcessed += [PSCustomObject]@{
        id               = $a.id
        name             = $a.name
        type             = $type
        start_date_local = $dateLocal
        distance_km      = $distanceKm
        moving_time_min  = $movingTimeMin
        average_hr       = if ($avgHr -ne $null) { [math]::Round($avgHr, 1) } else { $null }
        average_watts    = if ($avgWatts -ne $null) { [math]::Round($avgWatts, 1) } else { $null }
        suffer_score     = if ($load -ne $null) { [math]::Round($load, 1) } else { $null }
        training_load    = if ($load -ne $null) { [math]::Round($load, 1) } else { $null }
        power_load       = if ($powerLoad -ne $null) { [math]::Round($powerLoad, 1) } else { $null }
        hr_load          = if ($hrLoad -ne $null) { [math]::Round($hrLoad, 1) } else { $null }
        pace_load        = if ($paceLoad -ne $null) { [math]::Round($paceLoad, 1) } else { $null }
        strain_score     = if ($strainScore -ne $null) { [math]::Round($strainScore, 1) } else { $null }
        normalized_power = if ($npWatts -ne $null) { [math]::Round($npWatts, 0) } else { $null }
        variabilidade    = if ($vi -ne $null) { [math]::Round($vi, 2) } else { $null }
        notas            = $notes
        planejado        = $planSummary
    }
}

$prevWeight = $null
$prevHR = $null
$prevSleep = $null
$prevHRV = $null
$prevVo2Run = $null
$prevVo2Bike = $null
$wellnessFixed = @()

foreach ($w in ($wellness | Sort-Object id)) {
    $peso = if ($w.weight) { $w.weight } else { $prevWeight }
    $fc   = if ($w.restingHR) { $w.restingHR } else { $prevHR }
    $sono = if ($w.sleepSecs -gt 0) { [math]::Round($w.sleepSecs / 3600, 1) } else { $prevSleep }
    $hrv  = if ($w.hrv) { $w.hrv } else { $prevHRV }
    $vo2Run = Get-FirstValue @(
        $w.vo2MaxRun, $w.vo2maxRun, $w.vo2Run, $w.vo2max_run,
        $w.vo2Max, $w.vo2max
    )
    $vo2Bike = Get-FirstValue @(
        $w.vo2MaxBike, $w.vo2maxBike, $w.vo2Bike, $w.vo2max_bike,
        $w.vo2Max, $w.vo2max
    )
    if (-not $vo2Run) { $vo2Run = $prevVo2Run }
    if (-not $vo2Bike) { $vo2Bike = $prevVo2Bike }

    $wellnessFixed += [PSCustomObject]@{
        data      = $w.id
        peso      = $peso
        fc_reposo = $fc
        sono_h    = $sono
        hrv       = $hrv
        vo2_run   = $vo2Run
        vo2_bike  = $vo2Bike
        passos    = $w.steps
        ctl       = if ($w.ctl -ne $null) { [math]::Round($w.ctl, 1) } else { $null }
        atl       = if ($w.atl -ne $null) { [math]::Round($w.atl, 1) } else { $null }
        rampRate  = if ($w.rampRate -ne $null) { [math]::Round($w.rampRate, 1) } else { $null }
    }

    if ($peso) { $prevWeight = $peso }
    if ($fc) { $prevHR = $fc }
    if ($sono) { $prevSleep = $sono }
    if ($hrv) { $prevHRV = $hrv }
    if ($vo2Run) { $prevVo2Run = $vo2Run }
    if ($vo2Bike) { $prevVo2Bike = $vo2Bike }
}

$pesoAtual = ($wellnessFixed | Where-Object { $_.peso -gt 0 } | Sort-Object data | Select-Object -Last 1).peso
if (-not $pesoAtual) { $pesoAtual = $prevWeight }
$vo2RunAtual = ($wellnessFixed | Where-Object { $_.vo2_run -gt 0 } | Sort-Object data | Select-Object -Last 1).vo2_run
$vo2BikeAtual = ($wellnessFixed | Where-Object { $_.vo2_bike -gt 0 } | Sort-Object data | Select-Object -Last 1).vo2_bike
if (-not $vo2RunAtual) { $vo2RunAtual = $prevVo2Run }
if (-not $vo2BikeAtual) { $vo2BikeAtual = $prevVo2Bike }

function Smooth([double[]]$values) {
    if ($values.Count -le 2) { return $values }
    $smoothed = @($values[0])
    for ($i = 1; $i -lt $values.Count - 1; $i++) {
        $smoothed += [math]::Round(($values[$i - 1] + $values[$i] + $values[$i + 1]) / 3, 1)
    }
    $smoothed += $values[-1]
    return $smoothed
}

$ctlValues = @($wellnessFixed | ForEach-Object { $_.ctl })
$atlValues = @($wellnessFixed | ForEach-Object { $_.atl })
$CTLs = if ($ctlValues.Count -gt 0) { Smooth($ctlValues) } else { @() }
$ATLs = if ($atlValues.Count -gt 0) { Smooth($atlValues) } else { @() }
$TSBs = @()
for ($i = 0; $i -lt $CTLs.Count; $i++) {
    $TSBs += if ($ATLs.Count -gt $i) { [math]::Round(($CTLs[$i] - $ATLs[$i]), 1) } else { $null }
}
$rampRate = if ($CTLs.Count -gt 1) { [math]::Round(($CTLs[-1] - $CTLs[0]) / ($CTLs.Count / 7), 1) } else { $null }

$totalTSS = ($activitiesProcessed | Measure-Object -Property suffer_score -Sum).Sum
$totalTempo = ($activitiesProcessed | Measure-Object -Property moving_time_min -Sum).Sum / 60
$totalDist = ($activitiesProcessed | Measure-Object -Property distance_km -Sum).Sum

$weeklyNotes = @($events | Where-Object { $_.category -eq "NOTE" -and $_.for_week } | ForEach-Object {
    [PSCustomObject]@{
        start_date_local = $_.start_date_local
        end_date_local   = $_.end_date_local
        name             = $_.name
        description      = $_.description
    }
})

$report = [PSCustomObject]@{
    semana = @{
        inicio             = $StartDate
        fim                = $EndDate
        tempo_total_horas  = [math]::Round($totalTempo, 1)
        distancia_total_km = [math]::Round($totalDist, 1)
        carga_total_tss    = if ($totalTSS -ne $null) { [math]::Round($totalTSS, 0) } else { 0 }
    }
    metricas = @{
        peso_atual = if ($pesoAtual -ne $null) { [math]::Round($pesoAtual, 1) } else { $null }
        vo2_run   = if ($vo2RunAtual -ne $null) { [math]::Round($vo2RunAtual, 1) } else { $null }
        vo2_bike  = if ($vo2BikeAtual -ne $null) { [math]::Round($vo2BikeAtual, 1) } else { $null }
        CTL        = if ($CTLs.Count -gt 0) { $CTLs[-1] } else { $null }
        ATL        = if ($ATLs.Count -gt 0) { $ATLs[-1] } else { $null }
        TSB        = if ($TSBs.Count -gt 0) { $TSBs[-1] } else { $null }
        RampRate   = $rampRate
    }
    atividades = $activitiesProcessed
    treinos_planejados = $plannedEvents
    bem_estar  = $wellnessFixed
    notas_semana = $weeklyNotes
    analise_gpt = "Versao 7.5: inclui treinos planejados e match com atividades."
}

$fileName = "report_{0}_{1}.json" -f $StartDate, $EndDate
$outDir = Join-Path $PSScriptRoot "Relatorios_Intervals"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$jsonPath = Join-Path -Path $outDir -ChildPath $fileName
$report | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

function Write-PlannedMarkdown {
    param (
        [string]$OutputPath,
        [object[]]$Planned,
        [string]$StartDate,
        [string]$EndDate
    )

    $lines = @()
    $lines += "# Treinos Planejados ($StartDate a $EndDate)"
    $lines += ""

    if (-not $Planned -or $Planned.Count -eq 0) {
        $lines += "Nenhum treino planejado encontrado no calendario."
        Set-Content -Path $OutputPath -Value $lines -Encoding UTF8
        return
    }

    $sorted = $Planned | Sort-Object start_date_local, name
    $lines += "| Data | Tipo | Nome | Duracao (min) | Distancia (km) | External ID |"
    $lines += "|------|------|------|---------------|----------------|------------|"
    foreach ($p in $sorted) {
        $date = if ($p.start_date_local) { (Get-Date $p.start_date_local -Format "yyyy-MM-dd") } else { "" }
        $dur = if ($p.moving_time_min -ne $null) { $p.moving_time_min } else { "" }
        $dist = if ($p.distance_km -ne $null) { $p.distance_km } else { "" }
        $ext = if ($p.external_id) { $p.external_id } else { "" }
        $lines += "| $date | $($p.type) | $($p.name) | $dur | $dist | $ext |"
    }

    $lines += ""
    $lines += "## Descricoes"
    foreach ($p in $sorted) {
        $date = if ($p.start_date_local) { (Get-Date $p.start_date_local -Format "yyyy-MM-dd") } else { "" }
        $lines += "### $date - $($p.type) - $($p.name)"
        $lines += ""
        if ($p.description) {
            $lines += ($p.description -replace "`r`n", "`n") -split "`n"
        } else {
            $lines += "_Sem descricao._"
        }
        $lines += ""
        $lines += "- event_id: $($p.event_id)"
        if ($p.external_id) { $lines += "- external_id: $($p.external_id)" }
        $lines += ""
    }

    Set-Content -Path $OutputPath -Value $lines -Encoding UTF8
}

$plannedFileName = "planned_{0}_{1}.md" -f $StartDate, $EndDate
$plannedPath = Join-Path -Path $outDir -ChildPath $plannedFileName
Write-PlannedMarkdown -OutputPath $plannedPath -Planned $plannedEvents -StartDate $StartDate -EndDate $EndDate

Write-Host ""
Write-Host "Report exported: $jsonPath"
Write-Host "Planned workouts exported: $plannedPath"
