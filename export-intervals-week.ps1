# ==============================================
# RELATÓRIO SEMANAL COMPLETO - VERSÃO 6.0
# Integra Strava + Intervals.icu
# Inclui CTL, ATL, TSB, RampRate, NP estimado, e métricas de bem-estar persistentes
# ==============================================

# === CONFIGURAÇÕES ===
$config = @{
    strava_client_id     = "188936"
    strava_client_secret = "c8e9bc5d1d2b574d6f4f11e6fbd54bcbc5e86b25"
    strava_refresh_token = "aa8213479e2d282af61148ec66f43d1abbf0312c"
    intervals_api_key    = "3o68ipgae5ndvi5u445tvqd90"
    token_file           = "strava_token.json"
}

# === FUNÇÃO: ATUALIZAR TOKEN STRAVA ===
function Get-StravaToken {
    param (
        [string]$clientId, [string]$clientSecret, [string]$refreshToken, [string]$tokenFile
    )

    if (Test-Path $tokenFile) {
        $saved = Get-Content $tokenFile | ConvertFrom-Json
        $now = [int][double]::Parse((Get-Date -UFormat %s))
        if ($saved.expires_at -gt $now) {
            Write-Host "Usando token Strava salvo (válido até $([DateTimeOffset]::FromUnixTimeSeconds($saved.expires_at).DateTime))"
            return $saved.access_token
        }
        Write-Host "Token expirado, renovando..."
    }

    try {
        $body = @{
            client_id     = $clientId
            client_secret = $clientSecret
            refresh_token = $refreshToken
            grant_type    = "refresh_token"
        }
        $response = Invoke-RestMethod -Uri "https://www.strava.com/oauth/token" -Method Post -Body $body -ErrorAction Stop
        $response | ConvertTo-Json | Out-File -FilePath $tokenFile -Encoding UTF8
        Write-Host "Novo token Strava obtido e salvo."
        return $response.access_token
    } catch {
        Write-Host "Erro ao renovar token Strava: $($_.Exception.Message)"
        exit
    }
}

# === OBTÉM TOKEN STRAVA ===
$access_token = Get-StravaToken -clientId $config.strava_client_id -clientSecret $config.strava_client_secret -refreshToken $config.strava_refresh_token -tokenFile $config.token_file

# === SEMANA ATUAL ===
$today  = Get-Date
$dayOfWeek = [int]$today.DayOfWeek
if ($dayOfWeek -eq 0) { $monday = $today.AddDays(-6) } else { $monday = $today.AddDays(-($dayOfWeek - 1)) }
$sunday = $monday.AddDays(6)
$oldest = $monday.ToString("yyyy-MM-dd")
$newest = $sunday.ToString("yyyy-MM-dd")

$athleteId = if ($env:INTERVALS_ATHLETE_ID) { $env:INTERVALS_ATHLETE_ID } else { "0" }

$after  = [int][double]::Parse((Get-Date "$oldest 00:00:00" -UFormat %s))
$before = [int][double]::Parse((Get-Date "$newest 23:59:59" -UFormat %s))

Write-Host "Semana atual: $oldest até $newest"

# === CABEÇALHOS ===
$headersStrava = @{ "Authorization" = "Bearer $access_token"; "Accept" = "application/json" }
$pair = "API_KEY:$($config.intervals_api_key)"
$base64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))
$headersIntervals = @{ "Authorization" = "Basic $base64"; "Accept" = "application/json" }

# === CONSULTAR ATIVIDADES STRAVA ===
$uriActivities = "https://www.strava.com/api/v3/athlete/activities?after=$after&before=$before&per_page=100"
Write-Host "Consultando atividades Strava..."
try {
    $activities = Invoke-RestMethod -Uri $uriActivities -Headers $headersStrava -Method Get -ErrorAction Stop
} catch {
    Write-Host "Erro ao consultar atividades: $($_.Exception.Message)"
    exit
}
Write-Host "Atividades carregadas: $($activities.Count)"

# === PROCESSAR ATIVIDADES ===
$activitiesProcessed = @()
foreach ($a in $activities) {
    $avgWatts = $a.average_watts
    $np = $a.normalized_power
    if (-not $np -and $avgWatts -gt 0) { $np = [math]::Round($avgWatts * 1.08, 0) } # NP estimado
    $vi = if ($avgWatts -gt 0 -and $np) { [math]::Round($np / $avgWatts, 2) } else { $null }

    $activitiesProcessed += [PSCustomObject]@{
        id               = $a.id
        name             = $a.name
        type             = $a.type
        start_date_local = (Get-Date $a.start_date_local -Format "yyyy-MM-dd")
        distance_km      = [math]::Round($a.distance / 1000, 1)
        moving_time_min  = [math]::Round($a.moving_time / 60, 1)
        average_hr       = $a.average_heartrate
        average_watts    = $avgWatts
        suffer_score     = $a.suffer_score
        normalized_power = $np
        variabilidade    = $vi
    }
}

# === WELLNESS INTERVALS ===
$uriWellness = "https://intervals.icu/api/v1/athlete/$athleteId/wellness?oldest=$oldest&newest=$newest"
Write-Host "Consultando dados de bem-estar..."
try {
    $wellness = Invoke-RestMethod -Uri $uriWellness -Headers $headersIntervals -Method Get -ErrorAction Stop
    Write-Host "Dados de bem-estar carregados: $($wellness.Count)"
} catch {
    Write-Host "Erro ao buscar wellness: $($_.Exception.Message)"
    $wellness = @()
}

# === PERSISTÊNCIA DE BEM-ESTAR ===
$prevWeight = $null; $prevHR = $null; $prevSleep = $null; $prevHRV = $null
$prevVo2Run = $null; $prevVo2Bike = $null
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
        ctl       = [math]::Round($w.ctl, 1)
        atl       = [math]::Round($w.atl, 1)
        rampRate  = [math]::Round($w.rampRate, 1)
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

# === SUAVIZAÇÃO DE CTL / ATL ===
function Smooth([double[]]$values) {
    if ($values.Count -le 2) { return $values }
    $smoothed = @($values[0])
    for ($i=1; $i -lt $values.Count-1; $i++) {
        $smoothed += [math]::Round(($values[$i-1] + $values[$i] + $values[$i+1]) / 3, 1)
    }
    $smoothed += $values[-1]
    return $smoothed
}
$CTLs = Smooth(($wellnessFixed | ForEach-Object { $_.ctl }))
$ATLs = Smooth(($wellnessFixed | ForEach-Object { $_.atl }))
$TSBs = for ($i=0; $i -lt $CTLs.Count; $i++) { [math]::Round(($CTLs[$i] - $ATLs[$i]), 1) }
$rampRate = [math]::Round(($CTLs[-1] - $CTLs[0]) / ($CTLs.Count / 7), 1)

# === RELATÓRIO FINAL ===
$totalTSS = ($activitiesProcessed | Measure-Object -Property suffer_score -Sum).Sum
$totalTempo = ($activitiesProcessed | Measure-Object -Property moving_time_min -Sum).Sum / 60
$totalDist = ($activitiesProcessed | Measure-Object -Property distance_km -Sum).Sum

$report = [PSCustomObject]@{
    semana = @{
        inicio             = $oldest
        fim                = $newest
        tempo_total_horas  = [math]::Round($totalTempo, 1)
        distancia_total_km = [math]::Round($totalDist, 1)
        carga_total_tss    = [math]::Round($totalTSS, 0)
    }
    metricas = @{
        peso_atual = [math]::Round($pesoAtual, 1)
        vo2_run   = if ($vo2RunAtual -ne $null) { [math]::Round($vo2RunAtual, 1) } else { $null }
        vo2_bike  = if ($vo2BikeAtual -ne $null) { [math]::Round($vo2BikeAtual, 1) } else { $null }
        CTL        = $CTLs[-1]
        ATL        = $ATLs[-1]
        TSB        = $TSBs[-1]
        RampRate   = $rampRate
    }
    atividades = $activitiesProcessed
    bem_estar  = $wellnessFixed
    analise_gpt = "Versão 6.0: inclui CTL/ATL/TSB suavizados, NP estimado e persistência total de dados fisiológicos."
}

# === EXPORTAR JSON ===
$fileName = "report_{0}_{1}.json" -f $oldest, $newest
$jsonPath = Join-Path -Path (Get-Location) -ChildPath $fileName
$report | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

Write-Host ""
Write-Host "Relatório semanal completo exportado:"
Write-Host "   $jsonPath"
Write-Host "Inclui TSB, NP estimado, CTL/ATL suavizados e peso atual persistente."
