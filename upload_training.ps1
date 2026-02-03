# upload_training.ps1 - Intervals.icu Calendar Events (bulk upsert)

param(
  [string]$IntervalsApiKey = "",
  [string]$ApiKeyPath = "$PSScriptRoot\api_key.txt",
  [string]$TrainingsFile = "trainings.json",
  [string]$StartTimeLocal = "",
  [switch]$WriteBackNormalized
)

$logPath = Join-Path $PSScriptRoot "upload_training.log.jsonl"

function Write-Log {
  param(
    [string]$Level,
    [string]$Message,
    [hashtable]$Data
  )

  $entry = [ordered]@{
    timestamp = (Get-Date).ToString("s")
    level     = $Level
    message   = $Message
    data      = $Data
  }

  ($entry | ConvertTo-Json -Depth 6 -Compress) | Add-Content -Path $logPath -Encoding UTF8
}

function Resolve-ApiKey {
  param(
    [string]$ProvidedKey,
    [string]$Path
  )

  if ($ProvidedKey) { return $ProvidedKey }
  if ($env:INTERVALS_API_KEY) { return $env:INTERVALS_API_KEY }
  if (Test-Path $Path) { return (Get-Content $Path -Raw).Trim() }
  $localPath = $null
  if ($env:USERPROFILE) { $localPath = Join-Path $env:USERPROFILE ".intervals\\api_key.txt" }
  elseif ($env:HOME) { $localPath = Join-Path $env:HOME ".intervals\\api_key.txt" }
  if ($localPath -and (Test-Path $localPath)) { return (Get-Content $localPath -Raw).Trim() }
  return ""
}

function Normalize-Description {
  param(
    [string]$Description,
    [string]$Type
  )

  if (-not $Description) { return "" }
  $normalized = ($Description -replace "`r`n", "`n").Trim()
  $lines = $normalized -split "`n"
  $output = @()

  foreach ($line in $lines) {
    $current = $line.TrimEnd()

    if ($Type -eq "Swim") {
      if ($current -match "^\s*-\s") {
        $current = $current -replace "(\d+)\s+meters", '$1meters'
        if ($current -match "\bmeters\b" -and $current -notmatch "\bpace\b") {
          $current = $current.TrimEnd() + " pace"
        }
        $current = $current -replace "\bPace\b", "pace"
      }
    } else {
      if ($current -match "^\s*-\s") {
        $current = $current -replace "\s+in\s+", " "
      }
      $current = $current -replace "\bPace\b", "pace"
    }

    $output += $current
  }

  return ($output -join "`n").Trim()
}

function Validate-Event {
  param(
    [hashtable]$Event
  )

  $errors = @()
  $warnings = @()

  $required = @("external_id", "category", "start_date_local", "type", "name", "description")
  foreach ($field in $required) {
    if (-not $Event.ContainsKey($field) -or [string]::IsNullOrWhiteSpace([string]$Event[$field])) {
      $errors += "Campo obrigatorio ausente: $field"
    }
  }

  $validTypes = @("Ride", "Run", "Swim", "WeightTraining")
  if ($Event.type -and ($validTypes -notcontains $Event.type)) {
    $errors += "Tipo invalido: $($Event.type)"
  }

  if ($Event.start_date_local) {
    try { [void][DateTime]::Parse($Event.start_date_local) }
    catch { $errors += "start_date_local invalido: $($Event.start_date_local)" }
  }

  if ($Event.type -eq "Swim") {
    if ($Event.description -match "\b\d+m\b") {
      $errors += "Natacao com 'm' detectado (minutos). Use 'meters'."
    }
    if ($Event.description -notmatch "\d+meters") {
      $warnings += "Natacao sem 'meters' encontrado no texto."
    }
    if ($Event.description -notmatch "\bpace\b") {
      $warnings += "Natacao sem 'pace' encontrado no texto."
    }
  }

  return @{
    Errors = $errors
    Warnings = $warnings
  }
}

function Apply-StartTime {
  param(
    [string]$StartDateLocal,
    [string]$StartTime
  )

  try {
    $date = [DateTime]$StartDateLocal
    if (-not [string]::IsNullOrWhiteSpace($StartTime)) {
      $time = [TimeSpan]::Parse($StartTime)
      $date = $date.Date + $time
    }
    return $date.ToString("yyyy-MM-ddTHH:mm:ss")
  } catch {
    return $StartDateLocal
  }
}

$IntervalsApiKey = Resolve-ApiKey -ProvidedKey $IntervalsApiKey -Path $ApiKeyPath
if (-not $IntervalsApiKey) {
  Write-Host "API key not provided."
  Write-Log -Level "error" -Message "API key missing" -Data @{ file = $ApiKeyPath }
  exit 1
}

if (-not (Test-Path $TrainingsFile)) {
  Write-Host "Arquivo trainings.json nao encontrado."
  Write-Log -Level "error" -Message "Arquivo nao encontrado" -Data @{ file = $TrainingsFile }
  exit 1
}

$pair   = "API_KEY:$IntervalsApiKey"
$base64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
$headers = @{
  Authorization = "Basic $base64"
  Accept        = "application/json"
  "Content-Type"= "application/json; charset=utf-8"
}

$uri = "https://intervals.icu/api/v1/athlete/0/events/bulk?upsert=true"

$rawEvents = Get-Content $TrainingsFile -Raw | ConvertFrom-Json
$events = @($rawEvents)

$normalizedEvents = @()
$hasChanges = $false
$allErrors = @()

foreach ($event in $events) {
  $eventHash = @{}
  $event.PSObject.Properties | ForEach-Object { $eventHash[$_.Name] = $_.Value }

  $eventHash.description = Normalize-Description -Description $eventHash.description -Type $eventHash.type
  $eventHash.start_date_local = Apply-StartTime -StartDateLocal $eventHash.start_date_local -StartTime $StartTimeLocal

  $validation = Validate-Event -Event $eventHash
  foreach ($warning in $validation.Warnings) {
    Write-Host "Aviso: $warning (external_id=$($eventHash.external_id))"
  }
  foreach ($error in $validation.Errors) {
    $allErrors += "$error (external_id=$($eventHash.external_id))"
  }

  if ($eventHash.description -ne $event.description -or $eventHash.start_date_local -ne $event.start_date_local) {
    $hasChanges = $true
  }

  $normalizedEvents += [PSCustomObject]$eventHash
}

if ($allErrors.Count -gt 0) {
  $allErrors | ForEach-Object { Write-Host "Erro: $_" }
  Write-Log -Level "error" -Message "Validacao falhou" -Data @{ errors = $allErrors }
  exit 1
}

if ($WriteBackNormalized -and $hasChanges) {
  $normalizedEvents | ConvertTo-Json -Depth 12 | Out-File -FilePath $TrainingsFile -Encoding UTF8
  Write-Host "trainings.json normalizado e regravado."
  Write-Log -Level "info" -Message "Arquivo normalizado" -Data @{ file = $TrainingsFile }
}

$supportsSkip = $false
try {
  $supportsSkip = (Get-Command Invoke-WebRequest).Parameters.ContainsKey("SkipHttpErrorCheck")
} catch { }

function Invoke-IntervalsUpload {
  param(
    [string]$Body
  )

  $bodyBytes = [Text.Encoding]::UTF8.GetBytes($Body)

  if ($PSVersionTable.PSVersion.Major -lt 6) {
    try {
      [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch { }
  }

  if ($supportsSkip) {
    $response = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -Body $bodyBytes -SkipHttpErrorCheck
    return @{
      StatusCode = [int]$response.StatusCode
      Reason     = $response.StatusDescription
      Content    = $response.Content
      Error      = $null
    }
  }

  try {
    $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $bodyBytes
    $content = $null
    try { $content = $response | ConvertTo-Json -Depth 12 } catch { }
    return @{
      StatusCode = 200
      Reason     = "OK"
      Content    = $content
      Error      = $null
    }
  } catch {
    $errorBody = $null
    $statusCode = $null
    $reasonPhrase = $null
    try {
      if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $reasonPhrase = $_.Exception.Response.StatusDescription
        if ($_.Exception.Response.GetResponseStream()) {
          $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
          $errorBody = $reader.ReadToEnd()
        }
      }
    } catch { }

    if (-not $errorBody -and $_.ErrorDetails -and $_.ErrorDetails.Message) {
      $errorBody = $_.ErrorDetails.Message
    }

    return @{
      StatusCode = $statusCode
      Reason     = $reasonPhrase
      Content    = $errorBody
      Error      = $_.Exception.Message
    }
  }
}

try {
  $body = ($normalizedEvents | ConvertTo-Json -Depth 12)
  if ($body.TrimStart().StartsWith("{")) { $body = "[$body]" }

  $bulkResult = Invoke-IntervalsUpload -Body $body
  $statusCode = $bulkResult.StatusCode
  $statusDescription = $bulkResult.Reason
  $content = $bulkResult.Content

  $bulkFailed = (-not $statusCode) -or ($statusCode -lt 200 -or $statusCode -ge 300)
  if ($bulkFailed) {
    Write-Host "Failed to upload trainings (bulk):"
    if ($statusCode) {
      Write-Host "HTTP status: $statusCode $statusDescription"
    } else {
      Write-Host $bulkResult.Error
    }
    if ($content) {
      Write-Host "Server response body:"
      Write-Host $content
    }
    Write-Log -Level "error" -Message "Upload falhou (bulk)" -Data @{
      status = $statusCode
      reason = $statusDescription
      body   = $content
      file   = $TrainingsFile
    }

    Write-Host "Tentando upload individual para identificar erro..."
    $singleFailures = @()
    foreach ($event in $normalizedEvents) {
      $singleBody = ($event | ConvertTo-Json -Depth 12)
      if ($singleBody.TrimStart().StartsWith("{")) { $singleBody = "[$singleBody]" }
      $singleResult = Invoke-IntervalsUpload -Body $singleBody
      $singleStatus = $singleResult.StatusCode
      $singleReason = $singleResult.Reason
      $singleContent = $singleResult.Content

      $singleFailed = (-not $singleStatus) -or ($singleStatus -lt 200 -or $singleStatus -ge 300)
      if ($singleFailed) {
        $failure = [ordered]@{
          external_id = $event.external_id
          status      = $singleStatus
          reason      = $singleReason
          body        = $singleContent
        }
        $singleFailures += $failure
        Write-Host "Falha no evento: $($event.external_id)"
        if ($singleStatus) {
          Write-Host "HTTP status: $singleStatus $singleReason"
        }
        if ($singleContent) {
          Write-Host "Server response body:"
          Write-Host $singleContent
        }
      }
    }

    if ($singleFailures.Count -gt 0) {
      Write-Log -Level "error" -Message "Upload falhou (single)" -Data @{ failures = $singleFailures }
      exit 1
    }

    Write-Host "Bulk falhou, mas todos os uploads individuais foram OK."
    Write-Log -Level "warn" -Message "Bulk falhou, single OK" -Data @{ file = $TrainingsFile }
    $content = $null
  }

  $parsed = $null
  $count = $null
  if ($content) {
    try { $parsed = $content | ConvertFrom-Json } catch { }
  }
  if ($parsed -ne $null) { $count = $parsed.Count }

  if ($count -ne $null) {
    Write-Host "Upload OK. Eventos criados/atualizados: $count"
  } else {
    Write-Host "Upload OK."
  }
  Write-Log -Level "info" -Message "Upload OK" -Data @{ count = $count; file = $TrainingsFile }
}
catch {
  Write-Host "Failed to upload trainings:"
  Write-Host $_.Exception.Message
  Write-Log -Level "error" -Message "Upload falhou" -Data @{ message = $_.Exception.Message }
  exit 1
}
