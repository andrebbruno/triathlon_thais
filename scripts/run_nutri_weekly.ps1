# run_nutri_weekly.ps1
# Local automation for MFP + nutrition report (weekly).

param(
  [string]$AthleteId = "0",
  [string]$ApiKeyPath = "$PSScriptRoot\..\api_key.txt",
  [int]$ChromeWaitSeconds = 20
)

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")

function Get-WeekRange([datetime]$date) {
  $dayOfWeek = [int]$date.DayOfWeek
  if ($dayOfWeek -eq 0) { $monday = $date.AddDays(-6) } else { $monday = $date.AddDays(-($dayOfWeek - 1)) }
  $sunday = $monday.AddDays(6)
  return @{
    Start = $monday.ToString("yyyy-MM-dd")
    End   = $sunday.ToString("yyyy-MM-dd")
  }
}

$today = Get-Date
$mfpUser = $env:MFP_USERNAME

if ($AthleteId -eq "0") {
  Write-Host "Defina o AthleteId (Intervals) ao rodar este script."
  exit 1
}
if (-not $mfpUser) {
  Write-Host "Defina MFP_USERNAME no ambiente antes de rodar."
  exit 1
}
$current = Get-WeekRange -date $today
$next = Get-WeekRange -date ($today.AddDays(7))

Write-Host "Intervals current week: $($current.Start) to $($current.End)"
& "$repoRoot\export-intervals-week-com-notas.ps1" -AthleteId $AthleteId -ApiKeyPath $ApiKeyPath -StartDate $current.Start -EndDate $current.End

Write-Host "Intervals next week: $($next.Start) to $($next.End)"
& "$repoRoot\export-intervals-week-com-notas.ps1" -AthleteId $AthleteId -ApiKeyPath $ApiKeyPath -StartDate $next.Start -EndDate $next.End

Write-Host "Running MFP (Chrome) for last 7 days..."
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
if (-not (Test-Path $chromePath)) {
  $chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
}
if (-not (Test-Path $chromePath)) {
  Write-Host "Chrome not found. Install Google Chrome or update the path."
  exit 1
}

Start-Process $chromePath -ArgumentList @(
  "--remote-debugging-port=9223",
  "--user-data-dir=$env:LOCALAPPDATA\Google\Chrome\User Data MFP",
  "--profile-directory=Default",
  "https://www.myfitnesspal.com/food/diary/$mfpUser"
)
Start-Sleep -Seconds $ChromeWaitSeconds
& node "$repoRoot\scripts\mfp_connect_chrome.js" --days 7 --close

Write-Host "Generating nutrition report..."
& node "$repoRoot\scripts\nutri_report.js"

Write-Host "Done."
