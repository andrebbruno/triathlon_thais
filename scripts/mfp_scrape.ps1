# MyFitnessPal diary scraper (password-protected public diary)

param(
  [string]$Username = '',
  [string]$Date = '',
  [string]$Out = '',
  [switch]$Headless
)

if (-not $Username) {
  Write-Host 'Missing -Username. Ex: -Username "seu_usuario".'
  exit 1
}
if (-not $env:MFP_DIARY_PASSWORD) {
  Write-Host 'Missing MFP_DIARY_PASSWORD env var.'
  exit 1
}

$script = Join-Path $PSScriptRoot 'mfp_scrape.mjs'
$args = @('--username', $Username)
if ($Date) { $args += @('--date', $Date) }
if ($Out) { $args += @('--out', $Out) }
if ($Headless) { $args += @('--headless') }

node $script @args
