param(
  [Parameter(Mandatory = $true)]
  [ValidateSet('head','coachTri','IT_tech','nutricao','fisio')]
  [string]$Agent,
  [string]$Summary = ''
)

$map = @{
  head     = 'HEAD'
  coachTri = 'COACH_TRI'
  IT_tech  = 'IT_TECH'
  nutricao = 'NUTRICAO'
  fisio    = 'FISIO'
}

$label = $map[$Agent]
if (-not $label) {
  Write-Host "Unknown agent: $Agent"
  exit 1
}

if ([string]::IsNullOrWhiteSpace($Summary)) {
  $Summary = Read-Host 'Session summary (short)'
}

if ([string]::IsNullOrWhiteSpace($Summary)) {
  Write-Host 'Summary is required.'
  exit 1
}

$root = Split-Path -Parent $PSScriptRoot
$sessionPath = Join-Path $root ("sessions\{0}.md" -f $label)

$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
$entry = "`n## $timestamp`n$Summary`n"

Add-Content -Path $sessionPath -Encoding ASCII -Value $entry
Write-Host "Saved to $sessionPath"
