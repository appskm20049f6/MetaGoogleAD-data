param(
  [Parameter(Mandatory = $true)]
  [string]$Hostname,
  [string]$TunnelName = "metagooglead-data",
  [string]$OriginUrl = "http://localhost:3000"
)

$ErrorActionPreference = "Stop"

function Get-CloudflaredPath {
  $wingetPath = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter "cloudflared.exe" -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName
  if ($wingetPath) { return $wingetPath }

  $cmd = Get-Command cloudflared -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }

  throw "cloudflared not found. Install with: winget install --id Cloudflare.cloudflared -e"
}

$cf = Get-CloudflaredPath
Write-Host "Using cloudflared: $cf"

$cloudflaredDir = Join-Path $env:USERPROFILE ".cloudflared"
if (-not (Test-Path $cloudflaredDir)) {
  New-Item -ItemType Directory -Path $cloudflaredDir | Out-Null
}

$certPath = Join-Path $cloudflaredDir "cert.pem"
if (-not (Test-Path $certPath)) {
  Write-Host "No cert.pem found. Starting Cloudflare login..."
  & $cf tunnel login
  if (-not (Test-Path $certPath)) {
    throw "Login not completed. Finish browser auth and run setup again."
  }
}

$tunnelId = $null
try {
  $listJson = & $cf tunnel list --output json
  $list = $listJson | ConvertFrom-Json
  $existing = $list | Where-Object { $_.name -eq $TunnelName } | Select-Object -First 1
  if ($existing) {
    $tunnelId = $existing.id
    Write-Host "Found existing tunnel: $TunnelName ($tunnelId)"
  }
} catch {
  Write-Host "Failed to list tunnels. Will try create tunnel."
}

if (-not $tunnelId) {
  $createOutput = & $cf tunnel create $TunnelName
  $match = [regex]::Match(($createOutput -join "`n"), "([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})")
  if (-not $match.Success) {
    throw "Tunnel created but ID not found. Run: cloudflared tunnel list"
  }
  $tunnelId = $match.Groups[1].Value
  Write-Host "Created tunnel: $TunnelName ($tunnelId)"
}

$credentialPath = Join-Path $cloudflaredDir ("$tunnelId.json")
if (-not (Test-Path $credentialPath)) {
  throw "Credential file missing: $credentialPath"
}

$configPath = Join-Path $cloudflaredDir "config.yml"
$config = @"
tunnel: $tunnelId
credentials-file: $credentialPath

ingress:
  - hostname: $Hostname
    service: $OriginUrl
  - service: http_status:404
"@
$config | Set-Content -Path $configPath -Encoding UTF8

& $cf tunnel route dns $TunnelName $Hostname

Write-Host "Fixed tunnel setup complete."
Write-Host "Hostname: $Hostname"
Write-Host "Config:   $configPath"
Write-Host "Start command: npm run share:fixed"
