$ErrorActionPreference = "Stop"

function Get-CloudflaredPath {
  $wingetPath = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter "cloudflared.exe" -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName
  if ($wingetPath) { return $wingetPath }

  $cmd = Get-Command cloudflared -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }

  throw "找不到 cloudflared，請先安裝：winget install --id Cloudflare.cloudflared -e"
}

$cf = Get-CloudflaredPath
& $cf tunnel --config "$env:USERPROFILE\.cloudflared\config.yml" run
