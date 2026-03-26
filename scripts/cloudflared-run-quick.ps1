$ErrorActionPreference = 'Stop'

$cf = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter "cloudflared.exe" -ErrorAction SilentlyContinue |
  Select-Object -First 1 -ExpandProperty FullName

if (-not $cf) {
  throw "cloudflared.exe not found. Please install it first."
}

Write-Host "Using cloudflared: $cf"
Write-Host "Starting quick tunnel for http://localhost:3000 ..."
Write-Host "Keep this terminal open. Closing it will stop the share URL."

$quickUrl = $null

$oldPreference = $ErrorActionPreference
$ErrorActionPreference = 'Continue'

try {
  & $cf --config NUL tunnel --url http://localhost:3000 --no-autoupdate 2>&1 | ForEach-Object {
    $line = $_.ToString()
    Write-Host $line

    if (-not $quickUrl) {
      $m = [regex]::Match($line, 'https://[a-z0-9-]+\.trycloudflare\.com')
      if ($m.Success) {
        $quickUrl = $m.Value
        Write-Host "URL: $quickUrl"
      }
    }
  }
}
finally {
  $ErrorActionPreference = $oldPreference
}

if ($quickUrl) {
  Write-Host "URL: $quickUrl"
}
