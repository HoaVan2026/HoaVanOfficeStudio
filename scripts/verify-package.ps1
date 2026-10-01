$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dist = Get-ChildItem (Join-Path $root "dist") -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $dist) { throw "No dist folder. Run publish-portable.ps1 first." }
$app = Join-Path $dist.FullName "HoaVan.OfficeStudio.exe"
$engine = Join-Path $dist.FullName "tools\officecli-win-x64.exe"
Write-Host "App:    $(Test-Path $app) $app"
Write-Host "Engine: $(Test-Path $engine) $engine"
if (-not (Test-Path $app)) { throw "App executable missing." }
if (-not (Test-Path $engine)) { Write-Warning "OfficeCLI engine missing." }
