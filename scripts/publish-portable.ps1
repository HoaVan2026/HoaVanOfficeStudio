param(
  [string]$Runtime = "win-x64"
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root "src\HoaVan.OfficeStudio\HoaVan.OfficeStudio.csproj"
$out = Join-Path $root "dist\HoaVanOfficeStudio-$Runtime"
Write-Host "Publishing Hoa Van Office Studio V1..." -ForegroundColor Cyan
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
  throw ".NET 10 SDK was not found. Install it from https://dotnet.microsoft.com/ then run this script again."
}
Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
New-Item $out -ItemType Directory -Force | Out-Null

dotnet publish $project -c Release -r $Runtime --self-contained true -p:PublishSingleFile=true -o $out

$engine = Join-Path $root "tools\officecli-win-x64.exe"
if (Test-Path $engine) {
  $toolDir = Join-Path $out "tools"
  New-Item $toolDir -ItemType Directory -Force | Out-Null
  Copy-Item $engine (Join-Path $toolDir "officecli-win-x64.exe") -Force
  Write-Host "Bundled OfficeCLI engine." -ForegroundColor Green
} else {
  Write-Warning "tools\officecli-win-x64.exe is missing. Copy the OfficeCLI Windows x64 binary there before distributing the app."
}
Copy-Item (Join-Path $root "NOTICE-OFFICECLI.txt") (Join-Path $out "NOTICE-OFFICECLI.txt") -Force
Write-Host "Done: $out" -ForegroundColor Green
