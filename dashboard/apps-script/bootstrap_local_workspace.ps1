param(
  [string]$TargetPath = "$env:USERPROFILE\\Desktop\\AMORA_AppsScript",
  [switch]$InstallDependencies
)

$sourcePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$filesToCopy = @(
  "Code.gs",
  "appsscript.json",
  "package.json",
  ".claspignore",
  "README.md"
)

New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null

foreach ($file in $filesToCopy) {
  Copy-Item -LiteralPath (Join-Path $sourcePath $file) -Destination (Join-Path $TargetPath $file) -Force
}

if ($InstallDependencies) {
  Push-Location $TargetPath
  try {
    npm install
  } finally {
    Pop-Location
  }
}

Write-Host ""
Write-Host "Workspace local do Apps Script preparado em:" -ForegroundColor Cyan
Write-Host $TargetPath -ForegroundColor Green
Write-Host ""
Write-Host "Proximos passos:" -ForegroundColor Cyan
Write-Host "1. cd `"$TargetPath`""
Write-Host "2. npm install"
Write-Host "3. npx clasp login"
Write-Host "4. npx clasp create --type standalone --title `"AMORA Dashboard API`""
Write-Host "5. npx clasp push"
Write-Host "6. Publicar como Web App no Apps Script"
