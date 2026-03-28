$repoZipUrl = "https://github.com/onlinellx551-debug/shengtu/archive/refs/heads/main.zip"
$userHome = [Environment]::GetFolderPath("UserProfile")
$sourceRoot = Join-Path $userHome ".codex-plugin-sources"
$sourceRepo = Join-Path $sourceRoot "shengtu"
$tempRoot = Join-Path $env:TEMP "jp-menswear-plugin-install"
$zipPath = Join-Path $tempRoot "shengtu-main.zip"
$extractRoot = Join-Path $tempRoot "extract"
$extractedRepo = Join-Path $extractRoot "shengtu-main"

New-Item -ItemType Directory -Force -Path $sourceRoot | Out-Null
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

if (Test-Path $zipPath) {
  Remove-Item -Force $zipPath
}

if (Test-Path $extractRoot) {
  Remove-Item -Recurse -Force $extractRoot
}

Write-Output "Downloading latest repository archive..."
Invoke-WebRequest -Uri $repoZipUrl -OutFile $zipPath

Write-Output "Extracting repository archive..."
Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force

if (-not (Test-Path $extractedRepo)) {
  throw "Extracted repository not found: $extractedRepo"
}

if (Test-Path $sourceRepo) {
  Remove-Item -Recurse -Force $sourceRepo
}

Move-Item -LiteralPath $extractedRepo -Destination $sourceRepo

$syncScript = Join-Path $sourceRepo "sync_jp_menswear_plugin.ps1"
if (-not (Test-Path $syncScript)) {
  throw "Sync script not found: $syncScript"
}

Write-Output "Installing or updating plugin for current user..."
powershell -ExecutionPolicy Bypass -File $syncScript

Write-Output "Done. Restart Codex, then search Plugins for: JP Menswear 6-Step Workflow"
