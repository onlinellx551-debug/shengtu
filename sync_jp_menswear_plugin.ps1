$repoRoot = $PSScriptRoot
$repoPlugin = Join-Path $repoRoot "plugins\jp-menswear-workflow"
$userHome = [Environment]::GetFolderPath("UserProfile")
$homePluginRoot = Join-Path $userHome "plugins"
$homePlugin = Join-Path $homePluginRoot "jp-menswear-workflow"
$homeMarketplaceDir = Join-Path $userHome ".agents\plugins"
$homeMarketplace = Join-Path $homeMarketplaceDir "marketplace.json"

if (-not (Test-Path $repoPlugin)) {
  throw "Repository plugin source not found: $repoPlugin"
}

New-Item -ItemType Directory -Force -Path $homePluginRoot | Out-Null
New-Item -ItemType Directory -Force -Path $homeMarketplaceDir | Out-Null

if (Test-Path $homePlugin) {
  Remove-Item -Recurse -Force $homePlugin
}

Copy-Item -Recurse -Force $repoPlugin $homePlugin

@'
{
  "name": "local-user-marketplace",
  "interface": {
    "displayName": "Local User Plugins"
  },
  "plugins": [
    {
      "name": "jp-menswear-workflow",
      "source": {
        "source": "local",
        "path": "./plugins/jp-menswear-workflow"
      },
      "policy": {
        "installation": "AVAILABLE",
        "authentication": "ON_INSTALL"
      },
      "category": "Productivity"
    }
  ]
}
'@ | Set-Content -Encoding utf8 -LiteralPath $homeMarketplace

Write-Output "Repository root: $repoRoot"
Write-Output "Synced plugin to: $homePlugin"
Write-Output "Updated marketplace: $homeMarketplace"
