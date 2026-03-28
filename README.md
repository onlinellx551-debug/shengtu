# task-codex

This repository is the workspace for the Japanese menswear workflow project.

## What is included

- local bridge rules for the workflow skills
- the plugin source for the mixed delivery model
- plugin sync utilities
- business-facing prompt templates

## Main plugin source

- `plugins/jp-menswear-workflow`

## Install the plugin for the current user

Anyone on the team can:
1. run the one-click installer below
2. restart Codex
3. search in `Plugins` for `JP Menswear 6-Step Workflow`

## One-click team installer

```powershell
powershell -ExecutionPolicy Bypass -File ".\install_or_update_team_plugin.ps1"
```

This script:
- downloads the latest `main` branch from GitHub
- stores a local source copy under the current user's home directory
- syncs the plugin into the current user's local Codex plugin directory
- updates the current user's local marketplace entry

## Sync plugin into the current user's local Codex plugin directory

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Administrator\Desktop\Codex Project\task-codex\sync_jp_menswear_plugin.ps1"
```

The script installs into the current Windows user's home directory automatically.

## Installed local plugin path

- `C:\Users\Administrator\plugins\jp-menswear-workflow`

## Installed local marketplace path

- `C:\Users\Administrator\.agents\plugins\marketplace.json`
