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
1. clone this repository
2. open the repo root
3. run the sync script below
4. restart Codex
5. search in `Plugins` for `JP Menswear 6-Step Workflow`

## Sync plugin into the current user's local Codex plugin directory

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Administrator\Desktop\Codex Project\task-codex\sync_jp_menswear_plugin.ps1"
```

The script installs into the current Windows user's home directory automatically.

## Installed local plugin path

- `C:\Users\Administrator\plugins\jp-menswear-workflow`

## Installed local marketplace path

- `C:\Users\Administrator\.agents\plugins\marketplace.json`
