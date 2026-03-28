# task-codex

This repository is the workspace for the Japanese menswear workflow project.

## What is included

- local bridge rules for the workflow skills
- the plugin source for the mixed delivery model
- plugin sync utilities
- business-facing prompt templates

## Main plugin source

- `plugins/jp-menswear-workflow`

## Sync plugin into the local Codex plugin directory

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Administrator\Desktop\Codex Project\task-codex\sync_jp_menswear_plugin.ps1"
```

## Installed local plugin path

- `C:\Users\Administrator\plugins\jp-menswear-workflow`

## Installed local marketplace path

- `C:\Users\Administrator\.agents\plugins\marketplace.json`
