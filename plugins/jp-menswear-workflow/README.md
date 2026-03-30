# JP Menswear Workflow Plugin

This plugin is the repository-side source for the Japanese menswear workflow.

It combines:
- a Git-managed plugin entry
- six reusable workflow skills
- one orchestrator skill that coordinates the six steps

## Structure

- `.codex-plugin/plugin.json`
- `skills/`

The `skills/` folder includes:
- `jp-menswear-orchestrator`
- `jp-menswear-step-1-trends`
- `jp-menswear-step-2-validation`
- `jp-menswear-step-3-sku-selection`
- `jp-menswear-step-4-sourcing`
- `jp-menswear-step-5-same-style-confirm`
- `jp-menswear-step-6-material-pack`

## How we use it

- This folder is the Git source of truth.
- The sync script copies the plugin into the current user's:
  `plugins\jp-menswear-workflow`
- The local marketplace entry points Codex to the synced plugin through the current user's:
  `.agents\plugins\marketplace.json`

## Update flow

1. Update files in this repository plugin folder.
2. From the repo root, run:
   `powershell -ExecutionPolicy Bypass -File ".\sync_jp_menswear_plugin.ps1"`
3. Restart Codex.
4. Check the plugin list for:
   `JP Menswear 6-Step Workflow`

## How to verify the installed version

Check either of these files in the installed plugin copy:

- `.codex-plugin/plugin.json`
- `.codex-plugin/build-info.json`

The current expected plugin version is:
- `0.1.1`

The current build commit is:
- `e4d6659`

## Team distribution

For teammates who should not manage this repository manually, use the repo-root helper:

`powershell -ExecutionPolicy Bypass -File ".\install_or_update_team_plugin.ps1"`

That helper downloads the latest repository snapshot and installs the plugin for the current Windows user.

## Step 6 direct-link behavior

If a teammate explicitly asks to do Step 6 and provides a product link:
- run Step 6 directly
- do not backfill Steps 1 to 5 unless they explicitly ask for the full workflow
- treat historical outputs only as optional reference material
- still produce a direct-to-publish material pack with slot mapping, source trace, upload order, copy, and web preview
