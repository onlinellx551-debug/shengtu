# JP Menswear Workflow Plugin

This plugin is the repository-side source of truth for the Japanese menswear workflow.

It combines:
- one business-facing plugin entry
- six reusable workflow skills
- one orchestrator skill that coordinates the six steps

## Included skills

- `jp-menswear-orchestrator`
- `jp-menswear-step-1-trends`
- `jp-menswear-step-2-validation`
- `jp-menswear-step-3-sku-selection`
- `jp-menswear-step-4-sourcing`
- `jp-menswear-step-5-same-style-confirm`
- `jp-menswear-step-6-material-pack`

## What this plugin is for

This plugin is designed for a Japanese menswear workflow where teammates may need to:
- run the full six-step process from trend analysis to material-pack delivery
- continue from the correct step in an existing project
- jump directly into Step 6 and build a publish-ready material pack from a product link

## Current business entry behavior

### Full workflow mode

Use the orchestrator when the teammate wants the full workflow or wants the system to decide where to resume.

### Step 6 direct-link mode

If a teammate explicitly asks for Step 6 only and provides a product link, the plugin should:
- go directly into Step 6
- skip Steps 1-5 by default
- reuse older outputs only as optional reference
- produce a direct-to-publish material pack instead of a generic narrative summary
- surface deliverable entry points first instead of long shell-style process logs

## Update flow

1. Update files under this repository plugin folder.
2. From the repository root, run:
   `powershell -ExecutionPolicy Bypass -File ".\sync_jp_menswear_plugin.ps1"`
3. Restart Codex.
4. Check the plugin list for:
   `JP Menswear 6-Step Workflow v0.1.4`

## Team installation

For teammates who should not manage this repository manually, use:

`powershell -ExecutionPolicy Bypass -File ".\install_or_update_team_plugin.ps1"`

That helper downloads the latest repository snapshot and installs the plugin for the current Windows user.

## How to verify the installed version

Check either file in the installed plugin copy:
- `.codex-plugin/plugin.json`
- `.codex-plugin/build-info.json`

Current expected version:
- `0.1.4`

The Codex plugin UI should also display the version directly in:
- `displayName`
- `shortDescription`
