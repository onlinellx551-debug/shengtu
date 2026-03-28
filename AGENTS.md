# task-codex local skill bridge

This workspace relies on a local skill set stored under:

- `C:\Users\Administrator\.codex\skills`

If a user mentions any of the names below in this workspace, do not treat them as plain text. Read the mapped `SKILL.md` and the listed reference files, then execute according to those documents.

## Skill map

### Orchestrator

Triggers:
- `$jp-menswear-orchestrator`
- `jp-menswear-orchestrator`
- `Japanese menswear orchestrator`
- `six-step workflow`

Read first:
- `C:\Users\Administrator\.codex\skills\jp-menswear-orchestrator\SKILL.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-orchestrator\references\workflow-contract.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-orchestrator\references\completion-checklist.md`

Execution contract:
- detect the correct starting step
- reuse existing outputs when possible
- stop at Step 5 for user confirmation
- only proceed to Step 6 after same-style confirmation

### Step 1

Triggers:
- `$jp-menswear-step-1-trends`
- `jp-menswear-step-1-trends`

Read first:
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-1-trends\SKILL.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-1-trends\references\playbook.md`

### Step 2

Triggers:
- `$jp-menswear-step-2-validation`
- `jp-menswear-step-2-validation`

Read first:
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-2-validation\SKILL.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-2-validation\references\playbook.md`

### Step 3

Triggers:
- `$jp-menswear-step-3-sku-selection`
- `jp-menswear-step-3-sku-selection`

Read first:
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-3-sku-selection\SKILL.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-3-sku-selection\references\playbook.md`

### Step 4

Triggers:
- `$jp-menswear-step-4-sourcing`
- `jp-menswear-step-4-sourcing`

Read first:
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-4-sourcing\SKILL.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-4-sourcing\references\playbook.md`

### Step 5

Triggers:
- `$jp-menswear-step-5-same-style-confirm`
- `jp-menswear-step-5-same-style-confirm`

Read first:
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-5-same-style-confirm\SKILL.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-5-same-style-confirm\references\playbook.md`

### Step 6

Triggers:
- `$jp-menswear-step-6-material-pack`
- `jp-menswear-step-6-material-pack`

Read first:
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-6-material-pack\SKILL.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-6-material-pack\references\playbook.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-6-material-pack\references\web-slot-template.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-6-material-pack\references\sku-input-contract.md`
- `C:\Users\Administrator\.codex\skills\jp-menswear-step-6-material-pack\references\oakln-style-replication-spec.md`

## Project path rule

These skills are global, but they should treat this workspace as the project root when running inside this repository.

Path precedence:
1. If the user explicitly gives a project path, use it.
2. Otherwise, if the current workspace contains the expected `step1_output` to `step6_output` structure, use the current workspace.
3. If neither condition is true, ask one concise question to confirm the intended project path before executing.

## Step 6 special rule

If the user only gives a product link:
- still run Step 6
- first extract what can be inferred from the link
- then reuse `step4_output`, `step5_output`, and `step6_output`
- if the user also gives color, reference page, or original asset paths, prefer the user-supplied inputs

Step 6 output root is fixed to:
- `C:\Users\Administrator\Desktop\Codex Project\task-codex\step6_output\material_bundle`

Step 6 must deliver:
- `index.html`
- `preview_check.png`
- `preview_full.png`
- a full-page long screenshot image
