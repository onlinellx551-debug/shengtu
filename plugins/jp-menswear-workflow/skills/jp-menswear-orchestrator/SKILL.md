---
name: jp-menswear-orchestrator
description: 用于统一驱动日本男装六步流程；当需要从任意一步启动、自动判断交接、复用已有输出并在 Step 5 停点后继续到素材包交付时使用。若用户明确指定“只做第六步 / 只做素材包 / 基于链接直接做素材包”，则直接进入 Step 6，不补跑前五步，除非用户明确要求完整流程。
---

# 日本男装总控技能

## 何时使用

当用户不想手动逐个调用 Step 1 到 Step 6，而是希望我们：
- 从头跑完整流程
- 从中间某一步继续
- 自动复用已有结果
- 在 Step 5 停下来等确认
- 在确认后继续到 Step 6

就使用本技能。

典型触发词：
- “从头跑六步”
- “继续当前项目”
- “从第 4 步开始”
- “只做第六步”
- “总控”
- “orchestrator”

## 使用前先读

- [workflow-contract.md](references/workflow-contract.md)
- [completion-checklist.md](references/completion-checklist.md)
- [step-skill-map.md](references/step-skill-map.md)

## 管理的子技能

- `$jp-menswear-step-1-trends`
- `$jp-menswear-step-2-validation`
- `$jp-menswear-step-3-sku-selection`
- `$jp-menswear-step-4-sourcing`
- `$jp-menswear-step-5-same-style-confirm`
- `$jp-menswear-step-6-material-pack`

## 核心职责

总控本身不替代 6 个子技能，而是负责：
- 判断应该从哪一步开始
- 读取已有输出，避免重复劳动
- 按顺序调用正确的子技能
- 控制停点与恢复点
- 保证每一步的交付物能被下一步继续使用

## 默认项目根与输出结构

默认把当前工作区当成项目根目录；如果用户显式给了项目路径，则优先用用户给的路径。

标准输出目录：
- `step1_output`
- `step2_output`
- `step3_output`
- `step4_output`
- `step5_output`
- `step6_output\material_bundle`

## 启动判断规则

### 1. 用户明确指定起点

如果用户明确说：
- “从第 3 步开始”
- “重做第二步”
- “只做素材包”

就按用户指定的步骤执行。

### 2. 用户明确说“只做第六步”

如果用户明确要求：
- “只做第六步”
- “只做素材包”
- “基于这个链接直接做素材包”

则默认进入 **Step 6 直链模式**：
- 不补跑 Step 1 到 Step 5
- 历史步骤结果只作为可选参考
- 重点直接放在交付可上架素材包

### 3. 用户说“继续”

按现有输出判断最近一个高质量完成步骤，再从下一步继续：
- Step 5 已确认：进入 Step 6
- Step 5 仅待确认：停在 Step 5
- Step 4 完整：进入 Step 5
- Step 3 完整：进入 Step 4
- Step 2 完整：进入 Step 3
- Step 1 完整：进入 Step 2
- 否则从 Step 1 开始

## 停点规则

默认只在 Step 5 停住。

只有满足以下任一条件，才允许正式进入 Step 6：
- 用户在当前线程明确确认了真同款
- `step5_output` 中存在明确的已确认文件

例外：
- 如果用户明确要求“只做第六步”，则走 Step 6 直链模式，不等待 Step 5 停点

## 与用户的沟通规则

更新进度时要明确说明：
- 现在在第几步
- 依据哪些已有文件做判断
- 下一步要调用哪个子技能
- 是否会停在 Step 5

如果是 Step 6 直链模式，要明确写：
- 当前只执行 Step 6
- 不补跑前五步
- 将直接围绕产品链接交付素材包
