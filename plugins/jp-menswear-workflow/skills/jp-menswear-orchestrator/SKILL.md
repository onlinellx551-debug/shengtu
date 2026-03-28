---
name: jp-menswear-orchestrator
description: 用于统一驱动日本男装 6 步流程；当需要从任意一步启动、自动判断交接、复用已有输出并在 Step 5 停点后继续到素材包交付时使用。
---

# 日本男装总控技能

## 何时使用

当用户不想手动逐个调用 Step 1 到 Step 6，而是希望：
- 从头跑完整流程
- 从中间某一步继续
- 自动复用已有结果
- 在需要人工确认的地方停住
- 后续可以不断迭代改进每一步

就使用本技能。

典型触发词：
- “从头跑六步”
- “继续做下一步”
- “从第 4 步开始”
- “自动串起来”
- “总控”
- “orchestrator”

## 本技能管理的 6 个子技能

- `$jp-menswear-step-1-trends`
- `$jp-menswear-step-2-validation`
- `$jp-menswear-step-3-sku-selection`
- `$jp-menswear-step-4-sourcing`
- `$jp-menswear-step-5-same-style-confirm`
- `$jp-menswear-step-6-material-pack`

参考映射见：
- [step-skill-map.md](C:/Users/Administrator/.codex/skills/jp-menswear-orchestrator/references/step-skill-map.md)
- [workflow-contract.md](C:/Users/Administrator/.codex/skills/jp-menswear-orchestrator/references/workflow-contract.md)
- [completion-checklist.md](C:/Users/Administrator/.codex/skills/jp-menswear-orchestrator/references/completion-checklist.md)

## 总控职责

本技能不替代 6 个子技能，而是负责：
- 判断当前该从哪一步开始
- 读取已有输出，避免重复劳动
- 按顺序调用正确的子技能
- 控制停点与恢复点
- 保证每一步的交付文件可被下一步直接使用

每次运行前，先读取：
- [workflow-contract.md](C:/Users/Administrator/.codex/skills/jp-menswear-orchestrator/references/workflow-contract.md)
- [completion-checklist.md](C:/Users/Administrator/.codex/skills/jp-menswear-orchestrator/references/completion-checklist.md)

不要只凭当前线程里的一两句话猜测流程细节。

## 默认工作目录与输出约定

默认以当前工作区为项目根目录。

标准输出目录：
- `step1_output`
- `step2_output`
- `step3_output`
- `step4_output`
- `step5_output`
- `step6_output\\material_bundle`

如果历史项目里存在旧目录或旧版本文件：
- 优先复用已有结果
- 不强行搬迁旧文件
- 新结果优先写入标准目录
- 名称冲突时创建 `v2 / v3`

## 启动判断规则

### 1. 用户明确指定起点

如果用户说：
- “从第 3 步开始”
- “重做第二步”
- “只做素材包”

就以用户指定为准。

### 2. 用户说“继续”

按已有输出判断最近完成步：
- 如果 `step5_output` 里有用户已确认真同款结果，直接进入 Step 6
- 如果 `step5_output` 只有待确认表，没有确认结果，停在 Step 5
- 如果 `step4_output` 完整但没做同款确认，进入 Step 5
- 如果 `step3_output` 完整但没找商，进入 Step 4
- 如果 `step2_output` 完整但没选 SKU，进入 Step 3
- 如果 `step1_output` 完整但没做多来源验证，进入 Step 2
- 否则从 Step 1 开始

### 3. 用户说“从头重跑”

仍然先读取旧文件做参考，但重新生成新版本，不覆盖旧确认结果。

## 停点规则

唯一默认停点是 Step 5。

原因：
- Step 5 需要用户确认真同款
- 没有确认前，不能进入 Step 6 正式取素材和补图

满足以下任一条件时，才允许继续到 Step 6：
- 用户在当前线程明确确认了真同款
- `step5_output` 中存在清晰的已确认文件，并能识别确认结果

## 运行顺序

### 模式 A：全流程

1. 调用 `$jp-menswear-step-1-trends`
2. 调用 `$jp-menswear-step-2-validation`
3. 调用 `$jp-menswear-step-3-sku-selection`
4. 调用 `$jp-menswear-step-4-sourcing`
5. 调用 `$jp-menswear-step-5-same-style-confirm`
6. 在 Step 5 停住，等待用户确认
7. 用户确认后调用 `$jp-menswear-step-6-material-pack`

### 模式 B：断点续跑

从识别出的最近未完成步骤开始，逐步往下执行。

### 模式 C：单步修订

如果用户只要求某一步：
- 只跑该步
- 但先读取必要的前置输出
- 不破坏后续结果，除非用户明确要重算后续步骤

## 编排规则

### 1. 永远优先复用

不要重复生成已经足够好的结果。总控的默认行为是：
- 先读
- 再判断
- 最后执行

### 2. 永远保留证据链

后一步必须能追溯到前一步：
- Step 3 能追溯到 Step 1 和 Step 2
- Step 4 只跟随 Step 3 的 SKU
- Step 6 只使用 Step 5 已确认同款

### 3. 永远保持中文主输出

默认中文输出；必须出现英文时，用“中文（English）”。

### 4. 永远做可交接文件

每一步至少要有：
- Excel
- markdown 说明文件

需要时再附加：
- CSV
- JSON
- PNG
- HTML

## Step 6 的特殊总控规则

当总控进入 Step 6 时，必须额外检查：
- 是否真的有“已确认真同款”
- 是否输出到 `material_bundle`
- 是否包含：
  - `index.html`
  - `preview_check.png`
  - `preview_full.png`
  - `整页长图.png`
- 是否实际打开网页做过检查

如果项目涉及 Lovart：
- 只用 `AdsPower` 环境 `k1at9ows`
- 用 `Agent`
- 模型用 `Nano Banana Pro`
- 优先灯泡思考模式
- 卡住时优先重开同一个环境，不切换其他 Lovart 账号

## 用户沟通规则

总控技能应尽量减少让用户重新解释上下文。

优先做法：
- 直接读取现有文件
- 在有合理假设时先执行
- 只在真正需要人工确认时停住

不要做法：
- 每一步都重新问用户一次
- 明明能从文件判断，却要求用户重复说明

## 输出建议

当总控执行时，给用户的状态更新要说明：
- 现在处于第几步
- 读取了哪些已有结果
- 下一步会调用哪个子技能
- 是否会停点

## 示例

### 示例 1：从头跑

用户说：
- “把这六步完整跑一遍”

总控行为：
- 从 Step 1 开始
- 一直跑到 Step 5
- 停下来等用户确认同款

### 示例 2：继续

用户说：
- “继续”

总控行为：
- 检查 `step1_output` 到 `step6_output`
- 找到最近未完成步骤
- 从那一步继续

### 示例 3：只做素材包

用户说：
- “直接做第六步”

总控行为：
- 先检查 `step5_output` 是否真的有已确认真同款
- 有则进 Step 6
- 没有则提醒必须先完成 Step 5 确认

## 迭代建议

后续如果要继续增强总控技能，优先加：
- 自动识别“最新版本文件”
- 统一版本号策略
- 每一步的完成状态检测
- 一键生成流程总览表



