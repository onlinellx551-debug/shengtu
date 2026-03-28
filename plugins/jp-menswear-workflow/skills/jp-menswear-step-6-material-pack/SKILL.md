---
name: jp-menswear-step-6-material-pack
description: 用于在用户确认真同款后，基于同款素材制作可直接上架的素材包；当需要主图、详情图、文案、网页预览与质检时使用。
---

# 日本男装 Step 6：素材包

## 何时使用

当用户已经确认真同款供应商，且要基于这些同款素材做可直接上架的主图、详情图、文案、网页预览和上传包时，使用本技能。

典型触发词：
- “第六步”
- “素材包”
- “详情页素材”
- “网页预览”
- “可直接上架”

每次使用前，先读取：
- [playbook.md](C:/Users/Administrator/.codex/skills/jp-menswear-step-6-material-pack/references/playbook.md)
- [web-slot-template.md](C:/Users/Administrator/.codex/skills/jp-menswear-step-6-material-pack/references/web-slot-template.md)
- [sku-input-contract.md](C:/Users/Administrator/.codex/skills/jp-menswear-step-6-material-pack/references/sku-input-contract.md)
- [oakln-style-replication-spec.md](C:/Users/Administrator/.codex/skills/jp-menswear-step-6-material-pack/references/oakln-style-replication-spec.md)

## 联动契约

这是 6 步流程的第 6 步。

输入：
- `step5_output` 中用户已确认的真同款结果

输出目录：
- 固定使用 `C:\Users\Administrator\Desktop\Codex Project\task-codex\step6_output\material_bundle`

必须产出：
- 素材文件夹
- 1 个 Excel
- 1 个 markdown 说明文件
- 网页预览文件

必须固定包含：
- `index.html`
- `preview_check.png`
- `preview_full.png`
- `整页长图.png`

## 核心原则

- 只允许用“已确认真同款”的素材做基础
- 不能混入非同款素材
- 如果原始素材不够，可以补图，但补图也必须基于真同款素材
- 目标是“可直接上传使用”，不是只做演示页

## 工作流

### 1. 收集原始素材

优先收：
- 主图
- 细节图
- 长图
- 视频
- 参数
- 尺寸
- 色卡
- 文案

先按：
- [sku-input-contract.md](C:/Users/Administrator/.codex/skills/jp-menswear-step-6-material-pack/references/sku-input-contract.md)

检查输入是否够用。输入不足时，不要假装可以稳定复刻，必须标明缺口并补采。

默认输入策略：
- 最低只要求 1 个产品链接
- 如果用户额外给了颜色、卖点、原图路径、参考页链接等信息，优先使用这些显式输入
- 如果用户没给，自动从产品链接、已有项目输出、同款确认结果中推断

### 2. 槽位映射

如果项目里已有成功商品页结构，要先拆槽位，再按槽位补素材。不要先乱生图。

默认先按：
- [web-slot-template.md](C:/Users/Administrator/.codex/skills/jp-menswear-step-6-material-pack/references/web-slot-template.md)

去核对“主图区 / 中段卖点区 / 色卡区 / 推荐区 / 员工服装区”的槽位，再决定哪些图可复用、哪些图必须重生。

如果当前项目要求“换一个 SKU，也要做出和参照页一样的呈现方式”，再额外读取：
- [oakln-style-replication-spec.md](C:/Users/Administrator/.codex/skills/jp-menswear-step-6-material-pack/references/oakln-style-replication-spec.md)

此时默认目标是：
- 结构一致
- 槽位职责一致
- 质感一致
- 只替换为新 SKU 的真实产品信息和真实同款素材

### 3. 补图与增强

如果需要生成或增强，优先遵守当前项目固定约束：
- `Lovart` 只用 `AdsPower` 环境 `k1at9ows`
- 使用 `Agent`
- 模型选 `Nano Banana Pro`
- 优先使用灯泡思考模式
- 如果 `Lovart` 卡住，先重开同一个 AdsPower 环境，不换别的 Lovart 账号

### 4. 网页质检

每次出图后都要：
- 打开 `index.html`
- 检查中文是否乱码
- 检查图片是否加载
- 检查是否是整张商品网页，而不是半页或对比页

## 输出结构

建议包含：
- 主图上传顺序
- 详情槽位覆盖表
- 卖点图
- 色卡图
- 推荐图
- 日文标题 / 副标题 / 卖点 / 描述
- 网页预览

## 质量门槛

- 图能直接上传
- 网页能正常打开
- 没有坏图
- 没有混入非同款素材
- 高级感、真实感优先于速度
- 如果是“复刻参照页”任务，必须达到“换 SKU 不换表达结构”的一致性

## 与未来迭代的接口

本技能要持续可迭代，优先改：
- 槽位映射
- 补图质量
- 网页检查
- 主图和卖点图的高级感

如果用户后续长期使用，始终把：
- `material_bundle`
- 槽位清单
- 预览图

作为稳定入口，不要频繁改目录结构。





