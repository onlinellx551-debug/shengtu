# 日本男装 6 步技能映射

## 技能列表

1. `jp-menswear-step-1-trends`
2. `jp-menswear-step-2-validation`
3. `jp-menswear-step-3-sku-selection`
4. `jp-menswear-step-4-sourcing`
5. `jp-menswear-step-5-same-style-confirm`
6. `jp-menswear-step-6-material-pack`

## 输出目录

1. `step1_output`
2. `step2_output`
3. `step3_output`
4. `step4_output`
5. `step5_output`
6. `step6_output\material_bundle`

## 默认停点

- Step 5：真同款确认后停住，等待用户确认

## 默认恢复点

- 如果 Step 5 已确认，则恢复到 Step 6
- 否则恢复到最近一个未完成步骤

