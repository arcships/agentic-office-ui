---
id: xlsx-integ
scope: xlsx-core, vue-xlsx
status: pending
depends-on: [xlsx-003]
---

# XLSX 集成验证：composables → xlsx-core 真实调用

## objective

验证 vue-xlsx/composables.ts 对 xlsx-core 引擎层的调用是真实路径（不是 stub/mock）。覆盖集成点 X1-X3。

## context

- `docs/plan/analysis/migration-split.md` 集成关系枚举 X1-X3
- `docs/upstream-xlsx-feature-alignment.md` 引擎层对齐要点 1-22、编辑能力要点 23-38

## path

- `packages/xlsx-core/src/`（被调方，只读验证）
- `packages/vue-xlsx/src/composables.ts`（调用方，只读验证）

## verification

### X1: wasm 引擎调用

```bash
node --input-type=module - <<'NODE'
// 验证 composables 实际调用 xlsx-core 的 wasm API
// 1. loadWorkbookFromBuffer → Workbook.fromBytes
// 2. recalculate → workbook.calculate
// 3. setCellFormula → worksheet.setFormula
// 4. exportXlsx → workbook.saveXlsxBytes
// 验证：导出的 xlsx 用 openpyxl 打开，cell/formula 保留
NODE
```

通过标准：
- `fromBytes` → `calculate` → `setFormula` → `saveXlsxBytes` 完整路径执行
- openpyxl 验证导出 xlsx 的 cell 值和 formula 正确

### X2: worker 调用

```bash
node --input-type=module - <<'NODE'
// 验证 composables 的 worker 模式
// 1. useWorker=true → XlsxWorkerClient.loadWorkbook
// 2. getCellSnapshot 返回 {displayValue, formula}
// 3. getRowsBatch 返回行数据
NODE
```

通过标准：worker 加载成功，snapshot/batch 查询返回真实数据（非 null/空）

### X3: 图片/图表 assets

```bash
node --input-type=module - <<'NODE'
// 验证 composables 调用 images/charts 解析
// 加载 charts-images.xlsx
// 1. parseWorkbookImageAssets 返回非空 images
// 2. loadWorkbookChartAssets 返回非空 charts
// 3. mergeWorkbookImageAssets 在导出时合并图片
NODE
```

通过标准：images/charts 数组非空，导出 xlsx 图片保留

### blocking 判定

以下情况为 blocking：
- composables 中任何 `/* noop */` / `return null` / `return false` 替代真实 wasm 调用
- worker 消息协议不完整（缺少 load/getCellSnapshot/getRowsBatch 之一）
- 导出路径不走 `saveXlsxBytes` + sanitize + merge images
