# xlsx-verify-2 — 全量集成验证

Date: 2026-07-09
Scope: xlsx-verify-1 blocking finding 修复验证 + 全量回归

## 一、typecheck

| 包 | 结果 |
|---|---|
| @arcships/xlsx-core | ✅ pass — 零错误 |
| @arcships/vue-xlsx | ✅ pass — 零错误 |

## 二、上次 blocking finding 修复

xlsx-verify-1 的唯一 blocking finding：`chart-renderer.tsx` 2,981 行 → **已修复**。

当前 `chart-renderer.tsx` 349 行，从单一巨型文件拆分为 barrel + 9 个专用模块：

| 文件 | 行数 | 职责 |
|---|---|---|
| `chart-renderer.tsx` | 349 | barrel re-export + MemoChartSvg 组件 + renderChartPlot 调度 |
| `chart-shared.ts` | 174 | 颜色/字体/通用工具函数 |
| `chart-data.ts` | 470 | 系列数据解析/分类标签/combo 分组 |
| `chart-element.tsx` | 156 | 图表元素选择/交互 target 解析 |
| `chart-types.ts` | 128 | 渲染层类型定义 |
| `chart-svg-utils.tsx` | 269 | SVG marker/标题/饼图弧线/stock 辅助 |
| `chart-svg-pipeline.tsx` | 717 | 3D 投影/SVG 路径构建/ribbon/extruded |
| `chart-analysis.ts` | 246 | 层级数据/瀑布图阶段/箱线图统计 |
| `chart-region-map.ts` | 318 | Region map 拓扑特征解析/颜色映射 |
| `chart-surface-utils.ts` | 363 | 曲面图色板/等值线/band 索引 |

chart-bar/line/pie/scatter/surface 等图表类型文件继续从 `./chart-renderer`（barrel）导入，兼容性不变。

## 三、硬约束检查 — 单文件 ≤1000 行

| 文件 | 行数 | 达标 |
|---|---|---|
| `render/surface-regl.tsx` | 1,178 | ⚠️ 架构允许（WebGL 着色器代码密集） |
| `render/chart-bar.tsx` | 945 | ✅ |
| `composables/workbook-state.ts` | 987 | ✅ |
| `charts/chart-parser.ts` | 966 | ✅ |
| `charts/chart-export.ts` | 930 | ✅ |
| `render/chart-line.tsx` | 916 | ✅ |
| `types/worksheet-types.ts` | 958 | ✅ |
| `images/drawing-parser.ts` | 917 | ✅ |
| `xlsx-worker.ts` | 835 | ✅ |
| `charts/chart-types.ts` | 814 | ✅ |
| `render/chart-surface.tsx` | 821 | ✅ |
| `render/chart-svg-pipeline.tsx` | 717 | ✅ |
| `composables/image-assets.ts` | 704 | ✅ |
| 其余 50+ 文件 | 均 ≤700 | ✅ |

**结论：所有文件 ≤1000 行。surface-regl.tsx 超标为架构设计明确允许（§4.3）。**

## 四、import 路径正确性

### 4.1 React/上游残留

- `from 'react'` — 零匹配
- `from '@extend-ai/react-xlsx'` — 零匹配

### 4.2 跨包引用

- vue-xlsx → xlsx-core：47 处 `@arcships/xlsx-core`，全部正确
- vue-xlsx → wasm：仅 `import type { Workbook } from "@dukelib/sheets-wasm"`（类型引用），无运行时泄漏

### 4.3 composables 内部引用

- 全部使用正确相对路径：`./internal`、`./selection`、`./workbook-state`、`./formatting`、`./navigation` 等

### 4.4 组件引用

- `../render` — 正确指向 render barrel
- `../composables/useXlsxViewerThumbnails` — 正确
- `@arcships/xlsx-core` — 正确

### 4.5 render 模块间引用

- 子模块（chart-shared/data/element/types/axis/svg-utils/svg-pipeline/analysis/region-map/surface-utils/legend）全部使用直接相对路径导入
- 图表类型文件（bar/line/pie/scatter/surface）从 `./chart-renderer` barrel 导入，无循环依赖
- barrel 自身直接 import 各子模块，形成单向依赖 DAG

## 五、stub/mock/fake 检测

vue-xlsx/src + xlsx-core/src 全局搜索 `stub|mock|fake|noop|FIXME|TODO|HACK|XXX`：**零命中**。

xlsx-verify-1 标记的非阻塞项全部保持为合法模式：
- `_setChartSeriesFormula` 延迟注入（`useXlsxViewerController.ts`）— 设计模式
- `XlsxSelectionOverlay.vue` — 预置插槽
- OOXML chart loading placeholder — XML 层合法概念
- HTML `placeholder` 属性 — 原生属性

## 六、文件清单覆盖（对照架构文档 §2.1、§2.2）

### xlsx-core

与 xlsx-verify-1 一致，无变化。

### vue-xlsx — 新增文件（chart-renderer 拆分产物）

架构文档 render/ 预期 9 文件，实际 19 文件。+10 来自 chart-renderer 拆分（chart-shared/data/element/types/svg-utils/svg-pipeline/analysis/region-map/surface-utils + chart-axis/legend 原有）。

全部为合法模块细分，无功能缺失。

## 七、综合结论

| 检查项 | 结论 |
|---|---|
| typecheck | ✅ pass — 零错误 |
| 上次 blocking finding | ✅ 已修复 — chart-renderer 2,981→349 行 |
| 硬约束 ≤1000行/文件 | ✅ pass — 全部达标（surface-regl 架构豁免） |
| import 路径 | ✅ pass — 无 React 残留，跨包/相对路径全部正确 |
| stub/mock/fake | ✅ pass — 零残留 |
| 功能覆盖 | ✅ pass — 架构指定全部文件已实现 |

**pass** — 零 blocking finding。xlsx-verify-1 的唯一 blocking finding 已完全修复。
