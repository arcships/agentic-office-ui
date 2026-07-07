# XLSX 迁移架构设计

Date: 2026-07-07
Status: 设计阶段

## 背景

按 `docs/upstream-xlsx-feature-alignment.md` 方案,把上游 `@extend-ai/react-xlsx`(commit `f285a1c`,~41000 行)迁移到 `packages/xlsx-core` + `packages/vue-xlsx`。

与 DOCX 不同,XLSX 的 xlsx-core 已部分迁移完成且能工作,但存在单文件超 1000 行的问题(charts.ts 4366行、images.ts 3870行、types.ts 1307行、composables.ts 4724行)。vue-xlsx 的 XlsxViewer 是 stub,chart-renderer 和 surface-regl 缺失。

## 硬约束

- 单文件 ≤ 1000 行
- 功能对齐上游
- 通过验证清单
- 仅限 XLSX,不动 docx/pdf

## 一、当前状态评估

### xlsx-core(已迁移,需拆分大文件)

| 文件 | 行数 | 状态 | 问题 |
|---|---:|---|---|
| wasm.ts | 82 | ✅ | - |
| safe-calculate.ts | 99 | ✅ | - |
| worker-client.ts | 230 | ✅ | - |
| xlsx-worker.ts | 835 | ✅ | - |
| colors.ts | 224 | ✅ | - |
| index.ts | 166 | ✅ | - |
| wasm-url.d.ts | 4 | ✅ | - |
| charts.ts | 4366 | ⚠️ | 超 1000 行,需按子领域拆分 |
| images.ts | 3870 | ⚠️ | 超 1000 行,需按子领域拆分 |
| types.ts | 1307 | ⚠️ | 超 1000 行,需拆分 |

### vue-xlsx(stub,需重写)

| 文件 | 行数 | 状态 |
|---|---:|---|
| composables.ts | 4724 | ⚠️ 机械改写,单文件超 1000 行,需拆分 |
| index.ts | 30 | stub(XlsxViewer pending) |
| env.d.ts | 5 | ✅ |
| chart-renderer | - | ❌ 缺失 |
| surface-regl | - | ❌ 缺失 |
| XlsxViewer.vue | - | ❌ 缺失 |
| types.ts | - | ❌ 缺失 |

## 二、目标模块设计

### 2.1 xlsx-core/ 拆分(大文件按子领域拆)

```
xlsx-core/src/
├── wasm.ts          (82) ✅
├── safe-calculate.ts (99) ✅
├── worker-client.ts (230) ✅
├── xlsx-worker.ts   (835) ✅
├── colors.ts        (224) ✅
├── wasm-url.d.ts    (4) ✅
├── index.ts         (barrel)
├── types/           ← types.ts 1307行拆分
│   ├── chart-types.ts     (~450) 图表类型
│   ├── image-types.ts     (~400) 图片/绘图类型
│   ├── worksheet-types.ts (~250) 工作表类型
│   └── index.ts           (barrel)
├── charts/          ← charts.ts 4366行拆分
│   ├── chart-parser.ts    (~700) XML 解析
│   ├── chart-series.ts    (~600) 系列公式
│   ├── chart-colors.ts    (~500) 颜色/主题
│   ├── chart-types.ts     (~800) 图表类型识别
│   ├── chart-styles.ts    (~800) 样式解析
│   ├── chart-export.ts    (~500) 导出/序列化
│   └── index.ts           (barrel)
└── images/          ← images.ts 3870行拆分
    ├── image-parser.ts    (~700) 图片解析
    ├── drawing-parser.ts  (~700) 绘图解析
    ├── column-width.ts    (~400) 列宽计算
    ├── theme-palette.ts   (~500) 主题色板
    ├── grid-render.ts     (~800) 网格渲染辅助
    ├── image-export.ts    (~500) 图片导出
    └── index.ts           (barrel)
```

### 2.2 vue-xlsx/ 重写

```
vue-xlsx/src/
├── composables/      ← composables.ts 4724行拆分
│   ├── useXlsxViewerController.ts  (~700) 核心 controller
│   ├── workbook-state.ts           (~600) 工作簿状态管理
│   ├── selection.ts                (~500) 选区管理
│   ├── editing.ts                  (~600) 单元格编辑
│   ├── chart-controller.ts         (~500) 图表 controller
│   ├── clipboard.ts                (~400) 复制粘贴
│   ├── navigation.ts               (~400) 导航/滚动
│   ├── formatting.ts               (~500) 格式化
│   └── index.ts                    (barrel)
├── render/           ← chart-renderer.tsx 7174行 + surface-regl.tsx 1185行
│   ├── chart-renderer.ts           (~800) 图表渲染主逻辑
│   ├── chart-bar.ts                (~500) 柱状图
│   ├── chart-line.ts               (~500) 折线图
│   ├── chart-pie.ts                (~400) 饼图
│   ├── chart-scatter.ts            (~500) 散点图
│   ├── chart-axis.ts               (~600) 坐标轴
│   ├── chart-legend.ts             (~400) 图例
│   ├── surface-regl.ts            (800) WebGL surface(可超1000,WebGL 着色器)
│   └── index.ts                    (barrel)
├── components/       ← XlsxViewer.tsx 16615行拆分
│   ├── XlsxViewer.vue              (~800) 主 viewer 容器
│   ├── XlsxGrid.vue                (~700) 网格/单元格
│   ├── XlsxToolbar.vue             (~600) 工具栏
│   ├── XlsxSheetTabs.vue           (~400) 工作表标签
│   ├── XlsxChartOverlay.vue        (~600) 图表覆盖层
│   ├── XlsxImageLayer.vue          (~500) 图片层
│   ├── XlsxSelectionOverlay.vue    (~500) 选区覆盖层
│   ├── XlsxContextMenu.vue         (~400) 右键菜单
│   └── index.ts                    (barrel)
├── types.ts          (~200) Vue 层类型
├── env.d.ts          (5)
└── index.ts          (barrel)
```

## 三、迁移执行顺序

| 步骤 | 任务 | scope | model | 依赖 |
|---|---|---|---|---|
| 1 | xlsx-core-types-split | xlsx-core | deepseek-v4-pro | - |
| 2 | xlsx-core-charts-split | xlsx-core | deepseek-v4-pro | 1 |
| 3 | xlsx-core-images-split | xlsx-core | deepseek-v4-pro | 1 |
| 4 | xlsx-composables-split | vue-xlsx | glm-5.2 | 2,3 |
| 5 | xlsx-render | vue-xlsx | glm-5.2 | 4 |
| 6 | xlsx-components | vue-xlsx | glm-5.2 | 4,5 |
| 7 | xlsx-demo | demo | deepseek-v4-pro | 6 |
| 8 | xlsx-verify | - | deepseek-v4-pro | 7 |

## 四、关键技术决策

### 4.1 不回退重做
xlsx-core 已能工作,不像 DOCX 那样是 stub。只拆分大文件,不重写已验证的逻辑。

### 4.2 WASM
复用现有 wasm.ts + worker-client.ts,不动。wasm 二进制已在 demo/public/。

### 4.3 渲染
- 图表用 Canvas 2D(上游 chart-renderer 是 Canvas)
- 网格 surface 用 WebGL(regl),surface-regl.ts 可超 1000 行(WebGL 着色器代码密集)

### 4.4 composables 拆分
useXlsxViewerController(4724行)按职责拆:状态/选区/编辑/图表/剪贴板/导航/格式化。

## 五、风险

| 风险 | 缓解 |
|---|---|
| charts.ts 拆分可能破坏内部依赖 | 拆分后全量 typecheck + build |
| XlsxViewer.tsx 16615行重写最复杂 | 先拆 composables 再重写组件 |
| WebGL surface-regl 耦合度高 | 保留为单文件,允许超 1000 行 |
| 图表渲染 Canvas API 复杂 | 按图表类型分文件 |
