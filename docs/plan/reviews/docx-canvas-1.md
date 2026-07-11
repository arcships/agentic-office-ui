# docx-canvas Review #1

Date: 2026-07-07
Review scope: 架构文档 §2.4 canvas/(2 文件) + 前置依赖层（engine/、layout/）
Upstream: `@extend-ai/react-docx` commit `6f70b92`

## 1. Findings

### 1.1 canvas 层（§2.4 产出物）

| # | 严重程度 | 阻塞 | 描述 | 设计文档位置 | 代码位置 |
|---|---|---|---|---|---|
| 1 | — | — | `canvas/types.ts`（44行）与上游一致，仅 import 路径从 `@extend-ai/react-docx-layout-engine` 改为相对路径 `../layout/layout-engine`。类型定义完整，无遗漏。 | §2.4 表格 | `packages/docx-core/src/canvas/types.ts` |
| 2 | — | — | `canvas/layout-diagnostics.ts`（156行）核心算法与上游一致。上游调用 `resolveDocumentLayout` from `../section-layout`，当前适配为 `resolveDocumentForLayout` from `../layout`，pageSize 访问路径由 `documentLayout.pageWidthPx` 改为 `documentLayout.layout.pageSizePx.width`。函数 `buildDocxLayoutDiagnostics` 功能完整。 | §2.4 表格 | `packages/docx-core/src/canvas/layout-diagnostics.ts` |
| 3 | P2 | non-blocking | `buildDocxLayoutDiagnostics` 当前使用 layout 层的默认页面尺寸（816×1056 px），未解析 `<w:sectPr>` XML 获取文档实际 section 尺寸。上游通过 `section-layout.ts` 的 `resolveDocumentLayout` 实现此功能。`section-layout.ts` 属于 viewer/ 层（步骤 4），尚未迁移。待 viewer 层完成时需回填此调用。 | §2.4，上游对齐文档 §2.4 | `packages/docx-core/src/canvas/layout-diagnostics.ts:128-130` |

### 1.2 layout 层（§2.2 前置依赖）

| # | 严重程度 | 阻塞 | 描述 | 设计文档位置 | 代码位置 |
|---|---|---|---|---|---|
| 4 | P3 | non-blocking | `page-segmentation-core.ts` 实际 864 行，超出设计估算 ~650 行（+214 行）。`page-segmentation-table.ts` 实际 376 行，低于设计估算 ~573 行（-197 行）。合并总量 1240 行 vs 上游 1223 行（+17 行分包开销）。切分边界与设计预估不一致，但单文件均 ≤1000 行，满足硬约束。 | §2.2 表格 | `packages/docx-core/src/layout/page-segmentation-core.ts`、`packages/docx-core/src/layout/page-segmentation-table.ts` |
| 5 | — | — | layout 层 5 文件齐全（`layout-engine.ts` 359行、`pagination.ts` 689行、`page-segmentation-core.ts` 864行、`page-segmentation-table.ts` 376行、`index.ts` 383行），与 §2.2 设计一致。 | §2.2 | `packages/docx-core/src/layout/` |

### 1.3 engine 层（§2.1 已标注 ✅）

| # | 严重程度 | 阻塞 | 描述 | 设计文档位置 | 代码位置 |
|---|---|---|---|---|---|
| 6 | — | — | engine/ 层 8 文件齐全（`types.ts`、`clone.ts`、`normalize.ts`、`wasm.ts`、`ooxml-core.ts`、`serializer.ts`、`doc-model.ts`、`index.ts` + generated/），与 §2.1 一致。架构标注为 "已完成 ✅"。 | §2.1 | `packages/docx-core/src/engine/` |

### 1.4 editor 层（§2.5 前置依赖）

| # | 严重程度 | 阻塞 | 描述 | 设计文档位置 | 代码位置 |
|---|---|---|---|---|---|
| 7 | P2 | non-blocking | `editor/editor-ops.ts`（1329行）尚未按 §2.5 拆分为 `paragraph-ops.ts`（~480行）、`run-style-ops.ts`（~450行）、`table-ops.ts`（~399行）。当前为单文件拷贝自上游，超过 1000 行硬约束。此拆分是步骤 5-6 的产出物，不在 canvas 步骤范围内。 | §2.5 表格 "editor-ops.ts 拆分" | `packages/docx-core/src/editor/editor-ops.ts` |
| 8 | — | — | `editor/helpers/` 目录不存在。步骤 5 产出物（39 文件），尚未开始。架构文档已明确标注为待实现。 | §2.5 "editor/helpers/" 表格 | — |

### 1.5 viewer 层（§2.3）

| # | 严重程度 | 阻塞 | 描述 | 设计文档位置 | 代码位置 |
|---|---|---|---|---|---|
| 9 | — | — | `viewer/` 目录为空。步骤 4 产出物（15 文件），尚未开始。架构文档已明确标注为待实现。 | §2.3 | `packages/docx-core/src/viewer/`（空目录） |

### 1.6 vue-docx（§3）

| # | 严重程度 | 阻塞 | 描述 | 设计文档位置 | 代码位置 |
|---|---|---|---|---|---|
| 10 | — | — | `vue-docx` 当前仅有 `composables.ts`（单文件）、`env.d.ts`、`index.ts`。架构 §3 要求 26 个 composable + 18 个组件 + 6 个 render 文件。步骤 7-9 产出物，尚未开始。架构文档已明确标注为待实现。 | §3.1–3.3 | `packages/vue-docx/src/` |

### 1.7 交叉验证

| # | 严重程度 | 阻塞 | 描述 | 设计文档位置 | 代码位置 |
|---|---|---|---|---|---|
| 11 | — | — | **import 路径**：canvas 层所有 import 均为相对路径（`../engine/types`、`../layout/layout-engine`、`../layout`、`./types`），无 `@extend-ai/react-docx-*` 包引用残留。grep 确认 docx-core/src 下仅注释和错误消息中含上游包名。 | — | `packages/docx-core/src/canvas/` |
| 12 | — | — | **typecheck**：`pnpm --filter @arcships/docx-core typecheck`（`tsc --noEmit`）通过，零错误。 | — | — |
| 13 | — | — | **stub/mock/fake**：canvas 层两个文件均为完整实现，无 `TODO`、`FIXME`、`throw new Error("not implemented")` 等占位模式。 | — | `packages/docx-core/src/canvas/` |
| 14 | — | — | **循环依赖**：canvas → layout → engine，单向无环。 | §5 依赖关系图 | — |
| 15 | — | — | **docx-core/src/index.ts** barrel 导出：engine（`./engine`）、layout（`./layout`）、editor-ops（`./editor/editor-ops`）、canvas types（`./canvas/types`）、canvas layout-diagnostics（`./canvas/layout-diagnostics`）。路径正确，覆盖所有已实现模块。 | §2.6 | `packages/docx-core/src/index.ts` |

## 2. 结论

**pass**

canvas 层（§2.4）2 文件全部到位，行数与架构设计一致，功能对齐上游（适配本地 layout API），import 路径均为相对路径，typecheck 零错误通过，无残留 stub，模块边界清晰无循环依赖。

非阻塞 P2 项（#3、#7）属于后续步骤（viewer/、editor split）的衔接点，不影响 canvas 层本身的交付质量。

阻塞项：无。
