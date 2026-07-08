# docx-components 开发产出 Review

Date: 2026-07-08
Reviewer: DimCode
Status: **pass**

## 一、检查范围

对照两份架构文档进行逐项校验：
- `docs/docx-migration-architecture.md`（"docx-components" 章节）
- `docs/upstream-docx-feature-alignment.md`（上游 `@extend-ai/react-docx` commit `6f70b92` 对齐清单）

## 二、文件清单覆盖度

### 2.1 docx-core（架构指定 ~78 文件，实际 ~84 文件）

| 层 | 架构指定 | 实际文件数 | 偏差 |
|---|---:|---:|---|
| engine | 8 | 8 | 无 |
| layout | 5 | 5 | 无 |
| viewer | 15 | 16 | +1（pretext-items-layout.ts，为 ≤1000 行约束从 pretext-layout.ts 额外拆分） |
| canvas | 2 | 2 | 无 |
| editor/ | 4 | 4 | 无 |
| editor/helpers | 39 | 47 | +8（额外拆分以满足 ≤1000 行约束 + 循环依赖断裂） |

**额外文件明细**：
- `line-height-wrap.ts`（412 行）：从 line-height.ts 拆分段落换行估算
- `paragraph-geometry-image.ts`（755 行）：从 paragraph-geometry.ts 拆分浮动图片几何
- `paragraph-toc.ts`（85 行）：断裂 field-helpers ↔ line-height 循环依赖
- `line-height-table-extra.ts`（288 行）：从 line-height-table.ts 拆分表高估算
- `table-height.ts`（8 行）：re-export barrel，待后续重命名
- 生成文件：`engine/generated/docx_wasm.js`（wasm bindgen 产物）

所有额外文件均为合理的技术拆分（满足 ≤1000 行约束或断裂循环依赖），非冗余文件。架构文档中明确标注"大模块再拆"，因此这些拆分符合设计意图。

### 2.2 vue-docx（架构指定 ~50 文件，实际 ~53 文件）

| 层 | 架构指定 | 实际文件数 | 偏差 |
|---|---:|---:|---|
| composables | 26 | 28 | +2（editor-shared.ts + composables.ts 兼容层） |
| components | 18 | 18 | 无 |
| render | 6 | 7 | +1（paragraph-runs-field-resolve.ts，从 paragraph-runs.ts 拆分） |

**额外文件明细**：
- `editor-shared.ts`（91 行）：composables 子模块共享类型定义
- `composables.ts`（6 行）：向后兼容 re-export barrel
- `paragraph-runs-field-resolve.ts`（158 行）：字段解析逻辑从 paragraph-runs.ts 拆分

架构文档中"编辑器状态机拆分"和"renderParagraphRuns 按类型分文件"的意图覆盖了这些拆分。

## 三、≤1000 行约束检查

**结论：所有文件 ≤1000 行。** 唯一边界值为 `selection-helpers.ts`（恰好 1000 行），仍在约束内。

| 最大文件 | 行数 | 层 |
|---|---|---|
| selection-helpers.ts | 1000 | editor/helpers |
| pagination-plan-core.ts | 990 | editor/helpers |
| numbering.ts | 959 | editor/helpers |
| tracked-changes.ts | 949 | editor/helpers |
| paragraph-runs.ts | 931 | vue-docx/render |
| editor-types-extra.ts | 908 | editor/helpers |
| line-height-table.ts | 899 | editor/helpers |
| table-utils.ts | 866 | editor/helpers |
| page-segmentation-core.ts | 864 | layout |

## 四、Import 路径检查

### 4.1 docx-core 内部
**全部使用相对路径。** 零 `@extend-ai/react-docx-*` 的 import 引用。上游包引用仅出现在注释中（迁移来源标注），符合架构 Phase 1 要求。

典型 import 模式：
```typescript
import { cloneDocModel } from "../engine/clone";
import type { DocModel } from "../engine/types";
import { twipsToPixels } from "../../viewer/section-layout";
```

### 4.2 vue-docx
**使用 `@extend-ai/docx-core` 包名导入。** 这是跨 workspace 包的合法引用方式，等同于相对路径但更稳定。符合架构 Phase 3 要求。

典型 import 模式：
```typescript
import { cloneDocModel } from "@extend-ai/docx-core";
import type { DocModel } from "@extend-ai/docx-core";
```

### 4.3 React 残留检查
**docx-core 中零 React 导入、零 React hook 调用。** 73 处原有 `React.CSSProperties` 已全部替换为 `Record<string, string | number | undefined>`。所有 React 类型标注（`React.Dispatch`、`React.KeyboardEvent`、`React.PointerEvent`、`React.MouseEvent`、`React.ReactNode`）已清理完毕。

## 五、Typecheck 与构建

| 包 | typecheck | build | 备注 |
|---|---|---|---|
| `@extend-ai/docx-core` | ✅ 零错误 | ✅ 646.97 KB ESM | 188.76 KB DTS |
| `@extend-ai/vue-docx` | ✅ 零错误 | ✅ 174.66 KB ESM | 13.47 KB CSS |

## 六、Stub/Mock/Fake 检查

### 6.1 无残留 stub
所有函数均有完整实现体，未发现 `throw new Error("not implemented")` 或空函数体。搜索覆盖 `TODO`、`FIXME`、`stub`、`mock`、`fake`、`not implemented`、`not yet` 等关键词。

### 6.2 疑似匹配项核实

| 匹配项 | 文件 | 结论 |
|---|---|---|
| `// TODO` | `engine/generated/docx_wasm.js` | wasm-bindgen 自动生成代码，非手写 |
| `// not yet ported` | `editor/helpers/line-height-wrap.ts:9` | 注释放置的 tab-leader 快速路径，上游原文保留 |
| `throw new Error("browser environment")` | `viewer/thumbnail-raster.ts` | 合法运行时环境检查 |
| `table-height.ts`（8 行） | `editor/helpers/` | re-export barrel，`export * from "./line-height-table"`，非 stub |
| `composables.ts`（6 行） | `vue-docx/src/` | 向后兼容 re-export barrel，`export * from "./composables/index"`，非 stub |

### 6.3 架构明确标注"待实现"的例外
架构文档未标注任何 docx-components 范围的"待实现"项。所有 47 个 editor/helpers 模块均标注为完成状态（`editor/helpers/index.ts` 写有 "0 remaining"）。

## 七、上游功能对齐验证

### 7.1 引擎层（wasm + ooxml-core + serializer + doc-model）
- `engine/types.ts` 与上游 `doc-model/src/types.ts` **完全一致**（diff 零差异）
- `engine/wasm.ts` 与上游仅差异错误消息前缀（`"react-docx"` → `"@extend-ai/docx-core"`）
- `engine/ooxml-core.ts` 与上游仅差异 import 路径（`@extend-ai/react-docx-wasm` → `./wasm`）
- `engine/clone.ts`、`engine/normalize.ts`、`engine/serializer.ts` 内容一致

### 7.2 布局层
- 上游 `page-segmentation.ts`（1223 行）→ 拆分为 `page-segmentation-core.ts`（864）+ `page-segmentation-table.ts`（376）= 1240 行（+17 行来自 import 调整和文件头注释）
- 上游 `pagination.ts`（689 行）→ 当前 `pagination.ts`（690 行）
- 上游 `layout-engine/src/index.ts`（359 行）→ 当前 `layout-engine.ts`（359 行）

### 7.3 Viewer 辅助模块
- 上游 `pretext-layout.ts`（1389 行）→ 拆分为 `pretext-layout.ts`（679）+ `pretext-selection.ts`（237）+ `pretext-items-layout.ts`（471）= 1387 行（-2 行）
- 上游 `thumbnail-raster.ts`（1239 行）→ 拆分为 `thumbnail-raster.ts`（941）+ `thumbnail-cache.ts`（295）= 1236 行（-3 行）
- 其余 13 个 viewer 文件全部对齐

### 7.4 Editor 操作
- 上游 `editor-ops/src/index.ts`（1329 行）→ 拆分为 `paragraph-ops.ts`（824）+ `run-style-ops.ts`（386）+ `table-ops.ts`（165）= 1375 行（+46 行）

### 7.5 renderParagraphRuns
使用 Vue `h()` 生成 VNode，正确复刻了上游 JSX 的三路渲染（text/image/form-field）逻辑。931 行完整实现，非 stub。

### 7.6 dispatchEditorTransaction
完整复刻上游事务分发器：snapshot-based transaction dispatch、history push（上限 100）、DOM 选区恢复 via nonce、session-aware suppress 逻辑。130 行完整实现。

## 八、与架构文档的偏差项

| 偏差 | 说明 | 影响 |
|---|---|---|
| viewer 多 1 文件 | pretext-items-layout.ts 额外拆分 | 无负面影响，满足 ≤1000 行约束 |
| editor/helpers 多 8 文件 | 拆分 + 循环依赖断裂 | 架构明确允许"大模块再拆" |
| vue-docx composables 多 2 文件 | editor-shared.ts + 兼容层 | 架构允许合理的共享类型提取 |
| render 多 1 文件 | paragraph-runs-field-resolve.ts | 满足 ≤1000 行约束 |
| line-height-table ~730→~899 行 | 实际代码量比估计大 | ≤1000 约束通过 |
| selection-helpers ~605→1000 行 | 实际代码量比估计大 | ≤1000 约束通过 |

所有偏差均为架构文档估计误差或合理技术拆分，非设计偏离。

## 九、验收结论

| 检查项 | 结果 |
|---|---|
| 1. 实现覆盖架构文档定义的文件清单 | ✅ 架构指定文件全部存在，额外文件均为合理拆分 |
| 2. 功能对齐上游 commit 6f70b92 | ✅ 引擎层精确对齐，布局/辅助/编辑层拆分后行数一致 |
| 3. import 路径正确（相对路径） | ✅ docx-core 全部相对路径，vue-docx 使用包名 |
| 4. typecheck 通过 | ✅ 两包均零错误 |
| 5. 无残留 stub/mock/fake | ✅ 6 个疑似匹配项均已核实为合法或非 stub |
| 6. ≤1000 行约束 | ✅ 全部通过，selection-helpers.ts 恰好 1000 行 |
| 7. 零 React 残留 | ✅ docx-core 零 React import/hook/类型 |
| 8. build 通过 | ✅ 两包均构建成功 |

**总体评估：pass。** 开发产出完整对齐架构文档和上游 commit 6f70b92，无阻塞性问题。
