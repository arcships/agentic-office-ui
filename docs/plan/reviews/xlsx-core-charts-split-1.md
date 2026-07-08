# Review: xlsx-core-charts-split

Date: 2026-07-08
Reviewer: agent
Scope: `packages/xlsx-core/src/charts/` — charts.ts (4366行) 拆分为 6 个子文件 + barrel

## 参考文档

- 架构设计：`docs/xlsx-migration-architecture.md` "2.1 xlsx-core/ 拆分" 节
- 上游对齐：`docs/upstream-xlsx-feature-alignment.md` "2.7 charts.ts" 节
- 上游源码：`@extend-ai/react-xlsx` commit `f285a1c`，`charts.ts` (4366行)

## 文件清单对照

| 架构设计 | 设计行数 | 实际文件 | 实际行数 | 状态 |
|---|---|---|---|---|
| `charts/chart-parser.ts` | ~700 | ✅ 存在 | 1215 (非空 1129) | ⚠️ 超限 |
| `charts/chart-series.ts` | ~600 | ✅ 存在 | 661 | ✅ |
| `charts/chart-colors.ts` | ~500 | ✅ 存在 | 362 | ✅ |
| `charts/chart-types.ts` | ~800 | ✅ 存在 | 811 | ✅ |
| `charts/chart-styles.ts` | ~800 | ✅ 存在 | 629 | ✅ |
| `charts/chart-export.ts` | ~500 | ✅ 存在 | 930 (非空 850) | ✅ |
| `charts/index.ts` | barrel | ✅ 存在 | 10 | ✅ |

原 `charts.ts` 已删除，无残留引用。xlsx-core/index.ts 和 xlsx-worker.ts 的 import 路径已更新为 `"./charts"`（指向 barrel）。

## Findings

### F1 (P1, blocking) — `chart-parser.ts` 超出单文件 1000 行硬约束

- **位置**：`packages/xlsx-core/src/charts/chart-parser.ts` (1215 行, 非空行 1129)
- **设计文档**：`docs/xlsx-migration-architecture.md` 硬约束 "单文件 ≤ 1000 行"，该文件设计估计 ~700 行
- **分析**：`loadWorkbookChartAssets`（~200 行）及其辅助函数（`parseChartCacheValues`、`parseChartMultiLevelCacheValues`、XML 工具函数如 `ensureChild`/`setLeafValue` 等）集中在 chart-parser.ts 中。可将 XML 工具函数抽取到独立的 `chart-xml-utils.ts`，或将 chart cache 解析拆到单独文件。
- **修复建议**：将 `parseChartCacheValues` + `parseChartMultiLevelCacheValues`（~100 行）和 XML DOM 工具函数（`ensureChild`、`setLeafValue`、`setBooleanValue`、`setNumericValue`、`getLocalChildren` 等，~100 行）抽取到独立子文件，或在设计文档中明确标注 chart-parser.ts 可超 1000 行（类似 surface-regl 的例外）。

### F2 (P2, non-blocking) — 模块间存在循环依赖

- **代码位置**：`packages/xlsx-core/src/charts/` 下多个文件
- **设计文档**：`docs/xlsx-migration-architecture.md` 无循环依赖相关约束，但 `docs/INDEX.md` 隐含要求模块边界清晰
- **具体循环**：
  - `chart-parser.ts` ↔ `chart-types.ts`（双向 value import）
  - `chart-parser.ts` ↔ `chart-styles.ts`（双向 value import）
  - `chart-types.ts` → `chart-export.ts` → `chart-styles.ts` → `chart-parser.ts` → `chart-types.ts`（长链）
- **分析**：源自分拆单体文件 chart.ts（4366行），原文件内部函数相互调用，拆分后形成跨文件循环。所有 import 均为 value import（非 `import type`），在 ES module 环境下可能导致初始化顺序问题，但当前 typecheck 通过说明未触发实际错误。功能正确性需在运行时验证。
- **修复建议**：将 `chart-types.ts` 对 `chart-export.ts`、`chart-styles.ts`、`chart-parser.ts` 的 value import 评估是否可以改为 type import；或将公共依赖提取到独立的 util 文件打破循环。

### F3 (P3, non-blocking) — `normalizeWorksheetVisibility` 放置在 chart-colors.ts

- **位置**：`packages/xlsx-core/src/charts/chart-colors.ts:11`
- **设计文档**：架构设计将 chart-colors.ts 定义为"颜色/主题"
- **分析**：该函数是 worksheet visibility 解析（hidden/veryHidden/visible），由 chart-export.ts 的 `buildTabs` 使用。功能正确，但在 chart-colors 模块中语义不匹配。上游 charts.ts 也将其放在同一文件中。
- **修复建议**：可移至 `chart-export.ts` 或 `chart-parser.ts`，或忽略（上游原有分组）。

### F4 (P3, non-blocking) — 文件行数分布偏离设计估计

- **位置**：整体 `packages/xlsx-core/src/charts/`
- **设计文档**：`docs/xlsx-migration-architecture.md` 第二节行数估计
- **偏差**：chart-parser +515 行（超限见 F1），chart-export +430 行（930，接近上限），chart-colors -138 行（362），chart-styles -171 行（629）。总行数 4618 与原始 4366 的差异（+252）来自拆分后的跨文件 import 声明。
- **分析**：拆分粒度与设计估计存在偏差，但除 F1 外均在可接受范围。总行数增长 5.8% 合理。

## 验证结果

| 检查项 | 结果 | 说明 |
|---|---|---|
| typecheck (`tsc --noEmit`) | ✅ 通过 | 零错误 |
| stub/mock/fake 残留 | ✅ 无 | 搜索全 charts/ 目录无匹配 |
| import 路径（相对路径） | ✅ 正确 | 全部使用 `./` 或 `../` 相对路径 |
| Barrel 导出完整性 | ✅ 完整 | charts/index.ts 覆盖上游 charts.ts 的全部 9 个 export |
| 上游功能对齐 | ✅ 对齐 | 拆分保持原有逻辑不变，无功能删减 |
| 单文件 ≤1000 行 | ❌ F1 | chart-parser.ts 1215 行 |

## 结论

**blocked** — F1 违反架构设计硬约束（单文件 ≤1000 行），需要将 chart-parser.ts 拆分至 1000 行以内或在设计文档中标注例外。

F2（循环依赖）为拆分引入的结构问题，建议在拆分 F1 时一并解决。
