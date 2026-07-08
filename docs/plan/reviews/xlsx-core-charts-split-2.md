# Review: xlsx-core-charts-split (Round 2)

Date: 2026-07-08
Reviewer: agent
Scope: `packages/xlsx-core/src/charts/` — charts.ts (4366行) 拆分为子模块，修复 review-1 的 F1

## 参考文档

- 架构设计：`docs/xlsx-migration-architecture.md` "2.1 xlsx-core/ 拆分" 节
- 上游对齐：`docs/upstream-xlsx-feature-alignment.md` "2.7 charts.ts" 节
- 上游源码：`@extend-ai/react-xlsx` commit `f285a1c`，`charts.ts` (4366行)
- 前次 review：`docs/plan/reviews/xlsx-core-charts-split-1.md`

## Review-1 F1 修复验证

Review-1 F1（chart-parser.ts 1215 行超限）已修复：从 chart-parser.ts 中抽取 `parseChartCacheValues` + `parseChartMultiLevelCacheValues` → `chart-cache.ts` (112行)，XML DOM 工具函数 → `chart-xml-utils.ts` (158行)。chart-parser.ts 降至 966 行（≤1000）。

## 文件清单对照

| 架构设计 | 设计行数 | 实际文件 | 实际行数 | 状态 |
|---|---|---|---|---|
| `charts/chart-parser.ts` | ~700 | ✅ 存在 | 966 | ✅ |
| `charts/chart-series.ts` | ~600 | ✅ 存在 | 661 | ✅ |
| `charts/chart-colors.ts` | ~500 | ✅ 存在 | 362 | ✅ |
| `charts/chart-types.ts` | ~800 | ✅ 存在 | 814 | ✅ |
| `charts/chart-styles.ts` | ~800 | ✅ 存在 | 635 | ✅ |
| `charts/chart-export.ts` | ~500 | ✅ 存在 | 930 | ✅ |
| `charts/index.ts` | barrel | ✅ 存在 | 10 | ✅ |
| — | — | `chart-cache.ts` | 112 | ⚠️ 架构未列出 |
| — | — | `chart-xml-utils.ts` | 158 | ⚠️ 架构未列出 |

原 `charts.ts` 已删除。`xlsx-core/src/index.ts` 和 `xlsx-core/src/xlsx-worker.ts` 的 import 路径均为 `"./charts"`（指向 barrel），正确。

## Findings

### F1 (P2, non-blocking) — 新增 chart-cache.ts、chart-xml-utils.ts 未在架构文档中列出

- **设计文档**：`docs/xlsx-migration-architecture.md` 2.1 节 charts/ 文件清单未包含这两个文件
- **代码位置**：
  - `packages/xlsx-core/src/charts/chart-cache.ts` (112行)
  - `packages/xlsx-core/src/charts/chart-xml-utils.ts` (158行)
- **分析**：这两个文件是为修复 review-1 F1（chart-parser.ts 超 1000 行）而从 chart-parser.ts 抽取的。功能归属合理：chart-cache 处理 `parseChartCacheValues`/`parseChartMultiLevelCacheValues`（XML 缓存数据提取），chart-xml-utils 处理 `getLocalChildren`/`ensureChild`/`parseXml` 等 DOM 工具。但架构文档未同步更新以反映这两个文件。
- **修复建议**：更新 `docs/xlsx-migration-architecture.md` 2.1 节，在 charts/ 文件清单中追加 chart-cache.ts 和 chart-xml-utils.ts。

### F2 (P2, non-blocking) — 模块间循环依赖未解决

- **代码位置**：`packages/xlsx-core/src/charts/` 下多个文件
- **设计文档**：`docs/xlsx-migration-architecture.md` 无循环依赖约束，但合理代码组织应避免
- **具体循环（与 review-1 F2 相同）**：
  - `chart-parser.ts` ↔ `chart-types.ts`（双向 value import：chart-parser 导入 `normalizeChartExChart` 等，chart-types 导入 `normalizeLegendPosition`）
  - `chart-parser.ts` ↔ `chart-styles.ts`（双向 value import：chart-parser 导入 `applyChartSeriesStyleFromXml` 等，chart-styles 导入 `parseChartTypeFromXml` 等）
  - `chart-types.ts` → `chart-export.ts` → `chart-styles.ts` → `chart-parser.ts` → `chart-types.ts`（长链）
- **分析**：源自分拆单体 charts.ts (4366行)，原文件内部函数相互调用，拆分后形成跨文件循环。所有 import 均为 value import（非 `import type`），在 ES module 环境下可能导致初始化顺序问题，但当前 typecheck 通过说明未触发实际错误。
- **修复建议**：将 `chart-types.ts` → `chart-parser.ts` 的 `normalizeLegendPosition` import 评估迁移至 chart-parser（该函数使用者均在 chart-parser/export）；或将其提取到 chart-xml-utils 打破循环。

### F3 (P3, non-blocking) — `normalizeWorksheetVisibility` 语义位置偏差

- **位置**：`packages/xlsx-core/src/charts/chart-colors.ts:11`
- **设计文档**：架构设计将 chart-colors.ts 定义为"颜色/主题"模块
- **分析**：该函数解析 worksheet visibility（hidden/veryHidden/visible），仅在 chart-export.ts 的 `buildTabs` 中使用。放在 chart-colors.ts 与其"颜色/主题"定位不匹配。上游 charts.ts 也将其放在同一位置，属于历史遗留。
- **修复建议**：可移至 chart-export.ts 或忽略（上游原有分组，无功能影响）。

### F4 (P3, non-blocking) — 文件行数与设计估计偏差

- **位置**：整体 `packages/xlsx-core/src/charts/`
- **设计文档**：`docs/xlsx-migration-architecture.md` 2.1 节行数估计
- **偏差**：chart-parser +266（966 vs ~700），chart-export +430（930 vs ~500），chart-colors -138（362 vs ~500），chart-styles -165（635 vs ~800）。新增 chart-cache.ts (112) + chart-xml-utils.ts (158) 未在设计估计中。总行数 4648 vs 原始 4366（+282，+6.5%，来源：跨文件 import + barrel + 新增两个子文件）。
- **分析**：拆分粒度与设计估计存在偏差，但所有单文件均 ≤1000 行。总行数增长 6.5% 合理。

## 验证结果

| 检查项 | 结果 | 说明 |
|---|---|---|
| typecheck (`tsc --noEmit`) | ✅ 通过 | 零错误 |
| stub/mock/fake 残留 | ✅ 无 | 搜索 charts/ 目录，零 `TODO`/`FIXME`/stub/mock/fake |
| import 路径（相对路径） | ✅ 正确 | 全部使用 `./` 或 `../` 相对路径，无绝对路径引用 |
| Barrel 导出完整性 | ✅ 完整 | charts/index.ts 覆盖上游 charts.ts 的全部 9 个 export |
| 上游功能对齐 | ✅ 对齐 | 拆分保持原有逻辑不变，`loadWorkbookChartAssets` 等核心函数完整 |
| 单文件 ≤1000 行 | ✅ 通过 | 最大值 chart-export.ts (930) |
| review-1 F1 修复 | ✅ 已修复 | chart-parser.ts 1215→966 |
| 旧 charts.ts 残留 | ✅ 已删除 | `packages/xlsx-core/src/charts.ts` 不存在 |
| 消费方引用 | ✅ 正确 | xlsx-worker.ts 和 index.ts 引用 `"./charts"` 指向 barrel |

## 结论

**pass** — review-1 的阻塞性问题 F1（chart-parser.ts 超 1000 行）已通过抽取 chart-cache.ts + chart-xml-utils.ts 修复，所有单文件 ≤1000 行，typecheck 通过，上游功能对齐完整。F1（架构文档未列出新增子文件）和 F2（循环依赖）为非阻塞问题，建议后续迭代修复。
