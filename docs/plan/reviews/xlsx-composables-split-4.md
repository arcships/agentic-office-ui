# Review: xlsx-composables-split-4

Date: 2026-07-08
Task: xlsx-composables-split（架构设计 `docs/xlsx-migration-architecture.md` §2.2 / §三 步骤4 / §4.4）
Branch: task/docx-remigration
HEAD: ff0fa44
前置 review: `xlsx-composables-split-1.md`（blocked）、`-2.md`（blocked）、`-3.md`（blocked）

## 结论

**blocked** —— 本任务的核心目标（把 4724 行的 `composables.ts` 按职责物理拆分到 ≤1000 行的多个模块、并把 monolith 改为从子模块 import）依然未达成。HEAD `ff0fa44` 相比 review-3 多产出了 6 个子文件（9 文件中的 7 个非 barrel/非 controller 文件已存在），但：

1. **monolith `composables.ts` 一字未改，仍为 4724 行**（硬约束的 4.7 倍，`git diff 914001a..HEAD -- composables.ts` 为空）。
2. **monolith 不从任何子模块 import**，9 个子文件全是 monolith 既有逻辑的重复副本（函数体逐字相同，仅多了 `export` 关键字），运行时仍走 monolith 内的本地副本，子文件是死代码。
3. **typecheck 失败**（2 个编译错误），子模块自身都不可编译。
4. **无 `composables/index.ts` barrel**，`src/index.ts` 仍从 `./composables`（即 monolith `composables.ts`）导入，子模块无法被外部引用。
5. **架构要求的 `useXlsxViewerController.ts`（核心 controller 文件）缺失**，`useXlsxViewerController` 仍在 monolith:1899。
6. **存在 value-level 循环依赖**（`workbook-state` ↔ `clipboard`、`workbook-state` ↔ `editing`），即使接入也会有初始化顺序风险。

这是连续第四次 review 同一核心问题：monolith 没有被拆分，只是被复制。问题性质与 review-3 完全一致（"复制而非剪切迁移"），仅复制规模从 3 文件扩大到 9 文件。

## Findings

### F1 — monolith 未缩减，仍为 4724 行，硬约束 ≤1000 行未满足（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §硬约束（"单文件 ≤ 1000 行"）+ §2.2（composables/ 目录 9 文件清单）+ §4.4（composables 拆分策略）
- 代码位置：`packages/vue-xlsx/src/composables.ts:1-4724`

`composables.ts` 自 `914001a`（task/xlsx-003 机械改写产物）以来从未被修改：

```
$ git diff 914001a..HEAD -- packages/vue-xlsx/src/composables.ts
（空）
$ wc -l packages/vue-xlsx/src/composables.ts
4724 packages/vue-xlsx/src/composables.ts
```

4724 行是硬约束的 4.7 倍。本任务存在的唯一理由就是拆掉这个 monolith；monolith 未缩减一行，即 blocking。

### F2 — 架构要求的 9 文件只产出 7 个，缺核心 controller 文件与 barrel（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（composables/ 目录清单）
- 代码位置：`packages/vue-xlsx/src/composables/`

§2.2 要求的 9 文件对账：

| 设计文件（§2.2） | 预估行数 | 实际状态 |
|---|---:|---|
| `composables/useXlsxViewerController.ts` | ~700 | ❌ **缺失**（核心 controller `useXlsxViewerController` 仍在 monolith:1899） |
| `composables/workbook-state.ts` | ~600 | ⚠️ 存在（1002 行，超 1000，见 F7），但孤立未接入 |
| `composables/selection.ts` | ~500 | ⚠️ 存在（170 行），孤立未接入 |
| `composables/editing.ts` | ~600 | ⚠️ 存在（465 行），孤立未接入 |
| `composables/chart-controller.ts` | ~500 | ⚠️ 存在（620 行），孤立未接入 |
| `composables/clipboard.ts` | ~400 | ⚠️ 存在（492 行），孤立未接入 |
| `composables/navigation.ts` | ~400 | ⚠️ 存在（433 行），孤立未接入 |
| `composables/formatting.ts` | ~500 | ⚠️ 存在（285 行），孤立未接入 |
| `composables/index.ts` | barrel | ❌ **缺失**（无 barrel，子模块无法被 `src/index.ts` 引用） |

```
$ ls packages/vue-xlsx/src/composables/
chart-controller.ts  clipboard.ts  editing.ts  formatting.ts
image-assets.ts  internal.ts  navigation.ts  selection.ts  workbook-state.ts
（无 index.ts、无 useXlsxViewerController.ts）
```

`useXlsxViewerController.ts` 是 §2.2 清单的第一个文件、§4.4 明确的拆分核心（"useXlsxViewerController(4724行)按职责拆"），其缺失意味着拆分的主干未动。

### F3 — 子文件全是 monolith 的重复副本，monolith 未改为 import，子文件是死代码（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2 + §4.4（拆分后需接入）
- 代码位置：`packages/vue-xlsx/src/composables/*.ts`（9 文件）；`packages/vue-xlsx/src/composables.ts`（monolith）；`packages/vue-xlsx/src/index.ts`

证据链：

1. monolith 不从任何子模块 import：
   ```
   $ grep -n 'from "\./composables/' packages/vue-xlsx/src/composables.ts
   （无输出）
   ```
2. monolith 仍保留全部同名函数的本地定义（子文件是重复副本而非"剪切迁移"）：
   ```
   $ for s in escapeHtml resolveInheritedCellStyle buildSheetList applyCellMutationState \
             loadWorkbookImageAssets shouldSkipXmlParsingForWorkbook coerceUserEnteredValue downloadArrayBuffer; do
       grep -cE "function $s\b|const $s\b" packages/vue-xlsx/src/composables.ts
   done  # 每个均输出 1，即 monolith 内副本仍在
   ```
3. 函数体逐字相同（仅子文件多了 `export`）：
   - monolith `escapeHtml` 与 `composables/formatting.ts:227` 的 `escapeHtml` 函数体完全一致。
4. 全仓库无任何文件 import 子模块：
   ```
   $ grep -rn 'composables/' packages/vue-xlsx/src/   # 无输出
   ```
5. `src/index.ts:2-5` 仍 `from "./composables"`，解析到 monolith `composables.ts`，子模块无人引用。

正确做法是"从 monolith 移出 → monolith 删除本地实现并改为 import"，最终 `composables.ts` 应消失或仅剩极薄入口。当前是"复制一份留着"，与 review-3 的问题性质完全相同。

### F4 — typecheck 失败：2 个编译错误，子模块自身不可编译（typecheck 失败）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §三 步骤4（typecheck 要求）+ §风险（"拆分后全量 typecheck + build"）
- 代码位置：`packages/vue-xlsx/src/composables/clipboard.ts:11`；`packages/vue-xlsx/src/composables/workbook-state.ts:186`

```
$ pnpm --filter @arcships/vue-xlsx typecheck
src/composables/clipboard.ts(11,30): error TS2440: Import declaration conflicts with local declaration of 'escapeHtml'.
src/composables/workbook-state.ts(186,18): error TS2304: Cannot find name 'XlsxThemePalette'.
ERR_PNPM_RECURSIVE_RUN_FIRST_FAIL  Exit status 2
```

两处错误：

1. `clipboard.ts:11` 从 `./formatting` 导入 `escapeHtml`，又在 `clipboard.ts:19` 本地重新声明 `export function escapeHtml`，二者冲突（TS2440）。这是复制 monolith 代码时未删本地副本的典型痕迹。
2. `workbook-state.ts:186` 使用 `XlsxThemePalette` 类型但未 import（TS2304）。该类型在 `@arcships/xlsx-core` 导出，import 块漏了它。

review-3 的 typecheck 通过（仅覆盖 monolith）；本次因新增子文件被 tsc 扫描而暴露出子模块自身的类型错误，说明子模块从未经过编译验证。

### F5 — value-level 循环依赖：`workbook-state` ↔ `clipboard`、`workbook-state` ↔ `editing`（模块边界缺陷）

- 严重程度：**P2**
- 阻塞：**blocking**（即便修复 F1-F4 接入，循环依赖会引入初始化顺序 / TDZ 风险）
- 设计文档位置：`docs/xlsx-migration-architecture.md` §4.4（按职责拆分，隐含模块单向依赖）+ §风险（"composables 拆分可能破坏内部依赖"）
- 代码位置：见下

import 图（value import 用 `→`，type-only 用 `⇢`）：

```
workbook-state → editing       (value: applyCellMutationState)        editing.ts:17
editing         → workbook-state (value: buildSheetList, buildVisibleSheetIndexMap) workbook-state.ts(被引)
⇒ workbook-state ↔ editing  value 循环

workbook-state → clipboard     (value: downloadArrayBuffer/.../downloadUrl)  workbook-state.ts:67
clipboard       → workbook-state (value: resolveInheritedCellStyle)          clipboard.ts:12
⇒ workbook-state ↔ clipboard value 循环

navigation      → chart-controller (value: loadWorkbookImageAssets, shouldSkipXmlParsingForWorkbook)  navigation.ts:56
chart-controller ⇢ navigation      (type only: XlsxControllerContext)  chart-controller.ts:28
⇒ navigation → chart-controller 单向（chart-controller 仅 type 反向，可接受）
```

`workbook-state` 与 `clipboard`/`editing` 的双向 value import 是真循环：模块加载时若一方在另一方之前求值，被引用的绑定可能处于 TDZ（`const`/`function` 提升有限）。这违背 §4.4 "按职责拆"隐含的模块单向依赖。正确做法是把跨模块共享的基础设施（如 `XlsxControllerContext`、共享常量、`resolveInheritedCellStyle` 这类被多方依赖的纯函数）下沉到 `internal.ts`/独立层，打破循环。

### F6 — 上游功能对齐：monolith 自身对齐良好，拆分产出的对齐无法验证（对齐验证不可进行）

- 严重程度：**P2**
- 阻塞：**non-blocking**（对齐主体 monolith 未变，blocking 来自 F1-F5）
- 设计文档位置：`docs/upstream-xlsx-feature-alignment.md` §2.9（controller.tsx 对齐清单 A–L）+ §三（功能对齐要点 22–38）
- 代码位置：`packages/vue-xlsx/src/composables.ts`

上游 `@extend-ai/react-xlsx` commit `f285a1c` 的 `controller.tsx` 核心逻辑在 monolith 中均存在（review-3 已确认，本次 monolith 未变）：双层历史（`pushHistoryEntry`/`applyCellMutationState`/`isApplyingHistoryRef`）、`coerceUserEnteredValue`、`maybeRecalculateWorkbook`+`tryRecalculate`、`sanitizeSavedWorkbookBytes`、`loadWorkbookChartAssets`、`buildSheetList`、`resolveInheritedCellStyle`、`clampZoomScale`/`zoomScaleOverridesByTabId`、`sortTable` 等。

但本任务（拆分）的对齐验证无法进行：拆分未真正发生（monolith 未改、子文件是死代码），无法确认 17 项关键逻辑（A–L）在拆分后是否完整保留、模块边界是否割裂内聚逻辑（如历史模型的 snapshot/cell-edit/range-edit 三类若被分到不同文件需谨慎处理共享状态）。子文件内容是 monolith 的逐字副本，逻辑上与上游对齐等价于 monolith 对齐，但因 F5 循环依赖 + F4 编译错误，子文件的模块边界本身尚不可用。

`composables.ts` 与 `composables/*.ts` 无 stub/mock/fake 残留（grep 无命中）。唯一 stub 字样在 `src/index.ts:10` 的 `XlsxViewer` 占位，属 `xlsx-components` 任务范畴，非本任务范围。

### F7 — `workbook-state.ts` 1002 行，超硬约束 ≤1000 行（硬约束违反）

- 严重程度：**P2**
- 阻塞：**non-blocking**（仅超 2 行，且该文件当前孤立未被接入；接入前可一并修）
- 设计文档位置：`docs/xlsx-migration-architecture.md` §硬约束（"单文件 ≤ 1000 行"）+ §2.2（`workbook-state.ts` ~600）
- 代码位置：`packages/vue-xlsx/src/composables/workbook-state.ts`

```
$ wc -l packages/vue-xlsx/src/composables/workbook-state.ts
1002 packages/vue-xlsx/src/composables/workbook-state.ts
```

设计预估 ~600 行，实际 1002，超硬约束 2 行。即便完成接入，此文件仍需进一步拆分（如把 `resolveInheritedCellStyle`/样式继承单独抽出，或把 `buildSheetList`/`buildVisibleSheetIndexMap` 归入 selection/workbook-state 边界调整）。

### F8 — 设计未提及的自加模块 `image-assets.ts`、`internal.ts`（non-blocking）

- 严重程度：P3
- 阻塞：**non-blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（清单无此二文件）
- 代码位置：`packages/vue-xlsx/src/composables/image-assets.ts`（704 行）；`packages/vue-xlsx/src/composables/internal.ts`（208 行）

§2.2 的 8 模块按职责（state/selection/editing/chart/clipboard/navigation/formatting/controller）划分，未列 `image-assets.ts` 与 `internal.ts`。

- `internal.ts`：聚合跨模块共享的常量（`FORMULA_COUNT_THRESHOLD`、`DEFAULT_ROW_HEIGHT`、`EMU_PER_PIXEL`、`HISTORY_LIMIT` 等）、历史类型、idle/abort 工具。合理工程决策，设计未覆盖属细节补充。review-3 已标注此点，本次仍未在设计文档补说明。
- `image-assets.ts`（704 行，本次新增）：把图片资产解析/锚点/resize 等从 monolith 抽出。§2.2 把图片相关逻辑归在 `chart-controller`（图表 controller）与未来 `render/` 层，单列 `image-assets.ts` 是对职责的更细划分。合理，但同样未在设计文档登记。

建议：若重做拆分时保留此二文件，在设计文档 §2.2 补两行说明其职责，避免清单与实现长期背离。

## 证据汇总

```
$ git rev-parse HEAD
ff0fa44
$ git diff --stat 03585c3..ff0fa44 -- packages/vue-xlsx/
 .../vue-xlsx/src/composables/chart-controller.ts   | 620 ++++++++++
 .../vue-xlsx/src/composables/clipboard.ts          | 492 ++++++++++
 .../vue-xlsx/src/composables/editing.ts            | 465 +++++++++
 .../vue-xlsx/src/composables/image-assets.ts       | 704 +++++++++++++
 .../vue-xlsx/src/composables/navigation.ts         | 433 +++++++++
 .../vue-xlsx/src/composables/workbook-state.ts     |1002 +++++++++++++++++++
 6 files changed, 3716 insertions(+)
（composables.ts 未出现在 diff）

$ wc -l packages/vue-xlsx/src/composables.ts packages/vue-xlsx/src/composables/*.ts
4724 composables.ts
 620 chart-controller.ts
 492 clipboard.ts
 465 editing.ts
 285 formatting.ts
 704 image-assets.ts
 208 internal.ts
 433 navigation.ts
 170 selection.ts
1002 workbook-state.ts

$ grep -n 'from "\./composables/' packages/vue-xlsx/src/composables.ts   # monolith 不引用子模块
（无输出）
$ grep -rn 'composables/' packages/vue-xlsx/src/                         # 子模块无人引用
（无输出）
$ ls packages/vue-xlsx/src/composables/index.ts packages/vue-xlsx/src/composables/useXlsxViewerController.ts
ls: No such file or directory   (index.ts)
ls: No such file or directory   (useXlsxViewerController.ts)

$ pnpm --filter @arcships/vue-xlsx typecheck
src/composables/clipboard.ts(11,30): error TS2440: Import declaration conflicts with local declaration of 'escapeHtml'.
src/composables/workbook-state.ts(186,18): error TS2304: Cannot find name 'XlsxThemePalette'.
Exit status 2
```

## 与 review-1 / review-2 / review-3 的差异

- review-1（HEAD `914001a`）、review-2（HEAD `2dd846f`）：拆分产出完全缺失，`composables/` 为空。
- review-3（HEAD `03585c3`）：`composables/` 新增 3 文件（`formatting`/`selection`/`internal`），均为 monolith 孤立重复副本，未接入、无 barrel、monolith 未缩减。blocked。
- review-4（HEAD `ff0fa44`）：`composables/` 新增至 9 文件（含本次 +6），但全部为 monolith 重复副本、孤立、无 barrel、monolith 仍 4724 行。新增问题：typecheck 失败（F4）、value-level 循环依赖（F5）、`workbook-state.ts` 超 1000 行（F7）。blocking 状态未改变，且因子文件被 tsc 扫描暴露了更多缺陷。

连续四次 review 同一核心问题（monolith 未拆、子文件是复制副本）未解决。

## 建议下一步

重做 `xlsx-composables-split`，采用"剪切迁移"而非"复制"，并修复本次暴露的编译/循环依赖问题：

1. **剪切迁移**：按 §2.2 的 8 模块把 `composables.ts` 代码**移出**（state/selection/editing/chart/clipboard/navigation/formatting/controller），monolith 删除本地实现并改为从子模块 `import`，最终 `composables.ts` 消失或仅剩极薄入口。
2. **补 controller 文件**：`composables/useXlsxViewerController.ts`（§2.2 首项），把 `useXlsxViewerController`（monolith:1899）及其直接 setup 逻辑移入。
3. **补 barrel**：`composables/index.ts`，`src/index.ts` 从 `./composables` 导入（接入后生效）。
4. **修复 F4 编译错误**：`clipboard.ts` 删本地 `escapeHtml`（改用 `./formatting` 导入）；`workbook-state.ts` 补 `XlsxThemePalette` 的 import。
5. **打破 F5 循环依赖**：把被多方依赖的共享基础（`XlsxControllerContext`、`resolveInheritedCellStyle` 等纯函数、共享常量）下沉到 `internal.ts` 或新设的 `context.ts`，使 `workbook-state`/`clipboard`/`editing` 间无 value 循环。
6. **每文件 ≤1000 行**：`workbook-state.ts` 当前 1002 行，拆分时进一步切（如样式继承单独成文）。
7. **验证**：`pnpm --filter @arcships/vue-xlsx typecheck` 必须通过；`pnpm --filter @arcships/vue-xlsx build` 后运行 `packages/vue-xlsx/test/structure.mjs` 验证 controller 方法表完整；并 grep 确认 monolith 内不再保留已迁出函数的本地副本。
8. 依赖任务 `xlsx-core-charts-split`/`xlsx-core-images-split` 已完成，composables 从 `@arcships/xlsx-core` 统一入口导入（当前 monolith 已如此，拆分时保持）。
