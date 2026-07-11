# Review: xlsx-composables-split-5

Date: 2026-07-08
Task: xlsx-composables-split（架构设计 `docs/xlsx-migration-architecture.md` §2.2 / §三 步骤4 / §4.4）
Branch: task/docx-remigration
HEAD: b3417fa
前置 review: `xlsx-composables-split-1.md`（blocked）、`-2.md`（blocked）、`-3.md`（blocked）、`-4.md`（blocked）

## 结论

**blocked** —— 本任务的核心目标（把 4724 行的 `composables.ts` 按职责物理拆分到 ≤1000 行的多个模块、并把 monolith 改为从子模块 import）依然未达成。HEAD `b3417fa`（"fix #4"）相比 review-4（`ff0fa44`）仅修复了 review-4 F4 的两处子模块编译错误并打破了 F5 的循环依赖、把 `workbook-state.ts` 压到 987 行，但本任务存在的唯一理由——拆掉 monolith——一行都没有动：

1. **monolith `composables.ts` 仍为 4724 行，自 `914001a` 起从未修改**（硬约束的 4.7 倍，`git diff 914001a..HEAD -- composables.ts` 为空）。
2. **monolith 不从任何子模块 import**，子模块是 monolith 既有逻辑的重复副本 + 全新平行架构（domain factory + `XlsxControllerContext`），运行时仍走 monolith 内联逻辑，子文件是死代码。
3. **无 `composables/index.ts` barrel**，`src/index.ts` 仍从 `./composables`（即 monolith `composables.ts`）导入，子模块无法被外部引用。
4. **架构要求的 `useXlsxViewerController.ts`（核心 controller 文件）缺失**，`useXlsxViewerController` 仍在 monolith:1899，其 return（monolith:4606-4723）直接引用内联局部函数，从不调用任何 `createXxxDomain(ctx)` factory。

这是连续第五次 review 同一核心问题：monolith 没有被拆分，只是被复制并叠加了一套未被使用的平行抽象。review-4 的 6 条建议只有第 4、5、6 条（编译错误、循环依赖、行数）被部分落实，第 1、2、3 条（剪切迁移、补 controller 文件、补 barrel）——即拆分本体——完全未动。

本次相比 review-4 的实质改进（均为子模块内部质量，不改变 blocking 状态）：
- F4 编译错误已修复（typecheck 通过）。
- F5 value-level 循环依赖已打破，子模块 import 图现为 DAG。
- `workbook-state.ts` 从 1002 → 987 行，回到 ≤1000。

## Findings

### F1 — monolith 未缩减，仍为 4724 行，硬约束 ≤1000 行未满足（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §硬约束（"单文件 ≤ 1000 行"）+ §2.2（composables/ 目录 9 文件清单）+ §4.4（composables 拆分策略）
- 代码位置：`packages/vue-xlsx/src/composables.ts:1-4724`

```
$ git diff --stat 914001a..HEAD -- packages/vue-xlsx/src/composables.ts
（空）
$ wc -l packages/vue-xlsx/src/composables.ts
4724 packages/vue-xlsx/src/composables.ts
```

4724 行是硬约束的 4.7 倍。本任务存在的唯一理由就是拆掉这个 monolith；monolith 未缩减一行，即 blocking。

### F2 — 架构要求的 9 文件仍缺 2 个核心文件：controller 文件与 barrel（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（composables/ 目录清单）
- 代码位置：`packages/vue-xlsx/src/composables/`（无 `index.ts`、无 `useXlsxViewerController.ts`）

§2.2 要求的 9 文件对账：

| 设计文件（§2.2） | 预估行数 | 实际状态 |
|---|---:|---|
| `composables/useXlsxViewerController.ts` | ~700 | ❌ **缺失**（核心 controller `useXlsxViewerController` 仍在 monolith:1899，return 在 monolith:4606） |
| `composables/workbook-state.ts` | ~600 | ⚠️ 存在（987 行，已回到 ≤1000），但孤立未接入 |
| `composables/selection.ts` | ~500 | ⚠️ 存在（170 行），孤立未接入 |
| `composables/editing.ts` | ~600 | ⚠️ 存在（450 行），孤立未接入 |
| `composables/chart-controller.ts` | ~500 | ⚠️ 存在（621 行），孤立未接入 |
| `composables/clipboard.ts` | ~400 | ⚠️ 存在（482 行），孤立未接入 |
| `composables/navigation.ts` | ~400 | ⚠️ 存在（433 行），孤立未接入 |
| `composables/formatting.ts` | ~500 | ⚠️ 存在（305 行），孤立未接入 |
| `composables/index.ts` | barrel | ❌ **缺失**（无 barrel，子模块无法被 `src/index.ts` 引用） |

```
$ ls packages/vue-xlsx/src/composables/index.ts packages/vue-xlsx/src/composables/useXlsxViewerController.ts
ls: No such file or directory   (index.ts)
ls: No such file or directory   (useXlsxViewerController.ts)
```

`useXlsxViewerController.ts` 是 §2.2 清单的第一个文件、§4.4 明确的拆分核心（"useXlsxViewerController(4724行)按职责拆"），其缺失意味着拆分的主干未动。

### F3 — 子模块是 monolith 的重复副本 + 未被使用的平行架构，monolith 未改为 import，子文件是死代码（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2 + §4.4（拆分后需接入）
- 代码位置：`packages/vue-xlsx/src/composables/*.ts`（9 文件）；`packages/vue-xlsx/src/composables.ts`（monolith）；`packages/vue-xlsx/src/index.ts:2-5`

证据链：

1. monolith 不从任何子模块 import：
   ```
   $ grep -n 'from "\./composables/' packages/vue-xlsx/src/composables.ts
   （无输出）
   ```
2. monolith 仍保留全部同名纯函数的本地定义（子文件是重复副本而非"剪切迁移"）：
   ```
   $ for s in escapeHtml resolveInheritedCellStyle buildSheetList applyCellMutationState \
             loadWorkbookImageAssets shouldSkipXmlParsingForWorkbook coerceUserEnteredValue downloadArrayBuffer; do
       grep -cE "function $s\b" packages/vue-xlsx/src/composables.ts
   done  # 每个均输出 1，即 monolith 内副本仍在
   ```
3. 全仓库无任何文件 import 子模块：
   ```
   $ grep -rn 'from "\./composables/' packages/vue-xlsx/src/   # 无输出
   ```
4. `src/index.ts:2,4` 仍 `from "./composables"`，解析到 monolith `composables.ts`，子模块无人引用。

更关键：子模块并非 monolith 的等价改写，而是引入了一套**全新的平行抽象**——`XlsxControllerContext`（`navigation.ts:63-161`）与 domain factory（`createEditingDomain`/`createNavigationDomain`/`createChartImageDomain`/`createClipboardDomain`/`createHistoryDomain`）。monolith 的 `useXlsxViewerController` 从不构造 `XlsxControllerContext`、从不调用任何 factory，其 return（monolith:4606-4723）直接引用内联局部函数：
```
$ grep -nE "createChartImageDomain|createEditingDomain|createHistoryDomain|createClipboardDomain|createNavigationDomain|XlsxControllerContext" packages/vue-xlsx/src/composables.ts
（none - monolith uses inline logic）
```
即子模块整套设计从未被接入。正确做法是"从 monolith 移出 → monolith 删除本地实现并改为 import"，最终 `composables.ts` 应消失或仅剩极薄入口。

### F4 — typecheck 通过（review-4 F4 已修复）

- 严重程度：—
- 阻塞：**non-blocking**（本项无问题，记录验证结果）
- 设计文档位置：`docs/xlsx-migration-architecture.md` §三 步骤4（typecheck 要求）
- 代码位置：`packages/vue-xlsx/src/composables/clipboard.ts`；`packages/vue-xlsx/src/composables/workbook-state.ts`

review-4 F4 的两处编译错误已修复：
- `clipboard.ts` 删除了本地 `escapeHtml` 声明，改从 `./formatting` 导入（`clipboard.ts:11`）。
- `workbook-state.ts:27` 补齐了 `type XlsxThemePalette` 的 import。

```
$ pnpm --filter @arcships/vue-xlsx typecheck
> tsc --noEmit
（exit 0，无输出）
```

注意：typecheck 通过仅证明 monolith + 子模块都能编译，不能证明拆分已生效（子模块是死代码，见 F3）。

### F5 — 子模块 import 图现为 DAG，review-4 F5 循环依赖已打破（模块边界改善）

- 严重程度：—
- 阻塞：**non-blocking**（本项无问题，记录验证结果）
- 设计文档位置：`docs/xlsx-migration-architecture.md` §4.4
- 代码位置：`packages/vue-xlsx/src/composables/*.ts`

review-4 F5 的两个 value 循环已消除：
- `workbook-state` ↔ `clipboard`：`clipboard` 不再 import `workbook-state`，`resolveInheritedCellStyle` 改从 `./formatting` 导入（`clipboard.ts:11`）。`workbook-state → clipboard`（download*，`workbook-state.ts:68`）现为单向。
- `workbook-state` ↔ `editing`：`workbook-state` 不再 import `editing`，`applyCellMutationState` 已下沉到 `./internal`（`workbook-state.ts:66`）。`editing → workbook-state`（buildSheetList，`editing.ts:18`）现为单向。

当前 import 图（value import 用 `→`，type-only 用 `⇢`）：

```
internal → selection
formatting → internal
selection → (none)
image-assets → internal
clipboard → internal, selection, formatting
editing → selection, internal, workbook-state, ⇢navigation
workbook-state → selection, internal, image-assets, clipboard, ⇢navigation
navigation → (none)
chart-controller → internal, ⇢navigation
```

无环。被多方依赖的共享基础（常量、历史类型、`applyCellMutationState`、`resolveInheritedCellStyle`）已下沉到 `internal.ts`/`formatting.ts`，符合 §4.4 隐含的模块单向依赖。

### F6 — 子模块 `navigation.ts` 被当作 barrel 复用，职责混乱（代码组织）

- 严重程度：**P3**
- 阻塞：**non-blocking**（子模块是死代码，接入前可一并修）
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（`navigation.ts` 职责为"导航/滚动"）
- 代码位置：`packages/vue-xlsx/src/composables/navigation.ts:416-433`

`navigation.ts` 既定义 `XlsxControllerContext`（跨模块共享上下文接口，`navigation.ts:63-161`）与 `createNavigationDomain`，又在文件尾部 re-export 一批与导航无关的符号（`navigation.ts:416-433`）：

```
export {
  buildSheetList, buildVisibleSheetIndexMap, clampZoomScale,
  resolveDefaultZoomScale, resolveNextZoomScale, resolveWorkbookBuffer,
  parseWorkbookBuffer, tryRecalculate, scheduleLowPriorityTask,
  DEFAULT_DEFER_LOADING_ABOVE_BYTES, DEFAULT_MAX_FILE_SIZE_BYTES,
  DEFAULT_ZOOM_TAB_KEY, MAX_ZOOM_SCALE, MIN_ZOOM_SCALE
};
export type { UseXlsxViewerControllerOptions };
```

这些符号定义在 `workbook-state.ts`/`internal.ts`，`navigation.ts` 仅做转发。同时 `XlsxControllerContext` 这一全局上下文接口放在 `navigation.ts` 而非 `internal.ts`/独立 `context.ts`，与 §2.2 对 `navigation.ts` 的职责定位（"导航/滚动"）不符。接入时应：把 `XlsxControllerContext` 下沉到 `internal.ts` 或 `context.ts`；把 barrel 职责集中到 `composables/index.ts`；`navigation.ts` 只保留 `createNavigationDomain`。

### F7 — 子模块残留上游 `[react-xlsx debug]` 调试日志（代码质量）

- 严重程度：**P3**
- 阻塞：**non-blocking**（与 monolith 行为一致，非拆分新引入，但接入前应清理）
- 设计文档位置：—（上游 `controller.tsx` 残留，非设计文档要求）
- 代码位置：`packages/vue-xlsx/src/composables/chart-controller.ts:314,329,349`（对应 monolith:3657,3672,3692）

`chart-controller.ts` 的 `setChartRect` 含 3 处 `console.info("[react-xlsx debug] ...")`，与 monolith 逐字相同，证明子模块是直接复制。这些 debug 日志会进入生产运行时。接入时应删除（monolith 侧同位置也应一并清理，但属于 monolith 既有问题，非本任务范围）。

```
$ grep -n "console\." packages/vue-xlsx/src/composables/chart-controller.ts
314:    console.info("[react-xlsx debug] setChartRect", { ... })
329:    console.info("[react-xlsx debug] currentChart", { ... })
349:    console.info("[react-xlsx debug] updateWorkbookChartAnchor", { ... })
```

### F8 — 子模块 `editing.ts` 的 `setCellFormula` 有 dead code（与 monolith 一致的既有问题）

- 严重程度：**P3**
- 阻塞：**non-blocking**（与 monolith 逐字相同，非拆分新引入）
- 设计文档位置：—
- 代码位置：`packages/vue-xlsx/src/composables/editing.ts:130-134`

```
const trimmedFormula = formula.trim();   // 声明后未使用
if (!formula.trim()) {                   // 重复 trim，未用 trimmedFormula
  worksheet.setCell(cellAddressToA1(cell), "");
} else {
  worksheet.setFormula(cellAddressToA1(cell), formula);  // 用原始 formula，未用 trimmedFormula
}
```

`trimmedFormula` 变量声明后从未使用，`else` 分支用原始 `formula` 而非 `trimmedFormula`。与 monolith 同位置逻辑逐字一致，属上游/机械改写遗留。接入时可顺手修正为 `trimmedFormula`。

### F9 — 上游功能对齐：monolith 对齐良好；拆分产出的对齐因死代码无法验证（对齐验证不可进行）

- 严重程度：**P2**
- 阻塞：**non-blocking**（对齐主体 monolith 未变，blocking 来自 F1-F3）
- 设计文档位置：`docs/upstream-xlsx-feature-alignment.md` §2.9（controller.tsx 对齐清单 A–L）+ §三（功能对齐要点 22–38）
- 代码位置：`packages/vue-xlsx/src/composables.ts`（monolith）；`packages/vue-xlsx/src/composables/*.ts`（死代码）

上游 `@extend-ai/react-xlsx` commit `f285a1c` 的 `controller.tsx` 核心逻辑在 monolith 中均存在（review-4 已确认，本次 monolith 未变）：双层历史（`pushHistoryEntry`/`applyCellMutationState`/`isApplyingHistoryRef`，要点 23-26）、`coerceUserEnteredValue`（要点 27）、paste 区分 formula/value（要点 28-29）、`fillSelection` 平铺填充（要点 30）、`sortTable` 用计算值排序 + 无变化只 setSortState（要点 31-32）、merge/unmerge/resize 走 snapshot 历史（要点 33-34）、`setRangeStyle` 单次 wasm 批量（要点 35）、`maybeRecalculateWorkbook` trap 降级（要点 36）、image/chart rect 走 snapshot + px→EMU（要点 37）、zoom 纯前端按 tab 存（要点 38）。

子模块内容是 monolith 的逐字副本 + domain factory 包装，逻辑上与上游对齐等价于 monolith 对齐；但因 F3 子模块是死代码，拆分后的模块边界是否割裂内聚逻辑（如历史模型的 snapshot/cell-edit/range-edit 三类跨 `workbook-state`/`editing` 分布）无法在运行时验证。

`composables.ts` 与 `composables/*.ts` 无 stub/mock/fake 残留（grep 无命中）。唯一 stub 字样在 `src/index.ts:10` 的 `XlsxViewer` 占位，属 `xlsx-components` 任务范畴，非本任务范围。

## 关于 import 路径

子模块内部的相对 import 路径均正确（typecheck 通过佐证），无绝对路径/包内错误引用。问题在于"接线"缺失：`src/index.ts` 与 monolith 均不引用子模块，故子模块路径正确与否对运行时无意义（见 F3）。

## 证据汇总

```
$ git rev-parse HEAD
b3417fa4694927739d948645f94e0edbbb6881d4

$ git diff --stat 914001a..HEAD -- packages/vue-xlsx/src/composables.ts   # monolith 自 task-003 起未改
（空）

$ wc -l packages/vue-xlsx/src/composables.ts packages/vue-xlsx/src/composables/*.ts
4724 composables.ts
 621 chart-controller.ts
 482 clipboard.ts
 450 editing.ts
 305 formatting.ts
 704 image-assets.ts
 226 internal.ts
 433 navigation.ts
 170 selection.ts
 987 workbook-state.ts

$ grep -n 'from "\./composables/' packages/vue-xlsx/src/composables.ts   # monolith 不引用子模块
（无输出）
$ grep -rn 'from "\./composables/' packages/vue-xlsx/src/                # 子模块无人引用
（无输出）
$ grep -nE "createChartImageDomain|createEditingDomain|createHistoryDomain|createClipboardDomain|createNavigationDomain|XlsxControllerContext" packages/vue-xlsx/src/composables.ts
（none - monolith uses inline logic）
$ ls packages/vue-xlsx/src/composables/index.ts packages/vue-xlsx/src/composables/useXlsxViewerController.ts
ls: No such file or directory   (index.ts)
ls: No such file or directory   (useXlsxViewerController.ts)

$ pnpm --filter @arcships/vue-xlsx typecheck
> tsc --noEmit
（exit 0）
```

## 与历次 review 的差异

| review | HEAD | composables/ 文件 | monolith 行数 | 接入 | typecheck | 结论 |
|---|---|---|---:|---|---|---|
| -1 | 914001a | 0 | 4724 | 无 | (monolith) pass | blocked |
| -2 | 2dd846f | 0 | 4724 | 无 | (monolith) pass | blocked |
| -3 | 03585c3 | 3 | 4724 | 无 | (monolith) pass | blocked |
| -4 | ff0fa44 | 9 | 4724 | 无 | fail (2 err) | blocked |
| **-5** | **b3417fa** | **9** | **4724** | **无** | **pass** | **blocked** |

连续五次 review 同一核心问题（monolith 未拆、子文件是复制副本/平行死代码）未解决。本次仅修复了 review-4 暴露的子模块编译错误与循环依赖，拆分本体（剪切迁移 + controller 文件 + barrel + 接入）零进展。

## 建议下一步

review-4 的建议仍然适用，本次仅强调被忽略的部分：

1. **剪切迁移**（review-4 建议 1，未落实）：按 §2.2 的 8 模块把 `composables.ts` 代码**移出**，monolith 删除本地实现并改为从子模块 `import`，最终 `composables.ts` 消失或仅剩极薄入口。当前是"复制 + 加平行架构"，必须改为"剪切"。
2. **补 controller 文件**（review-4 建议 2，未落实）：`composables/useXlsxViewerController.ts`，把 `useXlsxViewerController`（monolith:1899）及其 setup/return（monolith:4606-4723）移入，return 改为聚合各 `createXxxDomain(ctx)` 的产物。
3. **补 barrel + 接线**（review-4 建议 3，未落实）：`composables/index.ts`，`src/index.ts` 从 `./composables`（解析到 barrel）导入。接入后 F3 死代码问题自然消除。
4. 清理 F6（`navigation.ts` 的 barrel/上下文职责下沉）、F7（debug 日志）、F8（dead code）。
5. 验证：`pnpm --filter @arcships/vue-xlsx typecheck` + `build` 后运行 `packages/vue-xlsx/test/structure.mjs`（注意：该测试目前从 `dist/index.js` 导入，构建的是 monolith，接入后需确认其仍覆盖 controller 方法表）；grep 确认 monolith 内不再保留已迁出函数的本地副本、`grep -rn 'from "\./composables/' packages/vue-xlsx/src/` 有命中。
