# Review: xlsx-composables-split-6

Date: 2026-07-08
Task: xlsx-composables-split（架构设计 `docs/xlsx-migration-architecture.md` §2.2 / §三 步骤4 / §4.4）
Branch: task/docx-remigration
HEAD: 1c86bdf
前置 review: `xlsx-composables-split-1.md`（blocked）、`-2.md`（blocked）、`-3.md`（blocked）、`-4.md`（blocked）、`-5.md`（blocked）

## 结论

**blocked** —— 本任务的核心目标（把 4724 行的 `composables.ts` 按职责物理拆分到 ≤1000 行的多个模块、并把 monolith 改为从子模块 import）依然未达成。HEAD `1c86bdf`（"fix #5"）相比 review-5（`b3417fa`）落实了 review-5 的三条 non-blocking 子模块内部质量问题（F6 上下文下沉、F7 debug 日志、F8 dead code），但本任务存在的唯一理由——拆掉 monolith——依然零进展：

1. **monolith `composables.ts` 仍为 4724 行，自 `914001a` 起从未修改**（硬约束的 4.7 倍，`git diff 914001a..HEAD -- composables.ts` 为空）。
2. **架构要求的 9 文件仍缺 2 个核心文件**：`useXlsxViewerController.ts`（核心 controller）与 `index.ts`（barrel）。`useXlsxViewerController` 仍在 monolith:1899。
3. **monolith 不从任何子模块 import**，子模块是 monolith 既有逻辑的重复副本 + 全新平行架构（domain factory + `XlsxControllerContext`），运行时仍走 monolith 内联逻辑，子文件是死代码。

这是连续第六次 review 同一核心问题：monolith 没有被拆分，只是被复制并叠加了一套未被使用的平行抽象。review-5 的 5 条建议只有第 4 条（清理 F6/F7/F8 子模块内部质量）被落实，第 1/2/3/5 条（剪切迁移、补 controller 文件、补 barrel、接线 + 验证）——即拆分本体——完全未动。

本次相比 review-5 的实质改进（均为子模块内部质量，不改变 blocking 状态）：
- F6 已修复：`XlsxControllerContext` 从 `navigation.ts` 下沉到 `internal.ts:251`；`navigation.ts` 不再 re-export 无关符号，仅导出 `createNavigationDomain`；barrel 职责集中问题消除。
- F7 已修复：`chart-controller.ts` 的 3 处 `[react-xlsx debug]` 日志已删除。
- F8 已修复：`editing.ts` 的 `trimmedFormula` dead code 已改为实际使用 `trimmedFormula`。
- `navigation.ts` 从 433 → 310 行，`internal.ts` 从 226 → 351 行（吸收下沉的上下文与常量），`chart-controller.ts` 从 621 → 604 行。

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

4724 行是硬约束的 4.7 倍。本任务存在的唯一理由就是拆掉这个 monolith；monolith 自 task 起未改一行，即 blocking。

### F2 — 架构要求的 9 文件仍缺 2 个核心文件：controller 文件与 barrel（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（composables/ 目录清单）
- 代码位置：`packages/vue-xlsx/src/composables/`（无 `index.ts`、无 `useXlsxViewerController.ts`）

§2.2 要求的 9 文件对账：

| 设计文件（§2.2） | 预估行数 | 实际状态 |
|---|---:|---|
| `composables/useXlsxViewerController.ts` | ~700 | ❌ **缺失**（核心 controller `useXlsxViewerController` 仍在 monolith:1899，return 在 monolith:4606-4723） |
| `composables/workbook-state.ts` | ~600 | ⚠️ 存在（987 行），孤立未接入 |
| `composables/selection.ts` | ~500 | ⚠️ 存在（170 行），孤立未接入 |
| `composables/editing.ts` | ~600 | ⚠️ 存在（450 行），孤立未接入 |
| `composables/chart-controller.ts` | ~500 | ⚠️ 存在（604 行），孤立未接入 |
| `composables/clipboard.ts` | ~400 | ⚠️ 存在（482 行），孤立未接入 |
| `composables/navigation.ts` | ~400 | ⚠️ 存在（310 行），孤立未接入 |
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
5. monolith 的 `useXlsxViewerController` 从不构造 `XlsxControllerContext`、从不调用任何 domain factory，其 return（monolith:4606-4723）直接引用内联局部函数：
   ```
   $ grep -nE "createChartImageDomain|createEditingDomain|createHistoryDomain|createClipboardDomain|createNavigationDomain|XlsxControllerContext" packages/vue-xlsx/src/composables.ts
   （none - monolith uses inline logic）
   ```

即子模块整套设计从未被接入。正确做法是"从 monolith 移出 → monolith 删除本地实现并改为 import"，最终 `composables.ts` 应消失或仅剩极薄入口。当前是"复制 + 加平行架构"，必须改为"剪切"。

结构测试 `packages/vue-xlsx/test/structure.mjs` 从 `dist/index.js` 导入，构建的是 monolith（`src/index.ts:2` → `./composables` → monolith），测试通过仅证明 monolith 的 controller 方法表完整，不能证明拆分已生效：

```
$ pnpm --filter @arcships/vue-xlsx build && node packages/vue-xlsx/test/structure.mjs
... All structure checks passed.   # 验证的是 monolith，子模块是死代码
```

### F4 — typecheck 通过

- 严重程度：—
- 阻塞：**non-blocking**（本项无问题，记录验证结果）
- 设计文档位置：`docs/xlsx-migration-architecture.md` §三 步骤4（typecheck 要求）
- 代码位置：`packages/vue-xlsx/src/composables/*.ts`；`packages/vue-xlsx/src/composables.ts`

```
$ pnpm --filter @arcships/vue-xlsx typecheck
> tsc --noEmit
（exit 0，无输出）
```

typecheck 通过仅证明 monolith + 子模块都能编译，不能证明拆分已生效（子模块是死代码，见 F3）。

### F5 — 子模块 import 图现为 DAG，无循环依赖（模块边界）

- 严重程度：—
- 阻塞：**non-blocking**（本项无问题，记录验证结果）
- 设计文档位置：`docs/xlsx-migration-architecture.md` §4.4
- 代码位置：`packages/vue-xlsx/src/composables/*.ts`

当前 import 图（value import 用 `→`，type-only 用 `⇢`）：

```
selection → (none)
formatting → (none)
internal → selection
image-assets → internal
clipboard → internal, selection, formatting
chart-controller → internal, image-assets（+ re-export image-assets 符号）
workbook-state → selection, internal, image-assets, clipboard, ⇢internal
editing → selection, workbook-state, ⇢internal
navigation → workbook-state, internal, formatting, chart-controller
```

拓扑序：selection / formatting → internal → image-assets → clipboard → chart-controller → workbook-state → editing → navigation。无环。被多方依赖的共享基础（常量、历史类型、`applyCellMutationState`、`XlsxControllerContext`、`resolveInheritedCellStyle`）已下沉到 `internal.ts`/`formatting.ts`，符合 §4.4 隐含的模块单向依赖。

注意：`navigation` 是图的叶子（无人 import 它），因 monolith 走内联逻辑、子模块无人接入。接入后 `useXlsxViewerController.ts` 应作为 `navigation` 等各 domain 的聚合消费者，成为新的（唯一）根。

### F6 — review-5 F6 已修复：navigation.ts 职责归位，XlsxControllerContext 下沉

- 严重程度：—
- 阻塞：**non-blocking**（本项已修复，记录验证结果）
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（`navigation.ts` 职责为"导航/滚动"）
- 代码位置：`packages/vue-xlsx/src/composables/navigation.ts`；`packages/vue-xlsx/src/composables/internal.ts:251`

review-5 F6 的两个问题均已解决：
- `XlsxControllerContext` 从 `navigation.ts` 移到 `internal.ts:251`，各子模块统一从 `./internal` import type。
- `navigation.ts` 不再 re-export 无关符号（review-5 记录的 `navigation.ts:416-433` re-export 块已删除），现仅导出 `createNavigationDomain`（`navigation.ts:59`）。

```
$ grep -n "^export" packages/vue-xlsx/src/composables/navigation.ts
59:export function createNavigationDomain(ctx: XlsxControllerContext) {
$ grep -n "export interface XlsxControllerContext" packages/vue-xlsx/src/composables/internal.ts
251:export interface XlsxControllerContext {
```

### F7 — review-5 F7 已修复：chart-controller.ts debug 日志已删除

- 严重程度：—
- 阻塞：**non-blocking**（子模块侧已修复；monolith 侧同位置仍有，属 monolith 既有问题，非本任务范围）
- 设计文档位置：—（上游 `controller.tsx` 残留，非设计文档要求）
- 代码位置：`packages/vue-xlsx/src/composables/chart-controller.ts`

review-5 F7 记录的 `chart-controller.ts:314,329,349` 三处 `console.info("[react-xlsx debug] ...")` 已删除：

```
$ grep -n "console\." packages/vue-xlsx/src/composables/chart-controller.ts
（无输出）
```

注意 monolith 同位置（`composables.ts:3657,3672,3692`）仍有这 3 处 debug 日志，属 monolith 既有问题（上游机械改写遗留），非拆分新引入。接入时 monolith 侧应一并清理，但属 monolith 既有问题，非本任务新增范围。

### F8 — review-5 F8 已修复：editing.ts trimmedFormula dead code 已修正

- 严重程度：—
- 阻塞：**non-blocking**（本项已修复，记录验证结果）
- 设计文档位置：—
- 代码位置：`packages/vue-xlsx/src/composables/editing.ts:130-134`

review-5 F8 记录的 `trimmedFormula` 声明后未使用问题已修正，现实际使用 `trimmedFormula`：

```
$ grep -n "trimmedFormula" packages/vue-xlsx/src/composables/editing.ts
130:    const trimmedFormula = formula.trim();
131:    if (!trimmedFormula) {
134:      worksheet.setFormula(cellAddressToA1(cell), trimmedFormula);
```

### F9 — 上游功能对齐：monolith 对齐良好；拆分产出的对齐因死代码无法验证（对齐验证不可进行）

- 严重程度：**P2**
- 阻塞：**non-blocking**（对齐主体 monolith 未变，blocking 来自 F1-F3）
- 设计文档位置：`docs/upstream-xlsx-feature-alignment.md` §2.9（controller.tsx 对齐清单 A–L）+ §三（功能对齐要点 22–38）
- 代码位置：`packages/vue-xlsx/src/composables.ts`（monolith）；`packages/vue-xlsx/src/composables/*.ts`（死代码）

上游 `@extend-ai/react-xlsx` commit `f285a1c` 的 `controller.tsx` 核心逻辑在 monolith 中均存在（monolith 自 review-4 起未变，本次复核确认）：

- 要点 23 双层历史：`pushHistoryEntry`/`applyCellMutationState`/`isApplyingHistoryRef`（monolith:1937,3272,3285），snapshot/cell-edit/range-edit 三类齐全（43 处历史相关命中）。
- 要点 25 `HISTORY_LIMIT=100` FIFO（monolith:75,1079）。
- 要点 26 `isApplyingHistoryRef` 防自记（monolith:1937）。
- 要点 37 image/chart rect 走 snapshot + px→EMU（`rectToImageAnchor` monolith:17，`EMU_PER_PIXEL=9525` monolith:82）。
- 要点 38 zoom 纯前端按 tab 存（`zoomScaleOverridesByTabId` monolith:1927，`clampZoomScale` monolith:153）。
- 其余要点 24/27-36（applyCellMutationState 优先 formula、coerceUserEnteredValue、paste 区分 formula/value、fillSelection 平铺、sortTable 计算值排序、merge/unmerge/resize snapshot、setRangeStyle 单次 wasm、maybeRecalculate trap 降级）review-4/5 已逐条确认，monolith 未变。

子模块内容是 monolith 的逐字副本 + domain factory 包装，逻辑上与上游对齐等价于 monolith 对齐；但因 F3 子模块是死代码，拆分后的模块边界是否割裂内聚逻辑（如历史模型的 snapshot/cell-edit/range-edit 三类跨 `workbook-state`/`editing` 分布）无法在运行时验证。

`composables.ts` 与 `composables/*.ts` 无 stub/mock/fake 残留（grep 无命中）。唯一 stub 字样在 `src/index.ts:10` 的 `XlsxViewer` 占位，属 `xlsx-components` 任务范畴，非本任务范围。

## 关于 import 路径

子模块内部的相对 import 路径均正确（typecheck 通过佐证），无绝对路径/包内错误引用。问题在于"接线"缺失：`src/index.ts` 与 monolith 均不引用子模块，故子模块路径正确与否对运行时无意义（见 F3）。

## 证据汇总

```
$ git rev-parse HEAD
1c86bdfdfabd8cf222f9a72a116cb9017b2531fd

$ git diff --stat 914001a..HEAD -- packages/vue-xlsx/src/composables.ts   # monolith 自 task-003 起未改
（空）

$ wc -l packages/vue-xlsx/src/composables.ts packages/vue-xlsx/src/composables/*.ts
4724 composables.ts
 604 chart-controller.ts
 482 clipboard.ts
 450 editing.ts
 305 formatting.ts
 704 image-assets.ts
 351 internal.ts
 310 navigation.ts
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

$ pnpm --filter @arcships/vue-xlsx build && node packages/vue-xlsx/test/structure.mjs
... All structure checks passed.   # 验证的是 monolith
```

## 与历次 review 的差异

| review | HEAD | composables/ 文件 | monolith 行数 | 接入 | typecheck | 结论 |
|---|---|---|---:|---|---|---|
| -1 | 914001a | 0 | 4724 | 无 | (monolith) pass | blocked |
| -2 | 2dd846f | 0 | 4724 | 无 | (monolith) pass | blocked |
| -3 | 03585c3 | 3 | 4724 | 无 | (monolith) pass | blocked |
| -4 | ff0fa44 | 9 | 4724 | 无 | fail (2 err) | blocked |
| -5 | b3417fa | 9 | 4724 | 无 | pass | blocked |
| **-6** | **1c86bdf** | **9** | **4724** | **无** | **pass** | **blocked** |

连续六次 review 同一核心问题（monolith 未拆、子文件是复制副本/平行死代码）未解决。本次仅落实了 review-5 的 non-blocking 子模块内部质量清理（F6 上下文下沉、F7 debug 日志、F8 dead code），拆分本体（剪切迁移 + controller 文件 + barrel + 接入）零进展。

## 建议下一步

review-5 的建议仍然适用，本次仅强调被忽略的部分：

1. **剪切迁移**（review-4/5 建议 1，未落实）：按 §2.2 的 8 模块把 `composables.ts` 代码**移出**，monolith 删除本地实现并改为从子模块 `import`，最终 `composables.ts` 消失或仅剩极薄入口。当前是"复制 + 加平行架构"，必须改为"剪切"。
2. **补 controller 文件**（review-4/5 建议 2，未落实）：`composables/useXlsxViewerController.ts`，把 `useXlsxViewerController`（monolith:1899）及其 setup/return（monolith:4606-4723）移入，return 改为聚合各 `createXxxDomain(ctx)` 的产物。
3. **补 barrel + 接线**（review-4/5 建议 3，未落实）：`composables/index.ts`，`src/index.ts` 从 `./composables`（解析到 barrel）导入。接入后 F3 死代码问题自然消除。
4. 清理 monolith 侧的 F7 debug 日志（`composables.ts:3657,3672,3692`）。
5. 验证：`pnpm --filter @arcships/vue-xlsx typecheck` + `build` 后运行 `packages/vue-xlsx/test/structure.mjs`（接入后需确认其从 `dist/index.js` 导入的 controller 方法表仍覆盖 requiredMethods）；grep 确认 monolith 内不再保留已迁出函数的本地副本、`grep -rn 'from "\./composables/' packages/vue-xlsx/src/` 有命中、`composables.ts` 行数 ≤1000 或消失。
