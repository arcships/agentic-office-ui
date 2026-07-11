# Review: xlsx-composables-split-7

Date: 2026-07-08
Task: xlsx-composables-split（架构设计 `docs/xlsx-migration-architecture.md` §2.2 / §三 步骤4 / §4.4）
Branch: task/docx-remigration
HEAD: 667787a
前置 review: `xlsx-composables-split-1.md`～`-6.md`（全部 blocked）

## 结论

**blocked** —— HEAD `667787a`（"fix #6"）相比 review-6（`1c86bdf`）**无任何代码改动**：该 commit 唯一变更是新增本任务 review-6 的 markdown 文件（`docs/plan/reviews/xlsx-composables-split-6.md`，+287 行）。`git log --oneline 1c86bdf..HEAD` 仅此一个 commit，`git diff 914001a..HEAD -- packages/vue-xlsx/src/composables.ts` 仍为空。

review-6 记录的全部 blocking 问题一字未改：

1. **monolith `composables.ts` 仍为 4724 行**，自 `914001a` 起未改一行（硬约束的 4.7 倍）。
2. **架构要求的 9 文件仍缺 2 个核心文件**：`useXlsxViewerController.ts`（核心 controller）与 `index.ts`（barrel）。
3. **子模块仍是死代码**：monolith 不 import 任何子模块，全仓库无 `from "./composables/"` 引用，`src/index.ts:2,4,5` 仍解析到 monolith `composables.ts`；子模块是 monolith 既有逻辑的重复副本 + 未被使用的平行 domain factory 架构。

这是连续第七次 review 同一核心问题。review-6 给出的 5 条建议（剪切迁移、补 controller 文件、补 barrel + 接线、清理 monolith debug 日志、验证）全部未落实，连第一步"剪切迁移"都未启动。review-6→review-7 之间零代码进展。

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
4. `src/index.ts:2,4,5` 仍 `from "./composables"`，解析到 monolith `composables.ts`，子模块无人引用。
5. monolith 的 `useXlsxViewerController` 从不构造 `XlsxControllerContext`、从不调用任何 domain factory，其 return（monolith:4606-4723）直接引用内联局部函数：
   ```
   $ grep -nE "createChartImageDomain|createEditingDomain|createHistoryDomain|createClipboardDomain|createNavigationDomain|XlsxControllerContext" packages/vue-xlsx/src/composables.ts
   （none - monolith uses inline logic）
   ```

即子模块整套设计从未被接入。正确做法是"从 monolith 移出 → monolith 删除本地实现并改为 import"，最终 `composables.ts` 应消失或仅剩极薄入口。当前是"复制 + 加平行架构"，必须改为"剪切"。

结构测试 `packages/vue-xlsx/test/structure.mjs` 从 `dist/index.js` 导入，构建的是 monolith（`src/index.ts:2` → `./composables` → monolith），测试通过仅证明 monolith 的 controller 方法表完整，不能证明拆分已生效。

### F4 — typecheck 通过（但仅验证 monolith + 孤立子模块可编译，不证明拆分生效）

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
formatting → internal
internal → selection
image-assets → internal
clipboard → internal, selection, formatting
chart-controller → internal, image-assets（+ re-export image-assets 符号）
workbook-state → selection, internal, image-assets, clipboard, formatting, ⇢internal
editing → selection, workbook-state, ⇢internal
navigation → workbook-state, internal, formatting, chart-controller
```

拓扑序：selection / formatting → internal → image-assets → clipboard → chart-controller → workbook-state → editing → navigation。无环。被多方依赖的共享基础（常量、历史类型、`applyCellMutationState`、`XlsxControllerContext`、`resolveInheritedCellStyle`）已下沉到 `internal.ts`/`formatting.ts`，符合 §4.4 隐含的模块单向依赖。

注意：`navigation` 是图的叶子（无人 import 它），因 monolith 走内联逻辑、子模块无人接入。接入后 `useXlsxViewerController.ts` 应作为 `navigation` 等各 domain 的聚合消费者，成为新的（唯一）根。

子模块内部的相对 import 路径均正确（typecheck 通过佐证），无绝对路径/包内错误引用。问题在于"接线"缺失：`src/index.ts` 与 monolith 均不引用子模块，故子模块路径正确与否对运行时无意义（见 F3）。

### F6 — 上游功能对齐：monolith 对齐良好；拆分产出的对齐因死代码无法验证（对齐验证不可进行）

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

`composables.ts` 与 `composables/*.ts` 无 stub/mock/fake 残留（grep 无命中）。唯一 stub 字样在 `src/index.ts:10` 的 `XlsxViewer` 占位（`xlsx-viewer-stub`），属 `xlsx-components` 任务范畴，非本任务范围。

## 证据汇总

```
$ git rev-parse HEAD
667787a92cf34ed62ce3e5185d75d532ff9c1019

$ git log --oneline 1c86bdf..HEAD          # review-6 → review-7 之间仅 1 个 commit
667787a xlsx-composables-split: fix #6
$ git show --stat 667787a                    # 该 commit 仅新增 review-6 markdown，无代码
 docs/plan/reviews/xlsx-composables-split-6.md | 287 ++++++++++++++++++++++++++
 1 file changed, 287 insertions(+)

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
```

## 与历次 review 的差异

| review | HEAD | composables/ 文件 | monolith 行数 | 接入 | typecheck | 结论 |
|---|---|---|---:|---|---|---|
| -1 | 914001a | 0 | 4724 | 无 | (monolith) pass | blocked |
| -2 | 2dd846f | 0 | 4724 | 无 | (monolith) pass | blocked |
| -3 | 03585c3 | 3 | 4724 | 无 | (monolith) pass | blocked |
| -4 | ff0fa44 | 9 | 4724 | 无 | fail (2 err) | blocked |
| -5 | b3417fa | 9 | 4724 | 无 | pass | blocked |
| -6 | 1c86bdf | 9 | 4724 | 无 | pass | blocked |
| **-7** | **667787a** | **9** | **4724** | **无** | **pass** | **blocked** |

连续七次 review 同一核心问题（monolith 未拆、子文件是复制副本/平行死代码）未解决。review-6→review-7 之间**零代码进展**——唯一的 commit 仅添加了 review-6 的 markdown 文件，`composables.ts` 与 `composables/*.ts` 均无任何改动。

## 建议下一步

review-6 的 5 条建议一字未落实，仍然适用：

1. **剪切迁移**（review-4/5/6 建议 1，未落实）：按 §2.2 的 8 模块把 `composables.ts` 代码**移出**，monolith 删除本地实现并改为从子模块 `import`，最终 `composables.ts` 消失或仅剩极薄入口。当前是"复制 + 加平行架构"，必须改为"剪切"。
2. **补 controller 文件**（review-4/5/6 建议 2，未落实）：`composables/useXlsxViewerController.ts`，把 `useXlsxViewerController`（monolith:1899）及其 setup/return（monolith:4606-4723）移入，return 改为聚合各 `createXxxDomain(ctx)` 的产物。
3. **补 barrel + 接线**（review-4/5/6 建议 3，未落实）：`composables/index.ts`，`src/index.ts` 从 `./composables`（解析到 barrel）导入。接入后 F3 死代码问题自然消除。
4. 清理 monolith 侧的 debug 日志（`composables.ts:3657,3672,3692`）。
5. 验证：`pnpm --filter @arcships/vue-xlsx typecheck` + `build` 后运行 `packages/vue-xlsx/test/structure.mjs`（接入后需确认其从 `dist/index.js` 导入的 controller 方法表仍覆盖 requiredMethods）；grep 确认 monolith 内不再保留已迁出函数的本地副本、`grep -rn 'from "\./composables/' packages/vue-xlsx/src/` 有命中、`composables.ts` 行数 ≤1000 或消失。

关键判断：本任务已连续七轮 blocking 且 review-6→review-7 零代码进展。继续重复 review 不会推进拆分；建议将任务退回执行方，明确要求"剪切迁移而非复制"这一唯一正确路径后再提交，避免第八次 review 同一问题。
