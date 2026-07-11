# Review: xlsx-composables-split-3

Date: 2026-07-08
Task: xlsx-composables-split（架构设计 `docs/xlsx-migration-architecture.md` §2.2 / §三 步骤4 / §4.4）
Branch: task/docx-remigration
HEAD: 03585c3
前置 review: `docs/plan/reviews/xlsx-composables-split-1.md`（blocked）、`xlsx-composables-split-2.md`（blocked）

## 结论

**blocked** —— 本任务的拆分目标仍未达成。HEAD `03585c3` 在 `composables/` 下新增了 3 个文件，但它们是孤立副本（未被任何代码引用），架构设计 §2.2 要求的 9 文件拆分（8 模块 + 1 barrel）未完成，4724 行的 `composables.ts` 原封不动，单文件 ≤1000 行的硬约束仍未满足。

本任务的核心目标只有一条：把 `composables.ts`（4724 行）按职责物理拆分到 ≤1000 行的多个模块。这是 review-1 / review-2 已经两次标记 blocking 的同一问题。本次产出相比 review-2 多了 3 个文件，但没有一个被接入，monolith 没有缩减一行。

## Findings

### F1 — 拆分未完成：monolith 仍是 4724 行，硬约束 ≤1000 行未满足（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §硬约束（"单文件 ≤ 1000 行"）+ §2.2（composables/ 目录 9 文件清单）+ §4.4（composables 拆分策略）
- 代码位置：`packages/vue-xlsx/src/composables.ts:1-4724`

`composables.ts` 行数未变，仍为 4724 行（硬约束的 4.7 倍）。HEAD `03585c3` 的提交 diff 只新增文件，未修改该 monolith。这是本任务存在的唯一理由，未完成即 blocking。

```
$ wc -l packages/vue-xlsx/src/composables.ts
4724 packages/vue-xlsx/src/composables.ts
```

### F2 — 架构要求的 9 个文件只产出 3 个，且无 barrel、无 controller/状态/编辑/图表/剪贴板/导航模块（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（composables/ 目录清单）
- 代码位置：`packages/vue-xlsx/src/composables/`

§2.2 要求的 9 文件对账：

| 设计文件（§2.2） | 预估行数 | 实际状态 |
|---|---:|---|
| `composables/useXlsxViewerController.ts` | ~700 | ❌ 缺失（核心 controller 仍在 monolith 内） |
| `composables/workbook-state.ts` | ~600 | ❌ 缺失 |
| `composables/selection.ts` | ~500 | ⚠️ 文件存在（170 行），但为孤立副本，未接入 |
| `composables/editing.ts` | ~600 | ❌ 缺失 |
| `composables/chart-controller.ts` | ~500 | ❌ 缺失 |
| `composables/clipboard.ts` | ~400 | ❌ 缺失 |
| `composables/navigation.ts` | ~400 | ❌ 缺失 |
| `composables/formatting.ts` | ~500 | ⚠️ 文件存在（285 行），但为孤立副本，未接入 |
| `composables/index.ts` | barrel | ❌ 缺失（无 barrel，无法被 `src/index.ts` 引用） |

```
$ ls packages/vue-xlsx/src/composables/
formatting.ts  internal.ts  selection.ts   # 无 index.ts barrel
```

另：`composables/internal.ts`（208 行）不在 §2.2 清单内，是实现自加的"常量/类型/工具"聚合文件（见 F4）。

### F3 — 已产出的 3 个文件是孤立副本，未被引用（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2 + §4.4（拆分后需接入）
- 代码位置：`packages/vue-xlsx/src/composables/formatting.ts`、`selection.ts`、`internal.ts`；`packages/vue-xlsx/src/composables.ts`、`packages/vue-xlsx/src/index.ts`

`composables.ts` 与 `src/index.ts` 均不从 `./composables/` 导入任何东西，`formatting.ts`/`selection.ts`/`internal.ts` 全仓库无引用方：

```
$ grep -rn "composables/" packages/vue-xlsx/src/   # 无输出
```

更严重的是，这 3 个文件的内容是 monolith 中既有逻辑的**重复副本**，不是"剪切后迁移"：

- `internal.ts` 的常量（`FORMULA_COUNT_THRESHOLD`、`DEFAULT_ROW_HEIGHT`、`EMU_PER_PIXEL`、`HISTORY_LIMIT`、`DEFAULT_ZOOM_SCALE` 等）与 `composables.ts:66-88` 逐一重复。
- `selection.ts` 的 `columnLabel`/`cellAddressToA1`/`parseA1CellReference`/`parseA1RangeReference`/`normalizeRange`/`rangeToA1`/`rangeContainsCell`/`parseWorksheetFreezePanes`/`parseWorksheetDataValidations`/`resolveSheetDisplayUsedRange` 与 `composables.ts:502-918` 重复。
- `formatting.ts` 的 `formatBinaryBytes`/`findZipEndOfCentralDirectoryOffset`/`readZipCentralDirectoryEntries`/`preflightWorkbookBuffer`/`createWorkbookTooLargeError`/`XlsxFileSizeLimitExceededError`/`resolveDisplayFileName`/`fileStem`/`pxToSheetRowHeight`/`cssColor`/`mapBorder`/`decodeHtmlEntities`/`escapeHtml`/`cloneBytes`/`sanitizeSavedWorkbookBytes` 与 `composables.ts:318-1148` 重复。

即 monolith 没有删掉对应实现，运行时仍走 monolith 内的本地副本，新增的 3 文件是死代码。正确的拆分应是"从 monolith 移出 → monolith 改为 import"，而非"复制一份留着"。

### F4 — `composables/internal.ts` 是设计未提及的自加模块（non-blocking）

- 严重程度：P3
- 阻塞：**non-blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（清单无 `internal.ts`）
- 代码位置：`packages/vue-xlsx/src/composables/internal.ts`

§2.2 的 8 模块按职责（状态/选区/编辑/图表/剪贴板/导航/格式化/controller）划分，未列 `internal.ts`。把常量、历史类型、idle/abort 工具聚合到一个 `internal.ts` 是合理的工程决策（这些确属跨模块共享的内部基础设施），设计文档未覆盖属细节补充。但鉴于 F2/F3，该文件当前同样孤立，无法评估其接入后的模块边界是否清晰。若后续重做拆分时保留此文件，建议在设计文档补一行说明其职责。

### F5 — typecheck 通过，但只覆盖 monolith，对"本任务拆分产出"无证明力

- 严重程度：P2
- 阻塞：**non-blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §三 步骤4（typecheck 要求）
- 代码位置：`packages/vue-xlsx/src/composables.ts`（monolith，可编译）；`packages/vue-xlsx/src/composables/*.ts`（孤立，无 import 关系，不构成可验证的模块图）

```
$ pnpm --filter @arcships/vue-xlsx typecheck
> tsc --noEmit
（无输出，退出 0）
```

typecheck 通过仅证明 monolith（task/xlsx-003 机械改写产物）可编译。3 个孤立文件因无人 import，其类型正确性、import 路径正确性、与 monolith 的接口契约是否一致，均未被编译器实际检验（孤立文件会被 tsc 扫描但不会触发跨模块约束检查）。真正的拆分（monolith 改为从子模块 import）完成后，typecheck 才有意义。

### F6 — 上游功能对齐：monolith 自身对齐良好，但拆分产出的对齐无法验证

- 严重程度：P2
- 阻塞：**non-blocking**（对齐主体未变；blocking 来自 F1-F3 的拆分缺失）
- 设计文档位置：`docs/upstream-xlsx-feature-alignment.md` §2.9（controller.tsx 对齐清单 A–L）+ §三（功能对齐要点 22–38）
- 代码位置：`packages/vue-xlsx/src/composables.ts`

上游 `@extend-ai/react-xlsx` commit `f285a1c` 的 `controller.tsx` 核心逻辑在 monolith 中均存在：双层历史（`pushHistoryEntry`/`applyCellMutationState`/`isApplyingHistoryRef`）、`coerceUserEnteredValue`、`maybeRecalculateWorkbook`+`tryRecalculate`、`sanitizeSavedWorkbookBytes`、`loadWorkbookChartAssets`、`buildSheetList`、`resolveInheritedCellStyle`、`clampZoomScale`/`zoomScaleOverridesByTabId`、`sortTable` 等。这部分是 task/xlsx-003 的成果，对齐良好。

但本任务（拆分）的对齐验证无法进行：拆分后的模块不存在，无法确认 17 项关键逻辑（A–L）在拆分后是否完整保留、模块边界是否割裂了原本内聚的逻辑（如历史模型的 snapshot/cell-edit/range-edit 三类若被分到不同文件需谨慎处理共享状态）。`composables.ts` 与 `composables/*.ts` 无 stub/mock/fake 残留（grep 无命中），唯一 stub 字样在 `index.ts:10` 的 `XlsxViewer` 占位，属 `xlsx-components` 任务范畴。

## 证据汇总

```
$ wc -l packages/vue-xlsx/src/composables.ts packages/vue-xlsx/src/composables/*.ts
    4724 packages/vue-xlsx/src/composables.ts
     285 packages/vue-xlsx/src/composables/formatting.ts
     208 packages/vue-xlsx/src/composables/internal.ts
     170 packages/vue-xlsx/src/composables/selection.ts
    5404 total

$ grep -rn "composables/" packages/vue-xlsx/src/      # 孤立文件无引用，无输出

$ ls packages/vue-xlsx/src/composables/index.ts        # 无 barrel
ls: ...: No such file or directory

$ git show --stat 03585c3 | tail -6
 packages/vue-xlsx/src/composables/formatting.ts | 285 +++++
 packages/vue-xlsx/src/composables/internal.ts   | 208 +++
 packages/vue-xlsx/src/composables/selection.ts  | 170 +++
（composables.ts 未出现在 diff）

$ pnpm --filter @arcships/vue-xlsx typecheck          # monolith 可编译
（无输出，退出 0）
```

## 与 review-1 / review-2 的差异

review-1（HEAD `914001a`）、review-2（HEAD `2dd846f`）：拆分产出完全缺失，`composables/` 为空目录。

review-3（HEAD `03585c3`）：`composables/` 新增 3 文件（`formatting.ts`/`selection.ts`/`internal.ts`），但均为 monolith 的孤立重复副本，未接入、无 barrel、monolith 未缩减。问题性质从"完全未实现"变为"实现无效"——blocking 状态未改变。连续三次 review 同一核心问题未解决。

## 建议下一步

重做 `xlsx-composables-split`，采用"剪切迁移"而非"复制"：

1. 按职责把 `composables.ts` 的代码**移出**到 8 模块（state/selection/editing/chart/clipboard/navigation/formatting/controller），monolith 改为从子模块 `import`，删除 monolith 内的重复实现，最终 `composables.ts` 应消失或仅剩极薄入口。
2. 补 `composables/index.ts` barrel，`src/index.ts` 从 `./composables` 导入（当前已如此，接入后即生效）。
3. 每文件 ≤1000 行（controller 若超 1000 需进一步按方法域拆，或参照 §4.4 接受 controller 聚合但其余模块独立）。
4. 拆分后 `pnpm --filter @arcships/vue-xlsx typecheck` 必须通过，并运行 `packages/vue-xlsx/test/structure.mjs` 验证 controller 方法表完整。
5. 依赖任务 `xlsx-core-charts-split`/`xlsx-core-images-split` 已完成，composables 从 `@arcships/xlsx-core` 统一入口导入即可（当前 monolith 已如此，拆分时保持）。
