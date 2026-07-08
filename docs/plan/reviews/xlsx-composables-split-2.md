# Review: xlsx-composables-split-2

Date: 2026-07-08
Task: xlsx-composables-split（架构设计 `docs/xlsx-migration-architecture.md` §2.2 / §三 步骤4 / §4.4）
Branch: task/docx-remigration
HEAD: 2dd846f
前置 review: `docs/plan/reviews/xlsx-composables-split-1.md`（结论 blocked）

## 结论

**blocked** —— 拆分产出仍完全缺失，与 review-1 相同状态未改变。

架构设计 §2.2 要求把 `packages/vue-xlsx/src/composables.ts`（4724 行单文件）按职责拆分为 `composables/` 目录下的 9 个文件（8 模块 + 1 barrel），硬约束"单文件 ≤ 1000 行"。实际仓库 HEAD `2dd846f` 中：

- `packages/vue-xlsx/src/` 下仍只有 `composables.ts`（4724 行，未拆分）、`index.ts`（12 行）、`env.d.ts`，**`composables/` 目录为空**（仅存在空目录，无任何文件）。
- 两个标记为该任务的提交均未包含拆分代码：
  - `914001a`（"xlsx-composables-split: vue-xlsx composables 拆分"）只新增了 `docs/plan/reviews/xlsx-core-images-split-1.md`（一个无关 review 文件，疑似提交错位）。
  - `2dd846f`（"xlsx-composables-split: fix #1"）只新增了 `docs/plan/reviews/xlsx-composables-split-1.md`（review-1 文档本身）。
- 工作树干净（`git status` 无变更、无未跟踪文件），不存在被遗漏的临时实现。

因此本次 review 仍无法评估拆分质量（import 路径、模块边界、循环依赖、上游功能对齐），只能记录"未实现"事实。这与 review-1 结论一致。

## Findings

### F1 — 拆分产出完全缺失（功能缺失）

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §2.2（vue-xlsx/ 重写，composables/ 目录 9 文件清单）+ §三 步骤4（xlsx-composables-split 任务定义，scope=vue-xlsx，依赖 2/3）+ §4.4（composables 拆分策略：状态/选区/编辑/图表/剪贴板/导航/格式化）
- 代码位置：`packages/vue-xlsx/src/composables.ts`（仍为 4724 行单文件）；`packages/vue-xlsx/src/composables/`（空目录）

设计要求的 9 个文件均不存在：

| 设计文件（§2.2） | 预估行数 | 实际状态 |
|---|---:|---|
| `composables/useXlsxViewerController.ts` | ~700 | ❌ 缺失 |
| `composables/workbook-state.ts` | ~600 | ❌ 缺失 |
| `composables/selection.ts` | ~500 | ❌ 缺失 |
| `composables/editing.ts` | ~600 | ❌ 缺失 |
| `composables/chart-controller.ts` | ~500 | ❌ 缺失 |
| `composables/clipboard.ts` | ~400 | ❌ 缺失 |
| `composables/navigation.ts` | ~400 | ❌ 缺失 |
| `composables/formatting.ts` | ~500 | ❌ 缺失 |
| `composables/index.ts` | barrel | ❌ 缺失 |

证据：
```
$ git ls-tree -r --name-only HEAD packages/vue-xlsx/src/
packages/vue-xlsx/src/composables.ts
packages/vue-xlsx/src/env.d.ts
packages/vue-xlsx/src/index.ts

$ git show --stat 914001a
 docs/plan/reviews/xlsx-core-images-split-1.md | 137 +++++
 1 file changed, 137 insertions(+)

$ git show --stat 2dd846f
 docs/plan/reviews/xlsx-composables-split-1.md | 95 +++++
 1 file changed, 95 insertions(+)

$ wc -l packages/vue-xlsx/src/composables.ts
4724 packages/vue-xlsx/src/composables.ts

$ ls packages/vue-xlsx/src/composables/   # 空目录，无文件
```

### F2 — 单文件超 1000 行硬约束未满足

- 严重程度：**P1**
- 阻塞：**blocking**
- 设计文档位置：`docs/xlsx-migration-architecture.md` §硬约束（"单文件 ≤ 1000 行"）
- 代码位置：`packages/vue-xlsx/src/composables.ts:1-4724`

`composables.ts` 4724 行，是硬约束 ≤1000 行的 4.7 倍。本任务的唯一目标就是消除此超标，产出未触及该文件。

### F3 — 功能对齐 / import 路径 / 循环依赖检查项无法执行

- 严重程度：P1
- 阻塞：**blocking**
- 设计文档位置：`docs/upstream-xlsx-feature-alignment.md` §2.9（controller.tsx 对齐清单 A–L）+ §三（功能对齐要点 22–38）
- 代码位置：N/A（无拆分后的模块代码可检）

由于不存在拆分后的模块文件，以下检查项无法执行：
- 上游 commit f285a1c `controller.tsx` 的双层历史模型、selection、cell edit、clipboard/fill、sort、merge/resize、style、recalculate、sheet 管理、image/chart selection、zoom 等逻辑在拆分后是否完整保留（现单文件 `composables.ts` 本身已是对齐版本，但拆分过程可能引入回归，无代码则无法验证）。
- 拆分模块间的相对 import 路径是否正确。
- `composables/index.ts` barrel 是否正确 re-export，是否与 `packages/vue-xlsx/src/index.ts`（当前仅 12 行，直接 `from "./composables"`）衔接。

## 已通过的旁证（非本任务产出）

为完整性记录：当前未拆分的 `composables.ts` 自身 typecheck 通过。

```
$ pnpm --filter @extend-ai/vue-xlsx typecheck
> tsc --noEmit
（无输出，退出 0）
```

这仅说明 `task/xlsx-003`（commit 2459282 "vue-xlsx controller 机械改写"）的机械改写产物可编译，不构成 `xlsx-composables-split` 的拆分产出。本任务拆分代码不存在，typecheck 检查对"本任务产出"而言是"无目标可检"而非"通过"。

另注：`composables.ts` 本身无 stub/mock/fake 残留（全文件 grep 上述关键词无命中）；唯一的 stub 字样出现在 `index.ts:10` 的 `XlsxViewer` 占位组件，那是 `xlsx-components` 任务的范畴，非本任务范围，不影响本结论。

## 与 review-1 的差异

无差异。review-1（HEAD `914001a`）记录"拆分产出完全缺失"；本 review（HEAD `2dd846f`）确认状态未改变——两个提交之间只新增了两份 review 文档，未引入任何拆分代码。`composables/` 目录虽已创建（mtime 显示 7/8 19:32）但为空。

## 建议下一步

重做 `xlsx-composables-split`：在 `task/xlsx-003` 机械改写的 `composables.ts` 基础上，按 §2.2 的 8 模块边界做物理拆分（状态/选区/编辑/图表/剪贴板/导航/格式化/controller），每文件 ≤1000 行，提供 `composables/index.ts` barrel，并确保 `pnpm --filter @extend-ai/vue-xlsx typecheck` 通过后再交付。依赖任务 `xlsx-core-charts-split`、`xlsx-core-images-split` 已完成（本分支可见 `packages/xlsx-core/src/charts/`、`images/`、`types/` 拆分目录），composables 拆分时需更新 import 至新 barrel 路径（`@extend-ai/xlsx-core` 仍为统一入口，当前 `composables.ts` 已从该入口导入，拆分时保持即可）。
