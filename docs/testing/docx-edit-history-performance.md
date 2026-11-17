# DOCX 编辑/历史性能记录（issue #37）

> 状态：已完成。对应 issue #37「DOCX editor deep-clones the whole document (including image bytes) and JSON-serialises the model for history on every edit」。

本页记录 DOCX 编辑路径的克隆与历史预算优化、基准数据、验证方式，以及本次明确未纳入范围的后续项。

## 1. 问题回顾

修改前每次编辑都要做两件 O(document) 的事情：

1. 所有编辑命令走 `cloneDocModel()`（`packages/docx-core/src/engine/clone.ts`），逐节点、逐页眉页脚、逐编号定义复制，并对图片 `data` 执行 `new Uint8Array(child.data)`。20 MB 图片的文档每次击键就复制 20 MB。
2. 核心历史状态机 `trimHistoryEntriesToBudget()` 在每次 push 时重新 JSON 序列化**每一条**保留条目，代价是 O(保留条目 × 文档)。32 MB 预算、1 MB 模型、32 条保留条目时，一次击键约 32 MB JSON。

## 2. 修改内容

| 位置 | 变更 |
|---|---|
| `engine/clone.ts` | 新增 `shareImageData()`：图片字节按引用共享（数组形态仍归一化为 `Uint8Array`），编辑路径的节点克隆默认使用它 |
| `engine/clone.ts` | 新增 `cloneDocModelShallow()`（写时复制外壳：只复制模型、`nodes` 数组和 `metadata` 对象）、`cloneDocNode()`（只深拷贝被编辑的顶层节点）、`cloneDocModelNodes()`（多节点命令：只克隆列出的顶层节点）、`cloneDocModelSectionLists()`（页眉/页脚：复制 metadata 中所有 section 列表外壳与 section 对象，区内节点共享） |
| `editor/paragraph-ops.ts` | `updateParagraphText`、`insertParagraph`、`removeParagraph`、`duplicateParagraph` 改为写时复制；`cloneImageRun` 共享图片字节 |
| `editor/run-style-ops.ts` | `applyRunStyle`、`toggleRunStyleFlag`、`setParagraphAlignment`、`setParagraphHeading`、`pasteParagraphs` 改为写时复制；`replaceText` 仍需全量克隆（它会重写全文字节点） |
| `editor/table-ops.ts` | 三个表格单元格文本命令只克隆被编辑的表格节点 |
| `editor/helpers/text-mutation.ts` | 文本区间/拆分辅助函数共享图片字节 |
| `layout/layout-engine.ts` | 布局结果按引用共享图片字节，不再每次布局复制 20 MB |
| `editor/helpers/history-budget.ts` | 新增 `cachedHistoryEntryBytes()`：以条目对象为弱键缓存字节估算；`trimHistoryEntriesToBudget()` 默认使用缓存版本。条目创建一次后只读，估算稳定；预算口径与数值完全不变 |
| `editor/helpers/section-manipulation.ts` | `updateSectionParagraphTextAtLocation`、`updateSectionImageFloatingAtLocation` 改用 `cloneDocModelSectionLists`，页眉页脚输入不再整篇克隆 |
| `vue-docx/composables/editor-text-input.ts` | `commitTableCellTextByRuntimeKey`、`commitTableCellParagraphTextByRuntimeKey`、`appendParagraph`、`replaceExpandedSelection`、`splitParagraphAtSelection`、`insertListItemAfterSelection` 改为写时复制；新增 `editor-clone-helpers.ts` 提供 `topLevelNodeIndex` / `cloneHostTableForRuntimeKey` |
| `vue-docx/composables/editor-format.ts` | `applySelectedStyleChange`、`setLink`、`setHeading`、`setParagraphStyle`、`setLineSpacing`、`setAlignment`、`applyBorderPreset`（表格、区间、单段三条分支）只克隆被格式化的顶层节点 |
| `vue-docx/composables/editor-image.ts` | `applyImageChange`（缩放/文本框/浮层/环绕）、`insertImageFile`、`moveImage`、`moveParagraphDropCap`、`setParagraphDropCapFontSizePt`、`setParagraphDropCapText` 只克隆图片所在节点（及移动目标节点）；`moveSectionFloatingImage` 走已迁移的 section 命令 |

公开入口新增 `cloneDocModelShallow`、`cloneDocNode`、`cloneDocModelNodes`、`cloneDocModelSectionLists`（`@arcships/docx-core` 根入口与 `./core`），供宿主复用同一套写时复制约定；`scripts/ci/public-api-contract.json` 的 docx-core 声明指纹同步更新为 `1203 / 8f8d72b0…`（并入 #30/#31/#32/#34/#35/#38 后的合并值）。

### 不变量

- 模型图仍然不可变：所有编辑命令返回新模型，历史快照引用旧模型。未修改的节点、页眉页脚、编号定义和 metadata 数组在多个模型之间共享。
- 被编辑的顶层节点必须是 `cloneDocNode()` 的结果，不能直接在共享节点上赋值；`updateParagraphText` 等命令已按此实现。
- 编辑路径（`cloneDocNode` / `cloneParagraph` / `cloneImageRun` 默认参数）图片 `data` 在所有副本中指向同一缓冲区，读取方（解码、布局、导出）只能读，不得原地写入。
- `cloneDocModel()` 是隔离边界：保留给导入、标准化和长生命周期检查点，仍然完整深拷贝文档图，并保留自己的一份图片字节（`{ shareImageBytes: false }`），因此与 `docx-model-preservation` 的“克隆之间无共享可变状态”契约一致。历史快照只存模型引用，不走这条路径。

## 3. 基准

脚本：[`scripts/bench/docx-edit-history.mjs`](../../scripts/bench/docx-edit-history.mjs)。它用 Node 类型擦除直接加载 `packages/docx-core/src` 下的源文件，不依赖构建产物或 `node_modules`，因此可以在任意 checkout（包括修改前的提交）里跑同一份脚本。

```bash
node scripts/bench/docx-edit-history.mjs        # 表格
node scripts/bench/docx-edit-history.mjs --json # 明细
```

场景：200 段（约 200 页）+ 1 张表格 + 1 个页眉、 4 张 5 MB 图片（共 20 MB）的文档；100 次击键、100 次格式切换、100 次表格单元格提交、100 次页眉文字提交、60 次历史 push（20 段文本模型、保留上限 100 条）、20 次布局。

测量环境：Linux x64、Node.js v26.8.1、单进程、无浏览器。绝对毫秒只用于同机前后对比。

| 场景 | 修改前 | 修改后 | 倍数 |
|---|---|---|---|
| 击键 `updateParagraphText`（每次） | 3.53 ms；重建 200 个节点；复制 20 971 520 B 图片 | 0.067 ms；重建 1 个节点；复制 0 B | ≈53× |
| 格式切换 `applyRunStyle`（每次） | 2.21 ms；重建 200 个节点；复制 20 971 520 B | 0.024 ms；重建 1 个节点；复制 0 B | ≈92× |
| 表格单元格 `updateTableCellText`（每次） | 2.2 ms（整篇克隆） | 0.043 ms；重建 1 个节点；复制 0 B | ≈51× |
| 页眉文字 `updateSectionParagraphTextAtLocation`（每次） | 2.2 ms（整篇克隆） | 0.037 ms；重建 0 个顶层节点；复制 0 B | ≈59× |
| 历史 push（保留 60 条） | 3.45 ms；32.5 次 JSON 序列化/push | 0.19 ms；2.97 次/push | ≈18× |
| `layoutDocument`（每次） | 2.08 ms；共享 0 B 图片 | 0.31 ms；共享 20 971 520 B | ≈6.7× |

说明：

- 历史 push 的 2.97 次序列化 = 新条目估算 1 次 + 文本 run 样式比较 2 次（`distributeTextAcrossRuns` 的小对象比较，与文档大小无关）。修改前该数字随保留条目线性增长（60 条时平均 32.5 次/push）。
- 两种实现的 `history.pastBytes` 完全一致（4 503 592 B），说明预算口径没有被放宽或收紧。
- 每次击键仍会对**当前模型**做一次 JSON 估算；这是每条目一次的 O(document) 成本，不再是 O(保留条目 × 文档)。

## 4. 验证

| 测试 | 覆盖 |
|---|---|
| `tests/unit/docx-edit-structural-sharing.test.mjs` | 图片字节共享；写时复制外壳共享未修改节点与 metadata；`cloneDocModelNodes` 只复制列出的节点；`cloneDocModelSectionLists` 复制列表外壳并共享未改节点；单段编辑/表格编辑只重建一个节点；200 页文档 100 次击键只重建 100 个节点且不复制图片字节；历史预算每个快照只序列化一次，且 `pastBytes` 与逐条全量序列化一致 |
| `tests/component/docx-edit-clone-sharing.test.mjs` | 通过 `useDocxEditor` 逐条验证交互命令（输入、格式、表格、追加段、图片缩放、插入表行、页眉文字）只改有限顶层节点、图片缓冲区保持同一引用；200 页文档 100 次击键重建 100 个节点、0 次图片复制 |
| `tests/component/docx-history-budget.test.mjs` | 通过 `useDocxEditor` 在 200 页文档（2×2 MB 图片）里输入 100 个字符：未修改段落与图片 `data` 的引用保持稳定、撤销仍可回到上一状态；原有字节预算与溢出条目语义不变 |
| `tests/unit/core-purity.test.mjs`（既有） | 深度冻结模型上的 `updateParagraphText` / `updateTableCellText` / `applyRunStyle` 仍返回新对象且不修改输入 |
| `tests/unit/docx-history-cache.test.mjs`（既有） | 历史按条数/字节淘汰、保留最新超限条目、实例隔离 |

本地运行（需先构建 `@arcships/docx-core` 的 `dist`，组件测试还需要 `@arcships/vue-docx` 的 `dist`）：

```bash
pnpm --filter=@arcships/docx-core build
pnpm --filter=@arcships/vue-docx build
node --test --test-concurrency=1 tests/unit/docx-edit-structural-sharing.test.mjs
node --test --test-concurrency=1 tests/unit/docx-history-cache.test.mjs tests/unit/docx-core.test.mjs
node --test --test-concurrency=1 tests/component/docx-history-budget.test.mjs tests/component/docx-edit-clone-sharing.test.mjs
```

两个新测试已分别登记到 `pnpm test:unit` 与 `pnpm test:component`（`scripts/ci/run-suite.mjs`）。

## 5. 仍使用完整克隆的位置（非交互编辑路径）

交互编辑路径已无 `cloneDocModel()`：`packages/vue-docx` 内只剩编辑器构建（`useDocxEditor.ts` 的 starter/clone）与 `newDocument()` 重置（`editor-import-export.ts`）。以下保留完整克隆且不使用 `cloneDocModelShallow`：

- `engine/clone.ts` 自身的 `cloneDocModel`——导入、标准化、检查点与上述构建/重置路径的隔离边界。
- `editor/run-style-ops.ts` `replaceText`——全文档查找替换，本就要重写所有段落，不存在“未修改节点”。
- `editor/section-break-ops.ts` `insertSectionBreak` / `deleteSectionBreak`——结构性命令（重排 `metadata.sections`、可能新增段落），未接入 `vue-docx` 编辑器 UI，目前只作为公开核心 API 使用。
- `editor/helpers/tracked-change-decisions.ts`——批量接受/拒绝修订，会跨任意多节点与表格行/单元格重写内容，属于批量文档级操作。

历史字节估算仍是每个条目一次整模型 JSON；若要把这次 O(document) 也消掉，可改为按顶层节点缓存字节估算（`WeakMap<DocNode, number>`）后求和，但会改变 `pastBytes` 的数值口径，需要同步评估预算语义。
