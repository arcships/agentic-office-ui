---
id: docx-003
scope: vue-docx
status: pending
depends-on: [docx-002]
---

> ⚠️ **已失效（2026-07-07）**：本任务按旧方案（机械复制上游巨型文件）设计，已回退。
> 实际执行以 `docs/plan/README.md` 的「DOCX 模块化重做任务」为准。
> 架构设计见 `docs/docx-migration-architecture.md`。


# vue-docx composables 改写（editor.tsx hooks 区 → composables.ts）

## objective

将上游 editor.tsx 24954-32870 行（useDocxEditor + composables，~7917行）机械改写为 `packages/vue-docx/src/composables.ts`，hook → Vue reactivity。

## context

- `docs/upstream-docx-feature-alignment.md` 第二节 2.6（editor.tsx 后半部分）、第三节编辑能力要点(24-34)、第四节 Phase 3
- 上游：editor.tsx 24954-32870行

## path

- `packages/vue-docx/src/composables.ts`

## steps

1. 提取 editor.tsx 24954-32870 行
2. 替换 import：React → Vue reactivity
3. 脚本预处理（同 xlsx-003 的模式）：
   - useState × ~20 → ref
   - useCallback × ~228 → 普通函数
   - useRef × ~79 → ref 或局部 let
   - useEffect × ~60 → watch/onMounted
   - useMemo × ~84 → computed
4. 手动调整：
   - state 镜像 ref 可删除
   - dispatchEditorTransaction 核心事务逻辑完全保留，setModel(x) → model.value = x
   - pendingRunStyle 机制保留
   - historyRestoreRequest + nonce 保留
   - selectionSession 机制保留
   - contentEditable draft → ref(Map)，commitParagraphDraftFromElement 逻辑原样
5. 保持核心逻辑不变：
   - importDocxFile（wasm worker）/ exportDocx（serializeDocx + basePackage）
   - 双层历史（snapshot + cell-edit/range-edit）
   - 所有编辑操作（formatting/table/image/form-field/list/copy-paste）
   - 11 个 composables（theme/styles/lineSpacing/borders/formFields/trackChanges/comments/pageLayout/pagination/thumbnails/imageWrapMenu）

## verification

```bash
pnpm --filter @extend-ai/vue-docx typecheck
pnpm --filter @extend-ai/vue-docx build
```

结构测试：验证 controller 关键方法存在且可调用。
