---
id: xlsx-003
scope: vue-xlsx
status: done
depends-on: [xlsx-001]
---

# vue-xlsx controller 机械改写（controller.tsx → composables.ts）

## objective

将上游 `controller.tsx`（4987行）机械改写为 `packages/vue-xlsx/src/composables.ts`，把 React hooks 转换为 Vue reactivity，逻辑原样保留。

## context

- `docs/upstream-xlsx-feature-alignment.md` 第二节 2.9（controller.tsx 对齐清单）、第三节编辑能力要点、第四节 Phase 3
- 上游：`/Users/eric8810/Code/extend-ui-upstream/react-xlsx/packages/react-xlsx/src/controller.tsx`

## path

设计文档（只读）：
- `docs/upstream-xlsx-feature-alignment.md` 第二节 2.9 + 第三节 + 第四节 Phase 3

- `packages/vue-xlsx/src/composables.ts`

## steps

1. 复制 controller.tsx 到 composables.ts
2. 替换 import：`import * as React from "react"` → `import { ref, computed, watch, onMounted, onUnmounted, shallowRef } from "vue"`
3. 脚本预处理：
   - `useState(x)` → `const x = ref(x)`（28处）
   - `useCallback(fn, [deps])` → `function fn() { ... }`（97处，删依赖数组）
   - `useRef(x)` → `const xRef = ref(x)` 或局部 `let`（12处）
   - `useEffect(fn, [deps])` → `watch(deps, fn)` / `onMounted(fn)`（6处）
   - `useMemo(fn, [deps])` → `const x = computed(() => fn())`（20处）
4. 手动调整：
   - `setState(x)` → `state.value = x`（所有 setX 调用）
   - state 镜像 ref（modelRef/selectionRef）可删除——Vue 的 ref 本身就是 .value
   - useEffect 依赖数组 → watch 监听源
   - useEffect 返回 cleanup → onUnmounted
   - `const workbook = null` → 真实 `Workbook` handle
5. 保持核心逻辑不变：
   - dispatchEditorTransaction / commitWorkbookMutation 事务逻辑
   - 双层历史（snapshot + cell-edit/range-edit）
   - 公式阈值 1000、safeCalculate、tryRecalculate
   - worker 四消息协议、transfer clone
   - preflight central directory、三阈值
   - saveXlsxBytes + sanitize + merge images
   - 所有编辑操作（setCell/setFormula/setCellStyle/merge/resize/paste/fill/sort）

## verification

```bash
pnpm --filter @extend-ai/vue-xlsx typecheck
pnpm --filter @extend-ai/vue-xlsx build
```

结构测试：验证 controller 关键方法存在且可调用（loadWorkbookFromBuffer/recalculate/exportXlsx/setCellFormula/mergeSelection/sortTable/undo/redo/resizeColumn/resizeRow/pasteText/fillSelection）。
