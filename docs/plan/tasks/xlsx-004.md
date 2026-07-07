---
id: xlsx-004
scope: vue-xlsx
status: pending
depends-on: [xlsx-001]
---

# vue-xlsx chart-renderer + surface-regl 改写

## objective

将上游 `chart-renderer.tsx`（7174行）和 `surface-regl.tsx`（1185行）局部改写为 Vue 兼容版本。纯逻辑函数直接保留，JSX → Vue render function。

## context

- `docs/upstream-xlsx-feature-alignment.md` 第二节 2.10-2.11、第四节 Phase 4-5
- 上游：`chart-renderer.tsx`、`surface-regl.tsx`

## path

- `packages/vue-xlsx/src/chart-renderer.ts`
- `packages/vue-xlsx/src/surface-regl.ts`

## steps

1. 复制两个文件到 vue-xlsx/src/
2. chart-renderer.tsx 改写：
   - `import * as React from "react"` → `import { defineComponent, h, memo } from "vue"`
   - 90 处 `return (<svg>...</svg>)` → `return h('svg', { ...attrs }, [children])` 或保留 SVG 生成逻辑
   - `React.memo(Component)` → `memo(defineComponent({ ... }))`
   - 2 个 `useCallback` → 普通函数
   - d3 逻辑、颜色解析、chart 数据读取等纯函数（~6000行）零修改
3. surface-regl.ts 改写：
   - 1 个 `useEffect` → `onMounted`/`watch`
   - 1 个 `useRef` → `ref`
   - 1 个 `React.memo` → `defineComponent`
   - WebGL 逻辑（buildSurfaceMesh、regl command、shader）直接复制
   - 失败兜底（try/catch → SVG fallback）保留

## verification

```bash
pnpm --filter @extend-ai/vue-xlsx typecheck
```

验证图表 SVG 生成逻辑可用（至少 bar/line/pie 三种类型）。
