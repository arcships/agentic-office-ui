---
id: xlsx-001
scope: xlsx-core
status: done
depends-on: []
---

# xlsx-core 引擎层复制 + types 清理

## objective

从上游 `@extend-ai/react-xlsx` 复制 7 个零 React 依赖文件 + types.ts 到 `packages/xlsx-core/src/`，清理 types.ts 的 React 类型标注，调整 import 路径，安装依赖。完成后 xlsx-core 是可独立 typecheck + build 的引擎包。

## context

- `docs/upstream-xlsx-feature-alignment.md` 第一节（文件迁移分类总表）、第二节 2.1-2.8、第四节 Phase 1-2
- 上游源码：`/Users/eric8810/Code/extend-ui-upstream/react-xlsx/packages/react-xlsx/src/`

## path

- `packages/xlsx-core/src/wasm.ts`
- `packages/xlsx-core/src/safe-calculate.ts`
- `packages/xlsx-core/src/worker-client.ts`
- `packages/xlsx-core/src/xlsx-worker.ts`
- `packages/xlsx-core/src/colors.ts`
- `packages/xlsx-core/src/images.ts`
- `packages/xlsx-core/src/charts.ts`
- `packages/xlsx-core/src/types.ts`
- `packages/xlsx-core/src/wasm-url.d.ts`
- `packages/xlsx-core/src/index.ts`
- `packages/xlsx-core/package.json`
- `docs/INDEX.md`（更新导出状态）

## steps

1. 复制 8 个文件（7 .ts + wasm-url.d.ts）到 `packages/xlsx-core/src/`
2. 复制 types.ts，sed 清理 React 类型标注：
   - 删除 `import type * as React from "react"`
   - `React.CSSProperties` → `Record<string, string | number | undefined>`
   - `React.ReactNode` → `unknown`
   - `React.PointerEvent<HTMLElement>` → `Record<string, unknown>`
   - `React.Ref<HTMLDivElement>` → `unknown`
   - `React.ButtonHTMLAttributes<...>` → `Record<string, unknown>`
   - `React.HTMLAttributes<...>` → `Record<string, unknown>`
3. 删除纯 React 组件 prop 接口（XlsxViewerProps 等），保留数据模型类型
4. 重写 `index.ts` 导出所有公开 API
5. package.json 增加依赖：`@dukelib/sheets-wasm`、`fflate`（已有）、d3 全家桶、`regl`、`topojson-client`、`us-atlas`、`world-atlas`
6. 复制 wasm 二进制或配置 copy 脚本

## verification

```bash
pnpm --filter @arcships/xlsx-core typecheck
pnpm --filter @arcships/xlsx-core build
grep -r "React\." packages/xlsx-core/src/  # 应为空
```

Node smoke：
```bash
node --input-type=module - <<'NODE'
import init, { Workbook } from '@dukelib/sheets-wasm'
// fromBytes → calculate → setFormula → saveXlsxBytes
// openpyxl 验证导出
NODE
```
