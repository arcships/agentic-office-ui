---
id: docx-001
scope: docx-core
status: pending
depends-on: []
---

> ⚠️ **已失效（2026-07-07）**：本任务按旧方案（机械复制上游巨型文件）设计，已回退。
> 实际执行以 `docs/plan/README.md` 的「DOCX 模块化重做任务」为准。
> 架构设计见 `docs/docx-migration-architecture.md`。


# docx-core 引擎+布局+辅助复制（~10500 行）

## objective

从上游 `@extend-ai/react-docx` 的 8 个子包复制 28 个零 React 依赖文件到 `packages/docx-core/src/`，调整 import 路径，安装依赖。

## context

- `docs/upstream-docx-feature-alignment.md` 第一节 1.1（文件迁移总表）、第二节 2.1-2.4、第四节 Phase 1
- 上游：`/Users/eric8810/Code/extend-ui-upstream/react-docx/packages/`

## path

- `packages/docx-core/src/` 下 28 个文件
- `packages/docx-core/package.json`（增加依赖）

## steps

1. 复制引擎层 4 包：
   - wasm → `wasm.ts` + `generated/`
   - ooxml-core → `ooxml-core.ts`
   - serializer → `serializer.ts`
   - doc-model → `types.ts` + `clone.ts` + `normalize.ts` + `doc-model.ts`
2. 复制编辑操作：
   - editor-ops → `editor-ops.ts`
3. 复制布局层 2 包：
   - layout-core → `pagination.ts` + `page-segmentation.ts` + `layout-core.ts`
   - layout-engine → `layout-engine.ts`
4. 复制 react-viewer 辅助模块 14 个文件：
   - pretext-layout.ts, thumbnail-raster.ts, layout-snapshot.ts, docx-import.ts, docx-import-worker.ts, section-layout.ts, pagination-breaks.ts, image-render.ts, page-count-reconciliation.ts, content-signature.ts, canvas/layout-diagnostics.ts, canvas/types.ts, wasm-source.ts, utif.d.ts
5. 复制 core/state.ts（特殊处理：import 了 editor.tsx 的 DocxEditorSelection/DocxTextRange 类型——提取这些类型到独立 `editor-types.ts`）
6. 调整所有 `@extend-ai/react-docx-*` import 为相对路径
7. 重写 `index.ts` 导出
8. package.json 增加 `@chenglou/pretext`、`fast-png`、`utif`、`fflate`
9. 复制 `docx_wasm_bg.wasm` 到合适位置

## verification

```bash
pnpm --filter @arcships/docx-core typecheck
pnpm --filter @arcships/docx-core build
```

Node smoke：`wasmBuildDocModelFromBytes` + `wasmSerializeDocx` 可用。
