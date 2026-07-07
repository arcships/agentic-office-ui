---
id: cross-001
scope: build-config
status: pending
depends-on: [xlsx-001, docx-001]
---

# 构建配置 + wasm 部署

## objective

配置 tsup/vite 构建，部署 wasm 二进制，使全 workspace 可构建运行。

## context

- `docs/upstream-xlsx-feature-alignment.md` 第九节（构建配置对齐）
- `docs/upstream-docx-feature-alignment.md` 第七节（构建配置对齐）

## path

- `packages/xlsx-core/tsup.config.ts`（新建或更新）
- `packages/docx-core/tsup.config.ts`（新建或更新）
- `apps/demo/vite.config.ts`（更新）
- `apps/demo/src/main.ts`（更新）
- `apps/demo/public/`（wasm 文件）

## steps

1. xlsx-core tsup.config：
   - 增加 worker entry（`src/xlsx-worker.ts`）
   - onSuccess 复制 `duke_sheets_wasm_bg.wasm` 到 dist
   - external: `@dukelib/sheets-wasm`
2. docx-core tsup.config：
   - 增加 worker entry（`src/docx-import-worker.ts`）
   - onSuccess 复制 `docx_wasm_bg.wasm` 到 dist
   - external: `@chenglou/pretext`, `fast-png`, `utif`
3. demo vite.config：
   - `worker: { format: "es" }`
   - `optimizeDeps: { exclude: ["us-atlas", "world-atlas"] }`
4. demo public：
   - 复制 `duke_sheets_wasm_bg.wasm`（3.4MB）到 `apps/demo/public/`
   - 复制 `docx_wasm_bg.wasm`（1MB）到 `apps/demo/public/`
5. demo main.ts：
   - `setWasmSource("/duke_sheets_wasm_bg.wasm")` (xlsx)
   - `setWasmSource("/docx_wasm_bg.wasm")` (docx)

## verification

```bash
pnpm typecheck && pnpm build
```

全 workspace 构建通过，无 wasm 加载错误。
