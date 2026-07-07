# cross-001 Review #1

**Review date**: 2026-07-07
**Verdict**: ✅ **pass**

---

## 1. Findings

### P3 — non-blocking

| # | Finding | Design Ref | Code Ref |
|---|---------|------------|----------|
| 1 | `xlsx-core/package.json` 缺少 `./xlsx-worker` export path，与 `docx-core` 的 `./docx-import-worker` 不一致。xlsx worker 仅通过 `new Worker(new URL(...))` 加载，不需要 import 路径，但显式声明 export 有利于类型解析一致性 | — | [`packages/xlsx-core/package.json:8`](../../packages/xlsx-core/package.json#L8) vs [`packages/docx-core/package.json:13`](../../packages/docx-core/package.json#L13) |
| 2 | `packages/docx-core/src/docx_wasm_bg.wasm`（1MB）作为二进制文件提交到源码仓库，会持续膨胀 git 历史。后续可考虑改为构建时从上游同步的脚本 | — | [`packages/docx-core/src/docx_wasm_bg.wasm`](../../packages/docx-core/src/docx_wasm_bg.wasm) |

---

## 2. 逐项核对

| Step | 要求 | 状态 |
|------|------|------|
| xlsx-core tsup.config: worker entry `src/xlsx-worker.ts` | ✅ | [`packages/xlsx-core/tsup.config.ts:8`](../../packages/xlsx-core/tsup.config.ts#L8) |
| xlsx-core tsup.config: onSuccess 复制 `duke_sheets_wasm_bg.wasm` 到 dist | ✅ | [`packages/xlsx-core/tsup.config.ts:14-22`](../../packages/xlsx-core/tsup.config.ts#L14) |
| xlsx-core tsup.config: `external: ["@dukelib/sheets-wasm"]` | ✅ | [`packages/xlsx-core/tsup.config.ts:12`](../../packages/xlsx-core/tsup.config.ts#L12) |
| docx-core tsup.config: worker entry `src/docx-import-worker.ts` | ✅ | [`packages/docx-core/tsup.config.ts:6`](../../packages/docx-core/tsup.config.ts#L6) |
| docx-core tsup.config: onSuccess 复制 `docx_wasm_bg.wasm` 到 dist | ✅ | [`packages/docx-core/tsup.config.ts:12-16`](../../packages/docx-core/tsup.config.ts#L12) |
| docx-core tsup.config: `external: ["@chenglou/pretext", "fast-png", "utif"]` | ✅ | [`packages/docx-core/tsup.config.ts:10`](../../packages/docx-core/tsup.config.ts#L10) |
| demo vite.config: `worker: { format: "es" }` | ✅ | [`apps/demo/vite.config.ts:7-9`](../../apps/demo/vite.config.ts#L7) |
| demo vite.config: `optimizeDeps: { exclude: ["us-atlas", "world-atlas"] }` | ✅ | [`apps/demo/vite.config.ts:10-12`](../../apps/demo/vite.config.ts#L10) |
| demo public: `duke_sheets_wasm_bg.wasm`（3.3MB ≈ 3.4MB） | ✅ | [`apps/demo/public/duke_sheets_wasm_bg.wasm`](../../apps/demo/public/duke_sheets_wasm_bg.wasm) |
| demo public: `docx_wasm_bg.wasm`（1MB） | ✅ | [`apps/demo/public/docx_wasm_bg.wasm`](../../apps/demo/public/docx_wasm_bg.wasm) |
| demo main.ts: `setWasmSource("/duke_sheets_wasm_bg.wasm")` | ✅ | [`apps/demo/src/main.ts:14`](../../apps/demo/src/main.ts#L14) |
| demo main.ts: `setWasmSource("/docx_wasm_bg.wasm")` | ✅ | [`apps/demo/src/main.ts:15`](../../apps/demo/src/main.ts#L15) |
| Gate: `pnpm typecheck && pnpm build` 全 workspace 通过 | ✅ | 6/7 packages typecheck 通过，全部 build 通过，demo 产物含 worker chunk + wasm asset |

---

## 3. 结论

**pass**。所有 5 个 path 文件均按设计文档实现，worker entry、wasm 复制、Vite 配置、public 部署、`setWasmSource` 调用全部到位。`pnpm typecheck && pnpm build` 全 workspace 通过，demo 构建产物正确包含 worker chunk 和 hashed wasm asset。2 个 P3 finding 不影响交付。
