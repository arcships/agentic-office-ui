---
id: docx-integ
scope: docx-core, vue-docx
status: pending
depends-on: [docx-003]
---

# DOCX 集成验证：composables → docx-core 真实调用

## objective

验证 vue-docx/composables.ts 对 docx-core 引擎层的调用是真实路径（不是 stub/mock）。覆盖集成点 D1-D5。

## context

- `docs/plan/analysis/migration-split.md` 集成关系枚举 D1-D5
- `docs/upstream-docx-feature-alignment.md` 引擎层对齐要点 1-8、编辑能力要点 24-34

## path

- `packages/docx-core/src/`（被调方，只读验证）
- `packages/vue-docx/src/composables.ts`（调用方，只读验证）

## verification

### D1: wasm 引擎调用

```bash
node --input-type=module - <<'NODE'
// 验证 composables 实际调用 docx-core 的 wasm API
// 1. importDocxFile → wasmBuildDocModelFromBytes
// 2. exportDocx → wasmSerializeDocx(model, basePackage)
// 验证：导出的 docx 用 python-docx 打开，段落/样式保留
NODE
```

通过标准：fromBytes → model → serialize 完整路径执行，python-docx 验证导出 docx

### D2: docx-import worker 调用

```bash
node --input-type=module - <<'NODE'
// 验证 composables 的 worker 模式
// 1. importDocxBuffer 走 worker 路径
// 2. 返回 {package, model, timings}
NODE
```

通过标准：worker 加载成功，返回真实 model（nodes 非空）

### D3: 分页调用

```bash
node --input-type=module - <<'NODE'
// 验证 composables 调用 editor-helpers 分页
// 1. buildDocumentPageNodeSegments(model, ...) 返回 pages
// 2. pages 非空且 pageCount > 0
NODE
```

通过标准：分页结果非空，页数合理

### D4: 编辑操作调用

```bash
node --input-type=module - <<'NODE'
// 验证 composables 调用 editor-ops
// 1. commitParagraphText → updateParagraphText
// 2. toggleBold → applyRunStyle / mutateParagraphTextStyleInRange
// 3. 验证 model 变化（sourceXml = undefined, run style 变化）
NODE
```

通过标准：编辑后 model 发生真实变化（非 stub 返回）

### D5: 序列化 + basePackage

```bash
node --input-type=module - <<'NODE'
// 验证 composables 的导出路径
// 1. serializeDocx(model, basePackage) — basePackage 非空
// 2. 导出 docx 保留 styles/numbering/headers
NODE
```

通过标准：basePackage 机制生效，python-docx 验证样式/编号保留

### blocking 判定

以下情况为 blocking：
- composables 中任何 `{}` / `return` / noop 替代真实 wasm/编辑调用
- importDocxFile 不走 worker 或主线程 wasm 路径
- exportDocx 不走 serializeDocx + basePackage
- dispatchEditorTransaction 不调用 cloneDocModel + editor-ops
