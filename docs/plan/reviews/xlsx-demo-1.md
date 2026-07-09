# xlsx-demo Review #1

Date: 2026-07-09
Scope: `apps/demo/src/pages/XlsxViewerPage.vue`, `apps/demo/src/main.ts`
Task: [xlsx-006](docs/plan/tasks/xlsx-006.md)

## Typecheck

```bash
pnpm --filter @extend-ai/vue-xlsx typecheck
```

✅ **零错误**。`tsc --noEmit` 通过，无类型错误。

## 构建

```bash
pnpm --filter demo build
```

✅ **构建成功**。586 modules transformed，dist 产出包含 xlsx-worker、wasm、demo bundle。仅 chunk size warning（>500KB），属于预期行为。

## Findings

### P1 — blocking

| # | 描述 | 位置 |
|---|---|---|
| 1 | **缺少 Ribbon 工具栏**：XlsxToolbar 仅包含缩放/下载/导出按钮，无 6 个功能 tab（Home、Insert、Page Layout、Formulas、Data、View），与上游 playground 不对齐 | [XlsxToolbar.vue](packages/vue-xlsx/src/components/XlsxToolbar.vue) |
| 2 | **缺少 Formula bar**：无公式编辑栏，用户无法查看或编辑单元格公式 | [XlsxViewer.vue](packages/vue-xlsx/src/components/XlsxViewer.vue) |
| 3 | **缺少缩略图（thumbnails）**：sheet tab 区域无缩略图预览功能 | [XlsxSheetTabs.vue](packages/vue-xlsx/src/components/XlsxSheetTabs.vue) |
| 4 | **缺少拖拽上传**：XlsxViewerPage 仅支持 `<input type="file">` 选择文件，无 drag-and-drop 交互 | [XlsxViewerPage.vue:16](apps/demo/src/pages/XlsxViewerPage.vue#L16) |
| 5 | **缺少 URL 输入加载**：仅可通过下拉选择本地 sample fixture，无用户输入 URL 加载远程文件的能力 | [XlsxViewerPage.vue:9-14](apps/demo/src/pages/XlsxViewerPage.vue#L9) |
| 6 | **缺少 Canvas 开关**：task 要求 canvas 渲染开关，当前仅 worker/readOnly toggle | [XlsxViewerPage.vue:18-22](apps/demo/src/pages/XlsxViewerPage.vue#L18) |

以上 6 项均为 [xlsx-006](docs/plan/tasks/xlsx-006.md) task spec 明确列出的功能点（步骤 3: "ribbon 6 tab、formula bar、sheet tabs、缩略图" + "文件上传/拖拽/URL 加载" + "canvas 开关"），当前均未实现。

### P2 — non-blocking

无。

### P3 — non-blocking

| # | 描述 | 位置 |
|---|---|---|
| 1 | Task 状态仍为 `status: pending`，未更新为 `in_progress` 或 `done` | [xlsx-006.md:3](docs/plan/tasks/xlsx-006.md#L3) |
| 2 | API Verification 表格中 `XlsxViewer` 和 `useXlsxViewerController()` 标记为 `⚠️ SCOPED` / `⚠️ PARTIAL`，应为 `✅`（当前通过 typecheck + build + structure test 验证） | [XlsxViewerPage.vue:55-57](apps/demo/src/pages/XlsxViewerPage.vue#L55) |

## 已通过验证

### setWasmSource 配置

```ts
// main.ts:11,14
import { setWasmSource as setXlsxWasmSource } from "@extend-ai/xlsx-core"
setXlsxWasmSource("/duke_sheets_wasm_bg.wasm")
```

✅ wasm 二进制已放置于 `apps/demo/public/duke_sheets_wasm_bg.wasm`（3.4MB），构建后自动复制到 dist。

### Controller + Viewer 接入

✅ `useXlsxViewerController` + `XlsxViewer` 正确 import 自 `@extend-ai/vue-xlsx`。
✅ Wrapper 组件通过 `defineComponent` 创建 controller 并传入 XlsxViewer，key 机制保证 source 切换时正确重载。
✅ Fixture 切换：5 个 sample（financial-model、sales-table、charts-images、large-grid、corrupted）通过 `<select>` + URL 加载。
✅ Worker toggle 正确传入 `useWorker` prop。
✅ Read-only toggle 正确传入 `readOnly` prop。

### Import 路径

全部 import 路径正确：
- `@extend-ai/xlsx-core`：wasm API、类型定义
- `@extend-ai/vue-xlsx`：controller + viewer 组件
- 相对路径 `./components/Xxx.vue`：组件间引用
- 无直接引用 dist、无跨 workspace 裸路径

### stub/mock/fake 检查

✅ 全量 grep `stub|mock|fake|noop|TODO|FIXME|HACK` 在 `packages/vue-xlsx/src/` 下零匹配。所有 `return null` / `return false` 均为合法 guard clause（无 workbook、无 chart data、无 DOM container 等边界处理）。

## 结论

**blocked** — 6 个 task spec 明确要求的 UI 功能点（ribbon 6 tab、formula bar、thumbnails、drag-drop、URL 输入加载、canvas toggle）均未实现。typecheck/build 通过、import 路径正确、无 stub/mock 残留，基础设施就绪。blocking 项补齐后可重新 review。
