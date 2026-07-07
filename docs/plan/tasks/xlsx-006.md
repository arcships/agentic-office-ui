---
id: xlsx-006
scope: demo
status: pending
depends-on: [xlsx-005]
---

# demo XlsxViewerPage 接入

## objective

更新 `apps/demo/src/pages/XlsxViewerPage.vue` 使用新 controller + viewer，加 setWasmSource，对齐上游 playground 功能。

## context

- `docs/upstream-xlsx-feature-alignment.md` 第八节（playground 对齐）
- 上游 playground：`/Users/eric8810/Code/extend-ui-upstream/react-xlsx/apps/playground/src/App.tsx`

## path

- `apps/demo/src/pages/XlsxViewerPage.vue`
- `apps/demo/src/main.ts`（加 setWasmSource）
- `apps/demo/public/`（放 wasm 二进制）

## steps

1. main.ts 增加 `setWasmSource("/duke_sheets_wasm_bg.wasm")`
2. 复制 `duke_sheets_wasm_bg.wasm` 到 `apps/demo/public/`
3. XlsxViewerPage.vue 更新：
   - import 新的 XlsxViewer + useXlsxViewerController
   - 对齐 playground 功能：ribbon 6 tab、formula bar、sheet tabs、缩略图
   - fixture 切换（保留本地多 fixture，比上游更丰富）
   - 文件上传/拖拽/URL 加载
   - worker 开关、只读开关、canvas 开关

## verification

浏览器验证 `http://localhost:5000/#/xlsx-viewer`：
- fixture 切换正常
- 编辑/公式/合并/排序/导出可用
- 4 视口（1440×900 / 1280×720 / 768×1024 / 390×844）布局正常
