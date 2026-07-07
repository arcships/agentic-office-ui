---
id: docx-006
scope: demo
status: pending
depends-on: [docx-004, docx-005]
---

# demo DocxViewer/Editor 接入

## objective

更新 DocxViewerPage.vue + DocxEditorPage.vue 使用新组件，加 setWasmSource，对齐上游 playground 功能。

## context

- `docs/upstream-docx-feature-alignment.md` 第八节（playground 对齐）
- 上游 playground：`/Users/eric8810/Code/extend-ui-upstream/react-docx/apps/playground/src/App.tsx`

## path

- `apps/demo/src/pages/DocxViewerPage.vue`
- `apps/demo/src/pages/DocxEditorPage.vue`
- `apps/demo/src/main.ts`（加 setWasmSource for docx）
- `apps/demo/public/`（放 docx_wasm_bg.wasm）

## steps

1. main.ts 增加 docx wasm 加载
2. 复制 `docx_wasm_bg.wasm` 到 `apps/demo/public/`
3. DocxViewerPage.vue 更新：import 新 DocxViewer，fixture 切换
4. DocxEditorPage.vue 更新：
   - import 新 DocxEditorViewer + useDocxEditor
   - 对齐 playground 26 项功能：toolbar（undo/redo/字体/字号/行距/格式/颜色/对齐/列表/边框/图片/表格/缩放/导入/导出/修订/评论/只读）
   - fixture 切换（保留本地多 fixture）

## verification

浏览器验证：
- `http://localhost:5000/#/docx-viewer`：fixture 切换 + 分页 + 渲染
- `http://localhost:5000/#/docx-editor`：编辑 + 格式化 + 导出
- 4 视口布局正常
