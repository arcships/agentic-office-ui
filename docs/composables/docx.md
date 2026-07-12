# DOCX 组合函数

## `useDocxEditor`

```ts
const editor = useDocxEditor({
  initialFileName: "document.docx",
  historyMaxEntries: 100,
  historyMaxBytes: 32 * 1024 * 1024,
})
```

它返回 `DocxEditorController`。主要状态包括：

- `model`、`fileName`、`status`、`isImporting` 和 `importError`；
- `selection`、`activeTextRange` 和当前段落、文字样式、链接；
- `canUndo`、`canRedo`、`currentPage` 和 `totalPages`；
- 批注、修订、主题、列表、表单域和边框状态。

常用命令：

| 分类 | 方法 |
|---|---|
| 文件 | `importDocxFile()`、`newDocument()`、`exportDocx()` |
| 历史 | `undo()`、`redo()` |
| 文字 | `toggleBold()`、`toggleItalic()`、`toggleUnderline()`、`setFontFamily()`、`setFontSize()`、`setTextColor()` |
| 段落 | `setHeading()`、`setParagraphStyle()`、`setAlignment()`、`setLineSpacing()`、`toggleList()` |
| 表格 | `insertTable()`、增删行列、调整列宽、删除表格 |
| 图片 | `insertImageFile()`、`resizeImage()`、`moveImage()`、`setImageWrapMode()` |
| 输入 | 提交段落或单元格文字、复制和粘贴 |

控制器可以传给 `<DocxEditor :editor="editor" />`。控制器与编辑器组件应在同一业务区域内使用，避免多个编辑界面同时修改同一控制器。

## `useDocxModel`

```ts
const file = ref<ArrayBuffer>()
const state = useDocxModel(() => file.value)
```

返回值使用普通读取属性：`state.model`、`state.isLoading` 和 `state.error`。来源变化会取消旧加载；作用域销毁时内部 Runtime 和加载器自动销毁。

它只负责解析，不提供编辑历史、选区或工具栏。需要编辑时使用 `useDocxEditor`。

## `useDocxPageThumbnails`

```ts
const thumbnails = useDocxPageThumbnails(editor, {
  maxWidthPx: 180,
  maxCacheEntries: 24,
  maxCacheBytes: 32 * 1024 * 1024,
})
```

该组合函数依赖已经挂载并发布页面表面的 `DocxEditor`。返回：

- `pageThumbnailStates`；
- `attachCanvasForPage()` 和 `detachCanvasForPage()`；
- `renderPageThumbnailToCanvas()`；
- `prefetchPageThumbnailSurface()`；
- `rerenderAttachedThumbnails()`。

如果页面尚未挂载，缩略图状态可能是 `unavailable`。旧的 `useDocxViewerThumbnails` 没有文档输入，只为兼容保留，新代码不要使用。

## 辅助组合函数

主题、段落样式、图片环绕、行距、边框、表单域、修订、批注、页面布局和分页组合函数都围绕 `DocxEditorController` 工作。它们用于拆分自定义工具栏，不创建第二份文档模型。

```ts
const editor = useDocxEditor()
const comments = useDocxComments(editor)
const pagination = useDocxPagination(editor)
const lineSpacing = useDocxLineSpacing(editor)
```

具体返回字段以包导出的 `UseDocx*Result` 类型为准。不要直接使用内部页面注册表或已经弃用的低层渲染组件来建立新接口。
