# XLSX 组合函数

## `useXlsxViewerController`

```ts
const controller = useXlsxViewerController({
  file: bytes,
  fileName: "forecast.xlsx",
  readOnly: false,
  historyMaxEntries: 100,
  historyMaxBytes: 32 * 1024 * 1024,
  onDiagnostic(event) {
    console.log(event.type, event.error?.code)
  },
})
```

它负责文件加载、Worker、工作簿对象、工作表标签、选区、编辑历史和导出。主要状态包括：

- `isLoading`、`sourceState`、`sourceError` 和 `error`；
- `sheets`、`tabs`、`activeSheet` 和 `activeTab`；
- `activeCell`、`selection`、`selectedRangeAddress`；
- `charts`、`images`、`shapes` 和 `formControls`；
- `canUndo`、`canRedo`、`canExport`、`readOnly` 和缩放能力。

常用命令：

- 单元格：`setCellValue()`、`setCellFormula()`、`setCellStyle()`、清除、复制、粘贴和填充；
- 结构：合并、调整行列、增删工作表和切换标签；
- 历史：`undo()`、`redo()`；
- 计算和导出：`recalculate()`、`exportXlsx()`、`exportCsv()`、`download()`；
- 图表和图片：选择、移动、缩放和读取对象。

URL 来源必须同时提供 `src` 和 `urlPolicy`。大文件可以通过 `deferLoadingAboveBytes` 延迟，并在用户确认后调用 `continueDeferredLoad()`。

控制器会销毁自己创建的 Runtime、Worker、工作簿和资源。传入 `runtime` 时，Runtime 仍由调用方负责。

## `useXlsxViewerThumbnails`

```ts
const controllerRef = computed(() => controller)
const { thumbnails } = useXlsxViewerThumbnails(controllerRef, {
  includeHeaders: true,
  resolution: { maxWidth: 200, maxHeight: 132 },
})
```

返回的每个缩略图包含工作表信息、来源区域、尺寸和 `paint(canvas)`。它从控制器读取当前工作簿，不创建第二个工作簿实例。

缩略图只预览有限行列，适合导航，不是完整工作表截图。工作簿变化后应使用响应式返回值重新绘制。
