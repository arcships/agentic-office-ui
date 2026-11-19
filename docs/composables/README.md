# 组合函数手册

组合函数适合需要自定义工具栏、状态栏、缩略图或播放界面的项目。它们不是全局单例：每次调用都会创建一份独立状态，通常应在一个组件的 `setup` 作用域内创建。

| 格式 | 组合函数 | 主要用途 |
|---|---|---|
| DOCX | `useDocxEditor` | 编辑状态、选区、历史记录和命令 |
| DOCX | `useDocxModel` | 把 DOCX 字节解析为 `DocModel` |
| DOCX | `useDocxPageThumbnails` | 从编辑器页面生成缩略图 |
| DOCX | 主题、样式、批注、修订等辅助组合函数 | 从编辑控制器派生局部能力 |
| XLSX | `useXlsxViewerController` | 文件、工作表、选区、编辑和导出 |
| XLSX | `useXlsxViewerThumbnails` | 从控制器生成工作表缩略图 |
| PPTX | `usePptxDocument` | 文档加载、纵向滚动跳页/单页跳转、缩放和会话生命周期 |
| PPTX | `usePptxPlayback` | 动画步骤、播放状态和媒体 |
| 通用 | `useViewerKeys` | DOCX/PDF/PPTX 共享键盘映射、页面/缩放/搜索/Escape 命令 |
| 通用 | `useFileDrop` | DOCX/XLSX/PDF/PPTX 共享拖放导入：类型校验、嵌套 drag 深度、轮廓和拒绝反馈 |

- [DOCX 组合函数](docx.md)
- [XLSX 组合函数](xlsx.md)
- [PPTX 组合函数](pptx.md)
- [查看器键盘导航](../guide/viewer-keyboard.md)
- [Runtime 与资源所有权](runtime.md)

## 拖放导入

`useFileDrop` 把四个查看器的拖放行为收拢到一处，宿主也可以拿它在自定义 Surface 上复用：

```ts
import { useFileDrop } from "@arcships/vue-ui"

const drop = useFileDrop({
  accept: { extensions: [".xlsx"], mimeTypes: ["text/csv"] },
  onFile: (file) => controller.loadFile(file),
  onReject: () => (message.value = "仅支持 .xlsx 和 .csv 文件"),
})
```

- `accept` 同时匹配扩展名和 MIME 类型（忽略大小写）；`onFile` 只会收到第一个匹配的文件。
- `onDragEnter`/`onDragOver`/`onDragLeave`/`onDrop` 直接绑定到落点根元素。深度来自 `dragenter`/`dragleave` 的嵌套层级，`dragover` 只更新 `dropEffect`，重复触发不会把轮廓“卡住”（对应 DOCX 多页外壳的轮廓缺陷 #27）。
- `isDragOver` 用于绘制 `FILE_DROP_OUTLINE`；放下、拖离、`dragend` 或在窗口其他位置放下都会清除。
- `enabled` 可以是布尔值或 getter，用于“只有可编辑时才接受拖入”这类条件。

## 共同规则

- 在 Vue 组件作用域中调用，作用域销毁时由组合函数清理自己创建的资源。
- 不要把一个控制器同时交给两个互不相关的编辑器实例。
- 调用方创建并传入的 Runtime 由调用方销毁。
- 快速切换来源时只使用最新结果，界面应观察公开的加载和错误状态。
- 组合函数返回的是状态和命令，不保证包含完整界面；布局、按钮和错误提示由调用方实现。
