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
| PPTX | `usePptxDocument` | 文档加载、普通翻页和缩放 |
| PPTX | `usePptxPlayback` | 动画步骤、播放状态和媒体 |

- [DOCX 组合函数](docx.md)
- [XLSX 组合函数](xlsx.md)
- [PPTX 组合函数](pptx.md)
- [Runtime 与资源所有权](runtime.md)

## 共同规则

- 在 Vue 组件作用域中调用，作用域销毁时由组合函数清理自己创建的资源。
- 不要把一个控制器同时交给两个互不相关的编辑器实例。
- 调用方创建并传入的 Runtime 由调用方销毁。
- 快速切换来源时只使用最新结果，界面应观察公开的加载和错误状态。
- 组合函数返回的是状态和命令，不保证包含完整界面；布局、按钮和错误提示由调用方实现。
