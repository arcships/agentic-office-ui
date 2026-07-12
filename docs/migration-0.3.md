# 从 0.2.0 升级到 0.3.0

`0.3.0` 把 PPTX 预览和播放加入公开包。八个公开包使用同一版本发布。

## 安装

```bash
pnpm add @arcships/pptx-core@0.3.0 @arcships/vue-pptx@0.3.0
```

如果项目同时使用其他格式，建议统一升级：

```bash
pnpm add @arcships/docx-core@0.3.0 @arcships/vue-docx@0.3.0
pnpm add @arcships/xlsx-core@0.3.0 @arcships/vue-xlsx@0.3.0
pnpm add @arcships/vue-pdf@0.3.0 @arcships/vue-ui@0.3.0
```

## Vue 接入

```ts
import { PptxViewer } from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"
```

`PptxViewer` 默认使用浏览模式。传入 `mode="present"` 后可使用下一步、上一步、暂停、继续、重播、跳页、媒体恢复和全屏。

## 核心入口

平台无关的播放类型和时间安排从 `@arcships/pptx-core` 导入；文档会话、预览和播放控制器从 `@arcships/pptx-core/browser` 导入。

补丁后的 PPTX 渲染器已经包含在浏览器入口中。项目不需要安装 `@aiden0z/pptx-renderer`，也不需要配置 pnpm 补丁。

## 兼容范围

本版本不承诺完整兼容 PowerPoint 全部动画。无法精确执行的内容会通过能力报告标记为近似、静态或未解析。发布前应使用项目自己的 PPTX 素材复核关键页面和动画。
