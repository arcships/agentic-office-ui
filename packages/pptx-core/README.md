# @arcships/pptx-core

## 当前状态

本包提供 PPTX 纵向列表/单页静态预览、播放模型、文档会话、播放控制器、页面切换和受控媒体，从 `0.3.0` 起公开发布。实现范围和验收结果见 [实现设计](../../docs/pptx-playback-implementation-design.md) 与 [正式播放验收记录](../../docs/pptx-playback-acceptance-results.md)。

```bash
pnpm add @arcships/pptx-core@0.5.2
```

浏览器能力从 `@arcships/pptx-core/browser` 导入。补丁后的渲染器代码已经包含在浏览器入口中，使用方不需要安装或配置 `@aiden0z/pptx-renderer`；其普通运行依赖由 npm 自动安装。

`createPptxPreviewSession()` 默认使用 `renderMode: "list"`，纵向连续渲染全部幻灯片并按可视区域挂载。`createPptxDocumentSession()` 默认使用 `renderMode: "slide"`，供动画播放控制器使用。两者都可以通过会话选项显式指定 `renderMode`；`listOptions` 控制窗口化、初始挂载页数和预渲染范围。

## 入口边界

- `src/index.ts`、`src/types.ts`：平台无关的公开类型、错误和默认限制；
- `src/browser-types.ts`：浏览器会话接口；
- `src/browser.ts`：静态预览、一次解包文档会话和浏览器出口；
- `src/playback/`：播放类型、对象身份、能力报告、时间树和属性轨道；
- `src/browser/`：对象标记、播放解析、时钟、切换、媒体和控制器。

根入口不得导入 `@aiden0z/pptx-renderer`，不得访问 `window`、`document`、`DOMParser`、`Worker` 或媒体元素。浏览器实现不得从 `apps/pptx-playback-lab` 导入代码；探索台只提供验证结论。

本包不会宣称完整兼容 PowerPoint 全部动画。无法精确执行的内容会按能力报告标记为近似、静态或未解析。

## 文档

- [PPTX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/pptx.md)
- [API 导航](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/README.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)

## 验证

仓库根目录提供统一命令：

```bash
pnpm typecheck:pptx
pnpm build:pptx
pnpm test:pptx
pnpm test:pptx:blackbox
```
