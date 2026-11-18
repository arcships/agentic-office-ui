# PPTX 使用指南

支持 `.pptx`、`.pptm`、`.ppsx`、`.ppsm`、`.potx` 和 `.potm`。宏内容不会执行；模板和放映文件复用现有浏览/演示模式，不模拟桌面 PowerPoint 的自动启动行为。

## 完整查看器

```vue
<script setup lang="ts">
import { PptxViewer } from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"
</script>

<template>
  <PptxViewer :source="file" mode="browse" height="720px" />
</template>
```

`mode="browse"` 像 PDF/DOCX 一样纵向连续展示全部幻灯片，并提供缩略图、搜索和滚动跳页。`mode="present"` 使用单页舞台；左右键执行上/下一播放步骤，只有当前页步骤结束后才由播放控制器跨页。

浏览模式按可视区域挂载页面，滚动时同步当前页、工具栏页码和缩略图状态。上/下方向键与 PageUp/PageDown 用于纵向跳页；左/右方向键只在演示模式中执行上一步/下一步。

常用属性：

| 属性 | 默认值 | 说明 |
|---|---|---|
| `source` | `null` | `ArrayBuffer`、`Uint8Array` 或文件对象 |
| `mode` | `browse` | 浏览或演示 |
| `height` | `100%` | 组件高度 |
| `minHeight` | 无 | 可选最小高度（数字按 px）；不传时组件不设置最小高度 |
| `initialSlide` | `0` | 初始页，从零开始 |
| `autoplay` | `true` | 演示模式自动执行非点击步骤 |
| `showHiddenSlides` | `false` | 播放时包含隐藏页 |
| `showToolbar` | `true` | 浏览工具栏 |
| `showSidebar` | `true` | 缩略图侧栏 |
| `showSearch` | `true` | 搜索框 |
| `showPlaybackControls` | `true` | 播放控制条 |
| `showWarnings` | `true` | 显示文档 `warnings`（备注不支持、宏已忽略等） |
| `openTimeoutMs` | `30000` | 加载超时毫秒数，超时后进入错误状态；`0` 关闭超时 |
| `externalMedia` | `disabled` | 是否允许外部媒体；设为 `allowed` 时仍需要 `externalMediaUrlPolicy` 放行 |
| `externalMediaUrlPolicy` | `undefined` | 外部媒体的 URL 规则：`baseUrl`、`allowedProtocols`、`allowedOrigins` 等 |

外部媒体默认禁用，只作为静态内容报告。只有同时设置 `externalMedia="allowed"` 和 `externalMediaUrlPolicy`，并且目标 URL 通过 `allowedProtocols` / `allowedOrigins` 校验时才真正加载；`javascript:`、`data:` 和未授权来源一律不加载，避免演示文稿自行发起跟踪请求。打印时 `@page` 尺寸取自幻灯片尺寸，工具栏、侧栏和播放控制条自动隐藏。

`limits` 支持完整的归档预算：`maxInputBytes`、`maxArchiveEntries`、`maxSingleEntryBytes`、`maxUncompressedBytes`、`maxMediaBytes`、`maxCompressionRatio`、`maxArchivePathLength`、`maxXmlBytes`、`maxSingleXmlBytes`、`maxXmlDepth`、`maxXmlAttributeBytes`、`maxTextNodeBytes`、`maxRelationships` 和 `maxConcurrency`。宿主只能收紧默认值，不能放宽硬上限。

模板引用公开 `getController()`、`next()`、`previous()`、`play()`、`pause()`、`resume()`、`reset()`、`goToSlide()`、`enterFullscreen()` 和 `exitFullscreen()`。

## 界面状态

`data-state` 暴露当前视图状态：`empty`、`waiting`（舞台尚未挂载）、`loading`、`ready`、`cancelled` 和 `error`。

- 加载失败、超时或取消才会进入 `error`/`cancelled`，并显示重试按钮；
- 单页渲染失败不会锁死查看器：上一页保持可见，顶部提示第几页失败，并可选择“重试”或“跳过”到相邻页；
- 切换 `browse`/`present` 会重新打开文档，但保留当前页；
- 全屏时组件自动占满屏幕，忽略 `height`；演示模式按幻灯片比例等比缩放并留黑边，不会裁切页面；
- 文档 `warnings` 显示在工具栏下方的提示条中，可用 `warnings` 插槽替换；单页失败提示可用 `slide-error` 插槽替换。

## 自定义播放器界面

`PptxStage`、`usePptxDocument` 和 `usePptxPlayback` 可以绕过完整工具栏。完整示例见 [vue-pptx README](../../packages/vue-pptx/README.md)。

普通纵向 Surface 使用 `session: { renderMode: "list" }`，由文档组合函数负责滚动跳页和缩放。播放器使用 `session: { renderMode: "slide" }`，由播放组合函数负责动画边界和演示跳页。演示模式中不要混用 `document.goTo()` 和 `playback.goToSlide()`，否则文档页面与播放状态可能不同步。

## Surface 事件

`PptxStage` 实际发出以下事件：

```vue
<PptxStage
  @selection-change="onSelection"
  @object-click="onObjectClick"
  @context-menu="onContextMenu"
/>
```

| 事件 | 参数 | 触发时机 |
|---|---|---|
| `selection-change` | `{ kind: "slide", slideIndex }` | 点击幻灯片页面或其中对象 |
| `object-click` | `{ kind: "object", slideIndex, objectKey }` | 点击带稳定对象标识的形状、图片、图表等对象 |
| `context-menu` | `PptxStageContextMenu` 判别联合 | 右键幻灯片或对象 |

`context-menu` 的精确形式为：

```ts
type PptxStageContextMenu = Position & (
  | { kind: "slide"; objectKey?: never }
  | { kind: "object"; objectKey: string }
)
```

其中 `Position` 包含 `slideIndex`、`clientX/Y` 和 `containerX/Y`；容器坐标相对 `PptxStage` 根元素，而不是宿主的外层滚动容器。`slideIndex` 从零开始。对象点击会先产生页面选中，再产生 `object-click`；右键只发 `context-menu`，其中已经包含页面和对象上下文。宿主决定是否显示菜单，不应解析 `objectKey` 的内部字符串格式。

## 兼容性说明

无法精确执行的 PowerPoint 内容会报告为近似、静态或未解析。本库不会用静态显示冒充动画完全兼容。程序可以监听 `capability`、`playback-warning` 和 `playback-error`。

宏格式加载后会通过文档 `warnings` 报告宏已忽略，演讲者备注目前同样以 `NOTES_UNSUPPORTED` 报告。`PptxViewer` 默认在工具栏下方显示这些警告；自定义界面可读取 `usePptxDocument` 返回的 `document.warnings`。
