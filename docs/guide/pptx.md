# PPTX 使用指南

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
| `initialSlide` | `0` | 初始页，从零开始 |
| `autoplay` | `true` | 演示模式自动执行非点击步骤 |
| `showHiddenSlides` | `false` | 播放时包含隐藏页 |
| `showToolbar` | `true` | 浏览工具栏 |
| `showSidebar` | `true` | 缩略图侧栏 |
| `showSearch` | `true` | 搜索框 |
| `showPlaybackControls` | `true` | 播放控制条 |
| `externalMedia` | `disabled` | 是否允许外部媒体 |

模板引用公开 `getController()`、`next()`、`previous()`、`play()`、`pause()`、`resume()`、`reset()`、`goToSlide()`、`enterFullscreen()` 和 `exitFullscreen()`。

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
