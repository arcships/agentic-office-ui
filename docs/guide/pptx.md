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

`mode="browse"` 用于预览、缩略图、搜索和普通翻页。`mode="present"` 用于逐步动画、媒体、页面切换和全屏。

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

文档组合函数负责普通翻页，播放组合函数负责动画边界和演示跳页。演示模式中不要混用 `document.goTo()` 和 `playback.goToSlide()`，否则文档页面与播放状态可能不同步。

## 兼容性说明

无法精确执行的 PowerPoint 内容会报告为近似、静态或未解析。本库不会用静态显示冒充动画完全兼容。程序可以监听 `capability`、`playback-warning` 和 `playback-error`。
