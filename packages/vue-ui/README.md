# @arcships/vue-ui

Vue 3 UI 组件库：Office 对象轮廓与区域框选原语，以及签名、上传、缩略图和基础 UI。

## 安装

```bash
pnpm add @arcships/vue-ui @arcships/office-interaction
```

## 使用

```vue
<script setup lang="ts">
import { ref } from "vue"
import { OfficeObjectOutlineLayer, OfficeRegionSelector } from "@arcships/vue-ui"
import type { NormalizedRect } from "@arcships/office-interaction"
import "@arcships/vue-ui/style.css"

const outlines = []
const region = ref<NormalizedRect | null>(null)
</script>

<template>
  <div class="surface-container">
    <!-- Surface 内容 -->
    <OfficeObjectOutlineLayer :items="outlines" />
    <OfficeRegionSelector v-model="region" />
  </div>
</template>
```

## 组件

| 组件 | 用途 |
|---|---|
| `OfficeObjectOutlineLayer` | 在单个 Surface 容器上显示规范化对象轮廓和候选焦点 |
| `OfficeRegionSelector` | 鼠标或键盘框选规范化区域 |
| `SignaturePad` | 签名板 |
| `FileUpload` | 文件上传（含校验） |
| `FileThumbnail` | 文件缩略图 |
| `BoundingBoxCitations` | 引用框标注 |
| `LayoutBlocks` | 版面块 |
| `Spinner` | 加载指示 |
| `Tooltip` | 提示 |
| `OfficeThemeProvider` | 提供 `light` / `dark` / `system` 主题、密度、locale 和共享主题上下文 |
| `OfficeIcon` | 基于 `@lucide/vue` 的统一图标入口，支持尺寸、线宽和可访问名称 |

## 主题与图标

```ts
import "@arcships/vue-ui/style.css"
import "@arcships/vue-ui/theme.css"
```

```vue
<script setup lang="ts">
import { OfficeIcon, OfficeThemeProvider } from "@arcships/vue-ui"
</script>

<template>
  <OfficeThemeProvider theme="system" density="comfortable" locale="zh-CN">
    <button><OfficeIcon name="check" /> 已保存</button>
  </OfficeThemeProvider>
</template>
```

`theme.css` 包含 `--office-*` 语义 token，并把所有声明限定在 `[data-office-theme]` 子树内；`@arcships/vue-ui/tokens.css` 可单独提供浅色 token 契约。默认字体为 Geist / Geist Mono，但包内不打包字体，宿主自行加载或通过 `[data-office-theme]` 覆盖 `--office-font-sans` / `--office-font-mono`。

`OfficeIcon` 的内置注册表只按需引入常用 Lucide 图标；可用 `provideOfficeIcons()` 扩展或覆盖，通过 `:icon` 直接传入自定义组件。完整 token 表、暗色示例、字体加载和宿主覆盖方式见 [Office 设计 token、主题与图标](https://github.com/arcships/agentic-office-ui/blob/master/docs/design/office-theme.md)。

## 事件

选择控件均为受控原语：轮廓层发出候选激活、确认、关闭和层级导航事件；区域选择器发出 `update:modelValue` 与开始、变化、提交、取消事件。组件不保存引用集合，也不规定确认后显示什么界面或执行什么动作。

`FileUpload` 发出 `files-accepted`（通过校验的 `File[]`）和 `files-rejected`（`FileUploadRejection[]`，含 `code`、`file`、`message`）。`SignaturePad` 发出 `update:signature`。

## 文档

- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [设计 token、主题与图标](https://github.com/arcships/agentic-office-ui/blob/master/docs/design/office-theme.md)
- [开始使用](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/getting-started.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)

## License

Apache-2.0
