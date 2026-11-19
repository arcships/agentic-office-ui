# @arcships/vue-ui

Vue 3 UI 组件库：Office 对象轮廓、区域框选、区域截图与标注原语，以及签名、上传、缩略图和基础 UI。

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

### 区域截图与标注

```vue
<script setup lang="ts">
import { computed, ref } from "vue"
import {
  OfficeRegionAnnotator,
  captureRegion,
  createOfficeRegionAnnotationDocument,
  releaseRegionCapture,
  stringifyOfficeRegionAnnotationDocument,
} from "@arcships/vue-ui"

const surfaceRef = ref()
const capture = ref()
const annotations = ref(createOfficeRegionAnnotationDocument())
const annotationsJson = computed(() => stringifyOfficeRegionAnnotationDocument(annotations.value))

async function captureRegionPreview() {
  releaseRegionCapture(capture.value)
  capture.value = await captureRegion({
    surface: surfaceRef.value,
    region: { x: 0.1, y: 0.2, width: 0.5, height: 0.25 },
    scale: window.devicePixelRatio,
  })
  annotations.value = createOfficeRegionAnnotationDocument({ source: capture.value.source })
}
</script>

<template>
  <OfficeRegionAnnotator
    v-if="capture"
    v-model="annotations"
    :image="capture.url"
    :width="capture.width"
    :height="capture.height"
  />
  <pre>{{ annotationsJson }}</pre>
</template>
```

详见[区域截图与标注](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/region-capture-and-annotation.md)。
```

## 组件

| 组件 | 用途 |
|---|---|
| `OfficeObjectOutlineLayer` | 在单个 Surface 容器上显示规范化对象轮廓和候选焦点 |
| `OfficeRegionSelector` | 鼠标或键盘框选规范化区域 |
| `OfficeRegionAnnotator` | 在截图上添加箭头、矩形、文字和高亮，输出可序列化标注文档 |
| `SignaturePad` | 签名板 |
| `FileUpload` | 文件上传（含校验） |
| `FileThumbnail` | 文件缩略图 |
| `BoundingBoxCitations` | 引用框标注 |
| `LayoutBlocks` | 版面块 |
| `Spinner` | 加载指示 |
| `Tooltip` | 提示 |

## 事件

选择控件均为受控原语：轮廓层发出候选激活、确认、关闭和层级导航事件；区域选择器发出 `update:modelValue` 与开始、变化、提交、取消事件。标注编辑器发出 `update:modelValue` 和 `annotation-add` / `annotation-remove` / `annotation-select` / `annotation-change` / `annotation-error`。组件不保存引用集合，也不规定确认后显示什么界面或执行什么动作。

`captureRegion()` 返回的 `url` 由调用方释放；`OfficeRegionAnnotator` 只为 `Blob` 类型的 `image` 创建 URL，并在替换/卸载时撤销。

`FileUpload` 发出 `files-accepted`（通过校验的 `File[]`）和 `files-rejected`（`FileUploadRejection[]`，含 `code`、`file`、`message`）。`SignaturePad` 发出 `update:signature`。

## 文档

- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [区域截图与标注](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/region-capture-and-annotation.md)
- [开始使用](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/getting-started.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)

## License

Apache-2.0
