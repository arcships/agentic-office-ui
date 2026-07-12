# 自定义 PDF 界面

当前稳定的高层入口是 `PdfViewer`。可以关闭下载或旋转按钮、修改来源规则、注入 Runtime，并通过事件把状态展示在外部。

```vue
<script setup lang="ts">
import { ref } from "vue"
import { PdfViewer, type PdfLoadError } from "@arcships/vue-pdf"
import "@arcships/vue-pdf/style.css"

const page = ref(1)
const error = ref<PdfLoadError>()
const props = defineProps<{ file: File | null }>()
</script>

<template>
  <header>
    第 {{ page }} 页
    <span v-if="error">{{ error.code }}：{{ error.message }}</span>
  </header>
  <PdfViewer
    :source="props.file"
    :show-download="false"
    :show-rotate-controls="false"
    @active-page-change="page = $event"
    @document-load-error="error = $event"
  />
</template>
```

当前版本没有公开可从外部驱动 `PdfViewer` 翻页和搜索的 Vue 控制器。因此可以自定义外围状态和部分按钮，但不能只靠稳定公开接口完全替换内部工具栏。需要更底层控制时使用 `createPdfRenderRuntime()` 自己实现页面渲染，并承担分页、缓存、取消和销毁。
