# 把用户选择变成 Agent 可定位的引用

> 本页对应 `0.6.0` 源码候选。npm 稳定版 `0.5.4` 尚未包含统一引用 API。

Agentic Office UI 只负责把 Office Surface 上的一次选择转换成精确引用，并通过事件交给宿主。组件不保存确认后的引用集合，不提供聊天框、提示词协议或 Agent 调用。

## 选择模式

四个最小 Surface 共用三个受控模式：

| `selectionMode` | 用户手势 | 典型结果 |
|---|---|---|
| `content` | 拖选文字、选择单元格或当前页 | 精确文字、单元格、范围或幻灯片引用 |
| `object` | 悬停并点击可见对象 | 图表、图片、形状、分组或其他格式对象引用 |
| `region` | 在页面、幻灯片或工作表内框选 | 与当前文档位置绑定的规范化区域引用 |

`selectionMode` 由宿主控制。Surface 只解释下一次手势，不决定按钮、工具栏或模式切换方式。

## 最小接入

直接使用共享类型或纯函数时，宿主应把 `@arcships/office-interaction` 声明为直接依赖。`0.6.0` 发布后可按需安装：

```bash
pnpm add @arcships/vue-xlsx @arcships/xlsx-core @arcships/office-interaction
```

下面的宿主只保存最后一次确认引用。实际产品可以把事件放入自己的引用列表、表单或对话状态：

```vue
<script setup lang="ts">
import { shallowRef } from "vue"
import type { OfficeReferenceConfirmEvent } from "@arcships/office-interaction"
import { XlsxSheetSurface, useXlsxViewerController } from "@arcships/vue-xlsx"

const props = defineProps<{ file: ArrayBuffer }>()
const controller = useXlsxViewerController({
  file: props.file,
  fileName: "budget-2026.xlsx",
})
const selectedReference = shallowRef<OfficeReferenceConfirmEvent["reference"]>()

function onReferenceConfirm(event: OfficeReferenceConfirmEvent) {
  selectedReference.value = event.reference
}
</script>

<template>
  <XlsxSheetSurface
    :controller="controller"
    document-id="budget-2026.xlsx"
    selection-mode="content"
    @reference-confirm="onReferenceConfirm"
  />
</template>
```

`documentId` 应使用宿主能够回到同一输入文件的标识，不要使用临时 DOM id 或 Blob URL。文件内容变化时，格式适配器会生成新的 document revision；宿主可用 `resolveReference()` 检查旧引用是否仍能精确定位。

## 四种 Surface 的统一合同

| Surface | 基础输入 | 阶段一引用 |
|---|---|---|
| `DocxDocumentSurface` | DOCX 文档模型与分页状态 | 精确文字、页面、页面区域 |
| `XlsxSheetSurface` | XLSX viewer controller | 工作表、单元格、范围、行列、图表、工作表区域 |
| `PptxStage` | PPTX 文档会话 | 幻灯片、对象内精确文字、可见对象、组合层级、幻灯片区域 |
| `PdfSurface` | PDF render document | 精确字符范围、页面、页面区域 |

共同属性：

- `documentId?: string`
- `selectionMode?: "content" | "object" | "region"`
- `emitReferenceCandidates?: boolean`

共同事件：

| 事件 | 宿主如何使用 |
|---|---|
| `document-revision-change` | 记录当前引用属于哪一版输入 |
| `reference-candidate-change` | 在对象模式显示候选轮廓或候选名称 |
| `reference-confirm` | 保存引用，或与用户语言指令一起交给宿主 Agent |
| `region-draft-change` | 显示区域拖选过程或辅助标注 UI |
| `selection-cancel` | 退出宿主自己的选择状态 |
| `reference-resolve` | 显示精确、迁移、歧义、未找到或不支持状态 |
| `reference-error` | 按结构化错误码处理失败，不解析界面文案 |

共同 expose：

- `getDocumentRevision()`
- `hitTest()`
- `describeReference()`
- `resolveReference()`
- `scrollToReference()`
- `captureReferencePreview()`

当前 Surface 不负责截图位图采集；`captureReferencePreview()` 会用 `CAPTURE_UNSUPPORTED` 明确拒绝。需要截图或标注时，宿主可沿用引用中的页面、幻灯片、工作表和规范化区域，接入自己的渲染与标注管线。

## 自定义选择界面

如果产品不想直接使用 Surface 内置手势，可以组合 `@arcships/vue-ui` 的两个受控原语：

- `OfficeObjectOutlineLayer`：显示候选轮廓，发出激活、确认、取消和层级导航事件；
- `OfficeRegionSelector`：在 `0..1` 规范坐标中发出区域开始、变化、确认和取消事件。

两者都不读取 Office 文件，也不生成格式 locator。宿主把原语事件交给对应格式适配器，或直接消费 Surface 的统一事件即可。

完整字段、解析状态和错误码见[公开 API 合同](../api/public-api-contract.md)；对象范围与后续层级见[对象语义与选择设计](../product/object-semantics-and-selection.md)。
