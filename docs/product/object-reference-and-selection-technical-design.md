# Office 对象引用与选择技术设计

> 状态：T1–T6 已完成；四格式阶段一统一引用、Surface 事件、demo 闭环与发布验收均已通过
>
> 代码基线：`0.6.0` 源码候选（npm 稳定版为 `0.5.4`）
>
> 产品范围：见[《Office 对象语义与意图选择》](./object-semantics-and-selection.md)
>
> 本文定义组件库内部和公开 API 的技术边界。它不定义 Agent RPC、提示词模板或文档修改协议。

## 1. 目标与约束

本设计要让宿主从四种 Office Surface 得到一种稳定、可序列化、可重新定位的“用户所指对象”，并始终保留页面区域兜底。

必须满足：

1. 用户看到的选择轮廓与交给宿主的引用指向同一对象；
2. 引用不包含 DOM 节点、Vue 实例、Blob URL 或只能在当前帧使用的像素坐标；
3. 原生对象、结构推断、视觉识别和人工区域不会被混成同一种准确度；
4. 对象类型、边界、层级和跨版本重定位可靠度分别表达；
5. 文件更新后，系统要么重新定位成功，要么返回歧义或失效，不能静默指向另一个对象；
6. 默认文字、单元格和幻灯片交互不因对象模式而退化；
7. 组件库只产出对象候选、稳定引用和选择事件；引用如何组织、展示、解释和交给 Agent，全部由宿主负责。

本期不做：

- 不建立通用 Agent 消息总线；
- 不定义目标/参考物等业务角色、提示词片段或 Agent 动作；
- 不规定引用托盘、工具栏、候选菜单、面包屑或图层面板的产品形态；
- 不把 OOXML 节点、PDF glyph 或 SVG path 直接暴露为用户对象；
- 不要求第一阶段完成 OCR、表格视觉识别或 3D 图层视图；
- 不承诺引用在任意大幅改写后的文档中自动恢复；
- 不把截图位图本身塞进公共 JSON 引用。

## 2. 设计启动时的事实与当前实现

本设计启动时已经有不少格式内定位基础，但还没有统一引用层：

| 格式 | 可复用基础 | 当前缺口 |
|---|---|---|
| DOCX | 段落、表格行列、图片、页眉页脚等 DOM 标记；编辑器已有精确文本范围类型 | 对外选择只有文字和页码，`objectClick` 声明但没有真实发出 |
| XLSX | 单元格范围；图表、图片、形状、表格及图表内部元素模型 | Surface 只发出范围选择，形状层不可点击，声明的 `objectClick` 未形成统一行为 |
| PPTX | `objectKey`、shape id、group path、名称和来源构成的对象身份；动画和动作已关联对象 | 对外对象仍是通用 `object`，没有类型、内部层级和文本范围 |
| PDF | `PageTextSlice`、字符几何、选区矩形和页面坐标转换 | 没有区域选择及图片、表格、链接、表单、批注对象层 |

因此，本设计没有重写格式解析器，而是在现有模型和 Surface 之间增加统一的定位器适配、选择会话和受控事件。当前阶段一已经把页面/幻灯片/工作表、文字、单元格/范围/行列和人工区域接入四种 Surface；阶段二之后的高频可见对象、层级与不可见行为仍按产品矩阵推进。

## 3. 核心架构决定

### 3.1 新增公开的无框架基础包

目标架构新增第九个公开包：

```text
@arcships/office-interaction
```

它是零 Vue 依赖、零格式解析依赖的 TypeScript 包，负责：

- 统一引用、对象描述、命中候选、可靠度和短生命周期选择会话类型；
- 候选排序、候选导航、确认/取消、矩形归一化和 JSON 校验等纯函数；
- 引用与解析结果的稳定序列化。

不把这些类型放入 `@arcships/vue-ui`，因为格式核心层不应依赖 Vue；也不把它们只放入私有的 `@arcships/office-runtime`，因为外部宿主必须有唯一、可直接安装的类型来源。

`@arcships/office-runtime` 继续保持私有，只提供输入、任务、取消、资源限制等无状态运行工具。新增公开包不能导入它，任何公开 `.d.ts` 也不能引用它。

目标依赖方向：

```text
                  @arcships/office-interaction
                     ↑          ↑
      docx-core / xlsx-core / pptx-core       vue-ui
                     ↑
      vue-docx / vue-xlsx / vue-pptx / vue-pdf
                     ↑
                 宿主应用与 demo

      office-runtime 仍由各格式 Runtime 私下使用，不进入公开引用链
```

新包已实现并进入公开 API、打包复现与外部消费门禁，因此当前包清单统一按九个公开包维护。`vue-ui` 已依赖该公开包并提供通用选择控件；DOCX、XLSX、PPTX 与 PDF 格式包均已依赖它并提供阶段一格式适配器。四种 Surface 已统一发出候选、确认、区域草稿、取消、revision、解析和错误事件；确认后的引用集合、意图 UI 与 Agent 工作流仍严格留在宿主层。

### 3.2 四层职责

| 层 | 职责 | 禁止承担 |
|---|---|---|
| 格式核心层 | 从原生模型生成语义对象、原生定位器、父子关系和内容证据 | 指针事件、Vue 状态、屏幕坐标 |
| Surface 适配层 | 把渲染元素与核心对象关联，完成坐标转换、命中和当前可见段索引 | Agent 调用、持久化宿主会话 |
| 通用选择层 | 管理短生命周期的模式、悬停、候选导航、区域草稿、确认和取消 | 保存引用集合、分配业务角色、决定确认后的动作 |
| 宿主层 | 保存已确认引用，决定多选、角色、视觉呈现、语言指令、Agent 调用与结果重载 | 依赖内部 DOM 标记或私有格式对象 |

`@arcships/vue-ui` 只把轮廓层、区域框等选择原语作为核心能力；它们保持受控、弱样式并只发出选择事件。候选菜单、面包屑、图层面板、引用托盘和标注编辑器属于宿主产品，不是本库公共组件，也不是 Surface 接入和阶段验收的前置条件。

### 3.3 选择器的输出边界

```text
格式对象索引 / Surface 命中
            ↓ candidates
       无头选择会话
            ↓ referenceConfirm / selectionCancel
   宿主 UI、引用集合、Agent 与业务工作流
```

选择器可以解释指针、键盘、触控和区域手势，但不能根据确认结果显示业务工具栏、生成提示词、调用修改命令或决定引用之间的角色。宿主也不需要采用库提供的视觉组件；只要消费同一 candidate/reference/event 合同，就可以实现菜单、聊天输入框、右键操作、3D 图层或其他交互。

## 4. 四种数据对象不能混用

实现中必须区分以下对象：

| 名称 | 生命周期 | 是否可序列化 | 用途 |
|---|---|---|---|
| `OfficeHitCandidate` | 一次悬停或点击 | 否 | 当前指针位置有哪些可能对象，允许携带 DOM 命中证据 |
| `OfficeObjectDescriptor` | 当前文档修订和布局 | 是，但不建议持久保存 | 当前对象的名称、层级、内容摘要和可见几何 |
| `OfficeObjectReference` | 可跨组件和修订保存 | 是 | 宿主持有的稳定对象引用 |
| `OfficeReferenceConfirmEvent` | 一次确认操作 | 是 | 把确认后的引用及输入来源交给宿主，不表达后续业务动作 |

DOM 元素只能出现在 `OfficeHitCandidate.runtimeTarget` 这一类内部字段中。形成 `OfficeObjectReference` 前必须转换为格式定位器、内容指纹和规范化区域。

## 5. 公共数据合同

以下类型是第一版目标合同。字段命名在实现前可以做小幅调整，但职责分离和判别联合不能删除。

### 5.1 文档身份与修订

```ts
export type OfficeFormat = "docx" | "xlsx" | "pptx" | "pdf"

export interface OfficeDocumentRevision {
  format: OfficeFormat
  documentId: string
  revision: string
  contentDigest?: `sha256:${string}`
}
```

- `documentId` 是同一业务文档的逻辑身份，优先由宿主传入；没有传入时，由 Surface 会话生成，只保证当前会话内唯一；
- `revision` 由格式 Document/Workbook session 单独负责，不是可由 Surface prop 随意覆盖的状态；它在加载新字节或核心模型提交修改后必须变化；
- `contentDigest` 在运行时已经持有完整字节时计算，用于发现同名文件内容已变化，不作为唯一业务身份；
- Surface 不从文件名猜测 `documentId`。

引用始终绑定创建时修订。新修订到来后先走解析器，不允许仅因为 `documentId` 相同就继续使用旧坐标。

revision 的生成规则固定如下：

- 从文件或 URL 加载时，Runtime 在允许选择前计算来源字节的 SHA-256；初始 revision 使用该 digest；
- 同一可编辑 session 内使用 `初始 digest 或 session UUID + 单调 modelVersion`，每次成功提交模型命令后递增；
- 直接传入内存模型且没有来源字节时，session 使用加密安全随机 UUID，不能用时间戳或模块级计数器；
- 重新挂载后若要恢复历史引用，宿主必须继续传入同一 `documentId`；新 session 以当前字节 digest 建立 revision，再由 resolver 比较旧引用；
- `documentRevisionChange` 事件和 `getDocumentRevision()` expose 让宿主观察 revision，但宿主不成为它的第二个写入者。

### 5.2 对象类型

```ts
export type OfficeObjectKind =
  | "document" | "page" | "section" | "paragraph" | "heading" | "list-item"
  | "text-range" | "text-block" | "table" | "table-row" | "table-cell" | "image"
  | "workbook" | "worksheet" | "cell" | "cell-range" | "row" | "column"
  | "chart" | "chart-series" | "chart-point" | "chart-legend-entry" | "shape" | "text-box" | "group"
  | "slide" | "layout" | "master" | "comment" | "annotation" | "link" | "form-control"
  | "region" | "formula" | "style" | "conditional-format" | "data-validation"
  | "animation" | "transition" | "action" | "note" | "bookmark" | "field"
  | "named-range" | "protection" | "tracked-change" | "cross-reference" | "form-rule"
  | "attachment" | "layer" | "signature" | "permission" | "unknown"
```

这是跨格式的产品语义，不取代格式核心模型中的精细类型。格式特有信息放入定位器和 `facets`，不要不断增加只对单一解析器有意义的通用 `kind`。

### 5.3 可靠度

```ts
export type ReliabilityLevel = "exact" | "likely" | "uncertain" | "unknown"

export interface ReliabilityDimension {
  level: ReliabilityLevel
  score?: number       // 仅当识别器经过校准时使用，范围 0..1
  reasonCodes: readonly string[]
}

export interface OfficeObjectReliability {
  semantic: ReliabilityDimension
  boundary: ReliabilityDimension
  hierarchy: ReliabilityDimension
  relocation: ReliabilityDimension
}

export type RecognitionSource = "native" | "structural" | "visual" | "manual"
```

四个维度不得折算成一个公开的“总置信度”。`score` 是可选校准结果，不是为了 UI 排序随手生成的小数。候选排序可以使用内部评分，但公共反馈使用 `level + reasonCodes`。reason code 是稳定、机器可读的英文标识，界面文案由各包的本地化层转换。

人工区域的典型结果是：边界为 `exact`，语义为 `unknown`。这能表达“用户准确画了这里，但程序不知道它是什么”。

### 5.4 格式定位器

```ts
export type OfficeObjectLocator =
  | { type: "format"; format: "docx"; value: DocxObjectLocator }
  | { type: "format"; format: "xlsx"; value: XlsxObjectLocator }
  | { type: "format"; format: "pptx"; value: PptxObjectLocator }
  | { type: "format"; format: "pdf"; value: PdfObjectLocator }
  | { type: "manual-region"; format: OfficeFormat; value: ManualRegionLocator }
```

首批定位器按以下原生信息建立：

| 格式 | 定位器组成 | 重定位证据 |
|---|---|---|
| DOCX | part/区域、段落或表格路径、run/字符偏移、图片 child index；可用时保存 OOXML 原生 id | 精确文字及前后文、样式/标题、祖先路径、媒体关系和附近区域 |
| XLSX | sheet id/name、A1 范围、table name、drawing/chart/shape id、内部系列或点索引 | 公式/显示值摘要、对象名称、锚定单元格、祖先表格或图表 |
| PPTX | slide id/index、现有 `objectKey`、shape id、source、group path；内部段落/单元格/图表路径 | 对象名称、文本摘要、类型、组路径和规范化边界 |
| PDF | page index、`PageTextSlice`、原生 annotation/form/link id；视觉对象使用页面区域 | 精确文字及前后文、页面尺寸、规范化区域和视觉特征摘要 |

索引可以作为同一修订内的精确定位依据，但不能单独承担跨修订身份。例如 `nodeIndex: 12` 或 `slideIndex: 4` 必须与修订、原生 id 或内容证据一起使用。

格式定位器使用以下判别联合。`path` 保存结构路径，不保存 DOM 查询器：

```ts
export type DocxPartLocator =
  | { kind: "body" }
  | { kind: "header" | "footer"; partName: string }
  | { kind: "footnote" | "endnote"; noteId: string }

export type DocxPathSegment =
  | { kind: "node" | "table" | "row" | "cell" | "paragraph" | "run"; index: number; nativeId?: string }

export type DocxBehaviorOwner =
  | { scope: "document" }
  | { scope: "part"; part: DocxPartLocator; path: readonly DocxPathSegment[] }

export interface DocxBehaviorLocator {
  kind: "behavior"
  behavior: "style" | "field" | "bookmark" | "section" | "tracked-change-state" | "comment-state" | "cross-reference"
  owner: DocxBehaviorOwner
  instanceId: string
}

export type DocxObjectLocator =
  | { kind: "page"; pageIndex: number }
  | { kind: "structure"; part: DocxPartLocator; path: readonly DocxPathSegment[] }
  | {
      kind: "text-range"
      part: DocxPartLocator
      start: { path: readonly DocxPathSegment[]; offset: number }
      end: { path: readonly DocxPathSegment[]; offset: number }
    }
  | {
      kind: "image"
      part: DocxPartLocator
      paragraphPath: readonly DocxPathSegment[]
      childIndex: number
      relationId?: string
    }
  | {
      kind: "comment"
      part: DocxPartLocator
      commentId: string
      anchorPath?: readonly DocxPathSegment[]
    }
  | DocxBehaviorLocator

export interface XlsxSheetLocator {
  sheetId?: string
  name: string
  index: number
}

export type XlsxBehaviorLocator =
  | {
      kind: "behavior"
      behavior: "formula" | "conditional-format" | "data-validation"
      scope: { kind: "sheet"; sheet: XlsxSheetLocator; appliesToA1: string }
      instanceId: string
    }
  | {
      kind: "behavior"
      behavior: "named-range" | "protection"
      scope: { kind: "workbook" } | { kind: "sheet"; sheet: XlsxSheetLocator }
      instanceId: string
    }

export type XlsxObjectLocator =
  | { kind: "worksheet"; sheet: XlsxSheetLocator }
  | { kind: "range"; sheet: XlsxSheetLocator; a1: string }
  | { kind: "row" | "column"; sheet: XlsxSheetLocator; start: number; end: number }
  | { kind: "table"; sheet: XlsxSheetLocator; tableName: string; id?: string }
  | {
      kind: "drawing"
      sheet: XlsxSheetLocator
      objectType: "chart" | "image" | "shape" | "form-control"
      id: string
      name?: string
    }
  | {
      kind: "chart-element"
      sheet: XlsxSheetLocator
      chartId: string
      element: "series"
      seriesIndex: number
    }
  | {
      kind: "chart-element"
      sheet: XlsxSheetLocator
      chartId: string
      element: "point"
      seriesIndex: number
      pointIndex: number
    }
  | {
      kind: "chart-element"
      sheet: XlsxSheetLocator
      chartId: string
      element: "legend-entry"
      legendIndex: number
      seriesIndex?: number
    }
  | { kind: "comment"; sheet: XlsxSheetLocator; cellA1: string; nativeId?: string }
  | XlsxBehaviorLocator

export interface PptxSlideLocator {
  slideId?: string
  index: number
}

export interface PptxObjectKeyLocator {
  objectKey: string
  shapeId?: string
  source: "slide" | "layout" | "master"
  groupPath: readonly string[]
}

export type PptxObjectLocator =
  | { kind: "slide"; slide: PptxSlideLocator }
  | { kind: "object"; slide: PptxSlideLocator; object: PptxObjectKeyLocator }
  | {
      kind: "text-range"
      slide: PptxSlideLocator
      object: PptxObjectKeyLocator
      start: { paragraphIndex: number; offset: number }
      end: { paragraphIndex: number; offset: number }
    }
  | {
      kind: "sub-object"
      slide: PptxSlideLocator
      object: PptxObjectKeyLocator
      path: readonly (
        | { kind: "table-cell"; rowIndex: number; cellIndex: number }
        | { kind: "chart-series"; seriesIndex: number }
        | { kind: "chart-point"; seriesIndex: number; pointIndex: number }
        | { kind: "chart-legend-entry"; legendIndex: number; seriesIndex?: number }
      )[]
    }
  | {
      kind: "behavior"
      slide: PptxSlideLocator
      behavior: "animation" | "transition" | "action" | "master" | "note"
      nativeId: string
      ownerObjectKey?: string
    }

export type PdfObjectLocator =
  | { kind: "page"; pageIndex: number }
  | { kind: "text-range"; pageIndex: number; charIndex: number; charCount: number }
  | { kind: "native-object"; pageIndex: number; objectType: "link" | "form" | "annotation"; nativeId: string }
  | { kind: "visual-object"; pageIndex: number; providerId: string; objectId: string; region: NormalizedRect }
  | {
      kind: "behavior"
      pageIndex?: number
      behavior: "link-action" | "form-rule" | "annotation-state" | "attachment" | "layer" | "signature" | "permission"
      nativeId: string
    }
```

阶段四行为在第一版合同中已经使用封闭判别值，不先公开任意字符串再破坏性收窄。未列入产品首批矩阵的行为以后通过新增联合成员和 schema 版本兼容规则扩展；阶段一、二不得因为类型已存在就宣称行为已经实现。

每个 behavior locator 的 `instanceId` 都是必填且在其 owner/scope 内唯一。优先使用文件原生 id；原生格式没有 id 时，格式核心层用名称、规范路径、规则索引和内容哈希生成确定性 id。重新解析同一字节必须得到同一 id，不能使用随机 UUID。随机 UUID 只用于当前 session 和用户引用的 `referenceId`。

### 5.5 人工区域、截图与可选标注坐标

```ts
export interface NormalizedRect {
  x: number
  y: number
  width: number
  height: number
}

export type ManualRegionLocator =
  | { space: "page"; pageIndex: number; rect: NormalizedRect }
  | { space: "slide"; slideIndex: number; rect: NormalizedRect }
  | {
      space: "sheet"
      sheetId: string
      start: { row: number; col: number; xOffset: number; yOffset: number }
      end: { row: number; col: number; xOffset: number; yOffset: number }
    }

export interface NormalizedPoint { x: number; y: number }

export type RegionAnnotation =
  | { id: string; kind: "rectangle" | "ellipse"; rect: NormalizedRect; color?: string }
  | { id: string; kind: "arrow"; start: NormalizedPoint; end: NormalizedPoint; color?: string }
  | { id: string; kind: "freehand"; points: readonly NormalizedPoint[]; color?: string }
  | { id: string; kind: "text" | "pin"; point: NormalizedPoint; text: string; color?: string }
```

页面和幻灯片区域使用 `[0, 1]` 规范化坐标；工作表区域使用单元格地址加单元格内部偏移，避免滚动和缩放改变引用。`RegionAnnotation` 只是供宿主复用的可选几何数据类型，不进入 `OfficeObjectReference`，也不是选择器必须保存的状态。宿主可以自行实现箭头、矩形、文字或其他标注界面，并决定如何保存和传输。所有 annotation 点和矩形都相对“已选区域自身”的左上角与宽高归一化，工作表标注也先映射到所选 sheet region，再解释 `[0, 1]` 坐标。

截图只是同一区域在某个修订、缩放和像素密度下的预览附件。公共引用只保存区域，不保存标注、截图或 Blob URL。宿主如需预览位图，应通过显式 `captureReferencePreview()` 得到临时 Blob，并自行决定是否标注、上传、保留或脱敏。

### 5.6 引用、描述和确认事件

```ts
export interface OfficeObjectFingerprint {
  exactText?: string
  prefixText?: string
  suffixText?: string
  objectName?: string
  contentHash?: string
  ancestorKeys?: readonly string[]
}

export type JsonValue =
  | null | boolean | number | string
  | readonly JsonValue[]
  | { readonly [key: string]: JsonValue }

export type DocxReferenceKind =
  | "page" | "section" | "paragraph" | "heading" | "list-item" | "text-range"
  | "table" | "table-row" | "table-cell" | "image" | "comment" | "region"
  | "style" | "field" | "bookmark" | "tracked-change" | "cross-reference"

export type XlsxReferenceKind =
  | "worksheet" | "cell" | "cell-range" | "row" | "column" | "table" | "chart"
  | "chart-series" | "chart-point" | "chart-legend-entry" | "image" | "shape" | "form-control" | "comment" | "region"
  | "formula" | "conditional-format" | "data-validation" | "named-range" | "protection"

export type PptxReferenceKind =
  | "slide" | "text-range" | "text-box" | "shape" | "image" | "table" | "chart" | "group"
  | "table-cell" | "chart-series" | "chart-point" | "chart-legend-entry" | "region"
  | "animation" | "transition" | "action" | "master" | "note"

export type PdfReferenceKind =
  | "page" | "text-range" | "text-block" | "image" | "table" | "table-row" | "table-cell"
  | "link" | "form-control" | "annotation" | "region" | "action" | "form-rule"
  | "attachment" | "layer" | "signature" | "permission"

type FormatLocator<F extends OfficeFormat, L> = { type: "format"; format: F; value: L }
type RegionLocator<F extends OfficeFormat, S extends ManualRegionLocator["space"]> = {
  type: "manual-region"
  format: F
  value: Extract<ManualRegionLocator, { space: S }>
}

interface OfficeObjectReferenceBase<F extends OfficeFormat, K extends OfficeObjectKind, L, R extends ManualRegionLocator> {
  schemaVersion: 1
  referenceId: string
  document: OfficeDocumentRevision & { format: F }
  kind: K
  source: RecognitionSource
  locator: L
  fingerprint?: OfficeObjectFingerprint
  fallbackRegion?: R
  reliability: OfficeObjectReliability
}

export type OfficeObjectReference =
  | OfficeObjectReferenceBase<
      "docx",
      DocxReferenceKind,
      FormatLocator<"docx", DocxObjectLocator> | RegionLocator<"docx", "page">,
      Extract<ManualRegionLocator, { space: "page" }>
    >
  | OfficeObjectReferenceBase<
      "xlsx",
      XlsxReferenceKind,
      FormatLocator<"xlsx", XlsxObjectLocator> | RegionLocator<"xlsx", "sheet">,
      Extract<ManualRegionLocator, { space: "sheet" }>
    >
  | OfficeObjectReferenceBase<
      "pptx",
      PptxReferenceKind,
      FormatLocator<"pptx", PptxObjectLocator> | RegionLocator<"pptx", "slide">,
      Extract<ManualRegionLocator, { space: "slide" }>
    >
  | OfficeObjectReferenceBase<
      "pdf",
      PdfReferenceKind,
      FormatLocator<"pdf", PdfObjectLocator> | RegionLocator<"pdf", "page">,
      Extract<ManualRegionLocator, { space: "page" }>
    >

type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never
export type OfficeObjectReferenceDraft = DistributiveOmit<OfficeObjectReference, "referenceId">

export interface OfficeObjectDescriptor {
  objectId: string // 只在当前 revision 的对象索引中稳定
  draft: OfficeObjectReferenceDraft
  label: string
  path: readonly { kind: OfficeObjectKind; label: string }[]
  parentObjectId?: string
  childrenState: "none" | "available" | "lazy" | "unknown"
  content?: { text?: string; value?: string; formula?: string }
  visual?: {
    fragments: readonly (
      | {
          container: { space: "page"; pageIndex: number } | { space: "slide"; slideIndex: number }
          rect: NormalizedRect
          zIndex?: number
        }
      | {
          container: { space: "sheet"; sheetId: string }
          region: Extract<ManualRegionLocator, { space: "sheet" }>
          zIndex?: number
        }
    )[]
    layoutVersion: string
  }
  facets?: Readonly<Record<string, JsonValue>>
}

export interface OfficeHitCandidate<RuntimeTarget = unknown> {
  candidateId: string
  draft: OfficeObjectReferenceDraft
  preview?: Pick<OfficeObjectDescriptor, "label" | "path" | "visual">
  runtimeTarget?: RuntimeTarget
  hit: "direct" | "inside" | "ancestor" | "inferred"
  depth: number
  zIndex?: number
}
```

`fingerprint` 默认只保留重定位所需的短证据。文本证据默认每段最多 256 个 Unicode 字符，完整选中文字只出现在当前修订的 `descriptor.content` 中，由宿主决定是否读取、保存或发送。任何遥测默认不得记录这些文本字段。

descriptor 和命中候选都只携带没有 `referenceId` 的 draft。只有确认函数可以把 draft 提升为 reference；它使用调用方传入的 `referenceId`，或显式调用库提供的 `createOfficeReferenceId()`，不能让纯 reducer 隐式生成随机值。悬停、对象列表和指针移动不能不断制造已提交引用。`objectId` 只用于当前对象索引和父子查询，不能作为持久引用。

确认事件只回答“用户确认了哪个对象”，不回答“接下来要对它做什么”：

```ts
export interface OfficeReferenceSnapshot {
  label: string
  path: readonly { kind: OfficeObjectKind; label: string }[]
  content?: {
    text?: string
    value?: string
    formula?: string
    truncated?: boolean
  }
}

export type OfficeSelectionTrigger = "pointer" | "keyboard" | "touch" | "programmatic"

export interface OfficeReferenceConfirmEvent {
  reference: OfficeObjectReference
  snapshot?: OfficeReferenceSnapshot
  trigger: OfficeSelectionTrigger
  additiveRequested: boolean
}
```

`reference` 负责真正定位；`snapshot` 只是确认时可选的人类可读快照，例如“第 3 页 › 表格 2 › 单元格 B4”。宿主可以丢弃快照、重新调用 `describe()`、把引用加入自己的多选集合，或转换成任意 prompt、tool arguments、MCP resource 和业务协议。`additiveRequested` 只忠实表达 Shift 等输入手势，不要求宿主一定追加目标。

公共包不导出 `OfficeReferenceRole`、`OfficeIntentSelection`、`OfficeSelectedReference` 或 `OfficeIntentContext`。宿主如需角色、引用集合或 Agent 上下文，应在自己的业务类型中组合 `OfficeObjectReference`。

### 5.7 Schema 不变量

`@arcships/office-interaction` 同时提供 TypeScript 类型和运行时校验器。校验器必须保证：

- `reference.document.format`、`reference.locator.format` 和格式定位器分支一致；
- DOCX/PDF 人工区域只能使用 `page`，PPTX 只能使用 `slide`，XLSX 只能使用 `sheet`；
- `referenceId` 必须是非空的不透明引用身份，格式对象原生 id 不能代替它；
- 所有索引和偏移是有限非负整数，所有规范化坐标在 `[0, 1]` 内；
- `score` 如存在必须是 `[0, 1]` 内有限数值；
- 指纹文本遵守长度上限，`facets` 只能包含 `JsonValue`；
- 未知 `schemaVersion` 或不可能的判别联合直接返回结构化校验错误。

kind 与 locator 的兼容表也是运行时 schema 的强制规则：

| 格式 locator | 允许的 kind |
|---|---|
| DOCX `page` / `text-range` / `image` / `comment` | 分别只能是 `page` / `text-range` / `image` / `comment` |
| DOCX `structure` | 由路径末端和核心模型校验为 `section`、`paragraph`、`heading`、`list-item`、`table`、`table-row` 或 `table-cell` |
| DOCX `behavior` | `style→style`、`field→field`、`bookmark→bookmark`、`section→section`、`tracked-change-state→tracked-change`、`comment-state→comment`、`cross-reference→cross-reference`；文档级 style 使用 document scope，其余按真实 owner 使用 document 或带 part/path 的 part scope，且每项必须有唯一 `instanceId` |
| XLSX `worksheet` / `row` / `column` / `table` / `comment` | 分别只能是同名语义 kind |
| XLSX `range` | 单格只能是 `cell`，多格只能是 `cell-range` |
| XLSX `drawing` | `objectType` 必须分别对应 `chart`、`image`、`shape` 或 `form-control` |
| XLSX `chart-element` | `series→chart-series`、`point→chart-point`、`legend-entry→chart-legend-entry` |
| XLSX `behavior` | `formula→formula`、`conditional-format→conditional-format`、`data-validation→data-validation`、`named-range→named-range`、`protection→protection`；前三者必须使用带 `appliesToA1` 的 sheet scope，命名区域和保护可使用 workbook 或 sheet scope，且每项必须有唯一 `instanceId` |
| PPTX `slide` / `text-range` | 分别只能是 `slide` / `text-range` |
| PPTX `object` | kind 必须与当前修订中 `objectKey` 的解析类型相符，只允许 `text-box`、`shape`、`image`、`table`、`chart` 或 `group` |
| PPTX `sub-object` | 路径末端分别对应 `table-cell`、`chart-series`、`chart-point` 或 `chart-legend-entry` |
| PPTX `behavior` | 分别对应 `animation`、`transition`、`action`、`master` 或 `note` |
| PDF `page` / `text-range` | 分别只能是 `page` / `text-range` |
| PDF `native-object` | `link→link`、`form→form-control`、`annotation→annotation` |
| PDF `visual-object` | provider 只能声明 `text-block`、`image`、`table`、`table-row` 或 `table-cell`，并保留 `source: "visual"` |
| PDF `behavior` | `link-action→action`、`form-rule→form-rule`、`annotation-state→annotation`，其余分别对应 `attachment`、`layer`、`signature`、`permission` |
| 任意 `manual-region` | kind 只能是 `region`，并遵守格式与 space 对应关系 |

`schemaVersion: 1` 的 kind、locator 和 behavior 判别值是封闭集合。增加可选证据字段可以保持 v1；增加新的对象或定位器分支必须发布新的 schema version 和显式升级函数，不能让旧校验器把新语义当作 v1 接受。

序列化器输出确定的字段顺序，便于快照测试和内容哈希；消费者不能依赖对象属性原始插入顺序。

## 6. 引用解析与失效

统一解析接口：

```ts
export type ResolveReferenceResult =
  | { status: "exact"; reference: OfficeObjectReference; descriptor: OfficeObjectDescriptor }
  | { status: "relocated"; reference: OfficeObjectReference; descriptor: OfficeObjectDescriptor; reasonCodes: readonly string[] }
  | { status: "ambiguous"; candidates: readonly OfficeObjectDescriptor[]; reasonCodes: readonly string[] }
  | { status: "not-found"; reasonCode: string }
  | { status: "unsupported"; reasonCode: string }

export interface OfficeReferenceResolver {
  resolve(
    reference: OfficeObjectReference,
    current: OfficeDocumentRevision,
    options?: { signal?: AbortSignal },
  ): Promise<ResolveReferenceResult>
}
```

`exact` 和 `relocated` 返回的 reference 必须绑定当前修订，并保留原 `referenceId`；descriptor 只描述当前对象。resolver 不直接修改宿主保存的引用，宿主收到结果后自行决定是否替换旧 reference。

解析顺序固定为：

1. 校验格式和逻辑文档身份；
2. 同一修订优先验证原生定位器；
3. 修订变化后使用原生 id、祖先路径和内容指纹寻找候选；
4. 用文字、名称、类型、层级和邻近区域交叉验证；
5. 唯一且达到格式策略阈值时返回 `relocated`；
6. 多个候选接近时返回 `ambiguous`，不得自动取第一项；
7. 无法确认时返回 `not-found`；是否保留旧快照、如何提示失效由宿主决定。

各格式可以使用不同的内部阈值，但必须测试“最低可接受分数”和“第一、第二候选的最小差距”。这些排序分数不写入公共 `reliability.score`，除非对应识别器已完成校准。

人工区域的跨修订规则更保守：PDF/PPTX 页面或幻灯片尺寸与顺序未变时可以映射；DOCX 重排版、XLSX 行列插删或页面顺序变化后，只有同时存在附近结构证据时才自动恢复，否则要求重新选择。

## 7. 对象索引与命中测试

### 7.1 适配器接口

每种 Surface 实现同一内部接口：

```ts
export interface OfficeSurfaceObjectAdapter {
  getRevision(): OfficeDocumentRevision
  hitTest(
    point: { clientX: number; clientY: number },
    options?: { kinds?: readonly OfficeObjectKind[]; maxCandidates?: number },
  ): OfficeHitCandidate[]
  describe(reference: OfficeObjectReference, signal?: AbortSignal): Promise<OfficeObjectDescriptor>
  resolve(reference: OfficeObjectReference, signal?: AbortSignal): Promise<ResolveReferenceResult>
  getAncestors(reference: OfficeObjectReference): Promise<OfficeObjectDescriptor[]>
  getChildren(reference: OfficeObjectReference): Promise<OfficeObjectDescriptor[]>
  scrollTo(reference: OfficeObjectReference): Promise<void>
}
```

格式核心层建立“语义对象索引”，Surface 建立“当前可见几何索引”。两者通过格式定位器关联。不能在每次 `pointermove` 时遍历全文 DOM 或整个工作簿。

### 7.2 命中管线

一次指针命中按以下顺序执行：

1. 从 client 坐标转换到页面、幻灯片或工作表规范坐标；
2. 只查询当前可见页面、幻灯片或网格窗口的几何索引；
3. 收集原生 DOM 标记、核心模型边界、结构推断和视觉识别候选；
4. 根据调用方给出的 kind 范围过滤；适配器本身不知道宿主正在展示哪种工具栏或工作流；
5. 按直接命中、可见性、深度、面积、Z 轴和来源证据稳定排序；
6. 生成候选链，但不在适配器中替用户作最终选择；
7. 没有语义候选时允许进入人工区域框选。

排序规则必须确定且可测试。不能仅依赖浏览器 `event.target`，也不能把“原生来源”简单等同于“用户一定想选它”。例如点击图表标题时，图表标题和整个图表都是原生对象，最终选择由模式、深度和用户的穿透操作共同决定。

悬停命中使用 `requestAnimationFrame` 合帧；同一位置最多返回 20 个候选。更多对象由宿主通过父子查询或格式索引按需取得，库不规定它们必须出现在图层面板中。对象索引按 `revision + container + layoutVersion` 缓存，布局、缩放或模型变化后精确失效。

## 8. 无头选择会话与交互语义

选择会话只保存一次“指向并确认”的临时状态。确认事件发出后，选择器不继续持有引用集合，也不记录目标、参考物、来源或插入位置等业务角色。

```ts
export type OfficeSelectionMode = "content" | "object" | "region"

export interface OfficeSelectionSessionState {
  mode: OfficeSelectionMode
  phase: "idle" | "pointing" | "choosing" | "drawing"
  candidates: readonly OfficeReferenceCandidatePreview[]
  activeCandidateId?: string
  regionDraft?: ManualRegionLocator
}
```

公共 reducer 只接受以下类别的 action：设置受控模式、替换候选、移动活动候选、开始/更新区域草稿和取消当前操作。纯函数 `confirmOfficeCandidate(state, input)` 接收明确的 `referenceId`、触发来源和修饰键，返回 `{ state, event: OfficeReferenceConfirmEvent }`；它不会把引用追加到 session。是否多选、如何排序、能否删除以及给引用分配什么业务角色，都由宿主自己的状态管理决定。

交互规则：

| 模式 | 默认指针行为 | 与现有能力的关系 |
|---|---|---|
| `content` | DOCX/PDF 拖文字、XLSX 选范围、PPTX 选幻灯片或文字 | 默认模式，保持现有手势 |
| `object` | 悬停对象轮廓，点击提交单个可见对象 | 只有显式进入后才覆盖文字拖选 |
| `region` | 在页面、幻灯片或工作表上框选区域 | 所有语义识别失败时的目标兜底 |

`layers` 不是选择模式，而是宿主对候选链和对象父子关系的一种可选呈现；列表、面包屑、2D/3D 图层视图都消费同一候选数据。`behavior` 也不是画布模式：公式、动画和设置由格式适配器从属性面板、时间线或结构树等入口转换成同一种 candidate/reference，宿主决定采用何种界面。

统一快捷行为：

- `Escape` 先取消候选或内部层级，再发出取消事件；是否切换模式由宿主决定；
- `Enter` 确认当前候选；
- `Tab` / `Shift+Tab` 在可见候选间移动；
- 上下方向键移动同级，左右方向键进入子级或返回父级；
- `Alt/Option + 点击` 穿透到下一个重叠候选；
- `Shift + 点击` 只在确认事件中设置 `additiveRequested: true`；是否追加由宿主决定；
- 触控长按打开同点候选，拖动手柄调整区域边界。

可选 Vue 原语必须暴露焦点与 `aria-live` 所需状态，使宿主能够构建可访问的候选界面。库不强制渲染候选列表；如果宿主选择渲染列表、面包屑或 3D 图层透镜，它们必须消费相同的对象树和候选状态，不能建立第二套对象身份。

## 9. Surface 公共 API

四种 Surface 在保留现有格式事件的同时，只增量暴露选择能力和原始事件，不保存宿主的引用集合：

```ts
interface OfficeReferenceSurfaceProps {
  documentId?: string
  selectionMode?: OfficeSelectionMode // 完全受控；默认 content
  emitReferenceCandidates?: boolean
}

interface OfficeReferenceSurfaceExposed {
  getDocumentRevision(): OfficeDocumentRevision
  hitTest(point: { clientX: number; clientY: number }): readonly OfficeReferenceCandidatePreview[]
  describeReference(reference: OfficeObjectReference, signal?: AbortSignal): Promise<OfficeObjectDescriptor>
  resolveReference(reference: OfficeObjectReference): Promise<ResolveReferenceResult>
  scrollToReference(reference: OfficeObjectReference): Promise<void>
  captureReferencePreview(
    reference: OfficeObjectReference,
    options?: { maxWidth?: number; maxHeight?: number; signal?: AbortSignal },
  ): Promise<Blob>
}

export interface OfficeReferenceCandidatePreview {
  candidateId: string
  draft: OfficeObjectReferenceDraft
  preview?: Pick<OfficeObjectDescriptor, "label" | "path" | "visual">
  hit: OfficeHitCandidate["hit"]
  depth: number
  zIndex?: number
}

export type OfficeReferenceOperation = "hit-test" | "describe" | "resolve" | "scroll" | "capture" | "provider"
export type OfficeReferenceErrorCode =
  | "INVALID_REFERENCE" | "REVISION_CHANGED" | "HIT_TEST_FAILED" | "DESCRIBE_FAILED"
  | "RESOLVE_FAILED" | "CAPTURE_UNSUPPORTED" | "CAPTURE_LIMIT_EXCEEDED"
  | "PROVIDER_FAILED" | "ABORTED"

export interface OfficeReferenceError {
  code: OfficeReferenceErrorCode
  operation: OfficeReferenceOperation
  format: OfficeFormat
  recoverable: boolean
  referenceId?: string
  message: string
}

export interface OfficeReferenceCandidateChange {
  candidates: readonly OfficeReferenceCandidatePreview[]
  activeCandidateId?: string
}

export interface OfficeRegionDraftChange {
  phase: "start" | "change"
  region: ManualRegionLocator
}

export interface OfficeSelectionCancelEvent {
  mode: OfficeSelectionMode
  reason: "escape" | "pointer-cancel" | "programmatic"
}

export interface OfficeReferenceResolveEvent {
  referenceId: string
  result: ResolveReferenceResult
}
```

`selectionMode` 只决定 Surface 如何解释接下来的指针手势，不包含已确认引用。Surface 不提供 `defaultReferenceSelection`、`getReferenceSelection()` 或内部引用托盘；宿主可以使用公共 reducer，也可以完全用自己的状态管理。候选事件的关闭、确认事件后的动作和模式切换都不由 Surface 猜测。

统一事件：

| 事件 | 触发条件 | 载荷 |
|---|---|---|
| `documentRevisionChange` | 初始加载或模型命令提交产生新修订 | `OfficeDocumentRevision` |
| `referenceCandidateChange` | 悬停或穿透选择导致候选链变化 | `OfficeReferenceCandidateChange`；只有 `emitReferenceCandidates` 为 true 时发出 |
| `referenceConfirm` | 用户确认文字、对象或人工区域 | `OfficeReferenceConfirmEvent` |
| `regionDraftChange` | 人工区域开始或变化 | `OfficeRegionDraftChange` |
| `selectionCancel` | 用户取消当前候选或区域操作 | `OfficeSelectionCancelEvent` |
| `referenceResolve` | 旧引用在新修订中得到结果 | `OfficeReferenceResolveEvent` |
| `referenceError` | 命中、描述、截图或解析过程失败 | `OfficeReferenceError`，不含文档正文 |

兼容策略：

- `selectionChange` 继续表达当前格式已有的原始选区；
- `objectClick` 在阶段二以前保留，阶段二完成后标记为被 `referenceConfirm` 取代；
- `0.x` 不删除旧事件，最早在 `1.0.0` 按迁移指南移除；
- 新 API 只从各包根入口导出，不要求消费端深层导入；
- 公开事件只含引用和序列化数据，不泄漏格式内部模型实例。

## 10. 各格式接入设计

### 10.1 DOCX

- 阶段一：把当前只有 `text + pageIndex` 的选区提升为 `DocxTextRange + text quote + rects`；补页面与人工区域定位器；
- 阶段二：从 `DocModel`、段落样式和现有 DOM 标记建立段落、标题、列表项、表格、图片，以及批注标记与批注内容索引；
- 阶段三：复用 table/row/cell/paragraph 路径和嵌套表格结构，建立父子关系；
- 阶段四：由核心模型增加样式、字段、书签、分节、修订、批注状态和交叉引用描述器；宿主自行决定通过何种界面展示这些不可见对象。

正文、页眉、页脚、脚注和尾注必须带 part/region，不能只用全局 `nodeIndex`。分页变化不会改变段落身份，页面索引只能作为显示几何和辅助证据。

当前 `@arcships/docx-core` 根入口已实现正文 `DocxTextRange`、页面和人工区域草稿，以及同修订验证与跨修订文字重定位；`DocxDocumentSurface` 已把原生文字选择、页面对象模式和区域模式转换为统一确认事件。正文定位器使用现有段落/表格/行/单元格路径；唯一文字匹配才自动迁移，删除与并列候选分别返回 `not-found` 和 `ambiguous`。页眉、页脚、脚注、尾注及阶段二可见对象仍属于后续增量；DOCX 页码与人工区域在修订变化后因可能重排版而要求重新选择。

### 10.2 XLSX

- 阶段一：直接复用 `XlsxCellRange`，增加工作表、整行、整列和带单元格内偏移的人工区域；
- 阶段二：把表格、控制器内部的图表/图片选择和批注标记与内容接入 Surface，并用 table name、drawing id、chart id 和批注所属单元格建立定位器；
- 阶段三：复用现有 chart series/point/legend DOM 标记和模型，支持图表内部和重叠对象候选；
- 阶段四：将公式、条件格式、数据验证、命名区域和保护规则作为作用于范围的行为对象。

工作表索引不是稳定身份。定位时优先 sheet id，再用名称和索引校验；表格优先使用名称，单元格和范围使用 A1 语义而不是像素坐标。

当前 `@arcships/xlsx-core` 根入口已实现工作表、单元格、A1 范围、整行、整列、图表和人工 sheet 区域草稿，以及同步解析与描述；`XlsxSheetSurface` 已把普通单元格/范围选择、表头行列/工作表命中和区域模式转换为统一确认事件。工作表重命名或重排时优先按 tab id 迁移；图表 drawing id 变化后，仅在名称与图表类型形成唯一候选时迁移，重复候选返回 `ambiguous`。表格、图片、形状和图表内部元素仍按阶段二、三矩阵后续补齐。

### 10.3 PPTX

- 阶段一：沿用幻灯片选择，新增文本范围和人工区域；
- 阶段二：将现有 `objectKey` 映射为文本框、形状、图片、表格和图表语义；
- 阶段三：利用 `groupPath`、段落标记和渲染树增加组合、单元格与图表内部层级；
- 阶段四：直接复用 playback 文档已有的动画目标、transition、action 和 master 来源，并补备注对象描述器；时间线或结构树由宿主实现并把所选描述器交给同一确认函数。

`objectKey` 是同一解析结果内的主键，跨修订时仍要结合 slide id、shape id、source、group path、名称和内容摘要校验。

当前 `@arcships/pptx-core` 根入口已实现幻灯片、对象内精确文本范围、形状/图片/表格/图表、嵌套组合对象和人工 slide 区域草稿；`PptxStage` 已把原生文字选择、渲染对象命中和区域模式转换为统一确认事件。解析时先按 slide part id 和 `objectKey`，再按 shape id、source、group path 校验；原生 id 变化后只有名称与节点类型唯一时迁移，精确文字只有唯一 quote 候选时迁移。表格单元格、图表内部元素和动画等不可见对象仍属于阶段三、四增量。

### 10.4 PDF

- 阶段一：复用 `PageTextSlice` 和字符几何生成精确文字引用；新增页面与规范化人工区域；
- 阶段二：优先暴露 PDF 原生链接、表单和批注；图片、文字块和表格通过可插拔视觉识别器生成候选；
- 阶段三：视觉识别器可补充版面块、表格行列和重叠标记层级；
- 阶段四：从 PDF 原生结构暴露链接动作、表单规则、批注状态、附件、可选内容组、签名和权限。

视觉识别是可选 provider，不作为 `vue-pdf` 默认包体的一部分：

```ts
export interface PdfVisualObjectDetection {
  objectId: string
  kind: "text-block" | "image" | "table" | "table-row" | "table-cell"
  region: NormalizedRect
  label?: string
  parentObjectId?: string
  content?: { text?: string }
  reliability: OfficeObjectReliability
}

export interface PdfVisualObjectProvider {
  readonly id: string
  recognize(input: {
    revision: OfficeDocumentRevision
    pageIndex: number
    image: ImageBitmap
    signal: AbortSignal
  }): Promise<readonly PdfVisualObjectDetection[]>
}
```

provider 只返回检测结果，不生成 `referenceId`；PDF 适配器把 `provider.id + objectId + region` 转换为 visual locator 和 reference draft。没有 provider 或识别失败时，文字与人工区域仍必须可用。远程 provider 的上传和隐私授权由宿主负责，组件库默认只接受本地 provider。

当前 `@arcships/vue-pdf` 根入口已实现 `PageTextSlice`、页面和人工区域草稿；`PdfSurface` 已把精确文字选择、页面对象模式和区域模式转换为统一确认事件。文字解析复用调用方传入的 `PdfRenderRuntime.getTextSlices()` 与 `search()`；页面和区域跨修订只有在页面尺寸与旋转签名不变时迁移。

定位器联合覆盖的是目标架构可表达范围，不代表对应阶段已经实现。每阶段的交付对象只以产品矩阵和本节阶段条目为准；例如 XLSX shape/form-control 类型是后续兼容扩展位，不属于首批阶段二验收。

## 11. 性能、并发与安全

- 对象索引按页面、幻灯片或当前工作表窗口懒加载；不因打开对象模式扫描全文；
- `pointermove` 只做同步几何查询，OCR、文字提取和复杂描述在确认候选后异步执行；
- 每个异步描述、截图和解析任务接受 `AbortSignal`，文件切换时取消，旧任务不得回写新修订；
- 引用解析使用 `documentId + revision + requestId` 校验最新结果；
- 页面预览和截图遵守现有图片像素、内存和任务超时预算；
- 序列化器拒绝函数、DOM、循环引用、Blob URL 和未知 schema 版本；
- 默认遥测只记录格式、kind、来源、耗时、候选数和结果状态，不记录正文、公式、批注或截图；
- `captureReferencePreview()` 返回的 Blob 由调用方释放和传输，Surface 不长期缓存未被引用的位图。

## 12. 分阶段实施与验收

阶段完成仍以产品文档第 4 节矩阵为准，不能只完成一个格式就宣布整个阶段结束。

### 阶段一：基础引用闭环

实施内容：

1. 建立 `@arcships/office-interaction`、schema 校验、选择 reducer 和解析结果类型；
2. 四种 Surface 接入 `documentId`、`revision`、受控 `selectionMode` 和统一选择事件；
3. DOCX/PDF 精确文字范围，XLSX 单元格/范围/行列，PPTX 幻灯片与文字范围；
4. 四种格式的页面、幻灯片、工作表和人工区域选择；
5. `vue-ui` 提供受控轮廓层和区域框两类选择原语；
6. demo 展示“选择 → 查看引用事件 JSON → 模拟文件更新 → 重新定位/失效”，不内置引用托盘或真实 Agent 服务。

验收门槛：

- 产品文档阶段一矩阵全部覆盖；
- 缩放、滚动和旋转后引用不变，轮廓仍与内容对齐；
- 新修订中的引用只会得到 `exact`、`relocated`、`ambiguous`、`not-found` 或 `unsupported`；
- 区域选择在无法识别对象的页面上始终可用；
- 宿主无需读取 DOM 即可从 `referenceConfirm` 得到稳定引用。

### 阶段二：单个可见对象

- 四种格式实现产品矩阵中的高频可见对象；
- 对象模式提供候选描述、悬停轮廓和稳定单击确认；宿主决定是否显示类型标签；
- 建立格式对象索引，不要求内部层级；
- 完成旧 `objectClick` 到统一引用事件的兼容适配。

### 阶段三：层级与重叠

- 增加父子查询和候选轮换，不规定面包屑、图层面板或其他呈现；
- 覆盖表格单元格、图表内部元素、组合、遮挡和 PDF 版面层级；
- 对象树和选择 reducer 稳定后，宿主可以用列表、面包屑或 3D 图层透镜消费同一状态。

### 阶段四：行为与设置

- 增加行为与设置的 descriptor/reference 适配器；
- 行为引用同时包含行为身份和作用对象/范围；
- 完成产品矩阵列出的公式、规则、动画、翻页、动作和其他不可见对象；宿主负责时间线、属性面板或结构树界面。

## 13. 测试与发布门禁

| 层级 | 必测内容 |
|---|---|
| 纯单元测试 | schema 解析、坐标归一化、选择 reducer、候选稳定排序、引用解析状态机 |
| 格式夹具测试 | 每类定位器、父子关系、同修订精确解析、插入/删除/重排后的重定位和歧义 |
| Vue 组件测试 | 鼠标、键盘、触控、缩放、滚动、受控模式/区域草稿、旧事件兼容和取消 |
| 浏览器黑盒 | 四种真实文件、轮廓截图、区域框选、文件更新后的恢复/失效、无控制台异常 |
| 可访问性 | 焦点顺序、候选播报、键盘进入/退出层级、颜色之外的可靠度提示 |
| 发布包消费 | 九个目标公开包的 JS、类型、CSS、子入口和外部项目安装；公开 `.d.ts` 不引用私有包 |

每个解析器至少要有三类变更夹具：对象未变但位置变化、对象被删除、出现两个同样可信的候选。只有第一类可以自动返回 `relocated`；后两类必须分别得到 `not-found` 和 `ambiguous`。

## 14. 第一批工程任务边界

具体开发从阶段一拆成以下可独立合并的任务：

| 顺序 | 任务 | 产物 |
|---|---|---|
| T1 | 公共合同与发布骨架 | 新包、类型、schema、纯函数、压缩包消费测试 |
| T2 | 无头选择会话与基础原语 | 临时 reducer、轮廓层、区域框、候选导航和键盘基础 |
| T2R | 核心边界收缩 | 确认事件合同；删除未发布的意图角色、引用托盘和核心业务动作 |
| T3 | DOCX 与 PDF 基础适配 | 精确文本引用、页面区域、解析与失效夹具 |
| T4 | XLSX 与 PPTX 首批格式适配 | 工作表/单元格/范围/图表/区域，以及幻灯片/对象/组合/区域引用与夹具 |
| T5 | 统一 Surface API 与 demo | 受控模式、候选/确认事件、expose、引用查看和重载验证 |
| T6 | 四格式阶段一正式验收 | 组件、黑盒、可访问性、性能和九包消费证据 |

T1 已完成公共包、v1 引用、运行时校验、稳定序列化、几何、候选排序和发布消费边界。T2 已完成候选导航、键盘解释器、对象轮廓层和区域框。T2R 已删除未发布的意图协议、引用集合 reducer、业务动作、引用标签和引用托盘，并新增确认事件、取消事件与短生命周期 session。T3 已完成 DOCX/PDF 精确文字、页面/区域草稿。T4 已完成 XLSX 与 PPTX 首批格式适配。T5 已完成四种 Surface 的受控模式、统一事件和 expose，并在 DOCX demo 展示引用 JSON、重载与重新解析。T6 已补齐 XLSX 行列和 PPTX 精确文字，完成四格式组件/黑盒/可访问性/性能与九包外部消费验收。正式发布证据位于 `output/acceptance/de7b2ba31fd538645627abd176ddb7d145d2f80e/BB-RELEASE/20260714T045733Z/release/`。

## 15. 与现有设计的关系

- [统一 Surface API 设计](../unified-surface-api-design.md)继续负责渲染面、缩放、滚动和现有事件；对象选择部分以后以本文为准；
- [Surface 搜索与 PDF 选择设计](../surface-search-and-pdf-selection-design.md)继续负责搜索会话和 PDF glyph 算法；本文把其结果转换为持久引用；
- [项目架构审查与目标设计](../architecture-review-and-target-design.md)继续负责 Runtime、加载、包体和发布原则；本文新增公开基础包是一次明确的目标架构扩展；
- 产品对象范围和阶段完成口径以[《Office 对象语义与意图选择》](./object-semantics-and-selection.md)为准；其中描述的菜单、图层、3D 视图和 Agent 闭环是宿主可构建的产品体验，不自动成为本库的成品组件承诺。
