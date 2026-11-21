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

浏览模式按可视区域挂载页面，滚动时同步当前页、工具栏页码和缩略图状态。浏览模式的 `PageUp`/`PageDown`、`↑`/`↓`、`Home`/`End` 和 `Space` 只滚动幻灯片列表，不会切换当前页；`PageUp`/`PageDown`、`Home`/`End`、方向键和 `Space` 的翻页/播放导航只在演示模式生效。两种模式都支持 `Ctrl/Cmd` + `+`/`=`、`-`、`0` 缩放，浏览模式还支持 `Ctrl/Cmd + F` 搜索；`Escape` 在浏览模式关闭搜索，在演示模式退出全屏。完整映射见[查看器键盘导航](viewer-keyboard.md)。

常用属性：

| 属性 | 默认值 | 说明 |
|---|---|---|
| `source` | `null` | `ArrayBuffer`、`Uint8Array` 或文件对象 |
| `mode` | `browse` | 浏览或演示；支持 `v-model:mode`，省略时由工具栏切换 |
| `height` | `100%` | 组件高度 |
| `minHeight` | 无 | 可选最小高度（数字按 px）；不传时组件不设置最小高度 |
| `initialSlide` | `0` | 初始页，从零开始 |
| `autoplay` | `true` | 演示模式自动执行非点击步骤 |
| `showHiddenSlides` | `false` | 播放时包含隐藏页 |
| `showNotes` | `false` | 显示演讲者备注面板；演示模式下变为演讲者视图（当前页备注 + 下一页缩略图 + 计时）。支持 `v-model:show-notes`，工具栏 `⋯` 菜单的备注项也会发 `update:showNotes` |
| `showNotesToggle` | `true` | 工具栏 `⋯` 菜单的备注切换项 |
| `showPresentToggle` | `true` | 工具栏 `⋯` 菜单的“演示”项；演示模式在播放控制条显示“退出演示” |
| `showFullscreenControl` | `true` | 工具栏 `⋯` 菜单的全屏项 |
| `showToolbar` | `true` | 浏览工具栏（`PptxViewerToolbar`，48px min-height + 1px 底边） |
| `viewMode` | `null` | 浏览布局：`single` 或 `sorter`；`null` 跟随内部菜单状态 |
| `defaultViewMode` | `single` | 未受控时的初始浏览布局 |
| `showSidebar` | `true` | 缩略图侧栏（工具栏首项可切换） |
| `showSearch` | `true` | 搜索入口：工具栏 `⋯` 菜单的“搜索幻灯片”，或 `Ctrl/Cmd+F`；打开后工具栏下方出现 find bar |
| `showPlaybackControls` | `true` | 播放控制条 |
| `showWarnings` | `true` | 显示文档 `warnings`（备注不支持、宏已忽略等） |
| `showContextMenu` | `true` | 内置链接右键菜单；设为 `false` 时只发 `context-menu` |
| `linkHandling` | `built-in` | `built-in` 由查看器按 URL 策略打开/复制链接；`host` 只发 `action`/`context-menu`，由宿主处理 |
| `openTimeoutMs` | `30000` | 加载超时毫秒数，超时后进入错误状态；`0` 关闭超时 |
| `externalMedia` | `disabled` | 是否允许外部媒体；设为 `allowed` 时仍需要 `externalMediaUrlPolicy` 放行 |
| `externalMediaUrlPolicy` | `undefined` | 外部媒体的 URL 规则：`baseUrl`、`allowedProtocols`、`allowedOrigins` 等 |

外部媒体默认禁用，只作为静态内容报告。只有同时设置 `externalMedia="allowed"` 和 `externalMediaUrlPolicy`，并且目标 URL 通过 `allowedProtocols` / `allowedOrigins` 校验时才真正加载；`javascript:`、`data:` 和未授权来源一律不加载，避免演示文稿自行发起跟踪请求。打印时 `@page` 尺寸取自幻灯片尺寸，工具栏、侧栏和播放控制条自动隐藏。

`limits` 支持完整的归档预算：`maxInputBytes`、`maxArchiveEntries`、`maxSingleEntryBytes`、`maxUncompressedBytes`、`maxMediaBytes`、`maxCompressionRatio`、`maxArchivePathLength`、`maxXmlBytes`、`maxSingleXmlBytes`、`maxXmlDepth`、`maxXmlAttributeBytes`、`maxTextNodeBytes`、`maxRelationships` 和 `maxConcurrency`。宿主只能收紧默认值，不能放宽硬上限。

模板引用公开 `getController()`、`next()`、`previous()`、`play()`、`pause()`、`resume()`、`reset()`、`goToSlide()`、`enterFullscreen()` 和 `exitFullscreen()`。`enterFullscreen()` 只投影舞台（`data-testid="pptx-audience-surface"`）：工具栏、侧栏、播放控制条和演讲者备注 rail 都不进入全屏；备注投到第二块屏需要宿主自建窗口（#202）。

### 超链接与右键菜单

`a:hlinkClick` 和 `a:hlinkHover` 在浏览与演示模式下都可用。形状点击先经过共享 URL 策略：只允许 `http:`、`https:`、`mailto:` 和文档内跳页，`javascript:`、`data:`、`vbscript:` 等目标一律拒绝（不会被当作可点击链接，也不会到达 `window.open`）。外部链接在新标签页打开（`noopener,noreferrer`），内部跳转调用同一个 `goTo()`，因此浏览模式会滚到目标页、演示模式沿用播放控制器的跳页规则。两种模式都继续抛出 `action` 事件（`open-url` / `mailto` / `go-to-slide` / `unsupported`），宿主可以照旧接管。`linkHandling="host"` 关闭内置打开/复制行为，只保留事件。

右键带链接的形状会显示共享 `OfficeDropdownMenu` 渲染的内置菜单（内容节点 `data-testid="pptx-context-menu"`，`role="menu"`）：外部链接提供“打开链接”（`pptx-context-menu-open-link`）和“复制链接”（`pptx-context-menu-copy-link`），内部跳转提供“跳到第 N 页”。菜单以右击点为零尺寸锚点，由 Reka 处理视口翻转；键盘语义来自共享 primitive（`↑`/`↓`/`Home`/`End` 移动焦点，typeahead，`Esc`/`Tab` 或外部点击关闭），关闭后焦点回到查看器根节点。文档、页码或模式变化时菜单自动收起。无论是否显示内置菜单，`context-menu` 事件都会先抛出，宿主可用 `showContextMenu="false"` 完全接管；自建舞台可复用导出的 `PptxLinkMenu`。

渲染完成后，标记层会把链接信息写到形状元素上：`title` 为 `tooltip` 属性（没有时回退到 URL），并设置 `data-pptx-hyperlink`（`click`/`hover`）、`data-pptx-hyperlink-kind`、`data-pptx-hyperlink-url`、`data-pptx-hyperlink-slide` 和 `data-pptx-hyperlink-tooltip`。点击链接的形状获得 `role="link"` 和 `tabindex="0"`，`Enter`/`Space` 与鼠标点击走同一条链接路由；缩略图里的链接只保留 tooltip，不会在缩略图按钮里再嵌套可聚焦元素。悬停链接（`a:hlinkHover`）只显示 tooltip，不会自动执行；播放能力报告把悬停动作记为 `static`。同一份链接元数据也会出现在 `PptxObjectIdentity.hyperlinks`（`PptxObjectHyperlink[]`）和 `PptxSlideAction.tooltip` 中，供无障碍、引用适配器和自建界面使用。

### 幻灯片浏览（Sorter）

浏览模式支持两种布局，工具栏右侧 `⋯` 菜单（`data-testid="pptx-actions"`）里的“单页视图”/“幻灯片浏览”（`data-testid="pptx-view-mode-single"`、`pptx-view-mode-sorter`）切换，也可以由 `viewMode` 属性控制：

- `single`（默认）：保持原来的纵向幻灯片列表。
- `sorter`：网格缩略图，每页一张 `PptxSlideCard`（上游缩略图卡片的 `p-2` / `gap-2` / `rounded-md` / `shadow-xs` 结构），与侧栏共用同一套 `PptxThumbnail` 与懒加载门控；缩略图加载时显示 `OfficeSkeleton`，渲染失败显示错误 badge，隐藏页显示 `OfficeBadge`。网格只挂载可视行附近的缩略图（上方和下方有等高占位块），滚动与 `PageUp`/`PageDown`/`Home`/`End` 只滚动网格，不切换当前页；点击缩略图会切到该页但网格保持打开，选中项保持高亮。网格打开时隐藏侧栏，舞台仍然挂载（会话仍渲染到它）但在网格下方。移动端 390px 下 2 列，卡片不横向溢出，网格可滚动。

两种布局共用同一个活动页（`data-slide-index`），切换布局时保持当前页与选中项；`sorter` 打开时会把当前页滚到可视区域。演示模式始终使用单页舞台，不显示布局按钮。事件新增 `update:viewMode`，网格根节点为 `data-testid="pptx-sorter"`，每个格子为 `data-testid="pptx-sorter-item"` 并带 `data-slide-index` 与 `data-active`。

### 拖放导入

把 `.pptx`、`.pptm`、`.ppsx`、`.ppsm`、`.potx` 或 `.potm` 拖到查看器上会直接打开，走与 `source` 属性相同的 `usePptxDocument` 会话管线；宿主更新 `source` 后，属性来源重新优先。`PptxViewer` 同时抛出 `fileDrop(file)` 事件。拖入时根元素出现虚线轮廓，拖离、放下或拖拽在窗口其他位置结束后清除；其他扩展名不会打开，并在查看器上方显示本地化的拒绝提示（`pptx.viewer.unsupportedDrop`）。`isSupportedPptxFileName()` 与 `PPTX_FILE_ACCEPT` 可用于自建上传入口。

### 演讲者备注与演讲者视图

加载时会读取每页的备注部件（`ppt/notesSlides/notesSlide*.xml`），并按段落换行写入 `PptxPreviewDocument.slides[i].notes`。没有备注部件的页面不带该字段，空备注也不会产生空字符串。幻灯片图像、页码、页眉/页脚和日期占位符不算备注内容；备注只解析一次，且只解压备注部件本身，幻灯片 XML 和媒体不会再次解码。

- `showNotes`（默认 `false`）打开备注面板：浏览模式显示当前页备注，演示模式显示完整演讲者视图；工具栏 `⋯` 菜单的备注项也可切换。两种模式共用同一个面板根节点 `data-testid="pptx-presenter"`，由 `data-presenter-mode="browse"|"present"` 区分，内部是 `OfficeCard` + `OfficeSeparator` + `OfficeScrollArea`。
- 演讲者视图包含当前页备注（`pptx-presenter-notes`）、下一页缩略图（`showHiddenSlides=false` 时跳过隐藏页；最后一页显示 `pptx-presenter-last`）和从开始演示算起的计时（`pptx-presenter-timer`，`mm:ss`，超过一小时为 `h:mm:ss`）。
- 备注面板可用 `notes` 插槽整体替换，插槽参数为 `{ notes, slide }`。没有备注的页面显示本地化的空状态。
- 备注部件损坏时只影响该页备注，不会阻止演示文稿打开；但超过 XML 预算或加载被取消仍会使打开失败。

```vue
<PptxViewer :source="file" mode="present" :show-notes="true" height="100vh" />
```

## 界面状态

`data-state` 暴露当前视图状态：`empty`、`waiting`（舞台尚未挂载）、`loading`、`ready`、`cancelled` 和 `error`。状态面板由共享 `OfficeEmptyState` / `OfficeLoadingState` / `OfficeErrorState` 渲染；`empty`/`waiting`/`loading`/`cancelled`/`error` 插槽仍可替换描述内容。

- 加载失败、超时或取消才会进入 `error`/`cancelled`，并显示重试按钮；
- 密码保护或旧版 OLE/CFB 演示文稿会返回 `UNSUPPORTED_CONTENT`，查看器显示锁图标的“受保护”状态（`data-testid="pptx-encrypted"`，根节点 `data-encrypted="true"`）；`data-state` 仍为 `error`；
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
  @link-activate="onLinkActivate"
/>
```

| 事件 | 参数 | 触发时机 |
|---|---|---|
| `selection-change` | `{ kind: "slide", slideIndex }` | 点击幻灯片页面或其中对象 |
| `object-click` | `{ kind: "object", slideIndex, objectKey }` | 点击带稳定对象标识的形状、图片、图表等对象 |
| `context-menu` | `PptxStageContextMenu` 判别联合 | 右键幻灯片或对象 |
| `link-activate` | `{ slideIndex, objectKey }` | 用 `Enter`/`Space` 激活已聚焦的链接形状 |

`context-menu` 的精确形式为：

```ts
type PptxStageContextMenu = Position & (
  | { kind: "slide"; objectKey?: never }
  | { kind: "object"; objectKey: string }
)
```

其中 `Position` 包含 `slideIndex`、`clientX/Y` 和 `containerX/Y`；容器坐标相对 `PptxStage` 根元素，而不是宿主的外层滚动容器。`slideIndex` 从零开始。对象点击会先产生页面选中，再产生 `object-click`；右键只发 `context-menu`，其中已经包含页面和对象上下文。宿主决定是否显示菜单，不应解析 `objectKey` 的内部字符串格式。`PptxViewer` 会把同一个 `context-menu` 再抛给宿主，并另外提供内置链接菜单（见[超链接与右键菜单](#超链接与右键菜单)）。

## 兼容性说明

无法精确执行的 PowerPoint 内容会报告为近似、静态或未解析。本库不会用静态显示冒充动画完全兼容。程序可以监听 `capability`、`playback-warning` 和 `playback-error`。

宏格式加载后会通过文档 `warnings` 报告宏已忽略。演讲者备注已经解析，不再产生 `NOTES_UNSUPPORTED` 警告（该代码保留在类型中，方便既有消费方继续编译）。`PptxViewer` 默认在工具栏下方显示这些警告；自定义界面可读取 `usePptxDocument` 返回的 `document.warnings`。

### 页面切换效果

页面切换按 `<p:transition>` 的子元素名分派。以下效果按 PowerPoint 语义渲染：

| 效果 | 说明 |
|---|---|
| `cut` / `fade` | 直接切换 / 透明度交叉淡入淡出 |
| `push` | 新旧页按源边缘方向同时平移；支持 `l`/`r`/`u`/`d` 与四角方向 |
| `cover` | 新页盖住静止的旧页 |
| `uncover` | 旧页移开露出新页 |
| `wipe` | 从源边缘逐条显示新页；支持 `l`/`r`/`u`/`d`、`lu`/`ld`/`ru`/`rd` |
| `split` | `orient` 决定分割线方向（缺省 `vert` 左右、`horz` 上下），`dir="out"` 从中间向外展开、`dir="in"` 从两侧向内收拢 |

`dissolve` 以交叉淡入淡出近似（像素级溶解无法用 CSS 复现），`morph` 依赖对象身份匹配；两者都属于近似效果，只有在 `approximation="safe"` 时才执行，否则按静态页面处理。未知效果同样只在 `approximation="safe"` 时以淡入淡出显示。能力报告用 `transition:<效果名>` 记录处置：`cover`、`uncover`、`split` 和所有方向的 `wipe` 为 `strict`，`dissolve` 为 `approximate`（默认 `static`）。`morph byWord`/`byChar` 按词补间，对象级 Morph 按对象补间。

三维切换（`cube`、`gallery`、`pan` 等）固定报告为 `transition:3d:<效果名>` 并保持静态页面，即使开启 `approximation="safe"` 也不会退化成二维淡入；切换声音照常由独立的音频层播放，被阻止时进入 `blockedMediaIds`。

`pptx-core` 根入口公开同一套能力判断：`pptxTransitionCapability()`、`isPptxTransitionExecutable()`，以及方向、擦除裁剪、分割遮罩换算和进度换算工具（`pptxTransitionDirection()`、`pptxTransitionWipeInsets()`、`pptxTransitionWipeClip()`、`pptxTransitionSplitMask()`、`pptxTransitionProgress()`），宿主自建播放器时可直接复用。

### 对象引用与宿主证据

`createPptxObjectRegistry()` 和 `createPptxObjectIndex()` 把幻灯片、形状、图片、表格、图表、组合、表格行/单元格、图表内部元素与阶段四行为组织为统一 descriptor。表格行/单元格需要宿主 `getObjectTable(slideIndex, objectKey)` 返回渲染网格；图表内部元素需要 `getObjectChart(slideIndex, objectKey)` 返回系列名称和点标签。缺少这些证据时引用创建被拒绝，解析返回结构化 `unsupported`。

图表内部元素的可见轮廓是 provider-dependent：`PptxChartSnapshot.geometry` 使用 #184 的 `ChartGeometrySnapshot`（chart-local CSS px），宿主再用 `geometryViewport` 给出该 frame 在 slide 上的归一化矩形。对象索引把 series/point/legend 矩形投影为 `visual` 和 zIndex；不提供时仍可从系列/点标签创建引用，但不会伪造坐标。默认 ECharts adapter 仍由宿主提供。

### 从右到左 (RTL) 段落

渲染器会把 `a:pPr/@rtl="1"` 的段落画成 `direction: rtl`。渲染完成后查看器会把同一决定同步到标准 `dir` 属性（并写入 `data-pptx-paragraph-direction`），因此宿主 CSS、文本选择和引用范围都能按段落方向处理。能力报告同时记录：

| 特性 | 处置 | 说明 |
|---|---|---|
| `text:rtl-paragraph` | `strict` | 段落方向由渲染器执行，`dir="rtl"` 透传到 DOM |
| `text:rtl-columns` | `static` | `a:bodyPr/@rtlCol` 的多栏顺序尚未镜像，段落本身仍按 RTL 绘制 |

没有 RTL 段落的页面不会新增能力记录，现有报表计数不变。
