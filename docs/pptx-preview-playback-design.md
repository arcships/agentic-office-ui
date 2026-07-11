# PPTX 预览与播放技术设计

## 1. 文档状态

- 状态：静态预览第一版已实现；播放部分已由正式实现设计取代
- 日期：2026-07-11
- 适用范围：未来的 PPTX 只读预览与播放能力
- 不属于：`0.2.0` 候选版本范围
- 主方案：固定 `@aiden0z/pptx-renderer@1.2.4` 作为静态解析和渲染底座，在项目内实现播放层

本文保留静态预览的产品目标、选型结论、包边界和早期播放判断。当前代码已经提供文件加载、单页浏览、缩略图、搜索、缩放、隐藏页标记、资源限制、取消和销毁；这不表示 PPTX 播放已经可用，也不把第三方项目的功能说明当作本项目的验收证据。

播放探索已经完成。正式实现以 [PPTX 播放正式实现设计](pptx-playback-implementation-design.md)、[PPTX 播放类型与接口](pptx-playback-api-design.md)、[PPTX 播放兼容范围与验收](pptx-playback-compatibility-and-acceptance.md) 和 [PPTX 播放工程化路线图](pptx-playback-roadmap.md) 为准。本文后续涉及播放模型、Worker、公开接口、支持范围和实施顺序的内容只保留为历史背景；若有冲突，以正式实现设计为准。

当前实现位于：

- `packages/pptx-core/`：纯类型入口和显式浏览器入口；
- `packages/vue-pptx/`：`PptxViewer`、缩略图和静态浏览界面；
- `apps/demo/src/pages/PptxViewerPage.vue`：文件上传与浏览演示。

当前限制：演讲者备注、正式动画、页面切换和媒体播放尚未进入公开包；静态底座本身也不解析备注。正式静态视觉基线和工作区外真实压缩包消费仍需后续验收，因此 PPTX 包不并入现有 `0.2.0` 发布候选。播放是否使用 Worker 改为由正式性能结果决定，不再作为第一版前置条件。

## 2. 结论摘要

当前不存在同时满足以下条件的成熟现成库：

1. 完整开源；
2. 不依赖远程文档服务；
3. 直接运行在浏览器客户端；
4. 静态画面接近 PowerPoint；
5. 能执行原 PPTX 的页面切换、对象动画和触发顺序。

因此本项目不再继续寻找一个可以直接接入的“完整播放器”。目标方案拆成两个明确部分：

- `@aiden0z/pptx-renderer` 负责 PPTX 解包、静态模型和 HTML/SVG 渲染；
- 本项目负责读取 `p:transition`、`p:timing`、对象触发条件和媒体关系，并驱动浏览器播放。

`@aiden0z/pptx-renderer` 当前明确不支持动画和页面切换。本次选择它，是因为它是纯浏览器 TypeScript 项目，使用 HTML/SVG 保留单个对象，解析、模型、渲染三层分开，并公开页面和对象模型，适合作为后续开发底座；不是因为它已经满足播放要求。

在正式进入产品包之前，必须先在独立试验目录完成正式素材验证。若静态还原、稳定对象标识或基础时间轴中的任一项不成立，应停止接入，不得一边修改正式包一边继续选型。

## 3. 背景

项目目前已经有 DOCX、XLSX 和 PDF 相关能力，但没有 PPTX 查看器。PPTX 与 DOCX、XLSX 的主要差异不是页面外观，而是演示文稿具有时间和交互语义：

- 页面之间存在切换效果和持续时间；
- 一个对象可以有多段进入、强调、退出或运动动画；
- 动画可以在单击时开始、与上一动画同时开始或在上一动画结束后开始；
- 单击特定对象可以触发动画或跳转页面；
- 音频、视频和自动换页会参与时间安排；
- 隐藏页在普通播放中不应出现，但仍可能被内部链接访问。

PresentationML 把页面切换和对象动画分别保存在幻灯片的 `p:transition` 与 `p:timing` 中。微软的 Open XML 说明也明确区分了页面切换和页内时间动画：[Working with animation](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation)。

因此，只提供缩略图、全屏翻页或查看器自带淡入效果，不视为支持 PPTX 播放。

## 4. 产品标准

### 4.1 基础层：内容可见

基础层要求支持的内容不丢失：

- 文字、图片、表格、常用图表；
- 页面顺序、页面尺寸和背景；
- 母版、版式和主题的主要继承关系；
- 组合、连接线、常用形状和 SmartArt 的可用画面；
- 备注、超链接和媒体对象的存在信息。

### 4.2 静态层：画面还原

静态层要求在支持范围内尽量接近 PowerPoint：

- 对象位置、大小、旋转和层级；
- 字体、字号、行距、项目符号和文字方向；
- 填充、渐变、边框、裁剪和透明度；
- 表格样式、图表类型、坐标轴、图例和数据标签；
- 页面缩放后不重新排版。

静态预览是播放的基础，但不是最终产品目标。

### 4.3 播放层：原文件行为还原

播放层至少应处理：

- 普通播放自动跳过隐藏页；
- 读取页面切换类型、方向、持续时间和自动换页条件；
- 读取单击、与上一动画同时、上一动画之后三种常用开始方式；
- 保持同一次单击触发的并行动画关系；
- 支持进入、强调、退出和运动路径四类对象动画的批准子集；
- 支持延迟、持续时间、重复、自动反向和结束状态；
- 支持嵌入音频和视频的播放、暂停、循环和自动开始；
- 支持页面内跳转、外部链接和指定对象触发；
- 离开页面、重新进入、向后翻页和重新开始时恢复确定状态。

## 5. 范围与非目标

### 5.1 本轮目标

1. 提供只读 PPTX 浏览模式和演示模式；
2. 所有处理都在客户端完成；
3. 提供明确的支持清单和未支持项；
4. 为后续增加动画类型保留稳定模型；
5. 沿用当前项目的 Runtime、Worker、资源预算、取消和销毁规则；
6. 通过 Vue 组件提供产品界面，但把格式解析和播放状态留在核心层。

### 5.2 本轮不做

- 编辑 PPTX；
- 保存或重新生成 PPTX；
- 在线协作和多人播放同步；
- 服务端转 PDF、转图片或调用 Office；
- 宏、ActiveX 和嵌入程序执行；
- 第一阶段承诺 Morph、三维切换、复杂图表动画和逐字逐段动画；
- 为未支持效果伪造一个名称相似但行为不同的动画。

## 6. 约束

1. 第三方依赖必须使用允许修改和分发的开源许可证；主方案当前为 Apache 2.0。
2. 默认不上传文件，不依赖远程文档服务。
3. 浏览器端必须设置压缩包条目数、单项解压大小、总解压大小和媒体总量限制。
4. 大文件解析不得长期阻塞主线程。
5. 新文件加载必须取消旧任务；旧任务完成后不得覆盖新会话。
6. 组件卸载必须停止动画、暂停媒体、撤销对象 URL、终止 Worker 并释放图表实例。
7. 未支持内容必须产生结构化警告，不能静默宣称已经还原。
8. 正式包不能依赖 demo 路径、源码别名或未固定版本的远程资源。

## 7. 选型结论

### 7.1 主方案：`@aiden0z/pptx-renderer`

该项目当前提供：

- 浏览器端 TypeScript 实现；
- HTML/SVG 对象级渲染；
- 形状、文字、图片、表格、图表、SmartArt、组合和主题继承；
- `parseZipLazyMedia()`、`buildPresentation()`、`PresentationData`、`SlideData` 和对象模型；
- 按需页面解析、按需媒体解压和窗口化挂载；
- 压缩包资源限制；
- 基于 PowerPoint 输出的视觉对比流程。

其架构和已知限制见项目说明：[aiden0z/pptx-renderer](https://github.com/aiden0z/pptx-renderer)。当前版本明确把动画和页面切换列为未支持项，因此必须由本项目实现。

选择理由：

1. HTML/SVG 保留对象边界，浏览器可以单独控制对象；
2. `SlideData` 已经解析隐藏属性；
3. `BaseNodeData.id` 来自 PPTX 对象的 `cNvPr/@id`，可以作为动画目标映射的基础；
4. 原始幻灯片 XML 仍可从 `PptxFiles.slides` 读取，播放解析不需要再次解压文件；
5. 模型类型和手动解析流程公开，不必把播放器绑在第三方查看器的默认界面上；
6. TypeScript 代码可以直接审查、补丁和向上游贡献。

当前阻断：渲染得到的对象还没有稳定的 DOM 标记。开始播放开发前，必须为顶层对象和组合内对象补充 `data-pptx-node-id`，并验证它与 `cNvPr/@id` 一致。该改动优先向上游提交；在上游版本发布前，可以固定提交并维护最小补丁，但不建立长期分叉仓库。

### 7.2 不选择的方案

| 方案 | 不作为主方案的原因 |
|---|---|
| ZetaOffice / zetajs | 把完整 LibreOffice 运行时带入浏览器，体积、启动、嵌入和界面控制成本过高，浏览器端完整播放仍缺少正式验证 |
| OfficeCLI | 静态 HTML 预览较强，但现有预览器不执行 `p:timing` 和原页面切换；还要求本地可执行程序，不适合纯网页包 |
| `@silurus/ooxml` | 静态 Canvas 渲染可用，但官方明确不计划支持动画和切换；Canvas 也不利于逐对象控制 |
| `pptx-viewer` | 已有正式素材试验出现隐藏页、对象动画、页面切换和计时异常；在证据归档前不作为候选 |
| `@office-open/pptx` | 能解析和生成部分动画、切换数据，但没有可用的查看器和播放层；可作为实现参考，不同时引入第二套页面模型 |
| PPTXjs、PptxViewJS 等 | 提供静态查看或查看器自身的翻页效果，没有执行原文件时间轴的可靠证据 |

## 8. 总体架构

```text
File / ArrayBuffer / 受控 URL
            │
            ▼
@arcships/office-runtime
来源校验、下载、字节上限、取消、任务编号
            │
            ▼
@arcships/pptx-core/browser
            │
            ├── parseZipLazyMedia() ──► PptxFiles
            │                              │
            │                              ├── buildPresentation()
            │                              │        │
            │                              │        ▼
            │                              │   静态页面模型
            │                              │
            │                              └── 播放解析器
            │                                       │
            │                                       ▼
            │                                页面切换 + 时间安排
            │
            ▼
PptxSession
原始文件、静态模型、播放模型、媒体句柄、警告、dispose
            │
            ▼
@arcships/vue-pptx
            │
            ├── 浏览模式：页面列表、缩略图、搜索、缩放
            └── 演示模式：单页渲染、播放控制、全屏、媒体
                         │
                         ▼
              HTML/SVG 对象 + Web Animations API
```

### 8.1 包边界

未来新增两个公开包，但不加入 `0.2.0`：

| 包 | 职责 | 禁止承担 |
|---|---|---|
| `@arcships/pptx-core` | 公开类型、播放模型、纯状态转换、支持清单、警告和错误码 | Vue 组件、直接访问页面 DOM |
| `@arcships/pptx-core/browser` | 文件加载、第三方解析适配、播放 XML 解析、会话和媒体资源管理 | 产品按钮、工具栏和路由状态 |
| `@arcships/vue-pptx` | Vue 生命周期、渲染器适配、对象 DOM 映射、浏览和演示界面 | 重复解压 PPTX、持有第二份时间安排 |

`@aiden0z/pptx-renderer` 只允许出现在明确的浏览器入口和 Vue 渲染层依赖图中，不能让 `@arcships/pptx-core` 的纯类型入口在导入时访问 `window`、`document` 或 `DOMParser`。

### 8.2 状态所有权

| 对象 | 唯一负责的状态 |
|---|---|
| PPTX Runtime | 不可变配置、全局任务编号、Worker 工厂、资源限制和已创建 loader 清单 |
| loader | 当前请求、AbortController、最新任务 id 和一个专用 Worker |
| `PptxSession` | `PptxFiles`、静态模型、播放模型、对象 URL、媒体句柄、警告和释放状态 |
| 播放控制器 | 当前页面、当前单击步骤、播放状态、活动动画、活动媒体和本次播放编号 |
| Vue 控制器 | 当前 session、渲染句柄、全屏状态和用户界面状态 |

同一状态不得同时存在于 Vue 组件、播放控制器和第三方查看器内部。第三方查看器只负责生成页面对象，不负责决定当前应播放哪一步。

## 9. 播放数据模型

建议模型如下，名称可在实现阶段调整，但职责不能合并：

```ts
interface PptxPlaybackDocument {
  slides: PptxPlaybackSlide[]
  warnings: PptxPlaybackWarning[]
}

interface PptxPlaybackSlide {
  index: number
  hidden: boolean
  transition?: PptxSlideTransition
  steps: PptxPlaybackStep[]
  media: PptxMediaCue[]
  unsupported: PptxUnsupportedFeature[]
}

interface PptxPlaybackStep {
  id: string
  trigger: 'on-click' | 'with-previous' | 'after-previous' | 'on-shape-click'
  triggerShapeId?: string
  delayMs: number
  effects: PptxPlaybackEffect[]
}

interface PptxPlaybackEffect {
  targetShapeId: string
  kind: string
  durationMs: number
  repeat?: number
  autoReverse?: boolean
  holdEndState: boolean
  parameters: Record<string, string | number | boolean>
}
```

该接口是说明性草案。真实实现不能把 `p:seq` 和 `p:par` 简单压平成数组后丢失依赖关系。解析器应保留父子关系，再编译成可以执行的步骤和并行组。

### 9.1 对象映射

播放目标使用 PPTX 的形状 ID，不使用 DOM 顺序或数组下标：

```text
p:spTgt/@spid
      │
      ▼
BaseNodeData.id
      │
      ▼
[data-pptx-node-id="..."]
```

要求：

- 顶层形状、图片、表格、图表和组合必须有稳定标记；
- 组合内子对象也必须可单独定位；
- 母版、版式和页面内的相同数字 ID 需要加来源范围，避免冲突；
- 页面切换或窗口化卸载后，重新挂载必须得到相同标记；
- 找不到目标时记录警告，不得把动画错误应用到相邻对象。

逐段、逐行和逐字动画需要更细的文字标记，不属于第一阶段。第一阶段只批准以整个形状为目标的效果。

## 10. 解析设计

### 10.1 单次解包

文件只解包一次：

1. `office-runtime` 检查输入大小和来源；
2. 使用第三方解析器的严格资源限制读取 ZIP；
3. 得到 `PptxFiles`；
4. 从同一个 `PptxFiles` 构建静态模型和播放模型；
5. 两条路径共享页面关系、媒体关系和页面编号。

不得为了读取动画再用另一套 ZIP 库完整解包一次，也不得让 Vue 组件直接读取 XML。

### 10.2 页面切换

解析 `p:transition` 时至少保留：

- 类型；
- 方向或变体；
- 持续时间；
- 单击换页；
- 自动换页时间；
- 相关声音。

第一批批准类型建议为：无效果、切换、淡化、推进、擦除。其他类型先生成 `unsupported-transition` 警告。不同实现对同一种切换的像素细节可能不同，但方向、持续时间和遮挡顺序必须一致。

### 10.3 页内时间安排

解析器从 `p:timing` 读取：

- 顺序组和并行组；
- 通用时间节点、延迟、持续时间和重复；
- 开始条件和事件；
- 目标形状、文字范围或图表目标；
- 数值、颜色、透明度、位置、缩放和旋转变化；
- 运动路径；
- 音频和视频节点。

第一阶段先支持整个形状的常用效果。文字分段、图表内部序列、OLE 和复杂命令节点保留原始信息并报告未支持，不进入错误的近似播放。

## 11. 播放控制

### 11.1 播放状态

播放控制器至少有以下状态：

```text
idle → ready → playing → paused → ended
          ↑         │         │
          └──────── reset / seek
```

每次进入页面或重新开始都生成新的播放编号。旧编号创建的计时器、动画回调和媒体事件不得更新当前状态。

### 11.2 单击规则

单击处理顺序：

1. 如果单击对象有专用触发动画，执行该对象对应步骤；
2. 否则如果当前页面还有等待单击的步骤，执行下一组；
3. 否则进入下一张非隐藏页；
4. 如果当前对象是内部跳转或安全外部链接，按批准的动作规则处理，不同时触发普通翻页。

“与上一动画同时”加入同一并行组；“上一动画之后”由前一组完成事件触发，不消耗额外单击。

### 11.3 初始状态与重置

进入页面前先应用动画初始状态，例如进入动画的对象需要在开始前不可见。只有已批准且可以执行的效果才能修改初始状态；未支持的进入动画不能把对象永久隐藏。

页面重置必须：

- 取消所有浏览器动画；
- 恢复对象原始样式；
- 暂停并归零受控媒体；
- 清除临时遮罩；
- 把步骤指针恢复到页面开头。

### 11.4 页面切换

页面切换使用同时存在的旧页面层和新页面层。切换完成后立即释放旧页面渲染句柄。页面切换结束后，才能执行新页面的自动开始动画。

### 11.5 浏览器全屏与媒体

- 演示模式使用浏览器 Fullscreen API，不把隐藏工具栏的普通页面误称为全屏；
- 浏览器拒绝无用户操作的媒体播放时，界面显示明确的“点击开始播放”状态；
- 嵌入媒体使用受控对象 URL；
- 外部媒体默认不加载，只有通过 URL 规则和产品配置后才允许；
- 离开页面时暂停页面内媒体，除非未来明确支持跨页旁白。

## 12. Vue 产品界面

### 12.1 浏览模式

浏览模式提供：

- 页面列表或单页查看；
- 缩略图；
- 缩放和适应窗口；
- 页面跳转；
- 文本搜索；
- 隐藏页标记；
- 备注查看；
- 未支持项提示。

浏览模式不自动执行对象动画。

### 12.2 演示模式

演示模式提供：

- 单页舞台；
- 上一步、下一步；
- 播放、暂停、重新开始；
- 全屏进入和退出；
- 当前页和总页数；
- 可选的支持状态提示。

键盘至少支持方向键、空格、Home、End 和 Escape。触摸设备支持点击和左右滑动，但滑动不得抢占页面内对象触发。

### 12.3 公开接口草案

```ts
interface PptxViewerProps {
  source: File | ArrayBuffer | Uint8Array | string
  mode?: 'browse' | 'present'
  initialSlide?: number
  autoplay?: boolean
  showHiddenSlides?: boolean
  runtime?: PptxRuntime
}

interface PptxViewerController {
  next(): Promise<void>
  previous(): Promise<void>
  goToSlide(index: number): Promise<void>
  play(): Promise<void>
  pause(): void
  reset(): Promise<void>
  enterFullscreen(): Promise<void>
  exitFullscreen(): Promise<void>
  dispose(): void
}
```

事件至少包括：`ready`、`slidechange`、`stepchange`、`playstatechange`、`warning`、`unsupported` 和 `error`。事件字段和错误码必须在公开接口文档中固定，不能直接透出第三方异常文本作为稳定接口。

## 13. 未支持内容的处理

本项目不使用服务端图片或 PDF 作为替代路径。遇到未支持内容时：

1. 静态对象能够显示时，保持静态画面；
2. 未支持动画不执行，也不伪造成其他效果；
3. 未支持进入动画不能导致对象消失；
4. 未支持退出动画不能提前移除对象；
5. 页面和对象记录结构化警告；
6. 产品可以选择向用户展示“部分播放效果未还原”。

建议警告结构：

```ts
interface PptxPlaybackWarning {
  code: string
  slideIndex: number
  shapeId?: string
  feature: string
  message: string
}
```

## 14. 安全、资源与性能

### 14.1 输入安全

- 默认启用严格 ZIP 限制，不使用第三方库的无限制兼容模式；
- 限制条目数、单项解压大小、总解压大小和媒体总量；
- XML 解析不允许外部实体；
- 外部链接只允许批准协议；
- 不执行宏、嵌入程序、脚本或任意命令；
- 对象 URL 在会话销毁时统一撤销。

### 14.2 主线程与 Worker

早期设计曾要求把解包、XML 读取和时间安排编译全部放入 Worker。真实大文件探索后，正式设计改为第一版在主线程完成一次解包和第三方建模；只有固定环境三轮性能结果超过门槛时，才把可序列化的 XML 扫描和播放模型编译迁入 Worker。详细条件见正式实现设计，且任何方案都不能为 Worker 再解包一次 PPTX。

### 14.3 大文件

- 默认启用按需页面解析、按需媒体解压和窗口化页面挂载；
- 演示模式最多保留当前页、上一页和下一页的必要渲染结果；
- 预加载下一页时不得提前启动媒体和动画；
- 缩略图需要独立预算，不能永久保留全部页面的完整 DOM；
- 图表、媒体和对象 URL 必须随页面卸载释放。

性能预算不得在没有正式基线时随意填写。正式实现进入性能验收时，把批准环境、素材、冷启动、首次可见页、切页、内存峰值和释放后内存写入 `docs/testing/performance-baseline.md`。

## 15. 验证方案

### 15.1 素材分层

1. 最小构造样本：每个文件只验证一种 XML 规则；
2. Microsoft 官方演示文稿：验证真实主题、母版、图表、隐藏页、动画和切换组合；
3. NASA 或其他公开正式演示文稿：验证复杂页面和媒体；
4. 项目回归样本：保存每次发现问题的最小文件；
5. 大文件样本：验证按需解析、窗口化和资源释放。

正式素材必须保存来源、许可证或公开下载地址、文件哈希和基准 PowerPoint 版本。

### 15.2 自动检查

| 层级 | 检查内容 |
|---|---|
| 单元 | `p:transition`、`p:timing`、触发条件、隐藏页、媒体关系和未支持警告 |
| 核心状态 | 单击步骤、并行与顺序关系、暂停、继续、跳页、重置和旧任务失效 |
| 组件 | 对象 ID 映射、浏览器动画调用、全屏、键盘、媒体限制和销毁 |
| 视觉 | PowerPoint 基准图与浏览器静态截图逐页对比 |
| 播放 | PowerPoint 基准录屏与浏览器录屏按关键帧、顺序和持续时间对比 |
| 黑盒 | 正式压缩包安装后加载真实 PPTX，不允许源码深层导入 |

### 15.3 正式工程化验证

探索已经完成，正式门槛已经写入 [PPTX 播放兼容范围与验收](pptx-playback-compatibility-and-acceptance.md)。对象唯一命中、事件顺序、时间和属性误差、三种浏览器、资源释放、性能与外部消费必须按该文档执行，不能把探索结果直接写成正式兼容声明。

## 16. 实施顺序

包边界和静态浏览第一版已经完成，三段播放探索也已经收口。正式开发按 [PPTX 播放工程化路线图](pptx-playback-roadmap.md) 执行：先固定类型和渲染器补丁，再建立一次解包会话、稳定对象身份、时间节点树、属性轨道和播放控制器，最后接入切换、媒体、Vue 演示模式与发布验收。

探索代码继续作为证据保留，不直接复制成 `@arcships/pptx-core` 或 `@arcships/vue-pptx` 的正式实现。

## 17. 风险与停止条件

### 17.1 主要风险

1. 第三方渲染器仍在快速开发，公开模型可能变化；
2. DOM 对象标识需要上游改动或本地补丁；
3. 浏览器文字排版与 PowerPoint 不完全一致；
4. PPTX 时间安排是树状关系，简单效果之外的复杂度上升很快；
5. Morph 需要跨页对象匹配，不是普通页面切换；
6. 浏览器媒体自动播放策略会影响自动演示；
7. 大型媒体文档可能同时造成解压、内存和页面挂载压力。

### 17.2 停止条件

出现以下任一情况，应停止进入正式包并重新决策：

- 正式素材静态画面明显低于产品可接受水平；
- 不能建立稳定的形状 ID 到 DOM 对象映射；
- 基础时间安排无法从真实文件稳定编译；
- 为适配第三方库需要长期维护大面积分叉；
- 冷启动、切页或内存峰值超过批准预算且按需加载无法解决；
- 第一阶段支持范围不足以覆盖目标用户的常见演示文稿。

停止后只有两个诚实选择：缩小产品对播放的承诺，或者建设更完整的自有 PPTX 渲染与播放引擎。不得重新包装静态翻页并继续称为“原文件播放”。

## 18. 已形成的正式结论

探索之后已经确定：固定 `@aiden0z/pptx-renderer@1.2.4` 并维护范围明确的补丁；同页对象只用精确身份；Morph 强身份进入近似范围；外部媒体默认禁用；正式发布检查 Chromium、Firefox 和 WebKit；性能和画面门槛按兼容与验收文档执行；PPTX 继续保留为两个独立开发包，完成发布门槛前不进入公开发布合同。

## 19. 参考资料

- [aiden0z/pptx-renderer](https://github.com/aiden0z/pptx-renderer)
- [Office Open XML Viewer / @silurus/ooxml](https://github.com/yukiyokotani/office-open-xml-viewer)
- [Office Open / @office-open/pptx](https://github.com/DemoMacro/office-open)
- [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)
- [ZetaOffice](https://zetaoffice.net/)
- [Microsoft Learn：Working with animation](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation)
- [Microsoft Learn：PowerPoint 动画开始方式和时间](https://support.microsoft.com/en-US/PowerPoint/set-the-start-time-and-speed-of-an-animation-effect)
- [Microsoft Learn：Morph transition](https://support.microsoft.com/en-US/PowerPoint/morph-transition-tips-and-tricks)
