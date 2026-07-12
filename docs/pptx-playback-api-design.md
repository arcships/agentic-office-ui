# PPTX 播放类型与接口

## 1. 状态

- 状态：第一版接口已实现；接口变更需同步更新本文
- 日期：2026-07-11
- 关联设计：[PPTX 播放正式实现设计](pptx-playback-implementation-design.md)
- Vue 无界面补充：[PPTX 最小无界面 API 与原子组件设计](pptx-headless-api-design.md)

本文固定第一版公开名称和职责。实现中可以拆分文件，但不得改变根入口纯净性、会话所有权和事件字段含义。

## 2. 根入口类型

以下类型从 `@arcships/pptx-core` 导出，不访问浏览器对象。

### 2.1 编号规则

所有公开接口中的 `slideIndex`、`initialSlide` 和 `goToSlide(index)` 都从零开始，与现有静态预览一致。界面显示的页码从一开始，使用现有 `PptxSlideInfo.number`，不得把显示页码直接传给控制器。

### 2.2 对象身份

```ts
export type PptxObjectSource = "slide" | "layout" | "master"

export interface PptxObjectIdentity {
  key: string
  slidePath: string
  source: PptxObjectSource
  shapeId: string
  groupPath: readonly string[]
  name?: string
  explicitMorphName?: string
  creationId?: string
  nodeType: string
}

export type PptxMorphMatchMethod =
  | "explicit-name"
  | "creation-id"
  | "name-and-geometry"
  | "text-and-geometry"

export interface PptxMorphMatch {
  from: string
  to: string
  method: PptxMorphMatchMethod
  confidence: "strong" | "medium" | "weak"
  score: number
  unique: boolean
}
```

同页动画只接受 `PptxObjectIdentity.key` 的精确匹配。`PptxMorphMatch` 只用于相邻页 Morph。

### 2.3 时间节点

```ts
export type PptxTimeContainer =
  | "root"
  | "sequence"
  | "parallel"
  | "exclusive"
  | "behavior"
  | "unknown"

export type PptxTimeNodeKind =
  | "group"
  | "effect"
  | "media"
  | "command"
  | "unknown"

export type PptxTriggerEvent =
  | "delay"
  | "on-click"
  | "with-previous"
  | "after-previous"
  | "on-shape-click"
  | "on-begin"
  | "on-end"
  | "on-next"
  | "on-previous"
  | "on-media-bookmark"
  | "on-stop-audio"
  | "unknown"

export interface PptxTimeCondition {
  source: "start" | "end" | "previous" | "next"
  event: PptxTriggerEvent
  delayMs: number
  targetObjectKey?: string
  targetNodeId?: string
  bookmarkName?: string
  rawEvent?: string
}

export interface PptxTimeNode {
  id: string
  parentId?: string
  container: PptxTimeContainer
  kind: PptxTimeNodeKind
  nodeType?: string
  delayMs: number
  durationMs: number | "indefinite"
  repeatCount?: number | "indefinite"
  autoReverse: boolean
  fill: "hold" | "remove" | "freeze" | "unknown"
  restart: "always" | "when-not-active" | "never" | "unknown"
  acceleration: number
  deceleration: number
  conditions: readonly PptxTimeCondition[]
  childIds: readonly string[]
  effect?: PptxPlaybackEffect
  rawSummary?: string
}
```

节点使用编号引用，避免公开模型中递归对象过深，也便于 Worker 传输和序列化。

### 2.4 效果和属性轨道

```ts
export type PptxEffectKind =
  | "appear"
  | "disappear"
  | "fade-in"
  | "fade-out"
  | "wipe"
  | "scale"
  | "rotate"
  | "motion-path"
  | "emphasis"
  | "set"
  | "media-command"
  | "unknown"

export interface PptxPlaybackEffect {
  id: string
  kind: PptxEffectKind
  targetObjectKey?: string
  paragraphRange?: { start: number; end: number }
  presetClass?: string
  presetId?: string
  transition?: "in" | "out"
  filter?: string
  command?: string
  motionPath?: string
  values: Readonly<Record<string, string | number | boolean>>
}

export type PptxTrackProperty =
  | "display"
  | "opacity"
  | "translate-x"
  | "translate-y"
  | "scale-x"
  | "scale-y"
  | "rotate"
  | "clip-path"
  | "filter"
  | "fill-color"
  | "line-color"
  | "text-color"
  | "media-playback"
  | "media-time"
  | "media-volume"

export interface PptxTrackKeyframe {
  timeMs: number
  value: string | number | boolean
  easing?: string
  sourceNodeId: string
}

export interface PptxPropertyTrack {
  objectKey: string
  property: PptxTrackProperty
  initialValue: string | number | boolean
  keyframes: readonly PptxTrackKeyframe[]
  endTimeMs: number | "indefinite"
  repeatStartMs?: number
  repeatPeriodMs?: number
  repeatIndefinite?: boolean
}
```

### 2.5 页面、媒体、动作和能力报告

```ts
export interface PptxSlideTransition {
  kind: string
  direction?: string
  option?: string
  durationMs: number
  advanceOnClick: boolean
  advanceAfterMs?: number
  soundRelationId?: string
}

export interface PptxMediaBookmark {
  name: string
  timeMs: number
}

export interface PptxMediaItem {
  id: string
  objectKey?: string
  relationId: string
  sourcePath?: string
  contentType?: string
  kind: "audio" | "video"
  embedded: boolean
  trimStartMs?: number
  trimEndMs?: number
  loop: boolean
  volume: number
  bookmarks: readonly PptxMediaBookmark[]
}

export interface PptxSlideAction {
  id: string
  sourceObjectKey: string
  trigger: "click" | "hover"
  kind: "go-to-slide" | "open-url" | "mailto" | "unsupported"
  targetSlideIndex?: number
  url?: string
  rawAction?: string
}

export type PptxFeatureDisposition =
  | "strict"
  | "approximate"
  | "static"
  | "unparsed"

export interface PptxFeatureRecord {
  id: string
  slideIndex: number
  objectKey?: string
  feature: string
  disposition: PptxFeatureDisposition
  reason?: string
  sourceNodeId?: string
}

export interface PptxCapabilityReport {
  discovered: number
  strict: number
  approximate: number
  static: number
  unparsed: number
  features: readonly PptxFeatureRecord[]
}

export interface PptxPlaybackSlide {
  index: number
  hidden: boolean
  objects: readonly PptxObjectIdentity[]
  morphFromPrevious: readonly PptxMorphMatch[]
  transition?: PptxSlideTransition
  rootNodeId?: string
  nodes: Readonly<Record<string, PptxTimeNode>>
  media: readonly PptxMediaItem[]
  actions: readonly PptxSlideAction[]
  capability: PptxCapabilityReport
}

export interface PptxPlaybackDocument {
  slides: readonly PptxPlaybackSlide[]
  capability: PptxCapabilityReport
}
```

`trimStartMs` 是从开头裁掉的时长，`trimEndMs` 是从原媒体末尾裁掉的时长；两者都使用毫秒。播放器拿到媒体总时长后计算实际停止位置。

## 3. 错误和警告

现有 `PptxPreviewError` 保持兼容。播放增加独立错误：

```ts
export type PptxPlaybackErrorCode =
  | "PLAYBACK_NOT_READY"
  | "PLAYBACK_DISPOSED"
  | "TARGET_NOT_FOUND"
  | "TARGET_AMBIGUOUS"
  | "TIME_CONDITION_UNSUPPORTED"
  | "TRACK_COMPILE_FAILED"
  | "MEDIA_BLOCKED"
  | "MEDIA_FAILED"
  | "FULLSCREEN_REJECTED"
  | "UNSUPPORTED_FEATURE"

export class PptxPlaybackError extends Error {
  readonly code: PptxPlaybackErrorCode
  readonly slideIndex?: number
  readonly objectKey?: string
  readonly sourceNodeId?: string
  readonly cause?: unknown
}

export interface PptxPlaybackWarning {
  code: string
  message: string
  slideIndex: number
  objectKey?: string
  sourceNodeId?: string
  recoverable: boolean
}
```

单个效果错误通常转换为可恢复警告。文档解析失败、资源超限和会话销毁继续使用现有预览错误。

## 4. 浏览器入口

以下接口从 `@arcships/pptx-core/browser` 导出。

### 4.1 文档会话

```ts
export type PptxApproximationPolicy = "off" | "safe"

export interface PptxDocumentSessionOptions extends PptxPreviewSessionOptions {
  approximation?: PptxApproximationPolicy
  externalMedia?: "disabled" | "allowed"
}

export interface PptxDocumentSession extends PptxPreviewSession {
  readonly playbackDocument: PptxPlaybackDocument | null
  readonly capabilityReport: PptxCapabilityReport | null
  createPlaybackController(options?: PptxPlaybackControllerOptions): PptxPlaybackController
}

export function createPptxDocumentSession(
  container: HTMLElement,
  options?: PptxDocumentSessionOptions,
): PptxDocumentSession
```

`createPptxPreviewSession()` 继续返回现有接口，并内部委托给同一个会话实现。它不自动创建播放控制器。

### 4.2 播放控制器

```ts
export type PptxPlaybackStatus =
  | "idle"
  | "ready"
  | "transitioning"
  | "running"
  | "waiting"
  | "paused"
  | "blocked"
  | "ended"
  | "disposed"

export interface PptxPlaybackSnapshot {
  status: PptxPlaybackStatus
  slideIndex: number
  clickBoundary: number
  positionMs: number
  activeNodeIds: readonly string[]
  blockedMediaIds: readonly string[]
  generation: number
}

export interface PptxPlaybackControllerOptions {
  initialSlide?: number
  skipHiddenSlides?: boolean
  autoplay?: boolean
  approximation?: PptxApproximationPolicy
}

export type PptxPlaybackEvent =
  | { type: "statechange"; snapshot: PptxPlaybackSnapshot }
  | { type: "slidechange"; from: number; to: number; reason: string }
  | { type: "stepchange"; slideIndex: number; boundary: number }
  | { type: "effectstart"; slideIndex: number; nodeId: string; objectKey?: string }
  | { type: "effectend"; slideIndex: number; nodeId: string; objectKey?: string }
  | { type: "warning"; warning: PptxPlaybackWarning }
  | { type: "capability"; report: PptxCapabilityReport }
  | { type: "mediarequest"; mediaId: string; reason: "autoplay-blocked" }
  | { type: "action"; action: PptxActionRequest }
  | { type: "error"; error: PptxPlaybackError }

export interface PptxPlaybackController {
  readonly snapshot: PptxPlaybackSnapshot
  next(): Promise<void>
  activateObject(objectKey: string): Promise<boolean>
  previous(): Promise<void>
  play(): Promise<void>
  pause(): void
  resume(): Promise<void>
  reset(): Promise<void>
  goToSlide(index: number, options?: { includeHidden?: boolean }): Promise<void>
  resumeBlockedMedia(mediaId?: string): Promise<void>
  subscribe(listener: (event: PptxPlaybackEvent) => void): () => void
  dispose(): void
}
```

所有异步方法在控制器销毁或被新命令取代时安全结束，不允许旧操作更新新页面。

`activateObject()` 用于页面元素点击。找到对象触发动画或内部动作时返回 `true`；没有找到专用行为时返回 `false`。Vue 演示界面只在返回 `false` 后把该次点击交给普通 `next()`，避免一次点击同时触发对象序列和下一步。

### 4.3 动作请求

```ts
export type PptxActionRequest =
  | { kind: "go-to-slide"; slideIndex: number; sourceObjectKey: string }
  | { kind: "open-url"; url: string; sourceObjectKey: string }
  | { kind: "mailto"; url: string; sourceObjectKey: string }
  | { kind: "unsupported"; action: string; sourceObjectKey: string }
```

内部跳页由控制器执行，同时发出事件。外部 URL 只发事件，Vue 组件不直接调用 `window.open()`。

## 5. Vue 组件接口

`PptxViewer` 保留现有属性，并增加：

```ts
interface PptxViewerProps {
  source?: PptxPreviewSource | null
  mode?: "browse" | "present"
  initialSlide?: number
  autoplay?: boolean
  approximation?: "off" | "safe"
  showHiddenSlides?: boolean
  showPlaybackControls?: boolean
  showCapabilityStatus?: boolean
  externalMedia?: "disabled" | "allowed"
}

interface PptxViewerExpose {
  getController(): PptxPlaybackController | null
  next(): Promise<void>
  previous(): Promise<void>
  play(): Promise<void>
  pause(): void
  resume(): Promise<void>
  reset(): Promise<void>
  goToSlide(index: number): Promise<void>
  enterFullscreen(): Promise<void>
  exitFullscreen(): Promise<void>
}
```

新增事件：

```ts
interface PptxViewerEmits {
  playbackReady: [controller: PptxPlaybackController]
  playbackStateChange: [snapshot: PptxPlaybackSnapshot]
  stepChange: [slideIndex: number, boundary: number]
  playbackWarning: [warning: PptxPlaybackWarning]
  capability: [report: PptxCapabilityReport]
  mediaRequest: [mediaId: string]
  action: [action: PptxActionRequest]
  playbackError: [error: PptxPlaybackError]
}
```

现有 `loadStart`、`loadSuccess`、`loadError` 和 `slideChange` 不删除、不改名。

## 6. 兼容规则

- 根入口继续通过现有纯净性测试；
- 浏览器入口可以依赖第三方渲染器；
- 现有静态预览调用不需要增加任何属性；
- `mode` 默认 `browse`；
- `approximation` 默认 `off`；
- `externalMedia` 默认 `disabled`；
- 新事件只增加，不改变旧事件参数；
- `PptxPreviewSession` 不增加必须实现的方法，测试替身无需立即实现播放接口。

## 7. 内部接口

下列接口属于包内部，不从公共入口导出：

```ts
interface PptxRendererAdapter {
  renderSlide(index: number, layer: "current" | "incoming" | "outgoing"): Promise<PptxRenderedSlide>
  getObjectElement(key: string): HTMLElement | null
  getParagraphElements(key: string): readonly HTMLElement[]
  captureStaticState(key: string): PptxStaticObjectState
  releaseSlide(index: number): void
  dispose(): void
}

interface PptxClock {
  readonly positionMs: number
  play(): void
  pause(): void
  seek(positionMs: number): void
  subscribe(listener: (positionMs: number) => void): () => void
  dispose(): void
}

interface PptxTrackCompiler {
  compile(slide: PptxPlaybackSlide, staticState: PptxStaticState): PptxCompiledSlide
  evaluate(compiled: PptxCompiledSlide, positionMs: number): PptxObjectStateMap
}
```

内部接口可以调整，但对象查找必须返回唯一元素，轨道求值必须是无副作用纯计算。
