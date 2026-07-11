import {
  RECOMMENDED_ZIP_LIMITS,
  buildPresentation,
  parseZipLazyMedia,
  renderSlide,
  type BaseNodeData,
  type PptxFiles,
  type PresentationData,
  type SlideData,
  type SlideHandle,
} from "@aiden0z/pptx-renderer"
import { parseSlideTiming, type LabBatch, type LabEffect, type LabTimeNode, type LabTiming } from "./timing"
import "./style.css"

const root = document.querySelector<HTMLElement>("#app")!
root.innerHTML = `
  <main class="lab">
    <header class="lab__header">
      <div><h1>PPTX 播放探索台</h1><p>读取原始 timing，快速观察真实播放结果</p></div>
      <label class="upload">选择 PPTX<input id="file" type="file" accept=".pptx" /></label>
    </header>
    <section class="controls">
      <button id="prev-slide">上一页</button><button id="next-slide">下一页</button>
      <span id="slide-count">0 / 0</span>
      <input id="slide-number" class="page-input" type="number" min="1" value="1" aria-label="页面编号" /><button id="go-slide">跳页</button>
      <button id="next-effect" class="primary">下一动画</button>
      <button id="play-slide">播放本页</button><button id="pause">暂停</button>
      <button id="resume">继续</button><button id="reset">重置</button><button id="play-audio">播放音频</button>
      <label>速度 <select id="speed"><option value="0.25">0.25×</option><option value="0.5">0.5×</option><option value="1" selected>1×</option><option value="2">2×</option></select></label>
      <label><input id="show-targets" type="checkbox" /> 显示目标</label>
      <label><input id="auto-advance" type="checkbox" /> 自动换页</label>
    </section>
    <section class="workspace">
      <div class="stage-column">
        <div id="status" class="status">请选择包含动画的 PPTX。</div>
        <div id="stage" class="stage" tabindex="0"><div class="stage__empty">等待文件</div></div>
      </div>
      <aside class="debug">
        <section><h2>解析摘要</h2><pre id="summary">尚未加载</pre></section>
        <section><h2>事件记录</h2><pre id="trace"></pre></section>
        <details open><summary>时间树 XML</summary><pre id="timing-xml"></pre></details>
        <details><summary>页面切换 XML</summary><pre id="transition-xml"></pre></details>
      </aside>
    </section>
  </main>`

const fileInput = byId<HTMLInputElement>("file")
const stage = byId<HTMLElement>("stage")
const status = byId<HTMLElement>("status")
const slideCount = byId<HTMLElement>("slide-count")
const summary = byId<HTMLElement>("summary")
const trace = byId<HTMLElement>("trace")
const timingXml = byId<HTMLElement>("timing-xml")
const transitionXml = byId<HTMLElement>("transition-xml")
const speed = byId<HTMLSelectElement>("speed")
const showTargets = byId<HTMLInputElement>("show-targets")
const slideNumber = byId<HTMLInputElement>("slide-number")
const autoAdvance = byId<HTMLInputElement>("auto-advance")

let files: PptxFiles | null = null
let presentation: PresentationData | null = null
let slideHandle: SlideHandle | null = null
let slideIndex = 0
let timing: LabTiming = emptyTiming()
let batchIndex = 0
let targetMap = new Map<string, HTMLElement>()
let runningAnimations: Animation[] = []
let traceLines: string[] = []
let advanceId = 0
let advancing = false
let interactionCleanups: Array<() => void> = []
let mediaUrls: string[] = []
let audioTargets = new Map<string, HTMLAudioElement>()
let firedBookmarks = new Set<string>()
let renderedSlideIndex = -1
let renderedSlideXml = ""
let slideScale = 1

interface TimerJob {
  handle: number
  remaining: number
  startedAt: number
  callback: () => void
  active: boolean
}

let timerJobs: TimerJob[] = []
let timersPaused = false

function byId<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T
}

function emptyTiming(): LabTiming {
  return {
    batches: [], effects: [], timingXml: "", transitionXml: "", unknownPresetClasses: [],
    interactive: [], builds: [], transition: null, audioCues: [], mediaBookmarks: [],
    bookmarkSequences: [], conflicts: [],
    complexity: { timeNodeCount: 0, maxDepth: 0, conditionEvents: [], repeatNodeCount: 0, autoReverseNodeCount: 0, accelerationNodeCount: 0 },
    timeTree: null,
    treeValidation: { treeEffectIds: [], parsedEffectIds: [], matchedEffectIds: [], missingEffectIds: [], orderMatches: true, conditionEvents: [] },
  }
}

function log(message: string): void {
  const stamp = new Date().toISOString().slice(11, 23)
  traceLines.push(`${stamp} ${message}`)
  trace.textContent = traceLines.slice(-120).join("\n")
  trace.scrollTop = trace.scrollHeight
}

function clearPlayback(disposeView = false): void {
  advanceId += 1
  advancing = false
  for (const animation of runningAnimations) animation.cancel()
  for (const job of timerJobs) window.clearTimeout(job.handle)
  if (disposeView) {
    interactionCleanups.splice(0).forEach((cleanup) => cleanup())
    mediaUrls.splice(0).forEach((url) => URL.revokeObjectURL(url))
    audioTargets.clear()
    firedBookmarks.clear()
  }
  runningAnimations = []
  timerJobs = []
  timersPaused = false
}

function xmlElements(root: Document | Element, localName: string): Element[] {
  return Array.from(root.getElementsByTagNameNS("*", localName))
}

function relationshipPath(slidePath: string): string {
  const slash = slidePath.lastIndexOf("/")
  const dir = slidePath.slice(0, slash)
  const file = slidePath.slice(slash + 1)
  return `${dir}/_rels/${file}.rels`
}

function enclosingShapeId(node: Element): string {
  let current: Element | null = node
  while (current && !["pic", "sp", "graphicFrame"].includes(current.localName)) current = current.parentElement
  return current ? xmlElements(current, "cNvPr")[0]?.getAttribute("id") ?? "" : ""
}

async function hydrateMediaTargets(xml: string, slide: SlideData): Promise<void> {
  if (!files?.mediaResolver) return
  const doc = new DOMParser().parseFromString(xml, "application/xml")
  const relXml = files.slideRels.get(relationshipPath(slide.slidePath))
  if (!relXml) return
  const relDoc = new DOMParser().parseFromString(relXml, "application/xml")
  const relationships = new Map(xmlElements(relDoc, "Relationship").map((rel) => [
    rel.getAttribute("Id") ?? "",
    rel.getAttribute("Target") ?? "",
  ]))
  const mediaNodes = [...xmlElements(doc, "videoFile"), ...xmlElements(doc, "audioFile")]
  for (const mediaNode of mediaNodes) {
    const relationId = Array.from(mediaNode.attributes).find((attr) => ["link", "embed"].includes(attr.localName))?.value ?? ""
    const relationTarget = relationships.get(relationId)
    const shapeId = enclosingShapeId(mediaNode)
    const target = targetMap.get(shapeId)
    if (!relationTarget || !target) continue
    const media = await files.mediaResolver.resolve(relationTarget)
    if (!media) { log(`MEDIA RESOLVE MISS target=${shapeId} rel=${relationId}`); continue }
    const video = mediaNode.localName === "videoFile"
    const bytes = new Uint8Array(media.data.byteLength)
    bytes.set(media.data)
    const blob = new Blob([bytes.buffer], { type: video ? "video/mp4" : "audio/mpeg" })
    const url = URL.createObjectURL(blob)
    mediaUrls.push(url)
    const element = document.createElement(video ? "video" : "audio")
    element.src = url
    element.controls = true
    element.preload = "metadata"
    element.style.width = "100%"
    element.style.height = "100%"
    if (video) (element as HTMLVideoElement).style.objectFit = "contain"
    target.replaceChildren(element)
    log(`MEDIA HYDRATE target=${shapeId} type=${video ? "video" : "audio"} bytes=${media.data.byteLength}`)
  }
  for (const cue of timing.audioCues) {
    const relationTarget = relationships.get(cue.relationId)
    if (!relationTarget) continue
    const media = await files.mediaResolver.resolve(relationTarget)
    if (!media) { log(`AUDIO RESOLVE MISS rel=${cue.relationId}`); continue }
    const bytes = new Uint8Array(media.data.byteLength)
    bytes.set(media.data)
    const extension = relationTarget.split(".").pop()?.toLowerCase()
    const mime = extension === "wav" ? "audio/wav" : extension === "m4a" ? "audio/mp4" : "audio/mpeg"
    const url = URL.createObjectURL(new Blob([bytes.buffer], { type: mime }))
    mediaUrls.push(url)
    const audio = document.createElement("audio")
    audio.src = url
    audio.preload = "metadata"
    audio.hidden = true
    audio.dataset.pptxRelationId = cue.relationId
    stage.appendChild(audio)
    audioTargets.set(cue.relationId, audio)
    log(`AUDIO HYDRATE rel=${cue.relationId} name=${cue.name || "-"} bytes=${media.data.byteLength}`)
  }
}

function schedule(callback: () => void, delay: number): TimerJob {
  const job: TimerJob = { handle: 0, remaining: Math.max(0, delay), startedAt: performance.now(), callback, active: true }
  const start = () => {
    job.startedAt = performance.now()
    job.handle = window.setTimeout(() => {
      job.active = false
      callback()
    }, job.remaining)
  }
  if (!timersPaused) start()
  timerJobs.push(job)
  return job
}

function waitFor(delay: number): Promise<void> {
  return new Promise((resolve) => schedule(resolve, delay))
}

function fitSlide(): void {
  if (!presentation || !slideHandle) return
  const availableW = Math.max(stage.clientWidth - 24, 1)
  const availableH = Math.max(stage.clientHeight - 24, 1)
  const scale = Math.min(availableW / presentation.width, availableH / presentation.height)
  slideScale = scale
  slideHandle.element.style.position = "absolute"
  slideHandle.element.style.left = "50%"
  slideHandle.element.style.top = "50%"
  slideHandle.element.style.transformOrigin = "center"
  slideHandle.element.style.transform = `translate(-50%, -50%) scale(${scale})`
}

function mapTopLevelTargets(slide: SlideData, element: HTMLElement): Map<string, HTMLElement> {
  const direct = Array.from(element.children).filter((child): child is HTMLElement => child instanceof HTMLElement)
  const renderedNodes = direct.slice(Math.max(0, direct.length - slide.nodes.length))
  const mapped = new Map<string, HTMLElement>()
  slide.nodes.forEach((node: BaseNodeData, index) => {
    const target = renderedNodes[index]
    if (!target || !node.id) return
    target.dataset.pptxNodeId = node.id
    target.dataset.pptxNodeType = node.nodeType
    target.title = `PPTX ${node.nodeType} #${node.id} ${node.name}`
    mapped.set(node.id, target)
  })
  return mapped
}

function restoreTarget(target: HTMLElement): void {
  target.style.opacity = ""
  target.style.visibility = ""
  target.style.translate = ""
  target.style.scale = ""
  target.style.rotate = ""
  target.style.filter = ""
  target.style.clipPath = ""
  target.style.offsetPath = ""
  target.style.offsetDistance = ""
}

function paragraphElements(target: HTMLElement): HTMLElement[] {
  const candidates = Array.from(target.querySelectorAll<HTMLElement>("div")).filter((node) => {
    const text = node.textContent?.trim() ?? ""
    if (!text) return false
    return Array.from(node.children).some((child) => child instanceof HTMLSpanElement || child instanceof HTMLBRElement)
  })
  return candidates.filter((node) => !candidates.some((other) => other !== node && other.contains(node)))
}

function prepareInitialState(): void {
  for (const target of targetMap.values()) restoreTarget(target)
  for (const effect of timing.effects) {
    const target = targetMap.get(effect.targetId)
    if (!target) continue
    if (effect.presetClass === "entr") {
      const build = timing.builds.find((item) => item.kind === "paragraph" && item.targetId === effect.targetId && item.mode === "p")
      if (build) {
        target.style.visibility = "visible"
        paragraphElements(target).forEach((paragraph) => { paragraph.style.opacity = "0" })
      } else {
        target.style.opacity = "0"
        target.style.visibility = "hidden"
      }
    }
  }
}

function scaledMs(value: number): number {
  return value / Number(speed.value || 1)
}

function motionEnd(path: string): { x: number; y: number; curved: boolean } {
  const numbers = path.match(/-?\d*\.?\d+(?:E[+-]?\d+)?/gi)?.map(Number) ?? []
  const x = numbers.at(-2) ?? 0
  const y = numbers.at(-1) ?? 0
  return {
    x: x * (presentation?.width ?? 960),
    y: y * (presentation?.height ?? 540),
    curved: /\b[CGQ]\b/i.test(path),
  }
}

function resolveEffectTargets(effect: LabEffect, target: HTMLElement): HTMLElement[] {
  const paragraphs = paragraphElements(target)
  if (effect.paragraphRange && paragraphs.length) {
    return paragraphs.slice(effect.paragraphRange.start, effect.paragraphRange.end + 1)
  }
  const build = timing.builds.find((item) => item.kind === "paragraph" && item.targetId === effect.targetId && item.mode === "p")
  return build && paragraphs.length ? paragraphs : [target]
}

function runEffect(effect: LabEffect): void {
  const treePath = findTimeNodePath(timing.timeTree, effect.id)
  log(`TREE EXEC effect=${effect.id} path=${treePath || "missing"}`)
  const rootTarget = targetMap.get(effect.targetId)
  if (!rootTarget) {
    log(`MISS target=${effect.targetId} class=${effect.presetClass} preset=${effect.presetId}`)
    return
  }
  const targets = resolveEffectTargets(effect, rootTarget)
  if (effect.paragraphRange) {
    log(`TEXT RANGE target=${effect.targetId} start=${effect.paragraphRange.start} end=${effect.paragraphRange.end} matched=${targets.length}`)
  }
  if (targets.length > 1) {
    targets.forEach((target, index) => schedule(() => runEffectOnTarget(effect, target), scaledMs(index * 120)))
    log(`TEXT BUILD target=${effect.targetId} paragraphs=${targets.length}`)
    return
  }
  runEffectOnTarget(effect, targets[0] ?? rootTarget)
}

function findTimeNodePath(node: LabTimeNode | null, id: string, parentPath = ""): string {
  if (!node) return ""
  const path = `${parentPath}/${node.container || "root"}:${node.id || "-"}`
  if (node.id === id) return path
  for (const child of node.children) {
    const match = findTimeNodePath(child, id, path)
    if (match) return match
  }
  return ""
}

function runEffectOnTarget(effect: LabEffect, target: HTMLElement): void {
  const duration = scaledMs(effect.duration)
  target.style.visibility = "visible"
  if (effect.scaleBy || effect.scaleFrom || effect.scaleTo || effect.rotationBy || effect.attributes.length) {
    const scaleText = effect.scaleFrom || effect.scaleTo
      ? `${effect.scaleFrom?.x ?? 1},${effect.scaleFrom?.y ?? 1}->${effect.scaleTo?.x ?? 1},${effect.scaleTo?.y ?? 1}`
      : effect.scaleBy ? `${effect.scaleBy.x},${effect.scaleBy.y}` : "-"
    log(`PROPS target=${effect.targetId} scale=${scaleText} rotate=${effect.rotationBy} attrs=${effect.attributes.join("|") || "-"} values=${effect.values.slice(0, 6).join("|") || "-"}`)
  }
  if (effect.kind === "media") {
    const media = target.querySelector<HTMLMediaElement>("video, audio")
    if (!media) {
      log(`MEDIA MISS target=${effect.targetId} command=${effect.command || "-"}`)
      return
    }
    if (effect.command === "togglePause") {
      if (media.paused) void media.play().catch((error) => log(`MEDIA BLOCKED ${String(error)}`))
      else media.pause()
    } else {
      void media.play().catch((error) => log(`MEDIA BLOCKED ${String(error)}`))
    }
    log(`MEDIA target=${effect.targetId} command=${effect.command || "play"}`)
    return
  }
  const hasAnimatedProperties = Boolean(effect.scaleBy || effect.scaleFrom || effect.scaleTo || effect.rotationBy)
  if ((effect.kind === "appear" || effect.kind === "disappear") && !effect.filter && !hasAnimatedProperties) {
    const visible = effect.kind === "appear"
    target.style.opacity = visible ? "1" : "0"
    target.style.visibility = visible ? "visible" : "hidden"
    log(`INSTANT target=${effect.targetId} kind=${effect.kind} preset=${effect.presetId}`)
    return
  }
  let frames: Keyframe[]
  if (effect.filter.startsWith("wipe")) {
    const outgoing = effect.transition === "out"
    frames = outgoing
      ? [{ clipPath: "inset(0 0 0 0)" }, { clipPath: "inset(0 100% 0 0)" }]
      : [{ opacity: 1, clipPath: "inset(0 100% 0 0)" }, { opacity: 1, clipPath: "inset(0 0 0 0)" }]
  } else if (effect.kind === "fadeOut") {
    frames = [{ opacity: 1 }, { opacity: 0 }]
  } else if (effect.kind === "emphasis") {
    const scale = effect.scaleBy ? `${effect.scaleBy.x} ${effect.scaleBy.y}` : "1.08"
    const rotate = effect.rotationBy ? `${effect.rotationBy}deg` : "0deg"
    frames = [{ opacity: 1, scale: "1", rotate: "0deg", filter: "brightness(1)" }, { opacity: 0.65, scale, rotate, filter: "brightness(1.5)" }, { opacity: 1, scale: "1", rotate: "0deg", filter: "brightness(1)" }]
  } else if (effect.kind === "path") {
    const end = motionEnd(effect.motionPath)
    frames = [{ translate: "0 0" }, { translate: `${end.x}px ${end.y}px` }]
    log(`MOTION target=${effect.targetId} curved=${end.curved} path=${effect.motionPath || "-"}`)
  } else if (effect.scaleBy || effect.scaleFrom || effect.scaleTo || effect.rotationBy) {
    const fromScale = effect.scaleFrom ?? effect.scaleBy ?? { x: 1, y: 1 }
    const toScale = effect.scaleTo ?? { x: 1, y: 1 }
    frames = [
      { opacity: effect.presetClass === "entr" ? 0 : 1, scale: `${fromScale.x} ${fromScale.y}`, rotate: `${effect.rotationBy}deg` },
      { opacity: 1, scale: `${toScale.x} ${toScale.y}`, rotate: "0deg" },
    ]
  } else {
    frames = [{ opacity: 0 }, { opacity: 1 }]
  }
  const repeats = effect.repeatCount === "indefinite" ? 2 : effect.repeatCount ?? 1
  const iterations = repeats * (effect.autoReverse ? 2 : 1)
  const easing = effect.deceleration >= 0.8 ? "ease-out" : effect.acceleration >= 0.8 ? "ease-in" : "ease"
  const animation = target.animate(frames, {
    duration,
    fill: "forwards",
    easing,
    iterations,
    direction: effect.autoReverse ? "alternate" : "normal",
  })
  runningAnimations.push(animation)
  log(`START target=${effect.targetId} kind=${effect.kind} trigger=${effect.trigger} preset=${effect.presetId} filter=${effect.filter || "-"} duration=${effect.duration} iterations=${iterations} autoReverse=${effect.autoReverse} easing=${easing}`)
  animation.finished.then(() => {
    if (effect.kind === "fadeOut" || effect.transition === "out") target.style.visibility = "hidden"
    log(`END target=${effect.targetId} kind=${effect.kind}`)
  }).catch(() => undefined)
}

interface PropertyTrack {
  values: [string | number, string | number, string | number]
  owner: LabEffect
}

function effectTracks(effect: LabEffect): Map<string, PropertyTrack> {
  const tracks = new Map<string, PropertyTrack>()
  const add = (property: string, values: PropertyTrack["values"]) => tracks.set(property, { values, owner: effect })
  if (effect.filter.startsWith("wipe")) {
    const outgoing = effect.transition === "out"
    add("clipPath", outgoing
      ? ["inset(0 0 0 0)", "inset(0 50% 0 0)", "inset(0 100% 0 0)"]
      : ["inset(0 100% 0 0)", "inset(0 50% 0 0)", "inset(0 0 0 0)"])
    add("opacity", [1, 1, outgoing ? 0 : 1])
  } else if (effect.kind === "emphasis") {
    const scale = effect.scaleBy ? `${effect.scaleBy.x} ${effect.scaleBy.y}` : "1.08"
    add("opacity", [1, 0.65, 1])
    add("scale", ["1", scale, "1"])
    add("rotate", ["0deg", effect.rotationBy ? `${effect.rotationBy}deg` : "0deg", "0deg"])
    add("filter", ["brightness(1)", "brightness(1.5)", "brightness(1)"])
  } else if (effect.kind === "path") {
    const end = motionEnd(effect.motionPath)
    add("translate", ["0 0", `${end.x / 2}px ${end.y / 2}px`, `${end.x}px ${end.y}px`])
  } else {
    const outgoing = effect.kind === "fadeOut" || effect.kind === "disappear" || effect.transition === "out"
    const incoming = effect.kind === "fadeIn" || effect.kind === "appear" || effect.transition === "in"
    if (outgoing) add("opacity", [1, 0.5, 0])
    else if (incoming) add("opacity", [0, 0.5, 1])
  }
  if (effect.scaleBy || effect.scaleFrom || effect.scaleTo) {
    const from = effect.scaleFrom ?? effect.scaleBy ?? { x: 1, y: 1 }
    const to = effect.scaleTo ?? { x: 1, y: 1 }
    add("scale", [`${from.x} ${from.y}`, `${to.x} ${to.y}`, `${to.x} ${to.y}`])
  }
  if (effect.rotationBy) add("rotate", [`${effect.rotationBy}deg`, `${effect.rotationBy / 2}deg`, "0deg"])
  return tracks
}

function runEffectGroup(effects: LabEffect[]): void {
  const mediaEffects = effects.filter((effect) => effect.kind === "media")
  const visualEffects = effects.filter((effect) => effect.kind !== "media")
  mediaEffects.forEach(runEffect)
  if (visualEffects.length <= 1 || visualEffects.some((effect) => effect.paragraphRange || effect.repeatCount || effect.autoReverse)) {
    visualEffects.forEach(runEffect)
    return
  }
  const target = targetMap.get(visualEffects[0].targetId)
  if (!target) { visualEffects.forEach(runEffect); return }
  const preparedTracks = visualEffects.map((effect) => ({ effect, tracks: effectTracks(effect) }))
  if (preparedTracks.some((item) => item.tracks.size === 0)) { visualEffects.forEach(runEffect); return }
  const tracks = new Map<string, PropertyTrack>()
  preparedTracks.forEach(({ effect, tracks: effectTrackMap }) => {
    log(`TREE EXEC effect=${effect.id} path=${findTimeNodePath(timing.timeTree, effect.id) || "missing"}`)
    effectTrackMap.forEach((track, property) => tracks.set(property, track))
  })
  target.style.visibility = "visible"
  const frames: [Keyframe, Keyframe, Keyframe] = [{ offset: 0 }, { offset: 0.5 }, { offset: 1 }]
  tracks.forEach((track, property) => track.values.forEach((value, index) => { frames[index][property as keyof Keyframe] = value as never }))
  const duration = scaledMs(Math.max(...visualEffects.map((effect) => effect.duration)))
  const animation = target.animate(frames, { duration, fill: "forwards", easing: "ease" })
  runningAnimations.push(animation)
  const owners = Array.from(tracks, ([property, track]) => `${property}:${track.owner.id}`).join("|")
  log(`MERGE START target=${visualEffects[0].targetId} effects=${visualEffects.map((effect) => effect.id).join("+")} tracks=${owners}`)
  animation.finished.then(() => {
    const opacity = tracks.get("opacity")
    target.style.visibility = opacity && Number(opacity.values[2]) === 0 ? "hidden" : "visible"
    log(`MERGE END target=${visualEffects[0].targetId} opacity=${getComputedStyle(target).opacity} visibility=${getComputedStyle(target).visibility}`)
  }).catch(() => undefined)
}

function runBatch(batch: LabBatch): void {
  log(`BATCH ${batch.index + 1}/${timing.batches.length} trigger=${batch.trigger} effects=${batch.effects.length} duration=${batch.duration}`)
  const groups = new Map<string, LabEffect[]>()
  for (const effect of batch.effects) {
    const key = effect.targetId ? `${effect.targetId}:${Math.round(effect.start)}` : `effect:${effect.id}`
    groups.set(key, [...(groups.get(key) ?? []), effect])
  }
  for (const effects of groups.values()) {
    schedule(() => runEffectGroup(effects), scaledMs(effects[0].start))
  }
}

async function nextEffect(): Promise<void> {
  if (advancing) return
  if (batchIndex >= timing.batches.length) {
    log("本页动画已结束")
    return
  }
  advancing = true
  const id = ++advanceId
  try {
    do {
      const batch = timing.batches[batchIndex]
      runBatch(batch)
      batchIndex += 1
      updateSummary()
      await waitFor(scaledMs(Math.max(20, batch.duration + 20)))
      if (id !== advanceId) return
    } while (batchIndex < timing.batches.length && timing.batches[batchIndex].trigger !== "onClick")
  } finally {
    if (id === advanceId) advancing = false
  }
}

async function playSlide(): Promise<void> {
  resetSlide()
  for (const batch of timing.batches) {
    runBatch(batch)
    const wait = Math.max(250, batch.duration)
    await waitFor(scaledMs(wait + 120))
  }
  batchIndex = timing.batches.length
  updateSummary()
}

function pause(): void {
  runningAnimations.forEach((animation) => animation.pause())
  timersPaused = true
  const now = performance.now()
  timerJobs.filter((job) => job.active).forEach((job) => {
    window.clearTimeout(job.handle)
    job.remaining = Math.max(0, job.remaining - (now - job.startedAt))
  })
  log("PAUSE 动画与待执行计时器")
}

function resume(): void {
  runningAnimations.forEach((animation) => animation.play())
  timersPaused = false
  timerJobs.filter((job) => job.active).forEach((job) => {
    job.startedAt = performance.now()
    job.handle = window.setTimeout(() => { job.active = false; job.callback() }, job.remaining)
  })
  log("RESUME 动画与待执行计时器")
}

function resetSlide(): void {
  clearPlayback()
  firedBookmarks.clear()
  stage.querySelectorAll<HTMLMediaElement>("video, audio").forEach((media) => {
    media.pause()
    if (Number.isFinite(media.duration)) media.currentTime = 0
  })
  batchIndex = 0
  prepareInitialState()
  log("RESET")
  updateSummary()
}

function updateSummary(): void {
  const mapped = timing.effects.filter((effect) => targetMap.has(effect.targetId)).length
  const classes = Array.from(new Set(timing.effects.map((effect) => effect.presetClass))).join(", ") || "无"
  const kinds = Array.from(new Set(timing.effects.map((effect) => effect.kind))).join(", ") || "无"
  const triggers = Array.from(new Set(timing.effects.map((effect) => effect.trigger))).join(", ") || "无"
  summary.textContent = [
    `页面：${slideIndex + 1} / ${presentation?.slides.length ?? 0}`,
    `单击批次：${timing.batches.length}`,
    `动画节点：${timing.effects.length}`,
    `目标对应：${mapped} / ${timing.effects.length}`,
    `当前批次：${Math.min(batchIndex + 1, timing.batches.length || 0)}`,
    `类别：${classes}`,
    `效果：${kinds}`,
    `开始方式：${triggers}`,
    `未知类别：${timing.unknownPresetClasses.join(", ") || "无"}`,
    `交互序列：${timing.interactive.length}`,
    `媒体书签：${timing.mediaBookmarks.map((item) => `${item.name}@${item.time}`).join(", ") || "无"}`,
    `书签触发：${timing.bookmarkSequences.length}`,
    `构建：${timing.builds.map((item) => `${item.kind}:${item.targetId}:${item.mode}`).join(", ") || "无"}`,
    `音频提示：${timing.audioCues.map((item) => item.name || item.relationId).join(", ") || "无"}`,
    `页面切换：${timing.transition ? `${timing.transition.kind} ${timing.transition.direction}` : "无"}`,
    `自动换页：${timing.transition?.advanceMs ?? "无"}`,
    `时间树：节点 ${timing.complexity.timeNodeCount}，深度 ${timing.complexity.maxDepth}，事件 ${timing.complexity.conditionEvents.join("/") || "无"}`,
    `重复/往返/缓动：${timing.complexity.repeatNodeCount}/${timing.complexity.autoReverseNodeCount}/${timing.complexity.accelerationNodeCount}`,
    `并行动画冲突：${timing.conflicts.length}`,
    `节点树对应：${timing.treeValidation.matchedEffectIds.length}/${timing.treeValidation.parsedEffectIds.length}，顺序 ${timing.treeValidation.orderMatches ? "一致" : "不一致"}`,
    `节点树事件：${timing.treeValidation.conditionEvents.join("/") || "无"}`,
  ].join("\n")
}

function transitionFrames(transition: LabTiming["transition"], entering: boolean): Keyframe[] {
  const threeDimensional = new Set(["conveyor", "doors", "ferris", "flash", "flip", "flythrough", "gallery", "glitter", "honeycomb", "orbit", "pan", "prism", "reveal", "ripple", "shred", "switch", "vortex", "warp"])
  if (transition && threeDimensional.has(transition.kind)) {
    const direction = transition.direction === "r" ? 1 : -1
    return entering
      ? [{ opacity: 0, rotate: `y ${direction * 82}deg`, scale: 0.62, filter: "blur(8px)" }, { opacity: 1, rotate: "y 0deg", scale: 1, filter: "blur(0)" }]
      : [{ opacity: 1, rotate: "y 0deg", scale: 1 }, { opacity: 0.2, rotate: `y ${direction * -35}deg`, scale: 0.82 }]
  }
  if (transition?.kind === "push") {
    const horizontal = transition.direction !== "u" && transition.direction !== "d"
    const sign = transition.direction === "r" || transition.direction === "d" ? 1 : -1
    const from = horizontal ? `${sign * 100}% 0` : `0 ${sign * 100}%`
    return entering ? [{ translate: from }, { translate: "0 0" }] : [{ opacity: 1 }, { opacity: 0.5 }]
  }
  if (transition?.kind === "wipe") {
    return entering ? [{ clipPath: "inset(0 100% 0 0)" }, { clipPath: "inset(0 0 0 0)" }] : [{ opacity: 1 }, { opacity: 0.8 }]
  }
  if (transition?.kind === "cut") return [{ opacity: 1 }, { opacity: 1 }]
  return entering ? [{ opacity: 0 }, { opacity: 1 }] : [{ opacity: 1 }, { opacity: 0 }]
}

interface ObjectIdentity {
  targetId: string
  explicitName: string
  creationId: string
  name: string
  nodeType: string
  text: string
}

interface MorphSnapshot {
  identity: ObjectIdentity
  rect: { left: number; top: number; width: number; height: number }
}

function normalizedText(target: HTMLElement): string {
  const value = (target.textContent ?? "").replace(/\s+/g, " ").trim()
  return value.length <= 240 ? value : ""
}

function parseObjectIdentities(xml: string): Map<string, Omit<ObjectIdentity, "nodeType" | "text">> {
  const doc = new DOMParser().parseFromString(xml, "application/xml")
  const tree = xmlElements(doc, "spTree")[0]
  const identities = new Map<string, Omit<ObjectIdentity, "nodeType" | "text">>()
  if (!tree) return identities
  Array.from(tree.children).filter((node) => ["sp", "pic", "graphicFrame", "grpSp", "cxnSp"].includes(node.localName)).forEach((node) => {
    const properties = xmlElements(node, "cNvPr")[0]
    const targetId = properties?.getAttribute("id") ?? ""
    const name = properties?.getAttribute("name")?.trim() ?? ""
    const creation = properties ? xmlElements(properties, "creationId").find((item) => item.hasAttribute("id")) : undefined
    if (!targetId) return
    identities.set(targetId, {
      targetId,
      explicitName: name.startsWith("!!") ? name : "",
      creationId: creation?.getAttribute("id") ?? "",
      name,
    })
  })
  return identities
}

function identityFor(node: BaseNodeData, target: HTMLElement, raw: Map<string, Omit<ObjectIdentity, "nodeType" | "text">>): ObjectIdentity {
  const identity = raw.get(node.id)
  const name = identity?.name || (node.name ?? "").trim()
  return {
    targetId: node.id,
    explicitName: identity?.explicitName || (name.startsWith("!!") ? name : ""),
    creationId: identity?.creationId ?? "",
    name,
    nodeType: node.nodeType,
    text: normalizedText(target),
  }
}

function identityScore(previous: MorphSnapshot, current: ObjectIdentity, rect: DOMRect): { score: number; method: string } {
  const old = previous.identity
  let score = 0
  let method = ""
  if (old.explicitName && old.explicitName === current.explicitName) { score = 1000; method = "explicit" }
  else if (old.creationId && old.creationId === current.creationId) { score = 900; method = "creation" }
  else if (old.text && old.text === current.text) { score = 600; method = "text" }
  else if (old.name && old.name === current.name) { score = 400; method = "name" }
  if (!score) return { score: 0, method: "" }
  if (old.nodeType === current.nodeType) score += 60
  const dx = previous.rect.left + previous.rect.width / 2 - (rect.left + rect.width / 2)
  const dy = previous.rect.top + previous.rect.height / 2 - (rect.top + rect.height / 2)
  const distance = Math.hypot(dx, dy)
  score += Math.max(0, 40 - distance / 20)
  const widthRatio = rect.width && previous.rect.width ? Math.min(rect.width, previous.rect.width) / Math.max(rect.width, previous.rect.width) : 0
  const heightRatio = rect.height && previous.rect.height ? Math.min(rect.height, previous.rect.height) / Math.max(rect.height, previous.rect.height) : 0
  score += (widthRatio + heightRatio) * 20
  return { score, method }
}

function captureMorphSnapshots(): MorphSnapshot[] {
  if (!presentation || renderedSlideIndex < 0) return []
  const slide = presentation.slides[renderedSlideIndex]
  const raw = parseObjectIdentities(renderedSlideXml)
  return slide.nodes.flatMap((node): MorphSnapshot[] => {
    const target = targetMap.get(node.id)
    if (!target) return []
    const rect = target.getBoundingClientRect()
    return [{ identity: identityFor(node, target, raw), rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height } }]
  })
}

function applyMorphTransition(snapshots: MorphSnapshot[], slide: SlideData, xml: string): void {
  const raw = parseObjectIdentities(xml)
  const used = new Set<MorphSnapshot>()
  const methods = new Map<string, number>()
  let matched = 0
  let unmatched = 0
  for (const node of slide.nodes) {
    const target = targetMap.get(node.id)
    if (!target) continue
    const rect = target.getBoundingClientRect()
    const identity = identityFor(node, target, raw)
    const candidates = snapshots.filter((candidate) => !used.has(candidate)).map((candidate) => ({ candidate, ...identityScore(candidate, identity, rect) })).filter((item) => item.score >= 400).sort((left, right) => right.score - left.score)
    const best = candidates[0]
    if (!best) {
      unmatched += 1
      const animation = target.animate([{ opacity: 0 }, { opacity: 1 }], { duration: timing.transition?.duration ?? 1000, fill: "both", easing: "ease" })
      runningAnimations.push(animation)
      continue
    }
    const snapshot = best.candidate
    used.add(snapshot)
    matched += 1
    methods.set(best.method, (methods.get(best.method) ?? 0) + 1)
    const dx = (snapshot.rect.left - rect.left) / Math.max(slideScale, 0.001)
    const dy = (snapshot.rect.top - rect.top) / Math.max(slideScale, 0.001)
    const sx = rect.width ? snapshot.rect.width / rect.width : 1
    const sy = rect.height ? snapshot.rect.height / rect.height : 1
    target.style.transformOrigin = "center center"
    const animation = target.animate(
      [{ translate: `${dx}px ${dy}px`, scale: `${sx} ${sy}`, opacity: 0.75 }, { translate: "0 0", scale: "1 1", opacity: 1 }],
      { duration: timing.transition?.duration ?? 1000, fill: "both", easing: "ease-in-out" },
    )
    runningAnimations.push(animation)
  }
  log(`MORPH option=${timing.transition?.option || "-"} matched=${matched} unmatched=${unmatched} previous=${snapshots.length} methods=${Array.from(methods).map(([method, count]) => `${method}:${count}`).join("|") || "-"}`)
}

function attachInteractiveTriggers(): void {
  for (const sequence of timing.interactive) {
    const trigger = targetMap.get(sequence.triggerTargetId)
    if (!trigger) { log(`INTERACTIVE MISS trigger=${sequence.triggerTargetId}`); continue }
    trigger.style.cursor = "pointer"
    const handler = (event: Event) => {
      event.preventDefault()
      event.stopPropagation()
      log(`INTERACTIVE trigger=${sequence.triggerTargetId} batches=${sequence.batches.length}`)
      sequence.batches.forEach(runBatch)
    }
    trigger.addEventListener("click", handler)
    interactionCleanups.push(() => trigger.removeEventListener("click", handler))
  }
}

function attachBookmarkTriggers(): void {
  for (const bookmark of timing.mediaBookmarks) {
    const media = targetMap.get(bookmark.targetId)?.querySelector<HTMLMediaElement>("video, audio")
    if (!media) { log(`BOOKMARK MISS target=${bookmark.targetId} name=${bookmark.name}`); continue }
    const key = `${bookmark.targetId}:${bookmark.name}`
    const handler = () => {
      if (firedBookmarks.has(key) || media.currentTime * 1000 < bookmark.time) return
      firedBookmarks.add(key)
      const sequences = timing.bookmarkSequences.filter((item) => item.targetId === bookmark.targetId && item.bookmarkName === bookmark.name)
      log(`BOOKMARK HIT target=${bookmark.targetId} name=${bookmark.name} time=${Math.round(media.currentTime * 1000)} sequences=${sequences.length}`)
      sequences.forEach((sequence) => sequence.batches.forEach(runBatch))
    }
    media.addEventListener("timeupdate", handler)
    interactionCleanups.push(() => media.removeEventListener("timeupdate", handler))
  }
}

function inspectConflicts(): void {
  timing.conflicts.forEach((conflict) => {
    log(`CONFLICT target=${conflict.targetId} batch=${conflict.batchIndex + 1} effects=${conflict.effectIds.join("+")} properties=${conflict.properties.join("|")}`)
  })
}

function inspectBuildTargets(): void {
  for (const build of timing.builds.filter((item) => item.kind !== "paragraph")) {
    const target = targetMap.get(build.targetId)
    if (!target) { log(`BUILD MISS kind=${build.kind} target=${build.targetId}`); continue }
    const internal = target.querySelectorAll("canvas, svg, [data-pptx-node-id]").length
    log(`BUILD INSPECT kind=${build.kind} target=${build.targetId} mode=${build.mode} internalTargets=${internal}`)
    if (build.kind === "smartart") {
      const children = Array.from(target.children).filter((node): node is HTMLElement => node instanceof HTMLElement)
      children.forEach((child, index) => {
        child.animate([{ opacity: 0 }, { opacity: 1 }], { delay: index * 120, duration: 300, fill: "both" })
      })
      log(`BUILD APPROX smartart target=${build.targetId} domOrderChildren=${children.length}`)
    }
  }
}

function scheduleAutoAdvance(): void {
  const delay = timing.transition?.advanceMs
  if (!autoAdvance.checked || delay === null || delay === undefined) return
  const timelineDelay = timing.batches.every((batch) => batch.trigger !== "onClick")
    ? timing.batches.reduce((total, batch) => total + batch.duration, 0)
    : 0
  const total = timelineDelay + delay
  log(`AUTO ADVANCE scheduled=${total} timeline=${timelineDelay} advance=${delay}`)
  schedule(() => void moveSlide(1), scaledMs(total))
}

async function renderCurrentSlide(fade = false): Promise<void> {
  if (!presentation || !files) return
  const renderStarted = performance.now()
  const slide = presentation.slides[slideIndex]
  const xml = files.slides.get(slide.slidePath) ?? slide.sourceXml ?? ""
  const nextTiming = parseSlideTiming(xml)
  const morphSnapshots = fade && nextTiming.transition?.kind === "morph" ? captureMorphSnapshots() : []
  clearPlayback(true)
  if (fade && slideHandle && nextTiming.transition?.kind !== "morph") {
    await slideHandle.element.animate(transitionFrames(nextTiming.transition, false), { duration: Math.min(nextTiming.transition?.duration ?? 180, 500) }).finished.catch(() => undefined)
  }
  slideHandle?.dispose()
  stage.replaceChildren()
  timing = nextTiming
  slideHandle = renderSlide(presentation, slide, {
    onNodeError: (nodeId, error) => log(`RENDER ERROR node=${nodeId} ${String(error)}`),
  })
  stage.appendChild(slideHandle.element)
  await slideHandle.ready
  targetMap = mapTopLevelTargets(slide, slideHandle.element)
  await hydrateMediaTargets(xml, slide)
  batchIndex = 0
  timingXml.textContent = timing.timingXml || "本页没有 p:timing"
  transitionXml.textContent = timing.transitionXml || "本页没有 p:transition"
  slideCount.textContent = `${slideIndex + 1} / ${presentation.slides.length}`
  slideNumber.value = String(slideIndex + 1)
  slideNumber.max = String(presentation.slides.length)
  prepareInitialState()
  attachInteractiveTriggers()
  attachBookmarkTriggers()
  inspectBuildTargets()
  inspectConflicts()
  fitSlide()
  updateSummary()
  status.textContent = `已渲染第 ${slideIndex + 1} 页：${targetMap.size} 个顶层对象，${timing.effects.length} 个动画节点`
  log(`SLIDE ${slideIndex + 1} path=${slide.slidePath} nodes=${slide.nodes.length} effects=${timing.effects.length}`)
  log(`TREE VERIFY matched=${timing.treeValidation.matchedEffectIds.length}/${timing.treeValidation.parsedEffectIds.length} order=${timing.treeValidation.orderMatches} missing=${timing.treeValidation.missingEffectIds.join("|") || "-"} events=${timing.treeValidation.conditionEvents.join("|") || "-"}`)
  if (fade && timing.transition?.kind === "morph") applyMorphTransition(morphSnapshots, slide, xml)
  else if (fade) slideHandle.element.animate(transitionFrames(timing.transition, true), { duration: Math.min(timing.transition?.duration ?? 180, 800), fill: "both" })
  renderedSlideIndex = slideIndex
  renderedSlideXml = xml
  log(`PERF render=${Math.round(performance.now() - renderStarted)}ms slide=${slideIndex + 1}`)
  if (timing.batches[0] && timing.batches[0].trigger !== "onClick") {
    window.setTimeout(() => void nextEffect(), 0)
  }
  scheduleAutoAdvance()
}

async function openFile(file: File): Promise<void> {
  const openStarted = performance.now()
  status.textContent = `正在解析 ${file.name}…`
  traceLines = []
  trace.textContent = ""
  slideHandle?.dispose()
  slideHandle = null
  const buffer = await file.arrayBuffer()
  const parseStarted = performance.now()
  files = await parseZipLazyMedia(buffer, RECOMMENDED_ZIP_LIMITS)
  const parsedAt = performance.now()
  presentation = buildPresentation(files, { lazySlides: true })
  const builtAt = performance.now()
  slideIndex = 0
  renderedSlideIndex = -1
  log(`OPEN ${file.name} slides=${presentation.slides.length}`)
  await renderCurrentSlide()
  log(`PERF open parseZip=${Math.round(parsedAt - parseStarted)}ms build=${Math.round(builtAt - parsedAt)}ms firstRender=${Math.round(performance.now() - builtAt)}ms total=${Math.round(performance.now() - openStarted)}ms bytes=${buffer.byteLength}`)
}

async function moveSlide(delta: number): Promise<void> {
  if (!presentation) return
  const next = Math.min(Math.max(slideIndex + delta, 0), presentation.slides.length - 1)
  if (next === slideIndex) return
  slideIndex = next
  await renderCurrentSlide(true)
}

async function goToSlide(): Promise<void> {
  if (!presentation) return
  const requested = Number(slideNumber.value)
  if (!Number.isFinite(requested)) return
  const next = Math.min(Math.max(Math.trunc(requested) - 1, 0), presentation.slides.length - 1)
  if (next === slideIndex) return
  slideIndex = next
  await renderCurrentSlide(true)
}

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0]
  if (file) void openFile(file).catch((error) => {
    status.textContent = `失败：${error instanceof Error ? error.message : String(error)}`
    log(`OPEN ERROR ${String(error)}`)
  })
})
byId("prev-slide").addEventListener("click", () => void moveSlide(-1))
byId("next-slide").addEventListener("click", () => void moveSlide(1))
byId("go-slide").addEventListener("click", () => void goToSlide())
byId("next-effect").addEventListener("click", () => void nextEffect())
byId("play-slide").addEventListener("click", () => void playSlide())
byId("pause").addEventListener("click", pause)
byId("resume").addEventListener("click", resume)
byId("reset").addEventListener("click", resetSlide)
byId("play-audio").addEventListener("click", () => {
  const cue = timing.audioCues.find((item) => audioTargets.has(item.relationId))
  const audio = cue ? audioTargets.get(cue.relationId) : undefined
  if (!cue || !audio) { log("AUDIO 本页没有可播放的音频"); return }
  void audio.play().then(() => log(`AUDIO PLAY rel=${cue.relationId} name=${cue.name || "-"}`))
    .catch((error) => log(`AUDIO BLOCKED ${String(error)}`))
})
showTargets.addEventListener("change", () => root.classList.toggle("show-targets", showTargets.checked))
autoAdvance.addEventListener("change", scheduleAutoAdvance)
window.addEventListener("resize", fitSlide)
window.addEventListener("keydown", (event) => {
  if (event.code === "Space") { event.preventDefault(); void nextEffect() }
  if (event.key === "ArrowRight") void moveSlide(1)
  if (event.key === "ArrowLeft") void moveSlide(-1)
})

window.addEventListener("beforeunload", () => {
  clearPlayback(true)
  slideHandle?.dispose()
})
