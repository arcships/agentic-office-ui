import type { PptxFiles, PresentationData, SlideData } from "@aiden0z/pptx-renderer"

import { createPptxCapabilityReport } from "../playback/capability"
import { createPptxObjectKey, matchPptxMorphObjects } from "../playback/identity"
import type {
  PptxFeatureDisposition,
  PptxFeatureRecord,
  PptxMediaItem,
  PptxObjectIdentity,
  PptxPlaybackDocument,
  PptxPlaybackEffect,
  PptxPlaybackSlide,
  PptxSlideAction,
  PptxSlideTransition,
  PptxTimeCondition,
  PptxTimeContainer,
  PptxTimeNode,
  PptxTriggerEvent,
} from "../playback/types"
import type { PptxApproximationPolicy } from "../browser-types"

export interface PptxPlaybackParserOptions {
  approximation?: PptxApproximationPolicy
  externalMedia?: "disabled" | "allowed"
}

const objectNodeNames = new Set(["sp", "pic", "graphicFrame", "grpSp", "cxnSp"])

function elements(root: Document | Element, localName: string): Element[] {
  return Array.from(root.getElementsByTagNameNS("*", localName))
}

function directChild(root: Element, localName: string): Element | undefined {
  return Array.from(root.children).find((node) => node.localName === localName)
}

function localAttr(node: Element | undefined, name: string): string | undefined {
  if (!node) return undefined
  return Array.from(node.attributes).find((attribute) => attribute.localName === name)?.value
}

function parseXml(xml: string, slideIndex: number): Document {
  if (/<!DOCTYPE/i.test(xml)) throw new Error(`第 ${slideIndex + 1} 页包含不允许的 DOCTYPE。`)
  const document = new DOMParser().parseFromString(xml, "application/xml")
  const error = document.querySelector("parsererror")
  if (error) throw new Error(`第 ${slideIndex + 1} 页 XML 解析失败：${error.textContent ?? "未知错误"}`)
  return document
}

function timeValue(value: string | undefined, fallback = 0): number | "indefinite" {
  if (value === "indefinite") return "indefinite"
  const number = Number(value)
  return Number.isFinite(number) && number >= 0 ? number : fallback
}

function ratioValue(value: string | undefined): number {
  const number = Number(value)
  return Number.isFinite(number) ? number / 100000 : 0
}

function objectType(node: Element): string {
  if (node.localName === "sp" || node.localName === "cxnSp") return "shape"
  if (node.localName === "pic") return "picture"
  if (node.localName === "grpSp") return "group"
  if (elements(node, "tbl").length) return "table"
  if (elements(node, "chart").length) return "chart"
  return "unknown"
}

function nonVisualProperties(node: Element): Element | undefined {
  const container = Array.from(node.children).find((child) => [
    "nvSpPr",
    "nvPicPr",
    "nvGrpSpPr",
    "nvGraphicFramePr",
    "nvCxnSpPr",
  ].includes(child.localName))
  return container ? elements(container, "cNvPr")[0] : undefined
}

function collectObjects(
  parent: Element,
  slidePath: string,
  groupPath: readonly string[],
  result: PptxObjectIdentity[],
): void {
  for (const node of Array.from(parent.children)) {
    if (!objectNodeNames.has(node.localName)) continue
    const properties = nonVisualProperties(node)
    const shapeId = properties?.getAttribute("id")?.trim()
    if (!shapeId) continue
    const name = properties?.getAttribute("name")?.trim() || undefined
    const creation = properties ? elements(properties.parentElement ?? properties, "creationId")[0] : undefined
    const identity: PptxObjectIdentity = Object.freeze({
      key: createPptxObjectKey({ slidePath, source: "slide", shapeId, groupPath }),
      slidePath,
      source: "slide",
      shapeId,
      groupPath: Object.freeze([...groupPath]),
      name,
      explicitMorphName: name?.startsWith("!!") ? name : undefined,
      creationId: localAttr(creation, "id") ?? localAttr(creation, "val"),
      nodeType: objectType(node),
    })
    result.push(identity)
    if (node.localName === "grpSp") collectObjects(node, slidePath, [...groupPath, shapeId], result)
  }
}

function parseObjects(document: Document, slidePath: string): readonly PptxObjectIdentity[] {
  const tree = elements(document, "spTree")[0]
  if (!tree) return Object.freeze([])
  const result: PptxObjectIdentity[] = []
  collectObjects(tree, slidePath, [], result)
  return Object.freeze(result)
}

function triggerEvent(value: string | undefined): PptxTriggerEvent {
  switch (value) {
    case "onClick": return "on-click"
    case "withEffect": return "with-previous"
    case "afterEffect": return "after-previous"
    case "onBegin": return "on-begin"
    case "onEnd": return "on-end"
    case "onNext": return "on-next"
    case "onPrev": return "on-previous"
    case "onMediaBookmark": return "on-media-bookmark"
    case "onStopAudio": return "on-stop-audio"
    case undefined:
    case "": return "delay"
    default: return "unknown"
  }
}

function nodeTypeTrigger(value: string | undefined): PptxTriggerEvent | undefined {
  if (value === "clickEffect") return "on-click"
  if (value === "withEffect") return "with-previous"
  if (value === "afterEffect") return "after-previous"
  return undefined
}

function findShapeTarget(root: Element): string | undefined {
  return elements(root, "spTgt")[0]?.getAttribute("spid") ?? undefined
}

function effectOwner(timeNode: Element): Element | undefined {
  const behaviorNames = new Set(["anim", "animEffect", "animMotion", "animRot", "animScale", "cmd", "set"])
  let current = timeNode.parentElement
  while (current && current.localName !== "cTn" && current.localName !== "childTnLst") {
    if (behaviorNames.has(current.localName)) return current
    current = current.parentElement
  }
  return undefined
}

function effectTargetId(timeNode: Element): string | undefined {
  return findShapeTarget(effectOwner(timeNode) ?? timeNode)
}

function parseConditionList(
  timeNode: Element,
  listName: string,
  source: PptxTimeCondition["source"],
  objectKeys: ReadonlyMap<string, string>,
): PptxTimeCondition[] {
  const list = directChild(timeNode, listName)
  if (!list) return []
  return Array.from(list.children)
    .filter((condition) => condition.localName === "cond")
    .map((condition) => {
      const bookmark = elements(condition, "bmkTgt")[0]
      const targetId = findShapeTarget(condition) ?? bookmark?.getAttribute("spid") ?? undefined
      const rawEvent = condition.getAttribute("evt") ?? undefined
      const parsedEvent = triggerEvent(rawEvent)
      return Object.freeze({
        source,
        event: parsedEvent === "on-click" && targetId ? "on-shape-click" : parsedEvent,
        delayMs: (() => {
          const value = timeValue(condition.getAttribute("delay") ?? undefined)
          return value === "indefinite" ? 0 : value
        })(),
        targetObjectKey: targetId ? objectKeys.get(targetId) : undefined,
        targetNodeId: elements(condition, "tn")[0]?.getAttribute("val") ?? undefined,
        bookmarkName: bookmark?.getAttribute("bmkName") ?? undefined,
        rawEvent,
      })
    })
}

function timeContainer(node: Element): PptxTimeContainer {
  switch (node.parentElement?.localName) {
    case "seq": return "sequence"
    case "par": return "parallel"
    case "excl": return "exclusive"
    case "timing": return "root"
    default: return node.hasAttribute("presetClass") ? "behavior" : "unknown"
  }
}

function directTimeChildren(node: Element): Element[] {
  const list = directChild(node, "childTnLst")
  if (!list) return []
  return elements(list, "cTn").filter((candidate) => {
    let parent = candidate.parentElement
    while (parent && parent !== list) {
      if (parent.localName === "cTn") return false
      parent = parent.parentElement
    }
    return parent === list
  })
}

function effectKind(presetClass: string, transition: string, filter: string, node: Element): PptxPlaybackEffect["kind"] {
  if (presetClass === "mediacall" || node.localName === "cmd" || elements(node, "cmd").length) return "media-command"
  if (presetClass === "path" || node.localName === "animMotion" || elements(node, "animMotion").length) return "motion-path"
  if (node.localName === "animScale" || elements(node, "animScale").length) return "scale"
  if (node.localName === "animRot" || elements(node, "animRot").length) return "rotate"
  if (presetClass === "emph") return "emphasis"
  if (presetClass === "entr") {
    if (filter.includes("fade") && transition === "in") return "fade-in"
    if (filter.includes("wipe")) return "wipe"
    return node.getAttribute("presetID") === "1" ? "appear" : "unknown"
  }
  if (presetClass === "exit") {
    if (filter.includes("fade") && transition === "out") return "fade-out"
    return node.getAttribute("presetID") === "1" ? "disappear" : "unknown"
  }
  if (node.localName === "set" || elements(node, "set").length) return "set"
  return "unknown"
}

function parseEffect(
  node: Element,
  objectKeys: ReadonlyMap<string, string>,
  slideWidth: number,
  slideHeight: number,
): PptxPlaybackEffect {
  const owner = effectOwner(node) ?? node
  const animEffect = owner.localName === "animEffect" ? owner : elements(owner, "animEffect")[0]
  const filter = animEffect?.getAttribute("filter") ?? ""
  const transition = animEffect?.getAttribute("transition") ?? ""
  const presetClass = node.getAttribute("presetClass") ?? ""
  const targetId = findShapeTarget(owner)
  const range = elements(owner, "pRg")[0]
  const motion = owner.localName === "animMotion" ? owner : elements(owner, "animMotion")[0]
  const rotation = owner.localName === "animRot" ? owner : elements(owner, "animRot")[0]
  const scale = owner.localName === "animScale" ? owner : elements(owner, "animScale")[0]
  const animation = owner.localName === "anim" ? owner : elements(owner, "anim")[0]
  const values: Record<string, string | number | boolean> = {}
  if (filter) values.filter = filter
  const path = motion?.getAttribute("path") ?? ""
  if (path) {
    values.motionPath = path
    values.motionScaleX = slideWidth
    values.motionScaleY = slideHeight
  }
  const rotationBy = Number(rotation?.getAttribute("by"))
  if (Number.isFinite(rotationBy)) values.rotationBy = rotationBy / 60000
  for (const point of ["by", "from", "to"] as const) {
    const element = scale ? directChild(scale, point) : undefined
    const x = Number(element?.getAttribute("x"))
    const y = Number(element?.getAttribute("y"))
    if (Number.isFinite(x) && Number.isFinite(y)) {
      values[`scale${point[0].toUpperCase()}${point.slice(1)}X`] = x / 100000
      values[`scale${point[0].toUpperCase()}${point.slice(1)}Y`] = y / 100000
    }
  }
  const attributeNames = animation ? elements(animation, "attrName").map((item) => item.textContent?.trim() ?? "") : []
  if (attributeNames.some((name) => /opacity|transparency/iu.test(name))) {
    const normalizeOpacity = (value: number): number => Math.max(0, Math.min(1, Math.abs(value) > 1 ? value / 100000 : value))
    const floatValues = elements(animation as Element, "fltVal")
      .map((item) => Number(item.getAttribute("val")))
      .filter(Number.isFinite)
    if (floatValues.length) {
      values.emphasisProperty = "opacity"
      values.emphasisFrom = normalizeOpacity(floatValues[0])
      values.emphasisTo = normalizeOpacity(floatValues[floatValues.length - 1])
    }
  }
  const command = (owner.localName === "cmd" ? owner : elements(owner, "cmd")[0])?.getAttribute("cmd") ?? undefined
  return Object.freeze({
    id: node.getAttribute("id") ?? "",
    kind: effectKind(presetClass, transition, filter, owner),
    targetObjectKey: targetId ? objectKeys.get(targetId) : undefined,
    paragraphRange: range ? {
      start: Number(range.getAttribute("st") ?? range.getAttribute("start") ?? 0),
      end: Number(range.getAttribute("end") ?? range.getAttribute("st") ?? 0),
    } : undefined,
    presetClass: presetClass || undefined,
    presetId: node.getAttribute("presetID") ?? undefined,
    transition: transition === "in" || transition === "out" ? transition : undefined,
    filter: filter || undefined,
    command,
    motionPath: path || undefined,
    values: Object.freeze(values),
  })
}

function effectDisposition(effect: PptxPlaybackEffect, targetId: string | undefined): PptxFeatureDisposition {
  if (!effect.targetObjectKey) return effect.kind === "unknown" && !targetId ? "unparsed" : "static"
  if (["appear", "disappear", "fade-in", "fade-out", "scale", "rotate", "media-command"].includes(effect.kind)) return "strict"
  if (effect.kind === "wipe") return /left|right/i.test(effect.filter ?? "") ? "strict" : "static"
  if (effect.kind === "motion-path") return /^[\s\d.,+\-mlehvz]+$/iu.test(effect.motionPath ?? "") ? "strict" : "static"
  if (effect.kind === "emphasis") return effect.values.emphasisProperty === "opacity" ? "strict" : "static"
  if (effect.kind === "set") return "static"
  return "unparsed"
}

function parseTimeNode(
  node: Element,
  slideIndex: number,
  objectKeys: ReadonlyMap<string, string>,
  nodes: Record<string, PptxTimeNode>,
  features: PptxFeatureRecord[],
  slideWidth: number,
  slideHeight: number,
  parentId?: string,
): string {
  const id = node.getAttribute("id") ?? `generated-${slideIndex}-${Object.keys(nodes).length}`
  const conditions = [
    ...parseConditionList(node, "stCondLst", "start", objectKeys),
    ...parseConditionList(node, "endCondLst", "end", objectKeys),
    ...parseConditionList(node, "prevCondLst", "previous", objectKeys),
    ...parseConditionList(node, "nextCondLst", "next", objectKeys),
  ]
  const fallbackEvent = nodeTypeTrigger(node.getAttribute("nodeType") ?? undefined)
  if (fallbackEvent && !conditions.some((condition) => condition.event === fallbackEvent)) {
    conditions.push(Object.freeze({ source: "start", event: fallbackEvent, delayMs: 0 }))
  }
  const effect = node.hasAttribute("presetClass")
    ? parseEffect(node, objectKeys, slideWidth, slideHeight)
    : undefined
  const childIds = directTimeChildren(node).map((child) => parseTimeNode(
    child,
    slideIndex,
    objectKeys,
    nodes,
    features,
    slideWidth,
    slideHeight,
    id,
  ))
  let rawDuration = timeValue(localAttr(node, "dur"), 0)
  if (effect && rawDuration === 0) {
    const behaviorDurations = elements(node, "cTn").flatMap((child): number[] => {
      const value = timeValue(localAttr(child, "dur"), 0)
      return typeof value === "number" ? [value] : []
    })
    rawDuration = Math.max(0, ...behaviorDurations)
  }
  const rawRepeat = localAttr(node, "repeatCount")
  const repeatRawNumber = Number(rawRepeat)
  const repeatNumber = Number.isFinite(repeatRawNumber) && repeatRawNumber >= 1000
    ? repeatRawNumber / 1000
    : repeatRawNumber
  const timeNode: PptxTimeNode = Object.freeze({
    id,
    parentId,
    container: parentId ? timeContainer(node) : "root",
    kind: effect?.kind === "media-command"
      ? "media"
      : effect
        ? "effect"
        : effectOwner(node)
          ? "command"
          : childIds.length
            ? "group"
            : "unknown",
    nodeType: node.getAttribute("nodeType") ?? undefined,
    delayMs: conditions.find((condition) => condition.source === "start")?.delayMs ?? 0,
    durationMs: rawDuration,
    repeatCount: rawRepeat === "indefinite" ? "indefinite" : Number.isFinite(repeatNumber) && repeatNumber > 0 ? repeatNumber : undefined,
    autoReverse: localAttr(node, "autoRev") === "1",
    fill: ["hold", "remove", "freeze"].includes(localAttr(node, "fill") ?? "")
      ? localAttr(node, "fill") as "hold" | "remove" | "freeze"
      : "unknown",
    restart: localAttr(node, "restart") === "always"
      ? "always"
      : localAttr(node, "restart") === "whenNotActive"
        ? "when-not-active"
        : localAttr(node, "restart") === "never"
          ? "never"
          : "unknown",
    acceleration: ratioValue(localAttr(node, "accel")),
    deceleration: ratioValue(localAttr(node, "decel")),
    conditions: Object.freeze(conditions),
    childIds: Object.freeze(childIds),
    effect,
    rawSummary: effect?.kind === "unknown" ? node.parentElement?.localName : undefined,
  })
  nodes[id] = timeNode
  if (effect) {
    const targetId = effectTargetId(node)
    const unsupportedCondition = conditions.find((condition) =>
      ["unknown", "on-next", "on-previous", "on-stop-audio"].includes(condition.event)
      || condition.event === "on-shape-click" && !condition.targetObjectKey
      || condition.event === "on-media-bookmark" && !condition.bookmarkName
      || (condition.event === "on-begin" || condition.event === "on-end") && !condition.targetNodeId,
    )
    const baseDisposition = effectDisposition(effect, targetId)
    const disposition: PptxFeatureDisposition = unsupportedCondition
      ? unsupportedCondition.event === "unknown" ? "unparsed" : "static"
      : baseDisposition
    features.push(Object.freeze({
      id: `effect:${slideIndex}:${id}`,
      slideIndex,
      objectKey: effect.targetObjectKey,
      feature: effect.kind,
      disposition,
      reason: targetId && !effect.targetObjectKey
        ? `动画目标 ${targetId} 没有唯一对象身份。`
        : unsupportedCondition
          ? `第一版尚未执行触发条件 ${unsupportedCondition.rawEvent ?? unsupportedCondition.event}。`
          : disposition === "static" ? "第一版尚未批准该效果参数。" : undefined,
      sourceNodeId: id,
    }))
    if (timeNode.repeatCount !== undefined) features.push(Object.freeze({
      id: `repeat:${slideIndex}:${id}`,
      slideIndex,
      objectKey: effect.targetObjectKey,
      feature: timeNode.repeatCount === "indefinite" ? "repeat:indefinite" : "repeat:finite",
      disposition: disposition === "strict" && typeof timeNode.durationMs === "number" && timeNode.durationMs > 0
        ? "strict"
        : "static",
      reason: disposition === "strict" ? undefined : "基础效果本身未进入严格执行范围。",
      sourceNodeId: id,
    }))
    if (timeNode.autoReverse) features.push(Object.freeze({
      id: `auto-reverse:${slideIndex}:${id}`,
      slideIndex,
      objectKey: effect.targetObjectKey,
      feature: "auto-reverse",
      disposition: disposition === "strict" && typeof timeNode.durationMs === "number" && timeNode.durationMs > 0
        ? "strict"
        : "static",
      reason: disposition === "strict" ? undefined : "基础效果本身未进入严格执行范围。",
      sourceNodeId: id,
    }))
  }
  return id
}

function parseTransition(node: Element | undefined): PptxSlideTransition | undefined {
  if (!node) return undefined
  const effect = Array.from(node.children).find((child) => !["sndAc", "extLst"].includes(child.localName))
  const speed = node.getAttribute("spd")
  const fallback = speed === "fast" ? 500 : speed === "slow" ? 2000 : 1000
  const duration = timeValue(localAttr(node, "dur"), fallback)
  const after = timeValue(node.getAttribute("advTm") ?? undefined, 0)
  return Object.freeze({
    kind: effect?.localName ?? "cut",
    direction: effect?.getAttribute("dir") ?? effect?.getAttribute("orient") ?? undefined,
    option: effect?.getAttribute("option") ?? undefined,
    durationMs: duration === "indefinite" ? fallback : duration,
    advanceOnClick: node.getAttribute("advClick") !== "0",
    advanceAfterMs: node.hasAttribute("advTm") && after !== "indefinite" ? after : undefined,
    soundRelationId: localAttr(elements(node, "snd")[0], "embed"),
  })
}

function enclosingObjectId(node: Element): string | undefined {
  let current: Element | null = node
  while (current && !objectNodeNames.has(current.localName)) current = current.parentElement
  return current ? nonVisualProperties(current)?.getAttribute("id") ?? undefined : undefined
}

function parseMedia(
  document: Document,
  objectKeys: ReadonlyMap<string, string>,
  slide: SlideData,
): readonly PptxMediaItem[] {
  const result = new Map<string, PptxMediaItem>()
  for (const node of [...elements(document, "videoFile"), ...elements(document, "audioFile")]) {
    const mediaContainer = node.parentElement ?? node
    const mediaExtension = elements(mediaContainer, "media")[0]
    const legacyRelationId = localAttr(node, "embed") ?? localAttr(node, "link")
    const extensionRelationId = localAttr(mediaExtension, "embed") ?? localAttr(mediaExtension, "link")
    const extensionRelation = extensionRelationId ? slide.rels.get(extensionRelationId) : undefined
    const relationId = extensionRelation && extensionRelation.targetMode?.toLowerCase() !== "external"
      ? extensionRelationId
      : legacyRelationId ?? extensionRelationId
    if (!relationId) continue
    const relation = slide.rels.get(relationId)
    const external = relation?.targetMode?.toLowerCase() === "external"
    const sourcePath = relation
      ? external ? relation.target : resolveTargetPath(slide.slidePath, relation.target)
      : undefined
    const shapeId = enclosingObjectId(node)
    const trim = mediaExtension ? elements(mediaExtension, "trim")[0] : undefined
    const trimStart = Number(trim?.getAttribute("st"))
    const trimEnd = Number(trim?.getAttribute("end"))
    const bookmarks = elements(mediaContainer, "bmk").flatMap((bookmark) => {
      const name = bookmark.getAttribute("name")
      const time = Number(bookmark.getAttribute("time"))
      return name && Number.isFinite(time) ? [{ name, timeMs: time }] : []
    })
    result.set(relationId, Object.freeze({
      id: relationId,
      objectKey: shapeId ? objectKeys.get(shapeId) : undefined,
      relationId,
      sourcePath,
      contentType: sourcePath?.toLowerCase().endsWith(".mp4")
        ? "video/mp4"
        : sourcePath?.toLowerCase().endsWith(".webm")
          ? "video/webm"
          : sourcePath?.toLowerCase().endsWith(".mp3")
            ? "audio/mpeg"
            : sourcePath?.toLowerCase().endsWith(".wav") ? "audio/wav" : undefined,
      kind: node.localName === "videoFile" ? "video" : "audio",
      embedded: !external && sourcePath?.startsWith("ppt/") === true,
      trimStartMs: Number.isFinite(trimStart) && trim?.hasAttribute("st") ? trimStart : undefined,
      trimEndMs: Number.isFinite(trimEnd) && trim?.hasAttribute("end") ? trimEnd : undefined,
      loop: false,
      volume: 1,
      bookmarks: Object.freeze(bookmarks),
    }))
  }
  for (const mediaNode of elements(document, "cMediaNode")) {
    const shapeId = findShapeTarget(mediaNode)
    const objectKey = shapeId ? objectKeys.get(shapeId) : undefined
    const entry = [...result.values()].find((item) => item.objectKey === objectKey)
    if (!entry) continue
    const timing = elements(mediaNode, "cTn")[0]
    const rawVolume = Number(mediaNode.getAttribute("vol"))
    const volume = Number.isFinite(rawVolume)
      ? Math.max(0, Math.min(1, rawVolume > 1 ? rawVolume / 100000 : rawVolume))
      : entry.volume
    const attributeNumber = (names: readonly string[]): number | undefined => {
      for (const element of [mediaNode, ...elements(mediaNode, "media")]) {
        for (const name of names) {
          const value = Number(element.getAttribute(name))
          if (Number.isFinite(value) && element.hasAttribute(name)) return value
        }
      }
      return undefined
    }
    const discoveredBookmarks = elements(mediaNode, "bmk").flatMap((bookmark) => {
      const name = bookmark.getAttribute("name")
      const time = Number(bookmark.getAttribute("time"))
      return name && Number.isFinite(time) ? [{ name, timeMs: time }] : []
    })
    const bookmarks = new Map(entry.bookmarks.map((bookmark) => [bookmark.name, bookmark]))
    for (const bookmark of discoveredBookmarks) bookmarks.set(bookmark.name, bookmark)
    result.set(entry.relationId, Object.freeze({
      ...entry,
      trimStartMs: attributeNumber(["trimSt", "st"]) ?? entry.trimStartMs,
      trimEndMs: attributeNumber(["trimEnd", "end"]) ?? entry.trimEndMs,
      loop: localAttr(timing, "repeatCount") === "indefinite" || mediaNode.getAttribute("loop") === "1",
      volume,
      bookmarks: Object.freeze([...bookmarks.values()]),
    }))
  }
  return Object.freeze([...result.values()])
}

function resolveTargetPath(basePath: string, target: string): string {
  const base = basePath.slice(0, basePath.lastIndexOf("/")).split("/")
  for (const part of target.replace(/\\/gu, "/").split("/")) {
    if (!part || part === ".") continue
    if (part === "..") base.pop()
    else base.push(part)
  }
  return base.join("/")
}

function parseActions(
  document: Document,
  slide: SlideData,
  presentation: PresentationData,
  objectKeys: ReadonlyMap<string, string>,
): readonly PptxSlideAction[] {
  const actions: PptxSlideAction[] = []
  for (const node of [...elements(document, "hlinkClick"), ...elements(document, "hlinkHover")]) {
    const shapeId = enclosingObjectId(node)
    const sourceObjectKey = shapeId ? objectKeys.get(shapeId) : undefined
    if (!sourceObjectKey) continue
    const relationId = localAttr(node, "id")
    const relation = relationId ? slide.rels.get(relationId) : undefined
    const rawAction = node.getAttribute("action") ?? undefined
    let kind: PptxSlideAction["kind"] = "unsupported"
    let targetSlideIndex: number | undefined
    let url: string | undefined
    if (rawAction?.startsWith("ppaction://hlinksldjump") && relation) {
      const path = resolveTargetPath(slide.slidePath, relation.target)
      targetSlideIndex = presentation.slides.findIndex((candidate) => candidate.slidePath === path)
      if (targetSlideIndex >= 0) kind = "go-to-slide"
    } else if (rawAction?.startsWith("ppaction://hlinkshowjump")) {
      const jump = new URL(rawAction).searchParams.get("jump")?.toLowerCase()
      if (jump === "nextslide") targetSlideIndex = slide.index + 1
      else if (jump === "previousslide") targetSlideIndex = slide.index - 1
      else if (jump === "firstslide") targetSlideIndex = 0
      else if (jump === "lastslide") targetSlideIndex = presentation.slides.length - 1
      if (targetSlideIndex !== undefined && targetSlideIndex >= 0 && targetSlideIndex < presentation.slides.length) {
        kind = "go-to-slide"
      }
    } else if (relation?.targetMode?.toLowerCase() === "external") {
      url = relation.target
      kind = url.toLowerCase().startsWith("mailto:") ? "mailto" : /^https?:/iu.test(url) ? "open-url" : "unsupported"
    }
    actions.push(Object.freeze({
      id: `action:${slide.index}:${actions.length}`,
      sourceObjectKey,
      trigger: node.localName === "hlinkHover" ? "hover" : "click",
      kind,
      targetSlideIndex,
      url,
      rawAction,
    }))
  }
  return Object.freeze(actions)
}

function transitionDisposition(kind: string, approximation: PptxApproximationPolicy): PptxFeatureDisposition {
  if (["cut", "fade", "push", "wipe"].includes(kind)) return "strict"
  if (kind === "morph") return approximation === "safe" ? "approximate" : "static"
  return approximation === "safe" ? "approximate" : "static"
}

function parseSlide(
  files: PptxFiles,
  presentation: PresentationData,
  slide: SlideData,
  options: PptxPlaybackParserOptions,
): PptxPlaybackSlide {
  const xml = files.slides.get(slide.slidePath)
  if (!xml) throw new Error(`缺少页面 XML：${slide.slidePath}`)
  const document = parseXml(xml, slide.index)
  const objects = parseObjects(document, slide.slidePath)
  const objectKeys = new Map<string, string>()
  const duplicateIds = new Set<string>()
  for (const object of objects) {
    if (objectKeys.has(object.shapeId)) {
      objectKeys.delete(object.shapeId)
      duplicateIds.add(object.shapeId)
    } else if (!duplicateIds.has(object.shapeId)) objectKeys.set(object.shapeId, object.key)
  }
  const features: PptxFeatureRecord[] = []
  for (const id of duplicateIds) features.push(Object.freeze({
    id: `target:${slide.index}:${id}`,
    slideIndex: slide.index,
    feature: "object-identity",
    disposition: "unparsed",
    reason: `对象编号 ${id} 在同一页中不唯一。`,
  }))
  const nodes: Record<string, PptxTimeNode> = {}
  const root = elements(document, "cTn").find((node) => node.getAttribute("nodeType") === "tmRoot")
  const rootNodeId = root ? parseTimeNode(
    root,
    slide.index,
    objectKeys,
    nodes,
    features,
    presentation.width,
    presentation.height,
  ) : undefined
  const transitionNodes = elements(document, "transition")
  const transition = parseTransition(
    transitionNodes.find((candidate) => Array.from(candidate.children).some(
      (child) => !["sndAc", "extLst"].includes(child.localName),
    )) ?? transitionNodes[0],
  )
  if (transition) features.push(Object.freeze({
    id: `transition:${slide.index}`,
    slideIndex: slide.index,
    feature: `transition:${transition.kind}`,
    disposition: transitionDisposition(transition.kind, options.approximation ?? "off"),
  }))
  const media = parseMedia(document, objectKeys, slide)
  for (const item of media) features.push(Object.freeze({
    id: `media:${slide.index}:${item.id}`,
    slideIndex: slide.index,
    objectKey: item.objectKey,
    feature: `media:${item.kind}`,
    disposition: item.embedded || options.externalMedia === "allowed" ? "strict" : "static",
    reason: !item.embedded && options.externalMedia !== "allowed" ? "外部媒体默认禁用。" : undefined,
  }))
  for (const build of elements(document, "bldChart")) features.push(Object.freeze({
    id: `chart-build:${slide.index}:${features.length}`,
    slideIndex: slide.index,
    feature: `chart-build:${build.getAttribute("bld") ?? "unknown"}`,
    disposition: "static",
    reason: "图表内部动画尚未保留语义对象身份。",
  }))
  for (const build of elements(document, "bldDgm")) features.push(Object.freeze({
    id: `smartart-build:${slide.index}:${features.length}`,
    slideIndex: slide.index,
    feature: `smartart-build:${build.getAttribute("bld") ?? "unknown"}`,
    disposition: "static",
    reason: "SmartArt 内部动画尚未保留语义对象身份。",
  }))
  const actions = parseActions(document, slide, presentation, objectKeys)
  for (const action of actions) features.push(Object.freeze({
    id: action.id,
    slideIndex: slide.index,
    objectKey: action.sourceObjectKey,
    feature: `action:${action.kind}`,
    disposition: action.kind === "unsupported" || action.trigger === "hover" ? "static" : "strict",
    reason: action.trigger === "hover" ? "悬停动作仅保留，不自动执行。" : undefined,
  }))
  return Object.freeze({
    index: slide.index,
    hidden: slide.hidden === true || ["0", "false", "off"].includes(
      (document.documentElement.getAttribute("show") ?? "").toLowerCase(),
    ),
    objects,
    morphFromPrevious: Object.freeze([]),
    transition,
    rootNodeId,
    nodes: Object.freeze(nodes),
    media,
    actions,
    capability: createPptxCapabilityReport(features),
  })
}

export function parsePptxPlaybackDocument(
  files: PptxFiles,
  presentation: PresentationData,
  options: PptxPlaybackParserOptions = {},
): PptxPlaybackDocument {
  const parsedSlides = presentation.slides.map((slide) => parseSlide(files, presentation, slide, options))
  const slides = Object.freeze(parsedSlides.map((slide, index): PptxPlaybackSlide => {
    const matches = index > 0 && slide.transition?.kind === "morph"
      ? matchPptxMorphObjects(parsedSlides[index - 1].objects, slide.objects)
      : Object.freeze([])
    const morphFeatures = matches.map((match): PptxFeatureRecord => Object.freeze({
      id: `morph:${index}:${match.method}:${match.to}`,
      slideIndex: index,
      objectKey: match.to,
      feature: `morph:${match.method}`,
      disposition: options.approximation === "safe" ? "approximate" : "static",
      reason: options.approximation === "safe"
        ? `使用${match.method === "explicit-name" ? "显式名称" : "创建编号"}进行强身份匹配。`
        : "近似播放未开启，保留静态页面。",
    }))
    return Object.freeze({
      ...slide,
      morphFromPrevious: matches,
      capability: createPptxCapabilityReport([...slide.capability.features, ...morphFeatures]),
    })
  }))
  return Object.freeze({
    slides,
    capability: createPptxCapabilityReport(slides.flatMap((slide) => slide.capability.features)),
  })
}
