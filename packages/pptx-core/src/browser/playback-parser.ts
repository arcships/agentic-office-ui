import type { PptxFiles, PresentationData, SlideData } from "@aiden0z/pptx-renderer"

import { createPptxCapabilityReport } from "../playback/capability"
import {
  createPptxObjectGeometry,
  createPptxObjectKey,
  matchPptxMorphObjects,
  normalizePptxObjectText,
} from "../playback/identity"
import {
  PPTX_CAPABILITY_FEATURES,
  pptxEffectTargetSupport,
  pptxTextBody,
  pptxTextBodyOffset,
  pptxTextGranularity,
} from "../playback/targets"
import { pptxTransitionCapability } from "../playback/transitions"
import type {
  PptxChartElementTarget,
  PptxCharacterRange,
  PptxFeatureDisposition,
  PptxFeatureRecord,
  PptxGraphicElementTarget,
  PptxMediaItem,
  PptxMorphMatch,
  PptxObjectGeometry,
  PptxObjectHyperlink,
  PptxObjectIdentity,
  PptxPlaybackDocument,
  PptxPlaybackEffect,
  PptxPlaybackSlide,
  PptxSlideAction,
  PptxSlideTransition,
  PptxSlideTransitionSound,
  PptxTableCellTarget,
  PptxTimeCondition,
  PptxTimeContainer,
  PptxTimeNode,
  PptxTriggerEvent,
} from "../playback/types"
import type { PptxApproximationPolicy } from "../browser-types"
import { withAutoplayMedia } from "../playback/media-command"
import { collectPptxTextDirectionFeatures } from "../playback/text-direction"
import { repeatCountNote, repeatCountValue } from "../playback/timing-values"

export interface PptxPlaybackParserOptions {
  approximation?: PptxApproximationPolicy
  externalMedia?: "disabled" | "allowed"
}

const objectNodeNames = new Set(["sp", "pic", "graphicFrame", "grpSp", "cxnSp"])

interface PptxTableLayoutCell extends PptxTableCellTarget {
  paragraphStart: number
  paragraphCount: number
}

interface PptxTextBodyInfo {
  paragraphs: readonly string[]
  table?: { cells: readonly PptxTableLayoutCell[] }
}

interface ParseContext {
  slideIndex: number
  objectKeys: ReadonlyMap<string, string>
  objectsById: ReadonlyMap<string, PptxObjectIdentity>
  textBodies: ReadonlyMap<string, PptxTextBodyInfo>
  slideWidth: number
  slideHeight: number
}

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
  if (elements(node, "relIds").length || /diagram/iu.test(
    elements(node, "graphicData")[0]?.getAttribute("uri") ?? "",
  )) return "smartart"
  return "unknown"
}

function directChildren(root: Element, localName: string): Element[] {
  return Array.from(root.children).filter((child) => child.localName === localName)
}

function paragraphTexts(body: Element): string[] {
  return directChildren(body, "p").map((paragraph) => paragraph.textContent ?? "")
}

function textBodyParagraphs(node: Element): string[] | undefined {
  const body = directChild(node, "txBody")
  if (!body) return undefined
  return paragraphTexts(body)
}

/**
 * Table cells in rendered row-major order. The renderer skips merged
 * continuation cells, so the parser mirrors that grid walk before it maps a
 * `pRg` paragraph index onto a cell.
 */
function tableTextBody(node: Element): PptxTextBodyInfo | undefined {
  if (node.localName !== "graphicFrame") return undefined
  const table = elements(node, "tbl")[0]
  if (!table) return undefined
  const cells: PptxTableLayoutCell[] = []
  const paragraphs: string[] = []
  const occupied = new Set<string>()
  let paragraphStart = 0
  directChildren(table, "tr").forEach((row, rowIndex) => {
    let column = 0
    for (const cell of directChildren(row, "tc")) {
      if (localAttr(cell, "hMerge") === "1") continue
      const gridSpan = Math.max(1, Math.trunc(Number(localAttr(cell, "gridSpan") ?? 1)) || 1)
      if (localAttr(cell, "vMerge") === "1") {
        column += gridSpan
        continue
      }
      while (occupied.has(`${rowIndex}:${column}`)) column += 1
      const rowSpan = Math.max(1, Math.trunc(Number(localAttr(cell, "rowSpan") ?? 1)) || 1)
      for (let rowOffset = 0; rowOffset < rowSpan; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < gridSpan; columnOffset += 1) {
          occupied.add(`${rowIndex + rowOffset}:${column + columnOffset}`)
        }
      }
      const body = directChild(cell, "txBody")
      const cellParagraphs = body ? paragraphTexts(body) : []
      const effective = cellParagraphs.length ? cellParagraphs : [""]
      paragraphs.push(...effective)
      cells.push(Object.freeze({
        row: rowIndex,
        column,
        paragraphStart,
        paragraphCount: effective.length,
      }))
      paragraphStart += effective.length
      column += gridSpan
    }
  })
  return Object.freeze({ paragraphs: Object.freeze(paragraphs), table: Object.freeze({ cells: Object.freeze(cells) }) })
}

function textBodyInfo(node: Element): PptxTextBodyInfo | undefined {
  const table = tableTextBody(node)
  if (table) return table
  const paragraphs = textBodyParagraphs(node)
  return paragraphs ? Object.freeze({ paragraphs: Object.freeze(paragraphs) }) : undefined
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
  textBodies: Map<string, PptxTextBodyInfo>,
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
      // 组合子对象的坐标属于子坐标系，不参与跨页几何匹配。
      geometry: groupPath.length > 0 ? undefined : objectGeometry(node),
      text: objectText(node),
    })
    result.push(identity)
    const body = textBodyInfo(node)
    if (body) textBodies.set(shapeId, body)
    if (node.localName === "grpSp") collectObjects(node, slidePath, [...groupPath, shapeId], result, textBodies)
  }
}

function objectGeometry(node: Element): PptxObjectGeometry | undefined {
  const properties = directChild(node, node.localName === "grpSp" ? "grpSpPr" : "spPr")
  const transform = directChild(node, "xfrm") ?? (properties ? directChild(properties, "xfrm") : undefined)
  if (!transform) return undefined
  const offset = directChild(transform, "off")
  const extent = directChild(transform, "ext")
  return createPptxObjectGeometry({
    x: offset?.getAttribute("x"),
    y: offset?.getAttribute("y"),
    cx: extent?.getAttribute("cx"),
    cy: extent?.getAttribute("cy"),
  })
}

function objectText(node: Element): string | undefined {
  // 组合对象的文字等于子对象文字之和，不作为独立的文字身份。
  if (node.localName === "grpSp") return undefined
  return normalizePptxObjectText(elements(node, "t").map((item) => item.textContent ?? "").join(" "))
}

function parseObjects(
  document: Document,
  slidePath: string,
): { objects: readonly PptxObjectIdentity[]; textBodies: ReadonlyMap<string, PptxTextBodyInfo> } {
  const tree = elements(document, "spTree")[0]
  if (!tree) return { objects: Object.freeze([]), textBodies: Object.freeze(new Map()) }
  const result: PptxObjectIdentity[] = []
  const textBodies = new Map<string, PptxTextBodyInfo>()
  collectObjects(tree, slidePath, [], result, textBodies)
  return { objects: Object.freeze(result), textBodies }
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

function optionalNumber(element: Element | undefined, name: string): number | undefined {
  if (!element) return undefined
  const raw = element.getAttribute(name)
  if (raw === null || raw === "") return undefined
  const value = Number(raw)
  return Number.isFinite(value) ? value : undefined
}

function parseCharacterRange(node: Element | undefined): PptxCharacterRange | undefined {
  if (!node) return undefined
  const start = Number(node.getAttribute("st") ?? node.getAttribute("start") ?? 0)
  const end = Number(node.getAttribute("end") ?? node.getAttribute("st") ?? 0)
  if (!Number.isFinite(start) || !Number.isFinite(end)) return undefined
  return Object.freeze({ start: Math.max(0, Math.min(start, end)), end: Math.max(start, end) })
}

function chartElementTarget(owner: Element): PptxChartElementTarget | undefined {
  const chart = elements(owner, "oleChartEl")[0]
  if (!chart) return undefined
  const raw = (chart.getAttribute("type") ?? "chart").trim().toLowerCase()
  const type: PptxChartElementTarget["type"] = raw === "ser" || raw === "series"
    ? "series"
    : raw === "cat" || raw === "category"
      ? "category"
      : raw === "ptinser" || raw === "pointinseries"
        ? "point-in-series"
        : raw === "ptincat" || raw === "pointincategory"
          ? "point-in-category"
          : "chart"
  return Object.freeze({
    type,
    index: optionalNumber(chart, "idx"),
    level: optionalNumber(chart, "lvl"),
  })
}

function childObjectKey(
  identity: PptxObjectIdentity | undefined,
  sourceId: string | undefined,
): string | undefined {
  if (!identity || !sourceId || !/^[^|/]+$/u.test(sourceId)) return undefined
  return createPptxObjectKey({
    slidePath: identity.slidePath,
    source: identity.source,
    shapeId: sourceId,
    groupPath: [...identity.groupPath, identity.shapeId],
  })
}

function graphicElementTarget(
  owner: Element,
  identity: PptxObjectIdentity | undefined,
): PptxGraphicElementTarget | undefined {
  const subShape = elements(owner, "subSp")[0]
  if (subShape) {
    const sourceId = subShape.getAttribute("spid") ?? undefined
    return Object.freeze({
      type: "subSp",
      sourceId,
      objectKey: childObjectKey(identity, sourceId),
    })
  }
  const graphic = elements(owner, "graphicEl")[0]
  if (!graphic) return undefined
  const rawType = (graphic.getAttribute("type") ?? "graphicEl").trim() || "graphicEl"
  const sourceId = graphic.getAttribute("id")
    ?? graphic.getAttribute("spid")
    ?? undefined
  const index = optionalNumber(graphic, "idx")
  const level = optionalNumber(graphic, "lvl")
  // DrawingML diagram child ids are numeric shape ids. A non-numeric token
  // (for example `dgm`) has no rendered identity, so it stays reported only.
  const resolvedId = sourceId && /^\d+$/u.test(sourceId) ? sourceId : undefined
  return Object.freeze({
    type: rawType,
    sourceId,
    index,
    level,
    objectKey: childObjectKey(identity, resolvedId),
  })
}

function paragraphRangeFromCharacters(
  paragraphs: readonly string[],
  range: PptxCharacterRange,
): { start: number; end: number } | undefined {
  if (!paragraphs.length) return undefined
  let start: number | undefined
  let end: number | undefined
  for (let index = 0; index < paragraphs.length; index += 1) {
    const offset = pptxTextBodyOffset(paragraphs, index)
    const last = offset + paragraphs[index].length - 1
    if (range.end < offset || range.start > last) continue
    start ??= index
    end = index
  }
  return start === undefined || end === undefined ? undefined : { start, end }
}

function tableCellForRange(
  body: PptxTextBodyInfo | undefined,
  range: { start: number; end: number } | undefined,
): PptxTableCellTarget | undefined {
  if (!body?.table || !range) return undefined
  const matches = body.table.cells.filter((cell) =>
    cell.paragraphStart <= range.end
    && cell.paragraphStart + cell.paragraphCount - 1 >= range.start
  )
  return matches.length === 1
    ? Object.freeze({ row: matches[0].row, column: matches[0].column })
    : undefined
}

function parseEffect(node: Element, context: ParseContext): PptxPlaybackEffect {
  const owner = effectOwner(node) ?? node
  const animEffect = owner.localName === "animEffect" ? owner : elements(owner, "animEffect")[0]
  const filter = animEffect?.getAttribute("filter") ?? ""
  const transition = animEffect?.getAttribute("transition") ?? ""
  const presetClass = node.getAttribute("presetClass") ?? ""
  const targetId = findShapeTarget(owner)
  const targetIdentity = targetId ? context.objectsById.get(targetId) : undefined
  const targetBody = targetId ? context.textBodies.get(targetId) : undefined
  const range = elements(owner, "pRg")[0]
  const rawParagraphStart = Number(range?.getAttribute("st") ?? range?.getAttribute("start") ?? 0)
  const rawParagraphEnd = Number(range?.getAttribute("end") ?? range?.getAttribute("st") ?? 0)
  const paragraphRange = range
    && Number.isFinite(rawParagraphStart)
    && Number.isFinite(rawParagraphEnd)
    ? {
        start: Math.max(0, Math.min(rawParagraphStart, rawParagraphEnd)),
        end: Math.max(0, rawParagraphStart, rawParagraphEnd),
      }
    : undefined
  const characterRange = parseCharacterRange(elements(owner, "charRg")[0])
  const characterParagraphs = characterRange && targetBody
    ? paragraphRangeFromCharacters(targetBody.paragraphs, characterRange)
    : undefined
  const motion = owner.localName === "animMotion" ? owner : elements(owner, "animMotion")[0]
  const rotation = owner.localName === "animRot" ? owner : elements(owner, "animRot")[0]
  const scale = owner.localName === "animScale" ? owner : elements(owner, "animScale")[0]
  const animation = owner.localName === "anim" ? owner : elements(owner, "anim")[0]
  const values: Record<string, string | number | boolean> = {}
  if (filter) values.filter = filter
  const path = motion?.getAttribute("path") ?? ""
  if (path) {
    values.motionPath = path
    values.motionScaleX = context.slideWidth
    values.motionScaleY = context.slideHeight
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
  const cellRange = paragraphRange ?? characterParagraphs
  const granularity = characterRange && targetBody
    ? pptxTextGranularity(pptxTextBody(targetBody.paragraphs), characterRange)
    : undefined
  return Object.freeze({
    id: node.getAttribute("id") ?? "",
    kind: effectKind(presetClass, transition, filter, owner),
    targetObjectKey: targetId ? context.objectKeys.get(targetId) : undefined,
    targetShapeId: targetId,
    targetNodeType: targetIdentity?.nodeType,
    paragraphRange,
    characterRange,
    characterGranularity: granularity?.granularity,
    characterRangeUnavailable: (characterRange !== undefined && !targetBody) || undefined,
    tableCell: targetIdentity?.nodeType === "table" && !characterRange
      ? tableCellForRange(targetBody, cellRange)
      : undefined,
    chartElement: chartElementTarget(owner),
    graphicElement: graphicElementTarget(owner, targetIdentity),
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
  const support = pptxEffectTargetSupport(effect)
  if (support.disposition === "static") return "static"
  if (["appear", "disappear", "fade-in", "fade-out", "scale", "rotate", "media-command"].includes(effect.kind)) return "strict"
  if (effect.kind === "wipe") return /left|right/i.test(effect.filter ?? "") ? "strict" : "static"
  if (effect.kind === "motion-path") return /^[\s\d.,+\-mlehvz]+$/iu.test(effect.motionPath ?? "") ? "strict" : "static"
  if (effect.kind === "emphasis") return effect.values.emphasisProperty === "opacity" ? "strict" : "static"
  if (effect.kind === "set") return "static"
  return "unparsed"
}

function parseTimeNode(
  node: Element,
  context: ParseContext,
  nodes: Record<string, PptxTimeNode>,
  features: PptxFeatureRecord[],
  parentId?: string,
): string {
  const { slideIndex, objectKeys } = context
  const id = node.getAttribute("id") ?? `generated-${slideIndex}-${Object.keys(nodes).length}`
  const container = ["seq", "par", "excl"].includes(node.parentElement?.localName ?? "")
    ? node.parentElement as Element
    : undefined
  const conditions = [
    ...parseConditionList(node, "stCondLst", "start", objectKeys),
    ...parseConditionList(node, "endCondLst", "end", objectKeys),
    ...parseConditionList(node, "prevCondLst", "previous", objectKeys),
    ...parseConditionList(node, "nextCondLst", "next", objectKeys),
    ...(container ? parseConditionList(container, "prevCondLst", "previous", objectKeys) : []),
    ...(container ? parseConditionList(container, "nextCondLst", "next", objectKeys) : []),
  ]
  const fallbackEvent = nodeTypeTrigger(node.getAttribute("nodeType") ?? undefined)
  if (fallbackEvent && !conditions.some((condition) => condition.event === fallbackEvent)) {
    conditions.push(Object.freeze({ source: "start", event: fallbackEvent, delayMs: 0 }))
  }
  const effect = node.hasAttribute("presetClass")
    ? parseEffect(node, context)
    : undefined
  const childIds = directTimeChildren(node).map((child) => parseTimeNode(
    child,
    context,
    nodes,
    features,
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
  const repeatNumber = repeatCountValue(rawRepeat)
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
    repeatCount: repeatNumber,
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
    const targetSupport = pptxEffectTargetSupport(effect)
    const disposition: PptxFeatureDisposition = unsupportedCondition
      ? unsupportedCondition.event === "unknown" ? "unparsed" : "static"
      : baseDisposition
    const elementTarget = Boolean(
      effect.characterRange
      || effect.characterRangeUnavailable
      || effect.tableCell
      || effect.chartElement
      || effect.graphicElement,
    )
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
          : disposition === "static" ? targetSupport.reason ?? "第一版尚未批准该效果参数。" : undefined,
      sourceNodeId: id,
    }))
    if (elementTarget && targetSupport.feature) features.push(Object.freeze({
      id: `element-target:${slideIndex}:${id}`,
      slideIndex,
      objectKey: effect.targetObjectKey,
      feature: targetSupport.feature,
      disposition: unsupportedCondition
        ? unsupportedCondition.event === "unknown" ? "unparsed" : "static"
        : targetSupport.disposition,
      reason: unsupportedCondition
        ? `触发条件 ${unsupportedCondition.rawEvent ?? unsupportedCondition.event} 未执行。`
        : targetSupport.reason,
      sourceNodeId: id,
    }))
    if (timeNode.repeatCount !== undefined) {
      const repeatDisposition: PptxFeatureDisposition = disposition === "strict"
        && typeof timeNode.durationMs === "number"
        && timeNode.durationMs > 0
        ? "strict"
        : "static"
      const encoding = timeNode.repeatCount === "indefinite"
        ? undefined
        : repeatCountNote(rawRepeat, timeNode.repeatCount)
      features.push(Object.freeze({
        id: `repeat:${slideIndex}:${id}`,
        slideIndex,
        objectKey: effect.targetObjectKey,
        feature: timeNode.repeatCount === "indefinite" ? "repeat:indefinite" : "repeat:finite",
        disposition: repeatDisposition,
        reason: encoding
          ?? (repeatDisposition === "strict" ? undefined : "基础效果本身未进入严格执行范围。"),
        sourceNodeId: id,
      }))
    }
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

function mediaContentType(sourcePath: string | undefined): string | undefined {
  const path = sourcePath?.toLowerCase() ?? ""
  if (path.endsWith(".mp4")) return "video/mp4"
  if (path.endsWith(".webm")) return "video/webm"
  if (path.endsWith(".mp3")) return "audio/mpeg"
  if (path.endsWith(".wav")) return "audio/wav"
  if (path.endsWith(".m4a")) return "audio/mp4"
  if (path.endsWith(".ogg") || path.endsWith(".oga")) return "audio/ogg"
  return undefined
}

function parseTransitionSound(
  node: Element,
  slide: SlideData,
): PptxSlideTransitionSound | undefined {
  const soundAction = elements(node, "sndAc")[0]
  const sound = soundAction ? elements(soundAction, "snd")[0] : undefined
  const relationId = localAttr(sound, "embed") ?? localAttr(sound, "link")
  if (!relationId) return undefined
  const relation = slide.rels.get(relationId)
  const external = relation?.targetMode?.toLowerCase() === "external"
  const sourcePath = relation
    ? external ? relation.target : resolveTargetPath(slide.slidePath, relation.target)
    : undefined
  const embedded = !external && sourcePath?.startsWith("ppt/") === true
  return Object.freeze({
    relationId,
    sourcePath,
    contentType: mediaContentType(sourcePath),
    // External sounds stay reported but unloaded under the media policy.
    embedded,
    loop: localAttr(sound, "loop") === "1",
    volume: 1,
  })
}

function parseTransition(
  node: Element | undefined,
  slide: SlideData,
): PptxSlideTransition | undefined {
  if (!node) return undefined
  const effect = Array.from(node.children).find(
    (child) => !["sndAc", "endSnd", "extLst"].includes(child.localName),
  )
  const speed = node.getAttribute("spd")
  const fallback = speed === "fast" ? 500 : speed === "slow" ? 2000 : 1000
  const duration = timeValue(localAttr(node, "dur"), fallback)
  const after = timeValue(node.getAttribute("advTm") ?? undefined, 0)
  const sound = parseTransitionSound(node, slide)
  return Object.freeze({
    kind: effect?.localName ?? "cut",
    direction: effect?.getAttribute("dir") ?? effect?.getAttribute("orient") ?? undefined,
    orientation: effect?.getAttribute("orient") ?? undefined,
    option: effect?.getAttribute("option") ?? undefined,
    durationMs: duration === "indefinite" ? fallback : duration,
    advanceOnClick: node.getAttribute("advClick") !== "0",
    advanceAfterMs: node.hasAttribute("advTm") && after !== "indefinite" ? after : undefined,
    soundRelationId: localAttr(elements(node, "snd")[0], "embed"),
    sound,
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
      contentType: mediaContentType(sourcePath),
      kind: node.localName === "videoFile" ? "video" : "audio",
      embedded: !external && sourcePath?.startsWith("ppt/") === true,
      autoplay: false,
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
    // `numSld` 表示媒体跨页继续播放，`nodeType` 为 withEffect/afterEffect 时它随页面自动开始。
    const autoStartNodeType = localAttr(timing, "nodeType")
    const autoStart = mediaNode.hasAttribute("numSld")
      || autoStartNodeType === "withEffect"
      || autoStartNodeType === "afterEffect"
    result.set(entry.relationId, Object.freeze({
      ...entry,
      autoplay: entry.autoplay || autoStart,
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
    // `tooltip` is the visible affordance PowerPoint shows for the link; it is
    // kept even when the target itself is unsupported so a tooltip-only link
    // still reads correctly.
    const tooltip = node.getAttribute("tooltip")?.trim() || undefined
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
      tooltip,
      rawAction,
    }))
  }
  return Object.freeze(actions)
}

function withObjectHyperlinks(
  objects: readonly PptxObjectIdentity[],
  actions: readonly PptxSlideAction[],
): readonly PptxObjectIdentity[] {
  if (actions.length === 0) return objects
  const byObject = new Map<string, PptxObjectHyperlink[]>()
  for (const action of actions) {
    const hyperlink: PptxObjectHyperlink = Object.freeze({
      trigger: action.trigger,
      kind: action.kind,
      ...(action.targetSlideIndex !== undefined ? { targetSlideIndex: action.targetSlideIndex } : {}),
      ...(action.url !== undefined ? { url: action.url } : {}),
      ...(action.tooltip !== undefined ? { tooltip: action.tooltip } : {}),
    })
    const existing = byObject.get(action.sourceObjectKey)
    if (existing) existing.push(hyperlink)
    else byObject.set(action.sourceObjectKey, [hyperlink])
  }
  return Object.freeze(objects.map((object) => {
    const hyperlinks = byObject.get(object.key)
    return hyperlinks
      ? Object.freeze({ ...object, hyperlinks: Object.freeze(hyperlinks) })
      : object
  }))
}

function morphMatchDescription(method: PptxMorphMatch["method"]): string {
  switch (method) {
    case "explicit-name": return "显式名称"
    case "creation-id": return "创建编号"
    case "name-and-geometry": return "普通名称与几何"
    default: return "文字与几何"
  }
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
  const { objects, textBodies } = parseObjects(document, slide.slidePath)
  const objectKeys = new Map<string, string>()
  const objectsById = new Map<string, PptxObjectIdentity>()
  const duplicateIds = new Set<string>()
  for (const object of objects) {
    if (objectKeys.has(object.shapeId)) {
      objectKeys.delete(object.shapeId)
      duplicateIds.add(object.shapeId)
    } else if (!duplicateIds.has(object.shapeId)) objectKeys.set(object.shapeId, object.key)
    if (!objectsById.has(object.shapeId)) objectsById.set(object.shapeId, object)
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
  const context: ParseContext = {
    slideIndex: slide.index,
    objectKeys,
    objectsById,
    textBodies,
    slideWidth: presentation.width,
    slideHeight: presentation.height,
  }
  const root = elements(document, "cTn").find((node) => node.getAttribute("nodeType") === "tmRoot")
  const rootNodeId = root ? parseTimeNode(
    root,
    context,
    nodes,
    features,
  ) : undefined
  const transitionNodes = elements(document, "transition")
  const transition = parseTransition(
    transitionNodes.find((candidate) => Array.from(candidate.children).some(
      (child) => !["sndAc", "endSnd", "extLst"].includes(child.localName),
    )) ?? transitionNodes[0],
    slide,
  )
  if (transition) {
    const capability = pptxTransitionCapability(transition.kind, options.approximation ?? "off")
    features.push(Object.freeze({
      id: `transition:${slide.index}`,
      slideIndex: slide.index,
      feature: capability.feature,
      disposition: capability.disposition,
      reason: capability.reason,
    }))
  }
  if (transition?.sound) {
    const externalAllowed = options.externalMedia === "allowed" && Boolean(transition.sound.sourcePath)
    features.push(Object.freeze({
      id: `transition-sound:${slide.index}`,
      slideIndex: slide.index,
      feature: PPTX_CAPABILITY_FEATURES.transitionSound,
      disposition: transition.sound.embedded || externalAllowed ? "strict" : "static",
      reason: transition.sound.embedded || externalAllowed
        ? undefined
        : transition.sound.sourcePath
          ? "外部切换声音默认禁用。"
          : "切换声音的关系目标缺失。",
    }))
  }
  const media = withAutoplayMedia(parseMedia(document, objectKeys, slide), nodes, rootNodeId)
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
    reason: "图表构建清单只描述分组；元素级目标在渲染器没有系列 DOM 身份时保留静态。",
  }))
  for (const build of elements(document, "bldDgm")) features.push(Object.freeze({
    id: `smartart-build:${slide.index}:${features.length}`,
    slideIndex: slide.index,
    feature: `smartart-build:${build.getAttribute("bld") ?? "unknown"}`,
    disposition: "static",
    reason: "SmartArt 构建清单只描述分组；逐节点效果按渲染器暴露的子对象身份执行。",
  }))
  for (const build of elements(document, "bldGraphic")) {
    const targetType = objectsById.get(build.getAttribute("spid") ?? "")?.nodeType
    const mode = elements(build, "bldSub")[0]?.getAttribute("bld")
      ?? build.getAttribute("bld")
      ?? "unknown"
    if (targetType === "table") {
      features.push(Object.freeze({
        id: `table-build:${slide.index}:${features.length}`,
        slideIndex: slide.index,
        feature: `${PPTX_CAPABILITY_FEATURES.tableBuild}:${mode}`,
        disposition: "static",
        reason: "表格构建清单只描述分组；逐单元格效果按精确单元格执行。",
      }))
    } else if (targetType === "chart") {
      features.push(Object.freeze({
        id: `chart-build:${slide.index}:${features.length}`,
        slideIndex: slide.index,
        feature: `chart-build:${mode}`,
        disposition: "static",
        reason: "图表构建清单只描述分组；元素级目标在渲染器没有系列 DOM 身份时保留静态。",
      }))
    } else if (targetType === "smartart") {
      features.push(Object.freeze({
        id: `smartart-build:${slide.index}:${features.length}`,
        slideIndex: slide.index,
        feature: `smartart-build:${mode}`,
        disposition: "static",
        reason: "SmartArt 构建清单只描述分组；逐节点效果按渲染器暴露的子对象身份执行。",
      }))
    } else {
      features.push(Object.freeze({
        id: `graphic-build:${slide.index}:${features.length}`,
        slideIndex: slide.index,
        feature: `graphic-build:${mode}`,
        disposition: "static",
        reason: "图形构建清单的目标不是表格、图表或 SmartArt。",
      }))
    }
  }
  const actions = parseActions(document, slide, presentation, objectKeys)
  const linkedObjects = withObjectHyperlinks(objects, actions)
  features.push(...collectPptxTextDirectionFeatures(document, slide.index))
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
    objects: linkedObjects,
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
      // 弱匹配只用于调试报告，不进入任何播放模式。
      disposition: options.approximation === "safe" && match.confidence !== "weak"
        ? "approximate"
        : "static",
      reason: match.confidence === "weak"
        ? "文字弱匹配只用于调试报告，不执行补间。"
        : options.approximation === "safe"
          ? `使用${morphMatchDescription(match.method)}进行${match.confidence === "strong" ? "强" : "中等"}身份匹配，分数 ${match.score}。`
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
