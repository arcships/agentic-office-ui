export type LabPresetClass = "entr" | "exit" | "emph" | "path" | "mediacall" | string

export interface LabEffect {
  id: string
  targetId: string
  presetClass: LabPresetClass
  presetId: string
  nodeType: string
  trigger: "onClick" | "withPrevious" | "afterPrevious" | "unknown"
  duration: number
  delay: number
  start: number
  filter: string
  transition: string
  command: string
  kind: "appear" | "disappear" | "fadeIn" | "fadeOut" | "emphasis" | "path" | "media" | "approximate"
  motionPath: string
  rotationBy: number
  scaleBy: { x: number; y: number } | null
  scaleFrom: { x: number; y: number } | null
  scaleTo: { x: number; y: number } | null
  attributes: string[]
  values: string[]
  paragraphRange: { start: number; end: number } | null
  repeatCount: number | "indefinite" | null
  autoReverse: boolean
  acceleration: number
  deceleration: number
}

export interface LabBatch {
  index: number
  effects: LabEffect[]
  trigger: LabEffect["trigger"]
  duration: number
}

export interface LabInteractiveSequence {
  triggerTargetId: string
  batches: LabBatch[]
}

export interface LabBuild {
  kind: "paragraph" | "chart" | "smartart" | "graphic"
  targetId: string
  mode: string
  level: number | null
}

export interface LabTransition {
  kind: string
  direction: string
  option: string
  duration: number
  advanceMs: number | null
  advanceOnClick: boolean
}

export interface LabAudioCue {
  relationId: string
  name: string
}

export interface LabMediaBookmark {
  targetId: string
  relationId: string
  name: string
  time: number
}

export interface LabBookmarkSequence {
  targetId: string
  bookmarkName: string
  batches: LabBatch[]
}

export interface LabConflict {
  targetId: string
  batchIndex: number
  effectIds: string[]
  properties: string[]
}

export interface LabComplexity {
  timeNodeCount: number
  maxDepth: number
  conditionEvents: string[]
  repeatNodeCount: number
  autoReverseNodeCount: number
  accelerationNodeCount: number
}

export interface LabTimeCondition {
  source: "start" | "end" | "previous" | "next"
  event: string
  delay: number
  targetId: string
  bookmarkName: string
}

export interface LabTimeNode {
  id: string
  container: string
  nodeType: string
  presetClass: string
  duration: number | "indefinite"
  fill: string
  restart: string
  conditions: LabTimeCondition[]
  children: LabTimeNode[]
}

export interface LabTreeValidation {
  treeEffectIds: string[]
  parsedEffectIds: string[]
  matchedEffectIds: string[]
  missingEffectIds: string[]
  orderMatches: boolean
  conditionEvents: string[]
}

export interface LabTiming {
  batches: LabBatch[]
  effects: LabEffect[]
  timingXml: string
  transitionXml: string
  unknownPresetClasses: string[]
  interactive: LabInteractiveSequence[]
  builds: LabBuild[]
  transition: LabTransition | null
  audioCues: LabAudioCue[]
  mediaBookmarks: LabMediaBookmark[]
  bookmarkSequences: LabBookmarkSequence[]
  conflicts: LabConflict[]
  complexity: LabComplexity
  timeTree: LabTimeNode | null
  treeValidation: LabTreeValidation
}

function elements(root: ParentNode, localName: string): Element[] {
  if (!(root instanceof Element || root instanceof Document)) return []
  return Array.from(root.getElementsByTagNameNS("*", localName))
}

function directChild(root: Element, localName: string): Element | undefined {
  return Array.from(root.children).find((node) => node.localName === localName)
}

function localAttr(node: Element | undefined, name: string): string | null {
  if (!node) return null
  return Array.from(node.attributes).find((attr) => attr.localName === name)?.value ?? null
}

function readTime(value: string | null, fallback = 0): number {
  if (!value || value === "indefinite") return fallback
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : fallback
}

function directDelay(cTn: Element): number {
  const list = directChild(cTn, "stCondLst")
  const cond = list ? directChild(list, "cond") : undefined
  return readTime(cond?.getAttribute("delay") ?? null)
}

function inheritedDelay(effectNode: Element, batchNode: Element): number {
  let total = 0
  let current: Element | null = effectNode
  while (current && current !== batchNode) {
    if (current.localName === "cTn") total += directDelay(current)
    current = current.parentElement
  }
  return total
}

function findTarget(effectNode: Element): string {
  return elements(effectNode, "spTgt")[0]?.getAttribute("spid") ?? ""
}

function enclosingShapeId(node: Element): string {
  let current: Element | null = node
  while (current && !["pic", "sp", "graphicFrame"].includes(current.localName)) current = current.parentElement
  return current ? elements(current, "cNvPr")[0]?.getAttribute("id") ?? "" : ""
}

function descendantTimeAttribute(node: Element, name: string): string | null {
  if (localAttr(node, name) !== null) return localAttr(node, name)
  return elements(node, "cTn").map((child) => localAttr(child, name)).find((value) => value !== null) ?? null
}

function parseRepeatCount(value: string | null): LabEffect["repeatCount"] {
  if (!value) return null
  if (value === "indefinite") return "indefinite"
  const count = Number(value)
  return Number.isFinite(count) && count > 0 ? count : null
}

function numberAttr(node: Element | undefined, name: string): number | null {
  if (!node) return null
  const value = Number(node.getAttribute(name))
  return Number.isFinite(value) ? value : null
}

function parseScalePoint(effectNode: Element, name: "by" | "from" | "to"): LabEffect["scaleBy"] {
  const scale = elements(effectNode, "animScale")[0]
  if (!scale) return null
  const source = directChild(scale, name)
  const x = numberAttr(source, "x")
  const y = numberAttr(source, "y")
  return x === null || y === null ? null : { x: x / 100000, y: y / 100000 }
}

function parseParagraphRange(effectNode: Element): LabEffect["paragraphRange"] {
  const range = elements(effectNode, "pRg")[0]
  if (!range) return null
  const start = numberAttr(range, "st") ?? numberAttr(range, "start")
  const end = numberAttr(range, "end")
  return start === null || end === null ? null : { start, end }
}

function findMainSequence(doc: Document): Element | undefined {
  return elements(doc, "seq").find((seq) => {
    const cTn = directChild(seq, "cTn")
    return cTn?.getAttribute("nodeType") === "mainSeq"
  })
}

function triggerFor(nodeType: string): LabEffect["trigger"] {
  if (nodeType === "clickEffect") return "onClick"
  if (nodeType === "withEffect") return "withPrevious"
  if (nodeType === "afterEffect") return "afterPrevious"
  return "unknown"
}

function kindFor(presetClass: string, transition: string, filter: string): LabEffect["kind"] {
  if (presetClass === "mediacall") return "media"
  if (presetClass === "emph") return "emphasis"
  if (presetClass === "path") return "path"
  if (presetClass === "entr") {
    return transition === "in" && filter.includes("fade") ? "fadeIn" : "appear"
  }
  if (presetClass === "exit") {
    return transition === "out" && filter.includes("fade") ? "fadeOut" : "disappear"
  }
  return "approximate"
}

function parseEffect(cTn: Element, batchNode: Element): LabEffect {
  const animEffect = elements(cTn, "animEffect")[0]
  const ownDuration = readTime(cTn.getAttribute("dur"), 500)
  const childDuration = animEffect
    ? readTime(elements(animEffect, "cTn")[0]?.getAttribute("dur") ?? null, ownDuration)
    : ownDuration
  const nodeType = cTn.getAttribute("nodeType") ?? ""
  const command = elements(cTn, "cmd")[0]?.getAttribute("cmd") ?? ""
  const filter = animEffect?.getAttribute("filter") ?? ""
  const transition = animEffect?.getAttribute("transition") ?? ""
  const presetClass = cTn.getAttribute("presetClass") ?? "unknown"
  const motion = elements(cTn, "animMotion")[0]
  const rotation = elements(cTn, "animRot")[0]
  return {
    id: cTn.getAttribute("id") ?? crypto.randomUUID(),
    targetId: findTarget(cTn),
    presetClass,
    presetId: cTn.getAttribute("presetID") ?? "",
    nodeType,
    trigger: triggerFor(nodeType),
    duration: Math.max(1, childDuration),
    delay: inheritedDelay(cTn, batchNode),
    start: 0,
    filter,
    transition,
    command,
    kind: kindFor(presetClass, transition, filter),
    motionPath: motion?.getAttribute("path") ?? "",
    rotationBy: (numberAttr(rotation, "by") ?? 0) / 60000,
    scaleBy: parseScalePoint(cTn, "by"),
    scaleFrom: parseScalePoint(cTn, "from"),
    scaleTo: parseScalePoint(cTn, "to"),
    attributes: elements(cTn, "attrName").map((node) => node.textContent?.trim() ?? "").filter(Boolean),
    values: ["strVal", "fltVal", "intVal"].flatMap((name) => elements(cTn, name).map((node) => node.getAttribute("val") ?? "")).filter(Boolean),
    paragraphRange: parseParagraphRange(cTn),
    repeatCount: parseRepeatCount(descendantTimeAttribute(cTn, "repeatCount")),
    autoReverse: descendantTimeAttribute(cTn, "autoRev") === "1",
    acceleration: readTime(descendantTimeAttribute(cTn, "accel")) / 100000,
    deceleration: readTime(descendantTimeAttribute(cTn, "decel")) / 100000,
  }
}

function effectiveDuration(effect: LabEffect): number {
  const repeats = effect.repeatCount === "indefinite" ? 2 : effect.repeatCount ?? 1
  return effect.duration * repeats * (effect.autoReverse ? 2 : 1)
}

function scheduleBatch(index: number, effects: LabEffect[]): LabBatch {
  let previousStart = 0
  let previousEnd = 0
  effects.forEach((effect, effectIndex) => {
    if (effectIndex === 0 || effect.trigger === "onClick" || effect.trigger === "unknown") {
      effect.start = effect.delay
    } else if (effect.trigger === "withPrevious") {
      effect.start = previousStart + effect.delay
    } else {
      effect.start = previousEnd + effect.delay
    }
    previousStart = effect.start
    previousEnd = effect.start + effectiveDuration(effect)
  })
  return {
    index,
    effects,
    trigger: effects[0]?.trigger ?? "unknown",
    duration: Math.max(0, ...effects.map((effect) => effect.start + effectiveDuration(effect))),
  }
}

function parseBatchList(list: Element | undefined): LabBatch[] {
  if (!list) return []
  return Array.from(list.children).map((batchNode, index) => scheduleBatch(
    index,
    elements(batchNode, "cTn")
      .filter((node) => node.hasAttribute("presetClass"))
      .map((node) => parseEffect(node, batchNode)),
  )).filter((batch) => batch.effects.length > 0)
}

function parseBuilds(doc: Document): LabBuild[] {
  const list = elements(doc, "bldLst")[0]
  if (!list) return []
  return Array.from(list.children).map((node): LabBuild | null => {
    let kind: LabBuild["kind"] = "graphic"
    let mode = node.getAttribute("build") ?? "all"
    if (node.localName === "bldP") kind = "paragraph"
    else if (node.localName === "bldDgm") kind = "smartart"
    else if (elements(node, "bldDgm").length) {
      kind = "smartart"
      mode = elements(node, "bldDgm")[0]?.getAttribute("bld") ?? mode
    } else if (elements(node, "bldChart").length) {
      kind = "chart"
      mode = elements(node, "bldChart")[0]?.getAttribute("bld") ?? mode
    }
    const targetId = node.getAttribute("spid") ?? ""
    if (!targetId) return null
    return { kind, targetId, mode, level: numberAttr(node, "bldLvl") }
  }).filter((value): value is LabBuild => value !== null)
}

function parseTransition(node: Element | undefined): LabTransition | null {
  if (!node) return null
  const effect = Array.from(node.children).find((child) => !["sndAc", "extLst"].includes(child.localName))
  const speed = node.getAttribute("spd")
  const fallback = speed === "fast" ? 500 : speed === "slow" ? 2000 : 1000
  const advance = node.getAttribute("advTm")
  return {
    kind: effect?.localName ?? "cut",
    direction: effect?.getAttribute("dir") ?? effect?.getAttribute("orient") ?? "",
    option: effect?.getAttribute("option") ?? "",
    duration: readTime(localAttr(node, "dur"), fallback),
    advanceMs: advance === null ? null : readTime(advance),
    advanceOnClick: node.getAttribute("advClick") !== "0",
  }
}

function parseMediaBookmarks(doc: Document): LabMediaBookmark[] {
  return elements(doc, "media").flatMap((media): LabMediaBookmark[] => {
    const targetId = enclosingShapeId(media)
    const relationId = localAttr(media, "embed") ?? localAttr(media, "link") ?? ""
    if (!targetId) return []
    return elements(media, "bmk").flatMap((bookmark): LabMediaBookmark[] => {
      const name = bookmark.getAttribute("name") ?? ""
      const time = readTime(bookmark.getAttribute("time"), -1)
      return name && time >= 0 ? [{ targetId, relationId, name, time }] : []
    })
  })
}

function parseBookmarkSequences(doc: Document): LabBookmarkSequence[] {
  return elements(doc, "cond").flatMap((condition): LabBookmarkSequence[] => {
    if (condition.getAttribute("evt") !== "onMediaBookmark") return []
    const target = elements(condition, "bmkTgt")[0]
    if (!target) return []
    let timeNode = condition.parentElement
    while (timeNode && timeNode.localName !== "cTn") timeNode = timeNode.parentElement
    const childList = timeNode ? directChild(timeNode, "childTnLst") : undefined
    return [{
      targetId: target.getAttribute("spid") ?? "",
      bookmarkName: target.getAttribute("bmkName") ?? "",
      batches: parseBatchList(childList),
    }]
  }).filter((sequence) => sequence.targetId && sequence.bookmarkName)
}

function effectProperties(effect: LabEffect): string[] {
  const properties = new Set<string>()
  if (["appear", "disappear", "fadeIn", "fadeOut"].includes(effect.kind)) properties.add("opacity")
  if (effect.kind === "path") properties.add("translate")
  if (effect.kind === "emphasis") { properties.add("scale"); properties.add("filter"); properties.add("opacity") }
  if (effect.scaleBy || effect.scaleFrom || effect.scaleTo) properties.add("scale")
  if (effect.rotationBy) properties.add("rotate")
  effect.attributes.forEach((attribute) => properties.add(attribute))
  if (effect.kind === "media") properties.add("media")
  return Array.from(properties)
}

function findConflicts(batches: LabBatch[]): LabConflict[] {
  return batches.flatMap((batch): LabConflict[] => {
    const conflicts: LabConflict[] = []
    for (let leftIndex = 0; leftIndex < batch.effects.length; leftIndex += 1) {
      const left = batch.effects[leftIndex]
      for (let rightIndex = leftIndex + 1; rightIndex < batch.effects.length; rightIndex += 1) {
        const right = batch.effects[rightIndex]
        if (!left.targetId || left.targetId !== right.targetId) continue
        const overlaps = left.start < right.start + effectiveDuration(right) && right.start < left.start + effectiveDuration(left)
        if (!overlaps) continue
        const rightProperties = new Set(effectProperties(right))
        const properties = effectProperties(left).filter((property) => rightProperties.has(property))
        if (properties.length) conflicts.push({ targetId: left.targetId, batchIndex: batch.index, effectIds: [left.id, right.id], properties })
      }
    }
    return conflicts
  })
}

function parseComplexity(doc: Document): LabComplexity {
  const timeNodes = elements(doc, "cTn")
  const depth = (node: Element): number => {
    let current = node.parentElement
    let value = 1
    while (current) { if (current.localName === "cTn") value += 1; current = current.parentElement }
    return value
  }
  return {
    timeNodeCount: timeNodes.length,
    maxDepth: Math.max(0, ...timeNodes.map(depth)),
    conditionEvents: Array.from(new Set(elements(doc, "cond").map((node) => node.getAttribute("evt") ?? "").filter(Boolean))),
    repeatNodeCount: timeNodes.filter((node) => localAttr(node, "repeatCount") !== null || localAttr(node, "repeatDur") !== null).length,
    autoReverseNodeCount: timeNodes.filter((node) => localAttr(node, "autoRev") === "1").length,
    accelerationNodeCount: timeNodes.filter((node) => localAttr(node, "accel") !== null || localAttr(node, "decel") !== null).length,
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

function parseConditionList(node: Element, listName: string, source: LabTimeCondition["source"]): LabTimeCondition[] {
  const list = directChild(node, listName)
  if (!list) return []
  return Array.from(list.children).filter((condition) => condition.localName === "cond").map((condition) => {
    const bookmark = elements(condition, "bmkTgt")[0]
    return {
      source,
      event: condition.getAttribute("evt") ?? "delay",
      delay: readTime(condition.getAttribute("delay")),
      targetId: findTarget(condition) || bookmark?.getAttribute("spid") || "",
      bookmarkName: bookmark?.getAttribute("bmkName") ?? "",
    }
  })
}

function parseTimeNode(node: Element): LabTimeNode {
  const rawDuration = localAttr(node, "dur")
  const container = node.parentElement
  return {
    id: node.getAttribute("id") ?? "",
    container: container?.localName ?? "",
    nodeType: node.getAttribute("nodeType") ?? "",
    presetClass: node.getAttribute("presetClass") ?? "",
    duration: rawDuration === "indefinite" ? "indefinite" : readTime(rawDuration),
    fill: node.getAttribute("fill") ?? "",
    restart: node.getAttribute("restart") ?? "",
    conditions: [
      ...parseConditionList(node, "stCondLst", "start"),
      ...parseConditionList(node, "endCondLst", "end"),
      ...parseConditionList(node, "prevCondLst", "previous"),
      ...parseConditionList(node, "nextCondLst", "next"),
      ...(container?.localName === "seq" ? parseConditionList(container, "prevCondLst", "previous") : []),
      ...(container?.localName === "seq" ? parseConditionList(container, "nextCondLst", "next") : []),
    ],
    children: directTimeChildren(node).map(parseTimeNode),
  }
}

function flattenTimeTree(root: LabTimeNode | null): LabTimeNode[] {
  return root ? [root, ...root.children.flatMap((child) => flattenTimeTree(child))] : []
}

function validateTimeTree(root: LabTimeNode | null, effects: LabEffect[]): LabTreeValidation {
  const nodes = flattenTimeTree(root)
  const treeEffectIds = nodes.filter((node) => node.presetClass).map((node) => node.id)
  const parsedEffectIds = effects.map((effect) => effect.id)
  const treeIds = new Set(treeEffectIds)
  const matchedEffectIds = parsedEffectIds.filter((id) => treeIds.has(id))
  const missingEffectIds = parsedEffectIds.filter((id) => !treeIds.has(id))
  const parsedSet = new Set(parsedEffectIds)
  const comparableTreeIds = treeEffectIds.filter((id) => parsedSet.has(id))
  return {
    treeEffectIds,
    parsedEffectIds,
    matchedEffectIds,
    missingEffectIds,
    orderMatches: comparableTreeIds.join("|") === matchedEffectIds.join("|"),
    conditionEvents: Array.from(new Set(nodes.flatMap((node) => node.conditions.map((condition) => condition.event)))),
  }
}

export function parseSlideTiming(xml: string): LabTiming {
  const doc = new DOMParser().parseFromString(xml, "application/xml")
  const parserError = doc.querySelector("parsererror")
  if (parserError) throw new Error(parserError.textContent ?? "幻灯片 XML 解析失败")

  const timing = elements(doc, "timing")[0]
  const transition = elements(doc, "transition")[0]
  const mainSeq = findMainSequence(doc)
  const mainCTn = mainSeq ? directChild(mainSeq, "cTn") : undefined
  const childList = mainCTn ? directChild(mainCTn, "childTnLst") : undefined
  const batches = parseBatchList(childList)

  const interactive = elements(doc, "seq").flatMap((seq): LabInteractiveSequence[] => {
    const cTn = directChild(seq, "cTn")
    if (cTn?.getAttribute("nodeType") !== "interactiveSeq") return []
    const triggerList = directChild(cTn, "stCondLst")
    const triggerTargetId = triggerList ? findTarget(triggerList) : ""
    const list = directChild(cTn, "childTnLst")
    return triggerTargetId ? [{ triggerTargetId, batches: parseBatchList(list) }] : []
  })
  const bookmarkSequences = parseBookmarkSequences(doc)

  const effects = [
    ...batches.flatMap((batch) => batch.effects),
    ...interactive.flatMap((sequence) => sequence.batches.flatMap((batch) => batch.effects)),
    ...bookmarkSequences.flatMap((sequence) => sequence.batches.flatMap((batch) => batch.effects)),
  ]
  const known = new Set(["entr", "exit", "emph", "path", "mediacall"])
  const rootTimeNode = elements(doc, "cTn").find((node) => node.getAttribute("nodeType") === "tmRoot")
  const timeTree = rootTimeNode ? parseTimeNode(rootTimeNode) : null
  const audioCues = Array.from(new Map(
    elements(doc, "sndTgt").flatMap((node): Array<[string, LabAudioCue]> => {
      const relationId = Array.from(node.attributes).find((attr) => attr.localName === "embed")?.value ?? ""
      return relationId ? [[relationId, { relationId, name: node.getAttribute("name") ?? "" }]] : []
    }),
  ).values())
  return {
    batches,
    effects,
    timingXml: timing?.outerHTML ?? "",
    transitionXml: transition?.outerHTML ?? "",
    unknownPresetClasses: Array.from(new Set(
      effects.map((effect) => effect.presetClass).filter((value) => !known.has(value)),
    )),
    interactive,
    builds: parseBuilds(doc),
    transition: parseTransition(transition),
    audioCues,
    mediaBookmarks: parseMediaBookmarks(doc),
    bookmarkSequences,
    conflicts: findConflicts([
      ...batches,
      ...interactive.flatMap((sequence) => sequence.batches),
      ...bookmarkSequences.flatMap((sequence) => sequence.batches),
    ]),
    complexity: parseComplexity(doc),
    timeTree,
    treeValidation: validateTimeTree(timeTree, effects),
  }
}
