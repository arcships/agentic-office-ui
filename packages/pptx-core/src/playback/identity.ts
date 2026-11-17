import type {
  PptxMorphMatch,
  PptxObjectGeometry,
  PptxObjectIdentity,
  PptxObjectSource,
} from "./types"

export interface PptxObjectKeyParts {
  slidePath: string
  source: PptxObjectSource
  shapeId: string
  groupPath?: readonly string[]
}

function assertKeyPart(value: string, name: string): string {
  const normalized = value.trim()
  if (!normalized || /[|/]/u.test(normalized)) {
    throw new TypeError(`${name} 不能为空，也不能包含“|”或“/”。`)
  }
  return normalized
}

export function createPptxObjectKey(parts: PptxObjectKeyParts): string {
  const slidePath = parts.slidePath.trim().replace(/^\/+|\/+$/gu, "")
  if (!slidePath || slidePath.includes("|")) throw new TypeError("slidePath 无效。")
  const shapeId = assertKeyPart(parts.shapeId, "shapeId")
  const groupPath = (parts.groupPath ?? []).map((id) => `group:${assertKeyPart(id, "groupPath")}`)
  return `${slidePath}|${parts.source}|${[...groupPath, `shape:${shapeId}`].join("/")}`
}

/**
 * 中等和弱匹配需要“接近几何”。阈值故意宽松：Morph 本来就会把对象移到别处，
 * 几何只用来排除尺寸或位置完全不像的候选，并给匹配打分。
 */
export interface PptxMorphGeometryLimits {
  /** 中心距离与较大对象对角线的比值上限。 */
  maxDistanceRatio: number
  /** 长宽相似度（两轴较小值除以较大值的平均）下限。 */
  minScaleRatio: number
}

export const DEFAULT_PPTX_MORPH_GEOMETRY_LIMITS: PptxMorphGeometryLimits = Object.freeze({
  maxDistanceRatio: 6,
  minScaleRatio: 0.2,
})

export interface PptxObjectGeometryValues {
  x?: string | number | null
  y?: string | number | null
  cx?: string | number | null
  cy?: string | number | null
}

function numericAttribute(value: string | number | null | undefined): number | undefined {
  if (value === null || value === undefined || value === "") return undefined
  const number = typeof value === "number" ? value : Number(value)
  return Number.isFinite(number) ? number : undefined
}

/** 从 `a:off`/`a:ext` 的属性值构造对象几何；缺失或非法时返回 undefined。 */
export function createPptxObjectGeometry(values: PptxObjectGeometryValues): PptxObjectGeometry | undefined {
  const x = numericAttribute(values.x)
  const y = numericAttribute(values.y)
  const width = numericAttribute(values.cx)
  const height = numericAttribute(values.cy)
  if (x === undefined || y === undefined || width === undefined || height === undefined) return undefined
  if (width <= 0 || height <= 0) return undefined
  return Object.freeze({ x, y, width, height })
}

/** 对象文字用于弱匹配：折叠空白，过长或为空时不参与匹配。 */
export function normalizePptxObjectText(value: string | undefined): string | undefined {
  const text = (value ?? "").replace(/\s+/gu, " ").trim()
  if (!text || text.length > 240) return undefined
  return text
}

function geometryScore(
  from: PptxObjectGeometry | undefined,
  to: PptxObjectGeometry | undefined,
  limits: PptxMorphGeometryLimits,
): number | undefined {
  if (!from || !to) return undefined
  const scale = (
    Math.min(from.width, to.width) / Math.max(from.width, to.width)
    + Math.min(from.height, to.height) / Math.max(from.height, to.height)
  ) / 2
  const diagonal = Math.hypot(Math.max(from.width, to.width), Math.max(from.height, to.height))
  const distance = Math.hypot(
    from.x + from.width / 2 - (to.x + to.width / 2),
    from.y + from.height / 2 - (to.y + to.height / 2),
  )
  const distanceRatio = diagonal > 0 ? distance / diagonal : Number.POSITIVE_INFINITY
  if (scale < limits.minScaleRatio || distanceRatio > limits.maxDistanceRatio) return undefined
  const position = 1 / (1 + distanceRatio)
  return Math.round((scale * 0.6 + position * 0.4) * 1000) / 1000
}

/** 中等和弱匹配的几何分数；不可接受时返回 undefined。 */
export function pptxMorphGeometryScore(
  from: PptxObjectGeometry | undefined,
  to: PptxObjectGeometry | undefined,
  limits: PptxMorphGeometryLimits = DEFAULT_PPTX_MORPH_GEOMETRY_LIMITS,
): number | undefined {
  return geometryScore(from, to, limits)
}

function uniqueIdentityMap(
  objects: readonly PptxObjectIdentity[],
  value: (object: PptxObjectIdentity) => string | undefined,
): Map<string, PptxObjectIdentity> {
  const result = new Map<string, PptxObjectIdentity>()
  const duplicates = new Set<string>()
  for (const object of objects) {
    const key = value(object)?.trim()
    if (!key || duplicates.has(key)) continue
    if (result.has(key)) {
      result.delete(key)
      duplicates.add(key)
    } else result.set(key, object)
  }
  return result
}

function matchKey(value: string | undefined): string | undefined {
  return normalizePptxObjectText(value)
}

/**
 * 跨页 Morph 匹配，优先级与能力报告一致：
 * 1. `!!` 显式名称：强匹配；
 * 2. 相同创建编号：强匹配；
 * 3. 相同普通名称 + 对象类型 + 接近几何：中等匹配；
 * 4. 相同文字 + 对象类型 + 接近几何：弱匹配。
 * 每个候选键必须在前后页各自唯一；弱匹配只用于报告，不进入严格执行。
 */
export function matchPptxMorphObjects(
  fromObjects: readonly PptxObjectIdentity[],
  toObjects: readonly PptxObjectIdentity[],
  limits: PptxMorphGeometryLimits = DEFAULT_PPTX_MORPH_GEOMETRY_LIMITS,
): readonly PptxMorphMatch[] {
  const matches: PptxMorphMatch[] = []
  const usedFrom = new Set<string>()
  const usedTo = new Set<string>()
  const append = (
    fromObject: PptxObjectIdentity,
    toObject: PptxObjectIdentity,
    method: PptxMorphMatch["method"],
    confidence: PptxMorphMatch["confidence"],
    score: number,
  ): void => {
    usedFrom.add(fromObject.key)
    usedTo.add(toObject.key)
    matches.push(Object.freeze({
      from: fromObject.key,
      to: toObject.key,
      method,
      confidence,
      score,
      unique: true,
    }))
  }
  const appendStrong = (
    method: PptxMorphMatch["method"],
    from: ReadonlyMap<string, PptxObjectIdentity>,
    to: ReadonlyMap<string, PptxObjectIdentity>,
  ): void => {
    for (const [identity, fromObject] of from) {
      const toObject = to.get(identity)
      if (!toObject || usedFrom.has(fromObject.key) || usedTo.has(toObject.key)) continue
      append(fromObject, toObject, method, "strong", 1)
    }
  }
  const appendGraded = (
    method: PptxMorphMatch["method"],
    confidence: PptxMorphMatch["confidence"],
    keyOf: (object: PptxObjectIdentity) => string | undefined,
  ): void => {
    // 唯一性只看尚未匹配的对象：先被强匹配拿走的对象不再参与后两级候选。
    const from = uniqueIdentityMap(fromObjects.filter((object) => !usedFrom.has(object.key)), keyOf)
    const to = uniqueIdentityMap(toObjects.filter((object) => !usedTo.has(object.key)), keyOf)
    for (const [key, fromObject] of from) {
      const toObject = to.get(key)
      if (!toObject || usedFrom.has(fromObject.key) || usedTo.has(toObject.key)) continue
      const score = geometryScore(fromObject.geometry, toObject.geometry, limits)
      if (score === undefined) continue
      append(fromObject, toObject, method, confidence, score)
    }
  }

  appendStrong(
    "explicit-name",
    uniqueIdentityMap(fromObjects, (object) => object.explicitMorphName),
    uniqueIdentityMap(toObjects, (object) => object.explicitMorphName),
  )
  appendStrong(
    "creation-id",
    uniqueIdentityMap(fromObjects, (object) => object.creationId?.toLowerCase()),
    uniqueIdentityMap(toObjects, (object) => object.creationId?.toLowerCase()),
  )
  appendGraded(
    "name-and-geometry",
    "medium",
    (object) => {
      const name = object.explicitMorphName ? undefined : matchKey(object.name)
      return name ? `${name}\u0000${object.nodeType}` : undefined
    },
  )
  appendGraded(
    "text-and-geometry",
    "weak",
    (object) => {
      const text = matchKey(object.text)
      return text ? `${text}\u0000${object.nodeType}` : undefined
    },
  )
  return Object.freeze(matches)
}
