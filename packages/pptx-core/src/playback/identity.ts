import type { PptxMorphMatch, PptxObjectIdentity, PptxObjectSource } from "./types"

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

export function matchPptxMorphObjects(
  fromObjects: readonly PptxObjectIdentity[],
  toObjects: readonly PptxObjectIdentity[],
): readonly PptxMorphMatch[] {
  const matches: PptxMorphMatch[] = []
  const usedFrom = new Set<string>()
  const usedTo = new Set<string>()
  const appendMatches = (
    method: PptxMorphMatch["method"],
    from: ReadonlyMap<string, PptxObjectIdentity>,
    to: ReadonlyMap<string, PptxObjectIdentity>,
  ) => {
    for (const [identity, fromObject] of from) {
      const toObject = to.get(identity)
      if (!toObject || usedFrom.has(fromObject.key) || usedTo.has(toObject.key)) continue
      usedFrom.add(fromObject.key)
      usedTo.add(toObject.key)
      matches.push(Object.freeze({
        from: fromObject.key,
        to: toObject.key,
        method,
        confidence: "strong",
        score: 1,
        unique: true,
      }))
    }
  }
  appendMatches(
    "explicit-name",
    uniqueIdentityMap(fromObjects, (object) => object.explicitMorphName),
    uniqueIdentityMap(toObjects, (object) => object.explicitMorphName),
  )
  appendMatches(
    "creation-id",
    uniqueIdentityMap(fromObjects, (object) => object.creationId?.toLowerCase()),
    uniqueIdentityMap(toObjects, (object) => object.creationId?.toLowerCase()),
  )
  return Object.freeze(matches)
}
