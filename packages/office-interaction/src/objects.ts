import type {
  JsonValue,
  OfficeFormat,
  OfficeHitCandidate,
  OfficeObjectDescriptor,
  OfficeObjectFingerprint,
  OfficeObjectKind,
  OfficeObjectReferenceDraft,
  OfficeObjectReliability,
  OfficeReferenceCandidatePreview,
  OfficeVisualFragment,
} from "./types"
import { rankOfficeHitCandidates } from "./candidates"

/**
 * Extensible object registration shared by every format adapter.
 *
 * A provider turns one format's native model into *current-revision* object
 * descriptors. Providers are registered, not hard-coded, so a Surface or a
 * host can add an object family (for example a visual PDF recognizer) without
 * changing the selection contract. Descriptors are the only thing a provider
 * returns: they carry a reference draft without `referenceId`, so hovering or
 * listing objects never commits a reference.
 */
export interface OfficeObjectDescriptorInit {
  /** Stable inside the current revision's object index only. */
  objectId: string
  draft: OfficeObjectReferenceDraft
  label: string
  path: readonly { kind: OfficeObjectKind; label: string }[]
  parentObjectId?: string
  childrenState?: OfficeObjectDescriptor["childrenState"]
  content?: OfficeObjectDescriptor["content"]
  visual?: OfficeObjectDescriptor["visual"]
  facets?: Readonly<Record<string, JsonValue>>
}

function copyPath(path: readonly { kind: OfficeObjectKind; label: string }[]): OfficeObjectDescriptor["path"] {
  return path.map((entry) => ({ kind: entry.kind, label: entry.label }))
}

function copyFragments(
  fragments: readonly OfficeVisualFragment[],
): readonly OfficeVisualFragment[] {
  return fragments.map((fragment) => "rect" in fragment
    ? {
        container: { ...fragment.container },
        rect: { ...fragment.rect },
        ...(fragment.zIndex === undefined ? {} : { zIndex: fragment.zIndex }),
      }
    : {
        container: { ...fragment.container },
        region: {
          ...fragment.region,
          start: { ...fragment.region.start },
          end: { ...fragment.region.end },
        },
        ...(fragment.zIndex === undefined ? {} : { zIndex: fragment.zIndex }),
      })
}

/**
 * Build and defensively copy an object descriptor.
 *
 * The copy matters: descriptors are cached per revision and read by every
 * hover, so a host mutating a returned descriptor must not change later reads.
 */
export function createOfficeObjectDescriptor(init: OfficeObjectDescriptorInit): OfficeObjectDescriptor {
  const objectId = init.objectId.trim()
  if (!objectId) throw new TypeError("Office object descriptors require a non-empty objectId")
  if (!init.label.trim()) throw new TypeError("Office object descriptors require a non-empty label")
  const parentObjectId = init.parentObjectId?.trim()
  if (init.parentObjectId !== undefined && !parentObjectId) {
    throw new TypeError("Office object descriptors require a non-empty parentObjectId when present")
  }
  if (parentObjectId === objectId) {
    throw new RangeError("Office object descriptor cannot be its own parent")
  }
  return {
    objectId,
    draft: init.draft,
    label: init.label,
    path: copyPath(init.path),
    ...(parentObjectId === undefined ? {} : { parentObjectId }),
    childrenState: init.childrenState ?? "none",
    ...(init.content === undefined ? {} : { content: { ...init.content } }),
    ...(init.visual === undefined
      ? {}
      : {
          visual: {
            fragments: copyFragments(init.visual.fragments),
            layoutVersion: init.visual.layoutVersion,
          },
        }),
    ...(init.facets === undefined ? {} : { facets: { ...init.facets } }),
  }
}

/** Copy a fingerprint so callers cannot mutate cached evidence. */
export function createOfficeObjectFingerprint(
  init: OfficeObjectFingerprint,
): OfficeObjectFingerprint {
  return {
    ...(init.exactText === undefined ? {} : { exactText: init.exactText }),
    ...(init.prefixText === undefined ? {} : { prefixText: init.prefixText }),
    ...(init.suffixText === undefined ? {} : { suffixText: init.suffixText }),
    ...(init.objectName === undefined ? {} : { objectName: init.objectName }),
    ...(init.contentHash === undefined ? {} : { contentHash: init.contentHash }),
    ...(init.ancestorKeys === undefined ? {} : { ancestorKeys: Object.freeze([...init.ancestorKeys]) as readonly string[] }),
  }
}

export interface OfficeObjectListOptions {
  /** Restrict the list to these object kinds. Unknown kinds are simply absent. */
  kinds?: readonly OfficeObjectKind[]
  /** Skip descriptors without visualized geometry (stage 4 behaviors). */
  visibleOnly?: boolean
  /** Hard cap so one provider cannot flood a hover/list request. */
  maxObjects?: number
}

export const MAX_OFFICE_OBJECT_CANDIDATES = 200

/**
 * One extensible object family for a single format.
 *
 * Providers must be synchronous and side-effect free: the Surface calls them
 * from `pointermove` through a cached per-revision index. Anything expensive
 * (OCR, workbook scans) belongs behind the provider's own cache and must be
 * prepared before listing.
 */
export interface OfficeObjectProvider<TContext> {
  readonly id: string
  readonly format: OfficeFormat
  readonly kinds: readonly OfficeObjectKind[]
  list(context: TContext, options?: OfficeObjectListOptions): readonly OfficeObjectDescriptor[]
}

export interface OfficeObjectRegistry<TContext> {
  /** Register or replace a provider by id. Returns an idempotent unregister. */
  register(provider: OfficeObjectProvider<TContext>): () => void
  unregister(providerId: string): boolean
  get(providerId: string): OfficeObjectProvider<TContext> | undefined
  providers(): readonly OfficeObjectProvider<TContext>[]
  /** Describe every registered object for `context`, filtered and capped. */
  list(context: TContext, options?: OfficeObjectListOptions): readonly OfficeObjectDescriptor[]
}

function matchesKindFilter(kind: OfficeObjectKind, filter?: readonly OfficeObjectKind[]): boolean {
  return filter === undefined || filter.length === 0 || filter.includes(kind)
}

function normalizeListOptions(options?: OfficeObjectListOptions): Required<Pick<OfficeObjectListOptions, "maxObjects">> & OfficeObjectListOptions {
  const maxObjects = options?.maxObjects ?? MAX_OFFICE_OBJECT_CANDIDATES
  if (!Number.isSafeInteger(maxObjects) || maxObjects <= 0) {
    throw new RangeError("Office object list maxObjects must be a positive safe integer")
  }
  if (maxObjects > MAX_OFFICE_OBJECT_CANDIDATES) {
    throw new RangeError(`Office object list maxObjects must not exceed ${MAX_OFFICE_OBJECT_CANDIDATES}`)
  }
  return { ...options, maxObjects }
}

function listProvider<TContext>(
  provider: OfficeObjectProvider<TContext>,
  context: TContext,
  options: OfficeObjectListOptions,
): readonly OfficeObjectDescriptor[] {
  const descriptors = provider.list(context, options)
  const result: OfficeObjectDescriptor[] = []
  for (const descriptor of descriptors) {
    if (descriptor.draft.document.format !== provider.format) {
      throw new TypeError(`Office object provider ${provider.id} returned a ${descriptor.draft.document.format} descriptor`)
    }
    if (!matchesKindFilter(descriptor.draft.kind, options.kinds)) continue
    if (options.visibleOnly === true && !isOfficeObjectVisible(descriptor)) continue
    result.push(descriptor)
    if (result.length >= (options.maxObjects ?? MAX_OFFICE_OBJECT_CANDIDATES)) break
  }
  return result
}

export function createOfficeObjectRegistry<TContext>(
  initial: readonly OfficeObjectProvider<TContext>[] = [],
): OfficeObjectRegistry<TContext> {
  const registry = new Map<string, OfficeObjectProvider<TContext>>()
  for (const provider of initial) {
    if (!provider.id.trim()) throw new TypeError("Office object providers require a non-empty id")
    if (registry.has(provider.id)) throw new RangeError(`duplicate Office object provider id: ${provider.id}`)
    registry.set(provider.id, provider)
  }
  return {
    register(provider) {
      if (!provider.id.trim()) throw new TypeError("Office object providers require a non-empty id")
      registry.set(provider.id, provider)
      return () => {
        if (registry.get(provider.id) === provider) registry.delete(provider.id)
      }
    },
    unregister(providerId) {
      return registry.delete(providerId)
    },
    get(providerId) {
      return registry.get(providerId)
    },
    providers() {
      return Object.freeze([...registry.values()])
    },
    list(context, options) {
      const resolved = normalizeListOptions(options)
      const result: OfficeObjectDescriptor[] = []
      for (const provider of registry.values()) {
        if (!provider.kinds.some((kind) => matchesKindFilter(kind, resolved.kinds))) continue
        result.push(...listProvider(provider, context, resolved))
        if (result.length >= resolved.maxObjects) break
      }
      if (result.length > resolved.maxObjects) result.length = resolved.maxObjects
      return Object.freeze(result)
    },
  }
}

/**
 * True when a descriptor has rendered geometry.
 *
 * Stage 4 behaviors (formulas, animations, protections, …) are references
 * without an outline; the product contract says they are selected through a
 * timeline, property panel or structure tree, never a canvas rectangle.
 */
export function isOfficeObjectVisible(descriptor: OfficeObjectDescriptor): boolean {
  return (descriptor.visual?.fragments.length ?? 0) > 0
}

/** Root-to-object chain for one object id, derived from `parentObjectId`. */
export function officeObjectAncestors(
  descriptors: readonly OfficeObjectDescriptor[],
  objectId: string,
): readonly OfficeObjectDescriptor[] {
  const byId = new Map(descriptors.map((descriptor) => [descriptor.objectId, descriptor]))
  const chain: OfficeObjectDescriptor[] = []
  const seen = new Set<string>()
  let current = byId.get(objectId)
  while (current) {
    if (seen.has(current.objectId)) throw new RangeError(`Office object hierarchy contains a cycle at ${current.objectId}`)
    seen.add(current.objectId)
    chain.unshift(current)
    current = current.parentObjectId === undefined ? undefined : byId.get(current.parentObjectId)
  }
  return Object.freeze(chain)
}

/** Direct children of one object id in document order. */
export function officeObjectChildren(
  descriptors: readonly OfficeObjectDescriptor[],
  objectId: string,
): readonly OfficeObjectDescriptor[] {
  return Object.freeze(descriptors.filter((descriptor) => descriptor.parentObjectId === objectId))
}

/** One hover/click candidate built from a descriptor, before refinement. */
export interface OfficeObjectCandidate {
  candidateId: string
  descriptor: OfficeObjectDescriptor
  hit: OfficeHitCandidate["hit"]
  depth: number
  zIndex?: number
}

function candidateHit(entry: OfficeObjectCandidate): OfficeHitCandidate {
  return {
    candidateId: entry.candidateId,
    draft: entry.descriptor.draft,
    preview: {
      label: entry.descriptor.label,
      path: entry.descriptor.path,
      ...(entry.descriptor.visual === undefined ? {} : { visual: entry.descriptor.visual }),
    },
    hit: entry.hit,
    depth: entry.depth,
    ...(entry.zIndex === undefined ? {} : { zIndex: entry.zIndex }),
  }
}

/**
 * Deterministic candidate order: direct hits first, then deeper objects, then
 * smaller geometry, then z-order, semantic reliability and source. The sort is
 * reused from the hit-candidate ranker so object mode and the low-level API
 * cannot disagree.
 */
export function rankOfficeObjectCandidates(
  candidates: readonly OfficeObjectCandidate[],
): readonly OfficeObjectCandidate[] {
  const byId = new Map(candidates.map((candidate) => [candidate.candidateId, candidate]))
  return Object.freeze(
    rankOfficeHitCandidates(candidates.map(candidateHit))
      .map((candidate) => byId.get(candidate.candidateId))
      .filter((candidate): candidate is OfficeObjectCandidate => candidate !== undefined),
  )
}

/** Convert one object candidate into the preview Surface events carry. */
export function officeObjectCandidatePreview(
  candidate: OfficeObjectCandidate,
): OfficeReferenceCandidatePreview {
  const hit = candidateHit(candidate)
  return {
    candidateId: hit.candidateId,
    draft: hit.draft,
    ...(hit.preview === undefined ? {} : { preview: hit.preview }),
    hit: hit.hit,
    depth: hit.depth,
    ...(hit.zIndex === undefined ? {} : { zIndex: hit.zIndex }),
  }
}

/**
 * Descriptor reliability copy used by object providers. Kept here so every
 * format emits the same four dimensions without sharing private constants.
 */
export function officeObjectReliability(
  semantic: OfficeObjectReliability["semantic"],
  boundary: OfficeObjectReliability["boundary"],
  hierarchy: OfficeObjectReliability["hierarchy"],
  relocation: OfficeObjectReliability["relocation"],
): OfficeObjectReliability {
  return {
    semantic: { ...semantic, reasonCodes: [...semantic.reasonCodes] },
    boundary: { ...boundary, reasonCodes: [...boundary.reasonCodes] },
    hierarchy: { ...hierarchy, reasonCodes: [...hierarchy.reasonCodes] },
    relocation: { ...relocation, reasonCodes: [...relocation.reasonCodes] },
  }
}
