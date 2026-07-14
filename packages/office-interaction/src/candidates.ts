import type {
  ManualRegionLocator,
  OfficeHitCandidate,
  OfficeReferenceCandidatePreview,
  OfficeVisualFragment,
  ReliabilityLevel,
} from "./types"
import { normalizedRectArea } from "./geometry"

const hitOrder: Record<OfficeHitCandidate["hit"], number> = {
  direct: 0,
  inside: 1,
  inferred: 2,
  ancestor: 3,
}

const reliabilityOrder: Record<ReliabilityLevel, number> = {
  exact: 0,
  likely: 1,
  uncertain: 2,
  unknown: 3,
}

const sourceOrder = { native: 0, structural: 1, visual: 2, manual: 3 } as const

function sheetArea(region: Extract<ManualRegionLocator, { space: "sheet" }>): number {
  const rows = region.end.row - region.start.row + region.end.yOffset - region.start.yOffset
  const cols = region.end.col - region.start.col + region.end.xOffset - region.start.xOffset
  return Math.max(Number.EPSILON, rows) * Math.max(Number.EPSILON, cols)
}

function fragmentArea(fragment: OfficeVisualFragment): number {
  return "rect" in fragment ? normalizedRectArea(fragment.rect) : sheetArea(fragment.region)
}

function candidateArea(candidate: OfficeHitCandidate): number {
  const fragments = candidate.preview?.visual?.fragments
  if (!fragments?.length) return Number.POSITIVE_INFINITY
  return fragments.reduce((total, fragment) => total + fragmentArea(fragment), 0)
}

export function rankOfficeHitCandidates<T>(candidates: readonly OfficeHitCandidate<T>[]): OfficeHitCandidate<T>[] {
  return candidates
    .map((candidate, index) => ({ candidate, index, area: candidateArea(candidate) }))
    .sort((a, b) =>
      hitOrder[a.candidate.hit] - hitOrder[b.candidate.hit] ||
      b.candidate.depth - a.candidate.depth ||
      a.area - b.area ||
      (b.candidate.zIndex ?? 0) - (a.candidate.zIndex ?? 0) ||
      reliabilityOrder[a.candidate.draft.reliability.semantic.level] - reliabilityOrder[b.candidate.draft.reliability.semantic.level] ||
      sourceOrder[a.candidate.draft.source] - sourceOrder[b.candidate.draft.source] ||
      a.candidate.candidateId.localeCompare(b.candidate.candidateId) ||
      a.index - b.index,
    )
    .map(({ candidate }) => candidate)
}

export function toOfficeReferenceCandidatePreview(
  candidate: OfficeHitCandidate,
): OfficeReferenceCandidatePreview {
  const { candidateId, draft, preview, hit, depth, zIndex } = candidate
  return {
    candidateId,
    draft,
    hit,
    depth,
    ...(preview === undefined ? {} : { preview }),
    ...(zIndex === undefined ? {} : { zIndex }),
  }
}
