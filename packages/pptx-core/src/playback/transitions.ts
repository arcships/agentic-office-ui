import { PPTX_CAPABILITY_FEATURES } from "./targets"
import type {
  PptxFeatureDisposition,
  PptxMediaItem,
  PptxSlideTransitionSound,
} from "./types"
import type { PptxApproximationPolicy } from "../browser-types"

/**
 * ECMA-376 `p:transition` children that PowerPoint groups as 3D transitions.
 * They are always reported with a stable code and never approximated by the
 * flat fade fallback: a fake 2D fade is less honest than keeping the page.
 */
export const PPTX_3D_TRANSITION_KINDS: ReadonlySet<string> = new Set([
  "conveyor",
  "cube",
  "doors",
  "ferris",
  "ferriswheel",
  "flip",
  "flythrough",
  "gallery",
  "pan",
  "ripple",
  "switch",
  "vortex",
  "warp",
  "window",
])

/**
 * Two-dimensional transitions the controller can execute strictly.
 *
 * `cover`, `uncover` and `split` render exactly like PowerPoint. The wipe layer
 * covers every documented direction (`l`, `r`, `u`, `d` and the four corners),
 * and `push`/`cover`/`uncover` follow the same direction letters.
 */
export const PPTX_STRICT_TRANSITION_KINDS: ReadonlySet<string> = new Set([
  "cut",
  "fade",
  "push",
  "wipe",
  "cover",
  "uncover",
  "split",
])

/**
 * Two-dimensional transitions that are only approximated by a cross fade.
 * They stay static unless the caller opts into `approximation: "safe"`.
 */
export const PPTX_APPROXIMATE_TRANSITION_KINDS: ReadonlySet<string> = new Set([
  "dissolve",
])

export interface PptxTransitionCapability {
  disposition: PptxFeatureDisposition
  /** Stable capability feature name, e.g. `transition:3d:cube`. */
  feature: string
  /** Whether the controller should mount a transition layer. */
  executable: boolean
  reason?: string
}

export function pptxTransitionCapability(
  kind: string,
  approximation: PptxApproximationPolicy = "off",
): PptxTransitionCapability {
  const normalized = kind.trim().toLowerCase() || "cut"
  if (PPTX_STRICT_TRANSITION_KINDS.has(normalized)) {
    return { disposition: "strict", feature: `transition:${normalized}`, executable: true }
  }
  if (PPTX_APPROXIMATE_TRANSITION_KINDS.has(normalized)) {
    return approximation === "safe"
      ? {
          disposition: "approximate",
          feature: `transition:${normalized}`,
          executable: true,
          reason: "像素级溶解以交叉淡入淡出近似执行。",
        }
      : {
          disposition: "static",
          feature: `transition:${normalized}`,
          executable: false,
          reason: "近似播放未开启，保留静态页面。",
        }
  }
  if (normalized === "morph") {
    return approximation === "safe"
      ? {
          disposition: "approximate",
          feature: "transition:morph",
          executable: true,
          reason: "Morph 使用分级对象身份近似执行。",
        }
      : {
          disposition: "static",
          feature: "transition:morph",
          executable: false,
          reason: "近似播放未开启，保留静态页面。",
        }
  }
  if (PPTX_3D_TRANSITION_KINDS.has(normalized)) {
    return {
      disposition: "static",
      feature: `${PPTX_CAPABILITY_FEATURES.transition3d}${normalized}`,
      executable: false,
      reason: "三维切换尚未实现，按静态页面保留；不会退化为二维淡入。",
    }
  }
  return approximation === "safe"
    ? {
        disposition: "approximate",
        feature: `transition:${normalized}`,
        executable: true,
        reason: "未知二维切换使用淡入近似。",
      }
    : {
        disposition: "static",
        feature: `transition:${normalized}`,
        executable: false,
        reason: "近似播放未开启，保留静态页面。",
      }
}

export function isPptxTransitionExecutable(
  kind: string | undefined,
  approximation: PptxApproximationPolicy = "off",
): boolean {
  if (!kind) return false
  return pptxTransitionCapability(kind, approximation).executable
}

/** A unit offset pointing from the slide's final position towards its start. */
export interface PptxTransitionDirection {
  x: number
  y: number
}

/**
 * Direction letters name the edge the incoming slide enters from:
 * `l` slides in from the left, `rd` from the bottom-right corner.
 */
export function pptxTransitionDirection(direction: string | undefined): PptxTransitionDirection {
  switch ((direction ?? "l").trim().toLowerCase()) {
    case "r":
    case "right":
      return { x: -1, y: 0 }
    case "u":
    case "up":
      return { x: 0, y: 1 }
    case "d":
    case "down":
      return { x: 0, y: -1 }
    case "lu":
      return { x: 1, y: 1 }
    case "ld":
      return { x: 1, y: -1 }
    case "ru":
      return { x: -1, y: 1 }
    case "rd":
      return { x: -1, y: -1 }
    default:
      return { x: 1, y: 0 }
  }
}

function clampProgress(progress: number): number {
  return Math.max(0, Math.min(1, progress))
}

/**
 * Four `inset()` percentages (top, right, bottom, left) that keep the band the
 * wipe has already painted. The band always grows from the named edge, so the
 * hidden side is the opposite one.
 */
export function pptxTransitionWipeInsets(
  direction: string | undefined,
  progress: number,
): readonly [number, number, number, number] {
  const hidden = (1 - clampProgress(progress)) * 100
  switch ((direction ?? "l").trim().toLowerCase()) {
    case "r":
    case "right":
      return [0, 0, 0, hidden]
    case "u":
    case "up":
      return [hidden, 0, 0, 0]
    case "d":
    case "down":
      return [0, 0, hidden, 0]
    case "lu":
      return [0, hidden, hidden, 0]
    case "ld":
      return [hidden, hidden, 0, 0]
    case "ru":
      return [0, 0, hidden, hidden]
    case "rd":
      return [hidden, 0, 0, hidden]
    default:
      return [0, hidden, 0, 0]
  }
}

/** `inset()` value for `pptxTransitionWipeInsets`, rounded to 3 decimals. */
export function pptxTransitionWipeClip(
  direction: string | undefined,
  progress: number,
): string {
  const [top, right, bottom, left] = pptxTransitionWipeInsets(direction, progress)
  const value = (percent: number) => `${Math.round(percent * 1_000) / 1_000}%`
  return `inset(${value(top)} ${value(right)} ${value(bottom)} ${value(left)})`
}

/**
 * `mask-image` value for `<p:split>`.
 *
 * `orient` names the line the slide splits along, so `vert` (the default)
 * separates two side-by-side halves and `horz` two stacked halves. `dir="out"`
 * reveals the incoming slide from that line outwards, `dir="in"` keeps the
 * incoming slide at the edges and closes the gap towards the line.
 */
export function pptxTransitionSplitMask(
  orientation: string | undefined,
  direction: string | undefined,
  progress: number,
): string {
  const axis = (orientation ?? "").trim().toLowerCase() === "horz" ? "to bottom" : "to right"
  const amount = clampProgress(progress)
  const round = (value: number) => Math.round(value * 1_000) / 1_000
  if ((direction ?? "out").trim().toLowerCase() === "in") {
    const edge = round(amount * 50)
    const inner = round(100 - amount * 50)
    return `linear-gradient(${axis}, #000 0 ${edge}%, transparent ${edge}% ${inner}%, #000 ${inner}% 100%)`
  }
  const from = round(50 - amount * 50)
  const to = round(50 + amount * 50)
  return `linear-gradient(${axis}, transparent 0 ${from}%, #000 ${from}% ${to}%, transparent ${to}% 100%)`
}

/** Progress of a transition at `positionMs`, clamped into `0…1`. */
export function pptxTransitionProgress(positionMs: number, durationMs: number): number {
  if (!Number.isFinite(durationMs) || durationMs <= 0) return 1
  return clampProgress(positionMs / durationMs)
}

/** Stable id used by the controller snapshot and the blocked-media surface. */
export function pptxTransitionSoundMediaId(slideIndex: number): string {
  return `transition-sound:${slideIndex}`
}

/**
 * Transition sounds reuse the media resolver, so they get exactly the same
 * embedded/external policy and autoplay gating as audio shapes.
 */
export function pptxTransitionSoundMediaItem(
  slideIndex: number,
  sound: PptxSlideTransitionSound,
): PptxMediaItem {
  return Object.freeze({
    id: pptxTransitionSoundMediaId(slideIndex),
    relationId: sound.relationId,
    sourcePath: sound.sourcePath,
    contentType: sound.contentType,
    kind: "audio",
    embedded: sound.embedded,
    autoplay: true,
    loop: sound.loop,
    volume: sound.volume,
    bookmarks: Object.freeze([]),
  })
}
