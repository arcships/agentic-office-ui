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

/** Two-dimensional transitions the controller can execute strictly. */
export const PPTX_STRICT_TRANSITION_KINDS: ReadonlySet<string> = new Set([
  "cut",
  "fade",
  "push",
  "wipe",
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
  kind: string,
  approximation: PptxApproximationPolicy = "off",
): boolean {
  return pptxTransitionCapability(kind, approximation).executable
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
