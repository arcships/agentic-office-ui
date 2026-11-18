import { compilePptxSlideSchedule } from "./time-tree"
import type { PptxMediaItem, PptxTimeNode } from "./types"

/**
 * 媒体命令的纯文本规则，解析器、轨道编译器和浏览器播放器共用同一套判断，
 * 避免“解析出来的命令”和“执行时的命令”各有一套解释。
 */
export function normalizeMediaCommand(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined
  const command = value.trim().toLowerCase().replace(/\s+/gu, "")
  if (!command) return undefined
  if (command.includes("togglepause")) return "toggle-pause"
  if (command.includes("pause")) return "pause"
  if (command.includes("stop")) return "stop"
  // ECMA-376 §19.5.28 的 reserved verb `resume` 等价于从当前位置继续播放。
  if (command === "resume") return "play"
  if (command.includes("play")) return "play"
  return command
}

/** `play`、`playFrom(...)` 等命令都表示开始播放。 */
export function isMediaPlayCommand(value: unknown): boolean {
  return normalizeMediaCommand(value) === "play" || mediaSeekMs(value) !== undefined
}

/** `playFrom(2.5)` 表示从媒体第 2.5 秒开始播放；其他命令不产生定位。 */
export function mediaSeekMs(value: unknown): number | undefined {
  if (typeof value !== "string") return undefined
  const match = /^play\s*from\s*\(\s*(-?\d+(?:\.\d+)?)\s*\)$/iu.exec(value.trim())
  if (!match) return undefined
  const seconds = Number(match[1])
  return Number.isFinite(seconds) ? Math.max(0, seconds * 1000) : undefined
}

export interface PptxMediaLoopPolicy {
  /** 是否交给浏览器原生 `loop`。只有完全没有裁剪时才可靠。 */
  nativeLoop: boolean
  /** 手动循环（或原生循环结束时）回到的时间点，单位毫秒。 */
  restartMs?: number
}

/**
 * 只要设置了裁剪起点或终点，原生 `loop` 就会跳回 0，重播被裁掉的开头。
 * 因此裁剪过的媒体一律关闭原生循环，改由 `restartMs` 手动回绕。
 */
export function pptxMediaLoopPolicy(
  item: Pick<PptxMediaItem, "loop" | "trimStartMs" | "trimEndMs">,
): PptxMediaLoopPolicy {
  if (!item.loop) return { nativeLoop: false }
  const trimmed = item.trimStartMs !== undefined || item.trimEndMs !== undefined
  return trimmed
    ? { nativeLoop: false, restartMs: Math.max(0, item.trimStartMs ?? 0) }
    : { nativeLoop: true }
}

/**
 * 媒体是否随页面自动开始：主时间树里未绑定单击的 `p:cmd` 播放命令就是“自动播放”。
 * 与时间表使用同一套规则，解析器和控制器不会对同一页得出不同结论。
 */
export function withAutoplayMedia(
  media: readonly PptxMediaItem[],
  nodes: Readonly<Record<string, PptxTimeNode>>,
  rootNodeId: string | undefined,
): readonly PptxMediaItem[] {
  if (!rootNodeId || !nodes[rootNodeId]) return media
  const automaticTargets = new Set<string>()
  for (const group of compilePptxSlideSchedule({ rootNodeId, nodes }).groups) {
    if (group.key !== "auto") continue
    for (const scheduled of group.effects) {
      const target = scheduled.effect.targetObjectKey
      if (target && scheduled.effect.kind === "media-command" && isMediaPlayCommand(scheduled.effect.command)) {
        automaticTargets.add(target)
      }
    }
  }
  if (automaticTargets.size === 0) return media
  return Object.freeze(media.map((item) => item.autoplay || !item.objectKey || !automaticTargets.has(item.objectKey)
    ? item
    : Object.freeze({ ...item, autoplay: true })))
}
