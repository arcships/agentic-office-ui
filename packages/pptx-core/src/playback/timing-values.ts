/**
 * ECMA-376 §19.5.33：`repeatCount` 以千分位表示重复次数，XSD 默认值为 1000（即一次）。
 * 因此 1000 表示 1 次、2000 表示 2 次、500 表示 0.5 次，`indefinite` 保持无限。
 * 这里始终按千分位换算，不再用“是否大于等于 1000”猜测编码：那样会把 999 读成 999 次，
 * 也会让 1500 恰好正确、而 1001 变成 1.001 次这样难以解释的结果。
 */
export function repeatCountValue(raw: string | null | undefined): number | "indefinite" | undefined {
  if (!raw) return undefined
  if (raw === "indefinite") return "indefinite"
  const thousandths = Number(raw)
  return Number.isFinite(thousandths) && thousandths > 0 ? thousandths / 1000 : undefined
}

/** 把原始 `repeatCount` 文本写进能力报告，便于核对千分位编码是否被正确解读。 */
export function repeatCountNote(raw: string | null | undefined, repeatCount: number | "indefinite"): string | undefined {
  if (!raw || repeatCount === "indefinite") return undefined
  return `repeatCount 以千分位记录重复次数（ECMA-376 §19.5.33）：原始值 ${raw} 表示 ${repeatCount} 次。`
}
