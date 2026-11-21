/**
 * Minimal Vue renderer for `@hugeicons/core-free-icons` data.
 *
 * Extend UI renders its upload glyphs with Hugeicons; using the same icon data
 * keeps the FileUpload parity surface pixel-identical. The data package is a
 * runtime dependency (MIT); no SVG path is copied into this repository.
 */
import { defineComponent, h, type PropType, type VNode } from "vue"

/** Shape of the icon data exported by `@hugeicons/core-free-icons`. */
export type HugeIconData = ReadonlyArray<
  readonly [string, Readonly<Record<string, string | number>>]
>

const SVG_ATTRIBUTE_NAMES: Record<string, string> = {
  strokeWidth: "stroke-width",
  strokeLinecap: "stroke-linecap",
  strokeLinejoin: "stroke-linejoin",
  fillRule: "fill-rule",
  clipRule: "clip-rule",
}

const RESERVED_ATTRIBUTES = new Set(["key"])

function normalizeAttributes(
  attributes: Readonly<Record<string, string | number>>,
): Record<string, string | number> {
  const normalized: Record<string, string | number> = {}
  for (const [name, value] of Object.entries(attributes)) {
    if (RESERVED_ATTRIBUTES.has(name)) continue
    normalized[SVG_ATTRIBUTE_NAMES[name] ?? name] = value
  }
  return normalized
}

/**
 * Render Hugeicons icon data as an accessible inline SVG.
 *
 * The SVG keeps the upstream 24x24 viewBox; callers size it with CSS, exactly
 * like the React `HugeiconsIcon` wrapper used by the Extend reference.
 */
export const HugeIcon = defineComponent({
  name: "HugeIcon",
  props: {
    icon: {
      type: Array as PropType<HugeIconData>,
      required: true,
    },
    size: {
      type: [Number, String] as PropType<number | string>,
      default: 24,
    },
  },
  setup(props, { attrs }) {
    return (): VNode =>
      h(
        "svg",
        {
          xmlns: "http://www.w3.org/2000/svg",
          width: props.size,
          height: props.size,
          viewBox: "0 0 24 24",
          fill: "none",
          color: "currentColor",
          "aria-hidden": "true",
          focusable: "false",
          ...attrs,
        },
        props.icon.map(([tag, attributes], index) =>
          h(tag, { ...normalizeAttributes(attributes), key: index }),
        ),
      )
  },
})
