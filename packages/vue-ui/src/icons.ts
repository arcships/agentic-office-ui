import { inject, provide, type Component, type InjectionKey } from "vue"
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  ArrowUp,
  Bold,
  Calendar,
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  ChevronsUpDown,
  CircleCheck,
  CircleX,
  Clock,
  Columns2,
  Copy,
  Download,
  Eye,
  EyeOff,
  File,
  FileText,
  Folder,
  FolderOpen,
  Image,
  Info,
  Italic,
  Link,
  List,
  ListOrdered,
  LoaderCircle,
  LayoutGrid,
  Lock,
  Maximize,
  Minimize,
  Minus,
  Moon,
  MoreHorizontal,
  Pause,
  Play,
  Plus,
  Printer,
  Redo2,
  Save,
  Search,
  Settings,
  StickyNote,
  Strikethrough,
  Sun,
  Table,
  Trash2,
  TriangleAlert,
  Underline,
  Undo2,
  Upload,
  X,
} from "@lucide/vue"

/** A renderable icon component (a Lucide icon or a host component). */
export type OfficeIconComponent = Component

/** Map of kebab-case icon name to component. */
export interface OfficeIconRegistry {
  readonly [name: string]: OfficeIconComponent
}

/**
 * Icons shipped with `@arcships/vue-ui`.
 *
 * Only these named imports reach the bundle; the other Lucide icons stay out
 * of a consumer build. Hosts extend the set with `provideOfficeIcons()` or
 * pass a component through `OfficeIcon`'s `icon` prop.
 */
export const OFFICE_ICONS = Object.freeze({
  "align-center": AlignCenter,
  "align-left": AlignLeft,
  "align-right": AlignRight,
  "arrow-down": ArrowDown,
  "arrow-left": ArrowLeft,
  "arrow-right": ArrowRight,
  "arrow-up": ArrowUp,
  bold: Bold,
  calendar: Calendar,
  check: Check,
  "chevron-down": ChevronDown,
  "chevron-left": ChevronLeft,
  "chevron-right": ChevronRight,
  "chevron-up": ChevronUp,
  "chevrons-up-down": ChevronsUpDown,
  "circle-check": CircleCheck,
  "circle-x": CircleX,
  clock: Clock,
  "columns-2": Columns2,
  copy: Copy,
  download: Download,
  eye: Eye,
  "eye-off": EyeOff,
  file: File,
  "file-text": FileText,
  folder: Folder,
  "folder-open": FolderOpen,
  image: Image,
  info: Info,
  italic: Italic,
  link: Link,
  "layout-grid": LayoutGrid,
  list: List,
  "list-ordered": ListOrdered,
  "loader-circle": LoaderCircle,
  lock: Lock,
  maximize: Maximize,
  minimize: Minimize,
  minus: Minus,
  moon: Moon,
  "more-horizontal": MoreHorizontal,
  pause: Pause,
  play: Play,
  plus: Plus,
  printer: Printer,
  "redo-2": Redo2,
  save: Save,
  search: Search,
  settings: Settings,
  "sticky-note": StickyNote,
  strikethrough: Strikethrough,
  sun: Sun,
  table: Table,
  "trash-2": Trash2,
  "triangle-alert": TriangleAlert,
  underline: Underline,
  "undo-2": Undo2,
  upload: Upload,
  x: X,
} satisfies OfficeIconRegistry)

export type OfficeIconName = keyof typeof OFFICE_ICONS

/** Stable list of built-in icon names, useful for docs and pickers. */
export const OFFICE_ICON_NAMES: readonly OfficeIconName[] = Object.freeze(
  Object.keys(OFFICE_ICONS) as OfficeIconName[],
)

export const OFFICE_ICON_REGISTRY_INJECTION_KEY: InjectionKey<OfficeIconRegistry> =
  Symbol("arcships.office-icon-registry")

/**
 * Accept `chevron-down`, `ChevronDown`, `chevron_down` and `chevron down`
 * as the same name so hosts can mirror Lucide's PascalCase exports.
 */
export function normalizeOfficeIconName(name: string): string {
  return name
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, "$1-$2")
    .replace(/([a-zA-Z])(\d)/g, "$1-$2")
    .replace(/[\s_]+/g, "-")
    .replace(/-+/g, "-")
    .toLowerCase()
}

/** Resolve a name against a registry without importing the whole icon pack. */
export function resolveOfficeIcon(
  name: string,
  registry: OfficeIconRegistry = OFFICE_ICONS,
): OfficeIconComponent | undefined {
  const direct = registry[name]
  if (direct) return direct

  const normalized = normalizeOfficeIconName(name)
  if (!normalized) return undefined

  const normalizedKey = registry[normalized]
  if (normalizedKey) return normalizedKey

  for (const [key, icon] of Object.entries(registry)) {
    if (normalizeOfficeIconName(key) === normalized) return icon
  }
  return undefined
}

/**
 * Extend the icon registry for a subtree. Later calls win over inherited
 * entries, so a host can also replace a built-in icon.
 */
export function provideOfficeIcons(icons: OfficeIconRegistry): OfficeIconRegistry {
  const merged = Object.freeze({ ...useOfficeIcons(), ...icons }) satisfies OfficeIconRegistry
  provide(OFFICE_ICON_REGISTRY_INJECTION_KEY, merged)
  return merged
}

/** Read the nearest icon registry; falls back to the built-in set. */
export function useOfficeIcons(): OfficeIconRegistry {
  return inject(OFFICE_ICON_REGISTRY_INJECTION_KEY, OFFICE_ICONS)
}
