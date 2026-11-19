/** Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. */
/**
 * Shared contracts for the Office primitive layer (issue #157).
 *
 * The primitives are token-driven Vue components on top of `reka-ui`. They do
 * not use Tailwind: every visual value comes from the `--office-*` custom
 * properties declared in `packages/vue-ui/src/styles/tokens.css`, with a small
 * set of component-local metrics on `.office-ui`.
 */

import type { OfficeIconName } from "../../icons"

/** Variants shared by Button, IconButton, Badge and Toggle. */
export type OfficeVariant =
  | "default"
  | "secondary"
  | "ghost"
  | "outline"
  | "destructive"
  | "link"

/** Text-control sizes shared by Button, Input, Textarea and Select. */
export type OfficeControlSize = "xs" | "sm" | "default" | "lg"

/** Button-style sizes, including the square icon-only forms. */
export type OfficeButtonSize = OfficeControlSize | "icon" | "icon-sm" | "icon-lg"

/** Dialog / Sheet width presets. */
export type OfficeOverlaySize = "sm" | "default" | "lg" | "xl" | "full"

/** Side used by Tooltip, Popover, DropdownMenu and Sheet. */
export type OfficeSide = "top" | "right" | "bottom" | "left"

/** Alignment used by Tooltip, Popover and DropdownMenu. */
export type OfficeAlign = "start" | "center" | "end"

/** Direction of a Tabs list. */
export type OfficeOrientation = "horizontal" | "vertical"

/** A single option rendered inside `OfficeSelect`. */
export interface OfficeSelectOption {
  value: string
  label: string
  disabled?: boolean
  icon?: OfficeIconName
}

/** Options sharing a heading inside `OfficeSelect`. */
export interface OfficeSelectGroup {
  type: "group"
  label?: string
  options: OfficeSelectOption[]
}

/** Visual divider inside `OfficeSelect`. */
export interface OfficeSelectSeparator {
  type: "separator"
}

/** Anything `OfficeSelect` accepts in its `options` prop. */
export type OfficeSelectEntry =
  | OfficeSelectOption
  | OfficeSelectGroup
  | OfficeSelectSeparator

/** A selectable command in a dropdown or context menu. */
export interface OfficeMenuItem {
  type?: "item"
  /** Stable id reported with `select` and used as the Vue key. */
  id?: string
  label: string
  icon?: OfficeIconName
  /** Display-only keyboard hint, e.g. `⌘K`. */
  shortcut?: string
  disabled?: boolean
  variant?: "default" | "destructive"
  /** Called after the item is selected. */
  onSelect?: (item: OfficeMenuItem) => void
}

/** A checkable command. Values are controlled by the host's `items` array. */
export interface OfficeMenuCheckboxItem {
  type: "checkbox"
  id?: string
  label: string
  icon?: OfficeIconName
  shortcut?: string
  disabled?: boolean
  checked: boolean
  onSelect?: (item: OfficeMenuCheckboxItem) => void
}

/** A visual divider inside a menu. */
export interface OfficeMenuSeparator {
  type: "separator"
  id?: string
}

/** A non-interactive heading inside a menu. */
export interface OfficeMenuLabel {
  type: "label"
  label: string
  id?: string
}

/** A labelled group of menu entries. */
export interface OfficeMenuGroup {
  type: "group"
  label?: string
  id?: string
  items: OfficeMenuEntry[]
}

/** A nested submenu, opened with ArrowRight / hover. */
export interface OfficeMenuSubmenu {
  type: "sub"
  label: string
  icon?: OfficeIconName
  id?: string
  disabled?: boolean
  items: OfficeMenuEntry[]
}

/** Anything a dropdown or context menu can render. */
export type OfficeMenuEntry =
  | OfficeMenuItem
  | OfficeMenuCheckboxItem
  | OfficeMenuSeparator
  | OfficeMenuLabel
  | OfficeMenuGroup
  | OfficeMenuSubmenu

/** A tab descriptor rendered by `OfficeTabs`. */
export interface OfficeTab {
  value: string
  label: string
  icon?: OfficeIconName
  disabled?: boolean
}

/**
 * Join class names without pulling in `clsx`.
 *
 * Falsy values are skipped, which keeps conditional classes readable in the
 * templates (`:class="officeClass('x', active && 'x--active')"`).
 */
export function officeClass(
  ...values: Array<string | false | null | undefined>
): string {
  return values.filter(Boolean).join(" ")
}

/** True when an entry is a plain select option. */
export function isOfficeSelectOption(
  entry: OfficeSelectEntry,
): entry is OfficeSelectOption {
  return !("type" in entry) || !entry.type
}

/** True when a menu entry is a plain command. */
export function isOfficeMenuItem(
  entry: OfficeMenuEntry,
): entry is OfficeMenuItem {
  return !entry.type || entry.type === "item"
}

/** Format a CSS custom-property value from a numeric or string length. */
export function toCssLength(value: number | string): string {
  return typeof value === "number" ? `${value}px` : value
}
