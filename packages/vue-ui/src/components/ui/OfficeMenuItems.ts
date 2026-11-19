/** Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. */
/**
 * Internal recursive renderer for `OfficeDropdownMenu` and `OfficeContextMenu`.
 *
 * The public API is the `items` array; this file maps each entry to the
 * matching Reka UI menu part. Keeping one renderer means both menus stay
 * visually and behaviourally identical (highlight, typeahead, submenu, ARIA).
 */

import { defineComponent, h, type Component, type PropType, type VNode } from "vue"
import {
  ContextMenuCheckboxItem,
  ContextMenuGroup,
  ContextMenuItem,
  ContextMenuItemIndicator,
  ContextMenuLabel,
  ContextMenuSeparator,
  ContextMenuSub,
  ContextMenuSubContent,
  ContextMenuSubTrigger,
  DropdownMenuCheckboxItem,
  DropdownMenuGroup,
  DropdownMenuItem,
  DropdownMenuItemIndicator,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
} from "reka-ui"
import OfficeIcon from "../OfficeIcon.vue"
import type { OfficeIconName } from "../../icons"
import type { OfficeMenuEntry } from "./types"

export type OfficeMenuKind = "dropdown" | "context"

interface MenuParts {
  CheckboxItem: Component
  Group: Component
  Item: Component
  ItemIndicator: Component
  Label: Component
  Separator: Component
  Sub: Component
  SubContent: Component
  SubTrigger: Component
}

const PARTS: Record<OfficeMenuKind, MenuParts> = {
  dropdown: {
    CheckboxItem: DropdownMenuCheckboxItem,
    Group: DropdownMenuGroup,
    Item: DropdownMenuItem,
    ItemIndicator: DropdownMenuItemIndicator,
    Label: DropdownMenuLabel,
    Separator: DropdownMenuSeparator,
    Sub: DropdownMenuSub,
    SubContent: DropdownMenuSubContent,
    SubTrigger: DropdownMenuSubTrigger,
  },
  context: {
    CheckboxItem: ContextMenuCheckboxItem,
    Group: ContextMenuGroup,
    Item: ContextMenuItem,
    ItemIndicator: ContextMenuItemIndicator,
    Label: ContextMenuLabel,
    Separator: ContextMenuSeparator,
    Sub: ContextMenuSub,
    SubContent: ContextMenuSubContent,
    SubTrigger: ContextMenuSubTrigger,
  },
}

export interface OfficeMenuItemsOptions {
  kind: OfficeMenuKind
  items: OfficeMenuEntry[]
  onSelect: (entry: OfficeMenuEntry, event: Event) => void
}

/** Render a list of menu entries for the given Reka menu flavour. */
export function renderOfficeMenuItems(
  options: OfficeMenuItemsOptions,
): VNode[] {
  const parts = PARTS[options.kind]
  return options.items.map((entry, index) => renderEntry(entry, index, parts, options))
}

/**
 * Internal menu body: a padded viewport plus the mapped Reka menu parts.
 * Rendered by `OfficeDropdownMenu` and `OfficeContextMenu` so both share one
 * implementation of separators, groups, submenus and checkbox rows.
 */
export const OfficeMenuList = defineComponent({
  name: "OfficeMenuList",
  props: {
    kind: {
      type: String as PropType<OfficeMenuKind>,
      required: true,
    },
    items: {
      type: Array as PropType<OfficeMenuEntry[]>,
      required: true,
    },
    onSelect: {
      type: Function as PropType<(entry: OfficeMenuEntry, event: Event) => void>,
      required: true,
    },
  },
  setup(props) {
    return () =>
      h(
        "div",
        { class: "office-menu__viewport", "data-slot": "menu-viewport" },
        renderOfficeMenuItems({
          kind: props.kind,
          items: props.items,
          onSelect: props.onSelect,
        }),
      )
  },
})

function entryKey(entry: OfficeMenuEntry, index: number, kind: OfficeMenuKind): string {
  return entry.id ?? `${kind}-${index}`
}

function renderIcon(entry: { icon?: OfficeIconName }): VNode | null {
  if (!entry.icon) return null
  return h(OfficeIcon, { name: entry.icon, class: "office-menu__item-icon" })
}

function renderEntry(
  entry: OfficeMenuEntry,
  index: number,
  parts: MenuParts,
  options: OfficeMenuItemsOptions,
): VNode {
  const key = entryKey(entry, index, options.kind)

  if (entry.type === "separator") {
    return h(parts.Separator, {
      key,
      class: "office-menu__separator",
      "data-slot": "menu-separator",
    })
  }

  if (entry.type === "label") {
    return h(
      parts.Label,
      { key, class: "office-menu__label", "data-slot": "menu-label" },
      () => entry.label,
    )
  }

  if (entry.type === "group") {
    return h(
      parts.Group,
      { key, class: "office-menu__group", "data-slot": "menu-group" },
      () => [
        entry.label
          ? h(
              parts.Label,
              { class: "office-menu__label", "data-slot": "menu-label" },
              () => entry.label,
            )
          : null,
        ...renderOfficeMenuItems({
          ...options,
          items: entry.items,
        }),
      ],
    )
  }

  if (entry.type === "sub") {
    return h(parts.Sub, { key, disabled: entry.disabled }, () => [
      h(
        parts.SubTrigger,
        {
          class: "office-menu__item office-menu__sub-trigger",
          "data-slot": "menu-sub-trigger",
          disabled: entry.disabled,
        },
        () => [
          renderIcon(entry),
          h("span", { class: "office-menu__item-label" }, entry.label),
        ],
      ),
      h(
        parts.SubContent,
        {
          class: "office-ui office-menu__content",
          "data-slot": "menu-sub-content",
          sideOffset: 4,
        },
        () =>
          h("div", { class: "office-menu__viewport" }, renderOfficeMenuItems({
            ...options,
            items: entry.items,
          })),
      ),
    ])
  }

  if (entry.type === "checkbox") {
    return h(
      parts.CheckboxItem,
      {
        key,
        checked: entry.checked,
        class: "office-menu__item",
        "data-slot": "menu-checkbox-item",
        disabled: entry.disabled,
        onSelect: (event: Event) => options.onSelect(entry, event),
      },
      () => [
        h(
          "span",
          { class: "office-menu__item-indicator" },
          () => h(parts.ItemIndicator, null, () => h(OfficeIcon, { name: "check" })),
        ),
        renderIcon(entry),
        h("span", { class: "office-menu__item-label" }, entry.label),
        entry.shortcut
          ? h("span", { class: "office-menu__item-shortcut" }, entry.shortcut)
          : null,
      ],
    )
  }

  return h(
    parts.Item,
    {
      key,
      class: "office-menu__item",
      "data-slot": "menu-item",
      "data-variant": entry.variant ?? "default",
      disabled: entry.disabled,
      onSelect: (event: Event) => options.onSelect(entry, event),
    },
    () => [
      renderIcon(entry),
      h("span", { class: "office-menu__item-label" }, entry.label),
      entry.shortcut
        ? h("span", { class: "office-menu__item-shortcut" }, entry.shortcut)
        : null,
    ],
  )
}
