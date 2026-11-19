/** Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. */
/**
 * Office primitive layer (issue #157).
 *
 * Token-driven Vue 3 components on top of `reka-ui`:
 * Button, IconButton, Input, Textarea, Select, DropdownMenu, ContextMenu,
 * Dialog, Sheet, Popover, Tooltip, Tabs, Toggle, Badge, Card, Separator,
 * Skeleton, Spinner, ScrollArea and Toast.
 *
 * Styling lives in `src/styles/ui.css`; every visual value resolves from the
 * `--office-*` design tokens (plus `.office-ui` local spacing metrics).
 */

export { default as OfficeButton } from "./OfficeButton.vue"
export { default as OfficeIconButton } from "./OfficeIconButton.vue"
export { default as OfficeInput } from "./OfficeInput.vue"
export { default as OfficeTextarea } from "./OfficeTextarea.vue"
export { default as OfficeSelect } from "./OfficeSelect.vue"
export { default as OfficeDropdownMenu } from "./OfficeDropdownMenu.vue"
export { default as OfficeContextMenu } from "./OfficeContextMenu.vue"
export { default as OfficeDialog } from "./OfficeDialog.vue"
export { default as OfficeSheet } from "./OfficeSheet.vue"
export { default as OfficePopover } from "./OfficePopover.vue"
export { default as OfficeTooltip } from "./OfficeTooltip.vue"
export { default as OfficeTabs } from "./OfficeTabs.vue"
export { default as OfficeToggle } from "./OfficeToggle.vue"
export { default as OfficeBadge } from "./OfficeBadge.vue"
export { default as OfficeCard } from "./OfficeCard.vue"
export { default as OfficeSeparator } from "./OfficeSeparator.vue"
export { default as OfficeSkeleton } from "./OfficeSkeleton.vue"
export { default as OfficeSpinner } from "./OfficeSpinner.vue"
export { default as OfficeScrollArea } from "./OfficeScrollArea.vue"
export { default as OfficeToastProvider } from "./OfficeToastProvider.vue"

export { useOfficeToast } from "./office-toast"
export type {
  OfficeToastApi,
  OfficeToastOptions,
  OfficeToastPromiseOptions,
  OfficeToastStatus,
  OfficeToastUpdate,
} from "./office-toast"

export type {
  OfficeAlign,
  OfficeButtonSize,
  OfficeControlSize,
  OfficeMenuCheckboxItem,
  OfficeMenuEntry,
  OfficeMenuGroup,
  OfficeMenuItem,
  OfficeMenuLabel,
  OfficeMenuSeparator,
  OfficeMenuSubmenu,
  OfficeOrientation,
  OfficeOverlaySize,
  OfficeSelectEntry,
  OfficeSelectGroup,
  OfficeSelectOption,
  OfficeSelectSeparator,
  OfficeSide,
  OfficeTab,
  OfficeVariant,
} from "./types"

export type { OfficeButtonProps } from "./OfficeButton.vue"
export type { OfficeIconButtonProps } from "./OfficeIconButton.vue"
export type { OfficeInputProps } from "./OfficeInput.vue"
export type { OfficeTextareaProps } from "./OfficeTextarea.vue"
export type { OfficeSelectProps } from "./OfficeSelect.vue"
export type { OfficeDropdownMenuProps } from "./OfficeDropdownMenu.vue"
export type { OfficeContextMenuProps } from "./OfficeContextMenu.vue"
export type { OfficeDialogProps } from "./OfficeDialog.vue"
export type { OfficeSheetProps } from "./OfficeSheet.vue"
export type { OfficePopoverProps } from "./OfficePopover.vue"
export type { OfficeTooltipProps } from "./OfficeTooltip.vue"
export type { OfficeTabsProps } from "./OfficeTabs.vue"
export type { OfficeToggleProps } from "./OfficeToggle.vue"
export type { OfficeBadgeProps } from "./OfficeBadge.vue"
export type { OfficeCardProps } from "./OfficeCard.vue"
export type { OfficeSeparatorProps } from "./OfficeSeparator.vue"
export type { OfficeSkeletonProps } from "./OfficeSkeleton.vue"
export type { OfficeSpinnerProps } from "./OfficeSpinner.vue"
export type { OfficeScrollAreaProps } from "./OfficeScrollArea.vue"
export type { OfficeToastProviderProps } from "./OfficeToastProvider.vue"
export type { OfficeToastViewportProps } from "./OfficeToastViewport.vue"
