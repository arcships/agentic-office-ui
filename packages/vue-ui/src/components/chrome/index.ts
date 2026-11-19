/**
 * Shared Office chrome (issue #158).
 *
 * Token-driven toolbar, page/zoom/search/status controls, empty/error/loading
 * states and the canonical document-sidebar shell used by all four viewers and
 * editors. Nothing here reads an Office file or owns document state: hosts
 * pass values in and receive `update:*` / intent events out.
 *
 * Styling lives in `src/styles/chrome.css` and resolves every visual value
 * from the `--office-*` tokens; there is no Tailwind, preflight or reset.
 */

export { default as OfficeToolbar } from "./OfficeToolbar.vue"
export { default as OfficeToolbarGroup } from "./OfficeToolbarGroup.vue"
export { default as OfficeToolbarSeparator } from "./OfficeToolbarSeparator.vue"
export { default as OfficePageControl } from "./OfficePageControl.vue"
export { default as OfficeZoomControl } from "./OfficeZoomControl.vue"
export { default as OfficeSearchPopover } from "./OfficeSearchPopover.vue"
export { default as OfficeStatusBar } from "./OfficeStatusBar.vue"
export { default as OfficeEmptyState } from "./OfficeEmptyState.vue"
export { default as OfficeErrorState } from "./OfficeErrorState.vue"
export { default as OfficeLoadingState } from "./OfficeLoadingState.vue"
export { default as DocumentViewerSidebar } from "./DocumentViewerSidebar.vue"
export { default as DocumentViewerSidebarSkeleton } from "./DocumentViewerSidebarSkeleton.vue"

export type { OfficeToolbarProps } from "./OfficeToolbar.vue"
export type { OfficeToolbarGroupProps } from "./OfficeToolbarGroup.vue"
export type { OfficeToolbarSeparatorProps } from "./OfficeToolbarSeparator.vue"
export type { OfficePageControlProps } from "./OfficePageControl.vue"
export type { OfficeZoomControlProps } from "./OfficeZoomControl.vue"
export type { OfficeSearchPopoverProps } from "./OfficeSearchPopover.vue"
export type { OfficeStatusBarProps } from "./OfficeStatusBar.vue"
export type { OfficeEmptyStateProps } from "./OfficeEmptyState.vue"
export type { OfficeErrorStateProps } from "./OfficeErrorState.vue"
export type { OfficeLoadingStateProps } from "./OfficeLoadingState.vue"
export type { DocumentViewerSidebarProps } from "./DocumentViewerSidebar.vue"
export type { DocumentViewerSidebarSkeletonProps } from "./DocumentViewerSidebarSkeleton.vue"

export type {
  OfficeSearchStatus,
  OfficeSidebarMode,
  OfficeSidebarResolvedMode,
  OfficeStateIcon,
  OfficeToolbarAlign,
  OfficeToolbarGap,
  OfficeZoomChange,
  OfficeZoomFitMode,
} from "./types"
