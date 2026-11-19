/**
 * Shared types for the Office chrome layer (issue #158).
 *
 * The chrome components are token-driven and framework-neutral apart from
 * Vue: they never read Office files and never own document state. Hosts pass
 * values in and receive `update:*` / intent events out.
 */

import type { OfficeIconName } from "../../icons"

/** Search session state shared by DOCX, XLSX, PPTX and PDF hosts. */
export type OfficeSearchStatus = "idle" | "searching" | "ready" | "error"

/** Requested sidebar layout. `auto` follows the measured container width. */
export type OfficeSidebarMode = "auto" | "inline" | "overlay"

/** Resolved sidebar layout emitted after the ResizeObserver measurement. */
export type OfficeSidebarResolvedMode = "inline" | "overlay"

/** Toolbar cluster spacing: 4px (`sm`) or 8px (`md`). */
export type OfficeToolbarGap = "sm" | "md"

/** Cross-axis alignment for a toolbar cluster. */
export type OfficeToolbarAlign = "start" | "center" | "end"

/** How the current zoom value was produced. */
export type OfficeZoomFitMode = "none" | "width" | "page"

/** Payload of `OfficeZoomControl`'s `change` event. */
export interface OfficeZoomChange {
  /** Zoom percentage (100 = actual size). */
  zoom: number
  fitMode: OfficeZoomFitMode
}

/** Resolved icon default for a state component. */
export type OfficeStateIcon = OfficeIconName
