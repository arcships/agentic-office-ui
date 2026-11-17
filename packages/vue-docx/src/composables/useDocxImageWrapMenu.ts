// Vue composable: useDocxImageWrapMenu
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31375-31446
//
// Exposes image wrap mode state and menu options from the editor controller.

import {
  DOCX_IMAGE_WRAP_MODE_ACTIONS,
  resolveDocxImageWrapState,
  imageWrapModeActionId,
  type DocxContextMenuRenderProps,
  type DocxImageWrapMode,
  type DocxImageWrapState,
  type DocxImageWrapMenuOption,
} from "@arcships/docx-core"
import { officeMessagesEn } from "@arcships/office-interaction"
import type { OfficeDocxMessages } from "@arcships/office-interaction"

export type DocxImageWrapMessages = OfficeDocxMessages["image"]

function wrapModeLabel(
  mode: DocxImageWrapMode,
  messages: DocxImageWrapMessages,
): string {
  switch (mode) {
    case "inline": return messages.wrapInline
    case "square": return messages.wrapSquare
    case "tight": return messages.wrapTight
    case "through": return messages.wrapThrough
    case "topAndBottom": return messages.wrapTopAndBottom
    case "behindText": return messages.wrapBehind
    case "inFrontOfText": return messages.wrapFront
    default: return mode
  }
}

export interface UseDocxImageWrapMenuResult {
  state: DocxImageWrapState
  wrapOptions: DocxImageWrapMenuOption[]
  positioningOptions: DocxImageWrapMenuOption[]
  editWrapBoundaryOption: DocxImageWrapMenuOption
  moreLayoutOptionsOption: DocxImageWrapMenuOption
  setMode: (mode: DocxImageWrapMode) => void
  setMoveWithText: (moveWithText: boolean) => void
}

export function useDocxImageWrapMenu(
  menu: Pick<DocxContextMenuRenderProps, "context" | "runAction">,
  messages: DocxImageWrapMessages = officeMessagesEn.docx.image,
): UseDocxImageWrapMenuResult | undefined {
  const wrapState =
    menu.context.kind === "image"
      ? menu.context.image?.wrap ??
        resolveDocxImageWrapState(menu.context.image?.floating)
      : undefined

  if (!wrapState) {
    return undefined
  }

  const setMode = (mode: DocxImageWrapMode): void => {
    menu.runAction(imageWrapModeActionId(mode))
  }

  const setMoveWithText = (moveWithText: boolean): void => {
    menu.runAction(
      moveWithText ? "image-move-with-text" : "image-fix-position-on-page"
    )
  }

  return {
    state: wrapState,
    wrapOptions: DOCX_IMAGE_WRAP_MODE_ACTIONS.map((action) => ({
      actionId: action.actionId,
      label: wrapModeLabel(action.mode, messages),
      checked: wrapState.mode === action.mode,
      onSelect: () => menu.runAction(action.actionId),
    })),
    positioningOptions: [
      {
        actionId: "image-move-with-text",
        label: messages.moveWithText,
        checked: wrapState.moveWithText,
        disabled: wrapState.mode === "inline",
        onSelect: () => menu.runAction("image-move-with-text"),
      },
      {
        actionId: "image-fix-position-on-page",
        label: messages.fixPositionOnPage,
        checked: wrapState.fixedPositionOnPage,
        disabled: wrapState.mode === "inline",
        onSelect: () => menu.runAction("image-fix-position-on-page"),
      },
    ],
    editWrapBoundaryOption: {
      actionId: "image-edit-wrap-boundary",
      label: messages.editWrapBoundary,
      checked: false,
      disabled: true,
      onSelect: () => menu.runAction("image-edit-wrap-boundary"),
    },
    moreLayoutOptionsOption: {
      actionId: "image-more-layout-options",
      label: messages.moreLayoutOptions,
      checked: false,
      disabled: true,
      onSelect: () => menu.runAction("image-more-layout-options"),
    },
    setMode,
    setMoveWithText,
  }
}
