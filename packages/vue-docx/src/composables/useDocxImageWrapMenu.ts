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
  menu: Pick<DocxContextMenuRenderProps, "context" | "runAction">
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
      label: action.label,
      checked: wrapState.mode === action.mode,
      onSelect: () => menu.runAction(action.actionId),
    })),
    positioningOptions: [
      {
        actionId: "image-move-with-text",
        label: "Move with Text",
        checked: wrapState.moveWithText,
        disabled: wrapState.mode === "inline",
        onSelect: () => menu.runAction("image-move-with-text"),
      },
      {
        actionId: "image-fix-position-on-page",
        label: "Fix Position on Page",
        checked: wrapState.fixedPositionOnPage,
        disabled: wrapState.mode === "inline",
        onSelect: () => menu.runAction("image-fix-position-on-page"),
      },
    ],
    editWrapBoundaryOption: {
      actionId: "image-edit-wrap-boundary",
      label: "Edit Wrap Boundary",
      checked: false,
      disabled: true,
      onSelect: () => menu.runAction("image-edit-wrap-boundary"),
    },
    moreLayoutOptionsOption: {
      actionId: "image-more-layout-options",
      label: "More Layout Options...",
      checked: false,
      disabled: true,
      onSelect: () => menu.runAction("image-more-layout-options"),
    },
    setMode,
    setMoveWithText,
  }
}
