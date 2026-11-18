// Image layer order commands (upstream editor.tsx 48734-48757, 49373-49450).
//
// Word stores the paint order of floating objects in
// `wp:anchor/@relativeHeight`, which the compact model exposes as
// `ImageRunNode.floating.zIndex`. The four context-menu commands only rewrite
// that value inside Word's 0..251658240 range; the front/behind text layer is
// controlled separately by `behindDocument` ("Behind Text" / "In Front of
// Text") and is not what these commands change.

import type { ImageRunNode } from "../../engine/types";
import {
  DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS,
  WORD_IMAGE_Z_INDEX_DEFAULT,
  WORD_IMAGE_Z_INDEX_MAX,
  WORD_IMAGE_Z_INDEX_MIN,
  WORD_IMAGE_Z_INDEX_STEP,
} from "./constants";
import type {
  DocxContextMenuAction,
  DocxContextMenuActionId,
} from "./editor-types-extra";

/** Image layer order command, one per built-in context-menu action. */
export type DocxImageLayerOrderCommand =
  | "bringToFront"
  | "bringForward"
  | "sendBackward"
  | "sendToBack";

/**
 * What the four layer commands would do to one image.
 *
 * `zIndex` is the current value clamped into Word's range. A missing zIndex
 * uses `WORD_IMAGE_Z_INDEX_DEFAULT` (the Word default), which leaves every
 * "bring" command disabled until the image is sent backward.
 */
export interface DocxImageLayerOrderState {
  /** Only floating images have a paint order. */
  floating: boolean;
  /** Clamped Word `relativeHeight` the commands start from. */
  zIndex: number;
  canBringToFront: boolean;
  canBringForward: boolean;
  canSendBackward: boolean;
  canSendToBack: boolean;
}

/** One flattened default layer action with its resolved disabled state. */
export interface DocxImageLayerOrderMenuAction {
  command: DocxImageLayerOrderCommand;
  actionId: DocxContextMenuActionId;
  /** Default (English) label from `DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS`. */
  label: string;
  disabled: boolean;
}

const COMMAND_ACTION_IDS: Record<
  DocxImageLayerOrderCommand,
  DocxContextMenuActionId
> = {
  bringToFront: "image-bring-to-front",
  bringForward: "image-bring-forward",
  sendBackward: "image-send-backward",
  sendToBack: "image-send-to-back",
};

const ACTION_ID_COMMANDS: Record<string, DocxImageLayerOrderCommand> = {
  "image-bring-to-front": "bringToFront",
  "image-bring-forward": "bringForward",
  "image-send-backward": "sendBackward",
  "image-send-to-back": "sendToBack",
};

// Menu order: front-most first, back-most last, matching Word's Arrange menu.
const MENU_COMMAND_ORDER: DocxImageLayerOrderCommand[] = [
  "bringToFront",
  "bringForward",
  "sendBackward",
  "sendToBack",
];

function clampWordImageZIndex(value: number): number {
  return Math.max(
    WORD_IMAGE_Z_INDEX_MIN,
    Math.min(WORD_IMAGE_Z_INDEX_MAX, Math.round(value))
  );
}

/** Clamp/normalize a model zIndex the same way Word writes `relativeHeight`. */
export function normalizeWordImageZIndex(value: number | undefined): number {
  return Number.isFinite(value)
    ? clampWordImageZIndex(value as number)
    : WORD_IMAGE_Z_INDEX_DEFAULT;
}

/** Resolve command availability for one floating image. */
export function resolveDocxImageLayerOrderState(
  floating?: ImageRunNode["floating"]
): DocxImageLayerOrderState {
  const hasFloatingImage = Boolean(floating);
  const zIndex = normalizeWordImageZIndex(floating?.zIndex);
  return {
    floating: hasFloatingImage,
    zIndex,
    canBringToFront:
      hasFloatingImage && zIndex < WORD_IMAGE_Z_INDEX_MAX,
    canBringForward:
      hasFloatingImage && zIndex < WORD_IMAGE_Z_INDEX_MAX,
    canSendBackward:
      hasFloatingImage && zIndex > WORD_IMAGE_Z_INDEX_MIN,
    canSendToBack:
      hasFloatingImage && zIndex > WORD_IMAGE_Z_INDEX_MIN,
  };
}

/** Whether one command changes anything for the given image. */
export function isDocxImageLayerOrderCommandEnabled(
  state: DocxImageLayerOrderState,
  command: DocxImageLayerOrderCommand
): boolean {
  switch (command) {
    case "bringToFront":
      return state.canBringToFront;
    case "bringForward":
      return state.canBringForward;
    case "sendBackward":
      return state.canSendBackward;
    case "sendToBack":
      return state.canSendToBack;
    default:
      return false;
  }
}

/**
 * The floating patch a layer command writes.
 *
 * Upstream always clears `behindDocument`: the four commands reorder the
 * image inside the front/behind paint layer the user sees; switching that
 * layer is the separate wrap-mode action ("Behind Text" / "In Front of
 * Text"). A non-floating image has no paint order, so there is no patch.
 */
export function resolveDocxImageLayerOrderPatch(
  command: DocxImageLayerOrderCommand,
  floating?: ImageRunNode["floating"]
): Partial<NonNullable<ImageRunNode["floating"]>> | undefined {
  if (!floating) {
    return undefined;
  }

  const currentZIndex = normalizeWordImageZIndex(floating.zIndex);
  switch (command) {
    case "bringToFront":
      return {
        behindDocument: false,
        zIndex: WORD_IMAGE_Z_INDEX_MAX,
      };
    case "bringForward":
      return {
        behindDocument: false,
        zIndex: clampWordImageZIndex(
          currentZIndex + WORD_IMAGE_Z_INDEX_STEP
        ),
      };
    case "sendBackward":
      return {
        behindDocument: false,
        zIndex: clampWordImageZIndex(
          currentZIndex - WORD_IMAGE_Z_INDEX_STEP
        ),
      };
    case "sendToBack":
      return {
        behindDocument: false,
        zIndex: WORD_IMAGE_Z_INDEX_MIN,
      };
    default:
      return undefined;
  }
}

export function docxImageLayerOrderActionId(
  command: DocxImageLayerOrderCommand
): DocxContextMenuActionId {
  return COMMAND_ACTION_IDS[command];
}

export function docxImageLayerOrderCommandFromActionId(
  actionId: DocxContextMenuActionId | (string & {})
): DocxImageLayerOrderCommand | undefined {
  return ACTION_ID_COMMANDS[String(actionId)];
}

/**
 * Flatten `DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS` into the four leaf
 * actions a menu renders, with the disabled state resolved from the image
 * that opened the menu. The default constants remain the single source of
 * action ids and labels.
 */
export function resolveDocxImageLayerOrderMenuActions(
  floating?: ImageRunNode["floating"]
): DocxImageLayerOrderMenuAction[] {
  const state = resolveDocxImageLayerOrderState(floating);
  const defaults = new Map<string, DocxContextMenuAction>();
  for (const action of DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS) {
    if (action.children?.length) {
      for (const child of action.children) {
        if (child.id) defaults.set(String(child.id), child);
      }
      continue;
    }
    if (action.id) defaults.set(String(action.id), action);
  }

  return MENU_COMMAND_ORDER.map((command) => {
    const actionId = COMMAND_ACTION_IDS[command];
    const defaultAction = defaults.get(String(actionId));
    return {
      command,
      actionId,
      label: defaultAction?.label ?? String(actionId),
      disabled: !isDocxImageLayerOrderCommandEnabled(state, command),
    };
  });
}
