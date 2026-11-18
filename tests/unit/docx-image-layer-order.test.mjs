// Regression coverage for issue #145 — DOCX image layer order commands.
//
// The issue bullets map to these tests:
//   1. bringToFront / bringForward / sendBackward / sendToBack on the image
//      model ......................................... layer order state + patches
//   2. four actions in the image context menu with correct disabled states
//      ............................................... menu action resolution
//   3. record history and export the new z-order ..... component + export test
//      (controller/history/export is covered in
//      tests/component/docx-image-layer-order.test.mjs)

import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);

const MAX = core.WORD_IMAGE_Z_INDEX_MAX;
const MIN = core.WORD_IMAGE_Z_INDEX_MIN;
const STEP = core.WORD_IMAGE_Z_INDEX_STEP;

function floating(overrides = {}) {
  return { xPx: 0, yPx: 0, behindDocument: false, ...overrides };
}

test("layer order state clamps relativeHeight and defaults like Word", () => {
  const unset = core.resolveDocxImageLayerOrderState(floating());
  assert.equal(unset.floating, true);
  assert.equal(unset.zIndex, MAX, "Word's default relativeHeight is the top");
  assert.equal(unset.canBringToFront, false);
  assert.equal(unset.canBringForward, false);
  assert.equal(unset.canSendBackward, true);
  assert.equal(unset.canSendToBack, true);

  const low = core.resolveDocxImageLayerOrderState(floating({ zIndex: 65536 }));
  assert.equal(low.zIndex, 65536);
  assert.equal(low.canBringToFront, true);
  assert.equal(low.canBringForward, true);
  assert.equal(low.canSendBackward, true);
  assert.equal(low.canSendToBack, true);

  const atMin = core.resolveDocxImageLayerOrderState(floating({ zIndex: MIN }));
  assert.equal(atMin.canSendBackward, false);
  assert.equal(atMin.canSendToBack, false);
  assert.equal(atMin.canBringForward, true);
  assert.equal(atMin.canBringToFront, true);

  // Out-of-range values are clamped before the disabled decision, exactly like
  // the upstream context menu.
  const aboveMax = core.resolveDocxImageLayerOrderState(
    floating({ zIndex: MAX + STEP }),
  );
  assert.equal(aboveMax.zIndex, MAX);
  assert.equal(aboveMax.canBringForward, false);

  const belowMin = core.resolveDocxImageLayerOrderState(
    floating({ zIndex: -STEP }),
  );
  assert.equal(belowMin.zIndex, MIN);
  assert.equal(belowMin.canSendBackward, false);
});

test("a non-floating image has no layer order commands", () => {
  const state = core.resolveDocxImageLayerOrderState(undefined);
  assert.equal(state.floating, false);
  assert.equal(state.canBringToFront, false);
  assert.equal(state.canBringForward, false);
  assert.equal(state.canSendBackward, false);
  assert.equal(state.canSendToBack, false);
  assert.equal(core.resolveDocxImageLayerOrderPatch("bringToFront", undefined), undefined);
});

test("the four commands rewrite zIndex in Word layer steps and clamp at both ends", () => {
  const initial = floating({ zIndex: 131072, behindDocument: true });

  assert.deepEqual(core.resolveDocxImageLayerOrderPatch("bringToFront", initial), {
    behindDocument: false,
    zIndex: MAX,
  });
  assert.deepEqual(core.resolveDocxImageLayerOrderPatch("bringForward", initial), {
    behindDocument: false,
    zIndex: 131072 + STEP,
  });
  assert.deepEqual(core.resolveDocxImageLayerOrderPatch("sendBackward", initial), {
    behindDocument: false,
    zIndex: 131072 - STEP,
  });
  assert.deepEqual(core.resolveDocxImageLayerOrderPatch("sendToBack", initial), {
    behindDocument: false,
    zIndex: MIN,
  });

  assert.equal(
    core.resolveDocxImageLayerOrderPatch("bringForward", floating({ zIndex: MAX })).zIndex,
    MAX,
    "bring forward clamps at the top",
  );
  assert.equal(
    core.resolveDocxImageLayerOrderPatch("sendBackward", floating({ zIndex: MIN })).zIndex,
    MIN,
    "send backward clamps at the bottom",
  );

  // Missing/NaN zIndex starts from Word's default (the top), so only the
  // "send" direction can move the image.
  assert.equal(
    core.resolveDocxImageLayerOrderPatch("sendBackward", floating()).zIndex,
    MAX - STEP,
  );
  assert.equal(
    core.resolveDocxImageLayerOrderPatch("sendBackward", floating({ zIndex: Number.NaN })).zIndex,
    MAX - STEP,
  );
});

test("the default image layer constants produce four menu actions in Arrange order", () => {
  const actions = core.resolveDocxImageLayerOrderMenuActions(
    floating({ zIndex: 131072 }),
  );
  assert.deepEqual(
    actions.map((action) => [action.actionId, action.label, action.disabled]),
    [
      ["image-bring-to-front", "Bring to Front", false],
      ["image-bring-forward", "Bring Forward", false],
      ["image-send-backward", "Send Backward", false],
      ["image-send-to-back", "Send to Back", false],
    ],
  );

  const atTop = core.resolveDocxImageLayerOrderMenuActions(
    floating({ zIndex: MAX }),
  );
  assert.deepEqual(
    atTop.map((action) => action.disabled),
    [true, true, false, false],
    "an image already at the top can only be sent backward",
  );

  const atBottom = core.resolveDocxImageLayerOrderMenuActions(
    floating({ zIndex: MIN }),
  );
  assert.deepEqual(
    atBottom.map((action) => action.disabled),
    [false, false, true, true],
    "an image already at the bottom can only be brought forward",
  );

  const inline = core.resolveDocxImageLayerOrderMenuActions(undefined);
  assert.ok(inline.every((action) => action.disabled));
});

test("layer order action ids map both ways for context-menu dispatch", () => {
  for (const command of ["bringToFront", "bringForward", "sendBackward", "sendToBack"]) {
    const actionId = core.docxImageLayerOrderActionId(command);
    assert.equal(core.docxImageLayerOrderCommandFromActionId(actionId), command);
  }
  assert.equal(
    core.docxImageLayerOrderCommandFromActionId("image-wrap-square"),
    undefined,
  );
});

test("floating image CSS stacking follows relativeHeight and behind-text", () => {
  assert.equal(core.resolveFloatingImageStackingZIndex(undefined), undefined);
  assert.equal(
    core.resolveFloatingImageStackingZIndex(floating({ zIndex: 65536 })),
    1,
  );
  assert.equal(
    core.resolveFloatingImageStackingZIndex(floating({ zIndex: 131072 })),
    2,
  );
  assert.equal(
    core.resolveFloatingImageStackingZIndex(floating({ zIndex: MAX })),
    MAX / STEP,
  );
  assert.ok(
    core.resolveFloatingImageStackingZIndex(
      floating({ zIndex: 65536, behindDocument: true }),
    ) < 0,
    "behind-text images paint below the paragraph text",
  );
});
