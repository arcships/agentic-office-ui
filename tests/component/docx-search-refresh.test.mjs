import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { DocxDocumentSurface } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function modelWithText(text) {
  const base = createBlankDocumentModel();
  return { ...base, nodes: [paragraph(text)] };
}

test("DOCX surface keeps results and re-resolves offsets when an edit replaces the model", async (t) => {
  // The shared test renderer mounts Vue into plain host objects, so give them
  // the few layout APIs the surface uses while projecting search highlights.
  const stubs = ["getBoundingClientRect", "querySelectorAll", "scrollTo", "scrollBy"];
  const originals = new Map();
  for (const key of stubs) {
    originals.set(key, Object.getOwnPropertyDescriptor(Object.prototype, key));
    Object.defineProperty(Object.prototype, key, {
      configurable: true,
      value: key === "getBoundingClientRect"
        ? () => ({ left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 })
        : key === "querySelectorAll"
          ? () => []
          : () => {},
    });
  }
  t.after(() => {
    for (const [key, descriptor] of originals) {
      if (descriptor) Object.defineProperty(Object.prototype, key, descriptor);
      else delete Object.prototype[key];
    }
  });

  const model = vue.shallowRef(modelWithText("alpha needle"));
  const surfaceRef = vue.ref();
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(DocxDocumentSurface, { model: model.value, ref: surfaceRef });
    },
  });
  const mounted = await mount(Harness);
  await vue.nextTick();

  const surface = surfaceRef.value;
  assert.ok(surface, "the document surface must expose its search API");
  const first = await surface.search("needle");
  assert.equal(first.matches.length, 1);
  assert.equal(first.matches[0].range.start.offset, 6);
  assert.equal(first.matches[0].range.end.offset, 12);

  // An edit clones the model; the open search must re-run against the new
  // object instead of clearing the result list.
  model.value = modelWithText("longer alpha needle");
  await waitFor(
    () => surface.getSearchState().matches[0]?.range.start.offset === 13,
    "the search result offsets must follow the edited model",
  );

  const refreshed = surface.getSearchState();
  assert.equal(refreshed.query, "needle");
  assert.equal(refreshed.matches.length, 1);
  assert.equal(refreshed.matches[0].match, "needle");
  assert.equal(refreshed.matches[0].range.start.offset, 13);
  assert.equal(refreshed.matches[0].range.end.offset, 19);
  mounted.app.unmount();
});
