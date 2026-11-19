// Issue #163 — FileUpload / FileThumbnail parity and interaction contract.
//
// The repository's lightweight host renderer drives the component handlers
// directly (no browser). Drag depth, selection/rejection, file-list rows and
// thumbnail reveal/fallback states are covered here; static CSS assertions
// cover the coarse-pointer target and long-name truncation that the renderer
// cannot compute.
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { FileThumbnail, FileUpload } = await importFromDemo("@arcships/vue-ui");

const ROOT = new URL("../../", import.meta.url);

function source(relativePath) {
  return readFileSync(new URL(relativePath, ROOT), "utf8");
}

function classNames(value) {
  if (Array.isArray(value)) return value.map(classNames).filter(Boolean).join(" ");
  if (value && typeof value === "object") {
    return Object.entries(value)
      .filter(([, enabled]) => enabled)
      .map(([name]) => name)
      .join(" ");
  }
  return String(value ?? "");
}

function byClass(root, className) {
  return walk(root).filter((node) =>
    classNames(node.props?.class).split(/\s+/).includes(className),
  );
}

function dropzone(root) {
  const node = walk(root).find((candidate) => candidate.props?.role === "button");
  assert.ok(node, "FileUpload must render a role=button drop zone");
  return node;
}

function dragEvent(files = []) {
  return {
    preventDefault() {},
    dataTransfer: {
      files,
      dropEffect: "",
    },
  };
}

function fileInput(root) {
  const input = walk(root).find(
    (node) => node.type === "input" && node.props?.type === "file",
  );
  assert.ok(input, "FileUpload must render a file input");
  return input;
}

function selectFiles(root, files) {
  const input = fileInput(root);
  const target = { files, value: "selected" };
  input.props.onChange({ target });
  return target;
}

function imageLoad(naturalWidth = 120, naturalHeight = 90) {
  return { currentTarget: { naturalWidth, naturalHeight } };
}

test("FileUpload renders the Extend structure with host-compatible defaults", async () => {
  const mounted = await mount(FileUpload, {
    accept: ".pdf",
    title: "Click to upload or drop files",
    description: "PDF, DOC, XLSX",
    showFileList: true,
  });

  assert.equal(byClass(mounted.root, "file-upload__icon-card").length, 3);
  assert.match(textContent(mounted.root), /Click to upload or drop files/);
  assert.match(textContent(mounted.root), /PDF, DOC, XLSX/);
  assert.match(textContent(mounted.root), /Browse files/);
  assert.equal(dropzone(mounted.root).props.tabindex, "0");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("FileUpload keeps the internal file list opt-in for existing hosts", async () => {
  const accepted = [];
  const legacy = await mount(FileUpload, {
    accept: ".pdf",
    onFilesAccepted: (files) => accepted.push(files),
  });
  selectFiles(legacy.root, [
    { name: "quarterly-report.pdf", type: "application/pdf", size: 14, lastModified: 1 },
  ]);
  await vue.nextTick();
  assert.equal(accepted.length, 1);
  assert.equal(byClass(legacy.root, "file-upload__list").length, 0);
  legacy.app.unmount();
});

test("FileUpload tracks drag depth so nested dragleave does not drop the state", async () => {
  const mounted = await mount(FileUpload, {});
  const zone = dropzone(mounted.root);

  zone.props.onDragenter(dragEvent());
  zone.props.onDragenter(dragEvent());
  await vue.nextTick();
  assert.match(classNames(zone.props.class), /is-dragging/);
  zone.props.onDragleave(dragEvent());
  await vue.nextTick();
  assert.match(classNames(zone.props.class), /is-dragging/);
  zone.props.onDragleave(dragEvent());
  await vue.nextTick();
  assert.doesNotMatch(classNames(zone.props.class), /is-dragging/);

  zone.props.onDragenter(dragEvent());
  zone.props.onDrop(
    dragEvent([
      { name: "contract.pdf", type: "application/pdf", size: 10, lastModified: 1 },
    ]),
  );
  await vue.nextTick();
  assert.doesNotMatch(classNames(zone.props.class), /is-dragging/);
  mounted.app.unmount();
});

test("FileUpload drops, selects and lists multiple accepted files", async () => {
  const accepted = [];
  const mounted = await mount(FileUpload, {
    accept: ".pdf,.png",
    multiple: true,
    showFileList: true,
    onFilesAccepted: (files) => accepted.push(files),
  });
  const files = [
    { name: "quarterly-report.pdf", type: "application/pdf", size: 14, lastModified: 1 },
    { name: "contract-page.png", type: "image/png", size: 2048, lastModified: 2 },
  ];
  selectFiles(mounted.root, files);
  await vue.nextTick();

  assert.equal(accepted.length, 1);
  assert.deepEqual(
    accepted[0].map((file) => file.name),
    ["quarterly-report.pdf", "contract-page.png"],
  );
  assert.equal(byClass(mounted.root, "file-upload__row").length, 2);
  assert.match(textContent(mounted.root), /quarterly-report\.pdf/);
  assert.match(textContent(mounted.root), /application\/pdf - 14 B/);
  assert.match(textContent(mounted.root), /image\/png - 2\.0 KB/);
  assert.match(textContent(mounted.root), /Ready/);

  dropzone(mounted.root).props.onDrop(dragEvent(files));
  await vue.nextTick();
  assert.equal(byClass(mounted.root, "file-upload__row").length, 2);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("FileUpload reports rejected files and shows the reference type copy", async () => {
  const rejected = [];
  const mounted = await mount(FileUpload, {
    accept: ".pdf",
    showFileList: true,
    onFilesRejected: (items) => rejected.push(items),
  });
  selectFiles(mounted.root, [
    { name: "notes.txt", type: "text/plain", size: 14, lastModified: 1 },
  ]);
  await vue.nextTick();

  assert.equal(rejected.length, 1);
  assert.equal(rejected[0][0].code, "FILE_TYPE_NOT_ACCEPTED");
  assert.match(textContent(mounted.root), /This file type is not supported here\./);
  assert.doesNotMatch(textContent(mounted.root), /notes\.txt is not an accepted file type/);
  assert.equal(byClass(mounted.root, "file-upload__row").length, 0);
  mounted.app.unmount();
});

test("FileUpload keeps detailed rejection messages for size and mixed batches", async () => {
  const rejected = [];
  const mounted = await mount(FileUpload, {
    accept: ".pdf",
    maxSize: 8,
    onFilesRejected: (items) => rejected.push(items),
  });
  selectFiles(mounted.root, [
    { name: "ok.pdf", type: "application/pdf", size: 4, lastModified: 1 },
    { name: "large.pdf", type: "application/pdf", size: 12, lastModified: 2 },
    { name: "notes.txt", type: "text/plain", size: 4, lastModified: 3 },
  ]);
  await vue.nextTick();

  assert.deepEqual(
    rejected[0].map(({ code, file }) => ({ code, name: file.name })),
    [
      { code: "FILE_TOO_LARGE", name: "large.pdf" },
      { code: "FILE_TYPE_NOT_ACCEPTED", name: "notes.txt" },
    ],
  );
  const text = textContent(mounted.root);
  assert.match(text, /notes\.txt.*not an accepted file type/);
  assert.match(text, /large\.pdf.*size limit/);
  mounted.app.unmount();
});

test("FileUpload ignores drag, drop, click and keyboard input while disabled", async () => {
  const accepted = [];
  const mounted = await mount(FileUpload, {
    disabled: true,
    onFilesAccepted: (files) => accepted.push(files),
  });
  const zone = dropzone(mounted.root);
  assert.equal(zone.props["aria-disabled"], "true");

  zone.props.onDragenter(dragEvent());
  await vue.nextTick();
  assert.doesNotMatch(classNames(zone.props.class), /is-dragging/);
  zone.props.onDrop(
    dragEvent([
      { name: "contract.pdf", type: "application/pdf", size: 10, lastModified: 1 },
    ]),
  );
  zone.props.onClick();
  const keydown = zone.props.onKeydown;
  if (typeof keydown === "function") keydown(new Event("keydown"));
  await vue.nextTick();
  assert.equal(accepted.length, 0);
  assert.match(classNames(zone.props.class), /is-disabled/);
  mounted.app.unmount();
});

test("FileUpload truncates long file names without breaking the row layout", async () => {
  const longName = `${"very-long-contract-name-".repeat(8)}.pdf`;
  const mounted = await mount(FileUpload, {
    accept: ".pdf",
    showFileList: true,
  });
  selectFiles(mounted.root, [
    { name: longName, type: "application/pdf", size: 14, lastModified: 1 },
  ]);
  await vue.nextTick();

  const name = byClass(mounted.root, "file-upload__file-name")[0];
  assert.ok(name, "file name row must render");
  assert.equal(name.props.title, longName);
  assert.equal(
    walk(name).find((node) => node.type === "#text")?.text,
    longName,
  );

  const uploadCss = source("packages/vue-ui/src/components/FileUpload.vue");
  assert.match(uploadCss, /\.file-upload__file-name[\s\S]*?text-overflow:\s*ellipsis/u);
  assert.match(uploadCss, /\.file-upload__file-name[\s\S]*?white-space:\s*nowrap/u);
  mounted.app.unmount();
});

test("Hugeicons stays a runtime dependency and is external in the library build", () => {
  const packageJson = JSON.parse(
    source("packages/vue-ui/package.json"),
  );
  assert.equal(packageJson.dependencies["@hugeicons/core-free-icons"], "4.4.0");
  const viteConfig = source("packages/vue-ui/vite.config.ts");
  assert.match(viteConfig, /"@hugeicons\/core-free-icons"/u);
  assert.match(source("packages/vue-ui/src/components/FileUpload.vue"), /HugeIcon/u);
});

test("FileUpload reserves a 44px coarse-pointer target for the zone and rows", () => {
  const uploadCss = source("packages/vue-ui/src/components/FileUpload.vue");
  assert.match(uploadCss, /@media \(pointer: coarse\)/u);
  assert.match(uploadCss, /--office-ui-touch-target/u);
  assert.match(
    uploadCss,
    /@media \(pointer: coarse\)[\s\S]*?\.file-upload__dropzone[\s\S]*?min-height:\s*max\(11rem, var\(--office-ui-touch-target\)\)/u,
  );
  assert.match(
    uploadCss,
    /@media \(pointer: coarse\)[\s\S]*?\.file-upload__chip,[\s\S]*?\.file-upload__row[\s\S]*?min-height:\s*var\(--office-ui-touch-target\)/u,
  );
});

test("file component styles stay scoped, token-driven and Tailwind-free", () => {
  for (const relative of [
    "packages/vue-ui/src/components/FileUpload.vue",
    "packages/vue-ui/src/components/FileThumbnail.vue",
  ]) {
    const component = source(relative);
    const style = [...component.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/gu)]
      .map((match) => match[1])
      .join("\n");
    assert.ok(style.length > 0, `${relative} must ship scoped styles`);
    assert.doesNotMatch(style, /@apply|tailwind|preflight/u);
    assert.match(style, /var\(--office-/u);
  }
});

test("FileThumbnail renders loading, error, fallback and type glyph layers", async () => {
  const loading = await mount(FileThumbnail, {
    file: { name: "contract.pdf", type: "application/pdf" },
    isLoading: true,
  });
  assert.equal(byClass(loading.root, "file-thumbnail__loading").length, 1);
  assert.equal(byClass(loading.root, "file-thumbnail__fallback").length, 0);
  loading.app.unmount();

  const error = await mount(FileThumbnail, {
    file: { name: "contract.pdf", type: "application/pdf" },
    hasError: true,
  });
  assert.equal(byClass(error.root, "file-thumbnail__loading").length, 0);
  assert.equal(byClass(error.root, "file-thumbnail__fallback").length, 1);
  assert.equal(walk(error.root).filter((node) => node.type === "svg").length, 0);
  error.app.unmount();

  const glyph = await mount(FileThumbnail, {
    file: { name: "contract.pdf", type: "application/pdf" },
    hasError: true,
    showTypeGlyph: true,
  });
  assert.equal(walk(glyph.root).filter((node) => node.type === "svg").length, 1);
  glyph.app.unmount();
});

test("FileThumbnail supports configurable aspect ratios and long meta text", async () => {
  const longName = `${"very-long-contract-name-".repeat(8)}.bin`;
  const mounted = await mount(FileThumbnail, {
    file: { name: longName, type: "" },
    size: 120,
    previewAspectRatio: 16 / 9,
    showMeta: true,
  });
  const root = walk(mounted.root).find(
    (node) => node.props?.["data-slot"] === "file-thumbnail",
  );
  const preview = walk(mounted.root).find(
    (node) => node.props?.["data-slot"] === "file-thumbnail-preview",
  );
  assert.equal(root.props.style.width, "120px");
  assert.equal(preview.props.style.aspectRatio, String(16 / 9));
  const text = textContent(mounted.root);
  assert.match(text, new RegExp(longName.replace(/\./gu, "\\.")));
  assert.match(text, /Unknown type/);
  mounted.app.unmount();
});

test("FileThumbnail hides the loading overlay after a successful image reveal", async () => {
  const mounted = await mount(FileThumbnail, {
    file: { name: "contract-page.png", type: "image/png" },
    previewImageUrl: "/samples/contract-page.png",
  });
  assert.equal(byClass(mounted.root, "file-thumbnail__loading").length, 1);

  const image = walk(mounted.root).find((node) => node.type === "img");
  assert.ok(image, "image preview must render");
  image.props.onLoad(imageLoad());
  await waitFor(
    () => byClass(mounted.root, "file-thumbnail__loading").length === 0,
    "image loading overlay reveal",
  );
  assert.equal(byClass(mounted.root, "file-thumbnail__fallback").length, 0);
  mounted.app.unmount();
});

test("FileThumbnail falls back when the image fails and recovers for a new URL", async () => {
  const mounted = await mount(FileThumbnail, {
    file: { name: "contract-page.png", type: "image/png" },
    previewImageUrl: "/samples/broken.png",
  });
  const image = walk(mounted.root).find((node) => node.type === "img");
  image.props.onError();
  await vue.nextTick();
  assert.equal(byClass(mounted.root, "file-thumbnail__fallback").length, 1);

  await mounted.app.unmount();
  const next = await mount(FileThumbnail, {
    file: { name: "contract-page.png", type: "image/png" },
    previewImageUrl: "/samples/contract-page.png",
  });
  assert.equal(byClass(next.root, "file-thumbnail__loading").length, 1);
  next.app.unmount();
});
