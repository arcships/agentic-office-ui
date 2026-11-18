// Page geometry resolution for document blocks and word ranges.
//
// Upstream: @extend-ai/react-docx `react-viewer/src/block-geometry.tsx`.
//
// The semantic block model is produced without a browser; this module measures
// the actual DOM the viewer rendered and projects every source reference into
// page CSS pixels. That is why a box stays aligned at any zoom: the page
// surface's client rect is divided by its CSS size, and page-relative boxes are
// computed in the surface's own (unscaled) coordinate space.

import type {
  BlockBoundingBox,
  DocxBlockFragment,
  DocxBlockGeometry,
  DocxSourceRef,
  DocumentBlock,
  DocumentWord,
} from "./types";

interface TextRangeEntry {
  node: Text;
  start: number;
  end: number;
}

interface TextRangeIndex {
  entries: TextRangeEntry[];
  prefixEnds: number[];
}

interface PageMetrics {
  rect: DOMRect;
  pageIndex: number;
  pageWidth: number;
  pageHeight: number;
  xScale: number;
  yScale: number;
}

interface GeometryCache {
  elementsById?: Map<string, HTMLElement[]>;
  boxes: Map<HTMLElement, BlockBoundingBox | undefined>;
  text: Map<HTMLElement, TextRangeIndex>;
  pages: Map<HTMLElement, PageMetrics | undefined>;
}

const PAGE_SURFACE_SELECTOR = "[data-docx-page-surface]";
const PAGE_WRAPPER_SELECTOR = "[data-docx-page-index]";
const SOURCE_SELECTOR = "[data-docx-source-id]";
const TEXT_RANGE_SELECTOR = "[data-docx-text-start][data-docx-text-end]";

function intersect(a: BlockBoundingBox, b: BlockBoundingBox): BlockBoundingBox {
  return {
    left: Math.max(a.left, b.left),
    top: Math.max(a.top, b.top),
    right: Math.min(a.right, b.right),
    bottom: Math.min(a.bottom, b.bottom),
  };
}

function visibleBox(
  element: HTMLElement,
  surface: HTMLElement,
  bounds?: BlockBoundingBox,
): BlockBoundingBox | undefined {
  const style = getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden") return undefined;
  const rect = bounds ?? element.getBoundingClientRect();
  let box: BlockBoundingBox = {
    left: rect.left,
    top: rect.top,
    right: rect.right,
    bottom: rect.bottom,
  };
  for (
    let parent = element.parentElement;
    parent && parent !== surface;
    parent = parent.parentElement
  ) {
    const parentStyle = getComputedStyle(parent);
    if (parentStyle.display === "none" || parentStyle.visibility === "hidden") {
      return undefined;
    }
    const parentRect = parent.getBoundingClientRect();
    if (/hidden|clip|scroll|auto/.test(parentStyle.overflowX)) {
      box = {
        ...box,
        left: Math.max(box.left, parentRect.left),
        right: Math.min(box.right, parentRect.right),
      };
    }
    if (/hidden|clip|scroll|auto/.test(parentStyle.overflowY)) {
      box = {
        ...box,
        top: Math.max(box.top, parentRect.top),
        bottom: Math.min(box.bottom, parentRect.bottom),
      };
    }
  }
  box = intersect(box, surface.getBoundingClientRect());
  return box.right > box.left && box.bottom > box.top ? box : undefined;
}

function createGeometryCache(): GeometryCache {
  return { boxes: new Map(), text: new Map(), pages: new Map() };
}

function cachedVisibleBox(
  element: HTMLElement,
  surface: HTMLElement,
  cache: GeometryCache,
): BlockBoundingBox | undefined {
  if (!cache.boxes.has(element)) {
    cache.boxes.set(element, visibleBox(element, surface));
  }
  return cache.boxes.get(element);
}

function textNodes(element: HTMLElement, includeNumbering = false): Text[] {
  const nodes: Text[] = [];
  const ignored =
    "[data-docx-generated-text],[data-docx-pagination-ignore],button,textarea,select,[aria-hidden='true'],[data-docx-note-marker],[data-docx-form-field]";
  const excluded = includeNumbering
    ? ignored
    : `[data-docx-numbering-label],${ignored}`;
  const doc = element.ownerDocument;
  if (!doc?.createTreeWalker) return nodes;
  const walker = doc.createTreeWalker(element, 4 /* NodeFilter.SHOW_TEXT */);
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    const parent = node.parentElement;
    if (!parent || parent.closest(excluded)) continue;
    const source = parent.closest<HTMLElement>(SOURCE_SELECTOR);
    if (source && source !== element) {
      const prefix = `${element.dataset.docxSourceId}/child/`;
      const id = source.dataset.docxSourceId ?? "";
      if (!id.startsWith(prefix) || !/^\d+$/.test(id.slice(prefix.length))) {
        continue;
      }
      if (source.matches("[data-docx-image-location],[data-docx-image-child-index],[role='math']") || parent.closest("svg")) {
        continue;
      }
    }
    nodes.push(node);
  }
  return nodes;
}

function textRangeIndex(element: HTMLElement, cache: GeometryCache): TextRangeIndex {
  const cached = cache.text.get(element);
  if (cached) return cached;
  const annotated = [...element.querySelectorAll<HTMLElement>(TEXT_RANGE_SELECTOR)]
    .filter((node) => node.closest(SOURCE_SELECTOR) === element);
  const chunks = new Map<HTMLElement, { offset: number; end: number }>(
    annotated.map((node) => [
      node,
      {
        offset: Number(node.dataset.docxTextStart),
        end: Number(node.dataset.docxTextEnd),
      },
    ]),
  );
  const fallback = { offset: 0, end: Number.POSITIVE_INFINITY };
  const entries: TextRangeEntry[] = [];
  for (const node of textNodes(element)) {
    const chunk = annotated.length
      ? chunks.get(node.parentElement!.closest<HTMLElement>(TEXT_RANGE_SELECTOR)!)
      : fallback;
    if (!chunk) continue;
    const start = chunk.offset;
    chunk.offset += node.length;
    const end = Math.min(chunk.offset, chunk.end);
    if (Number.isFinite(start) && end > start) entries.push({ node, start, end });
  }
  entries.sort((a, b) => a.start - b.start || a.end - b.end);
  let lastEnd = 0;
  const result: TextRangeIndex = {
    entries,
    prefixEnds: entries.map((entry) => (lastEnd = Math.max(lastEnd, entry.end))),
  };
  cache.text.set(element, result);
  return result;
}

function pageMetrics(
  surface: HTMLElement,
  wrapper: HTMLElement,
  cache: GeometryCache,
): PageMetrics | undefined {
  if (!cache.pages.has(surface)) {
    const rect = surface.getBoundingClientRect();
    const style = getComputedStyle(surface);
    const pageWidth = Number.parseFloat(style.width) || surface.offsetWidth;
    const pageHeight = Number.parseFloat(style.height) || surface.offsetHeight;
    cache.pages.set(
      surface,
      rect.width && rect.height && pageWidth && pageHeight
        ? {
            rect,
            pageIndex: Number(wrapper.dataset.docxPageIndex),
            pageWidth,
            pageHeight,
            xScale: rect.width / pageWidth,
            yScale: rect.height / pageHeight,
          }
        : undefined,
    );
  }
  return cache.pages.get(surface);
}

function inputTextBoxes(
  input: HTMLInputElement,
  surface: HTMLElement,
  start = 0,
  end = input.value.length,
): BlockBoundingBox[] {
  if (start >= input.value.length || end <= start) return [];
  const style = getComputedStyle(input);
  const frame = input.getBoundingClientRect();
  const px = (value: string): number => Number.parseFloat(value) || 0;
  const insetX = px(style.borderLeftWidth) + px(style.paddingLeft);
  const insetY = px(style.borderTopWidth) + px(style.paddingTop);
  const paddingX = insetX + px(style.borderRightWidth) + px(style.paddingRight);
  const paddingY = insetY + px(style.borderBottomWidth) + px(style.paddingBottom);
  const width = px(style.width) + (style.boxSizing === "border-box" ? 0 : paddingX);
  const height = px(style.height) + (style.boxSizing === "border-box" ? 0 : paddingY);
  if (!width || !height) return [];
  const xScale = frame.width / width;
  const yScale = frame.height / height;
  const doc = input.ownerDocument;
  const mirror = doc.createElement("span");
  mirror.dataset.docxGeometryMeasure = "true";
  Object.assign(mirror.style, {
    position: "fixed",
    left: "-100000px",
    top: "0",
    visibility: "hidden",
    display: "inline-block",
    whiteSpace: "pre",
    margin: "0",
    padding: "0",
    border: "0",
  });
  for (const property of [
    "font-family",
    "font-size",
    "font-weight",
    "font-style",
    "font-stretch",
    "font-variant",
    "font-kerning",
    "font-feature-settings",
    "font-variation-settings",
    "font-size-adjust",
    "font-optical-sizing",
    "line-height",
    "letter-spacing",
    "word-spacing",
    "text-transform",
    "direction",
  ]) {
    mirror.style.setProperty(property, style.getPropertyValue(property));
  }
  mirror.textContent = input.value;
  doc.body.append(mirror);
  try {
    const measured = mirror.getBoundingClientRect();
    const available = Math.max(0, width - paddingX - measured.width);
    const align = style.textAlign;
    const right =
      align === "right" ||
      (align === "start" && style.direction === "rtl") ||
      (align === "end" && style.direction !== "rtl");
    const left =
      frame.left +
      (insetX + (right ? available : align === "center" ? available / 2 : 0) - input.scrollLeft) *
        xScale;
    const top = frame.top + (insetY + (height - paddingY - measured.height) / 2) * yScale;
    const range = doc.createRange();
    range.setStart(mirror.firstChild!, start);
    range.setEnd(mirror.firstChild!, Math.min(end, input.value.length));
    return [...range.getClientRects()].flatMap((rect) => {
      const box = visibleBox(
        input,
        surface,
        intersect(frame, {
          left: left + (rect.left - measured.left) * xScale,
          top: top + (rect.top - measured.top) * yScale,
          right: left + (rect.right - measured.left) * xScale,
          bottom: top + (rect.bottom - measured.top) * yScale,
        }),
      );
      return box ? [box] : [];
    });
  } finally {
    mirror.remove();
  }
}

function rangeBoxes(
  element: HTMLElement,
  surface: HTMLElement,
  ref: DocxSourceRef,
  cache: GeometryCache,
  expectedText?: string,
): { boxes: BlockBoundingBox[]; mismatch: boolean } {
  const input = element.matches("[data-docx-drop-cap]")
    ? element.querySelector<HTMLInputElement>("input")
    : null;
  if (input) {
    const mismatch =
      expectedText !== undefined && input.value.slice(ref.start, ref.end) !== expectedText;
    return {
      boxes: mismatch ? [] : inputTextBoxes(input, surface, ref.start!, ref.end!),
      mismatch,
    };
  }
  const boxes: BlockBoundingBox[] = [];
  let mismatch = false;
  const start = ref.start!;
  const end = ref.end!;
  const { entries, prefixEnds } = textRangeIndex(element, cache);
  let low = 0;
  let high = entries.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (prefixEnds[middle]! <= start) low = middle + 1;
    else high = middle;
  }
  for (let index = low; index < entries.length && entries[index]!.start < end; index += 1) {
    const { node, start: nodeStart, end: nodeEnd } = entries[index]!;
    const from = Math.max(start, nodeStart);
    const to = Math.min(end, nodeEnd);
    if (to <= from) continue;
    if (
      expectedText !== undefined &&
      node.data.slice(from - nodeStart, to - nodeStart) !==
        expectedText.slice(from - start, to - start)
    ) {
      mismatch = true;
      continue;
    }
    const clip = cachedVisibleBox(node.parentElement!, surface, cache);
    if (!clip) continue;
    const range = node.ownerDocument.createRange();
    range.setStart(node, from - nodeStart);
    range.setEnd(node, to - nodeStart);
    for (const rect of range.getClientRects()) {
      const box = intersect(rect, clip);
      if (box.right > box.left && box.bottom > box.top) boxes.push(box);
    }
  }
  return { boxes, mismatch };
}

function paragraphTextBoxes(
  element: HTMLElement,
  surface: HTMLElement,
  cache: GeometryCache,
): BlockBoundingBox[] {
  const input = element.matches("[data-docx-drop-cap]")
    ? element.querySelector<HTMLInputElement>("input")
    : null;
  if (input) return inputTextBoxes(input, surface);
  const boxes: BlockBoundingBox[] = [];
  let previousExtent: BlockBoundingBox | undefined;
  const nodes = textNodes(element, true);
  for (const picture of element.querySelectorAll<HTMLElement>("[data-docx-numbering-label] img")) {
    if (picture.closest(SOURCE_SELECTOR) !== element) continue;
    const box = cachedVisibleBox(picture, surface, cache);
    if (box) boxes.push(box);
  }
  for (const node of nodes) {
    const start = node.data.search(/\S/u);
    if (start < 0) continue;
    const end = node.data.trimEnd().length;
    const clip = cachedVisibleBox(node.parentElement!, surface, cache);
    if (!clip) continue;
    const range = node.ownerDocument.createRange();
    range.setStart(node, start);
    range.setEnd(node, end);
    const textRects = [...range.getClientRects()];
    range.selectNodeContents(node);
    const fullRects = start > 0 || end < node.length ? [...range.getClientRects()] : textRects;
    for (const rect of textRects) {
      const box = intersect(rect, clip);
      if (box.right <= box.left || box.bottom <= box.top) continue;
      const fullRect = fullRects.find(
        (full) =>
          Math.abs(full.top - rect.top) < 0.5 &&
          Math.abs(full.bottom - rect.bottom) < 0.5 &&
          full.left <= rect.left &&
          full.right >= rect.right,
      );
      const extent = intersect(fullRect ?? rect, clip);
      const previous = boxes[boxes.length - 1];
      if (
        previous &&
        previousExtent &&
        Math.abs(previous.top - box.top) < 0.5 &&
        Math.abs(previous.bottom - box.bottom) < 0.5 &&
        extent.left <= previousExtent.right + 0.5 &&
        extent.right >= previousExtent.left - 0.5
      ) {
        previous.left = Math.min(previous.left, box.left);
        previous.right = Math.max(previous.right, box.right);
        previous.top = Math.min(previous.top, box.top);
        previous.bottom = Math.max(previous.bottom, box.bottom);
        previousExtent = {
          left: Math.min(previousExtent.left, extent.left),
          right: Math.max(previousExtent.right, extent.right),
          top: Math.min(previousExtent.top, extent.top),
          bottom: Math.min(previousExtent.bottom, extent.bottom),
        };
      } else {
        boxes.push(box);
        previousExtent = extent;
      }
    }
  }
  return boxes;
}

function resolveSourceGeometry(
  root: HTMLElement,
  refs: readonly DocxSourceRef[],
  cache: GeometryCache = createGeometryCache(),
  expectedText?: string,
  paragraphContent = false,
): DocxBlockGeometry {
  const fragments: DocxBlockFragment[] = [];
  const unresolvedSourceIds: string[] = [];
  let foundElement = false;
  const issues: DocxBlockGeometry["issues"] = [];
  let textOffset = 0;

  for (const ref of refs) {
    const ranged = ref.start !== undefined || ref.end !== undefined;
    const textLength = Math.max(0, (ref.end ?? 0) - (ref.start ?? 0));
    const sourceText = expectedText?.slice(textOffset, textOffset + textLength);
    textOffset += textLength;
    if (
      ranged &&
      (!Number.isInteger(ref.start) ||
        !Number.isInteger(ref.end) ||
        ref.start! < 0 ||
        ref.end! <= ref.start!)
    ) {
      issues.push({ nodeId: ref.nodeId, reason: "unsupported_range" });
      unresolvedSourceIds.push(ref.nodeId);
      continue;
    }

    const elements = cache.elementsById
      ? cache.elementsById.get(ref.nodeId) ?? []
      : root.querySelectorAll<HTMLElement>(`[data-docx-source-id="${cssEscape(ref.nodeId)}"]`);
    let resolved = false;
    for (const element of elements) {
      const documentId = element.closest<HTMLElement>("[data-docx-document-id]")?.dataset
        .docxDocumentId;
      if (ref.documentId && documentId && documentId !== ref.documentId) {
        issues.push({ nodeId: ref.nodeId, reason: "document_mismatch" });
        continue;
      }
      const surface = element.closest<HTMLElement>(PAGE_SURFACE_SELECTOR);
      const wrapper = surface?.closest<HTMLElement>(PAGE_WRAPPER_SELECTOR);
      if (!surface || !wrapper) {
        issues.push({ nodeId: ref.nodeId, reason: "not_on_page" });
        continue;
      }
      foundElement = true;
      const box = cachedVisibleBox(element, surface, cache);
      if (!box) continue;
      const metrics = pageMetrics(surface, wrapper, cache);
      if (!metrics) continue;
      const range = ranged ? rangeBoxes(element, surface, ref, cache, sourceText) : undefined;
      if (range?.mismatch) {
        issues.push({ nodeId: ref.nodeId, reason: "range_text_mismatch" });
      }
      const boxes = range
        ? range.boxes
        : paragraphContent || element.matches(
              "[data-docx-paragraph-host],[data-docx-section-paragraph-host],[data-docx-table-cell-paragraph-host],[data-docx-source-kind='paragraph'],[data-docx-drop-cap]",
            )
          ? paragraphTextBoxes(element, surface, cache)
          : [box];
      const { rect: page, pageWidth, pageHeight, pageIndex, xScale, yScale } = metrics;
      for (const boxed of boxes) {
        const boundingBox = {
          left: (boxed.left - page.left) / xScale,
          top: (boxed.top - page.top) / yScale,
          right: (boxed.right - page.left) / xScale,
          bottom: (boxed.bottom - page.top) / yScale,
        };
        const { left, top, right, bottom } = boundingBox;
        fragments.push({
          nodeId: ref.nodeId,
          pageIndex,
          pageWidth,
          pageHeight,
          boundingBox,
          polygon: [
            { x: left, y: top },
            { x: right, y: top },
            { x: right, y: bottom },
            { x: left, y: bottom },
          ],
          coordinateSpace: "page-css-px",
        });
        resolved = true;
      }
    }
    if (!resolved) unresolvedSourceIds.push(ref.nodeId);
  }

  return {
    status: fragments.length
      ? unresolvedSourceIds.length || issues.length
        ? "partially_rendered"
        : "rendered"
      : issues.length
        ? "unresolved"
        : foundElement
          ? "not_visible"
          : "not_mounted",
    scope: "mounted_pages",
    issues,
    fragments,
    unresolvedSourceIds,
  };
}

function cssEscape(value: string): string {
  if (typeof CSS !== "undefined" && typeof CSS.escape === "function") {
    return CSS.escape(value);
  }
  return value.replace(/["\\]/g, "\\$&");
}

/** Resolves the given source references against the mounted viewer DOM. */
export function resolveDocxSourceGeometry(
  root: HTMLElement,
  refs: readonly DocxSourceRef[],
): DocxBlockGeometry {
  return resolveSourceGeometry(root, refs);
}

/** Resolves one geometry entry per block id. */
export function resolveDocxBlockGeometries(
  root: HTMLElement,
  blocks: readonly DocumentBlock[],
): Map<string, DocxBlockGeometry> {
  const elementsById = new Map<string, HTMLElement[]>();
  for (const element of root.querySelectorAll<HTMLElement>(SOURCE_SELECTOR)) {
    const id = element.dataset.docxSourceId!;
    const elements = elementsById.get(id) ?? [];
    elements.push(element);
    elementsById.set(id, elements);
  }
  const cache = createGeometryCache();
  cache.elementsById = elementsById;
  return new Map(
    blocks.map((block) => [
      block.id,
      resolveSourceGeometry(
        root,
        block.sourceRefs,
        cache,
        block.sourceRefs.every((ref) => ref.start !== undefined && ref.end !== undefined)
          ? block.content
          : undefined,
        block.details.kind === "paragraph" ||
          block.type === "heading" ||
          block.type === "section_heading" ||
          block.type === "legend" ||
          block.type === "page_number" ||
          block.type === "key_value",
      ),
    ]),
  );
}

function wordTarget(word: DocumentWord): DocumentBlock {
  return {
    object: "block",
    id: word.id,
    type: "text",
    content: word.content,
    details: { kind: "node", blockId: word.blockId },
    sourceRefs: word.sourceRefs,
    children: [],
  };
}

/** Resolves word ranges through their UTF-16 source offsets. */
export function resolveDocxWordGeometries(
  root: HTMLElement,
  words: readonly DocumentWord[],
): Map<string, DocxBlockGeometry> {
  return resolveDocxBlockGeometries(root, words.map(wordTarget));
}

/** Fractional page coordinates for overlay consumers. */
export function normalizedDocxBoundingBox(fragment: DocxBlockFragment): {
  pageIndex: number;
  left: number;
  top: number;
  width: number;
  height: number;
} {
  const { boundingBox: box, pageWidth, pageHeight, pageIndex } = fragment;
  return {
    pageIndex,
    left: box.left / pageWidth,
    top: box.top / pageHeight,
    width: (box.right - box.left) / pageWidth,
    height: (box.bottom - box.top) / pageHeight,
  };
}

export interface DocxPageBlock extends Omit<DocumentBlock, "id" | "parentBlockId" | "children"> {
  id: string;
  sourceBlockId: string;
  parentBlockId?: string;
  metadata: { page: { number: number; width: number; height: number } };
  boundingBox: BlockBoundingBox;
  polygon: Array<{ x: number; y: number }>;
  fragments: DocxBlockFragment[];
}

/** Splits a block into its per-page occurrences, keeping the source identity. */
export function toDocxPageBlocks(
  block: DocumentBlock,
  geometry: DocxBlockGeometry,
): DocxPageBlock[] {
  const sourceIds = new Set(block.sourceRefs.map((ref) => ref.nodeId));
  const pages = new Map<number, DocxBlockFragment[]>();
  for (const fragment of geometry.fragments) {
    if (!sourceIds.has(fragment.nodeId)) continue;
    const list = pages.get(fragment.pageIndex) ?? [];
    list.push(fragment);
    pages.set(fragment.pageIndex, list);
  }
  return [...pages]
    .sort(([a], [b]) => a - b)
    .map(([pageIndex, fragments]) => {
      const box = {
        left: Math.min(...fragments.map((fragment) => fragment.boundingBox.left)),
        top: Math.min(...fragments.map((fragment) => fragment.boundingBox.top)),
        right: Math.max(...fragments.map((fragment) => fragment.boundingBox.right)),
        bottom: Math.max(...fragments.map((fragment) => fragment.boundingBox.bottom)),
      };
      const { children: _children, id, parentBlockId, ...content } = block;
      return {
        ...content,
        id: `${id}/page/${pageIndex}`,
        sourceBlockId: id,
        parentBlockId: parentBlockId ? `${parentBlockId}/page/${pageIndex}` : undefined,
        metadata: {
          page: {
            number: pageIndex + 1,
            width: fragments[0]!.pageWidth,
            height: fragments[0]!.pageHeight,
          },
        },
        boundingBox: box,
        polygon: [
          { x: box.left, y: box.top },
          { x: box.right, y: box.top },
          { x: box.right, y: box.bottom },
          { x: box.left, y: box.bottom },
        ],
        fragments,
      };
    });
}

/**
 * Scrolls the viewer to a source block or word range. Range lookups mount the
 * page through a temporary 1px target at the range centre; callers that render
 * virtualized pages should retry after the page mounts (handled internally for
 * a short frame budget).
 */
export function revealDocxSource(
  root: HTMLElement,
  source: string | DocxSourceRef,
): void {
  const nodeId = typeof source === "string" ? source : source.nodeId;
  const ref =
    typeof source === "string" || source.start === undefined || source.end === undefined
      ? undefined
      : source;
  const doc = root.ownerDocument;
  const win = doc?.defaultView;
  const revealRange = (): boolean => {
    if (!ref) return false;
    const fragment = resolveSourceGeometry(root, [ref]).fragments[0];
    if (!fragment) return false;
    const surface = root.querySelector<HTMLElement>(
      `[data-docx-page-index="${fragment.pageIndex}"] ${PAGE_SURFACE_SELECTOR}`,
    );
    if (!surface) return false;
    const target = doc.createElement("span");
    target.setAttribute("aria-hidden", "true");
    target.setAttribute("data-docx-pagination-ignore", "true");
    const box = fragment.boundingBox;
    Object.assign(target.style, {
      position: "absolute",
      width: "1px",
      height: "1px",
      pointerEvents: "none",
      left: `${(box.left + box.right) / 2}px`,
      top: `${(box.top + box.bottom) / 2}px`,
    });
    surface.append(target);
    target.scrollIntoView({ block: "center", inline: "nearest" });
    target.remove();
    return true;
  };
  if (ref) {
    if (revealRange()) return;
  } else {
    const element = root.querySelector<HTMLElement>(
      `[data-docx-source-id="${cssEscape(nodeId)}"]`,
    );
    if (element) {
      element.scrollIntoView({ block: "center", inline: "nearest" });
      return;
    }
  }
  const viewer = root.querySelector<HTMLElement>("[data-docx-editor-root]");
  (viewer ?? root).dispatchEvent(
    new CustomEvent("docx-reveal-source", {
      detail: { nodeId, start: ref?.start, end: ref?.end },
    }),
  );
  if (ref && win?.requestAnimationFrame) {
    let attempts = 0;
    const retry = (): void => {
      if (!root.isConnected || revealRange()) return;
      attempts += 1;
      if (attempts < 30) win.requestAnimationFrame(retry);
    };
    win.requestAnimationFrame(retry);
  }
}
