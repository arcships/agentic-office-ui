import type { ParagraphNode, TextRunNode } from "./types";

/**
 * Plain text of a paragraph, matching the text the renderer walks.
 *
 * Kept in the engine layer so pure core operations do not import editor
 * helpers (which carry DOM measurement state).
 */
export function paragraphTextContent(paragraph: ParagraphNode): string {
  return paragraph.children
    .filter((child): child is TextRunNode => child.type === "text")
    .map((run) => run.text)
    .join("");
}
