// editor barrel — paragraph, table, run-style, section-break, comment and
// clipboard editing operations.
export * from "./paragraph-ops";
export * from "./table-ops";
export * from "./run-style-ops";
export * from "./section-break-ops";
export * from "./comment-ops";

// Clipboard: the structured payload, the `text/html` render/parse pair and the
// paste insertion op. Parser internals (attribute/CSS stack machine, internal
// run helpers) stay module-private so the public surface stays intentional.
export {
  DOCX_CLIPBOARD_MIME,
  DOCX_CLIPBOARD_VERSION,
  MAX_DOCX_CLIPBOARD_PAYLOAD_CHARS,
  createDocxClipboardPayload,
  docxClipboardParagraphsForRange,
  insertDocxClipboardParagraphs,
  parseDocxClipboardPayload,
  serializeDocxClipboardPayload
} from "./clipboard-format";
export type {
  CreateDocxClipboardPayloadInput,
  DocxClipboardPayload,
  InsertDocxClipboardResult
} from "./clipboard-format";
export {
  MAX_DOCX_CLIPBOARD_HTML_CHARS,
  MAX_DOCX_CLIPBOARD_PARAGRAPHS,
  MAX_DOCX_CLIPBOARD_RUNS,
  clipboardParagraphsFromText,
  clipboardRunCss,
  clipboardTextFromParagraphs,
  decodeClipboardHtmlEntities,
  escapeClipboardHtml,
  renderDocxClipboardHtml,
  sanitizeClipboardColor,
  sanitizeClipboardParagraphs,
  sanitizeClipboardTextStyle
} from "./clipboard-html";
export type {
  DocxClipboardListInfo,
  DocxClipboardParagraph,
  DocxClipboardParseOptions,
  DocxClipboardRun
} from "./clipboard-html";
export { parseDocxClipboardHtml } from "./clipboard-html-parse";
export {
  MAX_DOCX_CLIPBOARD_RTF_CHARS,
  escapeDocxClipboardRtf,
  parseDocxClipboardRtf,
  renderDocxClipboardRtf
} from "./clipboard-rtf";
