// clipboard-rtf.ts — RTF flavour of the DOCX clipboard.
//
// Some desktop editors (and older Office builds) prefer `text/rtf` over
// `text/html`, so a copy also writes a compact RTF document built from the same
// clipboard runs. Paste reads RTF after the internal payload and `text/html`:
// the parser is deliberately conservative — it understands paragraph breaks,
// character formatting, colour tables, hyperlink fields and Word's list
// markers, and anything it cannot parse simply yields no blocks so the
// plain-text flavour still applies.
//
// No DOM APIs: the writer and parser are string work, so they run in a Worker
// and in unit tests.

import type { ParagraphAlignment, TextStyle } from "../engine/types"
import type {
  DocxClipboardListInfo,
  DocxClipboardParagraph,
  DocxClipboardParseOptions,
  DocxClipboardRun
} from "./clipboard-html"
import {
  HEADING_FONT_SIZES_PT,
  mergeAdjacentRuns,
  sanitizeClipboardColor,
  sanitizeClipboardParagraphs,
  trimRunEdges
} from "./clipboard-html"

/** Written RTF cap; a larger document is truncated at a paragraph boundary. */
export const MAX_DOCX_CLIPBOARD_RTF_CHARS = 512 * 1024

// ── writer ─────────────────────────────────────────────────────────

interface RtfTables {
  colors: string[]
  colorIndexes: Map<string, number>
}

function rtfColorEntry(color: string): string | undefined {
  const normalized = sanitizeClipboardColor(color)
  if (!normalized) return undefined
  const hex = normalized.startsWith("#") ? normalized.slice(1) : ""
  const expanded =
    hex.length === 3
      ? hex
          .split("")
          .map((part) => part + part)
          .join("")
      : hex
  if (!/^[0-9a-f]{6}$/i.test(expanded)) return undefined
  const red = Number.parseInt(expanded.slice(0, 2), 16)
  const green = Number.parseInt(expanded.slice(2, 4), 16)
  const blue = Number.parseInt(expanded.slice(4, 6), 16)
  return `\\red${red}\\green${green}\\blue${blue}`
}

function colorTableIndex(tables: RtfTables, color?: string): number | undefined {
  if (!color) return undefined
  const entry = rtfColorEntry(color)
  if (!entry) return undefined
  const existing = tables.colorIndexes.get(entry)
  if (existing !== undefined) return existing
  tables.colors.push(entry)
  const index = tables.colors.length
  tables.colorIndexes.set(entry, index)
  return index
}

function rtfUnicode(code: number): string {
  if (code <= 0xffff) {
    const signed = code > 0x7fff ? code - 0x10000 : code
    return `\\u${signed}?`
  }
  const offset = code - 0x10000
  const high = 0xd800 + (offset >> 10)
  const low = 0xdc00 + (offset & 0x3ff)
  return `\\u${high - 0x10000}?\\u${low - 0x10000}?`
}

/** Escapes text into RTF, using `\\uN?` for everything outside ASCII. */
export function escapeDocxClipboardRtf(value: string): string {
  let output = ""
  for (const character of value) {
    const code = character.codePointAt(0) ?? 0
    if (character === "\\") output += "\\\\"
    else if (character === "{") output += "\\{"
    else if (character === "}") output += "\\}"
    else if (character === "\n") output += "\\line "
    else if (character === "\t") output += "\\tab "
    else if (code >= 0x20 && code < 0x7f) output += character
    else if (code >= 0x7f) output += rtfUnicode(code)
  }
  return output
}

function rtfLinkTarget(value: string): string {
  // An RTF field instruction is plain text; a quote or backslash would break
  // it, and the sanitizer already rejected control characters.
  return escapeDocxClipboardRtf(value).replace(/"/g, "%22")
}

function rtfRun(run: DocxClipboardRun, tables: RtfTables): string {
  const controls: string[] = []
  const style = run.style
  if (style?.bold) controls.push("\\b")
  if (style?.italic) controls.push("\\i")
  if (style?.underline) controls.push("\\ul")
  if (style?.strike) controls.push("\\strike")
  if (style?.caps) controls.push("\\caps")
  if (style?.smallCaps) controls.push("\\scaps")
  if (style?.verticalAlign === "superscript") controls.push("\\super")
  if (style?.verticalAlign === "subscript") controls.push("\\sub")
  if (typeof style?.fontSizePt === "number") {
    const halfPoints = Math.max(2, Math.min(800, Math.round(style.fontSizePt * 2)))
    controls.push(`\\fs${halfPoints}`)
  }
  const colorIndex = colorTableIndex(tables, style?.color)
  if (colorIndex !== undefined) controls.push(`\\cf${colorIndex}`)
  const highlightIndex = colorTableIndex(
    tables,
    style?.backgroundColor ?? style?.highlight
  )
  if (highlightIndex !== undefined) controls.push(`\\highlight${highlightIndex}`)

  const text = escapeDocxClipboardRtf(run.text)
  const body = controls.length > 0 ? `{${controls.join("")} ${text}}` : text
  if (!run.link) return body
  return `{\\field{\\*\\fldinst{HYPERLINK "${rtfLinkTarget(run.link)}"}}{\\fldrslt{${body}}}}`
}

function rtfAlignment(align?: ParagraphAlignment): string {
  if (align === "center") return "\\qc"
  if (align === "right") return "\\qr"
  if (align === "justify") return "\\qj"
  if (align === "left") return "\\ql"
  return ""
}

function rtfParagraph(
  paragraph: DocxClipboardParagraph,
  tables: RtfTables,
  orderedIndexes: Map<string, number>
): string {
  const controls = ["\\pard"]
  const alignment = rtfAlignment(paragraph.style?.align)
  if (alignment) controls.push(alignment)
  if (paragraph.list) controls.push("\\li720\\fi-360")

  let marker = ""
  if (paragraph.list) {
    if (paragraph.list.type === "ordered") {
      const level = Math.max(0, Math.min(8, paragraph.list.level ?? 0))
      const key = `level-${level}`
      const next = (orderedIndexes.get(key) ?? 0) + 1
      orderedIndexes.set(key, next)
      marker = `${next}.\\tab `
    } else {
      marker = "\\u8226?\\tab "
    }
  }

  let body = paragraph.runs.map((run) => rtfRun(run, tables)).join("")
  const headingLevel = paragraph.style?.headingLevel
  if (headingLevel) {
    const size = HEADING_FONT_SIZES_PT[headingLevel - 1] * 2
    body = `{\\b\\fs${size} ${body}}`
  }
  return `${controls.join("")} ${marker}${body}\\par\n`
}

/** Renders clipboard paragraphs into a compact `text/rtf` document. */
export function renderDocxClipboardRtf(
  paragraphs: readonly DocxClipboardParagraph[],
  options?: {
    maxChars?: number
    /** Drops links the destination app would otherwise follow. */
    sanitizeLink?: (value: string) => string | undefined
  }
): string {
  const maxChars = Math.max(1, options?.maxChars ?? MAX_DOCX_CLIPBOARD_RTF_CHARS)
  const tables: RtfTables = { colors: [], colorIndexes: new Map() }
  const orderedIndexes = new Map<string, number>()

  let body = ""
  for (const paragraph of paragraphs) {
    if (!paragraph.list) orderedIndexes.clear()
    const runs = options?.sanitizeLink
      ? paragraph.runs.map((run) => {
          if (!run.link) return run
          const link = options.sanitizeLink?.(run.link)
          return link ? { ...run, link } : { text: run.text, style: run.style }
        })
      : paragraph.runs
    const chunk = rtfParagraph({ ...paragraph, runs }, tables, orderedIndexes)
    if (body.length + chunk.length > maxChars) break
    body += chunk
  }

  // A paragraph that does not fit the cap leaves a header-only document with no
  // text; return no flavour at all so the caller can fall back to plain text.
  if (!body) return ""

  const colorTable =
    tables.colors.length > 0 ? `{\\colortbl;${tables.colors.join(";")};}` : "{\\colortbl;}"
  return (
    `{\\rtf1\\ansi\\ansicpg1252\\deff0\n` +
    `{\\fonttbl{\\f0\\fnil Calibri;}}\n` +
    `${colorTable}\n` +
    `${body}}\n`
  )
}

// ── parser ─────────────────────────────────────────────────────────

interface RtfParseState {
  style: TextStyle
  link?: string
  align?: ParagraphAlignment
  skip: boolean
  starDestination: boolean
  inInstruction: boolean
  inResult: boolean
  pendingHighSurrogate?: number
}

interface RtfFieldInfo {
  /** Stack length while inside the field group. */
  depth: number
  instruction: string
  link?: string
  inInstruction: boolean
  inResult: boolean
}

const RTF_SKIP_DESTINATIONS = new Set([
  "fonttbl",
  "colortbl",
  "stylesheet",
  "info",
  "pict",
  "object",
  "header",
  "headerl",
  "headerr",
  "headerf",
  "footer",
  "footerl",
  "footerr",
  "footerf",
  "footnote",
  "annotation",
  "generator",
  "listtable",
  "listoverridetable",
  "themedata",
  "colorschememapping",
  "latentstyles",
  "datastore",
  "xmlnstbl",
  "rsidtbl",
  "pgptbl",
  "pnseclvl",
  "defchp",
  "defpap",
  "filetbl",
  "revtbl",
  "xmlname",
  "mmathpr",
  "wgrffmtfilter",
  "fchars",
  "lchars"
])

/** Parses the colortbl prefix into CSS colours indexed by RTF colour number. */
function parseRtfColorTable(source: string): Array<string | undefined> {
  const start = source.indexOf("{\\colortbl")
  if (start < 0) return []
  let depth = 0
  let end = start
  for (; end < source.length; end += 1) {
    if (source[end] === "{") depth += 1
    else if (source[end] === "}") {
      depth -= 1
      if (depth === 0) break
    }
  }
  const body = source.slice(start + 1, end)
  const colors: Array<string | undefined> = [undefined]
  for (const entry of body.split(";").slice(1)) {
    const red = /\\red(\d+)/.exec(entry)?.[1]
    const green = /\\green(\d+)/.exec(entry)?.[1]
    const blue = /\\blue(\d+)/.exec(entry)?.[1]
    if (red === undefined || green === undefined || blue === undefined) continue
    const channel = (value: string): string =>
      Math.max(0, Math.min(255, Number.parseInt(value, 10))).toString(16).padStart(2, "0")
    colors.push(`#${channel(red)}${channel(green)}${channel(blue)}`)
  }
  return colors
}

function decodeRtfHexByte(value: string): string {
  const code = Number.parseInt(value, 16)
  if (!Number.isFinite(code)) return ""
  // RTF `\'hh` is one byte of the document code page; our writer never emits
  // it, and Word's ASCII text is unaffected by the latin-1 approximation.
  return String.fromCharCode(code)
}

function parseRtfFontSizePt(parameter: number): number | undefined {
  if (!Number.isFinite(parameter) || parameter <= 0) return undefined
  return Math.max(1, Math.min(400, parameter / 2))
}

function parseRtfLinkInstruction(instruction: string): string | undefined {
  const match = /HYPERLINK\s+(?:"([^"]*)"|(\S+))/i.exec(instruction)
  return match?.[1] ?? match?.[2]
}

/**
 * Parses the `text/rtf` clipboard flavour into clipboard paragraphs.
 *
 * Best effort by design: paragraph breaks, character formatting, colours,
 * hyperlink fields and Word list markers are understood; every other control
 * word is ignored. A document the parser cannot read yields no blocks, so paste
 * falls back to `text/plain`.
 */
export function parseDocxClipboardRtf(
  rtf: string,
  options?: DocxClipboardParseOptions & { maxRtfChars?: number }
): DocxClipboardParagraph[] {
  if (!rtf) return []
  const maxChars = Math.max(1, options?.maxRtfChars ?? MAX_DOCX_CLIPBOARD_RTF_CHARS)
  const source = rtf.length > maxChars ? rtf.slice(0, maxChars) : rtf
  const colors = parseRtfColorTable(source)

  const initialState: RtfParseState = {
    style: {},
    skip: false,
    starDestination: false,
    inInstruction: false,
    inResult: false
  }
  const stack: RtfParseState[] = [{ ...initialState, style: {} }]
  const fields: RtfFieldInfo[] = []
  const paragraphs: DocxClipboardParagraph[] = []
  let runs: DocxClipboardRun[] = []
  let blockAlign: ParagraphAlignment | undefined
  let opened = false

  const current = (): RtfParseState => stack[stack.length - 1] ?? initialState
  const activeField = (): RtfFieldInfo | undefined => fields[fields.length - 1]

  function appendCodePoint(code: number, state: RtfParseState): void {
    if (code >= 0xd800 && code <= 0xdbff) {
      state.pendingHighSurrogate = code
      return
    }
    if (code >= 0xdc00 && code <= 0xdfff) {
      const high = state.pendingHighSurrogate
      state.pendingHighSurrogate = undefined
      appendChars(high !== undefined ? String.fromCharCode(high, code) : String.fromCharCode(code), state)
      return
    }
    state.pendingHighSurrogate = undefined
    appendChars(String.fromCharCode(code), state)
  }

  function appendChars(text: string, state: RtfParseState): void {
    if (!text) return
    const field = activeField()
    if (field?.inInstruction) {
      field.instruction += text
      return
    }
    if (state.skip) return
    if (!opened) opened = true
    const style = Object.keys(state.style).length > 0 ? { ...state.style } : undefined
    const link = state.link
    const previous = runs[runs.length - 1]
    if (previous && previous.link === link && JSON.stringify(previous.style ?? null) === JSON.stringify(style ?? null)) {
      previous.text += text
      return
    }
    const run: DocxClipboardRun = { text }
    if (style) run.style = style
    if (link) run.link = link
    runs.push(run)
  }

  function flushBlock(): void {
    if (!opened && runs.length === 0) return
    const trimmed = trimRunEdges(mergeAdjacentRuns(runs))
    if (trimmed.length === 0 && !opened) {
      runs = []
      return
    }
    const list = stripRtfListMarker(trimmed)
    const paragraph: DocxClipboardParagraph = { runs: list.runs }
    if (list.marker) paragraph.list = list.marker
    if (blockAlign) paragraph.style = { align: blockAlign }
    paragraphs.push(paragraph)
    runs = []
    opened = false
    blockAlign = undefined
  }

  let index = 0
  while (index < source.length) {
    const character = source[index]

    if (character === "{") {
      stack.push({ ...current(), style: { ...current().style }, starDestination: false })
      index += 1
      continue
    }

    if (character === "}") {
      const closing = stack[stack.length - 1]
      if (closing?.inInstruction) {
        const field = activeField()
        if (field) {
          const link = parseRtfLinkInstruction(field.instruction)
          field.link = link && options?.sanitizeLink ? options.sanitizeLink(link) : undefined
        }
      }
      if (stack.length > 1) stack.pop()
      while (fields.length > 0 && stack.length < fields[fields.length - 1].depth) {
        fields.pop()
      }
      index += 1
      continue
    }

    if (character !== "\\") {
      // Raw CR/LF are insignificant in RTF (formatting whitespace), unlike the
      // `\\line`/`\\par` control words.
      if (character !== "\n" && character !== "\r") {
        appendCodePoint(character.codePointAt(0) ?? 0, current())
      }
      index += 1
      continue
    }

    // Control symbol or control word.
    const next = source[index + 1]
    if (next === undefined) break
    if (!/[a-zA-Z]/.test(next)) {
      index += 2
      const state = current()
      switch (next) {
        case "\\":
        case "{":
        case "}":
          appendChars(next, state)
          break
        case "~":
          appendChars("\u00a0", state)
          break
        case "-":
        case "_":
          break
        case "*":
          state.starDestination = true
          break
        case "'": {
          const hex = source.slice(index, index + 2)
          appendChars(decodeRtfHexByte(hex), state)
          index += 2
          break
        }
        case "\n":
        case "\r":
          break
        default:
          break
      }
      continue
    }

    // Control word: letters, optional signed parameter and one delimiter.
    let wordEnd = index + 1
    while (wordEnd < source.length && /[a-zA-Z]/.test(source[wordEnd])) wordEnd += 1
    const word = source.slice(index + 1, wordEnd).toLowerCase()
    let parameterEnd = wordEnd
    if (source[parameterEnd] === "-" || /\d/.test(source[parameterEnd] ?? "")) {
      parameterEnd += 1
      while (parameterEnd < source.length && /\d/.test(source[parameterEnd])) parameterEnd += 1
    }
    const parameterText = source.slice(wordEnd, parameterEnd)
    const parameter = parameterText ? Number.parseInt(parameterText, 10) : undefined
    index = source[parameterEnd] === " " ? parameterEnd + 1 : parameterEnd

    const state = current()

    if (state.starDestination) {
      state.starDestination = false
      if (word === "fldinst") {
        const field = activeField()
        if (field) {
          field.inInstruction = true
          state.inInstruction = true
        }
        continue
      }
      state.skip = true
      continue
    }

    switch (word) {
      case "par":
      case "pard":
        if (word === "pard") {
          state.align = undefined
          blockAlign = undefined
          continue
        }
        flushBlock()
        continue
      case "line":
        appendChars("\n", state)
        continue
      case "tab":
        appendChars("\t", state)
        continue
      case "field": {
        fields.push({
          depth: stack.length,
          instruction: "",
          inInstruction: false,
          inResult: false
        })
        continue
      }
      case "fldinst": {
        const field = activeField()
        if (field) {
          field.inInstruction = true
          state.inInstruction = true
        }
        continue
      }
      case "fldrslt": {
        const field = activeField()
        if (field) {
          field.inInstruction = false
          field.inResult = true
          state.inInstruction = false
          state.inResult = true
          if (field.link) state.link = field.link
        }
        continue
      }
      case "b":
        state.style.bold = parameter !== 0
        continue
      case "i":
        state.style.italic = parameter !== 0
        continue
      case "ul":
        state.style.underline = parameter !== 0
        continue
      case "ulnone":
        state.style.underline = false
        continue
      case "strike":
      case "striked":
        state.style.strike = parameter !== 0
        continue
      case "caps":
        state.style.caps = parameter !== 0
        continue
      case "scaps":
        state.style.smallCaps = parameter !== 0
        continue
      case "super":
        state.style.verticalAlign = "superscript"
        continue
      case "sub":
        state.style.verticalAlign = "subscript"
        continue
      case "nosupersub":
        state.style.verticalAlign = undefined
        continue
      case "plain":
        state.style = {}
        state.link = undefined
        continue
      case "fs": {
        const fontSizePt = parameter !== undefined ? parseRtfFontSizePt(parameter) : undefined
        state.style.fontSizePt = fontSizePt
        continue
      }
      case "cf":
      case "highlight": {
        const color = parameter !== undefined ? colors[parameter] : undefined
        const sanitized = sanitizeClipboardColor(color)
        if (word === "cf") state.style.color = sanitized
        else state.style.highlight = sanitized
        continue
      }
      case "qc":
        state.align = "center"
        blockAlign = "center"
        continue
      case "qr":
        state.align = "right"
        blockAlign = "right"
        continue
      case "qj":
        state.align = "justify"
        blockAlign = "justify"
        continue
      case "ql":
        state.align = "left"
        blockAlign = "left"
        continue
      case "u": {
        if (parameter === undefined) continue
        const code = parameter < 0 ? parameter + 65536 : parameter
        appendCodePoint(code, state)
        // The RTF spec allows one fallback character after `\uN`.
        if (source[index] !== undefined && source[index] !== "\\") index += 1
        continue
      }
      default:
        if (RTF_SKIP_DESTINATIONS.has(word)) {
          state.skip = true
          continue
        }
        // `\pntext` group content is the list marker itself ("1.\tab"),
        // captured as text so the paragraph-level marker detection sees it.
        continue
    }
  }

  flushBlock()
  return sanitizeClipboardParagraphs(paragraphs, options)
}

/** Detects and strips Word's `marker\ttab` list prefix from a paragraph. */
function stripRtfListMarker(runs: DocxClipboardRun[]): {
  marker?: DocxClipboardListInfo
  runs: DocxClipboardRun[]
} {
  const text = runs.map((run) => run.text).join("")
  const match = /^\s*([\u00b7\u2022\u25cf\u25cb\u25a0\u25a1\u25aa\u25ab\u25c6\u25c7\u25e6\u2043-]|\d+[.)])\t\s*/.exec(text)
  if (!match) return { runs }
  const markerWidth = match[0].length
  let remaining = markerWidth
  const trimmed: DocxClipboardRun[] = []
  for (const run of runs) {
    if (remaining >= run.text.length) {
      remaining -= run.text.length
      continue
    }
    const text = remaining > 0 ? run.text.slice(remaining) : run.text
    remaining = 0
    if (text) trimmed.push({ ...run, text })
  }
  const markerText = match[1] ?? ""
  const ordered = /\d/.test(markerText)
  return {
    marker: { type: ordered ? "ordered" : "unordered", level: 0 },
    runs: trimmed
  }
}
