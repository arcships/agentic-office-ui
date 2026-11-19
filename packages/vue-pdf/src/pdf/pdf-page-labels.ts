/**
 * Read the `/PageLabels` number tree from PDF bytes and turn it into one
 * display label per page.
 *
 * PDFium exposes page labels through `FPDF_GetPageLabel`, but the EmbedPDF
 * worker engine does not forward that call. The label tree lives in the
 * document catalog and is cheap to read directly, so this module walks the
 * bytes for `/PageLabels` instead of adding an engine round-trip. Only the
 * label number tree is parsed: streams and object streams are ignored, and a
 * document whose labels are compressed falls back to decimal page numbers.
 */

interface PdfDictValue {
  kind: "dict"
  entries: Map<string, PdfValue>
}

type PdfValue =
  | PdfDictValue
  | { kind: "array"; items: PdfValue[] }
  | { kind: "name"; name: string }
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "reference"; number: number; generation: number }
  | { kind: "keyword"; keyword: string }

interface PdfPageLabelSpec {
  /** Zero-based page index where the range starts. */
  start: number
  prefix: string
  style: "D" | "R" | "r" | "A" | "a"
  startNumber: number
}

const PDF_WHITESPACE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20])
const PDF_DELIMITERS = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25])
const MAX_PDF_VALUE_DEPTH = 32

function asciiBytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length)
  for (let index = 0; index < text.length; index++) bytes[index] = text.charCodeAt(index) & 0xff
  return bytes
}

function indexOfSequence(haystack: Uint8Array, needle: Uint8Array, from = 0): number {
  if (needle.length === 0) return Math.min(Math.max(0, from), haystack.length)
  const first = needle[0]
  const limit = haystack.length - needle.length
  for (let index = Math.max(0, from); index <= limit; index++) {
    if (haystack[index] !== first) continue
    let offset = 1
    while (offset < needle.length && haystack[index + offset] === needle[offset]) offset++
    if (offset === needle.length) return index
  }
  return -1
}

function decodeStringBytes(bytes: Uint8Array): string {
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    let text = ""
    for (let index = 2; index + 1 < bytes.length; index += 2) {
      text += String.fromCharCode((bytes[index] << 8) | bytes[index + 1])
    }
    return text
  }
  let text = ""
  for (const byte of bytes) text += String.fromCharCode(byte)
  return text
}

function readHexDigit(byte: number): number {
  if (byte >= 0x30 && byte <= 0x39) return byte - 0x30
  if (byte >= 0x41 && byte <= 0x46) return byte - 0x37
  if (byte >= 0x61 && byte <= 0x66) return byte - 0x57
  return -1
}

class PdfValueReader {
  private position: number

  constructor(private readonly bytes: Uint8Array, position = 0) {
    this.position = position
  }

  readValue(depth = 0): PdfValue | undefined {
    if (depth > MAX_PDF_VALUE_DEPTH) return undefined
    this.skipSpace()
    if (this.position >= this.bytes.length) return undefined
    const byte = this.bytes[this.position]
    if (byte === 0x3c) {
      if (this.bytes[this.position + 1] === 0x3c) return this.readDictionary(depth)
      return this.readHexString()
    }
    if (byte === 0x5b) return this.readArray(depth)
    if (byte === 0x2f) return this.readName()
    if (byte === 0x28) return this.readLiteralString()
    return this.readTokenValue(depth)
  }

  private skipSpace(): void {
    while (this.position < this.bytes.length) {
      const byte = this.bytes[this.position]
      if (PDF_WHITESPACE.has(byte)) {
        this.position++
        continue
      }
      if (byte === 0x25) {
        while (
          this.position < this.bytes.length
          && this.bytes[this.position] !== 0x0a
          && this.bytes[this.position] !== 0x0d
        ) this.position++
        continue
      }
      break
    }
  }

  private readDictionary(depth: number): PdfDictValue {
    const entries = new Map<string, PdfValue>()
    this.position += 2
    for (;;) {
      this.skipSpace()
      if (this.position >= this.bytes.length) break
      if (this.bytes[this.position] === 0x3e && this.bytes[this.position + 1] === 0x3e) {
        this.position += 2
        break
      }
      if (this.bytes[this.position] !== 0x2f) {
        // A malformed entry cannot be skipped reliably; stop instead of
        // emitting a partial dictionary that looks authoritative.
        break
      }
      const key = this.readName().name
      const value = this.readValue(depth + 1)
      if (!value) break
      entries.set(key, value)
    }
    return { kind: "dict", entries }
  }

  private readArray(depth: number): PdfValue {
    const items: PdfValue[] = []
    this.position++
    for (;;) {
      this.skipSpace()
      if (this.position >= this.bytes.length) break
      if (this.bytes[this.position] === 0x5d) {
        this.position++
        break
      }
      const value = this.readValue(depth + 1)
      if (!value) break
      items.push(value)
    }
    return { kind: "array", items }
  }

  private readName(): { kind: "name"; name: string } {
    this.position++
    let name = ""
    while (this.position < this.bytes.length) {
      const byte = this.bytes[this.position]
      if (PDF_WHITESPACE.has(byte) || PDF_DELIMITERS.has(byte)) break
      if (byte === 0x23 && this.position + 2 < this.bytes.length) {
        const high = readHexDigit(this.bytes[this.position + 1])
        const low = readHexDigit(this.bytes[this.position + 2])
        if (high >= 0 && low >= 0) {
          name += String.fromCharCode((high << 4) | low)
          this.position += 3
          continue
        }
      }
      name += String.fromCharCode(byte)
      this.position++
    }
    return { kind: "name", name }
  }

  private readHexString(): PdfValue {
    this.position++
    const bytes: number[] = []
    let high = -1
    while (this.position < this.bytes.length) {
      const byte = this.bytes[this.position]
      this.position++
      if (byte === 0x3e) break
      const digit = readHexDigit(byte)
      if (digit < 0) continue
      if (high < 0) high = digit
      else {
        bytes.push((high << 4) | digit)
        high = -1
      }
    }
    if (high >= 0) bytes.push(high << 4)
    return { kind: "string", value: decodeStringBytes(Uint8Array.from(bytes)) }
  }

  private readLiteralString(): PdfValue {
    this.position++
    const bytes: number[] = []
    let depth = 1
    while (this.position < this.bytes.length) {
      const byte = this.bytes[this.position]
      this.position++
      if (byte === 0x5c) {
        const escaped = this.readEscape()
        if (escaped !== undefined) bytes.push(escaped)
        continue
      }
      if (byte === 0x28) {
        depth++
        bytes.push(byte)
        continue
      }
      if (byte === 0x29) {
        depth--
        if (depth === 0) break
        bytes.push(byte)
        continue
      }
      bytes.push(byte)
    }
    return { kind: "string", value: decodeStringBytes(Uint8Array.from(bytes)) }
  }

  private readEscape(): number | undefined {
    if (this.position >= this.bytes.length) return undefined
    const byte = this.bytes[this.position]
    this.position++
    switch (byte) {
      case 0x6e: return 0x0a
      case 0x72: return 0x0d
      case 0x74: return 0x09
      case 0x62: return 0x08
      case 0x66: return 0x0c
      case 0x0a: return undefined
      case 0x0d:
        if (this.bytes[this.position] === 0x0a) this.position++
        return undefined
      default:
        if (byte >= 0x30 && byte <= 0x37) {
          let value = byte - 0x30
          for (let count = 1; count < 3 && this.position < this.bytes.length; count++) {
            const next = this.bytes[this.position]
            if (next < 0x30 || next > 0x37) break
            value = value * 8 + (next - 0x30)
            this.position++
          }
          return value & 0xff
        }
        return byte
    }
  }

  private readRawToken(): string | undefined {
    this.skipSpace()
    const start = this.position
    while (this.position < this.bytes.length) {
      const byte = this.bytes[this.position]
      if (PDF_WHITESPACE.has(byte) || PDF_DELIMITERS.has(byte)) break
      this.position++
    }
    if (this.position === start) return undefined
    let token = ""
    for (let index = start; index < this.position; index++) token += String.fromCharCode(this.bytes[index])
    return token
  }

  private readTokenValue(depth: number): PdfValue | undefined {
    const start = this.position
    const token = this.readRawToken()
    if (token === undefined) {
      this.position++
      return undefined
    }
    if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/u.test(token)) {
      const value = Number(token)
      if (Number.isInteger(value)) {
        const reset = this.position
        const generation = this.readRawToken()
        const reference = this.readRawToken()
        if (
          generation !== undefined
          && reference === "R"
          && /^\d+$/u.test(generation)
        ) {
          return { kind: "reference", number: value, generation: Number(generation) }
        }
        this.position = reset
      }
      return { kind: "number", value }
    }
    if (token === "true" || token === "false" || token === "null") {
      return { kind: "keyword", keyword: token }
    }
    // Unknown keywords (obj, stream, endobj, ...) are values we cannot use; the
    // caller treats them as missing instead of misreading a following token.
    this.position = start + token.length
    return { kind: "keyword", keyword: token }
  }
}

function resolvePdfValue(
  value: PdfValue | undefined,
  bytes: Uint8Array,
  cache: Map<string, PdfValue | undefined>,
  depth = 0,
): PdfValue | undefined {
  if (!value || value.kind !== "reference" || depth > MAX_PDF_VALUE_DEPTH) return value
  const key = `${value.number} ${value.generation}`
  if (cache.has(key)) return cache.get(key)
  cache.set(key, undefined)
  const resolved = findPdfObjectValue(bytes, value.number, value.generation)
  cache.set(key, resolved)
  return resolved && resolved.kind === "reference"
    ? resolvePdfValue(resolved, bytes, cache, depth + 1)
    : resolved
}

function findPdfObjectValue(bytes: Uint8Array, number: number, generation: number): PdfValue | undefined {
  const pattern = asciiBytes(`${number} ${generation} obj`)
  let from = 0
  for (;;) {
    const at = indexOfSequence(bytes, pattern, from)
    if (at < 0) return undefined
    const before = at > 0 ? bytes[at - 1] : 0x20
    if (PDF_WHITESPACE.has(before) || before === 0x5d || before === 0x3e) {
      const reader = new PdfValueReader(bytes, at + pattern.length)
      const value = reader.readValue()
      if (value) return value
    }
    from = at + 1
  }
}

function readPageLabelDictionary(
  value: PdfValue | undefined,
  bytes: Uint8Array,
  cache: Map<string, PdfValue | undefined>,
): PdfPageLabelSpec | undefined {
  const resolved = resolvePdfValue(value, bytes, cache)
  if (!resolved || resolved.kind !== "dict") return undefined
  const styleValue = resolvePdfValue(resolved.entries.get("S"), bytes, cache)
  const prefixValue = resolvePdfValue(resolved.entries.get("P"), bytes, cache)
  const startValue = resolvePdfValue(resolved.entries.get("St"), bytes, cache)
  const rawStyle = styleValue?.kind === "name" ? styleValue.name : "D"
  const style = rawStyle === "R" || rawStyle === "r" || rawStyle === "A" || rawStyle === "a"
    ? rawStyle
    : "D"
  const prefix = prefixValue?.kind === "string" ? prefixValue.value : ""
  const startNumber = startValue?.kind === "number" && Number.isFinite(startValue.value)
    ? Math.max(1, Math.floor(startValue.value))
    : 1
  return { start: 0, prefix, style, startNumber }
}

function collectPageLabelSpecs(
  value: PdfValue | undefined,
  bytes: Uint8Array,
  cache: Map<string, PdfValue | undefined>,
  result: PdfPageLabelSpec[],
  depth = 0,
): void {
  if (depth > MAX_PDF_VALUE_DEPTH) return
  const resolved = resolvePdfValue(value, bytes, cache)
  if (!resolved || resolved.kind !== "dict") return

  const nums = resolvePdfValue(resolved.entries.get("Nums"), bytes, cache)
  if (nums?.kind === "array") {
    for (let index = 0; index + 1 < nums.items.length; index += 2) {
      const startValue = nums.items[index]
      if (startValue.kind !== "number" || !Number.isFinite(startValue.value)) continue
      const spec = readPageLabelDictionary(nums.items[index + 1], bytes, cache)
      if (!spec) continue
      const start = Math.max(0, Math.floor(startValue.value))
      result.push({ ...spec, start })
    }
  }

  const kids = resolvePdfValue(resolved.entries.get("Kids"), bytes, cache)
  if (kids?.kind === "array") {
    for (const kid of kids.items) collectPageLabelSpecs(kid, bytes, cache, result, depth + 1)
  }
}

function readPageLabelSpecs(bytes: Uint8Array): PdfPageLabelSpec[] {
  const needle = asciiBytes("/PageLabels")
  const cache = new Map<string, PdfValue | undefined>()
  const specs: PdfPageLabelSpec[] = []
  let from = 0
  for (;;) {
    const at = indexOfSequence(bytes, needle, from)
    if (at < 0) break
    from = at + needle.length
    const reader = new PdfValueReader(bytes, from)
    const value = reader.readValue()
    if (!value) continue
    collectPageLabelSpecs(value, bytes, cache, specs)
    // A well-formed document defines the tree once; the first readable tree
    // wins so a duplicated/stream-embedded token cannot mix ranges.
    if (specs.length > 0) break
  }
  return specs
}

/** `1` → `A`, `26` → `Z`, `27` → `AA`, the PDF letter-label sequence. */
function alphabeticPageLabel(value: number): string {
  let remaining = Math.max(1, Math.floor(value))
  let result = ""
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26
    result = String.fromCharCode(0x41 + remainder) + result
    remaining = Math.floor((remaining - 1) / 26)
  }
  return result
}

const ROMAN_VALUES: ReadonlyArray<readonly [number, string]> = [
  [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"],
  [100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
  [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
]

function romanPageLabel(value: number): string {
  let remaining = Math.max(1, Math.floor(value))
  let result = ""
  for (const [amount, numeral] of ROMAN_VALUES) {
    while (remaining >= amount) {
      result += numeral
      remaining -= amount
    }
  }
  return result
}

function formatPageLabelNumber(style: PdfPageLabelSpec["style"], value: number): string {
  if (!Number.isFinite(value)) return ""
  const number = Math.floor(value)
  switch (style) {
    case "R": return romanPageLabel(number)
    case "r": return romanPageLabel(number).toLowerCase()
    case "A": return alphabeticPageLabel(number)
    case "a": return alphabeticPageLabel(number).toLowerCase()
    default: return String(number)
  }
}

function defaultPageLabels(pageCount: number): string[] {
  const count = Number.isFinite(pageCount) ? Math.max(0, Math.floor(pageCount)) : 0
  return Array.from({ length: count }, (_, index) => String(index + 1))
}

/**
 * Page labels for a document. Always returns exactly `pageCount` entries: a
 * missing or unreadable `/PageLabels` tree yields decimal page numbers, and
 * pages before the first range keep their decimal number.
 */
export function readPdfPageLabels(bytes: Uint8Array, pageCount: number): string[] {
  const labels = defaultPageLabels(pageCount)
  if (labels.length === 0) return labels
  let specs: PdfPageLabelSpec[]
  try {
    specs = readPageLabelSpecs(bytes)
  } catch {
    return labels
  }
  if (specs.length === 0) return labels

  const byStart = new Map<number, PdfPageLabelSpec>()
  for (const spec of specs) {
    if (!byStart.has(spec.start)) byStart.set(spec.start, spec)
  }
  const sorted = [...byStart.values()].sort((left, right) => left.start - right.start)
  for (let index = 0; index < sorted.length; index++) {
    const spec = sorted[index]
    const start = Math.min(labels.length, spec.start)
    const nextStart = index + 1 < sorted.length ? sorted[index + 1].start : labels.length
    const end = Math.min(labels.length, Math.max(start, nextStart))
    for (let page = start; page < end; page++) {
      labels[page] = spec.prefix + formatPageLabelNumber(spec.style, spec.startNumber + (page - start))
    }
  }
  return labels
}
