export const OFFICE_LINK_PROTOCOLS = ["http:", "https:", "mailto:"] as const;

export type OfficeLinkProtocol = (typeof OFFICE_LINK_PROTOCOLS)[number];

const SAFE_PROTOCOLS = new Set<string>(OFFICE_LINK_PROTOCOLS);
const SCHEME_PATTERN = /^([A-Za-z][A-Za-z\d+.-]*):/;
const CONTROL_CHARACTER_PATTERN = /[\u0000-\u001F\u007F]/;
const RELATIVE_BASE = "https://office-link.invalid/";

/**
 * Returns the original target when it is safe to place on a link anchor.
 *
 * Absolute http(s) and mailto targets, fragment-only targets and relative
 * references are accepted. Other schemes, including javascript:, data: and
 * vbscript:, are rejected so callers can render plain text or refuse input.
 */
export function sanitizeOfficeLink(value: unknown): string | undefined {
  if (typeof value !== "string" || value.length === 0) return undefined;
  if (value.trim() !== value) return undefined;
  if (CONTROL_CHARACTER_PATTERN.test(value)) return undefined;

  if (value.startsWith("#")) return value;

  const scheme = SCHEME_PATTERN.exec(value);
  if (scheme) {
    const protocol = `${scheme[1].toLowerCase()}:`;
    if (!SAFE_PROTOCOLS.has(protocol)) return undefined;
    try {
      const parsed = new URL(value);
      if (parsed.protocol !== protocol) return undefined;
      if ((protocol === "http:" || protocol === "https:") && !parsed.hostname) {
        return undefined;
      }
      if (protocol === "mailto:" && parsed.pathname.length === 0) return undefined;
    } catch {
      return undefined;
    }
    return value;
  }

  // Relative references are resolved once to prove they stay on https.
  try {
    const parsed = new URL(value, RELATIVE_BASE);
    if (parsed.protocol !== "https:") return undefined;
  } catch {
    return undefined;
  }
  return value;
}

export function isSafeOfficeLink(value: unknown): value is string {
  return typeof value === "string" && sanitizeOfficeLink(value) !== undefined;
}
