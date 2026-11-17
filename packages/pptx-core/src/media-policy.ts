import { resolveOfficeUrl, type OfficeUrlPolicy } from "@arcships/office-runtime"

import type { PptxExternalMediaUrlPolicy } from "./browser-types"

/**
 * Authorize one external media URL against the host's policy.
 *
 * A deck can reference tracking pixels or `javascript:` targets, so external
 * media is only loaded when the host opted in *and* the shared URL policy allows
 * the resolved URL. Anything else stays blocked and the media element simply has
 * no source.
 */
export function resolvePptxExternalMediaUrl(
  url: string,
  policy: PptxExternalMediaUrlPolicy | undefined,
): string | undefined {
  try {
    return resolveOfficeUrl(url, policy as OfficeUrlPolicy | undefined).href
  } catch {
    return undefined
  }
}
