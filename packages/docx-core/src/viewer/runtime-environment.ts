/**
 * A DOM-like environment has a UI thread that must never run an unbounded
 * synchronous WASM parse or export. Node and other headless hosts do not.
 */
export function isUiEnvironment(): boolean {
  return typeof document !== "undefined" || typeof window !== "undefined";
}
