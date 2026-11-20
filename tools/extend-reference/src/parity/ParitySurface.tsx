import type { ReactNode } from "react";

export type ParityTheme = "light" | "dark";

/**
 * Harness contract shared with tools/vue-reference and scripts/ci/pixel-parity.py:
 *
 *   [data-parity-surface]  surface id (query param `surface`)
 *   [data-parity-theme]    "light" | "dark" (query param `theme`, `.dark` class applied here)
 *   [data-parity-cell]     state id; one cell per manifest state
 *   [data-parity-host]     wrapper whose first element child is the measured target
 *
 * The Playwright helper screenshots the cell (so focus rings and shadow bleed
 * stay inside the capture) and measures the target's box and computed styles.
 */
export function ParitySurface({
  surface,
  theme,
  children,
}: {
  surface: string;
  theme: ParityTheme;
  children: ReactNode;
}) {
  return (
    <div
      className={theme === "dark" ? "dark" : undefined}
      data-parity-app="extend-reference"
      data-parity-surface={surface}
      data-parity-theme={theme}
      style={{
        minHeight: "100vh",
        padding: 24,
        display: "flex",
        flexDirection: "column",
        alignItems: "flex-start",
        gap: 16,
        boxSizing: "border-box",
        colorScheme: theme === "dark" ? "dark" : "light",
        background: "var(--background)",
        color: "var(--foreground)",
      }}
    >
      {children}
    </div>
  );
}

export function ParityCell({
  state,
  matrixState,
  width,
  children,
}: {
  state: string;
  matrixState: string;
  width: number;
  children: ReactNode;
}) {
  return (
    <div
      data-parity-cell={state}
      data-parity-matrix-state={matrixState}
      style={{
        width,
        maxWidth: "100%",
        padding: 12,
        boxSizing: "border-box",
      }}
    >
      <div data-parity-host>{children}</div>
    </div>
  );
}
