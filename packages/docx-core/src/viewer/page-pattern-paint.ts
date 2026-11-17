import type { DocumentBackgroundPattern } from "../engine/types";

export interface PagePatternPaintStyle {
  backgroundColor?: string;
  backgroundImage?: string;
  backgroundSize?: string;
  backgroundRepeat?: "repeat";
}

export interface PagePatternCanvasArea {
  xPx?: number;
  yPx?: number;
  widthPx: number;
  heightPx: number;
}

function normalizedPreset(preset: string): string {
  return preset.trim().replace(/[\s_-]+/g, "").toLowerCase();
}

function densitySpacing(preset: string): number | undefined {
  const match = preset.match(/^pct(\d{1,3})$/);
  if (!match) return undefined;
  const percent = Math.max(1, Math.min(100, Number(match[1]) || 5));
  return Math.max(4, Math.round(100 / Math.sqrt(percent)));
}

function diagonalDirection(preset: string): "down" | "up" | undefined {
  if (/^(?:dndiag|dashdndiag|wddndiag)$/i.test(preset)) return "down";
  if (/^(?:updiag|dashupdiag|wdupdiag)$/i.test(preset)) return "up";
  return undefined;
}

/**
 * Builds a CSS background for a page pattern. Repeated gradients are used for
 * both line and dot presets so the pattern fills the whole paper without an
 * extra image element.
 */
export function pagePatternPaintStyle(
  pattern?: DocumentBackgroundPattern
): PagePatternPaintStyle {
  if (!pattern?.preset) return {};

  const preset = normalizedPreset(pattern.preset);
  const foreground = pattern.foregroundColor ?? "#000000";
  const base: PagePatternPaintStyle = {
    backgroundColor: pattern.backgroundColor,
    backgroundRepeat: "repeat",
  };
  const spacing = densitySpacing(preset);
  if (spacing !== undefined) {
    return {
      ...base,
      backgroundImage: `radial-gradient(circle, ${foreground} 1px, transparent 1.4px)`,
      backgroundSize: `${spacing}px ${spacing}px`,
    };
  }

  const line = (angle: number, widthPx = 1, periodPx = 8): string =>
    `repeating-linear-gradient(${angle}deg, ${foreground} 0, ${foreground} ${widthPx}px, transparent ${widthPx}px, transparent ${periodPx}px)`;

  if (/^(?:horz|lthorz|dkhorz|narhorz|dashhorz)$/i.test(preset)) {
    return { ...base, backgroundImage: line(0, preset.startsWith("dk") ? 2 : 1, 8) };
  }
  if (/^(?:vert|ltvert|dkvert|narvert|dashvert)$/i.test(preset)) {
    return { ...base, backgroundImage: line(90, preset.startsWith("dk") ? 2 : 1, 8) };
  }
  const diagonal = diagonalDirection(preset);
  if (diagonal === "down") {
    return { ...base, backgroundImage: line(135) };
  }
  if (diagonal === "up") {
    return { ...base, backgroundImage: line(45) };
  }
  if (preset === "cross") {
    return {
      ...base,
      backgroundImage: `${line(0)}, ${line(90)}`,
    };
  }
  if (/^(?:diagcross|smgrid|lggrid)$/i.test(preset)) {
    return {
      ...base,
      backgroundImage: `${line(45)}, ${line(135)}`,
    };
  }

  return {
    ...base,
    backgroundImage: `radial-gradient(circle, ${foreground} 1px, transparent 1.4px)`,
    backgroundSize: "8px 8px",
  };
}

function clampArea(area: PagePatternCanvasArea): {
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
} {
  return {
    xPx: Math.max(0, Math.round(area.xPx ?? 0)),
    yPx: Math.max(0, Math.round(area.yPx ?? 0)),
    widthPx: Math.max(0, Math.round(area.widthPx)),
    heightPx: Math.max(0, Math.round(area.heightPx)),
  };
}

/**
 * Paints the same page pattern onto a canvas thumbnail. The DOM page uses the
 * CSS gradients from `pagePatternPaintStyle`; this canvas counterpart keeps
 * thumbnail rails visually aligned with the paper.
 */
export function paintPagePatternOnCanvas(
  context: CanvasRenderingContext2D,
  pattern: DocumentBackgroundPattern | undefined,
  area: PagePatternCanvasArea
): void {
  if (!pattern?.preset) return;
  const { xPx, yPx, widthPx, heightPx } = clampArea(area);
  if (widthPx <= 0 || heightPx <= 0) return;

  const preset = normalizedPreset(pattern.preset);
  const foreground = pattern.foregroundColor ?? "#000000";
  const right = xPx + widthPx;
  const bottom = yPx + heightPx;
  const spacing = densitySpacing(preset);
  const diagonal = diagonalDirection(preset);

  context.save();
  context.fillStyle = foreground;
  context.strokeStyle = foreground;
  context.lineWidth = 1;

  const drawDots = (step: number): void => {
    for (let y = yPx + step / 2; y < bottom; y += step) {
      for (let x = xPx + step / 2; x < right; x += step) {
        context.fillRect(Math.round(x), Math.round(y), 1, 1);
      }
    }
  };

  const drawLines = (kind: "horizontal" | "vertical" | "down" | "up"): void => {
    const step = 8;
    context.beginPath();
    if (kind === "horizontal") {
      for (let y = yPx; y <= bottom; y += step) {
        context.moveTo(xPx, y);
        context.lineTo(right, y);
      }
    } else if (kind === "vertical") {
      for (let x = xPx; x <= right; x += step) {
        context.moveTo(x, yPx);
        context.lineTo(x, bottom);
      }
    } else if (kind === "down") {
      for (let x = xPx - heightPx; x <= right; x += step) {
        context.moveTo(x, yPx);
        context.lineTo(x + heightPx, bottom);
      }
    } else {
      for (let x = xPx - heightPx; x <= right; x += step) {
        context.moveTo(x, bottom);
        context.lineTo(x + heightPx, yPx);
      }
    }
    context.stroke();
  };

  if (spacing !== undefined) {
    drawDots(spacing);
  } else if (/^(?:horz|lthorz|dkhorz|narhorz|dashhorz)$/i.test(preset)) {
    drawLines("horizontal");
  } else if (/^(?:vert|ltvert|dkvert|narvert|dashvert)$/i.test(preset)) {
    drawLines("vertical");
  } else if (diagonal) {
    drawLines(diagonal);
  } else if (preset === "cross") {
    drawLines("horizontal");
    drawLines("vertical");
  } else if (/^(?:diagcross|smgrid|lggrid)$/i.test(preset)) {
    drawLines("down");
    drawLines("up");
  } else {
    drawDots(8);
  }

  context.restore();
}
