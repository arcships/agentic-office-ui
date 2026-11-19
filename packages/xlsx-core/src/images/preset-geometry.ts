/**
 * DrawingML preset geometry (`a:prstGeom@prst`).
 *
 * The path table follows the ECMA-376 preset shape definitions (the same
 * corpus Excel and PowerPoint use) and is adapted from the Apache-2.0 licensed
 * `@aiden0z/pptx-renderer` preset table. Every builder receives the shape's
 * width/height in the drawing's pixel space and the `avLst` guide adjustments,
 * and returns an SVG path in that same space. Values are `1/100000` of the
 * shape extents, exactly like DrawingML guides.
 */
export type PresetShapeAdjustments = ReadonlyMap<string, number>;

export type PresetShapeGeometry = {
  path: string;
  viewBox: {
    height: number;
    width: number;
  };
};

type PresetShapeBuilder = (
  width: number,
  height: number,
  adjustments: PresetShapeAdjustments
) => string;

function adjustmentValue(adjustments: PresetShapeAdjustments, name: string, fallback: number): number {
  return adjustments.get(name) ?? fallback;
}

function adjustmentFraction(adjustments: PresetShapeAdjustments, name: string, fallback: number): number {
  return adjustmentValue(adjustments, name, fallback) / 100000;
}

/** SVG arc from a start angle to an end angle on an ellipse. */
function ellipseArcPath(
  centerX: number,
  centerY: number,
  radiusX: number,
  radiusY: number,
  startAngleDeg: number,
  endAngleDeg: number
): string {
  const startAngle = (startAngleDeg * Math.PI) / 180;
  const endAngle = (endAngleDeg * Math.PI) / 180;
  const startX = centerX + radiusX * Math.cos(startAngle);
  const startY = centerY + radiusY * Math.sin(startAngle);
  const endX = centerX + radiusX * Math.cos(endAngle);
  const endY = centerY + radiusY * Math.sin(endAngle);
  let sweep = (((endAngleDeg - startAngleDeg) % 360) + 360) % 360;
  if (sweep === 0 && startAngleDeg !== endAngleDeg) sweep = 360;
  const largeArc = sweep > 180 ? 1 : 0;
  return `M${startX},${startY} A${radiusX},${radiusY} 0 ${largeArc},1 ${endX},${endY}`;
}

/** SVG arc endpoint used by rounded corners and action buttons. */
function arcSegment(
  startX: number,
  startY: number,
  radiusX: number,
  radiusY: number,
  startAngleDeg: number,
  sweepDeg: number
): { svg: string; x: number; y: number } {
  const startAngle = (startAngleDeg * Math.PI) / 180;
  const centerX = startX - radiusX * Math.cos(startAngle);
  const centerY = startY - radiusY * Math.sin(startAngle);
  const endAngle = ((startAngleDeg + sweepDeg) * Math.PI) / 180;
  const endX = centerX + radiusX * Math.cos(endAngle);
  const endY = centerY + radiusY * Math.sin(endAngle);
  const largeArc = Math.abs(sweepDeg) > 180 ? 1 : 0;
  const sweepFlag = sweepDeg >= 0 ? 1 : 0;
  return {
    svg: `A${radiusX},${radiusY} 0 ${largeArc},${sweepFlag} ${endX},${endY}`,
    x: endX,
    y: endY
  };
}

/** Evenly spaced star path (points = 2 * spikes). */
function starPath(width: number, height: number, points: number, innerRatio = 0.4): string {
  const centerX = width / 2;
  const centerY = height / 2;
  const outerX = width / 2;
  const outerY = height / 2;
  const innerX = outerX * innerRatio;
  const innerY = outerY * innerRatio;
  const commands: string[] = [];
  const step = (Math.PI * 2) / points;
  for (let index = 0; index < points; index += 1) {
    const outerAngle = (Math.PI * 2 * index) / points - Math.PI / 2;
    const isOuter = index % 2 === 0;
    const radiusX = isOuter ? outerX : innerX;
    const radiusY = isOuter ? outerY : innerY;
    const x = centerX + radiusX * Math.cos(outerAngle);
    const y = centerY + radiusY * Math.sin(outerAngle);
    commands.push(index === 0 ? `M${x},${y}` : `L${x},${y}`);
  }
  commands.push("Z");
  return commands.join(" ");
}

/** Mirror a path horizontally around `width` (used by the curved arrows). */
function flipPathHorizontally(path: string, width: number): string {
  const tokens = path.match(/[MLAZ]|-?\d*\.?\d+(?:e[-+]?\d+)?/gi);
  if (!tokens) return path;
  const output: string[] = [];
  let index = 0;
  while (index < tokens.length) {
    const command = tokens[index++];
    if (!command) break;
    output.push(command);
    if (command === "Z") continue;
    if (command === "M" || command === "L") {
      const x = Number(tokens[index++]);
      const y = Number(tokens[index++]);
      output.push(String(width - x), String(y));
      continue;
    }
    if (command === "A") {
      const rx = tokens[index++];
      const ry = tokens[index++];
      const rotation = tokens[index++];
      const largeArc = tokens[index++];
      const sweep = Number(tokens[index++]);
      const x = Number(tokens[index++]);
      const y = Number(tokens[index++]);
      output.push(rx, ry, rotation, largeArc, String(sweep ? 0 : 1), String(width - x), String(y));
      continue;
    }
    return path;
  }
  return output.join(" ");
}

/**
 * Geometry that is drawn as an open stroke: a fill would either be invisible
 * or paint the wrong area, so these presets always resolve to `noFill`.
 */
const STROKE_ONLY_GEOMETRIES = new Set([
  "arc",
  "bentconnector2",
  "bentconnector3",
  "bentconnector4",
  "bentconnector5",
  "bracepair",
  "bracketpair",
  "curvedconnector2",
  "curvedconnector3",
  "curvedconnector4",
  "curvedconnector5",
  "leftbrace",
  "leftbracket",
  "line",
  "lineinv",
  "rightbrace",
  "rightbracket",
  "straightconnector1"
]);

export function isStrokeOnlyShapeGeometry(geometry: string): boolean {
  return STROKE_ONLY_GEOMETRIES.has(geometry.trim().toLowerCase());
}

function presetBuilder(name: string): PresetShapeBuilder {
  const builder = PRESET_SHAPE_BUILDERS.get(name);
  if (!builder) {
    throw new Error(`Unknown preset shape: ${name}`);
  }
  return builder;
}

const PRESET_SHAPE_BUILDERS = new Map<string, PresetShapeBuilder>();
PRESET_SHAPE_BUILDERS.set("rect", (t: number, n: number) => `M0,0 L${t},0 L${t},${n} L0,${n} Z`);
PRESET_SHAPE_BUILDERS.set("roundRect", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 16667), s = Math.min(t, n) * o;
    return [
        `M${s},0`,
        `L${t - s},0`,
        `A${s},${s} 0 0,1 ${t},${s}`,
        `L${t},${n - s}`,
        `A${s},${s} 0 0,1 ${t - s},${n}`,
        `L${s},${n}`,
        `A${s},${s} 0 0,1 0,${n - s}`,
        `L0,${s}`,
        `A${s},${s} 0 0,1 ${s},0`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("plaque", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(Math.max(adjustmentValue(e, "adj", 16667), 0), 5e4), s = Math.min(t, n) * o / 1e5, i = t - s, r = n - s, c = arcSegment(0, s, s, s, 90, -90), l = arcSegment(i, 0, s, s, 180, -90), a = arcSegment(t, r, s, s, 270, -90), d = arcSegment(s, n, s, s, 0, -90);
    return [
        `M0,${s}`,
        c.svg,
        `L${i},0`,
        l.svg,
        `L${t},${r}`,
        a.svg,
        `L${s},${n}`,
        d.svg,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("cornerTabs", (t: number, n: number) => {
    const e = Math.sqrt(t * t + n * n) / 20;
    return [
        `M0,0 L${e},0 L0,${e} Z`,
        `M${t},0 L${t - e},0 L${t},${e} Z`,
        `M${t},${n} L${t - e},${n} L${t},${n - e} Z`,
        `M0,${n} L${e},${n} L0,${n - e} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("squareTabs", (t: number, n: number) => {
    const e = Math.sqrt(t * t + n * n) / 20;
    return [
        `M0,0 L${e},0 L${e},${e} L0,${e} Z`,
        `M${t - e},0 L${t},0 L${t},${e} L${t - e},${e} Z`,
        `M0,${n - e} L${e},${n - e} L${e},${n} L0,${n} Z`,
        `M${t - e},${n - e} L${t},${n - e} L${t},${n} L${t - e},${n} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("plaqueTabs", (t: number, n: number) => {
    const e = Math.sqrt(t * t + n * n) / 20;
    return [
        `M0,0 L${e},0 A${e},${e} 0 0,1 0,${e} Z`,
        `M${t},0 L${t - e},0 A${e},${e} 0 0,0 ${t},${e} Z`,
        `M0,${n} L0,${n - e} A${e},${e} 0 0,1 ${e},${n} Z`,
        `M${t},${n} L${t - e},${n} A${e},${e} 0 0,1 ${t},${n - e} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("ellipse", (t: number, n: number) => {
    const e = t / 2, o = n / 2;
    return [`M${t},${o}`, `A${e},${o} 0 1,1 0,${o}`, `A${e},${o} 0 1,1 ${t},${o}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("triangle", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 5e4);
    return `M${t * o},0 L${t},${n} L0,${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("isosTriangle", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 5e4);
    return `M${t * o},0 L${t},${n} L0,${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("rtTriangle", (t: number, n: number) => `M0,0 L${t},${n} L0,${n} Z`);
PRESET_SHAPE_BUILDERS.set("diamond", (t: number, n: number) => {
    const e = t / 2, o = n / 2;
    return `M${e},0 L${t},${o} L${e},${n} L0,${o} Z`;
});
PRESET_SHAPE_BUILDERS.set("pentagon", (t: number, n: number) => {
    const e = t / 2, o = e * 105146 / 1e5, s = n / 2 * 110557 / 1e5, i = s, r = o * Math.cos(18 * Math.PI / 180), c = o * Math.cos(54 * Math.PI / 180), l = s * Math.sin(18 * Math.PI / 180), a = s * Math.sin(54 * Math.PI / 180);
    return [
        `M${e - r},${i - l}`,
        // x1, y1 (upper-left)
        `L${e},0`,
        // hc, t (top)
        `L${e + r},${i - l}`,
        // x4, y1 (upper-right)
        `L${e + c},${i + a}`,
        // x3, y2 (lower-right)
        `L${e - c},${i + a}`,
        // x2, y2 (lower-left)
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("hexagon", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max(adjustmentValue(e, "adj", 25e3), 0), o > 0 ? 5e4 * t / o : 5e4), r = n / 2 * 115470 / 1e5, c = o * s / 1e5, l = t - c, a = n / 2, d = r * Math.sin(60 * Math.PI / 180), h = a - d, u = a + d;
    return [
        `M0,${a}`,
        `L${c},${h}`,
        `L${l},${h}`,
        `L${t},${a}`,
        `L${l},${u}`,
        `L${c},${u}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("octagon", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max(adjustmentValue(e, "adj", 29289), 0), 5e4), i = o * s / 1e5, r = t - i, c = n - i;
    return [
        `M0,${i}`,
        `L${i},0`,
        `L${r},0`,
        `L${t},${i}`,
        `L${t},${c}`,
        `L${r},${n}`,
        `L${i},${n}`,
        `L0,${c}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("heptagon", (t: number, n: number) => {
    const e = t / 2, o = e * 102572 / 1e5, s = n / 2 * 105210 / 1e5, i = n / 2 * 105210 / 1e5, r = o * 97493 / 1e5, c = o * 78183 / 1e5, l = o * 43388 / 1e5, a = s * 62349 / 1e5, d = s * 22252 / 1e5, h = s * 90097 / 1e5;
    return [
        `M${e - r},${i + d}`,
        // x1, y2 (left)
        `L${e - c},${i - a}`,
        // x2, y1 (upper-left)
        `L${e},0`,
        // hc, t (top: svc - shd2 = 0)
        `L${e + c},${i - a}`,
        // x5, y1 (upper-right)
        `L${e + r},${i + d}`,
        // x6, y2 (right)
        `L${e + l},${i + h}`,
        // x4, y3 (lower-right)
        `L${e - l},${i + h}`,
        // x3, y3 (lower-left)
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("decagon", (t: number, n: number) => {
    const e = t / 2, o = n / 2, s = o * 105146 / 1e5, i = e * Math.cos(36 * Math.PI / 180), r = e * Math.cos(72 * Math.PI / 180), c = s * Math.sin(72 * Math.PI / 180), l = s * Math.sin(36 * Math.PI / 180);
    return [
        `M0,${o}`,
        // l, vc
        `L${e - i},${o - l}`,
        // x1, y2
        `L${e - r},${o - c}`,
        // x2, y1
        `L${e + r},${o - c}`,
        // x3, y1
        `L${e + i},${o - l}`,
        // x4, y2
        `L${t},${o}`,
        // r, vc
        `L${e + i},${o + l}`,
        // x4, y3
        `L${e + r},${o + c}`,
        // x3, y4
        `L${e - r},${o + c}`,
        // x2, y4
        `L${e - i},${o + l}`,
        // x1, y3
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("dodecagon", (t: number, n: number) => {
    const e = t * 2894 / 21600, o = t * 7906 / 21600, s = t * 13694 / 21600, i = t * 18706 / 21600, r = n * 2894 / 21600, c = n * 7906 / 21600, l = n * 13694 / 21600, a = n * 18706 / 21600;
    return [
        `M0,${c}`,
        `L${e},${r}`,
        `L${o},0`,
        `L${s},0`,
        `L${i},${r}`,
        `L${t},${c}`,
        `L${t},${l}`,
        `L${i},${a}`,
        `L${s},${n}`,
        `L${o},${n}`,
        `L${e},${a}`,
        `L0,${l}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("parallelogram", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = o > 0 ? 1e5 * t / o : 1e5, i = Math.min(Math.max(adjustmentValue(e, "adj", 25e3), 0), s), r = o * i / 1e5, c = t - r;
    return `M0,${n} L${r},0 L${t},0 L${c},${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("trapezoid", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = o > 0 ? 5e4 * t / o : 5e4, i = Math.min(Math.max(adjustmentValue(e, "adj", 25e3), 0), s), r = o * i / 1e5, c = t - r;
    return `M0,${n} L${r},0 L${c},0 L${t},${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("nonIsoscelesTrapezoid", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = o > 0 ? 5e4 * t / o : 5e4, i = Math.min(Math.max(adjustmentValue(e, "adj1", 25e3), 0), s), r = Math.min(Math.max(adjustmentValue(e, "adj2", 25e3), 0), s), c = o * i / 1e5, l = o * r / 1e5, a = t - l;
    return `M0,${n} L${c},0 L${a},0 L${t},${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("corner", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max(adjustmentFraction(e, "adj1", 5e4), 0), 1), i = Math.min(Math.max(adjustmentFraction(e, "adj2", 5e4), 0), 1), r = o * i, c = o * s, l = n - c;
    return ["M0,0", `L${r},0`, `L${r},${l}`, `L${t},${l}`, `L${t},${n}`, `L0,${n}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("diagStripe", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(Math.max(adjustmentFraction(e, "adj", 5e4), 0), 1), s = t * o;
    return [`M0,${n * o}`, `L${s},0`, `L${t},0`, `L0,${n}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("star4", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 12500) * 2;
    return starPath(t, n, 4, Math.min(Math.max(o, 0), 1));
});
PRESET_SHAPE_BUILDERS.set("star5", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj")) ?? 19098, s = Math.min(Math.max(o, 0), 5e4), i = 105146, r = 110557, c = t / 2 * i / 1e5, l = n / 2 * r / 1e5, a = n / 2 * r / 1e5, d = c * s / 5e4, h = l * s / 5e4, u = t / 2, x = 2 * Math.PI / 5, p = x / 2, $ = -Math.PI / 2, g = [];
    for (let f = 0; f < 5; f++) {
        const y = $ + x * f, m = y + p, b = u + c * Math.cos(y), M = a + l * Math.sin(y), L = u + d * Math.cos(m), v = a + h * Math.sin(m);
        g.push(f === 0 ? `M${b},${M}` : `L${b},${M}`), g.push(`L${L},${v}`);
    }
    return g.push("Z"), g.join(" ");
});
PRESET_SHAPE_BUILDERS.set("star6", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj")) ?? 28868, s = Math.min(Math.max(o, 0), 5e4), r = t / 2 * 115470 / 1e5, c = n / 2, l = r * s / 5e4, a = c * s / 5e4, d = t / 2, h = n / 2, u = 2 * Math.PI / 6, x = u / 2, p = -Math.PI / 2, $ = [];
    for (let g = 0; g < 6; g++) {
        const f = p + u * g, y = f + x, m = d + r * Math.cos(f), b = h + c * Math.sin(f), M = d + l * Math.cos(y), L = h + a * Math.sin(y);
        $.push(g === 0 ? `M${m},${b}` : `L${m},${b}`), $.push(`L${M},${L}`);
    }
    return $.push("Z"), $.join(" ");
});
PRESET_SHAPE_BUILDERS.set("star7", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj")) ?? 34601, s = Math.min(Math.max(o, 0), 5e4), i = t / 2 * 102572 / 1e5, r = n / 2 * 105210 / 1e5, c = r, l = i * s / 5e4, a = r * s / 5e4, d = t / 2, h = 2 * Math.PI / 7, u = h / 2, x = -Math.PI / 2, p = [];
    for (let $ = 0; $ < 7; $++) {
        const g = x + h * $, f = g + u, y = d + i * Math.cos(g), m = c + r * Math.sin(g), b = d + l * Math.cos(f), M = c + a * Math.sin(f);
        p.push($ === 0 ? `M${y},${m}` : `L${y},${m}`), p.push(`L${b},${M}`);
    }
    return p.push("Z"), p.join(" ");
});
PRESET_SHAPE_BUILDERS.set("star8", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 37500) * 2;
    return starPath(t, n, 8, Math.min(Math.max(o, 0), 1));
});
PRESET_SHAPE_BUILDERS.set("star10", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj")) ?? 42533, s = Math.min(Math.max(o, 0), 5e4), r = t / 2 * 105146 / 1e5, c = n / 2, l = r * s / 5e4, a = c * s / 5e4, d = t / 2, h = n / 2, u = 2 * Math.PI / 10, x = u / 2, p = -Math.PI / 2, $ = [];
    for (let g = 0; g < 10; g++) {
        const f = p + u * g, y = f + x, m = d + r * Math.cos(f), b = h + c * Math.sin(f), M = d + l * Math.cos(y), L = h + a * Math.sin(y);
        $.push(g === 0 ? `M${m},${b}` : `L${m},${b}`), $.push(`L${M},${L}`);
    }
    return $.push("Z"), $.join(" ");
});
PRESET_SHAPE_BUILDERS.set("star12", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 37500) * 2;
    return starPath(t, n, 12, Math.min(Math.max(o, 0), 1));
});
PRESET_SHAPE_BUILDERS.set("star16", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 37500) * 2;
    return starPath(t, n, 16, Math.min(Math.max(o, 0), 1));
});
PRESET_SHAPE_BUILDERS.set("star24", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 37500) * 2;
    return starPath(t, n, 24, Math.min(Math.max(o, 0), 1));
});
PRESET_SHAPE_BUILDERS.set("star32", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 37500) * 2;
    return starPath(t, n, 32, Math.min(Math.max(o, 0), 1));
});
PRESET_SHAPE_BUILDERS.set("line", (t: number, n: number) => {
    const e = n || 1, o = t || 1;
    return t === 0 ? `M0.5,0 L0.5,${e}` : n === 0 ? `M0,0.5 L${o},0.5` : `M0,0 L${t},${n}`;
});
PRESET_SHAPE_BUILDERS.set("lineInv", (t: number, n: number) => {
    const e = n || 1, o = t || 1;
    return t === 0 ? `M0.5,0 L0.5,${e}` : n === 0 ? `M0,0.5 L${o},0.5` : `M${t},0 L0,${n}`;
});
PRESET_SHAPE_BUILDERS.set("straightConnector1", (t: number, n: number) => {
    const e = n || 1, o = t || 1;
    return t === 0 ? `M0.5,0 L0.5,${e}` : n === 0 ? `M0,0.5 L${o},0.5` : `M0,0 L${t},${n}`;
});
PRESET_SHAPE_BUILDERS.set("bentConnector2", (t: number, n: number) => `M0,0 L${t},0 L${t},${n}`);
PRESET_SHAPE_BUILDERS.set("bentConnector3", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 5e4), s = t * o;
    return `M0,0 L${s},0 L${s},${n} L${t},${n}`;
});
PRESET_SHAPE_BUILDERS.set("bentConnector4", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 5e4), s = adjustmentFraction(e, "adj2", 5e4), i = t * o, r = n * s;
    return `M0,0 L${i},0 L${i},${r} L${t},${r} L${t},${n}`;
});
PRESET_SHAPE_BUILDERS.set("curvedConnector2", (t: number, n: number) => `M0,0 C${t / 2},0 ${t},${n / 2} ${t},${n}`);
PRESET_SHAPE_BUILDERS.set("curvedConnector3", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = t * adjustmentFraction(e, "adj1", 5e4), s = o / 2, i = (t + o) / 2, r = n / 2, c = n / 4, l = n * 3 / 4;
    return `M0,0 C${s},0 ${o},${c} ${o},${r} C${o},${l} ${i},${n} ${t},${n}`;
});
PRESET_SHAPE_BUILDERS.set("curvedConnector4", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = t * adjustmentFraction(e, "adj1", 5e4), s = n * adjustmentFraction(e, "adj2", 5e4), i = o / 2, r = (t + o) / 2, c = (o + r) / 2, l = (r + t) / 2, a = s / 2, d = a / 2, h = (a + s) / 2, u = (n + s) / 2;
    return [
        "M0,0",
        `C${i},0 ${o},${d} ${o},${a}`,
        `C${o},${h} ${c},${s} ${r},${s}`,
        `C${l},${s} ${t},${u} ${t},${n}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("curvedConnector5", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = t * adjustmentFraction(e, "adj1", 5e4), s = n * adjustmentFraction(e, "adj2", 5e4), i = t * adjustmentFraction(e, "adj3", 5e4), r = (o + i) / 2, c = o / 2, l = (o + r) / 2, a = (i + r) / 2, d = (i + t) / 2, h = s / 2, u = h / 2, x = (h + s) / 2, p = (n + s) / 2, $ = (p + s) / 2, g = (p + n) / 2;
    return [
        "M0,0",
        `C${c},0 ${o},${u} ${o},${h}`,
        `C${o},${x} ${l},${s} ${r},${s}`,
        `C${a},${s} ${i},${$} ${i},${p}`,
        `C${i},${g} ${d},${n} ${t},${n}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("bentConnector5", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 5e4), s = adjustmentFraction(e, "adj2", 5e4), i = adjustmentFraction(e, "adj3", 5e4), r = t * o, c = n * s, l = t * i;
    return `M0,0 L${r},0 L${r},${c} L${l},${c} L${l},${n} L${t},${n}`;
});
PRESET_SHAPE_BUILDERS.set("rightArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 5e4), s = adjustmentFraction(e, "adj2", 5e4), i = Math.min(t, n), r = n * o / 2, c = i * s, l = n / 2, a = t - c;
    return [
        `M0,${l - r}`,
        `L${a},${l - r}`,
        `L${a},0`,
        `L${t},${l}`,
        `L${a},${n}`,
        `L${a},${l + r}`,
        `L0,${l + r}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("leftArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 5e4), s = adjustmentFraction(e, "adj2", 5e4), i = Math.min(t, n), r = n * o / 2, c = i * s, l = n / 2;
    return [
        `M${t},${l - r}`,
        `L${c},${l - r}`,
        `L${c},0`,
        `L0,${l}`,
        `L${c},${n}`,
        `L${c},${l + r}`,
        `L${t},${l + r}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("upArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 5e4), s = adjustmentFraction(e, "adj2", 5e4), i = t * o / 2, r = n * s, c = t / 2;
    return [
        `M${c - i},${n}`,
        `L${c - i},${r}`,
        `L0,${r}`,
        `L${c},0`,
        `L${t},${r}`,
        `L${c + i},${r}`,
        `L${c + i},${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("downArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 5e4), s = adjustmentFraction(e, "adj2", 5e4), i = t * o / 2, r = n * s, c = t / 2, l = n - r;
    return [
        `M${c - i},0`,
        `L${c + i},0`,
        `L${c + i},${l}`,
        `L${t},${l}`,
        `L${c},${n}`,
        `L0,${l}`,
        `L${c - i},${l}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("downArrowCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 25e3, s = (e == null ? void 0 : e.get("adj2")) ?? 25e3, i = (e == null ? void 0 : e.get("adj3")) ?? 25e3, r = (e == null ? void 0 : e.get("adj4")) ?? 64977, c = Math.min(t, n), l = Math.max(0, Math.min(s, 5e4 * t / Math.max(c, 1))), a = Math.max(0, Math.min(o, l * 2)), d = Math.max(0, Math.min(i, 1e5 * n / Math.max(c, 1))), h = d * c / Math.max(n, 1), u = Math.max(0, Math.min(r, 1e5 - h)), x = t / 2, p = c * l / 1e5, $ = c * a / 2e5, g = x - p, f = x - $, y = x + $, m = x + p, b = n - c * d / 1e5, M = n * u / 1e5;
    return [
        "M0,0",
        `L${t},0`,
        `L${t},${M}`,
        `L${y},${M}`,
        `L${y},${b}`,
        `L${m},${b}`,
        `L${x},${n}`,
        `L${g},${b}`,
        `L${f},${b}`,
        `L${f},${M}`,
        `L0,${M}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("rightArrowCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = 5e4 * n / Math.max(o, 1), i = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 25e3, s)), r = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 25e3, i * 2)), c = 1e5 * t / Math.max(o, 1), l = Math.max(0, Math.min((e == null ? void 0 : e.get("adj3")) ?? 25e3, c)), a = l * o / Math.max(t, 1), d = Math.max(0, Math.min((e == null ? void 0 : e.get("adj4")) ?? 64977, 1e5 - a)), h = n / 2, u = o * i / 1e5, x = o * r / 2e5, p = h - u, $ = h - x, g = h + x, f = h + u, y = o * l / 1e5, m = t - y, b = t * d / 1e5;
    return [
        "M0,0",
        `L${b},0`,
        `L${b},${$}`,
        `L${m},${$}`,
        `L${m},${p}`,
        `L${t},${h}`,
        `L${m},${f}`,
        `L${m},${g}`,
        `L${b},${g}`,
        `L${b},${n}`,
        `L0,${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("leftArrowCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = 5e4 * n / Math.max(o, 1), i = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 25e3, s)), r = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 25e3, i * 2)), c = 1e5 * t / Math.max(o, 1), l = Math.max(0, Math.min((e == null ? void 0 : e.get("adj3")) ?? 25e3, c)), a = l * o / Math.max(t, 1), d = Math.max(0, Math.min((e == null ? void 0 : e.get("adj4")) ?? 64977, 1e5 - a)), h = n / 2, u = o * i / 1e5, x = o * r / 2e5, p = h - u, $ = h - x, g = h + x, f = h + u, y = o * l / 1e5, m = t * d / 1e5, b = t - m;
    return [
        `M0,${h}`,
        `L${y},${p}`,
        `L${y},${$}`,
        `L${b},${$}`,
        `L${b},0`,
        `L${t},0`,
        `L${t},${n}`,
        `L${b},${n}`,
        `L${b},${g}`,
        `L${y},${g}`,
        `L${y},${f}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("upArrowCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = 5e4 * t / Math.max(o, 1), i = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 25e3, s)), r = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 25e3, i * 2)), c = 1e5 * n / Math.max(o, 1), l = Math.max(0, Math.min((e == null ? void 0 : e.get("adj3")) ?? 25e3, c)), a = l * o / Math.max(n, 1), d = Math.max(0, Math.min((e == null ? void 0 : e.get("adj4")) ?? 64977, 1e5 - a)), h = t / 2, u = o * i / 1e5, x = o * r / 2e5, p = h - u, $ = h - x, g = h + x, f = h + u, y = o * l / 1e5, m = n * d / 1e5, b = n - m;
    return [
        `M0,${b}`,
        `L${$},${b}`,
        `L${$},${y}`,
        `L${p},${y}`,
        `L${h},0`,
        `L${f},${y}`,
        `L${g},${y}`,
        `L${g},${b}`,
        `L${t},${b}`,
        `L${t},${n}`,
        `L0,${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("upDownArrowCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 25e3, s = (e == null ? void 0 : e.get("adj2")) ?? 25e3, i = (e == null ? void 0 : e.get("adj3")) ?? 25e3, r = (e == null ? void 0 : e.get("adj4")) ?? 48123, c = Math.min(t, n), l = Math.max(0, Math.min(s, 5e4 * t / Math.max(c, 1))), a = Math.max(0, Math.min(o, l * 2)), d = Math.max(0, Math.min(i, 5e4 * n / Math.max(c, 1))), h = d * c / Math.max(n, 1), u = Math.max(0, Math.min(r, 1e5 - h - h)), x = c * l / 1e5, p = c * a / 2e5, $ = t / 2, g = $ - x, f = $ - p, y = $ + p, m = $ + x, b = c * d / 1e5, M = n * u / 2e5, L = n / 2 - M, v = n / 2 + M, k = n - b;
    return [
        `M${$},0`,
        `L${m},${b}`,
        `L${y},${b}`,
        `L${y},${L}`,
        `L${t},${L}`,
        `L${t},${v}`,
        `L${y},${v}`,
        `L${y},${k}`,
        `L${m},${k}`,
        `L${$},${n}`,
        `L${g},${k}`,
        `L${f},${k}`,
        `L${f},${v}`,
        `L0,${v}`,
        `L0,${L}`,
        `L${f},${L}`,
        `L${f},${b}`,
        `L${g},${b}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("leftRightArrowCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 25e3, s = (e == null ? void 0 : e.get("adj2")) ?? 25e3, i = (e == null ? void 0 : e.get("adj3")) ?? 25e3, r = (e == null ? void 0 : e.get("adj4")) ?? 48123, c = Math.min(t, n), l = Math.max(0, Math.min(s, 5e4 * n / Math.max(c, 1))), a = Math.max(0, Math.min(o, l * 2)), d = Math.max(0, Math.min(i, 5e4 * t / Math.max(c, 1))), h = d * c / Math.max(t, 1), u = Math.max(0, Math.min(r, 1e5 - h - h)), x = c * l / 1e5, p = c * a / 2e5, $ = n / 2, g = $ - x, f = $ - p, y = $ + p, m = $ + x, b = c * d / 1e5, M = t * u / 2e5, L = t / 2 - M, v = t / 2 + M, k = t - b;
    return [
        `M0,${$}`,
        `L${b},${g}`,
        `L${b},${f}`,
        `L${L},${f}`,
        `L${L},0`,
        `L${v},0`,
        `L${v},${f}`,
        `L${k},${f}`,
        `L${k},${g}`,
        `L${t},${$}`,
        `L${k},${m}`,
        `L${k},${y}`,
        `L${v},${y}`,
        `L${v},${n}`,
        `L${L},${n}`,
        `L${L},${y}`,
        `L${b},${y}`,
        `L${b},${m}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("uturnArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 25e3, s = (e == null ? void 0 : e.get("adj2")) ?? 25e3, i = (e == null ? void 0 : e.get("adj3")) ?? 25e3, r = (e == null ? void 0 : e.get("adj4")) ?? 43750, c = (e == null ? void 0 : e.get("adj5")) ?? 75e3, l = Math.min(t, n), a = Math.max(0, Math.min(s, 25e3)), d = Math.max(0, Math.min(o, a * 2)), u = 1e5 - d * l / Math.max(n, 1), x = Math.max(0, Math.min(i, u * n / Math.max(l, 1))), p = (x + d) * l / Math.max(n, 1), $ = Math.max(p, Math.min(c, 1e5)), g = l * d / 1e5, f = l * a / 1e5, y = g / 2, m = f - y, b = n * $ / 1e5, M = l * x / 1e5, L = b - M, v = t - m, k = Math.min(v / 2, L), A = Math.max(0, Math.min(r, 1e5 * k / Math.max(l, 1))), S = l * A / 1e5, w = Math.max(S - g, 0), F = g + w, C = t - f, E = C - f, P = E + m, B = v - S, R = P - w;
    return [
        `M0,${n}`,
        `L0,${S}`,
        S > 0.1 ? `A${S},${S} 0 0,1 ${S},0` : "L0,0",
        `L${B},0`,
        S > 0.1 ? `A${S},${S} 0 0,1 ${v},${S}` : `L${v},0`,
        `L${v},${L}`,
        `L${t},${L}`,
        `L${C},${b}`,
        `L${E},${L}`,
        `L${P},${L}`,
        `L${P},${F}`,
        w > 0.1 ? `A${w},${w} 0 0,0 ${R},${g}` : `L${R},${g}`,
        `L${F},${g}`,
        w > 0.1 ? `A${w},${w} 0 0,0 ${g},${F}` : `L${g},${F}`,
        `L${g},${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("leftRightArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = n / 2, i = o > 0 ? 5e4 * t / o : 0, r = Math.min(Math.max((e == null ? void 0 : e.get("adj1")) ?? 5e4, 0), 1e5), c = Math.min(Math.max((e == null ? void 0 : e.get("adj2")) ?? 5e4, 0), i), l = o * c / 1e5, a = t - l, d = n * r / 2e5, h = s, u = h - d, x = h + d;
    return [
        `M0,${h}`,
        `L${l},0`,
        `L${l},${u}`,
        `L${a},${u}`,
        `L${a},0`,
        `L${t},${h}`,
        `L${a},${n}`,
        `L${a},${x}`,
        `L${l},${x}`,
        `L${l},${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("leftUpArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 25e3, 5e4)), s = o * 2, i = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 25e3, s)), r = 1e5 - s, c = Math.max(0, Math.min((e == null ? void 0 : e.get("adj3")) ?? 25e3, r)), l = Math.min(t, n), a = l * c / 1e5, d = l * o / 5e4, h = t - d, u = n - d, x = l * o / 1e5, p = t - x, $ = n - x, g = l * i / 2e5, f = p - g, y = p + g, m = $ - g, b = $ + g;
    return [
        `M0,${$}`,
        `L${a},${u}`,
        `L${a},${m}`,
        `L${f},${m}`,
        `L${f},${a}`,
        `L${h},${a}`,
        `L${p},0`,
        `L${t},${a}`,
        `L${y},${a}`,
        `L${y},${b}`,
        `L${a},${b}`,
        `L${a},${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("upDownArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 5e4, s = (e == null ? void 0 : e.get("adj2")) ?? 5e4, i = Math.min(t, n), r = 5e4 * n / Math.max(i, 1), c = Math.max(0, Math.min(s, r)), l = Math.max(0, Math.min(o, 1e5)), a = i * l / 2e5, d = i * c / 1e5, h = t / 2;
    return [
        `M${h},0`,
        `L${t},${d}`,
        `L${h + a},${d}`,
        `L${h + a},${n - d}`,
        `L${t},${n - d}`,
        `L${h},${n}`,
        `L0,${n - d}`,
        `L${h - a},${n - d}`,
        `L${h - a},${d}`,
        `L0,${d}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("notchedRightArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 5e4), s = adjustmentFraction(e, "adj2", 5e4), i = Math.min(t, n), r = n * o / 2, c = i * s, l = n / 2, a = t - c, d = l > 0 ? r * c / l : 0;
    return [
        `M0,${l - r}`,
        `L${a},${l - r}`,
        `L${a},0`,
        `L${t},${l}`,
        `L${a},${n}`,
        `L${a},${l + r}`,
        `L0,${l + r}`,
        `L${d},${l}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("chevron", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 5e4), i = Math.min(t, n) * o;
    return [
        "M0,0",
        `L${t - i},0`,
        `L${t},${n / 2}`,
        `L${t - i},${n}`,
        `L0,${n}`,
        `L${i},${n / 2}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("homePlate", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 5e4), i = Math.min(t, n) * o, r = t - i;
    return ["M0,0", `L${r},0`, `L${t},${n / 2}`, `L${r},${n}`, `L0,${n}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("stripedRightArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = o > 0 ? 84375 * t / o : 84375, i = Math.min(Math.max(adjustmentValue(e, "adj1", 5e4), 0), 1e5), r = Math.min(Math.max(adjustmentValue(e, "adj2", 5e4), 0), s), c = n * i / 2e5, l = o * r / 1e5, a = t - l, d = n / 2, h = d - c, u = d + c, x = o / 32, p = o / 16, $ = o / 8, g = o * 5 / 32;
    return [
        // Stripe 1: 0 to ssd32
        `M0,${h} L${x},${h} L${x},${u} L0,${u} Z`,
        // Stripe 2: ssd16 to ssd8
        `M${p},${h} L${$},${h} L${$},${u} L${p},${u} Z`,
        // Main body + arrowhead: x4 to r
        `M${g},${h}`,
        `L${a},${h}`,
        `L${a},0`,
        `L${t},${d}`,
        `L${a},${n}`,
        `L${a},${u}`,
        `L${g},${u}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("bentArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 25e3, 5e4)), i = s * 2, r = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 25e3, i)), c = Math.max(0, Math.min((e == null ? void 0 : e.get("adj3")) ?? 25e3, 5e4)), l = o * r / 1e5, a = o * s / 1e5, d = l / 2, h = a - d, u = o * c / 1e5, x = t - u, p = n - h, $ = Math.min(x, p), g = $ > 0 ? 1e5 * $ / o : 0, f = Math.max(0, Math.min((e == null ? void 0 : e.get("adj4")) ?? 43750, g)), y = o * f / 1e5, m = Math.max(y - l, 0), b = l + m, M = t - u, L = h + l, v = L + h, k = h + y, A = L + m, S = [
        `M0,${n}`,
        // bottom-left
        `L0,${k}`
        // up left edge to arc start
    ];
    return y > 0.1 ? S.push(`A${y},${y} 0 0,1 ${y},${h}`) : S.push(`L0,${h}`), S.push(`L${M},${h}`, 
    // horizontal to arrowhead base (top)
    `L${M},0`, 
    // up to arrowhead top-left wing
    `L${t},${a}`, 
    // arrowhead tip (pointing right)
    `L${M},${v}`, 
    // arrowhead bottom wing
    `L${M},${L}`, 
    // back to arrowhead base (bottom)
    `L${b},${L}`
    // horizontal back toward bend
    ), m > 0.1 ? S.push(`A${m},${m} 0 0,0 ${l},${A}`) : S.push(`L${l},${L}`), S.push(`L${l},${n}`, 
    // down right side of shaft to bottom
    "Z"), S.join(" ");
});
PRESET_SHAPE_BUILDERS.set("bentUpArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 25e3, 5e4)), s = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 25e3, 5e4)), i = Math.max(0, Math.min((e == null ? void 0 : e.get("adj3")) ?? 25e3, 5e4)), r = Math.min(t, n), c = r * i / 1e5, l = r * s / 5e4, a = t - l, d = r * s / 1e5, h = t - d, u = r * o / 2e5, x = h - u, p = h + u, $ = r * o / 1e5, g = n - $;
    return [
        `M0,${g}`,
        `L${x},${g}`,
        `L${x},${c}`,
        `L${a},${c}`,
        `L${h},0`,
        `L${t},${c}`,
        `L${p},${c}`,
        `L${p},${n}`,
        `L0,${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("curvedRightArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 25e3, s = (e == null ? void 0 : e.get("adj2")) ?? 5e4, i = (e == null ? void 0 : e.get("adj3")) ?? 25e3, r = 5e4, c = 1e5, l = n / 2, a = t, d = n, h = 0, u = 270, x = 180, p = 90, $ = Math.max(Math.min(t, n), 1), g = r * n / $, f = Math.max(0, Math.min(s, g)), y = Math.max(0, Math.min(o, f)), m = $ * y / c, b = $ * f / c, M = (m + b) / 4, L = l - M, v = L * 2, k = v * v, A = m * m, S = Math.max(k - A, 0), F = Math.sqrt(S) * t / Math.max(v, 1e-6), C = c * F / $, E = Math.max(0, Math.min(i, C)), P = $ * E / c, B = L + m, R = t * t, I = P * P, Z = Math.max(R - I, 0), q = Math.sqrt(Z) * L / Math.max(t, 1e-6), Q = L + q, G = B + q, D = (b - m) / 2, O = Q - D, ot = G + D, J = b / 2, W = d - J, it = a - P, lt = Math.atan(q / Math.max(P, 1e-6)), Pt = Math.PI - lt, Rt = -lt, Wt = m / 2, ft = Math.atan2(Wt, Math.max(F, 1e-6)) - Math.PI / 2, Zt = Pt * 180 / Math.PI, Dt = Rt * 180 / Math.PI, Ot = lt * 180 / Math.PI, Ut = ft * 180 / Math.PI, Y = (Bt: number, Ft: number, Xt: number, H: number, N: number, et: number) => {
        const nt = N * Math.PI / 180, at = et * Math.PI / 180, bt = Bt + Xt * Math.cos(nt), Mt = Ft + H * Math.sin(nt), tt = Bt + Xt * Math.cos(at), ht = Ft + H * Math.sin(at), ut = et - N, zt = Math.abs(ut) > 180 ? 1 : 0, te = ut >= 0 ? 1 : 0;
        return `M${bt},${Mt} A${Xt},${H} 0 ${zt},${te} ${tt},${ht}`;
    };
    return [
        `M${h},${L}`,
        Y(t, L, t, L, x, x + Dt).replace("M", "L"),
        `L${it},${Q}`,
        `L${it},${O}`,
        `L${a},${W}`,
        `L${it},${ot}`,
        `L${it},${G}`,
        Y(t, B, t, L, Zt, Zt + Ot).replace("M", "L"),
        "Z",
        Y(t, L, t, L, x, x + p),
        `L${a},${m}`,
        Y(t, B, t, L, u, u + Ut).replace("M", "L"),
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("curvedLeftArrow", (t: number, n: number, e: PresetShapeAdjustments) => flipPathHorizontally(presetBuilder("curvedRightArrow")(t, n, e), t));
PRESET_SHAPE_BUILDERS.set("curvedUpArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (G: number, T: number, D: number, O: number, ot: number, J: number) => {
        const W = ot * Math.PI / 180, it = J * Math.PI / 180, lt = G + D * Math.cos(W), Pt = T + O * Math.sin(W), Rt = G + D * Math.cos(it), Wt = T + O * Math.sin(it), Nt = J - ot, ft = Math.abs(Nt) > 180 ? 1 : 0, Zt = Nt >= 0 ? 1 : 0;
        return `M${lt},${Pt} A${D},${O} 0 ${ft},${Zt} ${Rt},${Wt}`;
    }, s = Math.min(t, n), i = t / 2, r = (e == null ? void 0 : e.get("adj1")) ?? 25e3, c = (e == null ? void 0 : e.get("adj2")) ?? 5e4, l = (e == null ? void 0 : e.get("adj3")) ?? 25e3, a = 5e4 * t / Math.max(s, 1), d = Math.max(0, Math.min(c, a)), h = Math.max(0, Math.min(r, 1e5)), u = s * h / 1e5, x = s * d / 1e5, p = (u + x) / 4, $ = i - p, g = $ * 2, f = Math.sqrt(Math.max(g * g - u * u, 0)) * n / Math.max(g, 1), y = 1e5 * f / Math.max(s, 1), m = Math.max(0, Math.min(l, y)), b = s * m / 1e5, M = $ + u, L = Math.sqrt(Math.max(n * n - b * b, 0)) * $ / Math.max(n, 1), v = $ + L, k = M + L, A = (x - u) / 2, S = v - A, w = k + A, F = t - x / 2, C = b, E = Math.atan2(L, b), P = Math.atan2(u / 2, f), B = Math.PI / 2 - P, R = P - E, I = Math.PI / 2 - E, Z = B * 180 / Math.PI, U = R * 180 / Math.PI, q = I * 180 / Math.PI, Q = E * 180 / Math.PI;
    return [
        o($, 0, $, n, Z, Z + U),
        `L${v},${C}`,
        `L${S},${C}`,
        `L${F},0`,
        `L${w},${C}`,
        `L${k},${C}`,
        o(M, 0, $, n, q, q + Q).replace("M", "L"),
        `L${$},${n}`,
        o($, 0, $, n, 90, 180).replace("M", "L"),
        `L${u},0`,
        o(M, 0, $, n, 180, 90).replace("M", "L"),
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("curvedDownArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (Q: number, G: number, T: number, D: number, O: number, ot: number) => {
        const J = O * Math.PI / 180, W = ot * Math.PI / 180, it = Q + T * Math.cos(J), lt = G + D * Math.sin(J), Pt = Q + T * Math.cos(W), Rt = G + D * Math.sin(W), Wt = ot - O, Nt = Math.abs(Wt) > 180 ? 1 : 0, ft = Wt >= 0 ? 1 : 0;
        return `M${it},${lt} A${T},${D} 0 ${Nt},${ft} ${Pt},${Rt}`;
    }, s = Math.min(t, n), i = t / 2, r = (e == null ? void 0 : e.get("adj1")) ?? 25e3, c = (e == null ? void 0 : e.get("adj2")) ?? 5e4, l = (e == null ? void 0 : e.get("adj3")) ?? 25e3, a = 5e4 * t / Math.max(s, 1), d = Math.max(0, Math.min(c, a)), h = Math.max(0, Math.min(r, 1e5)), u = s * h / 1e5, x = s * d / 1e5, p = (u + x) / 4, $ = i - p, g = $ * 2, f = Math.sqrt(Math.max(g * g - u * u, 0)) * n / Math.max(g, 1), y = 1e5 * f / Math.max(s, 1), m = Math.max(0, Math.min(l, y)), b = s * m / 1e5, M = $ + u, L = Math.sqrt(Math.max(n * n - b * b, 0)) * $ / Math.max(n, 1), v = $ + L, k = M + L, A = (x - u) / 2, S = v - A, w = k + A, F = t - x / 2, C = n - b, P = Math.atan2(L, b) * 180 / Math.PI, R = Math.atan2(u / 2, f) * 180 / Math.PI, I = 270 + P, Z = 270 - R, U = R - 90, q = 90 + R;
    return [
        `M${F},${n}`,
        `L${S},${C}`,
        `L${v},${C}`,
        o($, n, $, n, I, I - P).replace("M", "L"),
        `L${M},0`,
        o(M, n, $, n, 270, 270 + P).replace("M", "L"),
        `L${v + u},${C}`,
        `L${w},${C}`,
        "Z",
        `M${M},0`,
        o(M, n, $, n, Z, Z + U).replace("M", "L"),
        o($, n, $, n, 180, 180 + q).replace("M", "L"),
        "Z"
    ].join(" ");
});
function circularArrowPath(
  width: number,
  height: number,
  adjustments: PresetShapeAdjustments,
  mirrored = false,
  kind = "circularArrow"
): string {
  const t = width, n = height, e = adjustments, o = mirrored, s = kind;
    const i = t / 2, r = n / 2, c = t / 2, l = n / 2, a = Math.min(t, n), d = 108e5, h = (Ht: number) => Ht / 6e4 * Math.PI / 180, u = (Ht: number, oe: number) => Ht * Math.sin(h(oe)), x = (Ht: number, oe: number) => Ht * Math.cos(h(oe)), p = (Ht: number, oe: number, Ve: number) => Ht * Math.cos(Math.atan2(Ve, oe)), $ = (Ht: number, oe: number, Ve: number) => Ht * Math.sin(Math.atan2(Ve, oe)), g = (Ht: number, oe: number) => Math.atan2(oe, Ht) * 180 / Math.PI * 6e4, f = (Ht: number, oe: number, Ve: number) => Math.sqrt(Ht * Ht + oe * oe + Ve * Ve), y = s === "leftCircularArrow", m = (e == null ? void 0 : e.get("adj1")) ?? 12500, b = (e == null ? void 0 : e.get("adj2")) ?? (y ? -1142319 : 1142319), M = (e == null ? void 0 : e.get("adj3")) ?? (y ? 1142319 : 20457681), L = (e == null ? void 0 : e.get("adj4")) ?? 108e5, v = (e == null ? void 0 : e.get("adj5")) ?? 12500, k = Math.max(0, Math.min(v, 25e3)), A = k * 2, S = Math.max(0, Math.min(m, A)), w = Math.max(1, Math.min(M, 21599999)), F = Math.max(0, Math.min(L, 21599999)), C = a * S / 1e5, E = a * k / 1e5, P = C / 2, B = c + P - E, R = l + P - E, I = B - C, Z = R - C, U = I + P, q = Z + P, Q = u(U, w), G = x(q, w), T = p(U, G, Q), D = $(q, G, Q), O = i + T, ot = r + D, J = Math.min(I, Z), W = T * T, it = D * D, lt = J * J, Pt = W - lt, Rt = it - lt, Wt = it !== 0 ? Pt * Rt / W : 0, ft = 1 - (it !== 0 ? Wt / it : 0), Zt = Math.sqrt(Math.max(0, ft)), Dt = T !== 0 ? Pt / T : 0, Ot = D !== 0 ? Dt / D : 0, Ut = Ot !== 0 ? (1 + Zt) / Ot : 0, Y = g(1, Ut), Bt = Y + 216e5, Xt = (Y >= 0 ? Y : Bt) - w, H = Xt + 216e5, N = Xt >= 0 ? Xt : H, et = N - d, nt = N - 216e5, at = et >= 0 ? nt : N, bt = Math.abs(at);
    let Mt;
    if (y) {
        const Ht = -bt, oe = -Math.abs(b);
        Mt = Math.max(Ht, Math.min(oe, 0));
    }
    else
        Mt = Math.max(0, Math.min(b, bt));
    const tt = w + Mt, ht = u(U, tt), ut = x(q, tt), zt = p(U, ut, ht), te = $(q, ut, ht), Me = i + zt, yt = r + te, ee = u(B, F), ce = x(R, F), rt = p(B, ce, ee), $t = $(R, ce, ee), vt = i + rt, j = r + $t, At = x(E, tt), It = u(E, tt), ct = O + At, xt = ot + It, St = O - At, dt = ot - It, Vt = St - i, Ct = dt - r, Lt = ct - i, kt = xt - r, gt = Math.min(B, R), Qt = B !== 0 ? Vt * gt / B : 0, Yt = R !== 0 ? Ct * gt / R : 0, mt = B !== 0 ? Lt * gt / B : 0, _t = R !== 0 ? kt * gt / R : 0, fe = mt - Qt, $e = _t - Yt, le = f(fe, $e, 0), Jt = Qt * _t, Gt = mt * Yt, ae = Jt - Gt, cn = gt * gt, xe = le * le, Ue = cn * xe, qe = ae * ae, Do = Ue - qe, Pe = Math.max(Do, 0), Xn = Math.sqrt(Pe), Yn = $e * -1 >= 0 ? -1 : 1, Cn = Yn * fe * Xn, ye = ae * $e, Fn = xe !== 0 ? (ye + Cn) / xe : 0, kn = ye - Cn, ln = xe !== 0 ? kn / xe : 0, an = Math.abs($e) * Xn, qn = ae * fe * -1, No = xe !== 0 ? (qn + an) / xe : 0, ni = qn - an, oi = xe !== 0 ? ni / xe : 0, Bl = mt - Fn, Rl = mt - ln, Il = _t - No, Tl = _t - oi, zl = f(Bl, Il, 0), si = f(Rl, Tl, 0) - zl, Dl = si >= 0 ? Fn : ln, Ol = si >= 0 ? No : oi, ii = gt !== 0 ? Dl * B / gt : 0, ri = gt !== 0 ? Ol * R / gt : 0, Zo = i + ii, Go = r + ri, Qn = I !== 0 ? Vt * J / I : 0, Kn = Z !== 0 ? Ct * J / Z : 0, ci = I !== 0 ? Lt * J / I : 0, li = Z !== 0 ? kt * J / Z : 0, Ho = ci - Qn, Wo = li - Kn, ai = f(Ho, Wo, 0), Nl = Qn * li, Zl = ci * Kn, Jn = Nl - Zl, Gl = J * J, ze = ai * ai, Hl = Gl * ze, Wl = Jn * Jn, Ul = Hl - Wl, Vl = Math.max(Ul, 0), di = Math.sqrt(Vl), hi = Yn * Ho * di, ui = Jn * Wo, fi = ze !== 0 ? (ui + hi) / ze : 0, _l = ui - hi, $i = ze !== 0 ? _l / ze : 0, pi = Math.abs(Wo) * di, xi = Jn * Ho * -1, yi = ze !== 0 ? (xi + pi) / ze : 0, Xl = xi - pi, gi = ze !== 0 ? Xl / ze : 0, Yl = Qn - fi, ql = Qn - $i, Ql = Kn - yi, Kl = Kn - gi, Jl = f(Yl, Ql, 0), mi = f(ql, Kl, 0) - Jl, jl = mi >= 0 ? fi : $i, ta = mi >= 0 ? yi : gi, bi = J !== 0 ? jl * I / J : 0, Mi = J !== 0 ? ta * Z / J : 0, Uo = i + bi, Vo = r + Mi, _o = g(bi, Mi), ea = _o + 216e5, Xo = _o >= 0 ? _o : ea, dn = F - Xo;
    let Yo, wn;
    if (y) {
        const Ht = dn >= 0 ? dn : dn + 216e5;
        Yo = Xo + Ht, wn = -Ht;
    }
    else
        Yo = Xo, wn = dn >= 0 ? dn - 216e5 : dn;
    const na = Zo - Uo, oa = Go - Vo, jn = f(na, oa, 0) / 2 - E, Li = jn >= 0 ? Zo : ct, vi = jn >= 0 ? Go : xt, Ai = jn >= 0 ? Uo : St, Si = jn >= 0 ? Vo : dt, qo = g(ii, ri), sa = qo + 216e5, hn = (qo >= 0 ? qo : sa) - F;
    let Qo, En;
    if (y) {
        const Ht = hn >= 0 ? hn - 216e5 : hn;
        Qo = F + Ht, En = -Ht;
    }
    else {
        const Ht = hn >= 0 ? hn : hn + 216e5;
        Qo = F, En = Ht;
    }
    const Ci = Qo + En, Fi = u(B, Ci), ki = x(R, Ci), wi = i + p(B, ki, Fi), Ei = r + $(R, ki, Fi), Pi = Yo + wn, Bi = u(I, Pi), Ri = x(Z, Pi), Ii = i + p(I, Ri, Bi), Ti = r + $(Z, Ri, Bi), zi = Math.abs(En / 6e4) > 180 ? 1 : 0, Di = En > 0 ? 1 : 0, Oi = Math.abs(wn / 6e4) > 180 ? 1 : 0, Ni = wn > 0 ? 1 : 0;
    if (y) {
        const Ht = u(I, F), oe = x(Z, F), Ve = i + p(I, oe, Ht), ia = r + $(Z, oe, Ht);
        return [
            `M${vt},${j}`,
            `L${Ve},${ia}`,
            `A${I},${Z} 0 ${Oi},${Ni} ${Ii},${Ti}`,
            `L${Ai},${Si}`,
            `L${Me},${yt}`,
            `L${Li},${vi}`,
            `L${Zo},${Go}`,
            `A${B},${R} 0 ${zi},${Di} ${wi},${Ei}`,
            "Z"
        ].join(" ");
    }
    return [
        `M${vt},${j}`,
        `A${B},${R} 0 ${zi},${Di} ${wi},${Ei}`,
        `L${Li},${vi}`,
        `L${Me},${yt}`,
        `L${Ai},${Si}`,
        `L${Uo},${Vo}`,
        `A${I},${Z} 0 ${Oi},${Ni} ${Ii},${Ti}`,
        "Z"
    ].join(" ");
}
PRESET_SHAPE_BUILDERS.set("circularArrow", (t: number, n: number, e: PresetShapeAdjustments) => circularArrowPath(t, n, e, !1, "circularArrow"));
PRESET_SHAPE_BUILDERS.set("leftCircularArrow", (t: number, n: number, e: PresetShapeAdjustments) => circularArrowPath(t, n, e, !1, "leftCircularArrow"));
PRESET_SHAPE_BUILDERS.set("leftRightCircularArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = t / 400, s = n / 280, i = (S: number, w: number) => ({ x: S * o, y: w * s }), r = i(35, 140), c = i(19.9536, 89.9471), l = i(33.4296, 89.9471), a = i(74.6127, 28.1974), d = i(182.5744, 0.5489), h = i(274.5688, 28.1924), u = i(315.4978, 40.4912), x = i(348.2481, 62.4743), p = i(366.5707, 89.9471), $ = i(380.0463, 89.9471), g = i(365, 140), f = i(310.0463, 89.9471), y = i(320.9838, 89.9471), m = i(274.3848, 50.3095), b = i(182.4425, 40.5864), M = i(115.6249, 68.2298), L = i(101.3589, 74.1319), v = i(88.9651, 81.4842), k = i(79.0159, 89.947), A = i(89.9536, 89.9471);
    return [
        `M${r.x},${r.y}`,
        `L${c.x},${c.y}`,
        `L${l.x},${l.y}`,
        `C${a.x},${a.y} ${d.x},${d.y} ${h.x},${h.y}`,
        `C${u.x},${u.y} ${x.x},${x.y} ${p.x},${p.y}`,
        `L${$.x},${$.y}`,
        `L${g.x},${g.y}`,
        `L${f.x},${f.y}`,
        `L${y.x},${y.y}`,
        `C${m.x},${m.y} ${b.x},${b.y} ${M.x},${M.y}`,
        `C${L.x},${L.y} ${v.x},${v.y} ${k.x},${k.y}`,
        `L${A.x},${A.y}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("quadArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 22500, s = (e == null ? void 0 : e.get("adj2")) ?? 22500, i = (e == null ? void 0 : e.get("adj3")) ?? 22500, r = n / 2, c = t / 2, l = Math.min(t, n), a = Math.max(0, Math.min(s, 5e4)), d = Math.max(0, Math.min(o, 2 * a)), h = Math.max(0, Math.min(i, (1e5 - 2 * a) / 2)), u = l * h / 1e5, x = l * a / 1e5, p = c - x, $ = c + x, g = l * d / 2e5, f = c - g, y = c + g, m = t - u, b = r - x, M = r + x, L = r - g, v = r + g, k = n - u;
    return [
        `M0,${r}`,
        `L${u},${b}`,
        `L${u},${L}`,
        `L${f},${L}`,
        `L${f},${u}`,
        `L${p},${u}`,
        `L${c},0`,
        `L${$},${u}`,
        `L${y},${u}`,
        `L${y},${L}`,
        `L${m},${L}`,
        `L${m},${b}`,
        `L${t},${r}`,
        `L${m},${M}`,
        `L${m},${v}`,
        `L${y},${v}`,
        `L${y},${k}`,
        `L${$},${k}`,
        `L${c},${n}`,
        `L${p},${k}`,
        `L${f},${k}`,
        `L${f},${v}`,
        `L${u},${v}`,
        `L${u},${M}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("quadArrowCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = t / 2, i = n / 2, r = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 18515, 5e4)), c = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 18515, r * 2)), l = 5e4 - r, a = Math.max(0, Math.min((e == null ? void 0 : e.get("adj3")) ?? 18515, l)), d = a * 2, h = Math.max(c, Math.min((e == null ? void 0 : e.get("adj4")) ?? 48123, 1e5 - d)), u = o * r / 1e5, x = o * c / 2e5, p = o * a / 1e5, $ = t * h / 2e5, g = n * h / 2e5, f = t - p, y = s - $, m = s + $, b = s - u, M = s + u, L = s - x, v = s + x, k = n - p, A = i - g, S = i + g, w = i - u, F = i + u, C = i - x, E = i + x;
    return [
        `M0,${i}`,
        `L${p},${w}`,
        `L${p},${C}`,
        `L${y},${C}`,
        `L${y},${A}`,
        `L${L},${A}`,
        `L${L},${p}`,
        `L${b},${p}`,
        `L${s},0`,
        `L${M},${p}`,
        `L${v},${p}`,
        `L${v},${A}`,
        `L${m},${A}`,
        `L${m},${C}`,
        `L${f},${C}`,
        `L${f},${w}`,
        `L${t},${i}`,
        `L${f},${F}`,
        `L${f},${E}`,
        `L${m},${E}`,
        `L${m},${S}`,
        `L${v},${S}`,
        `L${v},${k}`,
        `L${M},${k}`,
        `L${s},${n}`,
        `L${b},${k}`,
        `L${L},${k}`,
        `L${L},${S}`,
        `L${y},${S}`,
        `L${y},${E}`,
        `L${p},${E}`,
        `L${p},${F}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("leftRightUpArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 25e3, 5e4)), s = o * 2, i = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 25e3, s)), c = (1e5 - s) / 2, l = Math.max(0, Math.min((e == null ? void 0 : e.get("adj3")) ?? 25e3, c)), a = Math.min(t, n), d = t / 2, h = a * l / 1e5, u = a * o / 1e5, x = d - u, p = d + u, $ = a * i / 2e5, g = d - $, f = d + $, y = t - h, m = a * o / 5e4, b = n - m, M = n - u, L = M - $, v = M + $;
    return [
        `M0,${M}`,
        `L${h},${b}`,
        `L${h},${L}`,
        `L${g},${L}`,
        `L${g},${h}`,
        `L${x},${h}`,
        `L${d},0`,
        `L${p},${h}`,
        `L${f},${h}`,
        `L${f},${L}`,
        `L${y},${L}`,
        `L${y},${b}`,
        `L${t},${M}`,
        `L${y},${n}`,
        `L${y},${v}`,
        `L${h},${v}`,
        `L${h},${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("swooshArrow", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = (e == null ? void 0 : e.get("adj1")) ?? 25e3, i = (e == null ? void 0 : e.get("adj2")) ?? 16667, r = Math.max(1, Math.min(s, 75e3)), c = 7e4 * t / o, l = Math.max(0, Math.min(i, c)), a = n * r / 1e5, d = o * l / 1e5, h = o / 8, u = n / 6, x = Math.PI / 2 / 14, p = Math.tan(x), $ = t - d, g = h, f = h * p, y = $ - f, m = a * p, b = g + a, M = $ + m, L = M + f, v = b + h, A = v / 2, S = n / 20, w = A + S, F = t / 6, C = u + u, E = u / 2, P = b + E, B = t / 4;
    return [
        `M0,${n}`,
        `Q${F},${C} ${$},${g}`,
        `L${y},0`,
        `L${t},${w}`,
        `L${L},${v}`,
        `L${M},${b}`,
        `Q${B},${P} 0,${n}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartProcess", (t: number, n: number) => `M0,0 L${t},0 L${t},${n} L0,${n} Z`);
PRESET_SHAPE_BUILDERS.set("flowChartDecision", (t: number, n: number) => {
    const e = t / 2, o = n / 2;
    return `M${e},0 L${t},${o} L${e},${n} L0,${o} Z`;
});
PRESET_SHAPE_BUILDERS.set("flowChartTerminator", (t: number, n: number) => {
    const e = t * 3475 / 21600, o = t * 18125 / 21600, s = e, i = n / 2;
    return [
        `M${e},0`,
        `L${o},0`,
        `A${s},${i} 0 0,1 ${o},${n}`,
        `L${e},${n}`,
        `A${s},${i} 0 0,1 ${e},0`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartDocument", (t: number, n: number) => {
    const e = n * 17322 / 21600, o = e, s = n * 23922 / 21600, i = n * 20172 / 21600;
    return ["M0,0", `L${t},0`, `L${t},${e}`, `C${t / 2},${o} ${t / 2},${s} 0,${i}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartInputOutput", (t: number, n: number) => {
    const e = t / 5;
    return `M${e},0 L${t},0 L${t - e},${n} L0,${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("flowChartPredefinedProcess", (t: number, n: number) => {
    const e = t / 8;
    return [
        // Outer rectangle
        `M0,0 L${t},0 L${t},${n} L0,${n} Z`,
        // Left inner line
        `M${e},0 L${e},${n}`,
        // Right inner line
        `M${t - e},0 L${t - e},${n}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartAlternateProcess", (t: number, n: number) => {
    const e = Math.min(t, n) / 6;
    return [
        `M${e},0`,
        `L${t - e},0`,
        `A${e},${e} 0 0,1 ${t},${e}`,
        `L${t},${n - e}`,
        `A${e},${e} 0 0,1 ${t - e},${n}`,
        `L${e},${n}`,
        `A${e},${e} 0 0,1 0,${n - e}`,
        `L0,${e}`,
        `A${e},${e} 0 0,1 ${e},0`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartManualInput", (t: number, n: number) => `M0,${n * 0.2} L${t},0 L${t},${n} L0,${n} Z`);
PRESET_SHAPE_BUILDERS.set("flowChartManualOperation", (t: number, n: number) => `M0,0 L${t},0 L${t * 4 / 5},${n} L${t / 5},${n} Z`);
PRESET_SHAPE_BUILDERS.set("flowChartPreparation", (t: number, n: number) => {
    const e = t * 0.2, o = n / 2;
    return `M${e},0 L${t - e},0 L${t},${o} L${t - e},${n} L${e},${n} L0,${o} Z`;
});
PRESET_SHAPE_BUILDERS.set("flowChartData", (t: number, n: number) => {
    const e = t * 0.15;
    return `M${e},0 L${t},0 L${t - e},${n} L0,${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("flowChartInternalStorage", (t: number, n: number) => {
    const e = t / 8, o = n / 8;
    return [
        `M0,0 L${t},0 L${t},${n} L0,${n} Z`,
        `M${e},0 L${e},${n}`,
        `M0,${o} L${t},${o}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartMagneticDisk", (t: number, n: number) => {
    const e = n / 6, o = e, s = n - e;
    return [
        // Top ellipse
        `M0,${o}`,
        `A${t / 2},${e} 0 1,1 ${t},${o}`,
        // Right side down
        `L${t},${s}`,
        // Bottom ellipse
        `A${t / 2},${e} 0 1,1 0,${s}`,
        // Left side up
        `L0,${o}`,
        "Z",
        // Top ellipse visible arc (back half)
        `M${t},${o}`,
        `A${t / 2},${e} 0 1,1 0,${o}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartDelay", (t: number, n: number) => {
    const e = t / 2, o = arcSegment(e, 0, e, n / 2, 270, 180);
    return ["M0,0", `L${e},0`, o.svg, `L0,${n}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartDisplay", (t: number, n: number) => {
    const e = t / 6, o = n / 6, s = e, i = o * 3, r = arcSegment(5 * e, 0, s, i, 270, 180);
    return [`M0,${3 * o}`, `L${e},0`, `L${5 * e},0`, r.svg, `L${e},${n}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartExtract", (t: number, n: number) => `M${t / 2},0 L${t},${n} L0,${n} Z`);
PRESET_SHAPE_BUILDERS.set("flowChartMerge", (t: number, n: number) => `M0,0 L${t},0 L${t / 2},${n} Z`);
PRESET_SHAPE_BUILDERS.set("flowChartOffpageConnector", (t: number, n: number) => {
    const e = n * 0.2;
    return ["M0,0", `L${t},0`, `L${t},${n - e}`, `L${t / 2},${n}`, `L0,${n - e}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartConnector", (t: number, n: number) => {
    const e = t / 2, o = n / 2;
    return [`M${t},${o}`, `A${e},${o} 0 1,1 0,${o}`, `A${e},${o} 0 1,1 ${t},${o}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartSort", (t: number, n: number) => {
    const e = t / 2, o = n / 2;
    return [`M${e},0 L${t},${o} L${e},${n} L0,${o} Z`, `M0,${o} L${t},${o}`].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartCollate", (t: number, n: number) => {
    const e = t / 2, o = n / 2;
    return [
        // top inverted triangle
        `M0,0 L${t},0 L${e},${o} Z`,
        // bottom upright triangle
        `M0,${n} L${t},${n} L${e},${o} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartPunchedTape", (t: number, n: number) => {
    const e = t / 20, o = n / 20, s = (x: number, p: number, $: number, g: number, f: number, y: number) => {
        const m = f / 6e4, b = y / 6e4, M = m * Math.PI / 180, L = (m + b) * Math.PI / 180, v = x - $ * Math.cos(M), k = p - g * Math.sin(M), A = v + $ * Math.cos(L), S = k + g * Math.sin(L), w = Math.abs(b) > 180 ? 1 : 0, F = b > 0 ? 1 : 0;
        return { endX: A, endY: S, svg: `A${$},${g} 0 ${w},${F} ${A},${S}` };
    }, i = 5 * e, r = 2 * o;
    let c = 0, l = 2 * o;
    const a = [`M${c},${l}`];
    let d = s(c, l, i, r, 108e5, -108e5);
    a.push(d.svg), c = d.endX, l = d.endY, d = s(c, l, i, r, 108e5, 108e5), a.push(d.svg), c = d.endX, l = d.endY;
    const h = 20 * e, u = 18 * o;
    return a.push(`L${h},${u}`), c = h, l = u, d = s(c, l, i, r, 0, -108e5), a.push(d.svg), c = d.endX, l = d.endY, d = s(c, l, i, r, 0, 108e5), a.push(d.svg), a.push("Z"), a.join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartPunchedCard", (t: number, n: number) => {
    const e = t / 5;
    return `M0,${n / 5} L${e},0 L${t},0 L${t},${n} L0,${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("flowChartSummingJunction", (t: number, n: number) => {
    const e = t / 2, o = n / 2, s = e * Math.cos(Math.PI / 4), i = o * Math.sin(Math.PI / 4), r = e - s, c = e + s, l = o - i, a = o + i;
    return [
        // Circle
        `M0,${o}`,
        `A${e},${o} 0 1,1 ${t},${o}`,
        `A${e},${o} 0 1,1 0,${o}`,
        "Z",
        // X cross
        `M${r},${l} L${c},${a}`,
        `M${c},${l} L${r},${a}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartOr", (t: number, n: number) => {
    const e = t / 2, o = n / 2;
    return [
        // Circle
        `M0,${o}`,
        `A${e},${o} 0 1,1 ${t},${o}`,
        `A${e},${o} 0 1,1 0,${o}`,
        "Z",
        // + cross
        `M${e},0 L${e},${n}`,
        `M0,${o} L${t},${o}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartOnlineStorage", (t: number, n: number) => {
    const e = t / 6;
    return [
        `M${e},0`,
        `L${t},0`,
        `A${e},${n / 2} 0 0,0 ${t},${n}`,
        `L${e},${n}`,
        `A${e},${n / 2} 0 0,1 ${e},0`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartMagneticDrum", (t: number, n: number) => {
    const e = t / 6, o = t * 5 / 6, s = n / 2;
    return [
        // Body
        `M${e},0`,
        `L${o},0`,
        `A${e},${s} 0 0,1 ${o},${n}`,
        `L${e},${n}`,
        `A${e},${s} 0 0,1 ${e},0`,
        "Z",
        // Right ellipse back-face (visible part)
        `M${o},${n}`,
        `A${e},${s} 0 0,1 ${o},0`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartMagneticTape", (t: number, n: number) => {
    const e = t / 2, o = n / 2, s = e, i = o, r = Math.atan2(n, t), c = i + o * Math.sin(Math.PI / 4), l = (g: number, f: number, y: number, m: number, b: number, M: number) => {
        const L = b * Math.PI / 180, v = (b + M) * Math.PI / 180, k = g - y * Math.cos(L), A = f - m * Math.sin(L), S = k + y * Math.cos(v), w = A + m * Math.sin(v), F = Math.abs(M) > 180 ? 1 : 0, C = M > 0 ? 1 : 0;
        return { endX: S, endY: w, svg: `A${y},${m} 0 ${F},${C} ${S},${w}` };
    };
    let a = s, d = n;
    const h = l(a, d, e, o, 90, 90);
    a = h.endX, d = h.endY;
    const u = l(a, d, e, o, 180, 90);
    a = u.endX, d = u.endY;
    const x = l(a, d, e, o, 270, 90);
    a = x.endX, d = x.endY;
    const p = r * 180 / Math.PI, $ = l(a, d, e, o, 0, p);
    return [`M${s},${n}`, h.svg, u.svg, x.svg, $.svg, `L${t},${c}`, `L${t},${n}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("flowChartMultidocument", (t: number, n: number) => {
    const e = (s: number) => t * s / 21600, o = (s: number) => n * s / 21600;
    return [
        // Front doc (bottom layer, with wave)
        `M0,${o(20782)}`,
        `C${e(9298)},${o(23542)} ${e(9298)},${o(18022)} ${e(18595)},${o(18022)}`,
        `L${e(18595)},${o(3675)} L0,${o(3675)} Z`,
        // Middle doc
        `M${e(1532)},${o(3675)} L${e(1532)},${o(1815)} L${e(2e4)},${o(1815)}`,
        `L${e(2e4)},${o(16252)}`,
        `C${e(19298)},${o(16252)} ${e(18595)},${o(16352)} ${e(18595)},${o(16352)}`,
        `L${e(18595)},${o(3675)} Z`,
        // Back doc (top layer)
        `M${e(2972)},${o(1815)} L${e(2972)},0 L${t},0`,
        `L${t},${o(14392)}`,
        `C${e(20800)},${o(14392)} ${e(2e4)},${o(14467)} ${e(2e4)},${o(14467)}`,
        `L${e(2e4)},${o(1815)} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("wedgeRectCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = t / 2, s = n / 2, i = t * ((e == null ? void 0 : e.get("adj1")) ?? -20833) / 1e5, r = n * ((e == null ? void 0 : e.get("adj2")) ?? 62500) / 1e5, c = o + i, l = s + r, a = i * n / t, d = Math.abs(r), h = Math.abs(a), u = d - h, x = t * (i >= 0 ? 7 : 2) / 12, p = t * (i >= 0 ? 10 : 5) / 12, $ = n * (r >= 0 ? 7 : 2) / 12, g = n * (r >= 0 ? 10 : 5) / 12, f = u > 0 || i >= 0 ? 0 : c, y = u > 0 ? r >= 0 ? x : c : x, m = u > 0 ? t : i >= 0 ? c : t, b = u > 0 && r >= 0 ? c : x, M = u > 0 || i >= 0 ? $ : l, L = u > 0 ? r >= 0 ? 0 : l : 0, v = u > 0 ? $ : i >= 0 ? l : $, k = u > 0 && r >= 0 ? l : n;
    return [
        "M0,0",
        `L${x},0`,
        `L${y},${L}`,
        `L${p},0`,
        `L${t},0`,
        `L${t},${$}`,
        `L${m},${v}`,
        `L${t},${g}`,
        `L${t},${n}`,
        `L${p},${n}`,
        `L${b},${k}`,
        `L${x},${n}`,
        `L0,${n}`,
        `L0,${g}`,
        `L${f},${M}`,
        `L0,${$}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("wedgeRoundRectCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = t / 2, s = n / 2, i = Math.min(t, n), r = t * ((e == null ? void 0 : e.get("adj1")) ?? -20833) / 1e5, c = n * ((e == null ? void 0 : e.get("adj2")) ?? 62500) / 1e5, l = i * ((e == null ? void 0 : e.get("adj3")) ?? 16667) / 1e5, a = o + r, d = s + c, h = r * n / t, u = Math.abs(c), x = Math.abs(h), p = u - x, $ = t - l, g = n - l, f = t * (r >= 0 ? 7 : 2) / 12, y = t * (r >= 0 ? 10 : 5) / 12, m = n * (c >= 0 ? 7 : 2) / 12, b = n * (c >= 0 ? 10 : 5) / 12, M = p > 0 || r >= 0 ? 0 : a, L = p > 0 ? c >= 0 ? f : a : f, v = p > 0 ? t : r >= 0 ? a : t, k = p > 0 && c >= 0 ? a : f, A = p > 0 || r >= 0 ? m : d, S = p > 0 ? c >= 0 ? 0 : d : 0, w = p > 0 ? m : r >= 0 ? d : m, F = p > 0 && c >= 0 ? d : n;
    return [
        `M0,${l}`,
        `A${l},${l} 0 0,1 ${l},0`,
        `L${f},0`,
        `L${L},${S}`,
        `L${y},0`,
        `L${$},0`,
        `A${l},${l} 0 0,1 ${t},${l}`,
        `L${t},${m}`,
        `L${v},${w}`,
        `L${t},${b}`,
        `L${t},${g}`,
        `A${l},${l} 0 0,1 ${$},${n}`,
        `L${y},${n}`,
        `L${k},${F}`,
        `L${f},${n}`,
        `L${l},${n}`,
        `A${l},${l} 0 0,1 0,${g}`,
        `L0,${b}`,
        `L${M},${A}`,
        `L0,${m}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("wedgeEllipseCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", -20833), s = adjustmentFraction(e, "adj2", 62500), i = t / 2, r = n / 2, c = i + t * o, l = r + n * s, a = Math.atan2(l - r, c - i), d = 0.15;
    return [
        ellipseArcPath(i, r, i, r, (a + d) * 180 / Math.PI, (a - d + 2 * Math.PI) * 180 / Math.PI),
        `L${c},${l}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("cloudCallout", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", -20833), s = adjustmentFraction(e, "adj2", 62500), i = t / 2 + t * o, r = n / 2 + n * s, c = presetBuilder("cloud")(t, n, e), l = t / 2, a = n / 2, d = i - l, h = r - a, u = Math.min(t, n) * 0.04, x = Math.min(t, n) * 0.025, p = l + d * 0.5, $ = a + h * 0.5, g = l + d * 0.75, f = a + h * 0.75;
    return [
        c,
        // Connector circles (approximated as small ellipses)
        `M${p + u},${$} A${u},${u} 0 1,1 ${p - u},${$} A${u},${u} 0 1,1 ${p + u},${$} Z`,
        `M${g + x},${f} A${x},${x} 0 1,1 ${g - x},${f} A${x},${x} 0 1,1 ${g + x},${f} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("borderCallout1", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = n * ((e == null ? void 0 : e.get("adj1")) ?? 18750) / 1e5, s = t * ((e == null ? void 0 : e.get("adj2")) ?? -8333) / 1e5, i = n * ((e == null ? void 0 : e.get("adj3")) ?? 112500) / 1e5, r = t * ((e == null ? void 0 : e.get("adj4")) ?? -38333) / 1e5;
    return `M0,0 L${t},0 L${t},${n} L0,${n} Z M${s},${o} L${r},${i}`;
});
PRESET_SHAPE_BUILDERS.set("cube", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 25e3), s = Math.min(t, n) * o;
    return [
        // Front face
        `M0,${s} L${t - s},${s} L${t - s},${n} L0,${n} Z`,
        // Top face
        `M0,${s} L${s},0 L${t},0 L${t - s},${s} Z`,
        // Right face
        `M${t - s},${s} L${t},0 L${t},${n - s} L${t - s},${n} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("plus", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max(adjustmentValue(e, "adj", 25e3), 0), 5e4), i = o * s / 1e5, r = t - i, c = n - i;
    return [
        `M0,${i}`,
        `L${i},${i}`,
        `L${i},0`,
        `L${r},0`,
        `L${r},${i}`,
        `L${t},${i}`,
        `L${t},${c}`,
        `L${r},${c}`,
        `L${r},${n}`,
        `L${i},${n}`,
        `L${i},${c}`,
        `L0,${c}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("heart", (t: number, n: number) => {
    const e = t / 2, o = n / 4, s = n / 3, i = t * 49 / 48, r = t * 10 / 48, c = e - i, l = e - r, a = e + r, d = e + i, h = -s;
    return [
        `M${e},${o}`,
        `C${a},${h} ${d},${o} ${e},${n}`,
        `C${c},${o} ${l},${h} ${e},${o}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("cloud", (t: number, n: number) => {
    const e = t / 43200, o = n / 43200, s = [
        [6753, 9190, -11429249, 7426832],
        [5333, 7267, -8646143, 5396714],
        [4365, 5945, -8748475, 5983381],
        [4857, 6595, -7859164, 7034504],
        [5333, 7273, -4722533, 6541615],
        [6775, 9220, -2776035, 7816140],
        [5785, 7867, 37501, 6842e3],
        [6752, 9215, 1347096, 6910353],
        [7720, 10543, 3974558, 4542661],
        [4360, 5918, -16496525, 8804134],
        [4345, 5945, -14809710, 9151131]
    ];
    let i = 3900 * e, r = 14370 * o;
    const c = [`M${i},${r}`];
    let l = 3900, a = 14370;
    for (const [d, h, u, x] of s) {
        const p = u / 6e4, $ = x / 6e4, g = p * Math.PI / 180, f = Math.atan2(d * Math.sin(g), h * Math.cos(g)), y = (p + $) * Math.PI / 180, m = Math.atan2(d * Math.sin(y), h * Math.cos(y)), b = l - d * Math.cos(f), M = a - h * Math.sin(f), L = b + d * Math.cos(m), v = M + h * Math.sin(m), k = L * e, A = v * o, S = d * e, w = h * o, F = Math.abs($) > 180 ? 1 : 0, C = $ > 0 ? 1 : 0;
        c.push(`A${S},${w} 0 ${F},${C} ${k},${A}`), l = L, a = v, i = k, r = A;
    }
    return c.push("Z"), c.join(" ");
});
PRESET_SHAPE_BUILDERS.set("frame", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 12500), s = Math.min(t, n) * o;
    return [
        // Outer rectangle
        `M0,0 L${t},0 L${t},${n} L0,${n} Z`,
        // Inner rectangle (counter-clockwise for hole)
        `M${s},${s} L${s},${n - s} L${t - s},${n - s} L${t - s},${s} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("halfFrame", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 33333, s = (e == null ? void 0 : e.get("adj2")) ?? 33333, i = Math.min(t, n), r = Math.max(0, Math.min(s, 1e5 * t / Math.max(i, 1))), c = i * r / 1e5, l = n * c / Math.max(t, 1), a = n - l, d = Math.max(0, Math.min(o, 1e5 * a / Math.max(i, 1))), h = i * d / 1e5, u = t - h * t / Math.max(n, 1), x = n - c * n / Math.max(t, 1);
    return ["M0,0", `L${t},0`, `L${u},${h}`, `L${c},${h}`, `L${c},${x}`, `L0,${n}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("donut", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max(adjustmentValue(e, "adj", 25e3), 0), 5e4), i = o * s / 1e5, r = t / 2, c = n / 2, l = Math.max(0, r - i), a = Math.max(0, c - i);
    return [
        // Outer circle (CW)
        `M0,${c}`,
        `A${r},${c} 0 1,1 ${t},${c}`,
        `A${r},${c} 0 1,1 0,${c}`,
        "Z",
        // Inner circle (CCW for evenodd hole)
        `M${i},${c}`,
        `A${l},${a} 0 1,0 ${t - i},${c}`,
        `A${l},${a} 0 1,0 ${i},${c}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("noSmoking", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max(adjustmentValue(e, "adj", 18750), 0), 5e4), i = o * s / 1e5, r = t / 2, c = n / 2, l = t / 2, a = n / 2, d = r - i, h = c - i, u = Math.atan2(n, t), x = h * Math.cos(u), p = d * Math.sin(u), $ = Math.sqrt(x * x + p * p) || 1, g = d * h / $, f = i / 2, y = Math.atan2(f, g), m = y * 2, b = -(Math.PI - m), M = u - y, L = M - Math.PI, v = (B: number) => {
        const R = h * Math.cos(B), I = d * Math.sin(B), Z = Math.sqrt(R * R + I * I) || 1, U = d * h / Z;
        return { x: l + U * Math.cos(B), y: a + U * Math.sin(B) };
    }, k = v(M), A = v(L), S = M + b, w = L + b, F = v(S), C = v(w), E = Math.abs(b) > Math.PI ? 1 : 0, P = b > 0 ? 1 : 0;
    return [
        // Outer circle (CW)
        `M0,${a}`,
        `A${r},${c} 0 1,1 ${t},${a}`,
        `A${r},${c} 0 1,1 0,${a}`,
        "Z",
        // First diagonal band arc (inner ellipse)
        `M${k.x},${k.y}`,
        `A${d},${h} 0 ${E},${P} ${F.x},${F.y}`,
        "Z",
        // Second diagonal band arc (opposite quadrant)
        `M${A.x},${A.y}`,
        `A${d},${h} 0 ${E},${P} ${C.x},${C.y}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("blockArc", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 108e5, s = (e == null ? void 0 : e.get("adj2")) ?? 0, i = (e == null ? void 0 : e.get("adj3")) ?? 25e3, r = Math.min(Math.max(o / 6e4, 0), 360), c = Math.min(Math.max(s / 6e4, 0), 360), l = (c - r + 360) % 360 || 360, a = r + l, d = c - l, h = t / 2, u = n / 2, x = Math.min(t, n) * Math.max(0, Math.min(i, 5e4)) / 1e5, p = Math.max(1, h - x), $ = Math.max(1, u - x), g = (L: number, v: number, k: number, A: number, S: number) => {
        const w = S * Math.PI / 180;
        return { x: L + k * Math.cos(w), y: v + A * Math.sin(w) };
    }, f = g(h, u, h, u, r), y = g(h, u, h, u, a), m = g(h, u, p, $, c), b = g(h, u, p, $, d), M = l > 180 ? 1 : 0;
    return [
        `M${f.x},${f.y}`,
        `A${h},${u} 0 ${M},1 ${y.x},${y.y}`,
        `L${m.x},${m.y}`,
        `A${p},${$} 0 ${M},0 ${b.x},${b.y}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("gear6", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 15e3, s = (e == null ? void 0 : e.get("adj2")) ?? 3526;
    return gearPath(t, n, 6, o, s);
});
PRESET_SHAPE_BUILDERS.set("gear9", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 1e4, s = (e == null ? void 0 : e.get("adj2")) ?? 1763;
    return gearPath(t, n, 9, o, s);
});
function gearPath(
  width: number,
  height: number,
  teeth: number,
  radiusAdjustment: number,
  toothAdjustment: number
): string {
  const t = width, n = height, e = teeth, o = radiusAdjustment, s = toothAdjustment;
    const i = t / 2, r = n / 2, c = Math.min(t, n), l = e === 6 ? 5358 : 2679, a = Math.min(Math.max(o, 0), 2e4), d = Math.min(Math.max(s, 0), l), h = c * a / 1e5, u = c * d / 1e5, x = t / 2 - h, p = n / 2 - h;
    if (x <= 0 || p <= 0)
        return `M0,0 L${t},0 L${t},${n} L0,${n} Z`;
    const $ = h / 2 + u / 2, g = Math.min(x, p), f = Math.atan2($, g), y = e === 6 ? [330, 30, 90, 150, 210, 270] : [310, 350, 30, 70, 110, 150, 190, 230, 270], m = [];
    for (let b = 0; b < y.length; b++) {
        const M = y[b] * Math.PI / 180, L = M - f, v = M + f, k = i + x * Math.cos(L), A = r + p * Math.sin(L), S = i + x * Math.cos(v), w = r + p * Math.sin(v), F = S - k, C = w - A, E = Math.sqrt(F * F + C * C), P = Math.cos(M), B = Math.sin(M);
        let R = P, I = B;
        E > 0 && (R = -C / E, I = F / E), R * P + I * B < 0 && (R = -R, I = -I);
        const Z = E > 0 ? F / E : 0, U = E > 0 ? C / E : 0, q = k + Z * u, Q = A + U * u, G = S - Z * u, T = w - U * u, D = q + R * h, O = Q + I * h, ot = G + R * h, J = T + I * h;
        if (b === 0) {
            const W = y[y.length - 1] * Math.PI / 180 + f, it = i + x * Math.cos(W), lt = r + p * Math.sin(W);
            m.push(`M${it},${lt}`), m.push(`A${x},${p} 0 0,1 ${k},${A}`);
        }
        if (m.push(`L${D},${O}`), m.push(`L${ot},${J}`), m.push(`L${S},${w}`), b < y.length - 1) {
            const W = y[b + 1] * Math.PI / 180 - f, it = i + x * Math.cos(W), lt = r + p * Math.sin(W);
            m.push(`A${x},${p} 0 0,1 ${it},${lt}`);
        }
    }
    return m.push("Z"), m.join(" ");
}
PRESET_SHAPE_BUILDERS.set("mathPlus", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max(adjustmentValue(e, "adj", 23520), 0), 73490), i = t * 73490 / 2e5, r = n * 73490 / 2e5, c = o * s / 2e5, l = t / 2, a = n / 2, d = l - i, h = l - c, u = l + c, x = l + i, p = a - r, $ = a - c, g = a + c, f = a + r;
    return [
        `M${d},${$}`,
        `L${h},${$}`,
        `L${h},${p}`,
        `L${u},${p}`,
        `L${u},${$}`,
        `L${x},${$}`,
        `L${x},${g}`,
        `L${u},${g}`,
        `L${u},${f}`,
        `L${h},${f}`,
        `L${h},${g}`,
        `L${d},${g}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("mathMinus", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(Math.max(adjustmentValue(e, "adj1", 23520), 0), 1e5), s = n * o / 2e5, i = t * 73490 / 2e5, r = t / 2, c = n / 2, l = r - i, a = r + i, d = c - s, h = c + s;
    return `M${l},${d} L${a},${d} L${a},${h} L${l},${h} Z`;
});
PRESET_SHAPE_BUILDERS.set("mathMultiply", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = t / 2, i = n / 2, r = Math.min(Math.max(adjustmentValue(e, "adj1", 23520), 0), 51965), c = o * r / 1e5, l = Math.atan2(n, t), a = Math.sin(l), d = Math.cos(l), h = a / d, u = Math.sqrt(t * t + n * n), x = u * 51965 / 1e5, p = u - x, $ = d * p / 2, g = a * p / 2, f = a * c / 2, y = d * c / 2, m = $ - f, b = g + y, M = $ + f, L = g - y, A = (s - M) * h + L, S = t - M, w = t - m, C = (i - b) / h, E = w - C, P = m + C, B = n - b, R = n - L, I = n - A;
    return [
        `M${m},${b}`,
        `L${M},${L}`,
        `L${s},${A}`,
        `L${S},${L}`,
        `L${w},${b}`,
        `L${E},${i}`,
        `L${w},${B}`,
        `L${S},${R}`,
        `L${s},${I}`,
        `L${M},${R}`,
        `L${m},${B}`,
        `L${P},${i}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("mathDivide", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 23520, s = (e == null ? void 0 : e.get("adj2")) ?? 5880, i = (e == null ? void 0 : e.get("adj3")) ?? 11760, r = Math.min(Math.max(o, 1e3), 36745), c = Math.min((73490 - r) / 4, 36745 * t / Math.max(n, 1)), l = Math.min(Math.max(i, 1e3), c), a = 73490 - 4 * l - r, d = Math.min(Math.max(s, 0), a), h = t / 2, u = n / 2, x = n * r / 2e5, p = n * d / 1e5, $ = n * l / 1e5, g = t * 73490 / 2e5, f = u - x, y = u + x, b = f - (p + $) - $, M = n - b, L = h - g, v = h + g;
    return [
        // Top dot
        `M${h + $},${b + $} A${$},${$} 0 1,1 ${h - $},${b + $} A${$},${$} 0 1,1 ${h + $},${b + $} Z`,
        // Bottom dot
        `M${h + $},${M - $} A${$},${$} 0 1,1 ${h - $},${M - $} A${$},${$} 0 1,1 ${h + $},${M - $} Z`,
        // Bar
        `M${L},${f} L${v},${f} L${v},${y} L${L},${y} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("mathEqual", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 23520, s = (e == null ? void 0 : e.get("adj2")) ?? 11760, i = Math.min(Math.max(o, 0), 36745), r = 1e5 - i * 2, c = Math.min(Math.max(s, 0), Math.max(r, 0)), l = n * i / 1e5, a = n * c / 2e5, d = t * 73490 / 2e5, h = t / 2, u = n / 2, x = u - a, p = u + a, $ = x - l, g = p + l, f = h - d, y = h + d;
    return [
        `M${f},${$} L${y},${$} L${y},${x} L${f},${x} Z`,
        `M${f},${p} L${y},${p} L${y},${g} L${f},${g} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("mathNotEqual", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 23520, s = e == null ? void 0 : e.get("adj2"), i = (e == null ? void 0 : e.get("adj3")) ?? 11760, r = t / 2, c = n / 2, l = n / 2, a = Math.min(Math.max(o, 0), 5e4), d = (() => {
        if (s === void 0)
            return 110 * Math.PI / 180;
        const lt = s / 6e4 * Math.PI / 180, Pt = 70 * Math.PI / 180, Rt = 110 * Math.PI / 180;
        return Math.min(Math.max(lt, Pt), Rt);
    })(), h = 1e5 - a * 2, u = Math.min(Math.max(i, 0), h), x = n * a / 1e5, p = n * u / 2e5, $ = t * 73490 / 2e5, g = r - $, f = r + $, y = c - p, m = c + p, b = y - x, M = m + x, L = d - Math.PI / 2, v = l * Math.tan(L), k = Math.hypot(v, l) || 1, A = k * x / l, S = A / 2, w = r + v - S, F = w - v * b / l, C = w - v * y / l, E = w - v * m / l, P = w - v * M / l, B = w + A, R = F + A, I = C + A, Z = E + A, U = P + A, q = x * l / k, Q = L > 0 ? w + q : B, G = L > 0 ? w : B - q, T = x * v / k, D = L > 0 ? T : 0, O = L > 0 ? 0 : -T, ot = t - Q, J = t - G, W = n - D, it = n - O;
    return [
        `M${g},${b}`,
        `L${F},${b}`,
        `L${G},${O}`,
        `L${Q},${D}`,
        `L${R},${b}`,
        `L${f},${b}`,
        `L${f},${y}`,
        `L${I},${y}`,
        `L${Z},${m}`,
        `L${f},${m}`,
        `L${f},${M}`,
        `L${U},${M}`,
        `L${J},${it}`,
        `L${ot},${W}`,
        `L${P},${M}`,
        `L${g},${M}`,
        `L${g},${m}`,
        `L${E},${m}`,
        `L${C},${y}`,
        `L${g},${y}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("round1Rect", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 16667), s = Math.min(t, n) * o;
    return ["M0,0", `L${t - s},0`, `A${s},${s} 0 0,1 ${t},${s}`, `L${t},${n}`, `L0,${n}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("round2SameRect", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 16667), s = adjustmentFraction(e, "adj2", 0), i = Math.min(t, n) * o, r = Math.min(t, n) * s;
    return [
        `M${i},0`,
        `L${t - i},0`,
        `A${i},${i} 0 0,1 ${t},${i}`,
        `L${t},${n - r}`,
        `A${r},${r} 0 0,1 ${t - r},${n}`,
        `L${r},${n}`,
        `A${r},${r} 0 0,1 0,${n - r}`,
        `L0,${i}`,
        `A${i},${i} 0 0,1 ${i},0`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("round2DiagRect", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 16667), s = adjustmentFraction(e, "adj2", 0), i = Math.min(t, n) * o, r = Math.min(t, n) * s;
    return [
        `M${i},0`,
        `L${t},0`,
        `L${t},${n - r}`,
        `A${r},${r} 0 0,1 ${t - r},${n}`,
        `L0,${n}`,
        `L0,${i}`,
        `A${i},${i} 0 0,1 ${i},0`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("snip1Rect", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 16667), s = Math.min(t, n) * o;
    return `M0,0 L${t - s},0 L${t},${s} L${t},${n} L0,${n} Z`;
});
PRESET_SHAPE_BUILDERS.set("snip2SameRect", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 16667), s = adjustmentFraction(e, "adj2", 0), i = Math.min(t, n) * o, r = Math.min(t, n) * s;
    return `M${i},0 L${t - i},0 L${t},${i} L${t},${n - r} L${t - r},${n} L${r},${n} L0,${n - r} L0,${i} Z`;
});
PRESET_SHAPE_BUILDERS.set("snip2DiagRect", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max((e == null ? void 0 : e.get("adj1")) ?? 0, 0), 5e4), i = Math.min(Math.max((e == null ? void 0 : e.get("adj2")) ?? 16667, 0), 5e4), r = o * s / 1e5, c = t - r, l = n - r, a = o * i / 1e5, d = t - a, h = n - a;
    return `M${r},0 L${d},0 L${t},${a} L${t},${l} L${c},${n} L${a},${n} L0,${h} L0,${r} Z`;
});
PRESET_SHAPE_BUILDERS.set("snipRoundRect", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj1", 16667), s = adjustmentFraction(e, "adj2", 16667), i = Math.min(t, n) * o, r = Math.min(t, n) * s;
    return [
        `M${i},0`,
        `L${t - r},0`,
        `L${t},${r}`,
        `L${t},${n}`,
        `L0,${n}`,
        `L0,${i}`,
        `A${i},${i} 0 0,1 ${i},0`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("bevel", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 12500), s = Math.min(t, n) * o;
    return [
        // Outer
        `M0,0 L${t},0 L${t},${n} L0,${n} Z`,
        // Inner
        `M${s},${s} L${s},${n - s} L${t - s},${n - s} L${t - s},${s} Z`,
        // Connecting triangles (top)
        `M0,0 L${t},0 L${t - s},${s} L${s},${s} Z`,
        // Right
        `M${t},0 L${t},${n} L${t - s},${n - s} L${t - s},${s} Z`,
        // Bottom
        `M${t},${n} L0,${n} L${s},${n - s} L${t - s},${n - s} Z`,
        // Left
        `M0,${n} L0,0 L${s},${s} L${s},${n - s} Z`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("foldedCorner", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 16667), s = Math.min(t, n) * o * 0.7;
    return [
        `M0,0 L${t},0 L${t},${n} L0,${n} Z`,
        // Fold triangle
        `M${t - s},${n} L${t},${n} L${t},${n - s}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("sun", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj")) ?? 25e3, s = Math.min(Math.max(o, 12500), 46875), i = 5e4 - s, r = i * 30274 / 32768, c = i * 12540 / 32768, l = 5e4 - r, a = 5e4 - c, d = l * 3 / 4, h = a * 3 / 4, u = d + 3662, x = h + 3662, p = h + 12500, $ = 1e5 - d, g = 1e5 - u, f = 1e5 - x, y = 1e5 - p, m = t / 2, b = n / 2, M = t * 18436 / 21600, L = n * 3163 / 21600, v = t * 3163 / 21600, k = n * 18436 / 21600, A = (W: number, it: number) => it * W / 1e5, S = A(d, t), w = A(u, t), F = A(x, t), C = A(p, t), E = A($, t), P = A(g, t), B = A(f, t), R = A(y, t), I = A(i, t), Z = A(i, n), U = A(d, n), q = A(u, n), Q = A(x, n), G = A(p, n), T = A($, n), D = A(g, n), O = A(f, n), ot = A(y, n), J = A(s, t);
    return [
        // Ray 0: right
        `M${t},${b} L${E},${ot} L${E},${G} Z`,
        // Ray 1: top-right
        `M${M},${L} L${P},${Q} L${B},${q} Z`,
        // Ray 2: top
        `M${m},0 L${R},${U} L${C},${U} Z`,
        // Ray 3: top-left
        `M${v},${L} L${F},${q} L${w},${Q} Z`,
        // Ray 4: left
        `M0,${b} L${S},${G} L${S},${ot} Z`,
        // Ray 5: bottom-left
        `M${v},${k} L${w},${O} L${F},${D} Z`,
        // Ray 6: bottom
        `M${m},${n} L${C},${T} L${R},${T} Z`,
        // Ray 7: bottom-right
        `M${M},${k} L${B},${D} L${P},${O} Z`,
        // Center ellipse (arcTo from x19,vc with wR,hR, startAngle=180°, sweep=360°)
        `M${J},${b}`,
        `A${I},${Z} 0 1,1 ${J + 2 * I},${b}`,
        `A${I},${Z} 0 1,1 ${J},${b}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("moon", (t: number, n: number, e: PresetShapeAdjustments) => {
    if (t <= 0 || n <= 0)
        return `M0,0 L${t},0 L${t},${n} L0,${n} Z`;
    const o = Math.min(t, n), s = n / 2, i = Math.min(Math.max((e == null ? void 0 : e.get("adj")) ?? 5e4, 0), 87500), r = o * i / 1e5, c = o - r;
    if (c <= 0)
        return `M0,0 L${t},0 L${t},${n} L0,${n} Z`;
    const l = r * t / o, a = (2 * o * o - r * r) / c, d = (a - r) * t / o, u = (a / 2 - r) * s / o, x = (d - l) / 2;
    return [
        `M${t},${n}`,
        `A${t},${s} 0 0,1 ${t},0`,
        // outer: (w,h) → left semicircle → (w,0)
        `A${x},${u} 0 0,0 ${t},${n}`,
        // inner: (w,0) → concave arc → (w,h)
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("lightningBolt", (t: number, n: number) => [
    `M${t * 0.3895},${n * 0}`,
    `L${t * 0},${n * 0.1821}`,
    `L${t * 0.3425},${n * 0.3845}`,
    `L${t * 0.2265},${n * 0.4452}`,
    `L${t * 0.5497},${n * 0.6391}`,
    `L${t * 0.453},${n * 0.683}`,
    `L${t * 0.9972},${n * 0.9983}`,
    `L${t * 0.6796},${n * 0.5919}`,
    `L${t * 0.7624},${n * 0.5514}`,
    `L${t * 0.5138},${n * 0.3153}`,
    `L${t * 0.5939},${n * 0.2816}`,
    "Z"
].join(" "));
PRESET_SHAPE_BUILDERS.set("bracketPair", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.min(Math.max(adjustmentValue(e, "adj", 16667), 0), 5e4), i = o * s / 1e5, r = t - i, c = n - i;
    return [
        // Left bracket: bottom-left arc → vertical → top-left arc
        `M${i},${n}`,
        `A${i},${i} 0 0,1 0,${c}`,
        `L0,${i}`,
        `A${i},${i} 0 0,1 ${i},0`,
        // Right bracket: top-right arc → vertical → bottom-right arc
        `M${r},0`,
        `A${i},${i} 0 0,1 ${t},${i}`,
        `L${t},${c}`,
        `A${i},${i} 0 0,1 ${r},${n}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("bracePair", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = adjustmentFraction(e, "adj", 8333), s = Math.min(t, n) * o, i = n / 2;
    return [
        // Left brace
        `M${s * 2},0`,
        `A${s},${s} 0 0,0 ${s},${s}`,
        `L${s},${i - s}`,
        `A${s},${s} 0 0,1 0,${i}`,
        `A${s},${s} 0 0,1 ${s},${i + s}`,
        `L${s},${n - s}`,
        `A${s},${s} 0 0,0 ${s * 2},${n}`,
        // Right brace
        `M${t - s * 2},0`,
        `A${s},${s} 0 0,1 ${t - s},${s}`,
        `L${t - s},${i - s}`,
        `A${s},${s} 0 0,0 ${t},${i}`,
        `A${s},${s} 0 0,0 ${t - s},${i + s}`,
        `L${t - s},${n - s}`,
        `A${s},${s} 0 0,1 ${t - s * 2},${n}`
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("leftBracket", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = o > 0 ? 5e4 * n / o : 0, i = Math.max(0, Math.min((e == null ? void 0 : e.get("adj")) ?? 8333, s)), r = o * i / 1e5, c = (h: number) => h / 6e4, l = (h: number, u: number, x: number, p: number, $: number, g: number) => {
        const f = c($) * Math.PI / 180, y = c(g) * Math.PI / 180, m = h - x * Math.cos(f), b = u - p * Math.sin(f), M = m + x * Math.cos(f + y), L = b + p * Math.sin(f + y), v = Math.abs(c(g)) > 180 ? 1 : 0;
        return { cmd: `A${x},${p} 0 ${v},1 ${M},${L}`, x: M, y: L };
    }, a = l(t, n, t, r, 54e5, 54e5), d = l(0, r, t, r, 108e5, 54e5);
    return [`M${t},${n}`, a.cmd, `L0,${r}`, d.cmd].join(" ");
});
PRESET_SHAPE_BUILDERS.set("rightBracket", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = o > 0 ? 5e4 * n / o : 0, i = Math.max(0, Math.min((e == null ? void 0 : e.get("adj")) ?? 8333, s)), r = o * i / 1e5, c = n - r, l = (u: number) => u / 6e4, a = (u: number, x: number, p: number, $: number, g: number, f: number) => {
        const y = l(g) * Math.PI / 180, m = l(f) * Math.PI / 180, b = u - p * Math.cos(y), M = x - $ * Math.sin(y), L = b + p * Math.cos(y + m), v = M + $ * Math.sin(y + m), k = Math.abs(l(f)) > 180 ? 1 : 0;
        return { cmd: `A${p},${$} 0 ${k},1 ${L},${v}`, x: L, y: v };
    }, d = a(0, 0, t, r, 162e5, 54e5), h = a(t, c, t, r, 0, 54e5);
    return ["M0,0", d.cmd, `L${t},${c}`, h.cmd].join(" ");
});
PRESET_SHAPE_BUILDERS.set("leftBrace", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 5e4, 1e5)), i = 1e5 - s, c = Math.min(i, s) / 2, l = o > 0 ? c * n / o : 0, a = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 8333, l)), d = o * a / 1e5, u = n * s / 1e5 + d, x = t / 2, p = t / 2, $ = (v: number) => v / 6e4, g = (v: number, k: number, A: number, S: number, w: number, F: number) => {
        const C = $(w) * Math.PI / 180, E = $(F) * Math.PI / 180, P = v - A * Math.cos(C), B = k - S * Math.sin(C), R = P + A * Math.cos(C + E), I = B + S * Math.sin(C + E), Z = Math.abs($(F)) > 180 ? 1 : 0, U = F >= 0 ? 1 : 0;
        return { cmd: `A${A},${S} 0 ${Z},${U} ${R},${I}`, x: R, y: I };
    };
    let f = t, y = n;
    const m = g(f, y, x, d, 54e5, 54e5);
    f = m.x, y = m.y;
    const b = g(p, u, x, d, 0, -54e5), M = g(b.x, b.y, x, d, 54e5, -54e5), L = g(p, d, x, d, 108e5, 54e5);
    return [
        `M${t},${n}`,
        m.cmd,
        `L${p},${u}`,
        b.cmd,
        M.cmd,
        `L${p},${d}`,
        L.cmd
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("rightBrace", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(t, n), s = Math.max(0, Math.min((e == null ? void 0 : e.get("adj2")) ?? 5e4, 1e5)), i = 1e5 - s, c = Math.min(i, s) / 2, l = o > 0 ? c * n / o : 0, a = Math.max(0, Math.min((e == null ? void 0 : e.get("adj1")) ?? 8333, l)), d = o * a / 1e5, u = n * s / 1e5 - d, x = n - d, p = t / 2, $ = t / 2, g = (L: number) => L / 6e4, f = (L: number, v: number, k: number, A: number, S: number, w: number) => {
        const F = g(S) * Math.PI / 180, C = g(w) * Math.PI / 180, E = L - k * Math.cos(F), P = v - A * Math.sin(F), B = E + k * Math.cos(F + C), R = P + A * Math.sin(F + C), I = Math.abs(g(w)) > 180 ? 1 : 0, Z = w >= 0 ? 1 : 0;
        return { cmd: `A${k},${A} 0 ${I},${Z} ${B},${R}`, x: B, y: R };
    }, y = f(0, 0, p, d, 162e5, 54e5), m = f($, u, p, d, 108e5, -54e5), b = f(m.x, m.y, p, d, 162e5, -54e5), M = f($, x, p, d, 0, 54e5);
    return ["M0,0", y.cmd, `L${$},${u}`, m.cmd, b.cmd, `L${$},${x}`, M.cmd].join(" ");
});
PRESET_SHAPE_BUILDERS.set("actionButtonBlank", (t: number, n: number) => `M0,0 L${t},0 L${t},${n} L0,${n} Z`);
PRESET_SHAPE_BUILDERS.set("wave", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(Math.max(adjustmentValue(e, "adj1", 12500), 0), 2e4), s = Math.min(Math.max(adjustmentValue(e, "adj2", 0), -1e4), 1e4), i = n * o / 1e5, r = i * 10 / 3, c = i - r, l = i + r, a = n - i, d = a - r, h = a + r, u = t * s / 5e4, x = u < 0 ? 0 : u, p = u < 0 ? u : 0, $ = -x, g = t - p, f = (g - $) / 3, y = $ + f, m = (y + g) / 2, b = p, M = t + x, L = b + (M - b) / 3, v = (L + M) / 2;
    return [
        `M${$},${i}`,
        `C${y},${c} ${m},${l} ${g},${i}`,
        `L${M},${a}`,
        `C${v},${h} ${L},${d} ${b},${a}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("doubleWave", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = Math.min(Math.max(adjustmentValue(e, "adj1", 6250), 0), 12500), s = Math.min(Math.max(adjustmentValue(e, "adj2", 0), -1e4), 1e4), i = n * o / 1e5, r = i * 10 / 3, c = i - r, l = i + r, a = n - i, d = a - r, h = a + r, u = t * s / 5e4, x = u < 0 ? 0 : u, p = u < 0 ? u : 0, $ = -x, g = t - p, f = (g - $) / 6, y = $ + f, m = (g - $) / 3, b = $ + m, M = ($ + g) / 2, L = M + f, v = (L + g) / 2, k = p, A = t + x, S = (A - k) / 6, w = k + S, F = k + (A - k) / 3, C = (k + A) / 2, E = C + S, P = (E + A) / 2;
    return [
        `M${$},${i}`,
        `C${y},${c} ${b},${l} ${M},${i}`,
        `C${L},${c} ${v},${l} ${g},${i}`,
        `L${A},${a}`,
        `C${P},${h} ${E},${d} ${C},${a}`,
        `C${F},${h} ${w},${d} ${k},${a}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("irregularSeal1", (t: number, n: number) => {
    const e = (s: number) => t * s / 21600, o = (s: number) => n * s / 21600;
    return [
        `M${e(10800)},${o(5800)}`,
        `L${e(14522)},0`,
        `L${e(14155)},${o(5325)}`,
        `L${e(18380)},${o(4457)}`,
        `L${e(16702)},${o(7315)}`,
        `L${e(21097)},${o(8137)}`,
        `L${e(17607)},${o(10475)}`,
        `L${e(21600)},${o(13290)}`,
        `L${e(16837)},${o(12942)}`,
        `L${e(18145)},${o(18095)}`,
        `L${e(14020)},${o(14457)}`,
        `L${e(13247)},${o(19737)}`,
        `L${e(10532)},${o(14935)}`,
        `L${e(8485)},${o(21600)}`,
        `L${e(7715)},${o(15627)}`,
        `L${e(4762)},${o(17617)}`,
        `L${e(5667)},${o(13937)}`,
        `L${e(135)},${o(14587)}`,
        `L${e(3722)},${o(11775)}`,
        `L0,${o(8615)}`,
        `L${e(4627)},${o(7617)}`,
        `L${e(370)},${o(2295)}`,
        `L${e(7312)},${o(6320)}`,
        `L${e(8352)},${o(2295)}`,
        "Z"
    ].join(" ");
});
PRESET_SHAPE_BUILDERS.set("irregularSeal2", (t: number, n: number) => [
    `M${t * 11462 / 21600},${n * 4342 / 21600}`,
    `L${t * 14790 / 21600},0`,
    `L${t * 14525 / 21600},${n * 5777 / 21600}`,
    `L${t * 18007 / 21600},${n * 3172 / 21600}`,
    `L${t * 16380 / 21600},${n * 6532 / 21600}`,
    `L${t},${n * 6645 / 21600}`,
    `L${t * 16985 / 21600},${n * 9402 / 21600}`,
    `L${t * 18270 / 21600},${n * 11290 / 21600}`,
    `L${t * 16380 / 21600},${n * 12310 / 21600}`,
    `L${t * 18877 / 21600},${n * 15632 / 21600}`,
    `L${t * 14640 / 21600},${n * 14350 / 21600}`,
    `L${t * 14942 / 21600},${n * 17370 / 21600}`,
    `L${t * 12180 / 21600},${n * 15935 / 21600}`,
    `L${t * 11612 / 21600},${n * 18842 / 21600}`,
    `L${t * 9872 / 21600},${n * 17370 / 21600}`,
    `L${t * 8700 / 21600},${n * 19712 / 21600}`,
    `L${t * 7527 / 21600},${n * 18125 / 21600}`,
    `L${t * 4917 / 21600},${n}`,
    `L${t * 4805 / 21600},${n * 18240 / 21600}`,
    `L${t * 1285 / 21600},${n * 17825 / 21600}`,
    `L${t * 3330 / 21600},${n * 15370 / 21600}`,
    `L0,${n * 12877 / 21600}`,
    `L${t * 3935 / 21600},${n * 11592 / 21600}`,
    `L${t * 1172 / 21600},${n * 8270 / 21600}`,
    `L${t * 5372 / 21600},${n * 7817 / 21600}`,
    `L${t * 4502 / 21600},${n * 3625 / 21600}`,
    `L${t * 8550 / 21600},${n * 6382 / 21600}`,
    `L${t * 9722 / 21600},${n * 1887 / 21600}`,
    "Z"
].join(" "));
PRESET_SHAPE_BUILDERS.set("teardrop", (t: number, n: number) => {
    const e = t / 2, o = n / 2;
    return [`M${t},${o}`, `A${e},${o} 0 1,1 ${e},0`, `L${t},0`, `L${t},${o}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("pie", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 0, s = (e == null ? void 0 : e.get("adj2")) ?? 162e5, i = o / 6e4 % 360, r = s / 6e4 % 360;
    let c = ((r - i) % 360 + 360) % 360;
    c === 0 && i !== r && (c = 360);
    const l = t / 2, a = n / 2, d = (m: number) => m * Math.PI / 180, h = (m: number) => Math.atan2(Math.sin(d(m)) / a, Math.cos(d(m)) / l), u = h(i), x = h(r), p = l + l * Math.cos(u), $ = a + a * Math.sin(u), g = l + l * Math.cos(x), f = a + a * Math.sin(x), y = c > 180 ? 1 : 0;
    return [`M${l},${a}`, `L${p},${$}`, `A${l},${a} 0 ${y},1 ${g},${f}`, "Z"].join(" ");
});
PRESET_SHAPE_BUILDERS.set("pieWedge", (t: number, n: number) => [`M0,${n}`, `A${t},${n} 0 0,1 ${t},0`, `L${t},${n}`, "Z"].join(" "));
PRESET_SHAPE_BUILDERS.set("arc", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 162e5, s = (e == null ? void 0 : e.get("adj2")) ?? 0, i = o / 6e4, r = s / 6e4, c = t / 2, l = n / 2, a = (m: number) => m * Math.PI / 180, d = (m: number) => Math.atan2(Math.sin(a(m)) / l, Math.cos(a(m)) / c), h = d(i), u = d(r), x = c + c * Math.cos(h), p = l + l * Math.sin(h), $ = c + c * Math.cos(u), g = l + l * Math.sin(u);
    let f = ((r - i) % 360 + 360) % 360;
    f === 0 && i !== r && (f = 360);
    const y = f > 180 ? 1 : 0;
    return `M${x},${p} A${c},${l} 0 ${y},1 ${$},${g}`;
});
PRESET_SHAPE_BUILDERS.set("chord", (t: number, n: number, e: PresetShapeAdjustments) => {
    const o = (e == null ? void 0 : e.get("adj1")) ?? 27e5, s = (e == null ? void 0 : e.get("adj2")) ?? 162e5, i = o / 6e4, r = s / 6e4, c = t / 2, l = n / 2, a = t / 2, d = n / 2, h = (M: number) => M * Math.PI / 180, u = (M: number) => Math.atan2(Math.sin(h(M)) / d, Math.cos(h(M)) / a), x = u(i), p = u(r), $ = c + a * Math.cos(x), g = l + d * Math.sin(x), f = c + a * Math.cos(p), y = l + d * Math.sin(p);
    let m = ((r - i) % 360 + 360) % 360;
    if (m === 0 && i !== r && (m = 360), m === 0)
        return `M${c - a},${l} A${a},${d} 0 1,1 ${c + a},${l} A${a},${d} 0 1,1 ${c - a},${l} Z`;
    const b = m > 180 ? 1 : 0;
    return `M${$},${g} A${a},${d} 0 ${b},1 ${f},${y} Z`;
});
PRESET_SHAPE_BUILDERS.set("funnel", (t: number, n: number) => {
    const e = Math.min(t, n), o = t / 2, s = n / 4, i = t / 2, r = n, c = e / 20, l = o - c, a = s - c, d = 8 * Math.PI / 180, h = o * Math.cos(d), u = s * Math.sin(d), x = Math.atan2(u, h), p = Math.PI - x, $ = Math.PI + 2 * x, g = Math.PI - 2 * x, f = o / 4, y = s / 4, m = s * Math.cos(p), b = o * Math.sin(p), M = Math.sqrt(m * m + b * b), L = o * s / M, v = L * Math.cos(p), k = L * Math.sin(p), A = i + v, S = s + k, w = p + $, F = s * Math.cos(w), C = o * Math.sin(w), E = Math.sqrt(F * F + C * C), P = o * s / E, B = P * Math.cos(w), R = P * Math.sin(w), I = i + B, Z = s + R, U = r - y, q = y * Math.cos(x), Q = f * Math.sin(x), G = Math.sqrt(q * q + Q * Q), T = f * y / G, D = T * Math.cos(x), O = T * Math.sin(x), ot = i + D, J = U + O, W = x + g, it = y * Math.cos(W), lt = f * Math.sin(W), Pt = Math.sqrt(it * it + lt * lt), Rt = f * y / Pt, Wt = Rt * Math.cos(W), Nt = Rt * Math.sin(W), ft = i + Wt, Zt = U + Nt, Dt = $ * 180 / Math.PI, Ot = Math.abs(Dt) > 180 ? 1 : 0, Ut = $ > 0 ? 1 : 0, Y = g * 180 / Math.PI, Bt = Math.abs(Y) > 180 ? 1 : 0, Ft = g > 0 ? 1 : 0, Xt = [
        `M${A},${S}`,
        `A${o},${s} 0 ${Ot},${Ut} ${I},${Z}`,
        `L${ot},${J}`,
        `A${f},${y} 0 ${Bt},${Ft} ${ft},${Zt}`,
        "Z"
    ].join(" "), H = o - l, N = o + l, et = [
        `M${H},${s}`,
        `A${l},${a} 0 1,0 ${N},${s}`,
        `A${l},${a} 0 1,0 ${H},${s}`,
        "Z"
    ].join(" ");
    return `${Xt} ${et}`;
});

/** `borderCallout2`/`callout2`: rectangle plus a two-segment leader line. */
PRESET_SHAPE_BUILDERS.set("borderCallout2", (t: number, n: number, e: PresetShapeAdjustments) => {
  const y1 = n * adjustmentValue(e, "adj1", 18750) / 1e5;
  const x1 = t * adjustmentValue(e, "adj2", -8333) / 1e5;
  const y2 = n * adjustmentValue(e, "adj3", 112500) / 1e5;
  const x2 = t * adjustmentValue(e, "adj4", -38333) / 1e5;
  const y3 = n * adjustmentValue(e, "adj5", 112500) / 1e5;
  const x3 = t * adjustmentValue(e, "adj6", -16667) / 1e5;
  return `M0,0 L${t},0 L${t},${n} L0,${n} Z M${x1},${y1} L${x2},${y2} L${x3},${y3}`;
});

/** `borderCallout3`/`callout3`: rectangle plus a three-segment leader line. */
PRESET_SHAPE_BUILDERS.set("borderCallout3", (t: number, n: number, e: PresetShapeAdjustments) => {
  const y1 = n * adjustmentValue(e, "adj1", 18750) / 1e5;
  const x1 = t * adjustmentValue(e, "adj2", -8333) / 1e5;
  const y2 = n * adjustmentValue(e, "adj3", 112500) / 1e5;
  const x2 = t * adjustmentValue(e, "adj4", -38333) / 1e5;
  const y3 = n * adjustmentValue(e, "adj5", 112500) / 1e5;
  const x3 = t * adjustmentValue(e, "adj6", -16667) / 1e5;
  const y4 = n * adjustmentValue(e, "adj7", 112500) / 1e5;
  const x4 = t * adjustmentValue(e, "adj8", 0) / 1e5;
  return `M0,0 L${t},0 L${t},${n} L0,${n} Z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x4},${y4}`;
});

// Geometry aliases: these preset names reuse the same path as their source
// shape but differ in their default fill/line style, which comes from the
// shape's `p:style` references rather than the geometry.
const PRESET_SHAPE_ALIASES: Record<string, string> = {
  accentbordercallout1: "borderCallout1",
  accentbordercallout2: "borderCallout2",
  accentbordercallout3: "borderCallout3",
  accentcallout1: "borderCallout1",
  accentcallout2: "borderCallout2",
  accentcallout3: "borderCallout3",
  callout1: "borderCallout1",
  callout2: "borderCallout2",
  callout3: "borderCallout3"
};

const PRESET_SHAPE_BUILDERS_BY_LOWER_NAME = new Map<string, PresetShapeBuilder>();
for (const [name, builder] of PRESET_SHAPE_BUILDERS) {
  PRESET_SHAPE_BUILDERS_BY_LOWER_NAME.set(name.toLowerCase(), builder);
}

function resolvePresetShapeBuilder(geometry: string): PresetShapeBuilder | undefined {
  const name = geometry.trim();
  if (!name) {
    return undefined;
  }
  const lowerName = name.toLowerCase();
  return PRESET_SHAPE_BUILDERS.get(name)
    ?? PRESET_SHAPE_BUILDERS_BY_LOWER_NAME.get(lowerName)
    ?? PRESET_SHAPE_BUILDERS_BY_LOWER_NAME.get(PRESET_SHAPE_ALIASES[lowerName]?.toLowerCase() ?? "");
}

function toAdjustmentMap(
  adjustments: Record<string, number> | PresetShapeAdjustments | undefined
): PresetShapeAdjustments {
  if (!adjustments) {
    return new Map();
  }
  const maybeMap = adjustments as PresetShapeAdjustments;
  if (typeof maybeMap.get === "function" && typeof maybeMap.entries === "function") {
    return maybeMap;
  }
  const map = new Map<string, number>();
  for (const [name, value] of Object.entries(adjustments)) {
    if (Number.isFinite(value)) {
      map.set(name, value);
    }
  }
  return map;
}

function positiveExtent(value: number | undefined): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : 1;
}

/**
 * Build the SVG path and view box for a DrawingML preset shape.
 *
 * @param geometry `a:prstGeom@prst` (case-insensitive; `custom` is not a
 *   preset and resolves to `undefined`).
 * @param adjustments Raw `a:avLst/a:gd@fmla` values in 1/100000 units.
 * @param size Shape extents in the renderer's pixel space. Zero-width
 *   connectors are clamped to one unit so the path stays visible.
 */
export function buildPresetShapePath(
  geometry: string | undefined,
  adjustments?: Record<string, number> | PresetShapeAdjustments,
  size?: { height?: number; width?: number }
): PresetShapeGeometry | undefined {
  if (!geometry) {
    return undefined;
  }
  const builder = resolvePresetShapeBuilder(geometry);
  if (!builder) {
    return undefined;
  }
  const width = positiveExtent(size?.width);
  const height = positiveExtent(size?.height);
  const path = builder(width, height, toAdjustmentMap(adjustments));
  if (!path.trim()) {
    return undefined;
  }
  return {
    path,
    viewBox: { height, width }
  };
}
