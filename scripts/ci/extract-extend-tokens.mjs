#!/usr/bin/env node

/**
 * Extracts the Extend design tokens from the vendored copy of
 * extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * `apps/v4/app/globals.css` and writes docs/design/extend-tokens.json.
 *
 * Everything the report claims is either:
 * - read from the vendored upstream CSS (semantic colors, radius, motion),
 * - read from the pinned Tailwind v4 default theme that Extend relies on for
 *   text sizes, spacing, shadows and default transition timing, or
 * - derived from those two by resolving `var()`, `--alpha()`, `--theme()` and
 *   `color-mix(in srgb, ...)`, with the raw value kept next to the result.
 *
 * Usage:
 *   node scripts/ci/extract-extend-tokens.mjs            # write the report
 *   node scripts/ci/extract-extend-tokens.mjs --check    # fail if stale
 *   node scripts/ci/extract-extend-tokens.mjs --extend-dir /tmp/extend-ui
 *       # additionally verify the vendored CSS body against the upstream file
 */
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const require = createRequire(import.meta.url);

const PINNED = {
  repo: "extend-hq/ui",
  commit: "70fe70c6f1e1eee53954c6923f3d9f583557a57e",
  commitDate: "2026-11-17",
  globalsPath: "apps/v4/app/globals.css",
  globalsSha256:
    "f929d2c1de31312f6fc53bd9a7fff9f6a3feff46d4d5fd65805fea2deddbc759",
  // sha256 of globals.css from `@custom-variant style-vega` to EOF. The
  // vendored copy only changes the `@import` block above that line, so this
  // hash proves the rest of the file is untouched upstream content.
  bodySha256:
    "44bf6997bde067b701806833e88db08cfa75ae62b1805b63ee9b382107b44eca",
  bodyMarker: "@custom-variant style-vega",
  fontsPath: "apps/v4/lib/fonts.ts",
};

const vendoredCssPath = path.join(
  root,
  "tools/extend-reference/src/styles/extend-globals.css",
);
const tokenReportPath = path.join(root, "docs/design/extend-tokens.json");
const sourceManifestPath = path.join(
  root,
  "docs/design/extend-token-sources.json",
);

/** Semantic token names in the order Extend declares them. */
const SEMANTIC_COLORS = [
  "background",
  "foreground",
  "card",
  "card-foreground",
  "popover",
  "popover-foreground",
  "primary",
  "primary-foreground",
  "secondary",
  "secondary-foreground",
  "muted",
  "muted-foreground",
  "accent",
  "accent-foreground",
  "destructive",
  "destructive-foreground",
  "border",
  "input",
  "ring",
  "chart-1",
  "chart-2",
  "chart-3",
  "chart-4",
  "chart-5",
  "sidebar",
  "sidebar-foreground",
  "sidebar-primary",
  "sidebar-primary-foreground",
  "sidebar-accent",
  "sidebar-accent-foreground",
  "sidebar-border",
  "sidebar-ring",
  "surface",
  "surface-foreground",
  "code",
  "code-foreground",
  "code-highlight",
  "code-number",
  "selection",
  "selection-foreground",
  "info",
  "info-foreground",
  "success",
  "success-foreground",
  "warning",
  "warning-foreground",
];

/** Proposed mapping to the #156 `--office-*` token layer. */
const OFFICE_TOKEN_BY_SEMANTIC = Object.fromEntries(
  SEMANTIC_COLORS.map((name) => [name, `--office-${name}`]),
);

const RADIUS_FACTORS = {
  sm: 0.6,
  md: 0.8,
  lg: 1,
  xl: 1.4,
  "2xl": 1.8,
  "3xl": 2.2,
  "4xl": 2.6,
};

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function readText(file) {
  return readFileSync(file, "utf8");
}

function parseArgs(argv) {
  const options = {
    check: false,
    extendDir: process.env.EXTEND_UI_DIR || "",
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--check") options.check = true;
    else if (arg === "--extend-dir") options.extendDir = argv[++index] ?? "";
    else if (arg.startsWith("--extend-dir=")) {
      options.extendDir = arg.slice("--extend-dir=".length);
    } else {
      throw new Error(`unknown argument: ${arg}`);
    }
  }
  return options;
}

/* ------------------------------------------------------------------ colors */

function parseOklch(value) {
  const match = value.match(/^oklch\(\s*([\d.]+)(%?)\s+([\d.]+)\s+([\d.]+)(?:\s*\/\s*([\d.]+)(%?))?\s*\)$/i);
  if (!match) return null;
  const lightness = Number(match[1]) / (match[2] === "%" ? 100 : 1);
  const chroma = Number(match[3]);
  const hue = Number(match[4]);
  let alpha = 1;
  if (match[5] !== undefined) {
    alpha = Number(match[5]) / (match[6] === "%" ? 100 : 1);
  }
  return { l: lightness, c: chroma, h: hue, alpha };
}

function linearToSrgb(channel) {
  const clamped = Math.min(1, Math.max(0, channel));
  return clamped <= 0.0031308
    ? 12.92 * clamped
    : 1.055 * Math.pow(clamped, 1 / 2.4) - 0.055;
}

function oklchToRgb({ l, c, h }) {
  const radians = (h * Math.PI) / 180;
  const a = c * Math.cos(radians);
  const b = c * Math.sin(radians);
  const lPrime = l + 0.3963377774 * a + 0.2158037573 * b;
  const mPrime = l - 0.1055613458 * a - 0.0638541728 * b;
  const sPrime = l - 0.0894841775 * a - 1.291485548 * b;
  const l3 = lPrime ** 3;
  const m3 = mPrime ** 3;
  const s3 = sPrime ** 3;
  const red = 4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const green = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const blue = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.707614701 * s3;
  return [red, green, blue].map((channel) =>
    Math.round(linearToSrgb(channel) * 255),
  );
}

function toHex([red, green, blue]) {
  return `#${[red, green, blue]
    .map((channel) => channel.toString(16).padStart(2, "0"))
    .join("")}`;
}

function toCssRgba({ rgb, alpha }) {
  if (alpha >= 1) return `rgb(${rgb[0]} ${rgb[1]} ${rgb[2]})`;
  return `rgb(${rgb[0]} ${rgb[1]} ${rgb[2]} / ${Number(alpha.toFixed(4))})`;
}

function parseAlphaFunction(value) {
  const match = value.match(/^--alpha\((.*)\)$/);
  if (!match) return null;
  const [colorPart, alphaPart] = splitTopLevel(match[1], "/");
  if (!alphaPart) throw new Error(`--alpha() without alpha: ${value}`);
  const alphaText = alphaPart.trim();
  const alpha = alphaText.endsWith("%")
    ? Number(alphaText.slice(0, -1)) / 100
    : Number(alphaText);
  return { colorPart: colorPart.trim(), alpha };
}

function splitTopLevel(value, separator) {
  let depth = 0;
  for (let index = 0; index < value.length; index += 1) {
    const char = value[index];
    if (char === "(") depth += 1;
    else if (char === ")") depth -= 1;
    else if (char === separator && depth === 0) {
      return [value.slice(0, index), value.slice(index + 1)];
    }
  }
  return [value, ""];
}

function parseColorMix(value) {
  const match = value.match(/^color-mix\(\s*in\s+([\w-]+)\s*,\s*(.*)\)$/i);
  if (!match) return null;
  const space = match[1].toLowerCase();
  const parts = splitTopLevel(match[2], ",").map((part) => part.trim());
  const stops = parts.map((part) => {
    const percentage = part.match(/^(.*?)\s+([\d.]+)%$/);
    if (percentage) {
      return { value: percentage[1].trim(), percentage: Number(percentage[2]) };
    }
    return { value: part, percentage: null };
  });
  const missing = stops.filter((stop) => stop.percentage === null).length;
  const remaining = 100 - stops.reduce((sum, stop) => sum + (stop.percentage ?? 0), 0);
  for (const stop of stops) {
    if (stop.percentage === null) {
      stop.percentage = missing > 0 ? remaining / missing : 100 / stops.length;
    }
  }
  return { space, stops };
}

function createResolver(cssText) {
  const declared = new Map();
  for (const match of cssText.matchAll(/--([\w-]+)\s*:\s*([^;]+);/g)) {
    // Later declarations (`.dark` etc.) are captured per block elsewhere; this
    // map is the flat lookup used for `--color-*` palette references.
    declared.set(`--${match[1]}`, match[2].trim());
  }
  return declared;
}

function parseColor(rawValue, palette, depth = 0) {
  if (depth > 8) throw new Error(`color resolution depth exceeded: ${rawValue}`);
  const value = rawValue.trim();

  if (value === "transparent") {
    return { rgb: [0, 0, 0], alpha: 0, kind: "keyword" };
  }
  if (value === "currentColor") {
    return { rgb: [0, 0, 0], alpha: 1, kind: "keyword" };
  }

  const hex = value.match(/^#([0-9a-f]{3,8})$/i);
  if (hex) {
    const digits = hex[1];
    const expand = (text) => (text.length === 1 ? text + text : text);
    const rgb = [
      parseInt(expand(digits.slice(0, digits.length === 3 || digits.length === 4 ? 1 : 2)), 16),
      parseInt(expand(digits.slice(digits.length === 3 || digits.length === 4 ? 1 : 2, digits.length === 3 || digits.length === 4 ? 2 : 4)), 16),
      parseInt(expand(digits.slice(digits.length === 3 || digits.length === 4 ? 2 : 4, digits.length === 3 || digits.length === 4 ? 3 : 6)), 16),
    ];
    let alpha = 1;
    if (digits.length === 4) alpha = parseInt(expand(digits[3]), 16) / 255;
    if (digits.length === 8) alpha = parseInt(digits.slice(6, 8), 16) / 255;
    return { rgb, alpha, kind: "hex" };
  }

  const oklch = parseOklch(value);
  if (oklch) {
    return { rgb: oklchToRgb(oklch), alpha: oklch.alpha, kind: "oklch" };
  }

  const rgbMatch = value.match(/^rgba?\((.*)\)$/i);
  if (rgbMatch) {
    const parts = rgbMatch[1]
      .replace(/\//g, " ")
      .split(/[\s,]+/)
      .filter(Boolean);
    const channel = (text) =>
      text.endsWith("%")
        ? Math.round((Number(text.slice(0, -1)) / 100) * 255)
        : Number(text);
    const rgb = [channel(parts[0]), channel(parts[1]), channel(parts[2])];
    const alpha = parts[3] === undefined
      ? 1
      : parts[3].endsWith("%")
        ? Number(parts[3].slice(0, -1)) / 100
        : Number(parts[3]);
    return { rgb, alpha, kind: "rgb" };
  }

  const themeFunction = value.match(/^--theme\((.*)\)$/);
  if (themeFunction) {
    return parseColor(themeFunction[1], palette, depth + 1);
  }

  const alphaFunction = parseAlphaFunction(value);
  if (alphaFunction) {
    const inner = parseColor(alphaFunction.colorPart, palette, depth + 1);
    return {
      rgb: inner.rgb,
      alpha: inner.alpha * alphaFunction.alpha,
      kind: "alpha",
    };
  }

  const variable = value.match(/^var\((--[\w-]+)(?:\s*,\s*(.*))?\)$/);
  if (variable) {
    const resolved = palette.get(variable[1]);
    if (resolved === undefined) {
      if (variable[2] !== undefined) {
        return parseColor(variable[2], palette, depth + 1);
      }
      throw new Error(`unresolved custom property: ${variable[1]}`);
    }
    return parseColor(resolved, palette, depth + 1);
  }

  const mix = parseColorMix(value);
  if (mix) {
    const stops = mix.stops.map((stop) => ({
      color: parseColor(stop.value, palette, depth + 1),
      percentage: stop.percentage,
    }));
    const total = stops.reduce((sum, stop) => sum + stop.percentage, 0) || 100;
    let rgb = [0, 0, 0];
    let alpha = 0;
    for (const stop of stops) {
      const weight = stop.percentage / total;
      rgb = rgb.map((channel, index) => channel + stop.color.rgb[index] * weight);
      alpha += stop.color.alpha * weight;
    }
    return {
      rgb: rgb.map((channel) => Math.round(channel)),
      alpha,
      kind: `color-mix-${mix.space}`,
      approximate: mix.space !== "srgb",
    };
  }

  throw new Error(`unsupported color value: ${rawValue}`);
}

function colorReport(rawValue, palette) {
  const parsed = parseColor(rawValue, palette);
  return {
    raw: rawValue,
    resolved: toCssRgba(parsed),
    hex: toHex(parsed.rgb),
    alpha: Number(parsed.alpha.toFixed(4)),
    kind: parsed.kind,
    ...(parsed.approximate
      ? { approximate: "mixed in sRGB; upstream mixes in oklab" }
      : {}),
  };
}

/* ----------------------------------------------------------------- parsing */

function extractBlock(cssText, selector) {
  const start = cssText.indexOf(`${selector} {`);
  if (start < 0) throw new Error(`selector not found: ${selector}`);
  let depth = 0;
  let index = cssText.indexOf("{", start);
  const bodyStart = index + 1;
  for (; index < cssText.length; index += 1) {
    const char = cssText[index];
    if (char === "{") depth += 1;
    else if (char === "}") {
      depth -= 1;
      if (depth === 0) {
        return cssText.slice(bodyStart, index);
      }
    }
  }
  throw new Error(`unterminated block: ${selector}`);
}

function extractDeclarations(blockText) {
  const declarations = new Map();
  for (const match of blockText.matchAll(/(^|\n)\s*(--[\w-]+)\s*:\s*([^;]+);/g)) {
    declarations.set(match[2], match[3].replace(/\s+/g, " ").trim());
  }
  return declarations;
}

function extractThemeInline(cssText) {
  const block = extractBlock(cssText, "@theme inline");
  return extractDeclarations(block);
}

function readTailwindTheme(root) {
  const tailwindThemePath = require.resolve("tailwindcss/theme.css", {
    paths: [path.join(root, "tools/extend-reference")],
  });
  const css = readText(tailwindThemePath);
  const palette = new Map();
  const defaults = new Map();
  for (const match of css.matchAll(/(--[\w-]+)\s*:\s*([^;]+);/g)) {
    const name = match[1];
    const value = match[2].replace(/\s+/g, " ").trim();
    defaults.set(name, value);
    if (name.startsWith("--color-")) palette.set(name, value);
  }
  return { palette, defaults, file: tailwindThemePath };
}

function rootRelative(file) {
  return path.relative(root, file).split(path.sep).join("/");
}

function walkSourceFiles(directories, extensions) {
  const files = [];
  for (const directory of directories) {
    const absolute = path.join(root, directory);
    if (!existsSync(absolute)) continue;
    for (const entry of readdirSync(absolute, { withFileTypes: true })) {
      const child = path.join(absolute, entry.name);
      if (entry.isDirectory()) {
        if (entry.name === "node_modules" || entry.name === "dist") continue;
        files.push(...walkSourceFiles([path.relative(root, child)], extensions));
      } else if (extensions.some((extension) => entry.name.endsWith(extension))) {
        files.push(child);
      }
    }
  }
  return files;
}

function collectCurrentUsage() {
  const sourceFiles = walkSourceFiles(
    ["packages/vue-ui/src", "packages/vue-pdf/src", "packages/vue-docx/src", "packages/vue-xlsx/src", "packages/vue-pptx/src"],
    [".vue", ".css", ".ts"],
  );
  const usage = new Map();
  for (const file of sourceFiles) {
    const text = readText(file);
    for (const match of text.matchAll(/var\(\s*(--[\w-]+)\s*,\s*([^()]*(?:\([^()]*\)[^()]*)*)\)/g)) {
      const token = match[1];
      const fallback = match[2].trim();
      if (!usage.has(token)) usage.set(token, []);
      const entries = usage.get(token);
      if (entries.length >= 6) continue;
      entries.push({ file: rootRelative(file), fallback });
    }
  }
  return usage;
}

function collectOfficeTokenPresence() {
  const files = walkSourceFiles(["packages", "apps"], [".vue", ".css", ".ts"]);
  const present = new Set();
  for (const file of files) {
    const text = readText(file);
    for (const match of text.matchAll(/--office-[\w-]+/g)) present.add(match[0]);
  }
  return present;
}

function collectComponentMotion(componentFiles) {
  const durations = new Map();
  const easings = new Map();
  for (const file of componentFiles) {
    const text = readText(file);
    for (const match of text.matchAll(/\bduration-(\d+)\b/g)) {
      const key = `duration-${match[1]}`;
      durations.set(key, { class: key, ms: Number(match[1]) });
    }
    for (const match of text.matchAll(/\bduration-\[([\d.]+)(ms|s)\]/g)) {
      const key = `duration-[${match[1]}${match[2]}]`;
      durations.set(key, {
        class: key,
        ms: Math.round(Number(match[1]) * (match[2] === "s" ? 1000 : 1)),
      });
    }
    for (const match of text.matchAll(/\bdelay-(\d+)\b/g)) {
      const key = `delay-${match[1]}`;
      durations.set(key, { class: key, ms: Number(match[1]) });
    }
    for (const match of text.matchAll(/ease-\[([^\]]+)\]/g)) {
      const key = `ease-[${match[1]}]`;
      easings.set(key, { class: key, value: match[1] });
    }
    for (const match of text.matchAll(/\bease-([\w-]+)\b/g)) {
      if (match[1].startsWith("in") || match[1] === "linear" || match[1] === "out") {
        easings.set(`ease-${match[1]}`, { class: `ease-${match[1]}`, value: match[1] });
      }
    }
  }
  return {
    durations: [...durations.values()].sort((left, right) => left.ms - right.ms),
    easings: [...easings.values()].sort((left, right) =>
      left.class.localeCompare(right.class),
    ),
  };
}

function main() {
  const options = parseArgs(process.argv.slice(2));
  const cssText = readText(vendoredCssPath);
  const markerIndex = cssText.indexOf(PINNED.bodyMarker);
  if (markerIndex < 0) {
    throw new Error(`vendored CSS has no ${PINNED.bodyMarker} marker`);
  }
  const body = cssText.slice(markerIndex);
  const bodySha = sha256(body);
  if (bodySha !== PINNED.bodySha256) {
    throw new Error(
      `vendored globals.css body changed (${bodySha} != ${PINNED.bodySha256}); ` +
        "re-vendor the pinned upstream file instead of editing it",
    );
  }

  let upstreamVerification = "not-run";
  if (options.extendDir) {
    const upstreamPath = path.join(options.extendDir, PINNED.globalsPath);
    if (!existsSync(upstreamPath)) {
      throw new Error(`--extend-dir has no ${PINNED.globalsPath}: ${options.extendDir}`);
    }
    const upstream = readText(upstreamPath);
    if (sha256(upstream) !== PINNED.globalsSha256) {
      throw new Error(
        `upstream globals.css hash mismatch in ${options.extendDir}: not the pinned commit`,
      );
    }
    const upstreamBody = upstream.slice(upstream.indexOf(PINNED.bodyMarker));
    if (sha256(upstreamBody) !== bodySha) {
      throw new Error("vendored CSS body differs from the upstream checkout");
    }
    upstreamVerification = `verified against ${rootRelative(upstreamPath)}`;
  }

  const tailwind = readTailwindTheme(root);
  const declarations = new Map([...tailwind.defaults, ...extractDeclarations(extractBlock(cssText, ":root"))]);
  // Palette + semantic declarations are needed to resolve `.dark` values
  // (`var(--sidebar)` refers to the .dark value in the .dark block).
  const palette = new Map([...tailwind.palette, ...declarations]);

  const themeInline = extractThemeInline(cssText);
  const lightDecls = extractDeclarations(extractBlock(cssText, ":root"));
  const darkDecls = extractDeclarations(extractBlock(cssText, ".dark"));

  const resolveTheme = (themeDecls) => {
    const colors = {};
    for (const name of SEMANTIC_COLORS) {
      const raw = themeDecls.get(`--${name}`);
      if (raw === undefined) {
        colors[name] = null;
        continue;
      }
      const localPalette = new Map(palette);
      for (const [key, value] of themeDecls) localPalette.set(key, value);
      colors[name] = colorReport(raw, localPalette);
    }
    return colors;
  };

  const light = resolveTheme(lightDecls);
  const dark = resolveTheme(darkDecls);

  const radiusRaw = lightDecls.get("--radius") ?? darkDecls.get("--radius");
  const radiusRem = Number.parseFloat(radiusRaw);
  const radiusPx = (factor) => `${Number((radiusRem * factor * 16).toFixed(4))}px`;
  const radius = {
    raw: radiusRaw,
    basePx: radiusPx(1),
    rem: radiusRem,
    scale: Object.fromEntries(
      Object.entries(RADIUS_FACTORS).map(([name, factor]) => [
        name,
        {
          raw: `calc(var(--radius) * ${factor})`,
          px: radiusPx(factor),
          varName: `--radius-${name}`,
          tailwindTheme: themeInline.get(`--radius-${name}`) ?? null,
        },
      ]),
    ),
  };

  const typography = {};
  for (const [name, value] of tailwind.defaults) {
    const size = name.match(/^--text-([\w-]+)$/);
    if (size) {
      const lineHeight = tailwind.defaults.get(`${name}--line-height`) ?? null;
      const letterSpacing = tailwind.defaults.get(`${name}--letter-spacing`) ?? null;
      const weight = tailwind.defaults.get(`${name}--font-weight`) ?? null;
      typography[size[1]] = {
        fontSize: value,
        fontSizePx: Number.parseFloat(value) * (value.endsWith("rem") ? 16 : 1),
        lineHeight,
        lineHeightPx: lineHeight?.endsWith("rem")
          ? Number((Number.parseFloat(lineHeight) * 16).toFixed(2))
          : lineHeight,
        letterSpacing,
        fontWeight: weight,
      };
    }
  }

  const spacingRaw = tailwind.defaults.get("--spacing") ?? "0.25rem";
  const spacingRem = Number.parseFloat(spacingRaw);

  const shadows = {};
  for (const [name, value] of tailwind.defaults) {
    if (/^--shadow-/.test(name)) {
      shadows[name.replace("--shadow-", "")] = { raw: value, utility: name };
    }
    if (/^--inset-shadow-/.test(name)) {
      shadows[`inset-${name.replace("--inset-shadow-", "")}`] = {
        raw: value,
        utility: name,
      };
    }
  }

  const motion = {
    defaultDuration:
      tailwind.defaults.get("--default-transition-duration") ?? null,
    defaultTimingFunction:
      tailwind.defaults.get("--default-transition-timing-function") ?? null,
    animation: {
      skeleton: themeInline.get("--animate-skeleton") ?? null,
      spin:
        tailwind.defaults.get("--animate-spin") ??
        "spin 1s linear infinite (Tailwind default)",
    },
    ...collectComponentMotion([
      path.join(root, "tools/extend-reference/src/vendor/ui/button.tsx"),
      path.join(root, "tools/extend-reference/src/vendor/ui/select.tsx"),
      path.join(root, "tools/extend-reference/src/vendor/extend/file-upload.tsx"),
      path.join(root, "tools/extend-reference/src/vendor/extend/file-thumbnail.tsx"),
      path.join(root, "tools/extend-reference/src/vendor/extend/document-viewer-sidebar.tsx"),
    ]),
  };

  const ring = {
    color: { light: light.ring, dark: dark.ring },
    focusRing: {
      source: "button.tsx default variant",
      classes: ["focus-visible:ring-2", "focus-visible:ring-offset-1", "focus-visible:ring-offset-background"],
      width: "2px",
      offset: "1px",
      offsetColor: { light: light.background, dark: dark.background },
      note: "Select trigger uses focus-visible:ring-[3px] instead of 2px (see mapping.ringWidth).",
    },
  };

  const currentUsage = collectCurrentUsage();
  const officeTokensPresent = collectOfficeTokenPresence();
  const mapping = SEMANTIC_COLORS.map((name) => {
    const officeToken = OFFICE_TOKEN_BY_SEMANTIC[name];
    const usage = currentUsage.get(`--${name}`) ?? [];
    return {
      extendToken: `--${name}`,
      officeToken,
      officeTokenStatus: officeTokensPresent.has(officeToken)
        ? "present"
        : "pending (#156 adds the --office-* layer; no --office-* custom property exists in the repo yet)",
      extendLight: light[name],
      extendDark: dark[name],
      currentRepoUsage: usage,
      currentRepoValue:
        usage.length === 0
          ? null
          : [...new Set(usage.map((entry) => entry.fallback))].join(" | "),
    };
  });

  const report = {
    $schema: "./extend-tokens.schema.json",
    generatedBy: "node scripts/ci/extract-extend-tokens.mjs",
    generatedAt: new Date().toISOString(),
    provenance: {
      repo: PINNED.repo,
      commit: PINNED.commit,
      commitDate: PINNED.commitDate,
      upstreamFile: PINNED.globalsPath,
      upstreamSha256: PINNED.globalsSha256,
      vendoredFile: rootRelative(vendoredCssPath),
      vendoredBodySha256: bodySha,
      vendoredBodyMarker: PINNED.bodyMarker,
      upstreamVerification,
      tailwindThemeFile: rootRelative(tailwind.file),
      fonts: {
        source: PINNED.fontsPath,
        sans: "Geist",
        mono: "Geist Mono",
        harnessStandIn: "@fontsource-variable/geist + @fontsource-variable/geist-mono",
      },
      notes: [
        "Text sizes, line heights, letter spacing, spacing, shadows and the default transition timing are Tailwind v4 defaults that Extend consumes implicitly; they are read from the pinned tailwindcss theme.css rather than from globals.css.",
        "--alpha(color / n%) is resolved as color-mix(in oklab, color n%, transparent). All Extend uses are black/white, where the sRGB and oklab results agree, so the resolved value is exact for this file.",
        "color-mix(in srgb, ...) is resolved by mixing sRGB channels; color-mix(in oklab, ...) is flagged approximate.",
      ],
    },
    colors: { light, dark },
    radius,
    typography: {
      scale: typography,
      fonts: {
        sans: "var(--font-sans)",
        mono: "var(--font-mono)",
        sansFallback:
          '"Geist Variable", ui-sans-serif, system-ui, -apple-system, "Segoe UI", sans-serif',
        monoFallback:
          '"Geist Mono Variable", ui-monospace, SFMono-Regular, Menlo, monospace',
        note: "apps/v4/lib/fonts.ts loads Geist/Geist Mono through next/font and assigns --font-sans / --font-mono; the harness uses the same families from @fontsource-variable.",
      },
    },
    spacing: {
      base: { raw: spacingRaw, px: spacingRem * 16 },
      scalePx: Object.fromEntries(
        [0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8].map((step) => [
          String(step),
          Number((spacingRem * 16 * step).toFixed(2)),
        ]),
      ),
      note: "Extend components use calc(--spacing(n) - 1px) for the 1px inset border, e.g. button px-[calc(--spacing(3)-1px)] = 11px.",
    },
    shadows,
    ring,
    motion,
    mapping,
  };

  const serialized = `${JSON.stringify(report, null, 2)}\n`;
  const sources = {
    $comment:
      "Pinned inputs for scripts/ci/extract-extend-tokens.mjs. Update both hashes together when bumping the reference commit, and re-vendor tools/extend-reference/src/styles/extend-globals.css.",
    repo: PINNED.repo,
    commit: PINNED.commit,
    commitDate: PINNED.commitDate,
    globalsPath: PINNED.globalsPath,
    globalsSha256: PINNED.globalsSha256,
    bodyMarker: PINNED.bodyMarker,
    bodySha256: PINNED.bodySha256,
    fontsPath: PINNED.fontsPath,
    tailwindVersion: "4.2.4 (matches the pinned repo's lockfile)",
  };

  if (options.check) {
    const current = existsSync(tokenReportPath) ? readText(tokenReportPath) : "";
    let matches = false;
    try {
      const parsed = JSON.parse(current);
      const strip = (value) => {
        const clone = JSON.parse(JSON.stringify(value));
        delete clone.generatedAt;
        return JSON.stringify(clone, null, 2);
      };
      matches = strip(parsed) === strip(JSON.parse(serialized));
    } catch {
      matches = false;
    }
    if (!matches) {
      console.error(
        `docs/design/extend-tokens.json is stale; run \`node scripts/ci/extract-extend-tokens.mjs\``,
      );
      process.exit(1);
    }
    console.log("extend tokens: up to date");
    return;
  }

  mkdirSync(path.dirname(tokenReportPath), { recursive: true });
  writeFileSync(tokenReportPath, serialized);
  writeFileSync(sourceManifestPath, `${JSON.stringify(sources, null, 2)}\n`);
  console.log(
    `extend tokens: wrote ${rootRelative(tokenReportPath)} (${SEMANTIC_COLORS.length} semantic colors, radius ${radius.basePx}, ${motion.durations.length} durations)`,
  );
}

main();
