// Issue #156 — OfficeThemeProvider / OfficeIcon runtime coverage.
//
// Uses the repository's Vue test renderer (no browser DOM) plus Vue's server
// renderer for the SSR contract. The suite starts from the built package, so it
// also proves the public exports and the emitted theme.css asset.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const {
  OFFICE_ICONS,
  OFFICE_ICON_NAMES,
  OfficeIcon,
  OfficeThemeProvider,
  normalizeOfficeIconName,
  officeSystemPrefersDark,
  provideOfficeIcons,
  resolveOfficeIcon,
  resolveOfficeTheme,
  useOfficeTheme,
} = await importFromDemo("@arcships/vue-ui");
const { renderToString } = await importFromDemo("vue/server-renderer");

const ROOT = new URL("../../", import.meta.url);

function read(relativePath) {
  return readFileSync(new URL(relativePath, ROOT), "utf8");
}

const foundSvg = (root) => walk(root).find((node) => node.type === "svg");

function installMatchMedia(initialMatches) {
  const queries = [];
  const listeners = new Set();
  globalThis.window = {
    matchMedia(query) {
      const mediaQuery = {
        matches: initialMatches,
        media: query,
        addEventListener(type, listener) {
          if (type === "change") listeners.add(listener);
        },
        removeEventListener(type, listener) {
          if (type === "change") listeners.delete(listener);
        },
      };
      queries.push(mediaQuery);
      return mediaQuery;
    },
  };
  return {
    queries,
    listenerCount: () => listeners.size,
    setMatches(matches) {
      for (const mediaQuery of queries) mediaQuery.matches = matches;
      for (const listener of listeners) {
        listener({ matches, media: "(prefers-color-scheme: dark)" });
      }
    },
  };
}

function removeWindow() {
  delete globalThis.window;
}

test("resolveOfficeTheme is pure and only system consults the OS preference", () => {
  assert.equal(resolveOfficeTheme("light", true), "light");
  assert.equal(resolveOfficeTheme("dark", false), "dark");
  assert.equal(resolveOfficeTheme("system", true), "dark");
  assert.equal(resolveOfficeTheme("system", false), "light");
});

test("OfficeThemeProvider SSR renders without window and keeps both attributes", async () => {
  assert.equal(officeSystemPrefersDark(), false, "no window must resolve to light");
  const html = await renderToString(
    vue.h(OfficeThemeProvider, { theme: "system" }, () =>
      vue.h(OfficeIcon, { name: "check", label: "Done" }),
    ),
  );

  assert.match(html, /data-office-theme="system"/);
  assert.match(html, /data-office-theme-resolved="light"/);
  assert.match(html, /data-office-density="comfortable"/);
  assert.match(html, /<svg[^>]*stroke-width="2"/);
  assert.match(html, /<svg[^>]*aria-label="Done"/);
  assert.match(html, /<svg[^>]*role="img"/);
  assert.doesNotMatch(html, /aria-hidden="true"[^>]*aria-label/u);

  const darkHtml = await renderToString(
    vue.h(
      OfficeThemeProvider,
      { theme: "dark", density: "compact", locale: "zh-CN" },
      () => vue.h("span", "文档"),
    ),
  );
  assert.match(darkHtml, /data-office-theme="dark"/);
  assert.match(darkHtml, /data-office-theme-resolved="dark"/);
  assert.match(darkHtml, /data-office-density="compact"/);
  assert.match(darkHtml, /lang="zh-CN"/);
  assert.match(darkHtml, /文档/);
});

test("OfficeThemeProvider switches light, dark and system at runtime", async (t) => {
  const theme = vue.ref("light");
  const Wrapper = vue.defineComponent({
    setup() {
      return () =>
        vue.h(
          OfficeThemeProvider,
          { theme: theme.value, density: "compact" },
          () => vue.h("span", { "data-testid": "child" }, "content"),
        );
    },
  });

  const { app, root } = await mount(Wrapper);
  t.after(() => app.unmount());

  const host = () => root.children[0];
  const child = () => findByTestId(root, "child");

  assert.equal(host().props["data-office-theme"], "light");
  assert.equal(host().props["data-office-theme-resolved"], "light");
  assert.equal(host().props["data-office-density"], "compact");
  assert.equal(textContent(child()), "content");

  theme.value = "dark";
  await vue.nextTick();
  assert.equal(host().props["data-office-theme"], "dark");
  assert.equal(host().props["data-office-theme-resolved"], "dark");

  theme.value = "system";
  await vue.nextTick();
  assert.equal(host().props["data-office-theme"], "system");
  assert.equal(host().props["data-office-theme-resolved"], "light");
});

test("system theme follows matchMedia changes and cleans up its listener", async (t) => {
  const media = installMatchMedia(true);
  t.after(removeWindow);

  const { app, root } = await mount(OfficeThemeProvider, { theme: "system" });

  const host = root.children[0];
  assert.equal(media.queries.length, 1);
  assert.equal(media.queries[0].media, "(prefers-color-scheme: dark)");
  assert.equal(host.props["data-office-theme"], "system");
  assert.equal(host.props["data-office-theme-resolved"], "dark");

  media.setMatches(false);
  await vue.nextTick();
  assert.equal(host.props["data-office-theme-resolved"], "light");

  media.setMatches(true);
  await vue.nextTick();
  assert.equal(host.props["data-office-theme-resolved"], "dark");

  app.unmount();
  assert.equal(media.listenerCount(), 0, "matchMedia listener must be removed on unmount");
});

test("OfficeThemeProvider publishes theme context and locale", async (t) => {
  const Probe = vue.defineComponent({
    setup() {
      const context = useOfficeTheme();
      return () =>
        vue.h(
          "span",
          { "data-testid": "probe" },
          `${context.theme.value}|${context.resolvedTheme.value}|${context.density.value}|${context.isDark.value}|${context.locale.value ?? ""}`,
        );
    },
  });

  const Wrapper = vue.defineComponent({
    setup: () => () =>
      vue.h(
        OfficeThemeProvider,
        { theme: "dark", density: "compact", locale: "zh-CN" },
        () => vue.h(Probe),
      ),
  });
  const { app, root } = await mount(Wrapper);
  t.after(() => app.unmount());
  assert.equal(textContent(findByTestId(root, "probe")), "dark|dark|compact|true|zh-CN");
  assert.equal(root.children[0].props.lang, "zh-CN");
});

test("nested providers override an ancestor without touching the host", async (t) => {
  const Probe = vue.defineComponent({
    setup() {
      const context = useOfficeTheme();
      return () =>
        vue.h("span", { "data-testid": "nested-probe" }, context.resolvedTheme.value);
    },
  });

  const { app, root } = await mount(
    vue.defineComponent({
      setup: () => () =>
        vue.h(OfficeThemeProvider, { theme: "dark" }, () =>
          vue.h(OfficeThemeProvider, { theme: "light" }, () => vue.h(Probe)),
        ),
    }),
  );
  t.after(() => app.unmount());

  const hosts = walk(root).filter(
    (node) => node.props && "data-office-theme" in node.props,
  );
  assert.deepEqual(
    hosts.map((host) => host.props["data-office-theme"]),
    ["dark", "light"],
  );
  assert.equal(textContent(findByTestId(root, "nested-probe")), "light");
});

test("OfficeIcon renders registry icons with size, stroke and a11y props", async () => {
  assert.ok(OFFICE_ICONS.check);
  assert.ok(Object.isFrozen(OFFICE_ICONS));
  assert.ok(Object.isFrozen(OFFICE_ICON_NAMES));
  assert.ok(OFFICE_ICON_NAMES.includes("chevron-down"));
  assert.equal(normalizeOfficeIconName("ChevronDown"), "chevron-down");
  assert.equal(normalizeOfficeIconName("chevron_down"), "chevron-down");
  assert.equal(normalizeOfficeIconName("Redo2"), "redo-2");
  assert.equal(resolveOfficeIcon("ChevronDown"), OFFICE_ICONS["chevron-down"]);
  assert.equal(resolveOfficeIcon("chevron-down"), OFFICE_ICONS["chevron-down"]);
  assert.equal(resolveOfficeIcon("Redo2"), OFFICE_ICONS["redo-2"]);
  assert.equal(resolveOfficeIcon("not-an-icon"), undefined);

  const plain = await mount(OfficeIcon, { name: "check" });
  const svg = foundSvg(plain.root);
  assert.ok(svg, "OfficeIcon must render an svg");
  assert.equal(svg.props["aria-hidden"], "true");
  assert.equal(svg.props.focusable, "false");
  assert.equal(svg.props["stroke-width"], 2);
  assert.equal(svg.props.width, 16, "default SVG width must match the 1rem token");
  assert.equal(svg.props.height, 16);
  assert.ok(String(svg.props.class).includes("office-icon"));
  plain.app.unmount();

  const labelled = await mount(OfficeIcon, { name: "search", label: "Search", strokeWidth: 1.5 });
  const labelledSvg = foundSvg(labelled.root);
  assert.equal(labelledSvg.props["aria-label"], "Search");
  assert.equal(labelledSvg.props.role, "img");
  assert.equal(labelledSvg.props["aria-hidden"], undefined);
  assert.equal(labelledSvg.props["stroke-width"], 1.5);
  labelled.app.unmount();

  const sized = await mount(OfficeIcon, { name: "plus", size: 20 });
  const sizedSvg = foundSvg(sized.root);
  assert.equal(sizedSvg.props.width, 20);
  assert.equal(sizedSvg.props.height, 20);
  assert.equal(sizedSvg.props.style["--office-icon-size"], "20px");
  sized.app.unmount();

  const cssSized = await mount(OfficeIcon, { name: "plus", size: "1.25rem" });
  const cssSvg = foundSvg(cssSized.root);
  assert.equal(cssSvg.props.style["--office-icon-size"], "1.25rem");
  assert.equal(cssSvg.props.width, 16);
  cssSized.app.unmount();
});

test("OfficeIcon supports custom components and a provided registry", async (t) => {
  const Custom = vue.defineComponent({
    setup: () => () => vue.h("svg", { "data-testid": "custom-icon" }),
  });

  const direct = await mount(OfficeIcon, { icon: Custom });
  assert.ok(findByTestId(direct.root, "custom-icon"));
  assert.equal(foundSvg(direct.root).props["aria-hidden"], "true");
  direct.app.unmount();

  const Star = vue.defineComponent({
    setup: () => () => vue.h("svg", { "data-testid": "star-icon" }),
  });

  const Wrapper = vue.defineComponent({
    setup() {
      provideOfficeIcons({ "custom-star": Star });
      return () => vue.h(OfficeIcon, { name: "custom-star", label: "Starred" });
    },
  });
  const { app, root } = await mount(Wrapper);
  t.after(() => app.unmount());

  const star = findByTestId(root, "star-icon");
  assert.ok(star, "provided registry entries must resolve");
  assert.equal(star.props["aria-label"], "Starred");
});

test("OfficeIcon fails soft for unknown names", async () => {
  const warnings = [];
  const originalWarn = console.warn;
  console.warn = (...args) => warnings.push(args.join(" "));
  try {
    const { app, root } = await mount(OfficeIcon, { name: "definitely-not-an-icon" });
    assert.equal(foundSvg(root), undefined);
    assert.equal(root.children[0]?.type, "#comment");
    app.unmount();
  } finally {
    console.warn = originalWarn;
  }
  assert.ok(
    warnings.some((warning) => warning.includes("Unknown icon")),
    "unknown icon must warn once with an actionable message",
  );
});

test("built package ships self-contained theme.css, tokens.css and icon styles", () => {
  const theme = read("packages/vue-ui/dist/theme.css");
  const tokens = read("packages/vue-ui/dist/tokens.css");
  const style = read("packages/vue-ui/dist/style.css");

  assert.match(theme, /--office-background:/u);
  assert.match(theme, /--office-border:/u);
  assert.match(theme, /\[data-office-theme="dark"\]/u);
  assert.match(theme, /@media \(prefers-color-scheme: dark\)/u);
  assert.match(theme, /--office-success:/u);
  assert.doesNotMatch(theme, /@import\s+["']\.\/tokens\.css["']/u, "dist theme.css must be self-contained");
  assert.match(tokens, /--office-radius: 0\.625rem/u);
  assert.match(style, /office-icon/u);
  assert.match(style, /var\(--office-icon-size/u);
});
