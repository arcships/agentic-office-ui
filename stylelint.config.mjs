/**
 * Component CSS style guards (#165).
 *
 * Component styles must consume the `--office-*` token layer instead of
 * hard-coding visual values, so an Extend parity change is a token change
 * rather than a hundred-component search. The four guards below forbid a bare
 * value; the token layer (`packages/*\/src/styles/tokens.css` and the theme
 * file that maps it) is exempt because defining those values is its job.
 *
 * Existing violations are frozen in `tests/style/component-style-baseline.json`
 * and policed by `scripts/ci/lint-styles.mjs`: the guard fails on new
 * violations only, and the baseline can only grow through an approved update
 * (`pnpm lint:style:baseline --approve ...`). The rules themselves are never
 * disabled for a whole directory.
 */
export default {
  // Vue SFC <style> blocks are linted through postcss-html; plain CSS uses
  // stylelint's default syntax.
  overrides: [
    {
      files: ["**/*.vue"],
      customSyntax: "postcss-html",
    },
    {
      // The token layer defines the raw values every other file must consume.
      files: [
        "packages/*/src/styles/tokens.css",
        "packages/*/src/styles/theme.css",
        "packages/*/src/**/tokens.css",
      ],
      rules: {
        "color-no-hex": null,
        "declaration-property-unit-disallowed-list": null,
        "declaration-property-value-disallowed-list": null,
      },
    },
  ],
  rules: {
    // 1. No bare hex colours: use a semantic `--office-*` token.
    "color-no-hex": true,
    // 2. No bare radius / font-size / shadow geometry / transition timing.
    "declaration-property-unit-disallowed-list": {
      "/^border(-(top|right|bottom|left))?-radius$/": [
        "px",
        "rem",
        "em",
        "pt",
        "%",
        "vw",
        "vh",
      ],
      "border-radius": ["px", "rem", "em", "pt", "%", "vw", "vh"],
      "font-size": ["px", "rem", "em", "pt", "%", "vw", "vh"],
      "/^box-shadow$/": ["px", "rem", "em"],
      "text-shadow": ["px", "rem", "em"],
      "/^transition(-duration|-delay)?$/": ["s", "ms"],
    },
    // 3. No bare shadow colour or literal easing inside a shadow/transition.
    "declaration-property-value-disallowed-list": {
      "/^(box-shadow|text-shadow)$/": [
        "/rgba?\\(/",
        "/hsla?\\(/",
        "/#[0-9a-fA-F]{3,8}/",
      ],
      // Timing functions that cannot come from a keyword (cubic-bezier/steps)
      // have to be a token; `ease`/`linear` stay legal because the 0.8 token
      // layer has no easing token and CSS keywords are the parity baseline.
      "/^transition(-timing-function)?$/": [
        "/cubic-bezier\\(/",
        "/steps\\(/",
      ],
    },
  },
};
