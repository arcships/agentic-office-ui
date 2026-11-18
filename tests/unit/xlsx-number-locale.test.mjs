import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const ts = requireFromXlsxCore("typescript");

const numberLocaleModule = (() => {
  const source = readFileSync(
    new URL("../../packages/vue-xlsx/src/composables/number-locale.ts", import.meta.url),
    "utf8",
  );
  const emitted = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return `data:text/javascript;base64,${Buffer.from(emitted).toString("base64")}`;
})();

const {
  localizeCellNumber,
  parseDisplayedNumber,
  parseLocaleNumber,
} = await import(numberLocaleModule);

test("localizeCellNumber applies the viewer locale to numeric formats", () => {
  assert.equal(localizeCellNumber("1,234.56", "#,##0.00", "de-DE"), "1.234,56");
  assert.equal(localizeCellNumber("1,234.56", "#,##0.00", "fr-FR"), "1\u202f234,56");
  assert.equal(localizeCellNumber("1,234.56", "#,##0.00", "en-US"), "1,234.56");
  assert.equal(localizeCellNumber("50.00%", "0.00%", "de-DE"), "50,00%");
  assert.equal(localizeCellNumber("1.23E+03", "0.00E+00", "de-DE"), "1,23E+03");
  assert.equal(localizeCellNumber("€1,234.56", "[$€-407]#,##0.00", "de-DE"), "€1.234,56");
});

test("a format's own locale section wins over the fallback locale", () => {
  assert.equal(localizeCellNumber("1,234.56", "[$-407]#,##0.00", "en-US"), "1.234,56");
  assert.equal(localizeCellNumber("1,234.56", "[$-409]#,##0.00", "de-DE"), "1,234.56");
});

test("localizeCellNumber protects literals, text and date formats", () => {
  // Quoted and escaped literals may contain digits and separators of their own.
  assert.equal(localizeCellNumber("kg1,234.56", '"kg"#,##0.00', "de-DE"), "kg1.234,56");
  assert.equal(localizeCellNumber("1,234.56 kg", '#,##0.00" kg"', "de-DE"), "1.234,56 kg");
  assert.equal(localizeCellNumber("v1.2-1,234.56", '"v1.2-"#,##0.00', "de-DE"), "v1.2-1.234,56");

  // Text and date/time formats keep their punctuation; only numeric formats
  // are rewritten.
  assert.equal(localizeCellNumber("1,234", "@", "de-DE"), "1,234");
  assert.equal(localizeCellNumber("1,234", "[$-407]@", "de-DE"), "1,234");
  assert.equal(localizeCellNumber("15.01.2024", "dd.mm.yyyy", "de-DE"), "15.01.2024");
  assert.equal(localizeCellNumber("1/15/2024", "m/d/yyyy", "de-DE"), "1/15/2024");
  assert.equal(localizeCellNumber("12:30:45", "hh:mm:ss", "de-DE"), "12:30:45");
  assert.equal(localizeCellNumber("1,234.56", "General", "de-DE"), "1,234.56");
});

test("parseLocaleNumber understands grouped and decimal marks", () => {
  // Both marks are unambiguous even without a locale.
  assert.equal(parseLocaleNumber("1.234,5"), 1234.5);
  assert.equal(parseLocaleNumber("1,234.5"), 1234.5);

  // With one mark the locale decides.
  assert.equal(parseLocaleNumber("1.234", "en-US"), 1.234);
  assert.equal(parseLocaleNumber("1.234", "de-DE"), 1234);
  assert.equal(parseLocaleNumber("1,5", "de-DE"), 1.5);
  assert.equal(parseLocaleNumber("1,234", "en-US"), 1234);

  // Spaces, apostrophes, signs and accounting parentheses are accepted.
  assert.equal(parseLocaleNumber("1 234,5", "de-DE"), 1234.5);
  assert.equal(parseLocaleNumber("1'234.5", "de-DE"), 1234.5);
  assert.equal(parseLocaleNumber("-1.234,5", "de-DE"), -1234.5);
  assert.equal(parseLocaleNumber("(1.234,5)", "de-DE"), -1234.5);
  assert.equal(parseLocaleNumber("not a number", "de-DE"), null);
});

test("parseDisplayedNumber follows the locale for a single separator", () => {
  assert.equal(parseDisplayedNumber("1.234,56", "de-DE"), 1234.56);
  assert.equal(parseDisplayedNumber("1,234.56", "en-US"), 1234.56);
  assert.equal(parseDisplayedNumber("$1,234.56", "en-US"), 1234.56);
  assert.equal(parseDisplayedNumber("(1,234.56)", "en-US"), -1234.56);
  assert.equal(parseDisplayedNumber("1.234", "de-DE"), 1234);
  assert.equal(parseDisplayedNumber("1.234"), 1.234);
  assert.equal(parseDisplayedNumber("50%", "de-DE"), 50);
});
