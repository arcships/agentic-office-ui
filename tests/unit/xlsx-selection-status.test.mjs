import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDemo = createRequire(
  new URL("../../apps/demo/package.json", import.meta.url),
);
const { formatSelectionStatValue } = await import(
  pathToFileURL(requireFromDemo.resolve("@arcships/vue-xlsx")).href
);

function format(code) {
  return { formatType: "custom", formatString: code };
}

test("General totals stay compact and ungrouped, like Excel's status bar", () => {
  assert.equal(formatSelectionStatValue(55, null, "en"), "55");
  assert.equal(formatSelectionStatValue(5.5, null, "en"), "5.5");
  assert.equal(formatSelectionStatValue(1 / 3, null, "en"), "0.3333333333");
  assert.equal(formatSelectionStatValue(-0, null, "en"), "0");
  assert.equal(formatSelectionStatValue(1234567, { formatType: "general" }, "en"), "1234567");
  assert.equal(formatSelectionStatValue(null, null, "en"), "");
  assert.equal(formatSelectionStatValue(Number.NaN, null, "en"), "");
});

test("number formats supply decimals and grouping", () => {
  assert.equal(formatSelectionStatValue(1234.56, format("#,##0.00"), "en"), "1,234.56");
  assert.equal(formatSelectionStatValue(1234.56, format("#,##0"), "en"), "1,235");
  assert.equal(formatSelectionStatValue(1234.56, format("0.000"), "en"), "1234.560");
  assert.equal(formatSelectionStatValue(1234.56, format("#,##0.00"), "de-DE"), "1.234,56");
});

test("currency formats keep their symbol on the side Excel puts it", () => {
  assert.equal(formatSelectionStatValue(1234.56, format("$#,##0.00"), "en"), "$1,234.56");
  assert.equal(formatSelectionStatValue(-1234.56, format("$#,##0.00"), "en"), "-$1,234.56");
  assert.equal(
    formatSelectionStatValue(1234.56, format("[$€-407]#,##0.00"), "en"),
    "€1.234,56",
  );
  assert.equal(
    formatSelectionStatValue(1234.56, format("#,##0.00 [$€-407]"), "en"),
    "1.234,56€",
  );
  assert.equal(
    formatSelectionStatValue(1234.56, format('[$$-409]#,##0.00'), "en"),
    "$1,234.56",
  );
});

test("percentage totals use the percent scale", () => {
  assert.equal(formatSelectionStatValue(0.75, format("0.00%"), "en"), "75.00%");
  assert.equal(formatSelectionStatValue(0.375, format("0%"), "en"), "38%");
  assert.equal(formatSelectionStatValue(0.75, format("0.00%"), "de-DE"), "75,00%");
});

test("scientific formats keep the exponent width", () => {
  assert.equal(formatSelectionStatValue(1234.5, format("0.00E+00"), "en"), "1.23E+03");
  assert.equal(formatSelectionStatValue(0.00045, format("0.00E+00"), "en"), "4.50E-04");
});

test("date, time and text formats fall back to General", () => {
  assert.equal(formatSelectionStatValue(45306, format("m/d/yyyy"), "en"), "45306");
  assert.equal(formatSelectionStatValue(45306, format("h:mm"), "en"), "45306");
  assert.equal(formatSelectionStatValue(45306, format("@"), "en"), "45306");
});
