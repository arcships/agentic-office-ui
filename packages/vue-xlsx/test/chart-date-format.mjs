import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

const requireFromPackage = createRequire(new URL("../package.json", import.meta.url));
const ts = requireFromPackage("typescript");

function transpile(fileUrl) {
  const source = readFileSync(fileUrl, "utf8");
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
}

const moduleUrl = (() => {
  const emitted = transpile(new URL("../src/render/chart-date.ts", import.meta.url));
  return `data:text/javascript;base64,${Buffer.from(emitted).toString("base64")}`;
})();

const {
  excelSerialToDate,
  formatExcelDateSerial,
  formatExcelSerial,
  isLikelyDateFormatCode,
} = await import(moduleUrl);

function runInTimeZone(timeZone) {
  const script = [
    `const mod = await import(${JSON.stringify(moduleUrl)});`,
    "console.log(JSON.stringify({",
    "  localized: mod.formatExcelSerial(45306, 'm/d/yyyy', { locale: 'en-US' }),",
    "  defaultLocale: mod.formatExcelSerial(45306, 'm/d/yyyy'),",
    "}));",
  ].join("\n");
  const result = spawnSync(process.execPath, ["--input-type=module", "-e", script], {
    encoding: "utf8",
    env: { ...process.env, TZ: timeZone },
  });
  assert.equal(result.status, 0, result.stderr);
  return JSON.parse(result.stdout.trim());
}

test("chart date serials format in UTC instead of the viewer timezone", () => {
  const losAngeles = runInTimeZone("America/Los_Angeles");
  const shanghai = runInTimeZone("Asia/Shanghai");

  assert.equal(losAngeles.localized, "1/15/2024");
  assert.equal(shanghai.localized, "1/15/2024");
  assert.equal(losAngeles.defaultLocale, shanghai.defaultLocale);
});

test("time-only chart formats render time instead of an 1899 date", () => {
  assert.equal(formatExcelSerial(0.5, "h:mm", { locale: "en-US" }), "12:00");
  assert.equal(formatExcelSerial(0.5, "h:mm AM/PM", { locale: "en-US" }), "12:00 PM");
  assert.equal(formatExcelSerial(0.5, "mm:ss", { locale: "en-US" }), "00:00");
});

test("elapsed-hour chart formats use total hours", () => {
  assert.equal(formatExcelSerial(1.25, "[h]:mm"), "30:00");
  assert.equal(formatExcelSerial(2.5, "[h]:mm:ss"), "60:00:00");
  assert.equal(formatExcelSerial(1.5, "[m]:ss"), "2160:00");
});

test("chart date formats honour the 1904 date system", () => {
  assert.equal(formatExcelSerial(1, "m/d/yyyy", { date1904: true }), "1/2/1904");
  assert.equal(excelSerialToDate(1, true).toISOString(), "1904-01-02T00:00:00.000Z");
  assert.equal(formatExcelDateSerial(45306, "m/d/yyyy", false), formatExcelSerial(45306, "m/d/yyyy"));
  assert.equal(excelSerialToDate(45306, false).toISOString(), "2024-01-15T00:00:00.000Z");
  assert.equal(excelSerialToDate(1, false).toISOString(), "1900-01-01T00:00:00.000Z");
  assert.equal(excelSerialToDate(59, false).toISOString(), "1900-02-28T00:00:00.000Z");
  assert.equal(excelSerialToDate(61, false).toISOString(), "1900-03-01T00:00:00.000Z");
});

test("date-like detection ignores quoted text and numeric formats", () => {
  assert.equal(isLikelyDateFormatCode("m/d/yyyy"), true);
  assert.equal(isLikelyDateFormatCode("h:mm"), true);
  assert.equal(isLikelyDateFormatCode("mm:ss"), true);
  assert.equal(isLikelyDateFormatCode("[h]:mm"), true);
  assert.equal(isLikelyDateFormatCode("0.00"), false);
  assert.equal(isLikelyDateFormatCode("0.00 \"m\""), false);
  assert.equal(isLikelyDateFormatCode("0.00 \\m"), false);
});
