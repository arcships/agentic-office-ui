#!/usr/bin/env node

import { createHash } from "node:crypto";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const failures = [];

function read(relative) {
  return readFileSync(path.join(root, relative), "utf8");
}

function readJson(relative) {
  return JSON.parse(read(relative));
}

function check(condition, message) {
  if (!condition) failures.push(message);
}

function sha256(relative) {
  return createHash("sha256").update(readFileSync(path.join(root, relative))).digest("hex");
}

function sourceFiles(relative) {
  const directory = path.join(root, relative);
  const result = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const child = path.join(relative, entry.name);
    if (entry.isDirectory()) result.push(...sourceFiles(child));
    else if (/\.(?:ts|vue)$/u.test(entry.name)) result.push(child);
  }
  return result;
}

const rootPackage = readJson("package.json");
const corePackage = readJson("packages/pptx-core/package.json");
const vuePackage = readJson("packages/vue-pptx/package.json");
const labPackage = readJson("apps/pptx-playback-lab/package.json");

check(corePackage.private === true, "@arcships/pptx-core 必须保持 private，直到完成正式发布验收。");
check(vuePackage.private === true, "@arcships/vue-pptx 必须保持 private，直到完成正式发布验收。");
check(labPackage.private === true, "pptx-playback-lab 必须保持 private。");
check(
  corePackage.dependencies?.["@aiden0z/pptx-renderer"] === "1.2.4",
  "pptx-core 必须固定 @aiden0z/pptx-renderer@1.2.4。",
);
check(
  labPackage.dependencies?.["@aiden0z/pptx-renderer"] === "1.2.4",
  "探索台必须和正式包使用相同的固定渲染器版本。",
);
const rendererPatch = rootPackage.pnpm?.patchedDependencies?.["@aiden0z/pptx-renderer@1.2.4"];
check(
  rendererPatch === "patches/@aiden0z__pptx-renderer@1.2.4.patch",
  "根 package.json 必须登记固定渲染器补丁。",
);
check(
  existsSync(path.join(root, rendererPatch ?? "")),
  "固定渲染器补丁文件不存在。",
);

for (const script of [
  "dev:pptx-lab",
  "build:pptx",
  "typecheck:pptx",
  "test:pptx",
  "capture:pptx-baselines",
]) {
  check(typeof rootPackage.scripts?.[script] === "string", `根 package.json 缺少 ${script}。`);
}
for (const script of ["build:lab", "typecheck"]) {
  check(typeof labPackage.scripts?.[script] === "string", `探索台缺少 ${script}。`);
}

for (const relative of [
  "docs/pptx-playback-implementation-design.md",
  "docs/pptx-playback-api-design.md",
  "docs/pptx-playback-compatibility-and-acceptance.md",
  "docs/pptx-playback-roadmap.md",
  "docs/pptx-development-guide.md",
  "packages/pptx-core/README.md",
  "packages/vue-pptx/README.md",
  "tests/fixtures/pptx/playback/README.md",
  "tests/fixtures/pptx/playback/manifest.json",
]) {
  check(existsSync(path.join(root, relative)), `缺少 PPTX 开发文件：${relative}`);
}

const expectedFixtures = [
  "entrance-exit-fade.pptx",
  "hidden-and-internal-link.pptx",
  "media-audio-video.pptx",
  "media-bookmark.pptx",
  "morph-explicit-name.pptx",
  "motion-straight.pptx",
  "overlap-same-property.pptx",
  "text-by-paragraph.pptx",
  "timing-click-with-after.pptx",
  "timing-delay-repeat-reverse.pptx",
  "transition-cut-fade-push-wipe.pptx",
  "trigger-shape-click.pptx",
  "wipe-scale-rotate.pptx",
];
const fixtureManifest = readJson("tests/fixtures/pptx/playback/manifest.json");
const entries = Array.isArray(fixtureManifest.entries) ? fixtureManifest.entries : [];
const fixtureNames = entries.map((entry) => entry.file).sort();
check(fixtureManifest.version === 1, "PPTX 素材清单版本必须为 1。");
check(
  JSON.stringify(fixtureNames) === JSON.stringify(expectedFixtures),
  "PPTX 最小素材清单与兼容验收设计不一致。",
);
check(new Set(fixtureNames).size === fixtureNames.length, "PPTX 素材清单存在重复文件名。");
for (const entry of entries) {
  const fixtureRelative = `tests/fixtures/pptx/playback/${entry.file}`;
  const stem = path.parse(entry.file).name;
  const fixtureAbsolute = path.join(root, fixtureRelative);
  check(entry.status === "ready", `${entry.file} 必须是 ready，正式验收不接受计划素材。`);
  check(typeof entry.source === "string" && entry.source.length > 0, `${entry.file} 缺少来源。`);
  check(typeof entry.license === "string" && entry.license.length > 0, `${entry.file} 缺少许可。`);
  check(/^[a-f0-9]{64}$/u.test(entry.sha256 ?? ""), `${entry.file} 缺少有效 SHA-256。`);
  check(
    typeof entry.powerpointVersion === "string" && entry.powerpointVersion.length > 0,
    `${entry.file} 缺少 PowerPoint 版本。`,
  );
  check(existsSync(fixtureAbsolute), `缺少正式素材：${entry.file}`);
  if (existsSync(fixtureAbsolute)) {
    const prefix = readFileSync(fixtureAbsolute).subarray(0, 4);
    check(prefix.equals(Buffer.from([0x50, 0x4b, 0x03, 0x04])), `${entry.file} 不是有效的 PPTX 压缩包。`);
    check(sha256(fixtureRelative) === entry.sha256, `${entry.file} 的 SHA-256 与清单不一致。`);
  }

  const baselineRelative = `tests/fixtures/pptx/playback/${entry.baseline ?? ""}`;
  const fixtureReadme = `tests/fixtures/pptx/playback/${stem}/README.md`;
  check(entry.baseline === `${stem}/expected-events.json`, `${entry.file} 的事件基准路径无效。`);
  check(existsSync(path.join(root, baselineRelative)), `${entry.file} 缺少事件基准。`);
  check(existsSync(path.join(root, fixtureReadme)), `${entry.file} 缺少基准说明。`);
  if (existsSync(path.join(root, baselineRelative))) {
    const baseline = readJson(baselineRelative);
    const capability = baseline.capability ?? {};
    check(baseline.schemaVersion === 1, `${entry.file} 的事件基准版本无效。`);
    check(baseline.fixture === entry.file, `${entry.file} 的事件基准文件名不一致。`);
    check(
      baseline.powerPointVersion === entry.powerpointVersion,
      `${entry.file} 的事件基准 PowerPoint 版本不一致。`,
    );
    check(
      baseline.focusSlideIndex === entry.focusSlideIndex,
      `${entry.file} 的事件基准焦点页不一致。`,
    );
    check(
      Number.isInteger(capability.discovered)
        && capability.discovered
          === (capability.strict ?? 0)
            + (capability.approximate ?? 0)
            + (capability.static ?? 0)
            + (capability.unparsed ?? 0),
      `${entry.file} 的能力统计不一致。`,
    );
    check(Array.isArray(baseline.eventSequence), `${entry.file} 缺少稳定事件序列。`);
    check(Array.isArray(baseline.schedule), `${entry.file} 缺少时间安排基准。`);
  }

}

const formalSource = [
  ...sourceFiles("packages/pptx-core/src"),
  ...sourceFiles("packages/vue-pptx/src"),
];
for (const relative of formalSource) {
  check(!read(relative).includes("pptx-playback-lab"), `${relative} 不得导入探索台代码。`);
}
for (const relative of sourceFiles("apps/pptx-playback-lab/src")) {
  check(!/@arcships\/(?:pptx-core|vue-pptx)/u.test(read(relative)), `${relative} 不得依赖正式 PPTX 包。`);
}

const coreRoot = `${read("packages/pptx-core/src/index.ts")}\n${read("packages/pptx-core/src/types.ts")}`;
check(!coreRoot.includes("@aiden0z/pptx-renderer"), "pptx-core 根入口不得导入第三方渲染器。");
check(!/\b(?:window|document|DOMParser|Worker)\b/u.test(coreRoot), "pptx-core 根入口不得访问浏览器对象。");

if (failures.length) {
  console.error(`FAIL: PPTX 开发准备检查发现 ${failures.length} 个问题。`);
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log(`PASS: PPTX 开发准备检查通过，正式源码 ${formalSource.length} 个，正式素材 ${entries.length} 个。`);
