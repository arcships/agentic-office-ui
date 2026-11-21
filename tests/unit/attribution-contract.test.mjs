import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const verifier = path.join(root, "scripts/ci/verify-attribution.mjs");

const COMMIT = "70fe70c6f1e1eee53954c6923f3d9f583557a57e";
const HEADER =
  "// Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.";
const MIT_TEXT = `MIT License

Copyright (c) 2026 CrowdView Inc, dba Extend
Portions Copyright (c) 2023 shadcn

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.
`;

function write(file, content) {
  mkdirSync(path.dirname(file), { recursive: true });
  writeFileSync(file, content);
}

function upstreamLicense(id) {
  switch (id) {
    case "extend-ui":
      return {
        license: "MIT",
        copyright: "Copyright (c) 2026 CrowdView Inc, dba Extend; Portions Copyright (c) 2023 shadcn",
        file: "docs/upstream/licenses/extend-ui-MIT.txt",
        markers: ["MIT License", "Copyright (c) 2026 CrowdView Inc, dba Extend", "Portions Copyright (c) 2023 shadcn"],
      };
    case "shadcn-vue":
      return {
        license: "MIT",
        copyright: "Copyright (c) 2023 unovue",
        file: "docs/upstream/licenses/shadcn-vue-MIT.txt",
        markers: ["MIT License", "Copyright (c) 2023 unovue"],
      };
    case "reka-ui":
      return {
        license: "MIT",
        copyright: "Copyright (c) 2023 radix-vue",
        file: "docs/upstream/licenses/reka-ui-MIT.txt",
        markers: ["MIT License", "Copyright (c) 2023 radix-vue"],
      };
    case "lucide":
      return {
        license: "ISC",
        copyright: "Copyright (c) 2026 Lucide Icons and Contributors",
        file: "docs/upstream/licenses/lucide-ISC.txt",
        markers: ["ISC License", "Copyright (c) 2026 Lucide Icons and Contributors"],
      };
    case "geist":
      return {
        license: "SIL Open Font License 1.1",
        copyright: "Copyright (c) 2023 Vercel, in collaboration with basement.studio",
        file: "docs/upstream/licenses/geist-OFL-1.1.txt",
        markers: [
          "SIL OPEN FONT LICENSE Version 1.1",
          "Copyright (c) 2023 Vercel, in collaboration with basement.studio",
        ],
      };
    default:
      throw new Error(`unknown upstream ${id}`);
  }
}

function manifestFixture() {
  const upstreams = ["extend-ui", "shadcn-vue", "reka-ui", "lucide", "geist"].map((id) => {
    const license = upstreamLicense(id);
    return {
      id,
      name: `Fixture ${id}`,
      aliases: [id],
      license: license.license,
      copyright: license.copyright,
      licenseFile: license.file,
      licenseMarkers: license.markers,
      sourceUrl: `https://example.invalid/${id}`,
      pinnedUrl: `https://example.invalid/${id}/tree/${COMMIT}`,
      licenseUrl: `https://example.invalid/${id}/LICENSE`,
      usage: `Fixture usage scope for ${id}.`,
    };
  });
  const index = [
    {
      upstreamFile: "apps/v4/components/ui/button.tsx",
      purpose: "Button primitive",
      target: "packages/vue-ui/src/components/ui/Button.vue",
      status: "adapted",
      issues: ["#157"],
      notes: "header required",
    },
    {
      upstreamFile: "apps/v4/components/extend/pdf-viewer.tsx",
      purpose: "PDF viewer chrome",
      target: "packages/vue-pdf/src/components/PdfViewer.vue",
      status: "reference-only",
      issues: ["#162"],
      notes: "no header required",
    },
    {
      upstreamFile: "tools vendored snapshot",
      purpose: "Vendored reference copy",
      target: "tools/extend-reference/src/vendor/",
      status: "copied",
      issues: ["#155"],
      notes: "planned directory",
    },
    {
      upstreamFile: "apps/v4/lib/utils.ts",
      purpose: "Tailwind helper",
      target: "",
      status: "not-used",
      issues: ["#166"],
      notes: "not used",
    },
  ];
  return {
    schemaVersion: 1,
    extend: {
      repository: "https://example.invalid/extend-ui",
      commit: COMMIT,
      license: "MIT",
      copyright: "Copyright (c) 2026 CrowdView Inc, dba Extend; Portions Copyright (c) 2023 shadcn",
    },
    header: {
      required:
        "Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.",
      markers: [
        {
          name: "upstream and pinned commit",
          pattern: "extend-hq/ui\\s+@?\\s*70fe70c6f1e1eee53954c6923f3d9f583557a57e",
        },
        {
          name: "license identifier",
          pattern: "\\(MIT\\)|MIT License|SPDX-License-Identifier:\\s*MIT",
        },
        {
          name: "Extend copyright",
          pattern: "Copyright \\(c\\) 2026 CrowdView Inc, dba Extend",
        },
        {
          name: "shadcn portions",
          pattern: "portions Copyright \\(c\\) 2023 shadcn",
        },
      ],
      commentPrefixes: ["//", "/*", "*", "<!--", "--", "/*!"],
    },
    upstreams,
    index,
    urlChecks: [{ url: "https://example.invalid/never-fetched" }],
  };
}

function noticesFixture(manifest) {
  const sections = manifest.upstreams.map((upstream) => {
    const text = readFileSync(path.join(manifest.__root, upstream.licenseFile), "utf8");
    return [
      `### ${upstream.name}`,
      "",
      `License: ${upstream.license}`,
      `Copyright: ${upstream.copyright}`,
      `Usage scope: ${upstream.usage}`,
      `Source: ${upstream.sourceUrl}`,
      `License copy: ${upstream.licenseFile}`,
      "",
      text,
    ].join("\n");
  });
  return `# Third-Party Notices\n\n${sections.join("\n\n")}\n`;
}

function indexDocFixture(manifest) {
  const rows = manifest.index
    .map(
      (entry) =>
        `| \`${entry.upstreamFile}\` | ${entry.purpose} | ${entry.target ? `\`${entry.target}\`` : "—"} | ${entry.status} | ${entry.notes} |`,
    )
    .join("\n");
  return `# Upstream index

Pinned commit: \`${manifest.extend.commit}\`.
Repository: ${manifest.extend.repository}
License: ${manifest.upstreams[0].licenseUrl}

\`\`\`text
${manifest.header.required}
\`\`\`

| Extend 文件 | 用途 | 我们的目标文件 | 状态 | 备注 |
|---|---|---|---|---|
${rows}
`;
}

function createFixture() {
  const fixtureRoot = mkdtempSync(path.join(tmpdir(), "attribution-fixture-"));
  write(
    path.join(fixtureRoot, "LICENSE"),
    "                                 Apache License\n                           Version 2.0, January 2004\n",
  );
  write(
    path.join(fixtureRoot, "README.md"),
    "# Fixture\n\n## Open-source attribution\n\nSee [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).\n",
  );
  write(
    path.join(fixtureRoot, "packages/vue-ui/README.md"),
    "# Fixture\n\n## Open-source attribution\n\nSee [THIRD_PARTY_NOTICES.md](../../THIRD_PARTY_NOTICES.md).\n",
  );
  write(
    path.join(fixtureRoot, "docs/INDEX.md"),
    "# Docs\n\n- [upstream/extend-ui.md](upstream/extend-ui.md)\n",
  );

  const manifest = manifestFixture();
  for (const upstream of manifest.upstreams) {
    const license = upstreamLicense(upstream.id);
    write(path.join(fixtureRoot, license.file), upstream.id === "extend-ui" ? MIT_TEXT : `${license.markers.join("\n")}\n`);
  }

  write(path.join(fixtureRoot, "THIRD_PARTY_NOTICES.md"), noticesFixture({ ...manifest, __root: fixtureRoot }));
  write(path.join(fixtureRoot, "docs/upstream/attribution.json"), `${JSON.stringify(manifest, null, 2)}\n`);
  write(path.join(fixtureRoot, "docs/upstream/extend-ui.md"), indexDocFixture(manifest));

  const adapted = path.join(fixtureRoot, "packages/vue-ui/src/components/ui/Button.vue");
  write(adapted, `${HEADER}\n<template><button><slot /></button></template>\n`);
  write(
    path.join(fixtureRoot, "packages/vue-pdf/src/components/UnindexedCopy.vue"),
    "<template><div /></template>\n",
  );
  return fixtureRoot;
}

function runVerifier(fixtureRoot) {
  return spawnSync(process.execPath, [verifier], {
    cwd: root,
    encoding: "utf8",
    env: {
      ...process.env,
      ATTRIBUTION_ROOT: fixtureRoot,
      ATTRIBUTION_SKIP_URLS: "1",
      ATTRIBUTION_EVIDENCE_DIR: path.join(fixtureRoot, "output"),
    },
  });
}

test("attribution verifier accepts a complete fixture and reports pending targets", () => {
  const fixtureRoot = createFixture();
  try {
    const result = runVerifier(fixtureRoot);
    assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
    const summary = JSON.parse(
      readFileSync(path.join(fixtureRoot, "output/summary.json"), "utf8"),
    );
    assert.equal(summary.result, "PASS");
    assert.deepEqual(summary.headerChecked, ["packages/vue-ui/src/components/ui/Button.vue"]);
    assert.deepEqual(summary.pendingTargets, [
      { target: "tools/extend-reference/src/vendor/", status: "copied", issues: ["#155"] },
    ]);
    assert.equal(summary.urlChecksSkipped, true);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

test("attribution verifier accepts the vendored-copy header variant", () => {
  const fixtureRoot = createFixture();
  try {
    write(
      path.join(fixtureRoot, "packages/vue-ui/src/components/ui/Button.vue"),
      `/**\n * Vendored from extend-hq/ui @ ${COMMIT}\n * Source: apps/v4/components/ui/button.tsx\n * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.\n * Licensed under the MIT License.\n */\n<template><button><slot /></button></template>\n`,
    );
    const result = runVerifier(fixtureRoot);
    assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

test("attribution verifier rejects a header that omits the license identifier", () => {
  const fixtureRoot = createFixture();
  try {
    write(
      path.join(fixtureRoot, "packages/vue-ui/src/components/ui/Button.vue"),
      `/**\n * Vendored from extend-hq/ui @ ${COMMIT}\n * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.\n */\n<template><button><slot /></button></template>\n`,
    );
    const result = runVerifier(fixtureRoot);
    assert.equal(result.status, 1, result.stdout);
    assert.match(result.stderr, /attribution header packages\/vue-ui\/src\/components\/ui\/Button\.vue/);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

test("attribution verifier requires the MIT notice next to a vendored snapshot", () => {
  const fixtureRoot = createFixture();
  try {
    write(
      path.join(fixtureRoot, "tools/extend-reference/src/vendor/ui/button.tsx"),
      `/**\n * Vendored from extend-hq/ui @ ${COMMIT}\n * Source: apps/v4/components/ui/button.tsx\n * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.\n * Licensed under the MIT License.\n */\nexport const button = "button";\n`,
    );
    const withoutNotice = runVerifier(fixtureRoot);
    assert.equal(withoutNotice.status, 1, withoutNotice.stdout);
    assert.match(withoutNotice.stderr, /vendored reference snapshot keeps the Extend MIT notice/);

    write(
      path.join(fixtureRoot, "tools/extend-reference/src/vendor/NOTICE.md"),
      `# Notice\n\nExtend UI is MIT licensed. Copyright (c) 2026 CrowdView Inc, dba Extend.\nSee docs/upstream/licenses/extend-ui-MIT.txt.\n`,
    );
    const withNotice = runVerifier(fixtureRoot);
    assert.equal(withNotice.status, 0, `${withNotice.stdout}\n${withNotice.stderr}`);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

test("attribution verifier fails when an existing adapted file has no header", () => {
  const fixtureRoot = createFixture();
  try {
    const adapted = path.join(fixtureRoot, "packages/vue-ui/src/components/ui/Button.vue");
    write(adapted, "<template><button><slot /></button></template>\n");
    const result = runVerifier(fixtureRoot);
    assert.equal(result.status, 1, result.stdout);
    assert.match(result.stderr, /attribution header packages\/vue-ui\/src\/components\/ui\/Button\.vue/);
    assert.match(result.stderr, /Adapted from extend-hq\/ui/);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

test("attribution verifier fails when a header-bearing file is missing from the index", () => {
  const fixtureRoot = createFixture();
  try {
    write(
      path.join(fixtureRoot, "packages/vue-pdf/src/components/UnindexedHeaderCopy.vue"),
      `${HEADER}\n<template><div /></template>\n`,
    );
    const result = runVerifier(fixtureRoot);
    assert.equal(result.status, 1, result.stdout);
    assert.match(result.stderr, /every file with an Extend header is in the index/);
    assert.match(result.stderr, /packages\/vue-pdf\/src\/components\/UnindexedHeaderCopy\.vue/);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

test("attribution verifier fails when the notice or license copy drifts", () => {
  const fixtureRoot = createFixture();
  try {
    rmSync(path.join(fixtureRoot, "docs/upstream/licenses/reka-ui-MIT.txt"));
    const result = runVerifier(fixtureRoot);
    assert.equal(result.status, 1, result.stdout);
    assert.match(result.stderr, /license file reka-ui/);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

test("attribution verifier fails when an index row is out of sync with the manifest", () => {
  const fixtureRoot = createFixture();
  try {
    const docPath = path.join(fixtureRoot, "docs/upstream/extend-ui.md");
    const doc = readFileSync(docPath, "utf8").replace("| adapted |", "| reference-only |");
    writeFileSync(docPath, doc);
    const result = runVerifier(fixtureRoot);
    assert.equal(result.status, 1, result.stdout);
    assert.match(result.stderr, /index doc row apps\/v4\/components\/ui\/button\.tsx/);
    assert.ok(existsSync(path.join(fixtureRoot, "docs/upstream/attribution.json")));
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});
