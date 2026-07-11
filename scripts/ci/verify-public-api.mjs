#!/usr/bin/env node

import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const contractPath = path.join(import.meta.dirname, "public-api-contract.json");
const output = path.resolve(
  process.env.PUBLIC_API_EVIDENCE_DIR ||
    path.join(root, "output", "acceptance", "public-api"),
);
const probe = process.env.PUBLIC_API_PROBE || "";

rmSync(output, { recursive: true, force: true });
mkdirSync(output, { recursive: true });

const contract = JSON.parse(readFileSync(contractPath, "utf8"));
const failures = [];
const packageChecks = [];
const publicDocumentationPath = path.join(root, "docs", "api", "public-api-contract.md");
let typescript;
try {
  const require = createRequire(path.join(root, "packages", "docx-core", "package.json"));
  typescript = require("typescript");
} catch (error) {
  failures.push(`TypeScript checker is unavailable: ${error instanceof Error ? error.message : String(error)}`);
}

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.keys(value).sort().map((key) => [key, stable(value[key])]),
    );
  }
  return value;
}

function collectExportTargets(value, result = []) {
  if (typeof value === "string") result.push(value);
  else if (value && typeof value === "object") {
    for (const child of Object.values(value)) collectExportTargets(child, result);
  }
  return result;
}

function walk(directory) {
  if (!existsSync(directory)) return [];
  const result = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) result.push(...walk(absolute));
    else result.push(absolute);
  }
  return result;
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function findDeprecatedDeclaration(files, symbol) {
  const symbolPattern = new RegExp(`\\b${escapeRegExp(symbol)}\\b`);
  for (const file of files) {
    const text = readFileSync(file, "utf8");
    const comments = text.matchAll(/\/\*\*[\s\S]*?\*\//g);
    for (const match of comments) {
      if (!match[0].includes("@deprecated")) continue;
      const start = (match.index ?? 0) + match[0].length;
      const suffix = text.slice(start, start + 1_200);
      const nextComment = suffix.indexOf("/**");
      const semicolon = suffix.indexOf(";");
      const limits = [suffix.length];
      if (nextComment >= 0) limits.push(nextComment);
      if (semicolon >= 0) limits.push(semicolon + 1);
      const declaration = suffix.slice(0, Math.min(...limits));
      if (symbolPattern.test(declaration)) {
        return {
          file: path.relative(root, file),
          declaration: declaration.trim().replace(/\s+/g, " ").slice(0, 500),
          annotation: match[0].replace(/\s+/g, " ").slice(0, 500),
        };
      }
    }
  }
  return null;
}

function declarationTarget(fromFile, specifier) {
  const absolute = path.resolve(path.dirname(fromFile), specifier);
  const candidates = [
    absolute,
    absolute.replace(/\.(?:mjs|cjs|js)$/, ".d.ts"),
    `${absolute}.d.ts`,
    path.join(absolute, "index.d.ts"),
  ];
  return candidates.find((candidate) => candidate.endsWith(".d.ts") && existsSync(candidate));
}

function validateDeclarationGraph(entry) {
  const missing = [];
  const visited = new Set();
  const queue = [entry];
  while (queue.length) {
    const file = queue.shift();
    if (!file || visited.has(file)) continue;
    visited.add(file);
    const text = readFileSync(file, "utf8");
    const specifiers = [
      ...text.matchAll(/(?:from\s+|import\s*)["']([^"']+)["']/g),
    ].map((match) => match[1]);
    for (const specifier of specifiers) {
      if (!specifier.startsWith(".")) continue;
      const target = declarationTarget(file, specifier);
      if (!target) {
        missing.push({ from: path.relative(root, file), specifier });
      } else {
        queue.push(target);
      }
    }
  }
  return {
    files: [...visited].map((file) => path.relative(root, file)).sort(),
    missing,
  };
}

function collectDeclarationExports(text) {
  const names = new Set();
  for (const match of text.matchAll(
    /export\s+(?:declare\s+)?(?:abstract\s+)?(?:class|interface|type|function|const|let|var|enum)\s+([A-Za-z_$][\w$]*)/g,
  )) {
    names.add(match[1]);
  }
  for (const match of text.matchAll(/export\s+(?:type\s+)?\{([\s\S]*?)\}\s*from\s*["']/g)) {
    for (const item of match[1].split(",")) {
      const cleaned = item.trim().replace(/^type\s+/, "");
      if (!cleaned) continue;
      const alias = cleaned.match(/\bas\s+([A-Za-z_$][\w$]*)$/);
      if (alias) names.add(alias[1]);
      else {
        const direct = cleaned.match(/^([A-Za-z_$][\w$]*)$/);
        if (direct) names.add(direct[1]);
      }
    }
  }
  return names;
}

function rootDeclarationFingerprint(entry) {
  if (!typescript) return null;
  const options = {
    module: typescript.ModuleKind.ESNext,
    moduleResolution: typescript.ModuleResolutionKind.Bundler,
    target: typescript.ScriptTarget.ES2022,
    skipLibCheck: true,
    allowArbitraryExtensions: true,
    noEmit: true,
  };
  const program = typescript.createProgram([entry], options);
  const source = program.getSourceFile(entry);
  const checker = program.getTypeChecker();
  const moduleSymbol = source ? checker.getSymbolAtLocation(source) : undefined;
  const names = moduleSymbol
    ? checker.getExportsOfModule(moduleSymbol).map((symbol) => symbol.getName()).sort()
    : [];
  const diagnostics = typescript.getPreEmitDiagnostics(program).map((diagnostic) => ({
    code: diagnostic.code,
    message: typescript.flattenDiagnosticMessageText(diagnostic.messageText, " "),
  }));
  return {
    count: names.length,
    sha256: createHash("sha256").update(names.join("\n")).digest("hex"),
    names,
    diagnostics,
  };
}

for (const expected of contract.packages) {
  const packageDir = path.join(root, expected.directory);
  const manifestPath = path.join(packageDir, "package.json");
  const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
  const exportMatches =
    JSON.stringify(stable(manifest.exports)) === JSON.stringify(stable(expected.exports));
  if (manifest.name !== expected.name) {
    failures.push(`${expected.directory}: expected package name ${expected.name}, got ${manifest.name}`);
  }
  if (!exportMatches) {
    failures.push(`${expected.name}: exports do not exactly match the public contract`);
  }

  const targets = [...new Set(collectExportTargets(expected.exports))];
  const missingTargets = targets.filter(
    (target) => !existsSync(path.join(packageDir, target.replace(/^\.\//, ""))),
  );
  for (const target of missingTargets) {
    failures.push(`${expected.name}: exported target is missing: ${target}`);
  }

  const declarationFiles = walk(path.join(packageDir, "dist"))
    .filter((file) => file.endsWith(".d.ts"));
  if (!declarationFiles.length) {
    failures.push(`${expected.name}: no built declarations found; run the formal build first`);
  }
  const declarationMaps = walk(path.join(packageDir, "dist"))
    .filter((file) => file.endsWith(".d.ts.map"));
  const leakPatterns = [
    { id: "private-office-runtime", pattern: /@arcships\/office-runtime/ },
    { id: "workspace-absolute-path", pattern: new RegExp(escapeRegExp(root)) },
    { id: "user-absolute-path", pattern: /(?:file:\/\/|\/Users\/|\/home\/|[A-Za-z]:[\\/])/ },
    { id: "source-deep-import", pattern: /@arcships\/[A-Za-z0-9_-]+\/(?:src|packages)\// },
  ];
  const declarationLeaks = [];
  for (const file of [...declarationFiles, ...declarationMaps]) {
    const text = readFileSync(file, "utf8");
    for (const leak of leakPatterns) {
      if (leak.pattern.test(text)) {
        declarationLeaks.push({ file: path.relative(root, file), kind: leak.id });
      }
    }
  }
  for (const leak of declarationLeaks) {
    failures.push(`${expected.name}: declaration leak ${leak.kind} in ${leak.file}`);
  }

  const deprecated = Object.fromEntries(
    expected.deprecatedSymbols.map((symbol) => [
      symbol,
      findDeprecatedDeclaration(declarationFiles, symbol),
    ]),
  );
  for (const [symbol, declaration] of Object.entries(deprecated)) {
    if (!declaration) {
      failures.push(`${expected.name}: built declaration for ${symbol} is missing @deprecated`);
    }
  }

  const fingerprint = rootDeclarationFingerprint(path.join(packageDir, "dist", "index.d.ts"));
  const expectedFingerprint = expected.rootDeclarationFingerprint;
  if (fingerprint?.diagnostics.length) {
    for (const diagnostic of fingerprint.diagnostics) {
      failures.push(`${expected.name}: declaration checker TS${diagnostic.code}: ${diagnostic.message}`);
    }
  }
  if (
    !fingerprint ||
    fingerprint.count !== expectedFingerprint.count ||
    fingerprint.sha256 !== expectedFingerprint.sha256
  ) {
    failures.push(
      `${expected.name}: root declaration export fingerprint changed; expected ` +
      `${expectedFingerprint.count}/${expectedFingerprint.sha256}, got ` +
      `${fingerprint?.count ?? "unavailable"}/${fingerprint?.sha256 ?? "unavailable"}`,
    );
  }

  packageChecks.push({
    package: expected.name,
    exportMatches,
    exportTargets: targets,
    missingTargets,
    declarationFiles: declarationFiles.map((file) => path.relative(root, file)),
    declarationLeaks,
    deprecated,
    rootDeclarationFingerprint: fingerprint,
    expectedRootDeclarationFingerprint: expectedFingerprint,
  });
}

const requiredDocumentationTerms = [
  ...contract.packages.map((entry) => entry.name),
  "0.2.0",
  "1.0.0",
  "files-rejected",
  "RUNTIME_DISPOSED",
  "ERR_PACKAGE_PATH_NOT_EXPORTED",
];
const publicDocumentation = {
  path: path.relative(root, publicDocumentationPath),
  exists: existsSync(publicDocumentationPath),
  requiredTerms: {},
};
if (!publicDocumentation.exists) {
  failures.push(`public API documentation is missing: ${publicDocumentation.path}`);
} else {
  const documentationText = readFileSync(publicDocumentationPath, "utf8");
  publicDocumentation.requiredTerms = Object.fromEntries(
    requiredDocumentationTerms.map((term) => [term, documentationText.includes(term)]),
  );
  for (const [term, present] of Object.entries(publicDocumentation.requiredTerms)) {
    if (!present) failures.push(`public API documentation is missing required term: ${term}`);
  }
}

const runtimeEntry = path.join(root, "packages", "docx-core", "dist", "runtime.d.ts");
let docxRuntimeCheck = {
  entry: path.relative(root, runtimeEntry),
  symbols: {},
  graph: { files: [], missing: [] },
};
if (!existsSync(runtimeEntry)) {
  failures.push("@arcships/docx-core/runtime: built declaration entry is missing");
} else {
  const runtimeText = readFileSync(runtimeEntry, "utf8");
  const runtimeExports = collectDeclarationExports(runtimeText);
  docxRuntimeCheck.symbols = Object.fromEntries(
    contract.docxRuntimeSymbols.map((symbol) => [
      symbol,
      runtimeExports.has(symbol),
    ]),
  );
  for (const [symbol, present] of Object.entries(docxRuntimeCheck.symbols)) {
    if (!present) {
      failures.push(`@arcships/docx-core/runtime: public signature type ${symbol} is not exported`);
    }
  }
  docxRuntimeCheck.graph = validateDeclarationGraph(runtimeEntry);
  for (const missing of docxRuntimeCheck.graph.missing) {
    failures.push(
      `@arcships/docx-core/runtime: declaration ${missing.from} cannot resolve ${missing.specifier}`,
    );
  }
}

if (probe === "missing-export") {
  failures.push("probe: simulated missing public export target");
} else if (probe === "private-type") {
  failures.push("probe: simulated @arcships/office-runtime declaration leak");
} else if (probe === "missing-deprecated") {
  failures.push("probe: simulated missing @deprecated declaration");
} else if (probe === "runtime-type") {
  failures.push("probe: simulated non-self-contained DOCX runtime declaration");
} else if (probe) {
  failures.push(`unknown PUBLIC_API_PROBE: ${probe}`);
}

const summary = {
  suite: "P2-PUBLIC-API-CONTRACT",
  result: failures.length === 0 ? "PASS" : "FAIL",
  probe: probe || null,
  contract: path.relative(root, contractPath),
  packageChecks,
  publicDocumentation,
  docxRuntimeCheck,
  positiveSpecifiers: contract.positiveSpecifiers,
  executableSpecifiers: contract.executableSpecifiers,
  deepImportRejections: contract.deepImportRejections,
  failures,
};
writeFileSync(path.join(output, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);

console.log(`${summary.result}: public API contract checked at ${output}`);
for (const failure of failures) console.error(`- ${failure}`);
process.exit(summary.result === "PASS" ? 0 : 1);
