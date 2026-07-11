import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";


const ROOT = path.resolve(import.meta.dirname, "../..");
const XLSX_GRID = path.join(
  ROOT,
  "packages/vue-xlsx/src/components/XlsxGrid.vue",
);
const HEAVY_FEATURE_PACKAGES = [
  "d3-geo",
  "d3-hierarchy",
  "d3-scale",
  "d3-shape",
  "regl",
  "topojson-client",
  "us-atlas",
  "world-atlas",
];

function scriptSource(file) {
  const source = readFileSync(file, "utf8");
  if (!file.endsWith(".vue")) return source;
  return [...source.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/g)]
    .map((match) => match[1])
    .join("\n");
}

function sourceFile(file) {
  return ts.createSourceFile(
    file,
    scriptSource(file),
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
}

function staticSpecifiers(file) {
  const result = [];
  const ast = sourceFile(file);
  for (const statement of ast.statements) {
    if (ts.isImportDeclaration(statement)) {
      if (statement.importClause?.isTypeOnly) continue;
      result.push(statement.moduleSpecifier.text);
    }
    if (ts.isExportDeclaration(statement) && statement.moduleSpecifier) {
      if (statement.isTypeOnly) continue;
      result.push(statement.moduleSpecifier.text);
    }
  }
  return result;
}

function resolveFile(candidate) {
  const candidates = [
    candidate,
    `${candidate}.ts`,
    `${candidate}.tsx`,
    `${candidate}.js`,
    `${candidate}.vue`,
    path.join(candidate, "index.ts"),
    path.join(candidate, "index.tsx"),
    path.join(candidate, "index.vue"),
  ];
  for (const item of candidates) {
    try {
      readFileSync(item);
      return path.normalize(item);
    } catch {
      // Try the next supported source extension.
    }
  }
  return null;
}

function resolveSpecifier(from, specifier) {
  const aliases = {
    "@arcships/vue-xlsx": path.join(ROOT, "packages/vue-xlsx/src/index.ts"),
    "@arcships/xlsx-core": path.join(ROOT, "packages/xlsx-core/src/index.ts"),
    "@arcships/office-runtime": path.join(ROOT, "packages/office-runtime/src/index.ts"),
  };
  if (aliases[specifier]) return aliases[specifier];
  if (!specifier.startsWith(".")) return null;
  return resolveFile(path.resolve(path.dirname(from), specifier));
}

function staticGraph(entry) {
  const pending = [path.normalize(entry)];
  const visited = new Set();
  const externals = [];
  while (pending.length) {
    const file = pending.pop();
    if (!file || visited.has(file)) continue;
    visited.add(file);
    for (const specifier of staticSpecifiers(file)) {
      const resolved = resolveSpecifier(file, specifier);
      if (resolved) pending.push(resolved);
      else externals.push({ from: path.relative(ROOT, file), specifier });
    }
  }
  return { visited, externals };
}

function descendants(node, predicate) {
  const matches = [];
  function visit(current) {
    if (predicate(current)) matches.push(current);
    ts.forEachChild(current, visit);
  }
  visit(node);
  return matches;
}

function declaration(ast, name) {
  return descendants(
    ast,
    (node) =>
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.name.text === name,
  )[0];
}

function functionNode(ast, name) {
  return descendants(
    ast,
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === name,
  )[0];
}

function startsAtZero(loop) {
  if (!ts.isForStatement(loop) || !ts.isVariableDeclarationList(loop.initializer)) {
    return false;
  }
  return loop.initializer.declarations.some(
    (item) => item.initializer?.kind === ts.SyntaxKind.NumericLiteral && item.initializer.text === "0",
  );
}

function zeroBasedLoops(node) {
  return descendants(node, (child) => ts.isForStatement(child) && startsAtZero(child));
}

test("P3-XLSX-LOAD-01 regression: default vue-xlsx entry excludes heavy optional features", () => {
  const graph = staticGraph(path.join(ROOT, "packages/vue-xlsx/src/index.ts"));
  const heavy = graph.externals.filter(({ specifier }) =>
    HEAVY_FEATURE_PACKAGES.some(
      (name) => specifier === name || specifier.startsWith(`${name}/`),
    ),
  );
  assert.deepEqual(
    heavy,
    [],
    `default entry statically reaches optional chart/map/WebGL packages:\n${JSON.stringify(heavy, null, 2)}`,
  );
});

test("P3-XLSX-LOAD-01 regression: Home does not statically load XLSX route code", () => {
  const graph = staticGraph(path.join(ROOT, "apps/demo/src/main.ts"));
  const xlsxPage = path.normalize(path.join(ROOT, "apps/demo/src/pages/XlsxViewerPage.vue"));
  assert.equal(
    graph.visited.has(xlsxPage),
    false,
    "Home entry statically imports XlsxViewerPage; route navigation cannot defer its code",
  );
});

test("P3-XLSX-SCROLL-01 regression: visible-range and offset lookup do not rescan from index zero", () => {
  const ast = sourceFile(XLSX_GRID);
  const targets = [
    declaration(ast, "visibleRowRange"),
    declaration(ast, "visibleColRange"),
    functionNode(ast, "getRowOffsetSum"),
    functionNode(ast, "getColOffsetSum"),
  ];
  assert.ok(targets.every(Boolean), "expected XlsxGrid range and offset functions were not found");
  const offenders = targets.flatMap((node) =>
    zeroBasedLoops(node).map((loop) => ({
      owner: node.name?.text ?? "computed range",
      code: loop.getText(ast),
    })),
  );
  assert.deepEqual(
    offenders,
    [],
    `scroll lookup still rescans row/column arrays from zero:\n${JSON.stringify(offenders, null, 2)}`,
  );
});

test("P3-XLSX-SCROLL-01 regression: scroll repaint does not rewrite Canvas dimensions", () => {
  const ast = sourceFile(XLSX_GRID);
  const resize = functionNode(ast, "resizeCanvas");
  assert.ok(resize, "resizeCanvas was not found");
  const assignments = descendants(resize, (node) => {
    if (!ts.isBinaryExpression(node) || node.operatorToken.kind !== ts.SyntaxKind.EqualsToken) {
      return false;
    }
    return (
      ts.isPropertyAccessExpression(node.left) &&
      ts.isIdentifier(node.left.expression) &&
      node.left.expression.text === "canvas" &&
      ["width", "height"].includes(node.left.name.text)
    );
  });
  const unconditional = assignments.filter((assignment) => {
    let parent = assignment.parent;
    while (parent && parent !== resize) {
      if (ts.isIfStatement(parent)) return false;
      parent = parent.parent;
    }
    return true;
  });
  const scrollDrivenPaint = descendants(ast, (node) => {
    if (!ts.isCallExpression(node) || node.expression.getText(ast) !== "watch") return false;
    const dependencyText = node.arguments[0]?.getText(ast) ?? "";
    const callbackText = node.arguments[1]?.getText(ast) ?? "";
    return (
      dependencyText.includes("scrollTop.value") &&
      dependencyText.includes("scrollLeft.value") &&
      callbackText.includes("paintGrid")
    );
  });
  assert.equal(
    scrollDrivenPaint.length > 0 && unconditional.length > 0,
    false,
    `scroll-driven paint reaches unconditional Canvas width/height writes: ${unconditional.map((node) => node.getText(ast)).join(", ")}`,
  );
});
