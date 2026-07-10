import { existsSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, extname, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const ts = requireFromDocxCore("typescript");

const entries = [
  resolve(rootDir, "packages/docx-core/src/core.ts"),
  resolve(rootDir, "packages/xlsx-core/src/core.ts"),
];

const platformIdentifiers = new Set([
  "Blob",
  "CanvasRenderingContext2D",
  "DOMParser",
  "Document",
  "Element",
  "FileReader",
  "HTMLCanvasElement",
  "HTMLElement",
  "ImageBitmap",
  "OffscreenCanvas",
  "Request",
  "Response",
  "URL",
  "Worker",
  "XMLDocument",
  "XMLSerializer",
  "console",
  "createImageBitmap",
  "document",
  "fetch",
  "globalThis",
  "navigator",
  "performance",
  "window",
]);
const mutableCollectionNames = new Set(["Map", "Set", "WeakMap", "WeakSet"]);
const visited = new Set();
const runtimeEdges = [];
const typeEdges = [];
const externalRuntimeImports = [];
const publicExports = [];
const violations = [];
const violationKeys = new Set();

function repoPath(filePath) {
  return relative(rootDir, filePath).split(sep).join("/");
}

function addViolation(code, filePath, node, message) {
  const sourceFile = node.getSourceFile();
  const position = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
  const violation = {
    code,
    column: position.character + 1,
    file: repoPath(filePath),
    line: position.line + 1,
    message,
  };
  const key = JSON.stringify(violation);
  if (violationKeys.has(key)) return;
  violationKeys.add(key);
  violations.push(violation);
}

function isTypeOnlyImport(statement) {
  const clause = statement.importClause;
  if (!clause) return false;
  if (clause.isTypeOnly) return true;
  if (clause.name) return false;
  const bindings = clause.namedBindings;
  return Boolean(
    bindings &&
    ts.isNamedImports(bindings) &&
    bindings.elements.length > 0 &&
    bindings.elements.every((element) => element.isTypeOnly),
  );
}

function isTypeOnlyExport(statement) {
  if (statement.isTypeOnly) return true;
  return Boolean(
    statement.exportClause &&
    ts.isNamedExports(statement.exportClause) &&
    statement.exportClause.elements.length > 0 &&
    statement.exportClause.elements.every((element) => element.isTypeOnly),
  );
}

function resolveLocalModule(fromFile, specifier) {
  if (!specifier.startsWith(".")) return null;
  const base = resolve(dirname(fromFile), specifier);
  const candidates = extname(base)
    ? [
        base,
        base.replace(/\.js$/u, ".ts"),
        base.replace(/\.mjs$/u, ".mts"),
      ]
    : [
        `${base}.ts`,
        `${base}.tsx`,
        `${base}.mts`,
        resolve(base, "index.ts"),
        resolve(base, "index.tsx"),
      ];
  return candidates.find((candidate) => existsSync(candidate)) ?? null;
}

function isBannedModulePath(specifier, resolvedPath) {
  const normalizedSpecifier = specifier.replaceAll("\\", "/");
  const normalizedResolved = resolvedPath?.replaceAll("\\", "/") ?? "";
  return (
    /(^|\/)vue(?:$|[-/])/u.test(normalizedSpecifier) ||
    normalizedSpecifier.includes("/demo") ||
    normalizedSpecifier.startsWith("demo") ||
    normalizedResolved.includes("/apps/") ||
    normalizedResolved.includes("/demo/") ||
    normalizedResolved.includes("/packages/vue-")
  );
}

function scanPlatformIdentifiers(filePath, sourceFile, entryOnly = false) {
  function visit(node) {
    if (ts.isIdentifier(node) && platformIdentifiers.has(node.text)) {
      addViolation(
        entryOnly ? "CORE_ENTRY_PLATFORM_TYPE" : "CORE_PLATFORM_DEPENDENCY",
        filePath,
        node,
        entryOnly
          ? `纯入口声明包含平台名称 ${node.text}。`
          : `纯运行图包含平台能力 ${node.text}。`,
      );
    }
    ts.forEachChild(node, visit);
  }
  visit(sourceFile);
}

function scanTopLevelState(filePath, sourceFile) {
  for (const statement of sourceFile.statements) {
    if (!ts.isVariableStatement(statement)) continue;
    const isConst = Boolean(statement.declarationList.flags & ts.NodeFlags.Const);
    for (const declaration of statement.declarationList.declarations) {
      const name = declaration.name.getText(sourceFile);
      if (!isConst) {
        addViolation(
          "CORE_TOP_LEVEL_MUTABLE_STATE",
          filePath,
          declaration,
          `纯运行图不允许顶层可变变量 ${name}。`,
        );
        continue;
      }

      const initializer = declaration.initializer;
      const isMutableCollection = Boolean(
        initializer &&
        ts.isNewExpression(initializer) &&
        ts.isIdentifier(initializer.expression) &&
        mutableCollectionNames.has(initializer.expression.text),
      );
      if (isMutableCollection || /cache|memo|registry|singleton/iu.test(name)) {
        addViolation(
          "CORE_TOP_LEVEL_MUTABLE_CACHE",
          filePath,
          declaration,
          `纯运行图不允许顶层可变缓存 ${name}。`,
        );
      }
    }
  }
}

function collectModuleReferences(filePath, sourceFile) {
  const references = [];
  for (const statement of sourceFile.statements) {
    if (ts.isImportDeclaration(statement) && ts.isStringLiteral(statement.moduleSpecifier)) {
      references.push({
        node: statement.moduleSpecifier,
        specifier: statement.moduleSpecifier.text,
        typeOnly: isTypeOnlyImport(statement),
      });
      continue;
    }
    if (ts.isExportDeclaration(statement) && statement.moduleSpecifier && ts.isStringLiteral(statement.moduleSpecifier)) {
      if (!statement.exportClause) {
        addViolation(
          "CORE_WILDCARD_EXPORT",
          filePath,
          statement,
          "纯入口及其运行图必须使用显式导出，不能使用 wildcard。",
        );
      }
      references.push({
        node: statement.moduleSpecifier,
        specifier: statement.moduleSpecifier.text,
        typeOnly: isTypeOnlyExport(statement),
      });
    }
  }
  return references;
}

function visitRuntimeFile(filePath, isEntry = false) {
  const absolutePath = resolve(filePath);
  if (visited.has(absolutePath)) return;
  visited.add(absolutePath);

  const sourceText = readFileSync(absolutePath, "utf8");
  const sourceFile = ts.createSourceFile(
    absolutePath,
    sourceText,
    ts.ScriptTarget.Latest,
    true,
    absolutePath.endsWith(".tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );
  scanPlatformIdentifiers(absolutePath, sourceFile, false);
  if (isEntry) scanPlatformIdentifiers(absolutePath, sourceFile, true);
  scanTopLevelState(absolutePath, sourceFile);

  if (isEntry) {
    for (const statement of sourceFile.statements) {
      if (ts.isExportDeclaration(statement) && statement.exportClause && ts.isNamedExports(statement.exportClause)) {
        for (const element of statement.exportClause.elements) {
          publicExports.push({
            entry: repoPath(absolutePath),
            name: element.name.text,
          });
        }
        continue;
      }
      const isExported = statement.modifiers?.some(
        (modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword,
      );
      if (!isExported) continue;
      if (
        (ts.isFunctionDeclaration(statement) ||
          ts.isClassDeclaration(statement) ||
          ts.isInterfaceDeclaration(statement) ||
          ts.isTypeAliasDeclaration(statement) ||
          ts.isEnumDeclaration(statement)) &&
        statement.name
      ) {
        publicExports.push({
          entry: repoPath(absolutePath),
          name: statement.name.text,
        });
      } else if (ts.isVariableStatement(statement)) {
        for (const declaration of statement.declarationList.declarations) {
          if (ts.isIdentifier(declaration.name)) {
            publicExports.push({
              entry: repoPath(absolutePath),
              name: declaration.name.text,
            });
          }
        }
      }
    }
  }

  for (const reference of collectModuleReferences(absolutePath, sourceFile)) {
    const resolvedPath = resolveLocalModule(absolutePath, reference.specifier);
    const edge = {
      from: repoPath(absolutePath),
      specifier: reference.specifier,
      to: resolvedPath ? repoPath(resolvedPath) : null,
    };
    if (reference.typeOnly) {
      typeEdges.push(edge);
      continue;
    }

    runtimeEdges.push(edge);
    if (isBannedModulePath(reference.specifier, resolvedPath)) {
      addViolation(
        "CORE_BANNED_MODULE",
        absolutePath,
        reference.node,
        `纯运行图不能依赖 ${reference.specifier}。`,
      );
    }
    if (!resolvedPath) {
      externalRuntimeImports.push({
        from: repoPath(absolutePath),
        specifier: reference.specifier,
      });
      continue;
    }
    visitRuntimeFile(resolvedPath, false);
  }
}

try {
  for (const entry of entries) {
    if (!existsSync(entry)) {
      violations.push({
        code: "CORE_ENTRY_MISSING",
        column: 1,
        file: repoPath(entry),
        line: 1,
        message: "纯入口文件不存在。",
      });
      continue;
    }
    visitRuntimeFile(entry, true);
  }
  const evidence = {
    entries: entries.map(repoPath),
    externalRuntimeImports: externalRuntimeImports.sort((left, right) =>
      `${left.from}:${left.specifier}`.localeCompare(`${right.from}:${right.specifier}`)
    ),
    runtimeEdges: runtimeEdges.sort((left, right) =>
      `${left.from}:${left.specifier}`.localeCompare(`${right.from}:${right.specifier}`)
    ),
    runtimeFiles: [...visited].map(repoPath).sort(),
    publicExports: publicExports.sort((left, right) =>
      `${left.entry}:${left.name}`.localeCompare(`${right.entry}:${right.name}`)
    ),
    status: violations.length === 0 ? "PASS" : "FAIL",
    typeEdges: typeEdges.sort((left, right) =>
      `${left.from}:${left.specifier}`.localeCompare(`${right.from}:${right.specifier}`)
    ),
    violations: violations.sort((left, right) =>
      `${left.file}:${left.line}:${left.column}:${left.code}`.localeCompare(
        `${right.file}:${right.line}:${right.column}:${right.code}`,
      )
    ),
  };
  process.stdout.write(`${JSON.stringify(evidence, null, 2)}\n`);
  if (violations.length > 0) process.exitCode = 1;
} catch (error) {
  process.stdout.write(`${JSON.stringify({
    error: error instanceof Error ? error.message : String(error),
    status: "ERROR",
  }, null, 2)}\n`);
  process.exitCode = 1;
}
