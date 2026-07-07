// Phase 2b (v2): Convert setter calls -> ref.value = ... with recursive argument transformation
// to handle nested setter calls inside updater bodies. Single pass, no overlapping edits.
import ts from "/Users/eric8810/Code/agentic-office-ui/node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/typescript.js";
import { readFileSync, writeFileSync } from "node:fs";

const FILE = "packages/vue-xlsx/src/composables.ts";
const setterMap = JSON.parse(readFileSync("/tmp/setterMap.json", "utf-8"));
const sourceText = readFileSync(FILE, "utf-8");
const sourceFile = ts.createSourceFile("composables.ts", sourceText, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

// Recursively transform a node into source text, converting any nested setter calls.
// Returns the text representation. We only need to special-case CallExpressions that are setters;
// everything else uses getText (which already reflects .current->.value from phase2a).
function transformNode(node) {
  // Setter call?
  if (ts.isCallExpression(node) && ts.isIdentifier(node.expression) && setterMap[node.expression.getText(sourceFile)]) {
    return transformSetterCall(node);
  }
  // For other nodes, we could recurse, but to keep it simple we only recurse into nodes that
  // might contain a setter call. Use a visit that finds setter calls within and replaces them.
  // Simplest correct approach: walk children; if a child is a setter call, splice its text.
  return spliceChildSetters(node);
}

// Produce text for `node`, replacing any descendant setter-call nodes with their transformed text.
function spliceChildSetters(node) {
  // Collect setter-call descendants (non-overlapping, outermost-first among siblings).
  const found = [];
  function find(n) {
    if (ts.isCallExpression(n) && ts.isIdentifier(n.expression) && setterMap[n.expression.getText(sourceFile)]) {
      found.push(n);
      return; // don't descend into a setter call; transformSetterCall handles its interior
    }
    ts.forEachChild(n, find);
  }
  ts.forEachChild(node, find);
  if (found.length === 0) return node.getText(sourceFile);
  // Sort by start desc.
  found.sort((a, b) => b.getStart(sourceFile) - a.getStart(sourceFile));
  let text = node.getText(sourceFile);
  const nodeStart = node.getStart(sourceFile);
  for (const fc of found) {
    const relStart = fc.getStart(sourceFile) - nodeStart;
    const relEnd = fc.getEnd() - nodeStart;
    const repl = transformSetterCall(fc);
    text = text.slice(0, relStart) + repl + text.slice(relEnd);
  }
  return text;
}

function transformSetterCall(node) {
  const callee = node.expression.getText(sourceFile);
  const refName = setterMap[callee];
  const args = node.arguments;
  if (args.length !== 1) return node.getText(sourceFile);
  const arg = args[0];
  if (ts.isArrowFunction(arg)) {
    const params = arg.parameters.map(p => p.getText(sourceFile)).join(", ");
    const body = arg.body;
    const isAsync = !!arg.modifiers?.some(m => m.kind === ts.SyntaxKind.AsyncKeyword);
    if (params.trim()) {
      const paramIdent = arg.parameters[0].name.getText(sourceFile);
      // Transform body text, converting nested setters first.
      let bodyText = spliceChildSetters(body);
      const re = new RegExp(`(?<![\\w$.])${paramIdent}(?![\\w$])`, "g");
      let newBody;
      if (ts.isBlock(body)) {
        const inner = bodyText.slice(1, -1);
        newBody = `{${inner.replace(re, `${refName}.value`)}}`;
      } else {
        newBody = bodyText.replace(re, `${refName}.value`);
      }
      let rhs;
      if (ts.isBlock(body)) {
        rhs = `(() => ${newBody})()`;
      } else {
        rhs = newBody;
      }
      if (isAsync) rhs = `(await (async () => ${newBody})())`;
      return `${refName}.value = ${rhs}`;
    }
    // no param arrow
    const bodyText = spliceChildSetters(body);
    return `${refName}.value = ${bodyText}`;
  }
  // plain value argument: transform nested setters within the argument.
  const argText = spliceChildSetters(arg);
  return `${refName}.value = ${argText}`;
}

// Top-level: find all top-level setter calls (not nested inside another setter call) and replace.
const edits = [];
function visitTop(node) {
  if (ts.isCallExpression(node) && ts.isIdentifier(node.expression) && setterMap[node.expression.getText(sourceFile)]) {
    edits.push([node.getStart(sourceFile), node.getEnd(), transformSetterCall(node)]);
    return; // don't descend; transformSetterCall handled interior
  }
  ts.forEachChild(node, visitTop);
}
visitTop(sourceFile);

edits.sort((a, b) => b[0] - a[0]);
let out = sourceText;
for (const [start, end, replacement] of edits) {
  out = out.slice(0, start) + replacement + out.slice(end);
}
writeFileSync(FILE, out);
console.log("Phase 2b v2 top-level setter edits:", edits.length);
