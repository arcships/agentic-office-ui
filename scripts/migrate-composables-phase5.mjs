// Phase 5: Convert the final `return React.useMemo(() => ({...}), [...deps])` into a plain
// returned object literal whose reactive-state properties become getters (to preserve reactivity).
// Functions and static values remain as-is.
import ts from "/Users/eric8810/Code/agentic-office-ui/node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/typescript.js";
import { readFileSync, writeFileSync } from "node:fs";

const FILE = "packages/vue-xlsx/src/composables.ts";
const sourceText = readFileSync(FILE, "utf-8");
const sourceFile = ts.createSourceFile("composables.ts", sourceText, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

let returnStmt = null;
function findReturn(node) {
  if (ts.isReturnStatement(node) && node.expression && ts.isCallExpression(node.expression) && node.expression.expression.getText(sourceFile) === "React.useMemo") {
    returnStmt = node;
    return;
  }
  ts.forEachChild(node, findReturn);
}
findReturn(sourceFile);

if (!returnStmt) { console.error("No return React.useMemo found"); process.exit(1); }

const call = returnStmt.expression;
const arrow = call.arguments[0];
let objLiteral = arrow.body; // could be ParenthesizedExpression wrapping ObjectLiteral
if (ts.isParenthesizedExpression(objLiteral)) objLiteral = objLiteral.expression;
const props = objLiteral.properties;

// Build new property list text.
const newProps = [];
for (const prop of props) {
  if (ts.isShorthandPropertyAssignment(prop)) {
    // function or static reference: keep as shorthand (e.g. addSheet, file, src)
    newProps.push(prop.getText(sourceFile));
  } else if (ts.isPropertyAssignment(prop)) {
    const key = prop.name.getText(sourceFile);
    const initText = prop.initializer.getText(sourceFile);
    // If initializer reads a .value (reactive) or contains a ref/computed access, make it a getter.
    if (/\.value\b/.test(initText) || /\bBoolean\(/.test(initText)) {
      newProps.push(`get ${key}() { return ${initText}; }`);
    } else {
      newProps.push(`${key}: ${initText}`);
    }
  } else {
    newProps.push(prop.getText(sourceFile));
  }
}

const replacement = `return {\n      ${newProps.join(",\n      ")}\n    };`;
const start = returnStmt.getStart(sourceFile);
const end = returnStmt.getEnd();
let out = sourceText.slice(0, start) + replacement + sourceText.slice(end);
writeFileSync(FILE, out);
console.log("Phase 5 (return object -> getters) done. props:", newProps.length);
