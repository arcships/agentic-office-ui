// Phase 2a: Convert NAME.current -> NAME.value for useRef-derived shallowRef names ONLY.
import ts from "/Users/eric8810/Code/agentic-office-ui/node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/typescript.js";
import { readFileSync, writeFileSync } from "node:fs";

const FILE = "packages/vue-xlsx/src/composables.ts";
const sourceText = readFileSync(FILE, "utf-8");
const sourceFile = ts.createSourceFile("composables.ts", sourceText, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

const refNames = new Set();
function collectRefs(node) {
  if (ts.isVariableStatement(node)) {
    for (const decl of node.declarationList.declarations) {
      const init = decl.initializer;
      if (init && ts.isCallExpression(init) && init.expression.getText(sourceFile) === "shallowRef") {
        if (ts.isIdentifier(decl.name)) refNames.add(decl.name.getText(sourceFile));
      }
    }
  }
  ts.forEachChild(node, collectRefs);
}
collectRefs(sourceFile);

const edits = [];
function visit(node) {
  if (
    ts.isPropertyAccessExpression(node) &&
    node.name.getText(sourceFile) === "current" &&
    ts.isIdentifier(node.expression) &&
    refNames.has(node.expression.getText(sourceFile))
  ) {
    const name = node.expression.getText(sourceFile);
    edits.push([node.getStart(sourceFile), node.getEnd(), `${name}.value`]);
  }
  ts.forEachChild(node, visit);
}
visit(sourceFile);

edits.sort((a, b) => b[0] - a[0]);
let out = sourceText;
for (const [start, end, replacement] of edits) {
  out = out.slice(0, start) + replacement + out.slice(end);
}
writeFileSync(FILE, out);
console.log("Phase 2a (.current) edits:", edits.length, "refNames:", [...refNames].join(", "));
