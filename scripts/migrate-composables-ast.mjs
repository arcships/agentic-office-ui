// AST-based migration of composables.ts React hooks -> Vue reactivity.
// Uses TypeScript compiler API for correct parsing/transformation.
import ts from "/Users/eric8810/Code/agentic-office-ui/node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/typescript.js";
import { readFileSync, writeFileSync } from "node:fs";

const FILE = "packages/vue-xlsx/src/composables.ts";
const sourceText = readFileSync(FILE, "utf-8");
const sourceFile = ts.createSourceFile("composables.ts", sourceText, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

// Collect transformations as [start, end, replacement] then apply from end.
const edits = [];
const setterMap = {}; // setX -> x (ref name)

function visit(node) {
  // const NAME = React.useCallback( (ARGS) => BODY, [DEPS] )
  if (
    ts.isVariableStatement(node) &&
    node.declarationList.declarations.length === 1
  ) {
    const decl = node.declarationList.declarations[0];
    const name = decl.name.getText(sourceFile);
    const init = decl.initializer;
    // const [x, setX] = React.useState(INIT)
    if (init && ts.isCallExpression(init) && init.expression.getText(sourceFile) === "React.useState") {
      if (ts.isArrayBindingPattern(decl.name)) {
        const elems = decl.name.elements;
        if (elems.length >= 1 && ts.isBindingElement(elems[0])) {
          const stateName = elems[0].name.getText(sourceFile);
          const initArg = init.arguments[0];
          const initText = initArg ? initArg.getText(sourceFile) : "";
          const replacement = `const ${stateName} = ref(${initText})`;
          edits.push([node.getStart(sourceFile), node.getEnd(), replacement]);
          // Record setter mapping if present: setX -> x
          if (elems.length >= 2 && ts.isBindingElement(elems[1])) {
            const setterName = elems[1].name.getText(sourceFile);
            setterMap[setterName] = stateName;
          }
          return;
        }
      }
    }
    // const NAME = React.useRef(INIT)
    if (init && ts.isCallExpression(init) && init.expression.getText(sourceFile) === "React.useRef") {
      const initArg = init.arguments[0];
      const initText = initArg ? initArg.getText(sourceFile) : "";
      const replacement = `const ${name} = shallowRef(${initText})`;
      edits.push([node.getStart(sourceFile), node.getEnd(), replacement]);
      return;
    }
    if (init && ts.isCallExpression(init) && init.expression.getText(sourceFile) === "React.useCallback") {
      const args = init.arguments;
      if (args.length >= 1 && ts.isArrowFunction(args[0])) {
        const arrow = args[0];
        const isAsync = !!arrow.modifiers?.some(m => m.kind === ts.SyntaxKind.AsyncKeyword) || !!(arrow.flags & 16777216); // AsyncFlag
        // Build params text
        const paramsText = arrow.parameters.map(p => p.getText(sourceFile)).join(", ");
        const body = arrow.body;
        let bodyText;
        if (ts.isBlock(body)) {
          bodyText = body.getText(sourceFile); // includes { }
          // Strip outer braces? We'll produce function NAME(params) { ... }
          const inner = bodyText.slice(1, -1);
          const replacement = `${isAsync ? "async " : ""}function ${name}(${paramsText}) {${inner}}`;
          edits.push([node.getStart(sourceFile), node.getEnd(), replacement]);
        } else {
          // expression body
          const exprText = body.getText(sourceFile);
          const replacement = `${isAsync ? "async " : ""}function ${name}(${paramsText}) {\n  return ${exprText};\n}`;
          edits.push([node.getStart(sourceFile), node.getEnd(), replacement]);
        }
        return;
      }
    }

    // const NAME = React.useMemo( () => EXPR, [DEPS] )
    if (init && ts.isCallExpression(init) && init.expression.getText(sourceFile) === "React.useMemo") {
      const args = init.arguments;
      if (args.length >= 1 && ts.isArrowFunction(args[0])) {
        const arrow = args[0];
        const body = arrow.body;
        let bodyText;
        if (ts.isBlock(body)) {
          bodyText = body.getText(sourceFile);
        } else {
          bodyText = `{ return ${body.getText(sourceFile)}; }`;
        }
        const replacement = `const ${name} = computed(() => ${bodyText})`;
        edits.push([node.getStart(sourceFile), node.getEnd(), replacement]);
        return;
      }
    }

    // const X = React.useDeferredValue(Y)
    if (init && ts.isCallExpression(init) && init.expression.getText(sourceFile) === "React.useDeferredValue") {
      const argText = init.arguments[0]?.getText(sourceFile) ?? "";
      const replacement = `const ${name} = ${argText}`;
      edits.push([node.getStart(sourceFile), node.getEnd(), replacement]);
      return;
    }
  }

  // React.useEffect(...) as an ExpressionStatement (top-level statement in function body)
  if (
    ts.isExpressionStatement(node) &&
    ts.isCallExpression(node.expression) &&
    node.expression.expression.getText(sourceFile) === "React.useEffect"
  ) {
    // Leave a marker for manual conversion; we'll handle these manually after.
    const text = node.getText(sourceFile);
    const replacement = `/* @vue-useEffect-manual */ ${text}`;
    edits.push([node.getStart(sourceFile), node.getEnd(), replacement]);
    return;
  }

  ts.forEachChild(node, visit);
}

visit(sourceFile);

// Apply edits from end to start.
edits.sort((a, b) => b[0] - a[0]);
let out = sourceText;
for (const [start, end, replacement] of edits) {
  out = out.slice(0, start) + replacement + out.slice(end);
}

writeFileSync(FILE, out);
console.log("AST edits applied:", edits.length);
console.log("Setter map entries:", Object.keys(setterMap).length);
console.log("Remaining React. count:", (out.match(/React\./g) || []).length);

// Save setter map for the next-phase script.
writeFileSync("/tmp/setterMap.json", JSON.stringify(setterMap, null, 2));
