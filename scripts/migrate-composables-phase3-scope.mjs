// Phase 3 (scope-aware): Add .value to reads of ref/shallowRef/computed, but ONLY when the
// identifier resolves to that ref declaration (not shadowed by a parameter/local of the same name).
import ts from "/Users/eric8810/Code/agentic-office-ui/node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/typescript.js";
import { readFileSync, writeFileSync } from "node:fs";

const FILE = "packages/vue-xlsx/src/composables.ts";
const sourceText = readFileSync(FILE, "utf-8");
const sourceFile = ts.createSourceFile("composables.ts", sourceText, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

// A "value declaration" is a VariableDeclaration whose initializer is ref(...)/shallowRef(...)/computed(...).
// We record, for each such declaration, the name and the source range [start,end) of the VariableStatement
// so we can test whether a given identifier position is in a scope where this declaration is visible AND
// not shadowed.

// Collect all ref/shallowRef/computed declarations globally (top-level + nested).
const decls = []; // { name, declStart, declEnd (of the VariableStatement), scopeNode }
function collectDecls(node) {
  if (ts.isVariableStatement(node)) {
    for (const d of node.declarationList.declarations) {
      const init = d.initializer;
      if (init && ts.isCallExpression(init)) {
        const calleeText = init.expression.getText(sourceFile);
        if (calleeText === "ref" || calleeText === "shallowRef" || calleeText === "computed") {
          if (ts.isIdentifier(d.name)) {
            decls.push({ name: d.name.getText(sourceFile), pos: d.name.getStart(sourceFile) });
          }
        }
      }
    }
  }
  ts.forEachChild(node, collectDecls);
}
collectDecls(sourceFile);
const valueDeclNames = new Set(decls.map(d => d.name));
console.log("valueDeclNames (" + valueDeclNames.size + "):", [...valueDeclNames].join(", "));

// For scope resolution: we walk the AST and maintain a stack of scope frames. Each frame is a Set of
// names declared in that scope (parameters, var/let/const, function declarations, catch bindings).
// For a given identifier position, we find the innermost declaring scope.
// To keep it tractable, we do a single recursive walk maintaining `scopeStack` (array of Sets).
// When entering a function, push a new frame with its parameters; on exit, pop.
// When we encounter a VariableDeclaration/BindingElement/Parameter in the current frame, add the name.
// For each Identifier in a "read" position, we look up the scope stack from innermost to outermost;
// if the name is declared in a frame as a NON-ref parameter/local, it shadows -> skip.
// If the name is declared in a frame as a ref/computed (we match by position: the declaration name pos),
// add .value.

// Build a map: name -> array of { frameDepth, kind } where kind is "ref" | "other".
// Simpler: during the walk, for each identifier read, resolve the nearest declaration by scanning
// the current scopeStack frames in reverse.

const scopeStack = []; // each: Map<name, "ref" | "other">
function pushScope() { scopeStack.push(new Map()); }
function popScope() { scopeStack.pop(); }
function declareInCurrent(name, kind) {
  if (scopeStack.length === 0) return;
  const frame = scopeStack[scopeStack.length - 1];
  // Don't overwrite an existing declaration in the same frame (first declaration wins, like hoisting for params).
  if (!frame.has(name)) frame.set(name, kind);
}

// Resolve: nearest frame that has the name. Returns "ref" | "other" | undefined.
function resolveName(name) {
  for (let i = scopeStack.length - 1; i >= 0; i--) {
    const frame = scopeStack[i];
    if (frame.has(name)) return frame.get(name);
  }
  return undefined;
}

// Position-based check: is this identifier position the same as a ref decl name? (then it's a declaration, skip)
const refDeclNamePositions = new Set(decls.map(d => d.pos));

const edits = [];

function needsValueAccess(identifier, parent) {
  const name = identifier.getText(sourceFile);
  if (!valueDeclNames.has(name)) return false;
  // Skip declaration positions
  if (refDeclNamePositions.has(identifier.getStart(sourceFile))) return false;
  // Resolve scope
  const kind = resolveName(name);
  if (kind !== "ref") return false; // shadowed by param/local, or not a ref in scope

  // Now apply the same position-exclusion rules as before:
  if (ts.isVariableDeclaration(parent) && parent.name === identifier) return false;
  if (ts.isBindingElement(parent) && parent.name === identifier) return false;
  if (ts.isBindingElement(parent) && parent.propertyName === identifier) return false;
  if (ts.isParameter(parent) && parent.name === identifier) return false;
  if (ts.isPropertySignature(parent) && parent.name === identifier) return false;
  if (ts.isMethodSignature(parent) && parent.name === identifier) return false;
  if (ts.isPropertyDeclaration(parent) && parent.name === identifier) return false;
  if (ts.isMethodDeclaration(parent) && parent.name === identifier) return false;
  if (ts.isFunctionDeclaration(parent) && parent.name === identifier) return false;
  if ((ts.isFunctionExpression(parent) || ts.isArrowFunction(parent)) && parent.name === identifier) return false;
  if (ts.isPropertyAssignment(parent) && parent.name === identifier && !ts.isComputedPropertyName(parent.name)) return false;
  if (ts.isShorthandPropertyAssignment(parent) && parent.name === identifier) return "shorthand";
  if (ts.isPropertyAccessExpression(parent) && parent.expression === identifier) {
    if (parent.name.getText(sourceFile) === "value") return false;
    return true;
  }
  if (ts.isPropertyAccessExpression(parent) && parent.name === identifier) return false;
  if (ts.isQualifiedName(parent) && parent.right === identifier) return false;
  if (ts.isQualifiedName(parent) && parent.left === identifier) return true;
  if (ts.isBinaryExpression(parent) && parent.operatorToken.kind === ts.SyntaxKind.EqualsToken && parent.left === identifier) return false;
  if (ts.isComputedPropertyName(parent) && parent.expression === identifier) return true;
  return true;
}

function walk(node) {
  // When entering function-like scopes, push a frame and declare parameters.
  let pushed = false;
  if (ts.isFunctionDeclaration(node) || ts.isFunctionExpression(node) || ts.isArrowFunction(node) || ts.isMethodDeclaration(node) || ts.isConstructorDeclaration(node) || ts.isGetAccessorDeclaration(node) || ts.isSetAccessorDeclaration(node)) {
    pushScope();
    pushed = true;
    // declare function name (for FunctionDeclaration) in the OUTER scope — handled by parent traversal.
    // declare parameters
    for (const p of node.parameters) {
      declareBindingPatternNames(p.name, "other");
    }
  }
  // Module/top-level scope: ensure at least one frame exists.
  if (scopeStack.length === 0) { pushScope(); pushed = true; }

  // Pre-declare hoisted function declarations and var/let/const in the current block scope.
  // For block scopes (Block, SourceFile, CaseClause, etc.), TS treats let/const as block-scoped.
  // We handle VariableDeclaration when we visit them; but hoisting matters for forward references.
  // To keep correct enough: declare const/let at point of declaration (we visit in order), and
  // function declarations are hoisted. We'll declare function declarations now.
  if (ts.isBlock(node) || ts.isSourceFile(node) || ts.isCaseClause(node) || ts.isDefaultClause(node)) {
    // declare hoisted function declarations in this block frame (we need a frame for blocks too)
  }

  // Visit children, but handle VariableDeclarations to declare names.
  ts.forEachChild(node, (child) => {
    // Declare variable names when entering a VariableDeclaration (const/let) — but these are block-scoped,
    // so ideally we'd push a block scope. For simplicity, declare in the current function frame.
    if (ts.isVariableStatement(child)) {
      // declare each declared name as "ref" or "other" depending on initializer
      for (const d of child.declarationList.declarations) {
        const init = d.initializer;
        let kind = "other";
        if (init && ts.isCallExpression(init)) {
          const ct = init.expression.getText(sourceFile);
          if (ct === "ref" || ct === "shallowRef" || ct === "computed") kind = "ref";
        }
        declareBindingPatternNames(d.name, kind);
      }
    }
    // Function declarations are hoisted into the enclosing function/module scope.
    if (ts.isFunctionDeclaration(child) && child.name) {
      declareInCurrent(child.name.getText(sourceFile), "other");
    }
    walk(child);
    // Handle identifier reads after declaring (so a reference before declaration in same scope
    // still resolves if hoisted). For const/let TDZ we ignore — rare in this file.
  });

  // Now check if `node` itself is an identifier read (we check identifiers as we encounter them).
  // Actually we check identifiers in the forEachChild above via recursion. Let's check `node`:
  if (ts.isIdentifier(node)) {
    const parent = node.parent;
    const res = needsValueAccess(node, parent);
    if (res === "shorthand") {
      const nm = node.getText(sourceFile);
      edits.push([node.getStart(sourceFile), node.getEnd(), `${nm}: ${nm}.value`]);
    } else if (res === true) {
      const nm = node.getText(sourceFile);
      edits.push([node.getStart(sourceFile), node.getEnd(), `${nm}.value`]);
    }
  }

  if (pushed && scopeStack.length > 0) popScope();
}

// Helper: declare all binding names in a binding pattern (identifier or Object/ArrayBindingPattern).
function declareBindingPatternNames(name, kind) {
  if (ts.isIdentifier(name)) {
    declareInCurrent(name.getText(sourceFile), kind);
  } else if (ts.isObjectBindingPattern(name)) {
    for (const e of name.elements) {
      if (ts.isBindingElement(e)) declareBindingPatternNames(e.name, kind);
    }
  } else if (ts.isArrayBindingPattern(name)) {
    for (const e of name.elements) {
      if (ts.isBindingElement(e)) declareBindingPatternNames(e.name, kind);
    }
  }
}

walk(sourceFile);

edits.sort((a, b) => b[0] - a[0]);
let out = sourceText;
let overlaps = 0; let lastStart = Infinity;
for (const [start, end, replacement] of edits) {
  if (end > lastStart) overlaps++;
  lastStart = start;
  out = out.slice(0, start) + replacement + out.slice(end);
}
writeFileSync(FILE, out);
console.log("Phase 3 scope-aware edits:", edits.length, "overlaps:", overlaps);
