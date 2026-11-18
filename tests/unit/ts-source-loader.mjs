import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Loads a small TypeScript source graph as data-URL modules.
 *
 * Unit tests that need an internal module (not exported by the package) would
 * otherwise depend on Node's experimental type stripping or a resolve hook.
 * Transpiling every relative import keeps the tests runnable on the repository's
 * supported Node versions without enabling process-wide flags.
 */
export function createTsModuleLoader(sourceRoot, typescript) {
  const rootPath = fileURLToPath(sourceRoot);
  const cache = new Map();

  const resolveRelative = (fromRelative, specifier) => {
    const base = path.posix.normalize(path.posix.join(path.posix.dirname(fromRelative), specifier));
    return path.posix.extname(base) ? base : `${base}.ts`;
  };

  const load = (relative) => {
    const cached = cache.get(relative);
    if (cached) return cached;
    let source = readFileSync(path.resolve(rootPath, relative), "utf8");
    source = source.replace(/(\bfrom\s+|\bimport\s+)["']([^"']+)["']/gu, (match, prefix, specifier) => {
      if (!specifier.startsWith(".")) return match;
      return `${prefix}"${load(resolveRelative(relative, specifier))}"`;
    });
    const output = typescript.transpileModule(source, {
      compilerOptions: {
        module: typescript.ModuleKind.ES2022,
        target: typescript.ScriptTarget.ES2022,
      },
    }).outputText;
    const url = `data:text/javascript;base64,${Buffer.from(output).toString("base64")}`;
    cache.set(relative, url);
    return url;
  };

  return load;
}
