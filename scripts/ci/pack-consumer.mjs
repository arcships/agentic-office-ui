#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  cpSync,
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const output = path.resolve(
  process.env.PACK_CONSUMER_EVIDENCE_DIR ||
    path.join(root, "output", "acceptance", "pack-consumer"),
);
const manifestPath = path.resolve(
  process.env.PACK_MANIFEST_PATH || path.join(output, "..", "real-tgz-manifests", "summary.json"),
);
const template = path.join(root, "tests", "consumer", "template");
const publicApiContract = JSON.parse(
  readFileSync(path.join(root, "scripts", "ci", "public-api-contract.json"), "utf8"),
);
const consumerDir = mkdtempSync(path.join(os.tmpdir(), "agentic-office-tgz-consumer-"));
const probe = process.env.PACK_CONSUMER_PROBE || "";
const logs = [];
let preview;
let result = "FAIL";
let failure;

rmSync(output, { recursive: true, force: true });
mkdirSync(output, { recursive: true });

function run(id, executable, args) {
  const commandResult = spawnSync(executable, args, {
    cwd: consumerDir,
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  });
  const log = `${commandResult.stdout || ""}${commandResult.stderr || ""}`;
  writeFileSync(path.join(output, `${id}.log`), log);
  logs.push({ id, command: [executable, ...args], exitCode: commandResult.status ?? 1 });
  if (commandResult.status !== 0) {
    throw new Error(`${id} exited ${commandResult.status}`);
  }
}

function walk(directory) {
  const result = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) result.push(...walk(absolute));
    else result.push(absolute);
  }
  return result;
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

function waitForServer(url, child) {
  return new Promise((resolve, reject) => {
    const deadline = Date.now() + 30_000;
    const attempt = () => {
      if (child.exitCode !== null) return reject(new Error(`consumer preview exited ${child.exitCode}`));
      const request = http.get(url, (response) => {
        response.resume();
        if (response.statusCode === 200) resolve();
        else if (Date.now() >= deadline) reject(new Error(`consumer preview returned ${response.statusCode}`));
        else setTimeout(attempt, 100);
      });
      request.on("error", () => {
        if (Date.now() >= deadline) reject(new Error("consumer preview did not become ready"));
        else setTimeout(attempt, 100);
      });
    };
    attempt();
  });
}

try {
  const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
  if (manifest.result !== "PASS" || manifest.entries.length !== 5) {
    throw new Error("real tgz manifest preflight did not pass for five packages");
  }
  const tgzDir = path.dirname(manifestPath);
  const packageFiles = Object.fromEntries(
    manifest.entries.map((entry) => [entry.package, path.join(tgzDir, "tgz", entry.filename)]),
  );
  cpSync(template, consumerDir, { recursive: true });
  mkdirSync(path.join(consumerDir, "public", "fixtures"), { recursive: true });
  for (const file of ["invoice-table.docx", "sales-table.xlsx", "sample.pdf"]) {
    cpSync(
      path.join(root, "apps", "demo", "public", "samples", file),
      path.join(consumerDir, "public", "fixtures", file),
    );
  }

  const packageJson = {
    name: "agentic-office-tgz-consumer",
    private: true,
    version: "0.0.0",
    type: "module",
    scripts: {
      typecheck: "tsc --noEmit",
      "vue-typecheck": "vue-tsc --noEmit",
      build: "vite build",
      preview: "vite preview",
    },
    dependencies: {
      ...Object.fromEntries(
        Object.entries(packageFiles).map(([name, file]) => [name, `file:${file}`]),
      ),
      vue: "3.5.39",
    },
    devDependencies: {
      "@vitejs/plugin-vue": "5.2.4",
      typescript: "5.9.3",
      vite: "6.4.3",
      "vue-tsc": "3.3.6",
    },
  };
  writeFileSync(path.join(consumerDir, "package.json"), `${JSON.stringify(packageJson, null, 2)}\n`);
  run("npm-install", "npm", ["install", "--ignore-scripts", "--no-audit", "--no-fund"]);

  const installChecks = [];
  const exportContractChecks = [];
  const realConsumerDir = realpathSync(consumerDir);
  for (const name of Object.keys(packageFiles)) {
    const installed = path.join(consumerDir, "node_modules", ...name.split("/"));
    installChecks.push({
      package: name,
      exists: existsSync(installed),
      symlink: lstatSync(installed).isSymbolicLink(),
      realPath: realpathSync(installed),
    });
    const expected = publicApiContract.packages.find((entry) => entry.name === name);
    if (!expected) throw new Error(`public API contract is missing ${name}`);
    const installedManifest = JSON.parse(readFileSync(path.join(installed, "package.json"), "utf8"));
    const exportMatches =
      JSON.stringify(stable(installedManifest.exports)) ===
      JSON.stringify(stable(expected.exports));
    const targets = [...new Set(collectExportTargets(expected.exports))];
    const missingTargets = targets.filter(
      (target) => !existsSync(path.join(installed, target.replace(/^\.\//, ""))),
    );
    exportContractChecks.push({ package: name, exportMatches, targets, missingTargets });
  }
  if (installChecks.some((item) => !item.exists || item.symlink || !item.realPath.startsWith(realConsumerDir))) {
    throw new Error("external consumer contains a missing or linked workspace package");
  }
  if (exportContractChecks.some((item) => !item.exportMatches || item.missingTargets.length)) {
    throw new Error(`installed tgz exports do not match the public contract: ${JSON.stringify(exportContractChecks)}`);
  }
  const forbiddenPublicResources = walk(path.join(consumerDir, "public"))
    .map((file) => path.relative(path.join(consumerDir, "public"), file))
    .filter((file) => file.endsWith(".wasm") || /worker/i.test(file));
  if (forbiddenPublicResources.length) {
    throw new Error(`consumer public contains copied runtime resources: ${forbiddenPublicResources.join(", ")}`);
  }
  writeFileSync(
    path.join(output, "consumer-metadata.json"),
    `${JSON.stringify({
      consumerDirectory: consumerDir,
      outsideWorkspace: !consumerDir.startsWith(root),
      installChecks,
      exportContractChecks,
      publicFiles: walk(path.join(consumerDir, "public")).map((file) => path.relative(consumerDir, file)),
      forbiddenPublicResources,
    }, null, 2)}\n`,
  );

  run("positive-public-entry-resolution", process.execPath, [
    "--input-type=module",
    "--eval",
    `import { existsSync, statSync } from "node:fs";
import { fileURLToPath } from "node:url";
const declared = ${JSON.stringify(publicApiContract.positiveSpecifiers)};
const executable = new Set(${JSON.stringify(publicApiContract.executableSpecifiers)});
const resolved = [];
for (const specifier of declared) {
  const url = import.meta.resolve(specifier);
  if (!url.startsWith("file:")) throw new Error("public entry did not resolve to an installed file: " + specifier + " -> " + url);
  const file = fileURLToPath(url);
  if (!existsSync(file) || !statSync(file).isFile()) throw new Error("public entry file is missing: " + specifier + " -> " + file);
  if (executable.has(specifier)) await import(specifier);
  resolved.push({ specifier, file, executed: executable.has(specifier) });
}
process.stdout.write(JSON.stringify({ status: "PASS", resolved }, null, 2));`,
  ]);

  run("deep-import-rejection", process.execPath, [
    "--input-type=module",
    "--eval",
    `const matrix = ${JSON.stringify(publicApiContract.deepImportRejections)};
for (const [packageName, forbidden] of Object.entries(matrix)) {
  for (const specifier of forbidden) {
    let wasRejected = false;
    try {
      await import(specifier);
    } catch (error) {
      if (error?.code !== "ERR_PACKAGE_PATH_NOT_EXPORTED") {
        throw new Error(packageName + " rejected " + specifier + " with " + error?.code, { cause: error });
      }
      wasRejected = true;
    }
    if (!wasRejected) throw new Error("deep import unexpectedly resolved: " + specifier);
  }
}
process.stdout.write(JSON.stringify({ status: "PASS", matrix }, null, 2));`,
  ]);

  if (probe === "missing-entry") {
    rmSync(path.join(consumerDir, "node_modules", "@extend-ai", "vue-docx", "dist", "index.js"));
  }

  run("typescript", "npm", ["run", "typecheck"]);
  run("vue-typescript", "npm", ["run", "vue-typecheck"]);
  run("vite-build", "npm", ["run", "build"]);

  const expectedResources = manifest.entries.flatMap((entry) =>
    entry.resources.map((resource) => ({ package: entry.package, ...resource })),
  );
  const expectedWasmResources = expectedResources.filter((item) => item.path.endsWith(".wasm"));
  const expectedWorkerResources = expectedResources.filter((item) => item.path.endsWith("worker.js"));
  const expectedPdfiumResource = expectedWasmResources.find(
    (item) => item.package === "@extend-ai/vue-extend" && item.path === "dist/pdfium.wasm",
  );
  if (
    expectedResources.length !== 5 ||
    expectedWasmResources.length !== 3 ||
    expectedWorkerResources.length !== 2 ||
    !expectedPdfiumResource ||
    expectedResources.some((item) => !item.sha256)
  ) {
    throw new Error(`expected two Workers and three WASM files, including the PDFium tgz asset: ${JSON.stringify(expectedResources)}`);
  }
  const expectedPath = path.join(output, "expected-tgz-resources.json");
  writeFileSync(expectedPath, `${JSON.stringify(expectedResources, null, 2)}\n`);

  const builtRuntimeFiles = walk(path.join(consumerDir, "dist"))
    .filter((file) => file.endsWith(".wasm") || /worker.*\.js$/.test(path.basename(file)))
    .map((file) => ({
      path: path.relative(consumerDir, file),
      sha256: createHash("sha256").update(readFileSync(file)).digest("hex"),
    }));
  const expectedWasmHashes = new Set(
    expectedResources.filter((item) => item.path.endsWith(".wasm")).map((item) => item.sha256),
  );
  const builtWasm = builtRuntimeFiles.filter((item) => item.path.endsWith(".wasm"));
  const builtWorkers = builtRuntimeFiles.filter((item) => /worker.*\.js$/.test(path.basename(item.path)));
  const builtWasmHashes = new Set(builtWasm.map((item) => item.sha256));
  if (
    builtRuntimeFiles.length !== 5 ||
    builtWasm.length !== 3 ||
    builtWorkers.length !== 2 ||
    builtWasmHashes.size !== expectedWasmHashes.size ||
    builtWasm.some((item) => !expectedWasmHashes.has(item.sha256))
  ) {
    throw new Error(`formal consumer resources do not match tgz hashes: ${JSON.stringify(builtRuntimeFiles)}`);
  }
  writeFileSync(
    path.join(output, "built-runtime-resources.json"),
    `${JSON.stringify(builtRuntimeFiles, null, 2)}\n`,
  );

  const port = 4191;
  const previewLog = path.join(output, "preview.log");
  const previewHandle = await import("node:fs").then(({ openSync }) => openSync(previewLog, "w"));
  preview = spawn("npm", ["run", "preview", "--", "--host", "127.0.0.1", "--port", String(port)], {
    cwd: consumerDir,
    detached: true,
    stdio: ["ignore", previewHandle, previewHandle],
  });
  const baseUrl = `http://127.0.0.1:${port}`;
  await waitForServer(baseUrl, preview);
  run("browser", process.execPath, [
    path.join(root, "scripts", "ci", "run-python.mjs"),
    path.join(root, "tests", "blackbox", "pack_consumer.py"),
    baseUrl,
    path.join(output, "browser"),
    consumerDir,
    expectedPath,
  ]);
  result = "PASS";
} catch (error) {
  failure = error instanceof Error ? error.stack || error.message : String(error);
} finally {
  if (preview && preview.exitCode === null) {
    try { process.kill(-preview.pid, "SIGTERM"); } catch { preview.kill("SIGTERM"); }
  }
  const cleanedPath = consumerDir;
  if (process.env.KEEP_PACK_CONSUMER !== "1") rmSync(consumerDir, { recursive: true, force: true });
  writeFileSync(
    path.join(output, "summary.json"),
    `${JSON.stringify({
      suite: "P1-PACK-EXTERNAL-CONSUMER",
      result,
      probe: probe || null,
      manifestPath,
      temporaryConsumer: cleanedPath,
      temporaryConsumerCleaned: !existsSync(cleanedPath),
      commands: logs,
      failure: failure || null,
    }, null, 2)}\n`,
  );
}

console.log(`${result}: external tgz consumer at ${output}`);
if (failure) console.error(failure);
process.exit(result === "PASS" ? 0 : 1);
