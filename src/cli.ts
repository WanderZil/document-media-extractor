#!/usr/bin/env node
import { realpathSync } from "node:fs";
import { access, mkdir, readdir, readFile, writeFile } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { extractDocumentMedia, type ExtractionPolicy } from "./index.js";

type CliOptions = {
  mode: "single" | "batch";
  inputPath: string;
  outputPath: string;
  policyPath?: string;
};

function parseCliOptions(args: string[]): CliOptions {
  const mode = args[0] === "--batch" ? "batch" : "single";
  const [inputPath, ...rest] = mode === "batch" ? args.slice(1) : args;
  const outputIndex = rest.indexOf("--out");
  const policyIndex = rest.indexOf("--policy");
  const outputPath = outputIndex >= 0 ? rest[outputIndex + 1] : undefined;
  const policyPath = policyIndex >= 0 ? rest[policyIndex + 1] : undefined;
  const expectedArgs = policyPath ? 4 : 2;
  if (!inputPath || !outputPath || outputIndex < 0 || ![2, 4].includes(rest.length) || rest.length !== expectedArgs) {
    throw new Error("USAGE: document-media-extractor <supported-file> --out <directory> [--policy policy.json]\n   Supports OOXML, OpenDocument, EPUB, CBZ, and ZIP containers.\n   or: document-media-extractor --batch <directory> --out <directory> [--policy policy.json]");
  }
  return { mode, inputPath, outputPath, policyPath };
}

async function createOutputDirectory(outputPath: string): Promise<string> {
  const destination = resolve(outputPath);
  try {
    await access(destination);
    throw new Error("OUTPUT_DIRECTORY_EXISTS");
  } catch (error) {
    if (error instanceof Error && error.message === "OUTPUT_DIRECTORY_EXISTS") throw error;
  }
  await mkdir(destination);
  return destination;
}

async function writeResult(destination: string, result: Awaited<ReturnType<typeof extractDocumentMedia>>): Promise<void> {
  const outputNames = result.assets.map((asset) => asset.exportName);
  if (new Set(outputNames).size !== outputNames.length || outputNames.includes("manifest.json")) {
    throw new Error("OUTPUT_NAME_COLLISION");
  }
  await Promise.all(result.assets.map((asset) => writeFile(resolve(destination, asset.exportName), asset.bytes)));
  await writeFile(resolve(destination, "manifest.json"), `${JSON.stringify(result.manifest, null, 2)}\n`);
}

function batchFolderName(inputName: string, index: number): string {
  const safeStem = basename(inputName, extname(inputName)).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "document";
  return `${String(index).padStart(3, "0")}-${safeStem}`;
}

async function supportedFiles(directory: string): Promise<string[]> {
  const entries = await readdir(directory, { withFileTypes: true });
  return entries
    .filter((entry) => entry.isFile() && [
      ".docx", ".docm", ".dotx", ".dotm",
      ".pptx", ".pptm", ".potx", ".potm",
      ".xlsx", ".xlsm", ".xltx", ".xltm",
      ".odt", ".ods", ".odp", ".epub", ".cbz", ".zip",
    ].includes(extname(entry.name).toLowerCase()))
    .map((entry) => entry.name)
    .sort((left, right) => left.localeCompare(right));
}

export async function runCli(args: string[]): Promise<void> {
  const { mode, inputPath, outputPath, policyPath } = parseCliOptions(args);
  let policy: ExtractionPolicy | undefined;
  if (policyPath) {
    try {
      policy = JSON.parse(await readFile(policyPath, "utf8")) as ExtractionPolicy;
    } catch {
      throw new Error("INVALID_POLICY");
    }
  }
  if (mode === "single") {
    const result = await extractDocumentMedia({
      sourceName: basename(inputPath),
      bytes: new Uint8Array(await readFile(inputPath)),
      policy,
    });
    const destination = await createOutputDirectory(outputPath);
    await writeResult(destination, result);
    return;
  }

  const names = await supportedFiles(inputPath);
  if (names.length === 0) throw new Error("NO_SUPPORTED_INPUTS");
  const destination = await createOutputDirectory(outputPath);
  const items: Array<Record<string, string>> = [];
  for (const [index, name] of names.entries()) {
    const outputDirectory = batchFolderName(name, index + 1);
    try {
      const result = await extractDocumentMedia({
        sourceName: name,
        bytes: new Uint8Array(await readFile(resolve(inputPath, name))),
        policy,
      });
      await mkdir(resolve(destination, outputDirectory));
      await writeResult(resolve(destination, outputDirectory), result);
      items.push({ inputName: name, outputDirectory, status: "success" });
    } catch (error) {
      items.push({ inputName: name, outputDirectory, status: "failure", error: error instanceof Error ? error.message : "UNKNOWN_ERROR" });
    }
  }
  await writeFile(
    resolve(destination, "batch-manifest.json"),
    `${JSON.stringify({ policy: policy ?? {}, items }, null, 2)}\n`,
  );
}

export function isCliEntrypoint(importMetaUrl: string, invocationPath: string | undefined): boolean {
  if (!invocationPath) return false;
  try {
    return realpathSync(invocationPath) === fileURLToPath(importMetaUrl);
  } catch {
    return false;
  }
}

if (isCliEntrypoint(import.meta.url, process.argv[1])) {
  runCli(process.argv.slice(2)).catch((error: unknown) => {
    console.error(error instanceof Error ? error.message : "UNKNOWN_ERROR");
    process.exitCode = 1;
  });
}
