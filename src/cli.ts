#!/usr/bin/env node
import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, resolve } from "node:path";

import { extractDocumentMedia, type ExtractionPolicy } from "./index.js";

type CliOptions = {
  inputPath: string;
  outputPath: string;
  policyPath?: string;
};

function parseCliOptions(args: string[]): CliOptions {
  const [inputPath, ...rest] = args;
  const outputIndex = rest.indexOf("--out");
  const policyIndex = rest.indexOf("--policy");
  const outputPath = outputIndex >= 0 ? rest[outputIndex + 1] : undefined;
  const policyPath = policyIndex >= 0 ? rest[policyIndex + 1] : undefined;
  const expectedArgs = policyPath ? 4 : 2;
  if (!inputPath || !outputPath || outputIndex < 0 || ![2, 4].includes(rest.length) || rest.length !== expectedArgs) {
    throw new Error("USAGE: document-media-extractor <file.docx|file.pptx|file.xlsx> --out <directory> [--policy policy.json]");
  }
  return { inputPath, outputPath, policyPath };
}

export async function runCli(args: string[]): Promise<void> {
  const { inputPath, outputPath, policyPath } = parseCliOptions(args);
  let policy: ExtractionPolicy | undefined;
  if (policyPath) {
    try {
      policy = JSON.parse(await readFile(policyPath, "utf8")) as ExtractionPolicy;
    } catch {
      throw new Error("INVALID_POLICY");
    }
  }
  const result = await extractDocumentMedia({
    sourceName: basename(inputPath),
    bytes: new Uint8Array(await readFile(inputPath)),
    policy,
  });
  const destination = resolve(outputPath);
  try {
    await access(destination);
    throw new Error("OUTPUT_DIRECTORY_EXISTS");
  } catch (error) {
    if (error instanceof Error && error.message === "OUTPUT_DIRECTORY_EXISTS") throw error;
  }
  const outputNames = result.assets.map((asset) => asset.exportName);
  if (new Set(outputNames).size !== outputNames.length || outputNames.includes("manifest.json")) {
    throw new Error("OUTPUT_NAME_COLLISION");
  }
  await mkdir(destination);
  await Promise.all(result.assets.map((asset) => writeFile(resolve(destination, asset.exportName), asset.bytes)));
  await writeFile(
    resolve(destination, "manifest.json"),
    `${JSON.stringify(result.manifest, null, 2)}\n`,
  );
}

if (import.meta.url === new URL(process.argv[1], "file:").href) {
  runCli(process.argv.slice(2)).catch((error: unknown) => {
    console.error(error instanceof Error ? error.message : "UNKNOWN_ERROR");
    process.exitCode = 1;
  });
}
