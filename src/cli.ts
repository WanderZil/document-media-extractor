import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, resolve } from "node:path";

import { extractDocumentMedia } from "./index.js";

type CliOptions = {
  inputPath: string;
  outputPath: string;
};

function parseCliOptions(args: string[]): CliOptions {
  const [inputPath, ...rest] = args;
  const outputIndex = rest.indexOf("--out");
  const outputPath = outputIndex >= 0 ? rest[outputIndex + 1] : undefined;
  if (!inputPath || !outputPath || outputIndex + 2 !== rest.length) {
    throw new Error("USAGE: document-media-extractor <file.docx> --out <directory>");
  }
  return { inputPath, outputPath };
}

export async function runCli(args: string[]): Promise<void> {
  const { inputPath, outputPath } = parseCliOptions(args);
  const result = await extractDocumentMedia({
    sourceName: basename(inputPath),
    bytes: new Uint8Array(await readFile(inputPath)),
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
