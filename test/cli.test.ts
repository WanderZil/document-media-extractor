import assert from "node:assert/strict";
import { execFile as execFileCallback } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import test from "node:test";
import JSZip from "jszip";

import { isCliEntrypoint } from "../src/cli.js";

const execFile = promisify(execFileCallback);
const PNG_BYTES = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

test("recognizes the npm bin symlink as the CLI entrypoint", async () => {
  const directory = await mkdtemp(join(tmpdir(), "document-media-extractor-bin-"));
  const source = join(process.cwd(), "src", "cli.ts");
  const bin = join(directory, "document-media-extractor");
  try {
    await symlink(source, bin);
    assert.equal(isCliEntrypoint(new URL("../src/cli.ts", import.meta.url).href, bin), true);
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("CLI writes original media and a manifest for a local DOCX", async () => {
  const directory = await mkdtemp(join(tmpdir(), "document-media-extractor-"));
  const inputPath = join(directory, "brief.docx");
  const outputPath = join(directory, "output");
  const archive = new JSZip();
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("word/document.xml", "<w:document />");
  archive.file("word/media/cover.png", PNG_BYTES);
  await writeFile(inputPath, await archive.generateAsync({ type: "uint8array" }));

  try {
    await execFile(join(process.cwd(), "node_modules/.bin/tsx"), [
      "src/cli.ts",
      inputPath,
      "--out",
      outputPath,
    ]);

    assert.deepEqual(new Uint8Array(await readFile(join(outputPath, "cover.png"))), PNG_BYTES);
    const manifest = JSON.parse(await readFile(join(outputPath, "manifest.json"), "utf8"));
    assert.equal(manifest.sourceName, "brief.docx");
    assert.equal(manifest.format, "docx");
    assert.deepEqual(manifest.policy, {});
    assert.deepEqual(manifest.assets[0], {
      sourcePath: "word/media/cover.png",
      originalName: "cover.png",
      exportName: "cover.png",
      mediaType: "image/png",
      byteSize: PNG_BYTES.byteLength,
      sha256: manifest.assets[0].sha256,
      included: true,
      reason: "INCLUDED",
    });
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("CLI refuses to overwrite an existing output directory", async () => {
  const directory = await mkdtemp(join(tmpdir(), "document-media-extractor-"));
  const inputPath = join(directory, "brief.docx");
  const outputPath = join(directory, "output");
  const archive = new JSZip();
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("word/document.xml", "<w:document />");
  archive.file("word/media/cover.png", PNG_BYTES);
  await writeFile(inputPath, await archive.generateAsync({ type: "uint8array" }));
  await writeFile(outputPath, "existing output");

  try {
    await assert.rejects(
      () =>
        execFile(join(process.cwd(), "node_modules/.bin/tsx"), [
          "src/cli.ts",
          inputPath,
          "--out",
          outputPath,
        ]),
      /OUTPUT_DIRECTORY_EXISTS/,
    );
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("CLI batch mode isolates failures and writes one output folder per supported document", async () => {
  const directory = await mkdtemp(join(tmpdir(), "document-media-extractor-"));
  const inputDirectory = join(directory, "inputs");
  const outputPath = join(directory, "output");
  const archive = new JSZip();
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("word/document.xml", "<w:document />");
  archive.file("word/media/cover.png", PNG_BYTES);
  await mkdir(inputDirectory);
  await writeFile(join(inputDirectory, "brief.docx"), await archive.generateAsync({ type: "uint8array" }));
  await writeFile(join(inputDirectory, "empty.pptx"), await new JSZip().generateAsync({ type: "uint8array" }));
  await writeFile(join(inputDirectory, "ignored.txt"), "not an OOXML document");

  try {
    await execFile(join(process.cwd(), "node_modules/.bin/tsx"), [
      "src/cli.ts",
      "--batch",
      inputDirectory,
      "--out",
      outputPath,
    ]);
    const batchManifest = JSON.parse(await readFile(join(outputPath, "batch-manifest.json"), "utf8"));
    assert.deepEqual(batchManifest.items, [
      { inputName: "brief.docx", outputDirectory: "001-brief", status: "success" },
      { inputName: "empty.pptx", outputDirectory: "002-empty", status: "failure", error: "INVALID_PPTX" },
    ]);
    assert.deepEqual(new Uint8Array(await readFile(join(outputPath, "001-brief", "cover.png"))), PNG_BYTES);
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});
