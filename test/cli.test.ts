import assert from "node:assert/strict";
import { execFile as execFileCallback } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import test from "node:test";
import JSZip from "jszip";

const execFile = promisify(execFileCallback);
const PNG_BYTES = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

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
    assert.deepEqual(JSON.parse(await readFile(join(outputPath, "manifest.json"), "utf8")), {
      sourceName: "brief.docx",
      format: "docx",
      assets: [
        {
          sourcePath: "word/media/cover.png",
          originalName: "cover.png",
          mediaType: "image/png",
          byteSize: PNG_BYTES.byteLength,
        },
      ],
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
