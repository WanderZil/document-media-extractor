import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import JSZip from "jszip";

import { extractDocumentMedia } from "../src/index.js";

const PNG_BYTES = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

async function makeDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types />");
  zip.file("word/document.xml", "<w:document />");
  zip.file("word/media/hero.png", PNG_BYTES);
  return zip.generateAsync({ type: "uint8array" });
}

function png(width: number, height: number, suffix: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  new DataView(bytes.buffer).setUint32(16, width);
  new DataView(bytes.buffer).setUint32(20, height);
  bytes[32] = suffix;
  return bytes;
}

function gif(width: number, height: number): Uint8Array {
  return new Uint8Array([71, 73, 70, 56, 57, 97, width & 255, width >> 8, height & 255, height >> 8]);
}

function jpeg(width: number, height: number): Uint8Array {
  return new Uint8Array([255, 216, 255, 192, 0, 17, 8, height >> 8, height & 255, width >> 8, width & 255, 3, 1, 17, 0, 2, 17, 0, 3, 17, 0, 255, 217]);
}

test("extracts original DOCX media and reports its provenance", async () => {
  const result = await extractDocumentMedia({
    sourceName: "launch-plan.docx",
    bytes: await makeDocx(),
  });

  assert.equal(result.manifest.format, "docx");
  assert.deepEqual(result.manifest.assets[0], {
    sourcePath: "word/media/hero.png",
    originalName: "hero.png",
    exportName: "hero.png",
    mediaType: "image/png",
    byteSize: PNG_BYTES.byteLength,
    sha256: result.manifest.assets[0].sha256,
    included: true,
    reason: "INCLUDED",
  });
  assert.equal(result.assets.length, 1);
  assert.equal(result.assets[0].sourcePath, "word/media/hero.png");
  assert.equal(result.assets[0].originalName, "hero.png");
  assert.equal(result.assets[0].exportName, "hero.png");
  assert.equal(result.assets[0].mediaType, "image/png");
  assert.deepEqual(result.assets[0].bytes, PNG_BYTES);
});

test("reports unsupported, malformed, and empty documents with stable errors", async () => {
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "legacy.doc", bytes: PNG_BYTES }),
    { message: "UNSUPPORTED_FORMAT" },
  );
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "broken.docx", bytes: PNG_BYTES }),
    { message: "INVALID_ARCHIVE" },
  );

  const emptyDocx = await new JSZip().generateAsync({ type: "uint8array" });
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "empty.docx", bytes: emptyDocx }),
    { message: "INVALID_DOCX" },
  );

  const structurallyEmptyDocx = new JSZip();
  structurallyEmptyDocx.file("[Content_Types].xml", "<Types />");
  structurallyEmptyDocx.file("word/document.xml", "<w:document />");
  const structurallyEmptyBytes = await structurallyEmptyDocx.generateAsync({ type: "uint8array" });
  await assert.rejects(
    () =>
      extractDocumentMedia({
        sourceName: "empty.docx",
        bytes: structurallyEmptyBytes,
      }),
    { message: "NO_MEDIA" },
  );
});

test("applies filters, exact duplicate policy, and deterministic names with auditable decisions", async () => {
  const archive = new JSZip();
  const original = png(400, 200, 1);
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("word/document.xml", "<w:document />");
  archive.file("word/media/hero image.png", original);
  archive.file("word/media/copy.png", original);
  archive.file("word/media/small.png", png(10, 10, 2));

  const result = await extractDocumentMedia({
    sourceName: "launch plan.docx",
    bytes: await archive.generateAsync({ type: "uint8array" }),
    policy: {
      minWidth: 100,
      exactDuplicates: "exclude",
      namingTemplate: "{source}-{index}-{name}",
    },
  });

  assert.deepEqual(result.assets.map((asset) => asset.exportName), ["launch-plan-1-copy.png"]);
  assert.deepEqual(
    result.manifest.assets.map((asset) => [asset.originalName, asset.included, asset.reason]),
    [
      ["copy.png", true, "INCLUDED"],
      ["hero image.png", false, "EXACT_DUPLICATE"],
      ["small.png", false, "MIN_WIDTH"],
    ],
  );
  assert.equal(result.manifest.assets[0].width, 400);
  assert.equal(result.manifest.assets[0].height, 200);
  assert.match(result.manifest.assets[0].sha256, /^[a-f0-9]{64}$/);
});

test("measures PNG, GIF, and JPEG dimensions for the common filter contract", async () => {
  const archive = new JSZip();
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("word/document.xml", "<w:document />");
  archive.file("word/media/animated.gif", gif(120, 40));
  archive.file("word/media/photo.jpg", jpeg(320, 240));
  archive.file("word/media/graphic.png", png(200, 100, 7));

  const result = await extractDocumentMedia({
    sourceName: "formats.docx",
    bytes: await archive.generateAsync({ type: "uint8array" }),
    policy: { minWidth: 100 },
  });

  assert.deepEqual(result.manifest.assets.map((asset) => [asset.originalName, asset.width, asset.height, asset.reason]), [
    ["animated.gif", 120, 40, "INCLUDED"],
    ["graphic.png", 200, 100, "INCLUDED"],
    ["photo.jpg", 320, 240, "INCLUDED"],
  ]);
});

test("uses a mapped OOXML accessibility description and otherwise falls back to the media name", async () => {
  const archive = new JSZip();
  archive.file("[Content_Types].xml", "<Types />");
  archive.file(
    "word/document.xml",
    `<w:document><w:drawing><wp:docPr descr="Board overview" /><a:blip r:embed="rId1" /></w:drawing></w:document>`,
  );
  archive.file(
    "word/_rels/document.xml.rels",
    `<Relationships><Relationship Id="rId1" Target="media/image1.png" /></Relationships>`,
  );
  archive.file("word/media/image1.png", png(100, 100, 1));
  archive.file("word/media/raw file.png", png(100, 100, 2));

  const result = await extractDocumentMedia({
    sourceName: "brief.docx",
    bytes: await archive.generateAsync({ type: "uint8array" }),
  });

  assert.deepEqual(result.assets.map((asset) => asset.exportName), ["board-overview.png", "raw-file.png"]);
  assert.equal(result.manifest.assets[0].accessibleDescription, "Board overview");
  assert.equal(result.manifest.assets[1].accessibleDescription, undefined);
});

test("keeps the policy contract reproducible against the checked-in fixture", async () => {
  const result = await extractDocumentMedia({
    sourceName: "launch plan.docx",
    bytes: new Uint8Array(await readFile(new URL("./fixtures/policy.docx", import.meta.url))),
    policy: { minWidth: 100, exactDuplicates: "exclude", namingTemplate: "{source}-{index}-{name}" },
  });

  assert.deepEqual(result.assets.map((asset) => asset.exportName), [
    "launch-plan-1-board-overview.png",
    "launch-plan-2-raw-file.png",
  ]);
  assert.deepEqual(
    result.manifest.assets.map((asset) => [asset.originalName, asset.included, asset.reason]),
    [
      ["copy.png", true, "INCLUDED"],
      ["hero image.png", false, "EXACT_DUPLICATE"],
      ["raw file.png", true, "INCLUDED"],
      ["small.png", false, "MIN_WIDTH"],
    ],
  );
});
