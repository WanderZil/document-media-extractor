import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import JSZip from "jszip";

import { extractDocumentMedia } from "../src/index.js";

function png(width: number, height: number, suffix: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  new DataView(bytes.buffer).setUint32(16, width);
  new DataView(bytes.buffer).setUint32(20, height);
  bytes[32] = suffix;
  return bytes;
}

async function makePptx(): Promise<Uint8Array> {
  const archive = new JSZip();
  const repeated = png(400, 200, 1);
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("ppt/presentation.xml", "<p:presentation />");
  archive.file(
    "ppt/slides/slide1.xml",
    `<p:sld><p:pic><p:nvPicPr><p:cNvPr descr="Cover graphic" /></p:nvPicPr><a:blip r:embed="rId1" /></p:pic></p:sld>`,
  );
  archive.file(
    "ppt/slides/_rels/slide1.xml.rels",
    `<Relationships><Relationship Id="rId1" Target="../media/copy.png" /></Relationships>`,
  );
  archive.file(
    "ppt/slides/slide2.xml",
    `<p:sld><p:pic><p:nvPicPr><p:cNvPr descr="" /></p:nvPicPr><a:blip r:embed="rId2" /></p:pic></p:sld>`,
  );
  archive.file(
    "ppt/slides/_rels/slide2.xml.rels",
    `<Relationships><Relationship Id="rId2" Target="../media/hero image.png" /></Relationships>`,
  );
  archive.file("ppt/media/copy.png", repeated);
  archive.file("ppt/media/hero image.png", repeated);
  archive.file("ppt/media/small.png", png(10, 10, 2));
  return archive.generateAsync({ type: "uint8array" });
}

test("applies the common policy to PPTX media and records slide provenance", async () => {
  const result = await extractDocumentMedia({
    sourceName: "launch.pptx",
    bytes: await makePptx(),
    policy: { minWidth: 100, exactDuplicates: "exclude", namingTemplate: "{source}-{index}-{name}" },
  });

  assert.equal(result.manifest.format, "pptx");
  assert.deepEqual(result.assets.map((asset) => asset.exportName), ["launch-1-cover-graphic.png"]);
  assert.deepEqual(
    result.manifest.assets.map((asset) => [asset.originalName, asset.included, asset.reason]),
    [
      ["copy.png", true, "INCLUDED"],
      ["hero image.png", false, "EXACT_DUPLICATE"],
      ["small.png", false, "MIN_WIDTH"],
    ],
  );
  assert.equal(result.manifest.assets[0].accessibleDescription, "Cover graphic");
  assert.deepEqual(result.manifest.assets[0].references, [
    { partPath: "ppt/slides/slide1.xml", relationshipId: "rId1", slideNumber: 1 },
  ]);
});

test("reports malformed and no-media PPTX files with stable errors", async () => {
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "broken.pptx", bytes: png(1, 1, 1) }),
    { message: "INVALID_ARCHIVE" },
  );

  const empty = new JSZip();
  empty.file("[Content_Types].xml", "<Types />");
  empty.file("ppt/presentation.xml", "<p:presentation />");
  const emptyBytes = await empty.generateAsync({ type: "uint8array" });
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "empty.pptx", bytes: emptyBytes }),
    { message: "NO_MEDIA" },
  );
});

test("keeps the multi-slide PPTX policy fixture reproducible", async () => {
  const result = await extractDocumentMedia({
    sourceName: "launch.pptx",
    bytes: new Uint8Array(await readFile(new URL("./fixtures/presentation.pptx", import.meta.url))),
    policy: { minWidth: 100, exactDuplicates: "exclude", namingTemplate: "{source}-{index}-{name}" },
  });

  assert.deepEqual(result.assets.map((asset) => asset.exportName), ["launch-1-cover-graphic.png"]);
  assert.deepEqual(result.manifest.assets[0].references, [
    { partPath: "ppt/slides/slide1.xml", relationshipId: "rId1", slideNumber: 1 },
  ]);
});
