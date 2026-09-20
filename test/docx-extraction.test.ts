import assert from "node:assert/strict";
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

test("extracts original DOCX media and reports its provenance", async () => {
  const result = await extractDocumentMedia({
    sourceName: "launch-plan.docx",
    bytes: await makeDocx(),
  });

  assert.equal(result.manifest.format, "docx");
  assert.deepEqual(result.manifest.assets, [
    {
      sourcePath: "word/media/hero.png",
      originalName: "hero.png",
      mediaType: "image/png",
      byteSize: PNG_BYTES.byteLength,
    },
  ]);
  assert.equal(result.assets.length, 1);
  assert.equal(result.assets[0].sourcePath, "word/media/hero.png");
  assert.equal(result.assets[0].originalName, "hero.png");
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
