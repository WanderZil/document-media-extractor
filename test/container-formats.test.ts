import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";

import { extractDocumentMedia } from "../src/index.js";

function png(marker: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  new DataView(bytes.buffer).setUint32(16, 400);
  new DataView(bytes.buffer).setUint32(20, 200);
  bytes[32] = marker;
  return bytes;
}

async function bytesFor(archive: JSZip): Promise<Uint8Array> {
  return archive.generateAsync({ type: "uint8array" });
}

test("extracts EPUB and OpenDocument image paths using the shared policy contract", async () => {
  const epub = new JSZip();
  epub.file("mimetype", "application/epub+zip");
  epub.file("META-INF/container.xml", "<container />");
  epub.file("OEBPS/images/cover.png", png(1));
  epub.file("OEBPS/images/duplicate.png", png(1));

  const epubResult = await extractDocumentMedia({
    sourceName: "book.epub",
    bytes: await bytesFor(epub),
    policy: { exactDuplicates: "exclude" },
  });
  assert.equal(epubResult.manifest.format, "epub");
  assert.deepEqual(epubResult.manifest.assets.map((asset) => [asset.sourcePath, asset.reason]), [
    ["OEBPS/images/cover.png", "INCLUDED"],
    ["OEBPS/images/duplicate.png", "EXACT_DUPLICATE"],
  ]);

  const odt = new JSZip();
  odt.file("mimetype", "application/vnd.oasis.opendocument.text");
  odt.file("content.xml", "<office:document-content />");
  odt.file("Pictures/photo.png", png(2));
  const odtResult = await extractDocumentMedia({ sourceName: "notes.odt", bytes: await bytesFor(odt) });
  assert.equal(odtResult.manifest.format, "odt");
  assert.equal(odtResult.assets[0]?.sourcePath, "Pictures/photo.png");
});

test("orders CBZ pages naturally and scans image files in ordinary ZIP archives", async () => {
  const cbz = new JSZip();
  cbz.file("10.png", png(10));
  cbz.file("2.png", png(2));
  cbz.file("ComicInfo.xml", "<ComicInfo />");
  const cbzResult = await extractDocumentMedia({ sourceName: "issue.cbz", bytes: await bytesFor(cbz) });
  assert.equal(cbzResult.manifest.format, "cbz");
  assert.deepEqual(cbzResult.assets.map((asset) => asset.originalName), ["2.png", "10.png"]);

  const archive = new JSZip();
  archive.file("assets/cover.png", png(3));
  archive.file("__MACOSX/noise.png", png(4));
  const archiveResult = await extractDocumentMedia({ sourceName: "assets.zip", bytes: await bytesFor(archive) });
  assert.equal(archiveResult.manifest.format, "zip");
  assert.deepEqual(archiveResult.assets.map((asset) => asset.sourcePath), ["assets/cover.png"]);
});

test("rejects containers that do not match their claimed EPUB or OpenDocument format", async () => {
  const archive = new JSZip();
  archive.file("image.png", png(1));
  const bytes = await bytesFor(archive);

  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "not-an-ebook.epub", bytes }),
    { message: "INVALID_EPUB" },
  );
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "not-a-document.odp", bytes }),
    { message: "INVALID_ODF" },
  );
});
