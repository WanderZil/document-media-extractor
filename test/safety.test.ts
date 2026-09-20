import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import JSZip from "jszip";

import { extractDocumentMedia, extractDocumentMediaBatch } from "../src/index.js";

async function makeDocx(media: Uint8Array): Promise<Uint8Array> {
  const archive = new JSZip();
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("word/document.xml", "<w:document />");
  archive.file("word/media/image.png", media);
  return archive.generateAsync({ type: "uint8array" });
}

test("enforces caller-selected input, archive, and media limits before returning results", async () => {
  const document = await makeDocx(new Uint8Array(64));

  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "brief.docx", bytes: document, policy: { limits: { maxInputBytes: 1 } } }),
    { message: "LIMIT_INPUT_BYTES" },
  );
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "brief.docx", bytes: document, policy: { limits: { maxArchiveEntries: 2 } } }),
    { message: "LIMIT_ARCHIVE_ENTRIES" },
  );
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "brief.docx", bytes: document, policy: { limits: { maxMediaBytes: 8 } } }),
    { message: "LIMIT_MEDIA_BYTES" },
  );
  await assert.rejects(
    () => extractDocumentMedia({ sourceName: "brief.docx", bytes: document, policy: { limits: { maxExpandedBytes: 1 } } }),
    { message: "LIMIT_EXPANDED_BYTES" },
  );
});

test("cancels explicitly and keeps other batch results usable", async () => {
  const valid = { sourceName: "brief.docx", bytes: await makeDocx(new Uint8Array(16)) };
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(() => extractDocumentMedia({ ...valid, signal: controller.signal }), { message: "CANCELLED" });

  const batch = await extractDocumentMediaBatch([
    valid,
    { sourceName: "broken.doc", bytes: new Uint8Array() },
  ]);
  assert.equal(batch[0].status, "success");
  assert.equal(batch[1].status, "failure");
  if (batch[1].status === "failure") assert.equal(batch[1].error.code, "UNSUPPORTED_FORMAT");
});

test("keeps successful DOCX, PPTX, and XLSX fixture results when one batch item fails", async () => {
  const batch = await extractDocumentMediaBatch([
    { sourceName: "ordinary.docx", bytes: new Uint8Array(await readFile(new URL("./fixtures/ordinary.docx", import.meta.url))) },
    { sourceName: "presentation.pptx", bytes: new Uint8Array(await readFile(new URL("./fixtures/presentation.pptx", import.meta.url))) },
    { sourceName: "workbook.xlsx", bytes: new Uint8Array(await readFile(new URL("./fixtures/workbook.xlsx", import.meta.url))) },
    { sourceName: "unsupported.odt", bytes: new Uint8Array() },
  ]);

  assert.deepEqual(batch.map((item) => item.status), ["success", "success", "success", "failure"]);
});
