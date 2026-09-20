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

async function makeXlsx(): Promise<Uint8Array> {
  const archive = new JSZip();
  const repeated = png(400, 200, 1);
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("xl/workbook.xml", `<workbook><sheets><sheet name="Overview" r:id="rId1" /><sheet name="Archive" r:id="rId2" /></sheets></workbook>`);
  archive.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml" /><Relationship Id="rId2" Target="worksheets/sheet2.xml" /></Relationships>`,
  );
  archive.file("xl/worksheets/sheet1.xml", `<worksheet><drawing r:id="rId1" /></worksheet>`);
  archive.file("xl/worksheets/sheet2.xml", `<worksheet><drawing r:id="rId2" /></worksheet>`);
  archive.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<Relationships><Relationship Id="rId1" Target="../drawings/drawing1.xml" /></Relationships>`,
  );
  archive.file(
    "xl/worksheets/_rels/sheet2.xml.rels",
    `<Relationships><Relationship Id="rId2" Target="../drawings/drawing2.xml" /></Relationships>`,
  );
  archive.file(
    "xl/drawings/drawing1.xml",
    `<xdr:wsDr><xdr:twoCellAnchor><xdr:cNvPr descr="Revenue chart" /><a:blip r:embed="rId1" /></xdr:twoCellAnchor></xdr:wsDr>`,
  );
  archive.file(
    "xl/drawings/drawing2.xml",
    `<xdr:wsDr><xdr:twoCellAnchor><xdr:cNvPr descr="" /><a:blip r:embed="rId2" /></xdr:twoCellAnchor></xdr:wsDr>`,
  );
  archive.file(
    "xl/drawings/_rels/drawing1.xml.rels",
    `<Relationships><Relationship Id="rId1" Target="../media/copy.png" /></Relationships>`,
  );
  archive.file(
    "xl/drawings/_rels/drawing2.xml.rels",
    `<Relationships><Relationship Id="rId2" Target="../media/hero image.png" /></Relationships>`,
  );
  archive.file("xl/media/copy.png", repeated);
  archive.file("xl/media/hero image.png", repeated);
  archive.file("xl/media/small.png", png(10, 10, 2));
  return archive.generateAsync({ type: "uint8array" });
}

test("applies the common policy to XLSX media and records workbook provenance", async () => {
  const result = await extractDocumentMedia({
    sourceName: "launch.xlsx",
    bytes: await makeXlsx(),
    policy: { minWidth: 100, exactDuplicates: "exclude", namingTemplate: "{source}-{index}-{name}" },
  });

  assert.equal(result.manifest.format, "xlsx");
  assert.deepEqual(result.assets.map((asset) => asset.exportName), ["launch-1-revenue-chart.png"]);
  assert.deepEqual(
    result.manifest.assets.map((asset) => [asset.originalName, asset.included, asset.reason]),
    [
      ["copy.png", true, "INCLUDED"],
      ["hero image.png", false, "EXACT_DUPLICATE"],
      ["small.png", false, "MIN_WIDTH"],
    ],
  );
  assert.deepEqual(result.manifest.assets[0].references, [
    {
      partPath: "xl/drawings/drawing1.xml",
      relationshipId: "rId1",
      workbookPartPath: "xl/workbook.xml",
      worksheetName: "Overview",
      worksheetPartPath: "xl/worksheets/sheet1.xml",
    },
  ]);
});

test("keeps the multi-worksheet XLSX policy fixture reproducible", async () => {
  const result = await extractDocumentMedia({
    sourceName: "launch.xlsx",
    bytes: new Uint8Array(await readFile(new URL("./fixtures/workbook.xlsx", import.meta.url))),
    policy: { minWidth: 100, exactDuplicates: "exclude", namingTemplate: "{source}-{index}-{name}" },
  });

  assert.deepEqual(result.assets.map((asset) => asset.exportName), ["launch-1-revenue-chart.png"]);
  assert.equal(result.manifest.assets[0].references?.[0]?.worksheetName, "Overview");
});
