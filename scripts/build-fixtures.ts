import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const fixtureDirectory = join(dirname(fileURLToPath(import.meta.url)), "../test/fixtures");
const PNG_BYTES = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const FIXTURE_DATE = new Date("2020-01-01T00:00:00.000Z");
const FIXTURE_FILE_OPTIONS = { createFolders: false, date: FIXTURE_DATE };

function png(width: number, height: number, suffix: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  new DataView(bytes.buffer).setUint32(16, width);
  new DataView(bytes.buffer).setUint32(20, height);
  bytes[32] = suffix;
  return bytes;
}

async function createDocx(name: string, includeMedia: boolean): Promise<void> {
  const archive = new JSZip();
  archive.file("[Content_Types].xml", "<Types />", FIXTURE_FILE_OPTIONS);
  archive.file("word/document.xml", "<w:document />", FIXTURE_FILE_OPTIONS);
  if (includeMedia) archive.file("word/media/hero.png", PNG_BYTES, FIXTURE_FILE_OPTIONS);
  await writeFile(join(fixtureDirectory, name), await archive.generateAsync({ type: "uint8array" }));
}

async function createPolicyDocx(): Promise<void> {
  const archive = new JSZip();
  const image = png(400, 200, 1);
  archive.file("[Content_Types].xml", "<Types />", FIXTURE_FILE_OPTIONS);
  archive.file(
    "word/document.xml",
    `<w:document><w:drawing><wp:docPr descr="Board overview" /><a:blip r:embed="rId1" /></w:drawing></w:document>`,
    FIXTURE_FILE_OPTIONS,
  );
  archive.file(
    "word/_rels/document.xml.rels",
    `<Relationships><Relationship Id="rId1" Target="media/copy.png" /></Relationships>`,
    FIXTURE_FILE_OPTIONS,
  );
  archive.file("word/media/copy.png", image, FIXTURE_FILE_OPTIONS);
  archive.file("word/media/hero image.png", image, FIXTURE_FILE_OPTIONS);
  archive.file("word/media/raw file.png", png(100, 100, 2), FIXTURE_FILE_OPTIONS);
  archive.file("word/media/small.png", png(10, 10, 3), FIXTURE_FILE_OPTIONS);
  await writeFile(
    join(fixtureDirectory, "policy.docx"),
    await archive.generateAsync({ type: "uint8array" }),
  );
}

async function createPptx(name: string, includeMedia: boolean): Promise<void> {
  const archive = new JSZip();
  const repeated = png(400, 200, 1);
  archive.file("[Content_Types].xml", "<Types />", FIXTURE_FILE_OPTIONS);
  archive.file("ppt/presentation.xml", "<p:presentation />", FIXTURE_FILE_OPTIONS);
  if (includeMedia) {
    archive.file(
      "ppt/slides/slide1.xml",
      `<p:sld><p:pic><p:nvPicPr><p:cNvPr descr="Cover graphic" /></p:nvPicPr><a:blip r:embed="rId1" /></p:pic></p:sld>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file(
      "ppt/slides/_rels/slide1.xml.rels",
      `<Relationships><Relationship Id="rId1" Target="../media/copy.png" /></Relationships>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file(
      "ppt/slides/slide2.xml",
      `<p:sld><p:pic><p:nvPicPr><p:cNvPr descr="" /></p:nvPicPr><a:blip r:embed="rId2" /></p:pic></p:sld>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file(
      "ppt/slides/_rels/slide2.xml.rels",
      `<Relationships><Relationship Id="rId2" Target="../media/hero image.png" /></Relationships>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file("ppt/media/copy.png", repeated, FIXTURE_FILE_OPTIONS);
    archive.file("ppt/media/hero image.png", repeated, FIXTURE_FILE_OPTIONS);
    archive.file("ppt/media/small.png", png(10, 10, 2), FIXTURE_FILE_OPTIONS);
  }
  await writeFile(join(fixtureDirectory, name), await archive.generateAsync({ type: "uint8array" }));
}

async function createXlsx(name: string, includeMedia: boolean): Promise<void> {
  const archive = new JSZip();
  const repeated = png(400, 200, 1);
  archive.file("[Content_Types].xml", "<Types />", FIXTURE_FILE_OPTIONS);
  archive.file(
    "xl/workbook.xml",
    `<workbook><sheets><sheet name="Overview" r:id="rId1" /><sheet name="Archive" r:id="rId2" /></sheets></workbook>`,
    FIXTURE_FILE_OPTIONS,
  );
  if (includeMedia) {
    archive.file(
      "xl/_rels/workbook.xml.rels",
      `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml" /><Relationship Id="rId2" Target="worksheets/sheet2.xml" /></Relationships>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file("xl/worksheets/sheet1.xml", `<worksheet><drawing r:id="rId1" /></worksheet>`, FIXTURE_FILE_OPTIONS);
    archive.file("xl/worksheets/sheet2.xml", `<worksheet><drawing r:id="rId2" /></worksheet>`, FIXTURE_FILE_OPTIONS);
    archive.file(
      "xl/worksheets/_rels/sheet1.xml.rels",
      `<Relationships><Relationship Id="rId1" Target="../drawings/drawing1.xml" /></Relationships>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file(
      "xl/worksheets/_rels/sheet2.xml.rels",
      `<Relationships><Relationship Id="rId2" Target="../drawings/drawing2.xml" /></Relationships>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file(
      "xl/drawings/drawing1.xml",
      `<xdr:wsDr><xdr:twoCellAnchor><xdr:cNvPr descr="Revenue chart" /><a:blip r:embed="rId1" /></xdr:twoCellAnchor></xdr:wsDr>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file(
      "xl/drawings/drawing2.xml",
      `<xdr:wsDr><xdr:twoCellAnchor><xdr:cNvPr descr="" /><a:blip r:embed="rId2" /></xdr:twoCellAnchor></xdr:wsDr>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file(
      "xl/drawings/_rels/drawing1.xml.rels",
      `<Relationships><Relationship Id="rId1" Target="../media/copy.png" /></Relationships>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file(
      "xl/drawings/_rels/drawing2.xml.rels",
      `<Relationships><Relationship Id="rId2" Target="../media/hero image.png" /></Relationships>`,
      FIXTURE_FILE_OPTIONS,
    );
    archive.file("xl/media/copy.png", repeated, FIXTURE_FILE_OPTIONS);
    archive.file("xl/media/hero image.png", repeated, FIXTURE_FILE_OPTIONS);
    archive.file("xl/media/small.png", png(10, 10, 2), FIXTURE_FILE_OPTIONS);
  }
  await writeFile(join(fixtureDirectory, name), await archive.generateAsync({ type: "uint8array" }));
}

await mkdir(fixtureDirectory, { recursive: true });
await createDocx("ordinary.docx", true);
await createDocx("empty.docx", false);
await createPolicyDocx();
await writeFile(join(fixtureDirectory, "malformed.docx"), PNG_BYTES);
await createPptx("presentation.pptx", true);
await createPptx("empty.pptx", false);
await writeFile(join(fixtureDirectory, "malformed.pptx"), PNG_BYTES);
await createXlsx("workbook.xlsx", true);
await createXlsx("empty.xlsx", false);
await writeFile(join(fixtureDirectory, "malformed.xlsx"), PNG_BYTES);
