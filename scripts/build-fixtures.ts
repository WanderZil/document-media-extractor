import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const fixtureDirectory = join(dirname(fileURLToPath(import.meta.url)), "../test/fixtures");
const PNG_BYTES = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const FIXTURE_DATE = new Date("2020-01-01T00:00:00.000Z");

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
  archive.file("[Content_Types].xml", "<Types />", { date: FIXTURE_DATE });
  archive.file("word/document.xml", "<w:document />", { date: FIXTURE_DATE });
  if (includeMedia) archive.file("word/media/hero.png", PNG_BYTES, { date: FIXTURE_DATE });
  await writeFile(join(fixtureDirectory, name), await archive.generateAsync({ type: "uint8array" }));
}

async function createPolicyDocx(): Promise<void> {
  const archive = new JSZip();
  const image = png(400, 200, 1);
  archive.file("[Content_Types].xml", "<Types />", { date: FIXTURE_DATE });
  archive.file(
    "word/document.xml",
    `<w:document><w:drawing><wp:docPr descr="Board overview" /><a:blip r:embed="rId1" /></w:drawing></w:document>`,
    { date: FIXTURE_DATE },
  );
  archive.file(
    "word/_rels/document.xml.rels",
    `<Relationships><Relationship Id="rId1" Target="media/copy.png" /></Relationships>`,
    { date: FIXTURE_DATE },
  );
  archive.file("word/media/copy.png", image, { date: FIXTURE_DATE });
  archive.file("word/media/hero image.png", image, { date: FIXTURE_DATE });
  archive.file("word/media/raw file.png", png(100, 100, 2), { date: FIXTURE_DATE });
  archive.file("word/media/small.png", png(10, 10, 3), { date: FIXTURE_DATE });
  await writeFile(
    join(fixtureDirectory, "policy.docx"),
    await archive.generateAsync({ type: "uint8array" }),
  );
}

await mkdir(fixtureDirectory, { recursive: true });
await createDocx("ordinary.docx", true);
await createDocx("empty.docx", false);
await createPolicyDocx();
await writeFile(join(fixtureDirectory, "malformed.docx"), PNG_BYTES);
