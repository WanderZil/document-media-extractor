import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const fixtureDirectory = join(dirname(fileURLToPath(import.meta.url)), "../test/fixtures");
const PNG_BYTES = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

async function createDocx(name: string, includeMedia: boolean): Promise<void> {
  const archive = new JSZip();
  archive.file("[Content_Types].xml", "<Types />");
  archive.file("word/document.xml", "<w:document />");
  if (includeMedia) archive.file("word/media/hero.png", PNG_BYTES);
  await writeFile(join(fixtureDirectory, name), await archive.generateAsync({ type: "uint8array" }));
}

await mkdir(fixtureDirectory, { recursive: true });
await createDocx("ordinary.docx", true);
await createDocx("empty.docx", false);
await writeFile(join(fixtureDirectory, "malformed.docx"), PNG_BYTES);
