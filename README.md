# Document Media Extractor

Extract original embedded media from local document containers without uploading files or running a backend. The current release supports OOXML Office documents.

## What it supports

The stable first release supports standard OOXML containers through a Node API and CLI. It is deliberately local-first: no backend, account, upload, or remote URL is involved.

| Input | Embedded media path | Provenance recorded |
| --- | --- | --- |
| DOCX | `word/media/` | Word accessibility descriptions, where relationship mapping is reliable |
| PPTX | `ppt/media/` | slide number and OOXML relationship |
| XLSX | `xl/media/` | workbook, worksheet, drawing, and OOXML relationship |

Legacy binary Office formats (`.doc`, `.ppt`, `.xls`) are intentionally not supported. Convert them to OOXML first.

Need a no-install, browser-only workflow for one document? Use the matching local browser tool: [Word image extractor](https://bulkimagedownload.com/word-extract), [PPT image extractor](https://bulkimagedownload.com/ppt-extract), or [Excel image extractor](https://bulkimagedownload.com/excel-extract).

## Get started from GitHub

```sh
git clone https://github.com/WanderZil/document-media-extractor.git
cd document-media-extractor
pnpm install
pnpm build
node dist/cli.js ./brief.docx --out ./extracted
```

The output folder contains the retained original media files and a `manifest.json`. The CLI refuses to overwrite an existing output directory.

For a policy file:

```sh
node dist/cli.js ./brief.xlsx --out ./extracted --policy ./policy.json
```

To process every supported document directly inside one folder (non-recursive), use batch mode. Each successful input gets its own numbered output folder, and `batch-manifest.json` records both successes and isolated failures.

```sh
node dist/cli.js --batch ./documents --out ./extracted-batch --policy ./policy.json
```

## Node API

```ts
import { readFile } from "node:fs/promises";
import { extractDocumentMedia } from "document-media-extractor";

const result = await extractDocumentMedia({
  sourceName: "brief.xlsx", // .docx, .pptx, and .xlsx are supported
  bytes: new Uint8Array(await readFile("brief.xlsx")),
  policy: {
    minWidth: 320,
    minPixels: 100_000,
    exactDuplicates: "exclude",
    namingTemplate: "{source}-{index}-{name}",
  },
});
```

`result.assets` contains only retained original media bytes. `result.manifest` retains every discovered item with source provenance, SHA-256, measurements, export name, and an `INCLUDED` or exclusion decision. The default policy retains every media file.

Image dimensions are read without decoding pixels for PNG, GIF, and baseline/progressive JPEG. Other media remains extractable, but dimension filters mark it `UNREADABLE_DIMENSIONS` instead of guessing.

`{name}` uses a Word accessibility description only when its OOXML relationship maps it reliably to one media part; otherwise it falls back to the embedded filename. Names are sanitized and collision-resolved deterministically.

### Safety and batches

Callers can opt into bounded work for CI or user-supplied files. Limit errors are explicit (for example, `LIMIT_EXPANDED_BYTES`), and `AbortSignal` cancellation reports `CANCELLED` rather than returning partial results. `extractDocumentMediaBatch(inputs)` isolates failures: successful inputs retain their results while other items report a stable error code.

```ts
policy: {
  limits: {
    maxInputBytes: 50_000_000,
    maxArchiveEntries: 2_000,
    maxExpandedBytes: 200_000_000,
    maxMediaCount: 500,
    maxMediaBytes: 20_000_000,
    maxTotalMediaBytes: 100_000_000,
  },
}
```

## Local review UI

Start a static server from this checkout, then open it in a modern Chromium browser:

```sh
python3 -m http.server 4173 --directory review
```

Choose a CLI output folder in the browser. The UI reads `manifest.json` and the exported files directly from the local file picker. It shows every Manifest decision; only retained files have a selectable preview because excluded originals are not written to the output folder.

It uses Canvas dHash (Hamming distance ≤ 8) to mark possible visual matches among browser-decodable retained images. This is an advisory UI feature, capped at 250 items to keep the browser responsive; it never changes the extractor’s exact-byte duplicate decision or removes files. Reviewers can select retained assets and download a new ZIP containing the original files plus a review annotation in its Manifest.

For the most useful review, extract with the default preservation policy so all media remains available to inspect.

## Scope and privacy

Processing is local. The project does not use a backend, upload documents, require an account, or fetch remote URLs. The browser UI is optional: the Node API and CLI are the primary extraction interfaces, while the UI is for manual visual review before a ZIP export.

## Related workflow

Need images from webpages rather than local Office documents? [Bulk Image Download](https://bulkimagedownload.com) handles URL-based, browser-aware collection.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). All fixtures must be safe to redistribute.

## License

[MIT](LICENSE)
