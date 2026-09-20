# Document Media Extractor

Extract original embedded media from local Office documents without uploading files or running a backend.

## Phase 1 status

The development build supports standard DOCX, PPTX, and XLSX extraction through a Node API and CLI, including local filtering, exact-content duplicate policy, SHA-256 provenance, deterministic names, bounded batches, and a decision-bearing Manifest. A static browser review surface is included for manual visual review.

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

`result.assets` contains only retained original media bytes. `result.manifest` retains every discovered item with source provenance, measurements where readable, SHA-256, export name, and an `INCLUDED` or exclusion decision. The default policy retains every media file.

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

## Development CLI

```sh
pnpm exec tsx src/cli.ts brief.xlsx --out ./extracted
```

This writes the extracted media and `manifest.json` to the chosen directory.

## Local review UI

Open [`review/index.html`](review/index.html) in a modern Chromium browser, then choose a CLI output folder. The UI reads `manifest.json` and the exported files directly from the local file picker. It uses Canvas dHash to mark possible visual matches among browser-decodable images; these are review hints only and never remove or upload files.

For the most useful review, extract with the default preservation policy so all media remains available to inspect.

## Scope and privacy

Processing is local. The project does not use a backend, upload documents, require an account, or fetch remote URLs.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). All fixtures must be safe to redistribute.

## License

[MIT](LICENSE)
