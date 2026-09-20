# Document Media Extractor

Extract original embedded media from local Office documents without uploading files or running a backend.

## Phase 1 status

The development build supports standard DOCX and PPTX extraction through a Node API and CLI, including local filtering, exact-content duplicate policy, SHA-256 provenance, deterministic names, and a decision-bearing Manifest. XLSX, visual duplicate review, and the static local review UI are planned next.

## Node API

```ts
import { readFile } from "node:fs/promises";
import { extractDocumentMedia } from "document-media-extractor";

const result = await extractDocumentMedia({
  sourceName: "brief.pptx", // .docx and .pptx are supported
  bytes: new Uint8Array(await readFile("brief.pptx")),
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

## Development CLI

```sh
pnpm exec tsx src/cli.ts brief.pptx --out ./extracted
```

This writes the extracted media and `manifest.json` to the chosen directory.

## Scope and privacy

Processing is local. The project does not use a backend, upload documents, require an account, or fetch remote URLs.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). All fixtures must be safe to redistribute.

## License

[MIT](LICENSE)
