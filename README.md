# Document Media Extractor

Extract original embedded media from local Office documents without uploading files or running a backend.

## Phase 1 status

The development build supports standard DOCX extraction through a Node API and CLI. PPTX, XLSX, policy controls, visual duplicate review, and the static local review UI are planned next.

## Node API

```ts
import { readFile } from "node:fs/promises";
import { extractDocumentMedia } from "document-media-extractor";

const result = await extractDocumentMedia({
  sourceName: "brief.docx",
  bytes: new Uint8Array(await readFile("brief.docx")),
});
```

`result.assets` contains original media bytes. `result.manifest` identifies the source document, OOXML media paths, filenames, media types, and byte sizes.

## Development CLI

```sh
pnpm exec tsx src/cli.ts brief.docx --out ./extracted
```

This writes the extracted media and `manifest.json` to the chosen directory.

## Scope and privacy

Processing is local. The project does not use a backend, upload documents, require an account, or fetch remote URLs.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). All fixtures must be safe to redistribute.

## License

[MIT](LICENSE)
