# Changelog

All notable changes to this project are documented in this file.

## 0.1.2 — 2026-09-21

### Documentation

- Align the README with the full supported-format scope.
- Add npm installation and `npx` CLI quick starts.
- Include this changelog in npm packages.

### Continuous integration

- Run tests, type checking, and production builds on GitHub Actions for Node.js 20 and 22.
- Pin pnpm 10.11.1 so CI uses the lockfile-compatible package manager.

## 0.1.1 — 2026-09-20

### Fixed

- Make the CLI run correctly when invoked through npm's `bin` symlink, including via `npx document-media-extractor`.

## 0.1.0 — 2026-09-20

### Added

- Local-first media extraction for OOXML, OpenDocument, EPUB, CBZ, and ZIP containers.
- Policy-based filtering, exact-byte deduplication, manifests, batch processing, and an optional local review UI.
