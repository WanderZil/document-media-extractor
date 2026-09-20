import JSZip from "jszip";

export type DocumentFormat = "docx" | "pptx" | "xlsx";

export type MediaAsset = {
  sourcePath: string;
  originalName: string;
  exportName: string;
  mediaType: string;
  bytes: Uint8Array;
};

export type ExactDuplicatePolicy = "keep" | "exclude";

export type ExtractionPolicy = {
  allowedMediaTypes?: readonly string[];
  minByteSize?: number;
  maxByteSize?: number;
  minWidth?: number;
  maxWidth?: number;
  minHeight?: number;
  maxHeight?: number;
  minPixels?: number;
  maxPixels?: number;
  minAspectRatio?: number;
  maxAspectRatio?: number;
  exactDuplicates?: ExactDuplicatePolicy;
  namingTemplate?: string;
  limits?: ExtractionLimits;
};

export type ExtractionLimits = {
  maxInputBytes?: number;
  maxArchiveEntries?: number;
  maxExpandedBytes?: number;
  maxMediaCount?: number;
  maxMediaBytes?: number;
  maxTotalMediaBytes?: number;
};

export type AssetDecision =
  | "INCLUDED"
  | "MEDIA_TYPE"
  | "MIN_BYTE_SIZE"
  | "MAX_BYTE_SIZE"
  | "UNREADABLE_DIMENSIONS"
  | "MIN_WIDTH"
  | "MAX_WIDTH"
  | "MIN_HEIGHT"
  | "MAX_HEIGHT"
  | "MIN_PIXELS"
  | "MAX_PIXELS"
  | "MIN_ASPECT_RATIO"
  | "MAX_ASPECT_RATIO"
  | "EXACT_DUPLICATE";

export type ManifestAsset = {
  sourcePath: string;
  originalName: string;
  accessibleDescription?: string;
  references?: OoxmlReference[];
  exportName?: string;
  mediaType: string;
  byteSize: number;
  width?: number;
  height?: number;
  pixels?: number;
  aspectRatio?: number;
  sha256: string;
  included: boolean;
  reason: AssetDecision;
};

export type OoxmlReference = {
  partPath: string;
  relationshipId: string;
  slideNumber?: number;
  workbookPartPath?: string;
  worksheetName?: string;
  worksheetPartPath?: string;
};

export type ExtractionManifest = {
  sourceName: string;
  format: DocumentFormat;
  policy: ExtractionPolicy;
  assets: ManifestAsset[];
};

export type ExtractionResult = {
  manifest: ExtractionManifest;
  assets: MediaAsset[];
};

export type ExtractDocumentMediaInput = {
  sourceName: string;
  bytes: Uint8Array;
  policy?: ExtractionPolicy;
  signal?: AbortSignal;
};

export type BatchExtractionResult =
  | { input: ExtractDocumentMediaInput; status: "success"; result: ExtractionResult }
  | { input: ExtractDocumentMediaInput; status: "failure"; error: { code: string } };

function mediaTypeFromName(name: string): string {
  const extension = name.split(".").at(-1)?.toLowerCase();
  const mediaTypes: Record<string, string> = {
    bmp: "image/bmp",
    gif: "image/gif",
    jpeg: "image/jpeg",
    jpg: "image/jpeg",
    png: "image/png",
    svg: "image/svg+xml",
    tif: "image/tiff",
    tiff: "image/tiff",
    webp: "image/webp",
  };
  return mediaTypes[extension ?? ""] ?? "application/octet-stream";
}

function documentFormat(sourceName: string): DocumentFormat {
  if (sourceName.toLowerCase().endsWith(".docx")) return "docx";
  if (sourceName.toLowerCase().endsWith(".pptx")) return "pptx";
  if (sourceName.toLowerCase().endsWith(".xlsx")) return "xlsx";
  throw new Error("UNSUPPORTED_FORMAT");
}

function mediaPrefix(format: DocumentFormat): string {
  if (format === "docx") return "word/media/";
  if (format === "pptx") return "ppt/media/";
  return "xl/media/";
}

function requiredDocumentPart(format: DocumentFormat): string {
  if (format === "docx") return "word/document.xml";
  if (format === "pptx") return "ppt/presentation.xml";
  return "xl/workbook.xml";
}

function invalidDocumentError(format: DocumentFormat): string {
  if (format === "docx") return "INVALID_DOCX";
  if (format === "pptx") return "INVALID_PPTX";
  return "INVALID_XLSX";
}

function assertNotCancelled(signal: AbortSignal | undefined): void {
  if (signal?.aborted) throw new Error("CANCELLED");
}

function assertLimit(value: number, maximum: number | undefined, error: string): void {
  if (maximum !== undefined && value > maximum) throw new Error(error);
}

function pngDimensions(bytes: Uint8Array): { width: number; height: number } | undefined {
  const signature = [137, 80, 78, 71, 13, 10, 26, 10];
  if (
    bytes.byteLength < 24 ||
    !signature.every((value, index) => bytes[index] === value) ||
    String.fromCharCode(...bytes.slice(12, 16)) !== "IHDR"
  ) {
    return undefined;
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return { width: view.getUint32(16), height: view.getUint32(20) };
}

function dimensionsFor(mediaType: string, bytes: Uint8Array): { width: number; height: number } | undefined {
  if (mediaType === "image/png") return pngDimensions(bytes);
  return undefined;
}

function usesDimensionFilter(policy: ExtractionPolicy): boolean {
  return [
    policy.minWidth,
    policy.maxWidth,
    policy.minHeight,
    policy.maxHeight,
    policy.minPixels,
    policy.maxPixels,
    policy.minAspectRatio,
    policy.maxAspectRatio,
  ].some((value) => value !== undefined);
}

function decisionFor(
  mediaType: string,
  bytes: Uint8Array,
  dimensions: { width: number; height: number } | undefined,
  policy: ExtractionPolicy,
): AssetDecision {
  if (policy.allowedMediaTypes && !policy.allowedMediaTypes.includes(mediaType)) return "MEDIA_TYPE";
  if (policy.minByteSize !== undefined && bytes.byteLength < policy.minByteSize) return "MIN_BYTE_SIZE";
  if (policy.maxByteSize !== undefined && bytes.byteLength > policy.maxByteSize) return "MAX_BYTE_SIZE";
  if (!dimensions) return usesDimensionFilter(policy) ? "UNREADABLE_DIMENSIONS" : "INCLUDED";

  const pixels = dimensions.width * dimensions.height;
  const aspectRatio = dimensions.width / dimensions.height;
  if (policy.minWidth !== undefined && dimensions.width < policy.minWidth) return "MIN_WIDTH";
  if (policy.maxWidth !== undefined && dimensions.width > policy.maxWidth) return "MAX_WIDTH";
  if (policy.minHeight !== undefined && dimensions.height < policy.minHeight) return "MIN_HEIGHT";
  if (policy.maxHeight !== undefined && dimensions.height > policy.maxHeight) return "MAX_HEIGHT";
  if (policy.minPixels !== undefined && pixels < policy.minPixels) return "MIN_PIXELS";
  if (policy.maxPixels !== undefined && pixels > policy.maxPixels) return "MAX_PIXELS";
  if (policy.minAspectRatio !== undefined && aspectRatio < policy.minAspectRatio) return "MIN_ASPECT_RATIO";
  if (policy.maxAspectRatio !== undefined && aspectRatio > policy.maxAspectRatio) return "MAX_ASPECT_RATIO";
  return "INCLUDED";
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  const digest = await globalThis.crypto.subtle.digest("SHA-256", copy.buffer);
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, "0")).join("");
}

function stem(value: string): string {
  return value
    .replace(/\.[^.]+$/, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "asset";
}

function extension(value: string): string {
  const match = /\.([a-z0-9]+)$/i.exec(value);
  return match ? `.${match[1].toLowerCase()}` : "";
}

function exportNameFor(
  sourceName: string,
  originalName: string,
  accessibleDescription: string | undefined,
  index: number,
  template: string | undefined,
  usedNames: Set<string>,
): string {
  const base = (template ?? "{name}")
    .replaceAll("{source}", stem(sourceName))
    .replaceAll("{index}", String(index))
    .replaceAll("{name}", stem(accessibleDescription ?? originalName));
  const safeBase = stem(base);
  const fileExtension = extension(originalName);
  let candidate = `${safeBase}${fileExtension}`;
  let suffix = 2;
  while (usedNames.has(candidate) || candidate === "manifest.json") {
    candidate = `${safeBase}-${suffix}${fileExtension}`;
    suffix += 1;
  }
  usedNames.add(candidate);
  return candidate;
}

function attribute(tag: string, name: string): string | undefined {
  const match = new RegExp(`\\b${name}\\s*=\\s*(["'])(.*?)\\1`, "i").exec(tag);
  return match?.[2]?.trim() || undefined;
}

function resolveRelationshipTarget(partPath: string, target: string): string {
  const directory = partPath.split("/").slice(0, -1);
  const segments = [...directory, ...target.split("/")];
  const resolved: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === ".") continue;
    if (segment === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }
  return resolved.join("/");
}

function relationshipPartPath(partPath: string): string {
  const segments = partPath.split("/");
  const filename = segments.pop();
  return `${segments.join("/")}/_rels/${filename}.rels`;
}

async function relationshipTargets(
  archive: JSZip,
  partPath: string,
): Promise<Map<string, string>> {
  const relationships = archive.file(relationshipPartPath(partPath));
  if (!relationships) return new Map();

  const targets = new Map<string, string>();
  const relationshipText = await relationships.async("string");
  for (const tag of relationshipText.match(/<Relationship\b[^>]*>/gi) ?? []) {
    const id = attribute(tag, "Id");
    const target = attribute(tag, "Target");
    if (id && target) targets.set(id, resolveRelationshipTarget(partPath, target));
  }
  return targets;
}

async function accessibleDescriptions(archive: JSZip): Promise<Map<string, string>> {
  const document = archive.file("word/document.xml");
  const relationships = archive.file("word/_rels/document.xml.rels");
  if (!document || !relationships) return new Map();

  const relationshipText = await relationships.async("string");
  const targetById = new Map<string, string>();
  for (const tag of relationshipText.match(/<Relationship\b[^>]*>/gi) ?? []) {
    const id = attribute(tag, "Id");
    const target = attribute(tag, "Target");
    if (id && target) targetById.set(id, resolveRelationshipTarget("word/document.xml", target));
  }

  const descriptions = new Map<string, string>();
  const documentText = await document.async("string");
  for (const drawing of documentText.match(/<w:drawing\b[\s\S]*?<\/w:drawing>/gi) ?? []) {
    const docProperties = drawing.match(/<wp:docPr\b[^>]*>/i)?.[0];
    const embedId = drawing.match(/<a:blip\b[^>]*\br:embed\s*=\s*(["'])(.*?)\1/i)?.[2];
    const description = docProperties ? attribute(docProperties, "descr") : undefined;
    const target = embedId ? targetById.get(embedId) : undefined;
    if (description && target && !descriptions.has(target)) descriptions.set(target, description);
  }
  return descriptions;
}

type AssetProvenance = {
  descriptions: Map<string, string>;
  references: Map<string, OoxmlReference[]>;
};

async function pptxProvenance(archive: JSZip): Promise<AssetProvenance> {
  const descriptions = new Map<string, string>();
  const references = new Map<string, OoxmlReference[]>();
  const slidePaths = Object.keys(archive.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path))
    .sort();

  for (const partPath of slidePaths) {
    const slide = archive.files[partPath];
    if (!slide) continue;
    const targetById = await relationshipTargets(archive, partPath);

    const slideNumber = Number(/slide(\d+)\.xml$/i.exec(partPath)?.[1]);
    const slideText = await slide.async("string");
    for (const picture of slideText.match(/<p:pic\b[\s\S]*?<\/p:pic>/gi) ?? []) {
      const properties = picture.match(/<p:cNvPr\b[^>]*>/i)?.[0];
      const embedId = picture.match(/<a:blip\b[^>]*\br:embed\s*=\s*(["'])(.*?)\1/i)?.[2];
      const target = embedId ? targetById.get(embedId) : undefined;
      if (!target || !embedId) continue;

      const reference = { partPath, relationshipId: embedId, slideNumber };
      const existing = references.get(target) ?? [];
      existing.push(reference);
      references.set(target, existing);

      const description = properties ? attribute(properties, "descr") : undefined;
      if (description && !descriptions.has(target)) descriptions.set(target, description);
    }
  }

  return { descriptions, references };
}

async function xlsxProvenance(archive: JSZip): Promise<AssetProvenance> {
  const descriptions = new Map<string, string>();
  const references = new Map<string, OoxmlReference[]>();
  const workbookPath = "xl/workbook.xml";
  const workbook = archive.file(workbookPath);
  if (!workbook) return { descriptions, references };

  const worksheetPathById = await relationshipTargets(archive, workbookPath);
  const workbookText = await workbook.async("string");
  for (const sheetTag of workbookText.match(/<sheet\b[^>]*>/gi) ?? []) {
    const relationshipId = attribute(sheetTag, "r:id");
    const worksheetName = attribute(sheetTag, "name");
    const worksheetPartPath = relationshipId ? worksheetPathById.get(relationshipId) : undefined;
    if (!worksheetPartPath || !worksheetName) continue;

    const worksheet = archive.file(worksheetPartPath);
    if (!worksheet) continue;
    const worksheetText = await worksheet.async("string");
    const drawingId = worksheetText.match(/<drawing\b[^>]*\br:id\s*=\s*(["'])(.*?)\1/i)?.[2];
    if (!drawingId) continue;

    const drawingPartPath = (await relationshipTargets(archive, worksheetPartPath)).get(drawingId);
    const drawing = drawingPartPath ? archive.file(drawingPartPath) : undefined;
    if (!drawing || !drawingPartPath) continue;
    const mediaPathById = await relationshipTargets(archive, drawingPartPath);
    const drawingText = await drawing.async("string");
    for (
      const anchor of drawingText.match(
        /<xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)>/gi,
      ) ?? []
    ) {
      const properties = anchor.match(/<xdr:cNvPr\b[^>]*>/i)?.[0];
      const embedId = anchor.match(/<a:blip\b[^>]*\br:embed\s*=\s*(["'])(.*?)\1/i)?.[2];
      const mediaPath = embedId ? mediaPathById.get(embedId) : undefined;
      if (!embedId || !mediaPath) continue;

      const existing = references.get(mediaPath) ?? [];
      existing.push({
        partPath: drawingPartPath,
        relationshipId: embedId,
        workbookPartPath: workbookPath,
        worksheetName,
        worksheetPartPath,
      });
      references.set(mediaPath, existing);

      const description = properties ? attribute(properties, "descr") : undefined;
      if (description && !descriptions.has(mediaPath)) descriptions.set(mediaPath, description);
    }
  }
  return { descriptions, references };
}

async function assetProvenance(archive: JSZip, format: DocumentFormat): Promise<AssetProvenance> {
  if (format === "pptx") return pptxProvenance(archive);
  if (format === "xlsx") return xlsxProvenance(archive);
  return { descriptions: await accessibleDescriptions(archive), references: new Map() };
}

export async function extractDocumentMedia(
  input: ExtractDocumentMediaInput,
): Promise<ExtractionResult> {
  const format = documentFormat(input.sourceName);
  const policy = input.policy ?? {};
  const limits = policy.limits ?? {};
  assertNotCancelled(input.signal);
  assertLimit(input.bytes.byteLength, limits.maxInputBytes, "LIMIT_INPUT_BYTES");

  let archive: JSZip;
  try {
    archive = await JSZip.loadAsync(input.bytes);
  } catch {
    throw new Error("INVALID_ARCHIVE");
  }
  assertNotCancelled(input.signal);

  assertLimit(Object.keys(archive.files).length, limits.maxArchiveEntries, "LIMIT_ARCHIVE_ENTRIES");

  if (!archive.file("[Content_Types].xml") || !archive.file(requiredDocumentPart(format))) {
    throw new Error(invalidDocumentError(format));
  }

  const mediaPaths = Object.keys(archive.files)
    .filter((path) => path.startsWith(mediaPrefix(format)) && !archive.files[path].dir)
    .sort();

  if (mediaPaths.length === 0) throw new Error("NO_MEDIA");
  assertLimit(mediaPaths.length, limits.maxMediaCount, "LIMIT_MEDIA_COUNT");
  const provenance = await assetProvenance(archive, format);
  assertNotCancelled(input.signal);

  const candidates = [];
  let expandedBytes = 0;
  for (const sourcePath of mediaPaths) {
    assertNotCancelled(input.signal);
    const originalName = sourcePath.split("/").at(-1) ?? sourcePath;
    const mediaType = mediaTypeFromName(originalName);
    const bytes = await archive.files[sourcePath].async("uint8array");
    expandedBytes += bytes.byteLength;
    assertNotCancelled(input.signal);
    assertLimit(bytes.byteLength, limits.maxMediaBytes, "LIMIT_MEDIA_BYTES");
    assertLimit(expandedBytes, limits.maxExpandedBytes, "LIMIT_EXPANDED_BYTES");
    assertLimit(expandedBytes, limits.maxTotalMediaBytes, "LIMIT_TOTAL_MEDIA_BYTES");
    const dimensions = dimensionsFor(mediaType, bytes);
    candidates.push({
      sourcePath,
      originalName,
      accessibleDescription: provenance.descriptions.get(sourcePath),
      references: provenance.references.get(sourcePath),
      mediaType,
      bytes,
      dimensions,
      sha256: await sha256(bytes),
    });
  }

  const usedNames = new Set<string>();
  const seenHashes = new Set<string>();
  const manifestAssets: ManifestAsset[] = [];
  const assets: MediaAsset[] = [];

  for (const candidate of candidates) {
    assertNotCancelled(input.signal);
    let reason = decisionFor(candidate.mediaType, candidate.bytes, candidate.dimensions, policy);
    if (
      reason === "INCLUDED" &&
      policy.exactDuplicates === "exclude" &&
      seenHashes.has(candidate.sha256)
    ) {
      reason = "EXACT_DUPLICATE";
    }

    const included = reason === "INCLUDED";
    const measurements = candidate.dimensions
      ? {
          width: candidate.dimensions.width,
          height: candidate.dimensions.height,
          pixels: candidate.dimensions.width * candidate.dimensions.height,
          aspectRatio: candidate.dimensions.width / candidate.dimensions.height,
        }
      : {};
    const manifestAsset: ManifestAsset = {
      sourcePath: candidate.sourcePath,
      originalName: candidate.originalName,
      ...(candidate.accessibleDescription
        ? { accessibleDescription: candidate.accessibleDescription }
        : {}),
      ...(candidate.references?.length ? { references: candidate.references } : {}),
      mediaType: candidate.mediaType,
      byteSize: candidate.bytes.byteLength,
      sha256: candidate.sha256,
      included,
      reason,
      ...measurements,
    };

    if (included) {
      const exportName = exportNameFor(
        input.sourceName,
        candidate.originalName,
        candidate.accessibleDescription,
        assets.length + 1,
        policy.namingTemplate,
        usedNames,
      );
      manifestAsset.exportName = exportName;
      assets.push({ ...candidate, exportName });
      seenHashes.add(candidate.sha256);
    }
    manifestAssets.push(manifestAsset);
  }

  return {
    manifest: {
      sourceName: input.sourceName,
      format,
      policy,
      assets: manifestAssets,
    },
    assets,
  };
}

export async function extractDocumentMediaBatch(
  inputs: readonly ExtractDocumentMediaInput[],
): Promise<BatchExtractionResult[]> {
  const results: BatchExtractionResult[] = [];
  for (const input of inputs) {
    try {
      results.push({ input, status: "success", result: await extractDocumentMedia(input) });
    } catch (error) {
      results.push({
        input,
        status: "failure",
        error: { code: error instanceof Error ? error.message : "UNKNOWN_ERROR" },
      });
    }
  }
  return results;
}
