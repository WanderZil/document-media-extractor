import JSZip from "jszip";

export type DocumentFormat = "docx";

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
};

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
  throw new Error("UNSUPPORTED_FORMAT");
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

function normalizeRelationshipTarget(target: string): string {
  return `word/${target.replace(/^\.\//, "").replace(/^\//, "")}`;
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
    if (id && target) targetById.set(id, normalizeRelationshipTarget(target));
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

export async function extractDocumentMedia(
  input: ExtractDocumentMediaInput,
): Promise<ExtractionResult> {
  const format = documentFormat(input.sourceName);
  const policy = input.policy ?? {};

  let archive: JSZip;
  try {
    archive = await JSZip.loadAsync(input.bytes);
  } catch {
    throw new Error("INVALID_ARCHIVE");
  }

  if (!archive.file("[Content_Types].xml") || !archive.file("word/document.xml")) {
    throw new Error("INVALID_DOCX");
  }

  const mediaPaths = Object.keys(archive.files)
    .filter((path) => path.startsWith("word/media/") && !archive.files[path].dir)
    .sort();

  if (mediaPaths.length === 0) throw new Error("NO_MEDIA");

  const descriptions = await accessibleDescriptions(archive);

  const candidates = await Promise.all(
    mediaPaths.map(async (sourcePath) => {
      const originalName = sourcePath.split("/").at(-1) ?? sourcePath;
      const mediaType = mediaTypeFromName(originalName);
      const bytes = await archive.files[sourcePath].async("uint8array");
      const dimensions = dimensionsFor(mediaType, bytes);
      return {
        sourcePath,
        originalName,
        accessibleDescription: descriptions.get(sourcePath),
        mediaType,
        bytes,
        dimensions,
        sha256: await sha256(bytes),
      };
    }),
  );

  const usedNames = new Set<string>();
  const seenHashes = new Set<string>();
  const manifestAssets: ManifestAsset[] = [];
  const assets: MediaAsset[] = [];

  for (const candidate of candidates) {
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
