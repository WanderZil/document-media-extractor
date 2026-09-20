import JSZip from "jszip";

export type DocumentFormat = "docx";

export type MediaAsset = {
  sourcePath: string;
  originalName: string;
  mediaType: string;
  bytes: Uint8Array;
};

export type ExtractionManifest = {
  sourceName: string;
  format: DocumentFormat;
  assets: Array<{
    sourcePath: string;
    originalName: string;
    mediaType: string;
    byteSize: number;
  }>;
};

export type ExtractionResult = {
  manifest: ExtractionManifest;
  assets: MediaAsset[];
};

export type ExtractDocumentMediaInput = {
  sourceName: string;
  bytes: Uint8Array;
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

export async function extractDocumentMedia(
  input: ExtractDocumentMediaInput,
): Promise<ExtractionResult> {
  const format = documentFormat(input.sourceName);

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

  const assets = await Promise.all(
    mediaPaths.map(async (sourcePath) => {
      const originalName = sourcePath.split("/").at(-1) ?? sourcePath;
      return {
        sourcePath,
        originalName,
        mediaType: mediaTypeFromName(originalName),
        bytes: await archive.files[sourcePath].async("uint8array"),
      };
    }),
  );

  return {
    manifest: {
      sourceName: input.sourceName,
      format,
      assets: assets.map(({ sourcePath, originalName, mediaType, bytes }) => ({
        sourcePath,
        originalName,
        mediaType,
        byteSize: bytes.byteLength,
      })),
    },
    assets,
  };
}
