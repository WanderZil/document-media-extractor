const folderInput = document.querySelector("#folder-input");
const status = document.querySelector("#status");
const summary = document.querySelector("#summary");
const groups = document.querySelector("#groups");
const grid = document.querySelector("#grid");
let reviewState;
const MAX_VISUAL_CANDIDATES = 250;

folderInput.addEventListener("change", async () => {
  const files = Array.from(folderInput.files || []);
  const manifestFile = files.find((file) => file.webkitRelativePath.endsWith("/manifest.json") || file.name === "manifest.json");
  if (!manifestFile) return showError("No manifest.json found in the selected folder.");

  try {
    const manifest = JSON.parse(await manifestFile.text());
    if (!manifest || !Array.isArray(manifest.assets) || typeof manifest.sourceName !== "string") {
      throw new Error("manifest.json does not match the extractor Manifest format.");
    }
    releaseObjectUrls();
    const fileByName = new Map(files.map((file) => [file.name, file]));
    const assets = manifest.assets.filter(isManifestAsset);
    const visualAssets = assets.filter((asset) => asset.included).slice(0, MAX_VISUAL_CANDIDATES);
    status.textContent = `Loaded ${manifest.sourceName}. Calculating local visual similarity for up to ${MAX_VISUAL_CANDIDATES} retained assets…`;
    const previewByPath = new Map(await Promise.all(visualAssets.map(async (asset) => {
      const hydrated = await hydrateAsset(asset, fileByName.get(asset.exportName));
      return [asset.sourcePath, hydrated];
    })));
    reviewState = {
      manifest,
      assets: assets.map((asset) => ({
        ...asset,
        ...(previewByPath.get(asset.sourcePath) || { file: undefined, hash: null, url: null }),
        selected: Boolean(asset.included && previewByPath.get(asset.sourcePath)?.file),
      })),
    };
    render();
  } catch (error) {
    showError(error instanceof Error ? error.message : "Could not read manifest.json.");
  }
});

async function hydrateAsset(asset, file) {
  if (!file || !file.type.startsWith("image/")) return { ...asset, file, hash: null, url: null };
  const url = URL.createObjectURL(file);
  try {
    return { ...asset, file, url, hash: await perceptualHash(url) };
  } catch {
    return { ...asset, file, url, hash: null };
  }
}

function isManifestAsset(asset) {
  return asset && typeof asset === "object" && typeof asset.sourcePath === "string" && typeof asset.originalName === "string" && typeof asset.mediaType === "string" && typeof asset.byteSize === "number" && typeof asset.included === "boolean" && typeof asset.reason === "string";
}

function releaseObjectUrls() {
  for (const asset of reviewState?.assets || []) {
    if (asset.url) URL.revokeObjectURL(asset.url);
  }
}

async function perceptualHash(url) {
  const image = new Image();
  image.src = url;
  await image.decode();
  const canvas = document.createElement("canvas");
  canvas.width = 9;
  canvas.height = 8;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  context.drawImage(image, 0, 0, 9, 8);
  const pixels = context.getImageData(0, 0, 9, 8).data;
  let hash = "";
  for (let row = 0; row < 8; row += 1) {
    for (let column = 0; column < 8; column += 1) {
      const offset = (row * 9 + column) * 4;
      const next = offset + 4;
      const brightness = pixels[offset] * 0.299 + pixels[offset + 1] * 0.587 + pixels[offset + 2] * 0.114;
      const nextBrightness = pixels[next] * 0.299 + pixels[next + 1] * 0.587 + pixels[next + 2] * 0.114;
      hash += brightness > nextBrightness ? "1" : "0";
    }
  }
  return hash;
}

function hamming(left, right) {
  return [...left].reduce((total, bit, index) => total + (bit === right[index] ? 0 : 1), 0);
}

function similarGroups(assets) {
  const candidates = assets.filter((asset) => asset.hash);
  const parent = new Map(candidates.map((asset) => [asset.exportName, asset.exportName]));
  const find = (name) => parent.get(name) === name ? name : find(parent.get(name));
  const merge = (left, right) => parent.set(find(left), find(right));
  for (let left = 0; left < candidates.length; left += 1) {
    for (let right = left + 1; right < candidates.length; right += 1) {
      if (hamming(candidates[left].hash, candidates[right].hash) <= 8) merge(candidates[left].exportName, candidates[right].exportName);
    }
  }
  const grouped = new Map();
  for (const asset of candidates) {
    const key = find(asset.exportName);
    grouped.set(key, [...(grouped.get(key) || []), asset]);
  }
  return [...grouped.values()].filter((group) => group.length > 1);
}

function render() {
  const { manifest, assets } = reviewState;
  const visualGroups = similarGroups(assets);
  const retained = assets.filter((asset) => asset.included).length;
  const visuallyCompared = assets.filter((asset) => asset.hash).length;
  summary.hidden = false;
  groups.hidden = false;
  grid.hidden = false;
  summary.innerHTML = `<div><strong>${retained}</strong><span>retained assets</span></div><div><strong>${manifest.assets.length}</strong><span>discovered in Manifest</span></div><div><strong>${visualGroups.length}</strong><span>possible visual groups</span></div><div class="actions"><button id="select-all" class="secondary-button" type="button">Select all retained</button><button id="export-zip" class="upload-button" type="button">Download selected ZIP</button></div>`;
  groups.innerHTML = visualGroups.length
    ? `<h2>Possible visual matches</h2><p>These are Canvas dHash candidates (Hamming distance ≤ 8). They are review hints only; no file is removed.</p><ul class="group-list">${visualGroups.map((group) => `<li>${group.map((asset) => escapeHtml(asset.exportName)).join(" · ")}</li>`).join("")}</ul>`
    : `<h2>No visual match candidates</h2><p>Only browser-decodable retained images are compared. The core extractor remains the source of truth for exact-byte duplicates.</p>`;
  if (assets.filter((asset) => asset.included).length > MAX_VISUAL_CANDIDATES) groups.insertAdjacentHTML("beforeend", `<p>Visual comparison is capped at ${MAX_VISUAL_CANDIDATES} retained assets to keep local review responsive.</p>`);
  grid.innerHTML = assets.map((asset) => `<article class="asset"><div class="preview">${asset.url ? `<img alt="" src="${asset.url}">` : `<span class="missing">${asset.included ? "Preview unavailable" : "Excluded by extraction policy"}</span>`}</div><div class="asset-copy"><label><input class="asset-selection" type="checkbox" data-path="${escapeHtml(asset.sourcePath)}" ${asset.selected ? "checked" : ""} ${asset.file ? "" : "disabled"}> Include in ZIP</label><h3 title="${escapeHtml(asset.exportName || asset.originalName)}">${escapeHtml(asset.exportName || asset.originalName)}</h3><p class="meta">${escapeHtml(asset.mediaType)}<br>${asset.byteSize.toLocaleString()} bytes</p><span class="badge">${escapeHtml(asset.reason)}</span></div></article>`).join("");
  document.querySelector("#select-all").addEventListener("click", () => {
    reviewState.assets.forEach((asset) => { asset.selected = Boolean(asset.included && asset.file); });
    render();
  });
  document.querySelector("#export-zip").addEventListener("click", downloadSelection);
  document.querySelectorAll(".asset-selection").forEach((checkbox) => checkbox.addEventListener("change", (event) => {
    const asset = reviewState.assets.find((candidate) => candidate.sourcePath === event.target.dataset.path);
    if (asset) asset.selected = event.target.checked;
  }));
  status.textContent = `Review is local to this browser tab. ${visuallyCompared} retained image(s) were browser-decodable.`;
}

async function downloadSelection() {
  const selected = reviewState.assets.filter((asset) => asset.selected && asset.file);
  if (selected.length === 0) return showError("Select at least one available file to create a ZIP.");
  status.textContent = `Creating a local ZIP with ${selected.length} file(s)…`;
  const zip = new JSZip();
  for (const asset of selected) zip.file(asset.exportName, asset.file);
  const exportManifest = {
    ...reviewState.manifest,
    review: {
      selectedExportNames: selected.map((asset) => asset.exportName),
      visualSimilarity: { algorithm: "dHash", threshold: 8, advisoryOnly: true },
    },
  };
  zip.file("manifest.json", `${JSON.stringify(exportManifest, null, 2)}\n`);
  const blob = await zip.generateAsync({ type: "blob" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${reviewState.manifest.sourceName.replace(/\.[^.]+$/, "")}-review.zip`;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
  status.textContent = "ZIP created locally. No files were uploaded.";
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[character]);
}

function showError(message) {
  status.textContent = message;
  summary.hidden = true;
  groups.hidden = true;
  grid.hidden = true;
}
import JSZip from "jszip";
