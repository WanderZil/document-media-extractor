const folderInput = document.querySelector("#folder-input");
const status = document.querySelector("#status");
const summary = document.querySelector("#summary");
const groups = document.querySelector("#groups");
const grid = document.querySelector("#grid");
let reviewState;

folderInput.addEventListener("change", async () => {
  const files = Array.from(folderInput.files || []);
  const manifestFile = files.find((file) => file.webkitRelativePath.endsWith("/manifest.json") || file.name === "manifest.json");
  if (!manifestFile) return showError("No manifest.json found in the selected folder.");

  try {
    const manifest = JSON.parse(await manifestFile.text());
    const fileByName = new Map(files.map((file) => [file.name, file]));
    const assets = manifest.assets.filter((asset) => asset.included);
    status.textContent = `Loaded ${manifest.sourceName}. Calculating local visual similarity…`;
    const previewAssets = await Promise.all(assets.map((asset) => hydrateAsset(asset, fileByName.get(asset.exportName))));
    reviewState = { manifest, assets: previewAssets.map((asset) => ({ ...asset, selected: true })) };
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
  summary.hidden = false;
  groups.hidden = false;
  grid.hidden = false;
  summary.innerHTML = `<div><strong>${assets.length}</strong><span>retained assets</span></div><div><strong>${manifest.assets.length}</strong><span>discovered in Manifest</span></div><div><strong>${visualGroups.length}</strong><span>possible visual groups</span></div><div class="actions"><button id="select-all" class="secondary-button" type="button">Select all</button><button id="export-zip" class="upload-button" type="button">Download selected ZIP</button></div>`;
  groups.innerHTML = visualGroups.length
    ? `<h2>Possible visual matches</h2><p>These are Canvas dHash candidates (Hamming distance ≤ 8). They are review hints only; no file is removed.</p><ul class="group-list">${visualGroups.map((group) => `<li>${group.map((asset) => escapeHtml(asset.exportName)).join(" · ")}</li>`).join("")}</ul>`
    : `<h2>No visual match candidates</h2><p>Only browser-decodable images are compared. The core extractor remains the source of truth for exact-byte duplicates.</p>`;
  grid.innerHTML = assets.map((asset) => `<article class="asset"><div class="preview">${asset.url ? `<img alt="" src="${asset.url}">` : `<span class="missing">Preview unavailable</span>`}</div><div class="asset-copy"><label><input class="asset-selection" type="checkbox" data-name="${escapeHtml(asset.exportName)}" ${asset.selected ? "checked" : ""}> Include in ZIP</label><h3 title="${escapeHtml(asset.exportName)}">${escapeHtml(asset.exportName)}</h3><p class="meta">${escapeHtml(asset.mediaType)}<br>${asset.byteSize.toLocaleString()} bytes</p><span class="badge">${escapeHtml(asset.reason)}</span></div></article>`).join("");
  document.querySelector("#select-all").addEventListener("click", () => {
    reviewState.assets.forEach((asset) => { asset.selected = true; });
    render();
  });
  document.querySelector("#export-zip").addEventListener("click", downloadSelection);
  document.querySelectorAll(".asset-selection").forEach((checkbox) => checkbox.addEventListener("change", (event) => {
    const asset = reviewState.assets.find((candidate) => candidate.exportName === event.target.dataset.name);
    if (asset) asset.selected = event.target.checked;
  }));
  status.textContent = "Review is local to this browser tab.";
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
  URL.revokeObjectURL(url);
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
