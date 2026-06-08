/* ═══════════════════════════════════════════════════════════════
   Avalie Team Icon Vault — PowerPoint Task Pane Logic
   ═══════════════════════════════════════════════════════════════ */

import "./taskpane.css";

/* global Office */

/* ──────────────────────── Types ──────────────────────── */

interface IconifyIcon {
  body: string;
  width?: number;
  height?: number;
}

interface IconifyJSON {
  prefix: string;
  width?: number;
  height?: number;
  icons: Record<string, IconifyIcon>;
}

interface IconEntry {
  set: string;
  key: string;
  body: string;
  width: number;
  height: number;
}

interface SearchResult {
  set: string;
  key: string;
  svg: string;
}

/* ──────────────────── Pack Definitions ──────────────────── */

const ICON_PACKS = [
  { id: "fluent", label: "Fluent UI", file: "icons/fluent.json" },
  { id: "lucide", label: "Lucide", file: "icons/lucide.json" },
  { id: "tabler", label: "Tabler", file: "icons/tabler.json" },
  { id: "heroicons", label: "Heroicons", file: "icons/heroicons.json" },
  { id: "ri", label: "Remix", file: "icons/ri.json" },
  { id: "fa6-regular", label: "FA6 Regular", file: "icons/fa6-regular.json" },
  { id: "fa6-solid", label: "FA6 Solid", file: "icons/fa6-solid.json" },
];

/* ──────────────────── State ──────────────────── */

const iconMap = new Map<string, IconEntry[]>();
const loadedPacks = new Set<string>();
let debounceTimer: number | null = null;
let toastTimer: number | null = null;
let totalIconCount = 0;

/* ──────────────────── Office Initialization ──────────────────── */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    initApp();
  } else {
    showState(
      "⚠️",
      "Unsupported Host",
      "This add-in only works in PowerPoint."
    );
  }
});

/* ──────────────────── App Initialization ──────────────────── */

function initApp(): void {
  // Build filter checkboxes
  buildFilters();

  // Bind search input
  const searchInput = document.getElementById(
    "search-input"
  ) as HTMLInputElement;
  searchInput.addEventListener("input", () => {
    if (debounceTimer) clearTimeout(debounceTimer);
    debounceTimer = window.setTimeout(performSearch, 200);
  });

  // Focus search
  searchInput.focus();

  // Load all icon packs in the background
  loadAllPacks();
}

/* ──────────────────── Filter Chips ──────────────────── */

function buildFilters(): void {
  const container = document.getElementById("filters")!;
  container.innerHTML = "";

  for (const pack of ICON_PACKS) {
    const label = document.createElement("label");
    label.className = "pack-filter";
    label.dataset.pack = pack.id;

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = pack.id;
    checkbox.checked = true;
    checkbox.addEventListener("change", performSearch);

    const chip = document.createElement("span");
    chip.className = "pack-chip";
    chip.innerHTML = `${pack.label} <span class="pack-count" id="count-${pack.id}"></span>`;

    label.appendChild(checkbox);
    label.appendChild(chip);
    container.appendChild(label);
  }
}

/* ──────────────────── Icon Loading ──────────────────── */

async function loadAllPacks(): Promise<void> {
  showState("", "Loading icon libraries…", "", true);

  let loaded = 0;
  const total = ICON_PACKS.length;

  const promises = ICON_PACKS.map(async (pack) => {
    try {
      const response = await fetch(pack.file);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const data: IconifyJSON = await response.json();

      const entries: IconEntry[] = [];
      const defaultWidth = data.width ?? 24;
      const defaultHeight = data.height ?? 24;

      for (const [key, icon] of Object.entries(data.icons)) {
        entries.push({
          set: pack.id,
          key,
          body: icon.body,
          width: icon.width ?? defaultWidth,
          height: icon.height ?? defaultHeight,
        });
      }

      iconMap.set(pack.id, entries);
      loadedPacks.add(pack.id);
      totalIconCount += entries.length;

      // Update count badge
      const countEl = document.getElementById(`count-${pack.id}`);
      if (countEl) countEl.textContent = `(${entries.length.toLocaleString()})`;
    } catch (err) {
      console.error(`[Icon Vault] Failed to load ${pack.id}:`, err);
      iconMap.set(pack.id, []);
    }

    loaded++;
    updateLoadingProgress(loaded, total);
  });

  await Promise.all(promises);

  // Show welcome
  showState(
    "✦",
    "Avalie Team Icon Vault",
    `${totalIconCount.toLocaleString()} icons loaded. Search by keyword, click to insert.`,
    false,
    "Shift+click to copy SVG to clipboard instead."
  );
}

function updateLoadingProgress(loaded: number, total: number): void {
  const bar = document.querySelector(".loading-bar-fill") as HTMLElement | null;
  if (bar) {
    bar.style.width = `${(loaded / total) * 100}%`;
  }
}

/* ──────────────────── Search ──────────────────── */

function getActivePacks(): string[] {
  return Array.from(
    document.querySelectorAll<HTMLInputElement>(".pack-filter input:checked")
  ).map((cb) => cb.value);
}

function performSearch(): void {
  const searchInput = document.getElementById(
    "search-input"
  ) as HTMLInputElement;
  const query = searchInput.value.trim();
  const activePacks = getActivePacks();

  if (!query) {
    showState(
      "✦",
      "Avalie Team Icon Vault",
      `${totalIconCount.toLocaleString()} icons loaded. Search by keyword, click to insert.`,
      false,
      "Shift+click to copy SVG to clipboard instead."
    );
    const statusBar = document.getElementById("status-bar")!;
    statusBar.style.display = "none";
    return;
  }

  const results = search(query, activePacks);
  renderResults(results, query);
}

function search(query: string, packs: string[]): SearchResult[] {
  const tokens = query
    .toLowerCase()
    .split(/[\s\-_]+/)
    .filter(Boolean);
  const results: SearchResult[] = [];
  const maxResults = 200;

  for (const packId of packs) {
    const entries = iconMap.get(packId);
    if (!entries) continue;

    for (const entry of entries) {
      if (results.length >= maxResults) break;

      const normalizedKey = entry.key.toLowerCase();
      if (tokens.every((t) => normalizedKey.includes(t))) {
        const svg = buildSvgString(entry);
        results.push({ set: entry.set, key: entry.key, svg });
      }
    }

    if (results.length >= maxResults) break;
  }

  return results;
}

function buildSvgString(entry: IconEntry): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${entry.width} ${entry.height}">${entry.body}</svg>`;
}

/* ──────────────────── Render Results ──────────────────── */

function renderResults(results: SearchResult[], query: string): void {
  const content = document.getElementById("content")!;
  const statusBar = document.getElementById("status-bar")!;
  const resultCountEl = document.getElementById("result-count")!;
  const resultCapEl = document.getElementById("result-cap")!;

  if (!results.length) {
    statusBar.style.display = "none";
    showState(
      "🔍",
      "No icons found",
      `No results for "${escapeHtml(query)}". Try a different keyword.`
    );
    return;
  }

  // Update status bar
  statusBar.style.display = "flex";
  resultCountEl.textContent = results.length.toString();
  resultCapEl.textContent =
    results.length >= 200 ? "(showing first 200)" : "";

  // Build grid
  const grid = document.createElement("div");
  grid.className = "icon-grid";

  for (const item of results) {
    const card = document.createElement("div");
    card.className = "icon-card";
    card.title = `${item.set} / ${item.key}\nClick to insert • Shift+click to copy`;
    card.dataset.svg = item.svg;
    card.dataset.set = item.set;
    card.dataset.key = item.key;

    // Badge
    const badge = document.createElement("span");
    badge.className = "icon-set-badge";
    badge.textContent = item.set;

    // SVG preview
    const preview = document.createElement("div");
    preview.className = "icon-preview";
    preview.innerHTML = item.svg;

    // Name
    const name = document.createElement("span");
    name.className = "icon-name";
    name.textContent = item.key;

    card.appendChild(badge);
    card.appendChild(preview);
    card.appendChild(name);

    // Click handler
    card.addEventListener("click", async (e) => {
      const svgStr = card.dataset.svg!;
      const iconKey = card.dataset.key!;
      const iconSet = card.dataset.set!;

      if (e.shiftKey) {
        // Shift+click → copy to clipboard
        try {
          await navigator.clipboard.writeText(svgStr);
          showToast(`Copied ${iconKey}`, "success");
        } catch {
          // Fallback for clipboard API restrictions
          fallbackCopy(svgStr);
          showToast(`Copied ${iconKey}`, "success");
        }
      } else {
        // Normal click → insert on slide
        card.classList.add("inserting");
        try {
          await insertIconOnSlide(svgStr);
          showToast(`Inserted ${iconKey} (${iconSet})`, "success");
        } catch (err) {
          console.error("Insert failed:", err);
          showToast("Insert failed — try Shift+click to copy", "error");
        } finally {
          card.classList.remove("inserting");
        }
      }
    });

    grid.appendChild(card);
  }

  content.innerHTML = "";
  content.appendChild(grid);
}

/* ──────────────────── Insert Icon onto Slide ──────────────────── */

async function insertIconOnSlide(svgString: string): Promise<void> {
  // Determine aspect ratio to size it reasonably
  let dispWidth = 100;
  let dispHeight = 100;

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(svgString, "image/svg+xml");
    const svgEl = doc.querySelector("svg");
    if (svgEl) {
      const viewBox = svgEl.getAttribute("viewBox");
      if (viewBox) {
        const parts = viewBox.split(/\s+/).map(Number);
        if (parts.length === 4 && !isNaN(parts[2]) && !isNaN(parts[3]) && parts[3] !== 0) {
          const aspect = parts[2] / parts[3];
          if (aspect > 1) {
            dispHeight = 100 / aspect;
          } else {
            dispWidth = 100 * aspect;
          }
        }
      }
    }
  } catch (e) {
    console.error("[Icon Vault] Failed to parse SVG aspect ratio:", e);
  }

  // 1. Try SVG insertion if supported (ImageCoercion 1.2)
  const canUseSvg =
    typeof Office !== "undefined" &&
    Office.context.requirements.isSetSupported("ImageCoercion", "1.2") &&
    !!(Office.CoercionType as any)["XmlSvg"];

  if (canUseSvg) {
    try {
      // Modify SVG to have explicit width/height matching the aspect ratio
      let insertSvg = svgString;
      if (!insertSvg.includes('width="')) {
        insertSvg = insertSvg.replace(
          "<svg",
          `<svg width="${dispWidth}" height="${dispHeight}"`
        );
      }

      await new Promise<void>((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(
          insertSvg,
          { coercionType: Office.CoercionType.XmlSvg },
          (result: Office.AsyncResult<void>) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(new Error(result.error.message));
            }
          }
        );
      });
      return;
    } catch (err) {
      console.warn("[Icon Vault] SVG insertion failed, falling back to PNG:", err);
    }
  }

  // 2. Fallback to PNG insertion
  const base64 = await svgToPngBase64(svgString, 512);
  await new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      base64,
      {
        coercionType: Office.CoercionType.Image,
        imageWidth: dispWidth,
        imageHeight: dispHeight,
      } as any,
      (result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}

function svgToPngBase64(svgString: string, targetSize = 512): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(svgString, "image/svg+xml");
      const svgEl = doc.querySelector("svg");
      if (!svgEl) {
        return reject(new Error("Invalid SVG element"));
      }

      let width = targetSize;
      let height = targetSize;

      const viewBox = svgEl.getAttribute("viewBox");
      if (viewBox) {
        const parts = viewBox.split(/\s+/).map(Number);
        if (parts.length === 4 && !isNaN(parts[2]) && !isNaN(parts[3]) && parts[3] !== 0) {
          const aspect = parts[2] / parts[3];
          if (aspect > 1) {
            height = targetSize / aspect;
          } else {
            width = targetSize * aspect;
          }
        }
      }

      svgEl.setAttribute("width", width.toString());
      svgEl.setAttribute("height", height.toString());

      const serializedSvg = new XMLSerializer().serializeToString(svgEl);
      const svgBlob = new Blob([serializedSvg], { type: "image/svg+xml;charset=utf-8" });
      const url = URL.createObjectURL(svgBlob);

      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.drawImage(img, 0, 0, width, height);
          const dataUrl = canvas.toDataURL("image/png");
          const base64 = dataUrl.split(",")[1];
          URL.revokeObjectURL(url);
          resolve(base64);
        } else {
          URL.revokeObjectURL(url);
          reject(new Error("Canvas context not available"));
        }
      };

      img.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error("Failed to render SVG onto canvas"));
      };

      img.src = url;
    } catch (e) {
      reject(e);
    }
  });
}

/* ──────────────────── Clipboard Fallback ──────────────────── */

function fallbackCopy(text: string): void {
  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.select();
  try {
    document.execCommand("copy");
  } catch {
    // Silently fail
  }
  document.body.removeChild(textarea);
}

/* ──────────────────── UI Helpers ──────────────────── */

function showState(
  icon: string,
  title: string,
  message: string,
  showSpinner = false,
  hint = ""
): void {
  const content = document.getElementById("content")!;

  let html = '<div class="state-message">';
  if (showSpinner) {
    html += '<div class="spinner"></div>';
    html += `<h3>${escapeHtml(title)}</h3>`;
    html += `<p>${escapeHtml(message)}</p>`;
    html += '<div class="loading-bar"><div class="loading-bar-fill"></div></div>';
  } else {
    if (icon) html += `<div class="state-icon">${icon}</div>`;
    html += `<h3>${escapeHtml(title)}</h3>`;
    if (message) html += `<p>${escapeHtml(message)}</p>`;
    if (hint) html += `<p class="hint">${escapeHtml(hint)}</p>`;
  }
  html += "</div>";

  content.innerHTML = html;
}

function showToast(
  message: string,
  type: "success" | "error" | "info" = "info"
): void {
  const toast = document.getElementById("toast")!;
  toast.textContent = type === "success" ? `✓ ${message}` : message;
  toast.className = `toast visible ${type}`;

  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => {
    toast.classList.remove("visible");
  }, 2000);
}

function escapeHtml(str: string): string {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}
