import { ensureDownloadLocationConfigured } from "../../api/config.js";
import { deleteTenantDocument, fetchTenantDocuments, uploadTenantDocument } from "../../api/sheets.js";
import { getLocalData, getLocalList, setLocalData, LOCAL_KEYS } from "../../api/localStore.js";
import { escapeHtml } from "../../utils/htmlUtils.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

const MAX_DOC_BYTES = 10 * 1024 * 1024;
const ALLOWED_EXTENSIONS = [
    "png",
    "jpg",
    "jpeg",
    "webp",
    "gif",
    "pdf",
    "doc",
    "docx",
    "xls",
    "xlsx",
    "csv",
    "txt",
    "rtf",
    "ppt",
    "pptx",
    "odt",
    "ods",
    "odp",
];
const ALLOWED_EXTENSION_SET = new Set(ALLOWED_EXTENSIONS);

const docsState = {
    tenant: null,
    docs: [],
    cache: new Map(),
    cacheLoaded: false,
    modalBound: false,
    uploadInProgress: false,
    downloadInProgress: false,
    downloads: new Map(),
    uploads: new Map(),
    uploadPlaceholders: new Map(),
    renderPending: false,
};

function scheduleDocsRender() {
    if (docsState.renderPending) return;
    docsState.renderPending = true;
    requestAnimationFrame(() => {
        docsState.renderPending = false;
        renderDocsList();
    });
}

function setDocDownloadState(docId, state) {
    if (!docId) return;
    if (!state) {
        docsState.downloads.delete(docId);
    } else {
        docsState.downloads.set(docId, state);
    }
    scheduleDocsRender();
}

function setDocUploadState(docId, state) {
    if (!docId) return;
    if (!state) {
        docsState.uploads.delete(docId);
    } else {
        docsState.uploads.set(docId, state);
    }
    scheduleDocsRender();
}

function normalizeId(value) {
    return value === undefined || value === null ? "" : String(value).trim();
}

function normalizeKey(value) {
    return normalizeId(value).toLowerCase();
}

function getTenantId(tenant) {
    return normalizeId(tenant?.tenantId || tenant?.tenant_id || tenant?.templateData?.tenant_id || "");
}

function getTenantName(tenant) {
    return (
        tenant?.tenantFullName ||
        tenant?.tenantName ||
        tenant?.tenant_name ||
        tenant?.full_name ||
        tenant?.templateData?.Tenant_Full_Name ||
        "Tenant"
    ).toString();
}

function getElements() {
    return {
        modal: document.getElementById("tenantDocsModal"),
        closeBtn: document.getElementById("tenantDocsClose"),
        tenantName: document.getElementById("tenantDocsTenantName"),
        fileInput: document.getElementById("tenantDocsFileInput"),
        typeSelect: document.getElementById("tenantDocsTypeSelect"),
        otherWrap: document.getElementById("tenantDocsOtherWrap"),
        otherInput: document.getElementById("tenantDocsOtherInput"),
        uploadBtn: document.getElementById("tenantDocsUploadBtn"),
        list: document.getElementById("tenantDocsList"),
        empty: document.getElementById("tenantDocsEmpty"),
        count: document.getElementById("tenantDocsCount"),
    };
}

function getDownloadModalElements() {
    return {
        modal: document.getElementById("tenantDocDownloadModal"),
        closeBtn: document.getElementById("tenantDocDownloadClose"),
        fileName: document.getElementById("tenantDocDownloadFileName"),
        location: document.getElementById("tenantDocDownloadLocation"),
        openLink: document.getElementById("tenantDocDownloadOpen"),
    };
}

function sanitizeFileSegment(value, fallback = "document") {
    const raw = normalizeId(value);
    if (!raw) return fallback;
    const cleaned = raw.replace(/[^\w.-]+/g, "_").replace(/^_+|_+$/g, "");
    return cleaned || fallback;
}

function getFileExtension(name = "") {
    const idx = name.lastIndexOf(".");
    if (idx < 0 || idx === name.length - 1) return "";
    return name.slice(idx + 1).toLowerCase();
}

function getFileStem(name = "") {
    const idx = name.lastIndexOf(".");
    if (idx <= 0) return name;
    return name.slice(0, idx);
}

const DOC_TYPE_LABELS = {
    aadhar: "AADHAR",
    pan: "PAN",
    agreement: "AGREEMENT",
    tenant_image: "TENANT",
    other: "OTHER",
};

function normalizeDocType(value) {
    const raw = normalizeKey(value);
    if (!raw) return "other";
    const cleaned = raw.replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
    if (cleaned === "aadhaar") return "aadhar";
    if (cleaned === "aadhar") return "aadhar";
    if (cleaned === "pan") return "pan";
    if (cleaned === "agreement" || cleaned === "rent_agreement") return "agreement";
    if (cleaned === "tenantimage" || cleaned === "tenant_image" || cleaned === "tenant_photo") return "tenant_image";
    if (cleaned === "other") return "other";
    return "other";
}

function resolveDocTypeSelection(value) {
    const raw = normalizeKey(value);
    if (!raw || raw === "--" || raw === "default") return "default";
    return normalizeDocType(raw);
}

function syncDocTypeFields() {
    const { typeSelect, otherWrap, otherInput } = getElements();
    if (!typeSelect || !otherWrap) return;
    const isOther = resolveDocTypeSelection(typeSelect.value) === "other";
    otherWrap.classList.toggle("hidden", !isOther);
    if (!isOther && otherInput) otherInput.value = "";
}

function getIconLabelFontSize(label) {
    const length = label.length;
    if (length <= 3) return 4.5;
    if (length === 4) return 4.0;
    if (length === 5) return 3.6;
    if (length === 6) return 3.3;
    if (length === 7) return 3.1;
    return 2.8;
}

const FILE_ICON_BASE_PATH =
    'M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z';
const FILE_ICON_VARIANTS = {
    image: {
        accent:
            '<path fill="currentColor" class="text-indigo-300" d="M8.5 11.5l-2.5 3h12l-3.5-4.5-2.5 3-1.5-2z M15.5 10a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3z"/>',
        barClass: "fill-indigo-600",
    },
    pdf: {
        accent:
            '<path fill="currentColor" class="text-red-300" d="M8 9c.5-1.5 2-2.5 3.5-1.5s2.5 2 1.5 3.5S10.5 13.5 9 12.5 6.5 10.5 8 9z" opacity="0.7"/>',
        barClass: "fill-red-600",
    },
    doc: {
        accent:
            '<g fill="currentColor" class="text-blue-400"><rect x="7" y="8" width="10" height="1.5" rx="0.5"/><rect x="7" y="11" width="10" height="1.5" rx="0.5"/><rect x="7" y="14" width="7" height="1.5" rx="0.5"/></g>',
        barClass: "fill-blue-600",
    },
    sheet: {
        accent:
            '<g fill="currentColor" class="text-green-400" opacity="0.8"><rect x="7" y="7" width="3.5" height="3.5" rx="0.5"/><rect x="11.5" y="7" width="3.5" height="3.5" rx="0.5"/><rect x="7" y="11.5" width="3.5" height="3.5" rx="0.5"/><rect x="11.5" y="11.5" width="3.5" height="3.5" rx="0.5"/></g>',
        barClass: "fill-green-600",
    },
    slide: {
        accent:
            '<path fill="currentColor" class="text-orange-300" d="M 8 14 L 8 10 L 10 10 L 10 14 Z M 11 14 L 11 7 L 13 7 L 13 14 Z M 14 14 L 14 11 L 16 11 L 16 14 Z"/>',
        barClass: "fill-orange-500",
    },
    aadhar: {
        accent:
            '<rect x="7" y="7.2" width="10" height="2.1" rx="0.6" fill="currentColor" class="text-orange-400"/><rect x="7" y="9.5" width="10" height="2.1" rx="0.6" fill="currentColor" class="text-slate-100"/><rect x="7" y="11.8" width="10" height="2.1" rx="0.6" fill="currentColor" class="text-emerald-500"/><circle cx="12" cy="10.6" r="0.8" fill="currentColor" class="text-blue-600"/>',
        barClass: "fill-emerald-600",
    },
    pan: {
        accent:
            '<rect x="7" y="7.5" width="10" height="6.5" rx="1" fill="currentColor" class="text-blue-100"/><rect x="8" y="9" width="5.8" height="1.2" rx="0.5" fill="currentColor" class="text-blue-500"/><rect x="8" y="11" width="7.2" height="1.2" rx="0.5" fill="currentColor" class="text-blue-500"/><circle cx="15.6" cy="10.2" r="1" fill="currentColor" class="text-blue-400"/>',
        barClass: "fill-blue-700",
    },
    agreement: {
        accent:
            '<g fill="currentColor" class="text-amber-300"><rect x="7" y="8" width="10" height="1.5" rx="0.5"/><rect x="7" y="11" width="6.5" height="1.5" rx="0.5"/></g><path d="M11.6 13.3l1.4 1.4 2.6-2.6" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" class="text-emerald-500"/>',
        barClass: "fill-emerald-600",
    },
    tenant_image: {
        accent:
            '<circle cx="12" cy="9.6" r="2.1" fill="currentColor" class="text-indigo-300"/><path fill="currentColor" class="text-indigo-400" d="M7.2 15.2c0-2.2 2.4-3.6 4.8-3.6s4.8 1.4 4.8 3.6v1.2H7.2z"/>',
        barClass: "fill-indigo-600",
    },
    neutral: {
        accent:
            '<g fill="currentColor" class="text-slate-300"><rect x="7" y="8" width="10" height="1.5" rx="0.5"/><rect x="7" y="11" width="7" height="1.5" rx="0.5"/></g>',
        barClass: "fill-slate-600",
    },
};

const FILE_ICON_MAP = {
    png: { label: "PNG", variant: "image" },
    jpg: { label: "JPG", variant: "image" },
    jpeg: { label: "JPG", variant: "image" },
    webp: { label: "WEBP", variant: "image" },
    gif: { label: "GIF", variant: "image" },
    pdf: { label: "PDF", variant: "pdf" },
    doc: { label: "DOC", variant: "doc" },
    docx: { label: "DOCX", variant: "doc" },
    odt: { label: "ODT", variant: "doc" },
    txt: { label: "TXT", variant: "doc" },
    rtf: { label: "RTF", variant: "doc" },
    xls: { label: "XLS", variant: "sheet" },
    xlsx: { label: "XLSX", variant: "sheet" },
    csv: { label: "CSV", variant: "sheet" },
    ods: { label: "ODS", variant: "sheet" },
    ppt: { label: "PPT", variant: "slide" },
    pptx: { label: "PPTX", variant: "slide" },
    odp: { label: "ODP", variant: "slide" },
};

function buildFileIconSvg({ label, variant }) {
    const safeLabel = (label || "FILE").toString().replace(/[^A-Za-z0-9]/g, "").toUpperCase() || "FILE";
    const iconVariant = FILE_ICON_VARIANTS[variant] || FILE_ICON_VARIANTS.neutral;
    const fontSize = getIconLabelFontSize(safeLabel);
    return `
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" class="w-14 h-14 pointer-events-none">
            <path fill="currentColor" class="text-gray-300" d="${FILE_ICON_BASE_PATH}"></path>
            ${iconVariant.accent}
            <rect x="3" y="16" width="18" height="7" rx="1" class="${iconVariant.barClass}"></rect>
            <text x="12" y="20.5" text-anchor="middle" font-size="${fontSize}" font-weight="bold" fill="white" style="pointer-events: none;">${escapeHtml(
                safeLabel
            )}</text>
        </svg>
    `;
}

function getFileIconSvg(fileName) {
    const ext = getFileExtension(fileName);
    const cleaned = ext.replace(/[^A-Za-z0-9]/g, "");
    const lookup = FILE_ICON_MAP[cleaned] || null;
    if (lookup) return buildFileIconSvg(lookup);
    const fallbackLabel = cleaned ? cleaned.toUpperCase() : "FILE";
    return buildFileIconSvg({ label: fallbackLabel, variant: "neutral" });
}

function getDocIconSvg(doc) {
    const fields = resolveDocFields(doc);
    const docType = normalizeDocType(fields.docType);
    if (docType === "other") return getFileIconSvg(fields.fileName);
    const label = DOC_TYPE_LABELS[docType] || DOC_TYPE_LABELS.other;
    return buildFileIconSvg({ label, variant: docType });
}

function isAllowedFile(file) {
    if (!file) return false;
    const ext = getFileExtension(file.name);
    if (file.type && file.type.startsWith("video/")) return false;
    return ALLOWED_EXTENSION_SET.has(ext);
}

function formatBytes(bytes) {
    const size = Number(bytes) || 0;
    if (!size) return "-";
    if (size < 1024) return `${size} B`;
    const kb = size / 1024;
    if (kb < 1024) return `${kb.toFixed(1)} KB`;
    return `${(kb / 1024).toFixed(1)} MB`;
}

function formatDateTime(value) {
    if (!value) return "-";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return date.toLocaleString();
}

function resolveDocFields(doc) {
    return {
        id: normalizeId(doc?.doc_id || doc?.docId || doc?.id),
        tenantId: normalizeId(doc?.tenant_id || doc?.tenantId || ""),
        tenantName: doc?.tenant_name || doc?.tenantName || "",
        docType: doc?.doc_type || doc?.docType || doc?.type || doc?.document_type || "",
        docDetails: normalizeId(doc?.doc_details || doc?.docDetails || doc?.details || ""),
        fileName: doc?.file_name || doc?.fileName || doc?.name || "",
        fileUrl: doc?.file_url || doc?.fileUrl || "",
        fileDriveId: doc?.file_drive_id || doc?.fileDriveId || "",
        fileMime: doc?.file_mime || doc?.fileMime || "",
        fileSize: Number(doc?.file_size ?? doc?.fileSize) || 0,
        uploadedAt: doc?.uploaded_at || doc?.uploadedAt || "",
    };
}

function docMatchesTenant(doc, tenant) {
    const docFields = resolveDocFields(doc);
    const tenantId = getTenantId(tenant);
    if (tenantId && docFields.tenantId) return normalizeKey(tenantId) === normalizeKey(docFields.tenantId);
    const tenantName = normalizeKey(getTenantName(tenant));
    return tenantName && tenantName === normalizeKey(docFields.tenantName);
}

function getRenderableDocs() {
    if (!docsState.uploadPlaceholders.size) return docsState.docs.slice();
    return [...docsState.uploadPlaceholders.values(), ...docsState.docs];
}

async function ensureDocsCacheLoaded() {
    if (docsState.cacheLoaded) return;
    const cached = (await getLocalData(LOCAL_KEYS.docsCache, {})) || {};
    Object.entries(cached).forEach(([key, value]) => {
        if (!key || !value) return;
        docsState.cache.set(key, value);
    });
    docsState.cacheLoaded = true;
}

async function persistDocsCache() {
    const obj = {};
    docsState.cache.forEach((value, key) => {
        obj[key] = value;
    });
    await setLocalData(LOCAL_KEYS.docsCache, obj);
}

async function removeCachedDocEntry(docId) {
    if (!docId) return;
    await ensureDocsCacheLoaded();
    if (!docsState.cache.has(docId)) return;
    docsState.cache.delete(docId);
    await persistDocsCache();
}

async function fileExists(filePath) {
    const fs = window.__TAURI__?.fs;
    if (!fs || !filePath) return null;
    if (typeof fs.exists === "function") {
        try {
            return await fs.exists(filePath);
        } catch (err) {
            return null;
        }
    }
    if (typeof fs.readFile === "function") {
        try {
            await fs.readFile(filePath);
            return true;
        } catch (err) {
            return null;
        }
    }
    return null;
}

async function ensureCacheDirectory(tenant) {
    const pathApi = window.__TAURI__?.path;
    const tauriInvoke = window.__TAURI__?.core?.invoke;
    if (typeof tauriInvoke === "function") {
        try {
            const result = await tauriInvoke("ensure_temp_dir", {
                tenantName: getTenantName(tenant),
            });
            if (typeof result === "string" && result.trim()) return result.trim();
            if (result?.path) return String(result.path);
        } catch (err) {
            // fall back to JS path API
        }
    }
    if (!pathApi) return "";
    const tenantSegment = sanitizeFileSegment(getTenantName(tenant));
    const bases = [];
    if (typeof pathApi.tempDir === "function") bases.push(await pathApi.tempDir());
    if (typeof pathApi.appCacheDir === "function") bases.push(await pathApi.appCacheDir());
    if (typeof pathApi.appDataDir === "function") bases.push(await pathApi.appDataDir());
    if (typeof pathApi.downloadDir === "function") bases.push(await pathApi.downloadDir());

    for (const baseDir of bases) {
        if (!baseDir) continue;
        const docsDir = await pathApi.join(baseDir, "Tenant_Docs");
        const tenantDir = await pathApi.join(docsDir, tenantSegment || "Tenant");
        return tenantDir;
    }
    return "";
}

function buildDriveDownloadUrl(doc) {
    const { fileDriveId, fileUrl } = resolveDocFields(doc);
    if (fileDriveId) {
        return `https://drive.google.com/uc?export=download&id=${fileDriveId}`;
    }
    const raw = normalizeId(fileUrl);
    const match = raw.match(/\/d\/([a-zA-Z0-9_-]+)/) || raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match) {
        return `https://drive.google.com/uc?export=download&id=${match[1]}`;
    }
    return raw;
}

async function downloadFileToPath(url, filePath, { onProgress, downloadId } = {}) {
    const tauriInvoke = window.__TAURI__?.core?.invoke;
    const tauriEvents = window.__TAURI__?.event;

    if (typeof tauriInvoke === "function") {
        const resolvedId = downloadId || `download-${Date.now()}-${Math.random().toString(16).slice(2)}`;
        let unlisten = null;
        if (onProgress && typeof tauriEvents?.listen === "function") {
            unlisten = await tauriEvents.listen("download-progress", (event) => {
                const payloadData = event?.payload || {};
                const id = payloadData.downloadId || payloadData.download_id || "";
                if (!id || id !== resolvedId) return;
                const loaded = Number(payloadData.loaded) || 0;
                const total = Number(payloadData.total) || 0;
                const percent = total ? (loaded / total) * 100 : null;
                onProgress({ loaded, total, percent });
                if (payloadData.done && typeof unlisten === "function") {
                    unlisten();
                    unlisten = null;
                }
            });
        }
        try {
            const result = await tauriInvoke("download_file_to_path", {
                url,
                filePath,
                downloadId: resolvedId,
            });
            if (typeof unlisten === "function") unlisten();
            return result;
        } catch (err) {
            if (typeof unlisten === "function") unlisten();
            throw err;
        }
    }

    const response = await fetch(url);
    if (!response.ok) throw new Error(`Download failed (${response.status})`);
    const total = Number(response.headers.get("Content-Length")) || 0;
    if (!response.body || !response.body.getReader) {
        const blob = await response.blob();
        const buffer = await blob.arrayBuffer();
        const bytes = new Uint8Array(buffer);
        if (typeof onProgress === "function") {
            onProgress({ loaded: bytes.length, total: bytes.length, percent: 100 });
        }
        if (window.__TAURI__?.fs?.writeFile) {
            await window.__TAURI__.fs.writeFile(filePath, bytes);
        }
        return;
    }

    const reader = response.body.getReader();
    let received = 0;
    const chunks = [];
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        if (value) {
            chunks.push(value);
            received += value.length;
            if (typeof onProgress === "function") {
                const percent = total ? (received / total) * 100 : null;
                onProgress({ loaded: received, total, percent });
            }
        }
    }
    const length = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const bytes = new Uint8Array(length);
    let offset = 0;
    chunks.forEach((chunk) => {
        bytes.set(chunk, offset);
        offset += chunk.length;
    });
    if (typeof onProgress === "function") {
        onProgress({ loaded: length, total: total || length, percent: 100 });
    }
    if (window.__TAURI__?.fs?.writeFile) {
        await window.__TAURI__.fs.writeFile(filePath, bytes);
    }
}

async function copyFileToPath(sourcePath, targetPath) {
    const fs = window.__TAURI__?.fs;
    if (!fs) throw new Error("File system is unavailable");
    if (typeof fs.copyFile === "function") {
        await fs.copyFile(sourcePath, targetPath);
        return;
    }
    const bytes = await fs.readFile(sourcePath);
    await fs.writeFile(targetPath, bytes);
}

async function getCachedDocPath(doc) {
    const id = resolveDocFields(doc).id;
    if (!id) return "";
    await ensureDocsCacheLoaded();
    const entry = docsState.cache.get(id);
    if (!entry?.path) return "";
    const exists = await fileExists(entry.path);
    if (exists === true || exists === null) return entry.path;
    docsState.cache.delete(id);
    await persistDocsCache();
    return "";
}

async function downloadDocToCache(doc, options = {}) {
    const tenant = docsState.tenant;
    const cacheDir = await ensureCacheDirectory(tenant);
    if (!cacheDir) throw new Error("Cache directory is unavailable");
    const { id, fileName } = resolveDocFields(doc);
    const pathApi = window.__TAURI__?.path;
    const baseName = sanitizeFileSegment(getFileStem(fileName || "document"));
    const ext = getFileExtension(fileName);
    const suffix = id ? `_${id.slice(0, 6)}` : `_${Date.now()}`;
    const safeName = ext ? `${baseName}${suffix}.${ext}` : `${baseName}${suffix}`;
    const separator = cacheDir.includes("\\") ? "\\" : "/";
    const filePath = pathApi?.join
        ? await pathApi.join(cacheDir, safeName)
        : `${cacheDir}${cacheDir.endsWith(separator) ? "" : separator}${safeName}`;

    const url = buildDriveDownloadUrl(doc);
    const resolvedOptions = { ...options };
    if (!resolvedOptions.downloadId) {
        resolvedOptions.downloadId = `doc-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    }
    const externalProgress = typeof resolvedOptions.onProgress === "function" ? resolvedOptions.onProgress : null;
    resolvedOptions.onProgress = ({ loaded, total, percent }) => {
        const safePercent = percent ?? (total ? (loaded / total) * 100 : 0);
        setDocDownloadState(id, { status: "downloading", percent: safePercent });
        if (externalProgress) externalProgress({ loaded, total, percent: safePercent });
    };
    setDocDownloadState(id, { status: "downloading", percent: 0 });
    try {
        await downloadFileToPath(url, filePath, resolvedOptions);
    } catch (err) {
        setDocDownloadState(id, null);
        throw err;
    }

    docsState.cache.set(id, { path: filePath, fileName: fileName || safeName, cachedAt: Date.now() });
    await persistDocsCache();
    setDocDownloadState(id, null);
    renderDocsList();
    return filePath;
}

async function getOrDownloadCachedDoc(doc) {
    const existing = await getCachedDocPath(doc);
    if (existing) return existing;
    return await downloadDocToCache(doc);
}

async function openLocalFile(filePath) {
    if (!window.__TAURI__?.opener?.openPath) {
        showToast("Open is available only in the desktop app.", "warning");
        return;
    }
    await window.__TAURI__.opener.openPath(filePath);
}

function updateDocCount(count) {
    const { count: countLabel } = getElements();
    if (countLabel) countLabel.textContent = `${count} file${count === 1 ? "" : "s"}`;
}

async function pruneMissingCachedDocs(list) {
    if (!window.__TAURI__) return;
    let changed = false;
    const checks = list
        .map((doc) => {
            const id = resolveDocFields(doc).id;
            if (!id) return null;
            const entry = docsState.cache.get(id);
            if (!entry?.path) return null;
            return fileExists(entry.path).then((exists) => {
                if (exists === false) {
                    docsState.cache.delete(id);
                    changed = true;
                }
            });
        })
        .filter(Boolean);

    if (checks.length) await Promise.all(checks);
    if (changed) await persistDocsCache();
}

function renderDocsList() {
    const { list, empty } = getElements();
    if (!list || !empty) return;
    list.innerHTML = "";

    const allDocs = getRenderableDocs();
    if (!allDocs.length) {
        empty.classList.remove("hidden");
        updateDocCount(0);
        return;
    }

    empty.classList.add("hidden");
    updateDocCount(allDocs.length);

    const icons = {
        download:
            '<svg class="h-3.5 w-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v12"></path><path d="M7 10l5 5 5-5"></path><path d="M5 21h14"></path></svg>',
        copy:
            '<svg class="h-3.5 w-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="9" width="13" height="13" rx="2"></rect><rect x="2" y="2" width="13" height="13" rx="2"></rect></svg>',
        delete:
            '<svg class="h-3.5 w-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18"></path><path d="M8 6V4h8v2"></path><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path><path d="M10 11v6"></path><path d="M14 11v6"></path></svg>',
    };
    const actionButtonBase =
        "inline-flex h-6 w-6 items-center justify-center rounded-md border text-[10px] font-semibold";
    const actionButtonNeutral = `${actionButtonBase} border-slate-200 text-slate-700 hover:bg-slate-100`;
    const actionButtonDanger = `${actionButtonBase} border-rose-200 text-rose-600 hover:bg-rose-50`;

    allDocs.forEach((doc) => {
        const fields = resolveDocFields(doc);
        const uploadState = docsState.uploads.get(fields.id);
        const downloadState = docsState.downloads.get(fields.id);
        const progressState = uploadState || downloadState;
        const downloaded = !uploadState && docsState.cache.has(fields.id);
        const row = document.createElement("div");
        row.className = "flex w-24 flex-col items-center gap-1 px-1 py-1.5 text-center";
        row.dataset.docId = fields.id;

        const rawFileTitle = fields.fileName || "Document";
        const rawSizeLabel = formatBytes(fields.fileSize);
        const fileTitle = escapeHtml(rawFileTitle);
        const sizeLabel = escapeHtml(rawSizeLabel);
        const hoverLabelText =
            rawSizeLabel && rawSizeLabel !== "-" ? `${rawFileTitle} (${rawSizeLabel})` : rawFileTitle;
        const hoverLabelAttr =
            sizeLabel && sizeLabel !== "-" ? `${fileTitle} (${sizeLabel})` : fileTitle;
        const detailsLabel = fields.docDetails ? escapeHtml(fields.docDetails) : "";
        const progressPercent = Math.round(Math.max(0, Math.min(100, Number(progressState?.percent) || 0)));
        const statusHtml = progressState
            ? `<div class="w-16">
                    <div class="h-1.5 w-full rounded-full bg-slate-100 overflow-hidden">
                        <div class="h-full bg-emerald-500 transition-all" style="width: ${Math.max(
                            5,
                            progressPercent
                        )}%"></div>
                    </div>
                </div>`
            : downloaded
                ? `<span class="inline-flex h-4 w-4 items-center justify-center rounded-full border border-emerald-200 bg-emerald-100 text-emerald-700" title="Downloaded">
                        <svg class="h-2.5 w-2.5" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                            <path d="M5 10l3 3 7-7"></path>
                        </svg>
                    </span>`
                : "<span class=\"h-4 w-4\"></span>";
        const iconHtml = getDocIconSvg(doc);
        const actionGridClass = "grid grid-cols-3 gap-1 w-full place-items-center";
        const actionButtonClass = uploadState
            ? `${actionButtonNeutral} opacity-50 pointer-events-none`
            : actionButtonNeutral;
        const actionDeleteClass = uploadState
            ? `${actionButtonDanger} opacity-50 pointer-events-none`
            : actionButtonDanger;
        row.title = hoverLabelText;

        row.innerHTML = `
            <div class="flex flex-col items-center gap-1">
                ${statusHtml}
                <div class="h-14 w-14 shrink-0 cursor-pointer rounded-xl transition-transform hover:scale-105" data-doc-icon="true" role="button" aria-label="Open document" title="${hoverLabelAttr}">
                    ${iconHtml}
                </div>
            </div>
            <div class="min-w-0">
                ${detailsLabel ? `<div class="text-[10px] text-slate-500 truncate">${detailsLabel}</div>` : ""}
            </div>
            <div class="${actionGridClass}">
                <button type="button" data-action="download" class="${actionButtonClass}" title="Download" aria-label="Download">
                    ${icons.download}
                </button>
                <button type="button" data-action="copy" class="${actionButtonClass}" title="Copy" aria-label="Copy">
                    ${icons.copy}
                </button>
                <button type="button" data-action="delete" class="${actionDeleteClass}" title="Delete" aria-label="Delete">
                    ${icons.delete}
                </button>
            </div>
        `;
        const iconTrigger = row.querySelector("[data-doc-icon]");
        if (iconTrigger) {
            iconTrigger.addEventListener("dblclick", () => {
                const previousTransform = iconTrigger.style.transform;
                iconTrigger.style.transform = "scale(1.08)";
                iconTrigger.classList.add("ring-2", "ring-emerald-200");
                setTimeout(() => {
                    iconTrigger.style.transform = previousTransform;
                    iconTrigger.classList.remove("ring-2", "ring-emerald-200");
                }, 180);
                handleDocOpen(doc);
            });
        }
        list.appendChild(row);
    });
}

async function setDocsForTenant(tenant) {
    const allDocs = await getLocalList(LOCAL_KEYS.docs, []);
    docsState.docs = allDocs
        .filter((doc) => docMatchesTenant(doc, tenant))
        .sort((a, b) => {
            const aTime = new Date(resolveDocFields(a).uploadedAt || 0).getTime();
            const bTime = new Date(resolveDocFields(b).uploadedAt || 0).getTime();
            return bTime - aTime;
        });
    await pruneMissingCachedDocs(docsState.docs);
    renderDocsList();
}

async function refreshDocsFromServer(tenant) {
    const result = await fetchTenantDocuments();
    if (Array.isArray(result?.docs)) {
        await setLocalData(LOCAL_KEYS.docs, result.docs);
        document.dispatchEvent(new CustomEvent("docs:updated", { detail: result.docs }));
    }
}

function setTenantLabel(tenant) {
    const { tenantName } = getElements();
    if (tenantName) tenantName.textContent = getTenantName(tenant);
}

function createUploadPlaceholder({ file, tenant, docType, docDetails, uploadId }) {
    return {
        doc_id: uploadId,
        tenant_id: getTenantId(tenant),
        tenant_name: getTenantName(tenant),
        doc_type: docType,
        doc_details: docDetails,
        file_name: file?.name || "document",
        file_size: Number(file?.size) || 0,
        uploaded_at: new Date().toISOString(),
    };
}

async function handleUploadClick() {
    if (docsState.uploadInProgress) return;
    const tenant = docsState.tenant;
    if (!tenant) {
        showToast("Select a tenant first.", "warning");
        return;
    }
    const { fileInput, typeSelect, otherInput } = getElements();
    if (!fileInput || !fileInput.files || !fileInput.files.length) {
        showToast("Choose a document to upload.", "warning");
        return;
    }
    const selection = resolveDocTypeSelection(typeSelect?.value || "");
    if (selection === "default") {
        showToast("Select a document type.", "warning");
        return;
    }
    const files = Array.from(fileInput.files);
    docsState.uploadInProgress = true;

    try {
        for (const file of files) {
            if (!isAllowedFile(file)) {
                showToast(`Unsupported format: ${file.name}`, "warning");
                continue;
            }
            if (file.size > MAX_DOC_BYTES) {
                showToast(`File exceeds 10MB: ${file.name}`, "warning");
                continue;
            }

            const docType = selection;
            const docDetails = docType === "other" ? normalizeId(otherInput?.value || "") : "";
            const uploadId = `upload-${Date.now()}-${Math.random().toString(16).slice(2)}`;
            const placeholder = createUploadPlaceholder({ file, tenant, docType, docDetails, uploadId });
            docsState.uploadPlaceholders.set(uploadId, placeholder);
            setDocUploadState(uploadId, { percent: 5 });
            renderDocsList();

            try {
                const base64 = await readFileAsBase64(file);

                const payload = {
                    tenantId: getTenantId(tenant),
                    tenantName: getTenantName(tenant),
                    docType,
                    docDetails,
                    fileName: file.name,
                    mimeType: file.type || "application/octet-stream",
                    size: file.size,
                    compression: "none",
                    dataBase64: base64,
                };

                setDocUploadState(uploadId, { percent: 12 });
                const result = await uploadTenantDocument(payload, {
                    onProgress: ({ loaded, total, percent }) => {
                        const safePercent = percent ?? (total ? (loaded / total) * 100 : 0);
                        setDocUploadState(uploadId, {
                            percent: Math.min(100, Math.max(12, safePercent)),
                        });
                    },
                });

                if (!result?.ok || !result?.doc) {
                    docsState.uploadPlaceholders.delete(uploadId);
                    setDocUploadState(uploadId, null);
                    showToast(`Upload failed for ${file.name}`, "error");
                    continue;
                }
                if (docType && !result.doc.doc_type && !result.doc.docType) {
                    result.doc.doc_type = docType;
                }
                if (docDetails && !result.doc.doc_details && !result.doc.docDetails) {
                    result.doc.doc_details = docDetails;
                }

                const current = await getLocalList(LOCAL_KEYS.docs, []);
                const next = current.filter((doc) => resolveDocFields(doc).id !== resolveDocFields(result.doc).id);
                next.unshift(result.doc);
                await setLocalData(LOCAL_KEYS.docs, next);
                docsState.uploadPlaceholders.delete(uploadId);
                setDocUploadState(uploadId, null);
                await setDocsForTenant(tenant);
                document.dispatchEvent(new CustomEvent("docs:updated", { detail: next }));
                showToast(`${file.name} uploaded`, "success");
            } catch (err) {
                docsState.uploadPlaceholders.delete(uploadId);
                setDocUploadState(uploadId, null);
                console.error("Document upload failed", err);
                showToast(`Document upload failed for ${file.name}.`, "error");
            }
        }
    } finally {
        docsState.uploadInProgress = false;
        if (fileInput) fileInput.value = "";
    }
}

async function handleDocOpen(doc) {
    if (!doc) return;
    const fields = resolveDocFields(doc);
    if (!window.__TAURI__) {
        if (fields.fileUrl) window.open(fields.fileUrl, "_blank", "noopener");
        return;
    }
    try {
        let filePath = await getCachedDocPath(doc);
        if (!filePath) {
            filePath = await downloadDocToCache(doc, {
                downloadId: `doc-open-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            });
        }
        try {
            await openLocalFile(filePath);
        } catch (err) {
            const exists = await fileExists(filePath);
            if (exists !== true) {
                const refreshedPath = await downloadDocToCache(doc, {
                    downloadId: `doc-open-retry-${Date.now()}-${Math.random().toString(16).slice(2)}`,
                });
                await openLocalFile(refreshedPath);
            } else {
                throw err;
            }
        }
    } catch (err) {
        console.error("Open document failed", err);
        showToast("Could not open document.", "error");
    }
}

async function handleDocDownload(doc) {
    const fields = resolveDocFields(doc);
    if (!window.__TAURI__) {
        if (fields.fileUrl) {
            const anchor = document.createElement("a");
            anchor.href = fields.fileUrl;
            anchor.download = fields.fileName || "document";
            anchor.target = "_blank";
            anchor.rel = "noopener";
            document.body.appendChild(anchor);
            anchor.click();
            anchor.remove();
        }
        return;
    }

    const location = await ensureDownloadLocationConfigured({ source: "tenant-docs" });
    if (!location?.ok || !location.path) return;
    showToast("Downloading document...", "info");
    const baseDir = location.path || "";
    const pathApi = window.__TAURI__?.path;
    const fs = window.__TAURI__?.fs;
    if (!fs) {
        showToast("Download is unavailable in this build.", "error");
        return;
    }

    const safeName = sanitizeFileSegment(getFileStem(fields.fileName || "document"));
    const ext = getFileExtension(fields.fileName);
    let finalName = ext ? `${safeName}.${ext}` : safeName;
    const separator = baseDir.includes("\\") ? "\\" : "/";
    let targetPath = pathApi?.join
        ? await pathApi.join(baseDir, finalName)
        : `${baseDir}${baseDir.endsWith(separator) ? "" : separator}${finalName}`;
    if (await fileExists(targetPath)) {
        const suffix = `_${Date.now()}`;
        finalName = ext ? `${safeName}${suffix}.${ext}` : `${safeName}${suffix}`;
        targetPath = pathApi?.join
            ? await pathApi.join(baseDir, finalName)
            : `${baseDir}${baseDir.endsWith(separator) ? "" : separator}${finalName}`;
    }

    try {
        let cachedPath = await getCachedDocPath(doc);
        if (cachedPath) {
            const exists = await fileExists(cachedPath);
            if (exists === false) cachedPath = "";
        }
        if (!cachedPath) {
            cachedPath = await downloadDocToCache(doc, {
                downloadId: `doc-cache-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            });
        }
        try {
            await copyFileToPath(cachedPath, targetPath);
        } catch (err) {
            const refreshedPath = await downloadDocToCache(doc, {
                downloadId: `doc-cache-retry-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            });
            await copyFileToPath(refreshedPath, targetPath);
        }
        showTenantDocDownloadModal(finalName, targetPath);
        showToast("Document downloaded", "success");
    } catch (err) {
        console.error("Download failed", err);
        showToast("Failed to download document.", "error");
    }
}

async function handleDocCopy(doc) {
    const fields = resolveDocFields(doc);
    if (!window.__TAURI__) {
        showToast("Copy is available only in the desktop app.", "warning");
        return;
    }

    try {
        let filePath = await getCachedDocPath(doc);
        if (!filePath) {
            filePath = await downloadDocToCache(doc, {
                downloadId: `doc-copy-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            });
        }
        let bytes;
        try {
            bytes = await window.__TAURI__.fs.readFile(filePath);
        } catch (readErr) {
            const refreshedPath = await downloadDocToCache(doc, {
                downloadId: `doc-copy-retry-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            });
            bytes = await window.__TAURI__.fs.readFile(refreshedPath);
            filePath = refreshedPath;
        }
        const tauriInvoke = window.__TAURI__?.core?.invoke;
        if (typeof tauriInvoke === "function") {
            await tauriInvoke("copy_file_to_clipboard", { filePath });
            showToast("File copied. Paste it in any folder.", "success");
        } else {
            showToast("File copy is unavailable. Use Download.", "warning");
        }
    } catch (err) {
        console.error("Copy failed", err);
        showToast("Failed to copy document.", "error");
    }
}

async function handleDocDelete(doc) {
    if (!doc) return;
    const fields = resolveDocFields(doc);
    if (!fields.id) {
        showToast("Document ID missing.", "error");
        return;
    }
    try {
        docsState.downloads.delete(fields.id);
        docsState.uploads.delete(fields.id);
        docsState.uploadPlaceholders.delete(fields.id);
        const result = await deleteTenantDocument(fields.id);
        await removeCachedDocEntry(fields.id);
        if (result?.ok === false) {
            showToast("Failed to delete document.", "error");
            return;
        }
        showToast("Document deleted", "success");
    } catch (err) {
        console.error("Delete document failed", err);
        showToast("Failed to delete document.", "error");
    }
}

function bindListActions() {
    const { list } = getElements();
    if (!list || list.dataset.bound === "true") return;
    list.dataset.bound = "true";

    list.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof Element)) return;
        const actionBtn = target.closest("button[data-action]");
        if (!actionBtn) return;
        const action = actionBtn.dataset.action;
        const row = actionBtn.closest("[data-doc-id]");
        if (!row) return;
        const docId = row.dataset.docId;
        const doc = docsState.docs.find((item) => resolveDocFields(item).id === docId);
        if (!doc) return;
        if (action === "download") {
            handleDocDownload(doc);
        } else if (action === "copy") {
            handleDocCopy(doc);
        } else if (action === "delete") {
            handleDocDelete(doc);
        }
    });

}

function showTenantDocDownloadModal(fileName, filePath) {
    const { modal, closeBtn, fileName: nameEl, location, openLink } = getDownloadModalElements();
    if (!modal) return;

    if (nameEl) nameEl.textContent = fileName || "document";
    if (location) {
        location.textContent = filePath ? `Saved to: ${filePath}` : "Saved to: -";
    }

    if (openLink) {
        openLink.removeAttribute("download");
        openLink.removeAttribute("href");
        openLink.target = "";
        if (filePath) {
            if (window.__TAURI__) {
                openLink.href = "#";
                openLink.dataset.filepath = filePath;
            } else {
                openLink.href = filePath;
                openLink.target = "_blank";
                openLink.rel = "noopener";
            }
            openLink.dataset.openReady = "true";
        }
        openLink.classList.toggle("pointer-events-none", !filePath);
        openLink.classList.toggle("opacity-50", !filePath);
    }

    if (closeBtn && !closeBtn.dataset.bound) {
        closeBtn.dataset.bound = "true";
        closeBtn.addEventListener("click", () => hideModal(modal));
    }

    if (openLink && !openLink.dataset.bound) {
        openLink.dataset.bound = "true";
        openLink.addEventListener("click", async (event) => {
            if (!openLink.dataset.filepath) return;
            if (window.__TAURI__?.opener?.openPath) {
                event.preventDefault();
                try {
                    await window.__TAURI__.opener.openPath(openLink.dataset.filepath);
                } catch (err) {
                    console.error("Failed to open file", err);
                    showToast("Failed to open file.", "error");
                }
                hideModal(modal);
            }
        });
    }

    showModal(modal);
}

function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const result = typeof reader.result === "string" ? reader.result : "";
            const comma = result.indexOf(",");
            resolve(comma >= 0 ? result.slice(comma + 1) : "");
        };
        reader.onerror = () => {
            reject(reader.error || new Error("Failed to read file"));
        };
        reader.readAsDataURL(file);
    });
}

function bindModalEvents() {
    if (docsState.modalBound) return;
    docsState.modalBound = true;
    const { modal, closeBtn, uploadBtn, typeSelect } = getElements();

    if (closeBtn) closeBtn.addEventListener("click", () => modal && hideModal(modal));
    if (modal) {
        modal.addEventListener("click", (event) => {
            if (event.target === modal) hideModal(modal);
        });
    }
    if (uploadBtn) uploadBtn.addEventListener("click", handleUploadClick);
    if (typeSelect) {
        typeSelect.addEventListener("change", syncDocTypeFields);
        syncDocTypeFields();
    }
    bindListActions();
}

export function initTenantDocuments() {
    bindModalEvents();
    document.addEventListener("docs:updated", () => {
        if (!docsState.tenant) return;
        setDocsForTenant(docsState.tenant);
    });
}

export async function openTenantDocumentsModal(tenant) {
    if (!tenant) {
        showToast("Select a tenant first.", "warning");
        return;
    }
    docsState.tenant = tenant;
    setTenantLabel(tenant);
    bindModalEvents();
    syncDocTypeFields();
    await ensureDocsCacheLoaded();
    await setDocsForTenant(tenant);
    const { modal } = getElements();
    if (modal) showModal(modal);
    refreshDocsFromServer(tenant).catch(() => null);
}

