import { ensureDownloadLocationConfigured } from "../../api/config.js";
import { fetchTenantDocuments, uploadTenantDocument } from "../../api/sheets.js";
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
};

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
        uploadBtn: document.getElementById("tenantDocsUploadBtn"),
        progressWrap: document.getElementById("tenantDocsProgressWrap"),
        progressLabel: document.getElementById("tenantDocsProgressLabel"),
        progressPercent: document.getElementById("tenantDocsProgressPercent"),
        progressBar: document.getElementById("tenantDocsProgressBar"),
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

function setProgressState({ show, label, percent }) {
    const { progressWrap, progressLabel, progressPercent, progressBar } = getElements();
    if (progressWrap) progressWrap.classList.toggle("hidden", !show);
    if (progressLabel) progressLabel.textContent = label || "";
    if (progressPercent) {
        const safePercent = Math.max(0, Math.min(100, Number(percent) || 0));
        progressPercent.textContent = `${Math.round(safePercent)}%`;
    }
    if (progressBar) {
        const safePercent = Math.max(0, Math.min(100, Number(percent) || 0));
        progressBar.style.width = `${safePercent}%`;
    }
}

function resolveDocFields(doc) {
    return {
        id: normalizeId(doc?.doc_id || doc?.docId || doc?.id),
        tenantId: normalizeId(doc?.tenant_id || doc?.tenantId || ""),
        tenantName: doc?.tenant_name || doc?.tenantName || "",
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

async function fileExists(filePath) {
    const fs = window.__TAURI__?.fs;
    if (!fs || !filePath) return false;
    if (typeof fs.exists === "function") {
        try {
            return await fs.exists(filePath);
        } catch (err) {
            return false;
        }
    }
    if (typeof fs.readFile === "function") {
        try {
            await fs.readFile(filePath);
            return true;
        } catch (err) {
            return false;
        }
    }
    return false;
}

async function ensureCacheDirectory(tenant) {
    const fs = window.__TAURI__?.fs;
    const pathApi = window.__TAURI__?.path;
    if (!fs || !pathApi) return "";
    const baseDir = typeof pathApi.appDataDir === "function"
        ? await pathApi.appDataDir()
        : typeof pathApi.appCacheDir === "function"
            ? await pathApi.appCacheDir()
            : typeof pathApi.downloadDir === "function"
                ? await pathApi.downloadDir()
                : "";
    if (!baseDir) return "";
    const docsDir = await pathApi.join(baseDir, "Tenant_Docs");
    const tenantSegment = sanitizeFileSegment(getTenantName(tenant));
    const tenantDir = await pathApi.join(docsDir, tenantSegment || "Tenant");
    try {
        await fs.createDir(tenantDir, { recursive: true });
    } catch (err) {
        // Ignore if exists.
    }
    return tenantDir;
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

async function getCachedDocPath(doc) {
    const id = resolveDocFields(doc).id;
    if (!id) return "";
    await ensureDocsCacheLoaded();
    const entry = docsState.cache.get(id);
    if (!entry?.path) return "";
    if (await fileExists(entry.path)) return entry.path;
    docsState.cache.delete(id);
    await persistDocsCache();
    return "";
}

async function downloadDocToCache(doc, options = {}) {
    const tenant = docsState.tenant;
    const cacheDir = await ensureCacheDirectory(tenant);
    if (!cacheDir) throw new Error("Cache directory is unavailable");
    const { id, fileName } = resolveDocFields(doc);
    const fs = window.__TAURI__?.fs;
    const pathApi = window.__TAURI__?.path;
    if (!fs || !pathApi) throw new Error("File system is unavailable");
    const baseName = sanitizeFileSegment(getFileStem(fileName || "document"));
    const ext = getFileExtension(fileName);
    const suffix = id ? `_${id.slice(0, 6)}` : `_${Date.now()}`;
    const safeName = ext ? `${baseName}${suffix}.${ext}` : `${baseName}${suffix}`;
    const filePath = await pathApi.join(cacheDir, safeName);

    const url = buildDriveDownloadUrl(doc);
    const resolvedOptions = { ...options };
    if (!resolvedOptions.downloadId) {
        resolvedOptions.downloadId = `doc-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    }
    await downloadFileToPath(url, filePath, resolvedOptions);

    docsState.cache.set(id, { path: filePath, fileName: fileName || safeName, cachedAt: Date.now() });
    await persistDocsCache();
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
                if (!exists) {
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

    if (!docsState.docs.length) {
        empty.classList.remove("hidden");
        updateDocCount(0);
        return;
    }

    empty.classList.add("hidden");
    updateDocCount(docsState.docs.length);

    docsState.docs.forEach((doc) => {
        const fields = resolveDocFields(doc);
        const downloaded = docsState.cache.has(fields.id);
        const row = document.createElement("div");
        row.className = "flex items-center justify-between gap-3 px-3 py-2";
        row.dataset.docId = fields.id;

        const nameHtml = escapeHtml(fields.fileName || "Document");
        const sizeLabel = escapeHtml(formatBytes(fields.fileSize));
        const dateLabel = escapeHtml(formatDateTime(fields.uploadedAt));
        const badge = downloaded
            ? `<span class="ml-2 text-[9px] px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 border border-emerald-200">Downloaded</span>`
            : "";

        row.innerHTML = `
            <div class="min-w-0">
                <div class="text-[11px] font-semibold text-slate-800 truncate">${nameHtml}${badge}</div>
                <div class="text-[10px] text-slate-500">${sizeLabel} • ${dateLabel}</div>
            </div>
            <div class="flex items-center gap-2 shrink-0">
                ${downloaded ? `<button type="button" data-action="open" class="px-2 py-1 rounded border border-slate-200 text-[10px] font-semibold text-slate-700 hover:bg-slate-100">Open</button>` : ""}
                <button type="button" data-action="download" class="px-2 py-1 rounded border border-slate-200 text-[10px] font-semibold text-slate-700 hover:bg-slate-100">Download</button>
                <button type="button" data-action="copy" class="px-2 py-1 rounded border border-slate-200 text-[10px] font-semibold text-slate-700 hover:bg-slate-100">Copy</button>
            </div>
        `;
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

async function handleUploadClick() {
    if (docsState.uploadInProgress) return;
    const tenant = docsState.tenant;
    if (!tenant) {
        showToast("Select a tenant first.", "warning");
        return;
    }
    const { fileInput } = getElements();
    if (!fileInput || !fileInput.files || !fileInput.files.length) {
        showToast("Choose a document to upload.", "warning");
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
            if (!window.PizZip) {
                showToast("Compression is unavailable in this build.", "error");
                continue;
            }

            setProgressState({ show: true, label: `Compressing ${file.name}...`, percent: 5 });
            const buffer = await file.arrayBuffer();
            const zip = new window.PizZip();
            zip.file(file.name, buffer);
            const zipped = zip.generate({ type: "uint8array", compression: "DEFLATE" });
            const base64 = bytesToBase64(zipped);

            const payload = {
                tenantId: getTenantId(tenant),
                tenantName: getTenantName(tenant),
                fileName: file.name,
                mimeType: file.type || "application/octet-stream",
                size: file.size,
                compression: "zip",
                dataBase64: base64,
            };

            setProgressState({ show: true, label: `Uploading ${file.name}...`, percent: 10 });
            const result = await uploadTenantDocument(payload, {
                onProgress: ({ loaded, total, percent }) => {
                    const safePercent = percent ?? (total ? (loaded / total) * 100 : 0);
                    setProgressState({
                        show: true,
                        label: `Uploading ${file.name}...`,
                        percent: Math.min(100, Math.max(10, safePercent)),
                    });
                },
            });

            if (!result?.ok || !result?.doc) {
                showToast(`Upload failed for ${file.name}`, "error");
                continue;
            }

            const current = await getLocalList(LOCAL_KEYS.docs, []);
            const next = current.filter((doc) => resolveDocFields(doc).id !== resolveDocFields(result.doc).id);
            next.unshift(result.doc);
            await setLocalData(LOCAL_KEYS.docs, next);
            await setDocsForTenant(tenant);
            document.dispatchEvent(new CustomEvent("docs:updated", { detail: next }));
            showToast(`${file.name} uploaded`, "success");
        }
    } catch (err) {
        console.error("Document upload failed", err);
        showToast("Document upload failed.", "error");
    } finally {
        docsState.uploadInProgress = false;
        setProgressState({ show: false, label: "", percent: 0 });
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
            setProgressState({ show: true, label: `Downloading ${fields.fileName || "document"}...`, percent: 5 });
            filePath = await downloadDocToCache(doc, {
                onProgress: ({ loaded, total, percent }) => {
                    const safePercent = percent ?? (total ? (loaded / total) * 100 : 0);
                    setProgressState({ show: true, label: "Downloading...", percent: safePercent });
                },
                downloadId: `doc-open-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            });
        }
        setProgressState({ show: true, label: "Opening file...", percent: 100 });
        try {
            await openLocalFile(filePath);
        } catch (err) {
            const exists = await fileExists(filePath);
            if (!exists) {
                setProgressState({ show: true, label: "File missing, downloading again...", percent: 10 });
                const refreshedPath = await downloadDocToCache(doc, {
                    onProgress: ({ loaded, total, percent }) => {
                        const safePercent = percent ?? (total ? (loaded / total) * 100 : 0);
                        setProgressState({ show: true, label: "Downloading...", percent: safePercent });
                    },
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
    } finally {
        setProgressState({ show: false, label: "", percent: 0 });
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
    if (!location?.ok) return;
    const baseDir = location.path || "";
    const pathApi = window.__TAURI__?.path;
    const fs = window.__TAURI__?.fs;
    if (!pathApi || !fs) {
        showToast("Download is unavailable in this build.", "error");
        return;
    }

    const safeName = sanitizeFileSegment(getFileStem(fields.fileName || "document"));
    const ext = getFileExtension(fields.fileName);
    let finalName = ext ? `${safeName}.${ext}` : safeName;
    let targetPath = await pathApi.join(baseDir, finalName);
    if (await fileExists(targetPath)) {
        const suffix = `_${Date.now()}`;
        finalName = ext ? `${safeName}${suffix}.${ext}` : `${safeName}${suffix}`;
        targetPath = await pathApi.join(baseDir, finalName);
    }

    try {
        setProgressState({ show: true, label: `Downloading ${fields.fileName || "document"}...`, percent: 5 });
        await downloadFileToPath(buildDriveDownloadUrl(doc), targetPath, {
            onProgress: ({ loaded, total, percent }) => {
                const safePercent = percent ?? (total ? (loaded / total) * 100 : 0);
                setProgressState({ show: true, label: "Downloading...", percent: safePercent });
            },
            downloadId: `doc-download-${Date.now()}-${Math.random().toString(16).slice(2)}`,
        });
        if (fields.id) {
            docsState.cache.set(fields.id, {
                path: targetPath,
                fileName: finalName,
                cachedAt: Date.now(),
                source: "download",
            });
            await persistDocsCache();
            renderDocsList();
        }
        setProgressState({ show: false, label: "", percent: 0 });
        showTenantDocDownloadModal(finalName, targetPath);
        showToast("Document downloaded", "success");
    } catch (err) {
        console.error("Download failed", err);
        showToast("Failed to download document.", "error");
        setProgressState({ show: false, label: "", percent: 0 });
    }
}

async function handleDocCopy(doc) {
    const fields = resolveDocFields(doc);
    if (!window.__TAURI__) {
        if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(fields.fileUrl || "");
            showToast("Link copied to clipboard", "success");
        }
        return;
    }

    try {
        setProgressState({ show: true, label: `Preparing ${fields.fileName || "document"}...`, percent: 5 });
        const filePath = await getOrDownloadCachedDoc(doc);
        const bytes = await window.__TAURI__.fs.readFile(filePath);
        const blob = new Blob([bytes], { type: fields.fileMime || "application/octet-stream" });

        if (navigator.clipboard && window.ClipboardItem) {
            const item = new ClipboardItem({ [blob.type]: blob });
            await navigator.clipboard.write([item]);
            showToast("Document copied to clipboard", "success");
        } else if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(filePath);
            showToast("File path copied to clipboard", "success");
        } else {
            showToast("Clipboard is unavailable.", "warning");
        }
    } catch (err) {
        console.error("Copy failed", err);
        if (navigator.clipboard?.writeText) {
            const fallbackText = fields.fileUrl || "";
            try {
                await navigator.clipboard.writeText(fallbackText);
                showToast("Document link copied to clipboard", "success");
                return;
            } catch (fallbackErr) {
                // fall through
            }
        }
        showToast("Failed to copy document.", "error");
    } finally {
        setProgressState({ show: false, label: "", percent: 0 });
    }
}

function bindListActions() {
    const { list } = getElements();
    if (!list || list.dataset.bound === "true") return;
    list.dataset.bound = "true";

    list.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        const actionBtn = target.closest("button[data-action]");
        if (!actionBtn) return;
        const action = actionBtn.dataset.action;
        const row = actionBtn.closest("[data-doc-id]");
        if (!row) return;
        const docId = row.dataset.docId;
        const doc = docsState.docs.find((item) => resolveDocFields(item).id === docId);
        if (!doc) return;
        if (action === "open") {
            handleDocOpen(doc);
        } else if (action === "download") {
            handleDocDownload(doc);
        } else if (action === "copy") {
            handleDocCopy(doc);
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

function bytesToBase64(bytes) {
    const chunkSize = 0x8000;
    let binary = "";
    for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
}

function bindModalEvents() {
    if (docsState.modalBound) return;
    docsState.modalBound = true;
    const { modal, closeBtn, uploadBtn } = getElements();

    if (closeBtn) closeBtn.addEventListener("click", () => modal && hideModal(modal));
    if (modal) {
        modal.addEventListener("click", (event) => {
            if (event.target === modal) hideModal(modal);
        });
    }
    if (uploadBtn) uploadBtn.addEventListener("click", handleUploadClick);
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
    await ensureDocsCacheLoaded();
    await setDocsForTenant(tenant);
    const { modal } = getElements();
    if (modal) showModal(modal);
    refreshDocsFromServer(tenant).catch(() => null);
}
