import { LOCAL_KEYS, getLocalList } from "../../api/localStore.js";

const MAX_ATTACHMENTS = 2;
const MAX_DOCS = 2;

const recentFilesState = {
    initialized: false,
    loadingPromise: null,
    attachments: [],
    docs: [],
};

function getElements() {
    return {
        widget: document.getElementById("recentFilesWidget"),
        attachmentCount: document.getElementById("recentFilesAttachmentCount"),
        attachmentEmpty: document.getElementById("recentFilesAttachmentsEmpty"),
        attachmentList: document.getElementById("recentFilesAttachmentsList"),
        docCount: document.getElementById("recentFilesDocCount"),
        docEmpty: document.getElementById("recentFilesDocsEmpty"),
        docList: document.getElementById("recentFilesDocsList"),
    };
}

function toText(value) {
    return value === undefined || value === null ? "" : String(value).trim();
}

function toTimestamp(value) {
    if (!value) return 0;
    const parsed = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(parsed.getTime())) return 0;
    return parsed.getTime();
}

function formatTimestamp(value) {
    const timestamp = Number(value) || 0;
    if (!timestamp) return "";
    const date = new Date(timestamp);
    if (Number.isNaN(date.getTime())) return "";
    return date.toLocaleString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
    });
}

function formatDocType(value) {
    const raw = toText(value).replace(/[_-]+/g, " ");
    if (!raw) return "";
    return raw
        .split(" ")
        .filter(Boolean)
        .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
        .join(" ");
}

async function openExternalUrl(url) {
    const target = toText(url);
    if (!target) return false;
    const normalizedTarget =
        /^(https?:|mailto:|tel:|file:|data:)/i.test(target) || target.startsWith("//")
            ? target
            : `https://${target}`;

    try {
        const opener = window.__TAURI__?.opener;
        if (opener?.openUrl) {
            await opener.openUrl(normalizedTarget);
            return true;
        }
    } catch (err) {
        console.error("Unable to open link via opener plugin", err);
    }

    try {
        const shell = window.__TAURI__?.shell;
        if (shell?.open) {
            await shell.open(normalizedTarget);
            return true;
        }
    } catch (err) {
        console.error("Unable to open link in default browser", err);
    }

    const anchor = document.createElement("a");
    anchor.href = normalizedTarget;
    anchor.target = "_blank";
    anchor.rel = "noopener noreferrer";
    anchor.style.display = "none";
    document.body.appendChild(anchor);
    anchor.click();
    requestAnimationFrame(() => anchor.remove());
    return true;
}

function buildPaymentLookup(payments) {
    const lookup = new Map();
    (Array.isArray(payments) ? payments : []).forEach((payment) => {
        const attachmentId = toText(payment?.attachmentId || payment?.attachment_id);
        if (!attachmentId) return;
        const current = lookup.get(attachmentId);
        const incomingTs = toTimestamp(
            payment?.paymentDateTime || payment?.createdAt || payment?.date || payment?.created_at
        );
        const existingTs = current?.__timestamp || 0;
        if (!current || incomingTs >= existingTs) {
            lookup.set(attachmentId, { ...payment, __timestamp: incomingTs });
        }
    });
    return lookup;
}

function dedupeRows(rows, keyBuilder) {
    const seen = new Set();
    return rows.filter((row) => {
        const key = toText(keyBuilder(row));
        if (!key || seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

function normalizeAttachmentRows(attachments, paymentLookup) {
    const rows = (Array.isArray(attachments) ? attachments : []).map((item) => {
        const attachmentId = toText(item?.attachment_id || item?.attachmentId || item?.id);
        const payment = attachmentId ? paymentLookup.get(attachmentId) : null;
        const fileName = toText(item?.file_name || item?.fileName || payment?.attachmentName) || "Attachment";
        const fileUrl = toText(item?.file_url || item?.fileUrl || payment?.attachmentUrl);
        const tenantName = toText(payment?.tenantName || payment?.tenant_name);
        const wing = toText(payment?.wing);
        const unit = toText(payment?.unitNumber || payment?.unit_number);
        const unitLabel = wing && unit ? `${wing} - ${unit}` : wing || unit;
        const monthLabel = toText(payment?.monthLabel || payment?.month_label || payment?.monthKey || payment?.month_key);
        const timestamp = toTimestamp(
            item?.uploaded_at ||
            item?.uploadedAt ||
            item?.created_at ||
            item?.createdAt ||
            payment?.paymentDateTime ||
            payment?.createdAt ||
            payment?.date
        );
        return {
            id: attachmentId || fileUrl || `${fileName}:${timestamp}`,
            title: fileName,
            url: fileUrl,
            context: [tenantName, unitLabel, monthLabel].filter(Boolean).join(" | "),
            timestamp,
        };
    });

    return dedupeRows(rows, (row) => row.id || row.url || row.title)
        .sort((a, b) => b.timestamp - a.timestamp);
}

function normalizeDocRows(docs) {
    const rows = (Array.isArray(docs) ? docs : []).map((item) => {
        const docId = toText(item?.doc_id || item?.docId || item?.id);
        const fileName = toText(item?.file_name || item?.fileName || item?.name) || "Document";
        const fileUrl = toText(item?.file_url || item?.fileUrl || item?.url);
        const tenantName = toText(item?.tenant_name || item?.tenantName);
        const docType = formatDocType(item?.doc_type || item?.docType);
        const details = toText(item?.doc_details || item?.docDetails);
        const timestamp = toTimestamp(
            item?.uploaded_at || item?.uploadedAt || item?.created_at || item?.createdAt
        );
        return {
            id: docId || fileUrl || `${fileName}:${timestamp}`,
            title: fileName,
            url: fileUrl,
            context: [tenantName, docType, details].filter(Boolean).join(" | "),
            timestamp,
        };
    });

    return dedupeRows(rows, (row) => row.id || row.url || row.title)
        .sort((a, b) => b.timestamp - a.timestamp);
}

function createRow(item) {
    const row = document.createElement("div");
    row.className = "rounded-md border border-slate-200 bg-white px-2 py-1.5";

    const top = document.createElement("div");
    top.className = "flex items-center justify-between gap-2";

    const title = document.createElement("p");
    title.className = "text-[10px] font-semibold text-slate-800 truncate";
    title.textContent = item.title || "-";

    const action = document.createElement("a");
    action.className = "text-[9px] font-semibold text-indigo-600 hover:underline shrink-0";
    action.textContent = "Open";

    if (item.url) {
        action.href = item.url;
        action.target = "_blank";
        action.rel = "noopener";
        action.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            void openExternalUrl(item.url);
        });
    } else {
        action.removeAttribute("href");
        action.classList.remove("text-indigo-600", "hover:underline");
        action.classList.add("text-slate-400", "cursor-not-allowed", "pointer-events-none");
        action.textContent = "No link";
    }

    top.appendChild(title);
    top.appendChild(action);
    row.appendChild(top);

    if (item.context) {
        const context = document.createElement("p");
        context.className = "text-[9px] text-slate-500 truncate mt-0.5";
        context.textContent = item.context;
        row.appendChild(context);
    }

    const when = formatTimestamp(item.timestamp);
    if (when) {
        const meta = document.createElement("p");
        meta.className = "text-[8px] text-slate-400 mt-0.5";
        meta.textContent = when;
        row.appendChild(meta);
    }

    return row;
}

function renderList({ list, empty, count, rows, emptyText, limit }) {
    if (!list || !empty || !count) return;
    list.innerHTML = "";

    const limited = rows.slice(0, limit);
    count.textContent = String(rows.length);

    if (!limited.length) {
        empty.textContent = emptyText;
        empty.classList.remove("hidden");
        return;
    }

    empty.classList.add("hidden");
    limited.forEach((item) => {
        list.appendChild(createRow(item));
    });
}

function renderLoading() {
    const { attachmentList, attachmentEmpty, docList, docEmpty } = getElements();
    if (attachmentList) attachmentList.innerHTML = "";
    if (docList) docList.innerHTML = "";
    if (attachmentEmpty) {
        attachmentEmpty.textContent = "Loading attachments...";
        attachmentEmpty.classList.remove("hidden");
    }
    if (docEmpty) {
        docEmpty.textContent = "Loading documents...";
        docEmpty.classList.remove("hidden");
    }
}

function renderAll() {
    const els = getElements();
    renderList({
        list: els.attachmentList,
        empty: els.attachmentEmpty,
        count: els.attachmentCount,
        rows: recentFilesState.attachments,
        emptyText: "No attachments yet.",
        limit: MAX_ATTACHMENTS,
    });
    renderList({
        list: els.docList,
        empty: els.docEmpty,
        count: els.docCount,
        rows: recentFilesState.docs,
        emptyText: "No documents yet.",
        limit: MAX_DOCS,
    });
}

async function loadRecentFiles(options = {}) {
    if (recentFilesState.loadingPromise && options.force !== true) {
        return recentFilesState.loadingPromise;
    }

    renderLoading();

    const run = (async () => {
        try {
            const [attachments, docs, payments] = await Promise.all([
                getLocalList(LOCAL_KEYS.attachments, []),
                getLocalList(LOCAL_KEYS.docs, []),
                getLocalList(LOCAL_KEYS.payments, []),
            ]);
            const paymentLookup = buildPaymentLookup(payments);
            recentFilesState.attachments = normalizeAttachmentRows(attachments, paymentLookup);
            recentFilesState.docs = normalizeDocRows(docs);
        } catch (err) {
            console.error("Failed to load recent files widget data", err);
            recentFilesState.attachments = [];
            recentFilesState.docs = [];
        }
        renderAll();
    })();

    recentFilesState.loadingPromise = run;
    try {
        await run;
    } finally {
        recentFilesState.loadingPromise = null;
    }
}

export function initRecentFiles() {
    if (recentFilesState.initialized) return;
    recentFilesState.initialized = true;

    const { widget } = getElements();
    if (!widget) return;

    document.addEventListener("payment:saved", () => {
        void loadRecentFiles({ force: true });
    });
    document.addEventListener("docs:updated", () => {
        void loadRecentFiles({ force: true });
    });
    document.addEventListener("sync:completed", () => {
        void loadRecentFiles({ force: true });
    });

    void loadRecentFiles();
}
