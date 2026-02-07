import { escapeHtml } from "../../utils/htmlUtils.js";
import { hideModal, showModal } from "../../utils/ui.js";

const MAX_LOG_ROWS = 500;
const STORAGE_KEY = "dashboard.syncDebugLogs.v1";

const syncDebugState = {
    initialized: false,
    logs: [],
    selectedId: "",
};

function getElements() {
    return {
        widget: document.getElementById("syncDebugWidget"),
        successCount: document.getElementById("syncDebugSuccessCount"),
        issueCount: document.getElementById("syncDebugIssueCount"),
        lastStatus: document.getElementById("syncDebugLastStatus"),
        expandBtn: document.getElementById("syncDebugExpand"),
        modal: document.getElementById("syncDebugModal"),
        modalClose: document.getElementById("syncDebugModalClose"),
        modalList: document.getElementById("syncDebugModalList"),
        modalEmpty: document.getElementById("syncDebugModalEmpty"),
        modalSuccessCount: document.getElementById("syncDebugModalSuccessCount"),
        modalIssueCount: document.getElementById("syncDebugModalIssueCount"),
        modalTotalCount: document.getElementById("syncDebugModalTotalCount"),
        clearBtn: document.getElementById("syncDebugClearAll"),
        detailMeta: document.getElementById("syncDebugModalDetailMeta"),
        detailBody: document.getElementById("syncDebugModalDetailBody"),
    };
}

function hasLocalStorage() {
    try {
        return typeof window !== "undefined" && !!window.localStorage;
    } catch (err) {
        return false;
    }
}

function normalizeStatus(raw) {
    const status = (raw || "").toString().trim().toLowerCase();
    if (status === "success" || status === "ok" || status === "synced") return "success";
    if (status === "warning" || status === "warn" || status === "paused" || status === "issue") return "warning";
    if (status === "error" || status === "failed" || status === "fail") return "error";
    return "info";
}

function formatStatusLabel(status) {
    if (status === "success") return "Success";
    if (status === "warning") return "Paused / Issue";
    if (status === "error") return "Error";
    return "Info";
}

function formatTimestamp(value) {
    const date = new Date(value || Date.now());
    if (Number.isNaN(date.getTime())) return "-";
    return date.toLocaleString("en-GB", {
        year: "numeric",
        month: "short",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
    });
}

function getStatusClasses(status, selected = false) {
    const selection = selected ? "ring-2 ring-slate-300" : "";
    if (status === "success") {
        return `border-emerald-200 bg-emerald-50 text-emerald-800 ${selection}`.trim();
    }
    if (status === "warning") {
        return `border-amber-200 bg-amber-50 text-amber-800 ${selection}`.trim();
    }
    if (status === "error") {
        return `border-rose-200 bg-rose-50 text-rose-800 ${selection}`.trim();
    }
    return `border-slate-200 bg-slate-50 text-slate-700 ${selection}`.trim();
}

function normalizeLog(raw = {}) {
    const timestamp = new Date(raw.timestamp || raw.ts || Date.now());
    const normalizedTimestamp = Number.isNaN(timestamp.getTime())
        ? new Date().toISOString()
        : timestamp.toISOString();
    return {
        id: (raw.id || `${Date.now()}-${Math.random().toString(16).slice(2, 10)}`).toString(),
        timestamp: normalizedTimestamp,
        stage: (raw.stage || "sync").toString(),
        action: (raw.action || "").toString(),
        label: (raw.label || "").toString(),
        status: normalizeStatus(raw.status),
        message: (raw.message || "").toString(),
        request: raw.request ?? null,
        response: raw.response ?? null,
        error: raw.error ? String(raw.error) : "",
        meta: raw.meta ?? null,
    };
}

function saveState() {
    if (!hasLocalStorage()) return;

    const selectedId = syncDebugState.selectedId || "";
    const candidates = [
        syncDebugState.logs.slice(0, MAX_LOG_ROWS),
        syncDebugState.logs.slice(0, 250),
        syncDebugState.logs.slice(0, 120),
        syncDebugState.logs.slice(0, 60),
        syncDebugState.logs.slice(0, 25),
    ];

    for (const logs of candidates) {
        try {
            const payload = { logs, selectedId };
            localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
            if (logs.length !== syncDebugState.logs.length) {
                syncDebugState.logs = logs;
                if (!syncDebugState.logs.some((entry) => entry.id === syncDebugState.selectedId)) {
                    syncDebugState.selectedId = syncDebugState.logs[0]?.id || "";
                }
            }
            return;
        } catch (err) {
            // Try next smaller payload.
        }
    }
}

function loadSavedState() {
    if (!hasLocalStorage()) return;
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return;
        const parsed = JSON.parse(raw);
        const logsRaw = Array.isArray(parsed)
            ? parsed
            : Array.isArray(parsed?.logs)
                ? parsed.logs
                : [];
        syncDebugState.logs = logsRaw.map((item) => normalizeLog(item)).slice(0, MAX_LOG_ROWS);
        syncDebugState.selectedId = (parsed?.selectedId || "").toString();
        if (!syncDebugState.logs.some((entry) => entry.id === syncDebugState.selectedId)) {
            syncDebugState.selectedId = syncDebugState.logs[0]?.id || "";
        }
    } catch (err) {
        syncDebugState.logs = [];
        syncDebugState.selectedId = "";
    }
}

function clearSavedState() {
    if (!hasLocalStorage()) return;
    try {
        localStorage.removeItem(STORAGE_KEY);
    } catch (err) {
        // Ignore storage errors.
    }
}

function summarizeRow(entry) {
    const label = entry.label || entry.action || "Sync row";
    if (entry.message) return entry.message;
    return `${label} (${formatStatusLabel(entry.status)})`;
}

function computeCounts(logs) {
    let success = 0;
    let issue = 0;
    logs.forEach((entry) => {
        if (entry.status === "success") success += 1;
        if (entry.status === "warning" || entry.status === "error") issue += 1;
    });
    return { success, issue };
}

function renderWidget() {
    const { widget, successCount, issueCount, lastStatus } = getElements();
    if (!widget) return;
    const { success, issue } = computeCounts(syncDebugState.logs);
    if (successCount) successCount.textContent = String(success);
    if (issueCount) issueCount.textContent = String(issue);

    if (!lastStatus) return;
    const latest = syncDebugState.logs[0];
    if (!latest) {
        lastStatus.className = "text-[10px] text-slate-500";
        lastStatus.textContent = "No sync logs yet.";
        return;
    }

    if (latest.status === "success") {
        lastStatus.className = "text-[10px] text-emerald-700";
    } else if (latest.status === "warning") {
        lastStatus.className = "text-[10px] text-amber-700";
    } else if (latest.status === "error") {
        lastStatus.className = "text-[10px] text-rose-700";
    } else {
        lastStatus.className = "text-[10px] text-slate-600";
    }
    lastStatus.textContent = summarizeRow(latest);
}

function renderDetail(entry) {
    const { detailMeta, detailBody } = getElements();
    if (!detailMeta || !detailBody) return;
    if (!entry) {
        detailMeta.textContent = "Select a row to inspect synced payload.";
        detailBody.textContent = "";
        return;
    }
    detailMeta.textContent = `${formatTimestamp(entry.timestamp)} | ${entry.stage || "sync"} | ${formatStatusLabel(
        entry.status
    )}`;

    const payload = {
        id: entry.id,
        timestamp: entry.timestamp,
        stage: entry.stage,
        action: entry.action,
        label: entry.label,
        status: entry.status,
        message: entry.message,
        request: entry.request,
        response: entry.response,
        error: entry.error || null,
        meta: entry.meta,
    };
    detailBody.textContent = JSON.stringify(payload, null, 2);
}

function renderModal() {
    const {
        modalList,
        modalEmpty,
        modalSuccessCount,
        modalIssueCount,
        modalTotalCount,
    } = getElements();
    if (!modalList) return;

    modalList.innerHTML = "";
    const { success, issue } = computeCounts(syncDebugState.logs);
    if (modalSuccessCount) modalSuccessCount.textContent = `${success} success`;
    if (modalIssueCount) modalIssueCount.textContent = `${issue} issue / paused`;
    if (modalTotalCount) modalTotalCount.textContent = `${syncDebugState.logs.length} total`;

    if (!syncDebugState.logs.length) {
        if (modalEmpty) modalEmpty.classList.remove("hidden");
        renderDetail(null);
        return;
    }
    if (modalEmpty) modalEmpty.classList.add("hidden");

    const selected = syncDebugState.logs.find((entry) => entry.id === syncDebugState.selectedId);
    if (!selected) {
        syncDebugState.selectedId = syncDebugState.logs[0].id;
    }

    syncDebugState.logs.forEach((entry) => {
        const row = document.createElement("button");
        row.type = "button";
        row.dataset.syncDebugId = entry.id;
        row.className = `w-full text-left px-2.5 py-2 rounded border transition-colors ${getStatusClasses(
            entry.status,
            entry.id === syncDebugState.selectedId
        )}`;

        const title = escapeHtml(entry.label || entry.action || "Sync row");
        const summary = escapeHtml(summarizeRow(entry));
        const time = escapeHtml(formatTimestamp(entry.timestamp));
        const stage = escapeHtml(entry.stage || "sync");
        const status = escapeHtml(formatStatusLabel(entry.status));
        row.innerHTML = `
            <div class="flex items-start justify-between gap-2">
                <div class="min-w-0">
                    <div class="text-[11px] font-semibold truncate">${title}</div>
                    <div class="text-[10px] opacity-90 break-words">${summary}</div>
                </div>
                <span class="shrink-0 text-[10px] font-semibold">${status}</span>
            </div>
            <div class="mt-1 text-[9px] opacity-80">${stage} | ${time}</div>
        `;
        modalList.appendChild(row);
    });

    const active = syncDebugState.logs.find((entry) => entry.id === syncDebugState.selectedId) || null;
    renderDetail(active);
}

function renderAll() {
    renderWidget();
    renderModal();
}

function handleLogClick(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const row = target.closest("[data-sync-debug-id]");
    if (!(row instanceof HTMLElement)) return;
    const id = row.dataset.syncDebugId || "";
    if (!id) return;
    syncDebugState.selectedId = id;
    saveState();
    renderModal();
}

function appendLog(rawDetail = {}) {
    const log = normalizeLog(rawDetail);
    syncDebugState.logs = [log, ...syncDebugState.logs].slice(0, MAX_LOG_ROWS);
    if (!syncDebugState.selectedId) syncDebugState.selectedId = log.id;
    saveState();
    renderAll();
}

function clearAllLogs() {
    const hasRows = syncDebugState.logs.length > 0;
    if (hasRows) {
        const confirmed = window.confirm("Clear all saved sync debug logs?");
        if (!confirmed) return;
    }
    syncDebugState.logs = [];
    syncDebugState.selectedId = "";
    clearSavedState();
    renderAll();
}

function bindModalEvents() {
    const { expandBtn, modal, modalClose, modalList, clearBtn } = getElements();

    if (expandBtn && !expandBtn.dataset.bound) {
        expandBtn.dataset.bound = "1";
        expandBtn.addEventListener("click", () => {
            if (modal) showModal(modal);
        });
    }

    if (modal && !modal.dataset.bound) {
        modal.dataset.bound = "1";
        modal.addEventListener("click", (event) => {
            if (event.target === modal) hideModal(modal);
        });
    }

    if (modalClose && !modalClose.dataset.bound) {
        modalClose.dataset.bound = "1";
        modalClose.addEventListener("click", () => {
            if (modal) hideModal(modal);
        });
    }

    if (modalList && !modalList.dataset.bound) {
        modalList.dataset.bound = "1";
        modalList.addEventListener("click", handleLogClick);
    }

    if (clearBtn && !clearBtn.dataset.bound) {
        clearBtn.dataset.bound = "1";
        clearBtn.addEventListener("click", clearAllLogs);
    }
}

export function initSyncDebug() {
    if (syncDebugState.initialized) return;
    syncDebugState.initialized = true;

    const { widget } = getElements();
    if (!widget) return;

    loadSavedState();
    bindModalEvents();
    document.addEventListener("sync:debug-log", (event) => {
        appendLog(event?.detail || {});
    });

    renderAll();
}
