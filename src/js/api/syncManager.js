import { ensureAppScriptUrl } from "./appscriptClient.js";
import {
    cacheDeletePrefix,
    queueAdd,
    queueCount,
    queueClear,
    queueDelete,
    queueList,
} from "./localDb.js";
import { callAppScript } from "./appscriptClient.js";
import { LOCAL_KEYS, getLocalData, setLocalData } from "./localStore.js";
import { showToast, updateSyncIndicator } from "../utils/ui.js";
import { STORAGE_KEYS } from "../constants.js";

let initialSyncRunning = false;
let queueFlushRunning = false;
let queueFlushRequested = false;
const SYNC_TIMEOUT_MS = 30000;

const WRITE_INVALIDATIONS = {
    addWing: ["wings"],
    removeWing: ["wings"],
    saveUnit: ["units"],
    deleteUnit: ["units"],
    saveLandlord: ["landlords"],
    deleteLandlord: ["landlords"],
    saveTenant: ["tenants", "units"],
    updateTenant: ["tenants", "units"],
    vacateTenancy: ["tenants", "units"],
    saveClauses: ["clauses"],
    savePayment: ["payments", "billsminimal", "generatedbills"],
    deleteAttachment: ["payments"],
    deleteTenantDocument: ["docs"],
    saveBillingRecord: ["generatedbills", "billsminimal"],
    saveRentRevision: ["tenants"],
    saveRentRevisionReminder: ["rentrevisionreminders"],
    saveNotes: ["notes"],
};

const LOCAL_STORAGE_KEYS = [
    "cache.wings",
    "cache.landlords",
    "cache.units",
    "cache.clauses",
];

function clearLocalStorageCaches() {
    try {
        LOCAL_STORAGE_KEYS.forEach((key) => localStorage.removeItem(key));
    } catch (err) {
        console.warn("Local cache clear failed", err);
    }
}

async function clearPendingDocDelete(docId) {
    const raw = docId ? String(docId) : "";
    if (!raw) return;
    const pending = await getLocalData(LOCAL_KEYS.docsDeleted, []);
    const list = Array.isArray(pending) ? pending : [];
    const next = list.filter((id) => String(id) !== raw);
    if (next.length !== list.length) {
        await setLocalData(LOCAL_KEYS.docsDeleted, next);
    }
}

function resolveTenantDocId(doc) {
    if (!doc || typeof doc !== "object") return "";
    const id = doc.doc_id || doc.docId || doc.id || "";
    return id ? String(id) : "";
}

async function removeDeletedDocFromLocal(docId) {
    const raw = docId ? String(docId) : "";
    if (!raw) return;
    const docs = await getLocalData(LOCAL_KEYS.docs, []);
    if (Array.isArray(docs)) {
        const next = docs.filter((doc) => resolveTenantDocId(doc) !== raw);
        if (next.length !== docs.length) {
            await setLocalData(LOCAL_KEYS.docs, next);
            dispatchUpdateEvent("docs:updated", next);
        }
    }
    const cached = await getLocalData(LOCAL_KEYS.docsCache, {});
    if (cached && typeof cached === "object" && cached[raw]) {
        const nextCache = { ...cached };
        delete nextCache[raw];
        await setLocalData(LOCAL_KEYS.docsCache, nextCache);
    }
}

function runWithTimeout(promise, timeoutMs, label) {
    if (!timeoutMs) return promise;
    let timeoutId = null;
    const timeout = new Promise((_, reject) => {
        timeoutId = setTimeout(() => {
            reject(new Error(`${label || "Sync task"} timed out`));
        }, timeoutMs);
    });
    return Promise.race([promise, timeout]).finally(() => {
        if (timeoutId) clearTimeout(timeoutId);
    });
}

function dispatchUpdateEvent(name, detail) {
    if (typeof document === "undefined") return;
    document.dispatchEvent(new CustomEvent(name, { detail }));
}

function emitSyncBusy() {
    dispatchUpdateEvent("sync:busy", { busy: initialSyncRunning || queueFlushRunning });
}

export function isSyncBusy() {
    return initialSyncRunning || queueFlushRunning;
}

async function updateSyncIndicatorWithCount(status, options = {}) {
    const count = typeof options.count === "number" ? options.count : await queueCount();
    let label = options.message;
    if (!label) {
        if (status === "syncing") {
            label = count > 0 ? `Syncing (${count})` : "Syncing";
        } else if (status === "pending") {
            label = count > 0 ? `Not synced (${count})` : "Not synced";
        } else if (status === "synced") {
            label = "Synced";
        }
    }
    updateSyncIndicator(status, label || "");
    return count;
}

async function refreshAllSheetsSnapshot(url) {
    if (!url) return;
    try {
        const data = await callAppScript({
            url,
            action: "allsheets",
            cache: { useLocal: false, revalidate: false },
        });
        if (Array.isArray(data?.allSheets)) {
            await setLocalData(LOCAL_KEYS.allSheets, data.allSheets);
        }
    } catch (err) {
        console.warn("All-sheets sync failed", err);
    }
}

async function storeWingsList(wings = []) {
    const next = Array.isArray(wings)
        ? wings.map((w) => (w || "").toString().trim()).filter(Boolean)
        : [];
    await setLocalData(LOCAL_KEYS.wings, next);
    dispatchUpdateEvent("wings:updated", next);
    return next;
}

async function storeUnitsList(units = []) {
    const next = Array.isArray(units) ? units : [];
    await setLocalData(LOCAL_KEYS.units, next);
    dispatchUpdateEvent("units:updated", next);
    return next;
}

async function storeLandlordsList(landlords = []) {
    const next = Array.isArray(landlords) ? landlords : [];
    await setLocalData(LOCAL_KEYS.landlords, next);
    dispatchUpdateEvent("landlords:updated", next);
    return next;
}

async function storeTenantsList(tenants = []) {
    const next = Array.isArray(tenants) ? tenants : [];
    await setLocalData(LOCAL_KEYS.tenants, next);
    dispatchUpdateEvent("tenants:updated", next);
    return next;
}

async function storeRentRevisions(revisions = []) {
    const next = Array.isArray(revisions) ? revisions : [];
    await setLocalData(LOCAL_KEYS.rentRevisionsAll, next);
    const grouped = next.reduce((acc, rev) => {
        const tenancyId = rev?.tenancy_id || rev?.tenancyId || "";
        if (!tenancyId) return acc;
        if (!acc[tenancyId]) acc[tenancyId] = [];
        acc[tenancyId].push(rev);
        return acc;
    }, {});
    await Promise.all(
        Object.entries(grouped).map(([tenancyId, list]) =>
            setLocalData(LOCAL_KEYS.rentRevisions(tenancyId), list)
        )
    );
    return next;
}

async function storeRentRevisionReminders(reminders = []) {
    const next = Array.isArray(reminders) ? reminders : [];
    await setLocalData(LOCAL_KEYS.rentRevisionReminders, next);
    dispatchUpdateEvent("rentRevisionReminders:updated", next);
    return next;
}

async function storeNotesList(notes = []) {
    const next = Array.isArray(notes) ? notes : [];
    await setLocalData(LOCAL_KEYS.notes, next);
    dispatchUpdateEvent("notes:updated", next);
    return next;
}

async function applyExportAll(data) {
    if (!data || typeof data !== "object") return;
    if (Array.isArray(data.wings)) await storeWingsList(data.wings);
    if (Array.isArray(data.units)) await storeUnitsList(data.units);
    if (Array.isArray(data.landlords)) await storeLandlordsList(data.landlords);
    if (Array.isArray(data.tenants)) await storeTenantsList(data.tenants);
    if (data.clauses && typeof data.clauses === "object") {
        const payload = {
            tenant: Array.isArray(data.clauses.tenant) ? data.clauses.tenant : [],
            landlord: Array.isArray(data.clauses.landlord) ? data.clauses.landlord : [],
            penalties: Array.isArray(data.clauses.penalties) ? data.clauses.penalties : [],
            misc: Array.isArray(data.clauses.misc) ? data.clauses.misc : [],
        };
        await setLocalData(LOCAL_KEYS.clauses, payload);
    }
    if (Array.isArray(data.payments)) {
        await setLocalData(LOCAL_KEYS.payments, data.payments);
    }
    if (Array.isArray(data.attachments)) {
        await setLocalData(LOCAL_KEYS.attachments, data.attachments);
    }
    if (Array.isArray(data.docs)) {
        await setLocalData(LOCAL_KEYS.docs, data.docs);
        dispatchUpdateEvent("docs:updated", data.docs);
    }
    if (Array.isArray(data.billLines)) {
        await setLocalData(LOCAL_KEYS.billLines, data.billLines);
    }
    if (Array.isArray(data.wingMonthlyConfig)) {
        await setLocalData(LOCAL_KEYS.wingMonthlyConfig, data.wingMonthlyConfig);
    }
    if (Array.isArray(data.tenantMonthlyReadings)) {
        await setLocalData(LOCAL_KEYS.tenantMonthlyReadings, data.tenantMonthlyReadings);
    }
    if (Array.isArray(data.tenancies)) {
        await setLocalData(LOCAL_KEYS.tenancies, data.tenancies);
    }
    if (Array.isArray(data.familyMembers)) {
        await setLocalData(LOCAL_KEYS.familyMembers, data.familyMembers);
    }
    if (Array.isArray(data.rentRevisions)) {
        await storeRentRevisions(data.rentRevisions);
    }
    if (Array.isArray(data.rentRevisionReminders)) {
        await storeRentRevisionReminders(data.rentRevisionReminders);
    }
    if (data.generatedBills && Array.isArray(data.generatedBills.bills)) {
        await setLocalData(LOCAL_KEYS.generatedBills, data.generatedBills);
    }
    if (Array.isArray(data.notes)) {
        await storeNotesList(data.notes);
    }
    if (Array.isArray(data.allSheets)) {
        await setLocalData(LOCAL_KEYS.allSheets, data.allSheets);
    }
}

function setProgressVisible(show) {
    const wrap = document.getElementById("syncProgressWrap");
    if (!wrap) return;
    if (isGlobalProgressSuppressed_()) {
        wrap.classList.add("hidden");
        return;
    }
    wrap.classList.toggle("hidden", !show);
}

function updateSyncProgress(percent, label) {
    dispatchUpdateEvent("sync:progress", { percent, label });
    if (isGlobalProgressSuppressed_()) return;
    const bar = document.getElementById("syncProgressBar");
    const percentLabel = document.getElementById("syncProgressPercent");
    const textLabel = document.getElementById("syncProgressLabel");
    if (bar) bar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
    if (percentLabel) percentLabel.textContent = `${Math.round(percent)}%`;
    if (textLabel && label) textLabel.textContent = label;
}

function isGlobalProgressSuppressed_() {
    if (typeof document === "undefined") return false;
    const modal = document.getElementById("appscriptModal");
    if (modal?.dataset?.syncState === "syncing") return true;
    const autoModal = document.getElementById("autoSyncModal");
    return autoModal?.dataset?.syncState === "syncing";
}

function buildSyncTasks(url) {
    let exportAllDone = false;
    const skipIfExported = (fn) => async () => {
        if (exportAllDone) return null;
        return fn();
    };
    return [
        {
            label: "Syncing full dataset",
            run: async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "exportall", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Full dataset"
                );
                if (data && data.ok) {
                    await applyExportAll(data);
                    exportAllDone = true;
                }
                return data;
            },
        },
        {
            label: "Syncing wings",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "wings", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Wings"
                );
                if (Array.isArray(data?.wings)) {
                    await storeWingsList(data.wings);
                }
                return data;
            }),
        },
        {
            label: "Syncing units",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "units", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Units"
                );
                if (Array.isArray(data?.units)) {
                    await storeUnitsList(data.units);
                }
                return data;
            }),
        },
        {
            label: "Syncing landlords",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "landlords", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Landlords"
                );
                if (Array.isArray(data?.landlords)) {
                    await storeLandlordsList(data.landlords);
                }
                return data;
            }),
        },
        {
            label: "Syncing tenants",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "tenants", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Tenants"
                );
                if (Array.isArray(data?.tenants)) {
                    await storeTenantsList(data.tenants);
                }
                return data;
            }),
        },
        {
            label: "Syncing clauses",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "clauses", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Clauses"
                );
                if (data && typeof data === "object") {
                    await setLocalData(LOCAL_KEYS.clauses, {
                        tenant: Array.isArray(data.tenant) ? data.tenant : [],
                        landlord: Array.isArray(data.landlord) ? data.landlord : [],
                        penalties: Array.isArray(data.penalties) ? data.penalties : [],
                        misc: Array.isArray(data.misc) ? data.misc : [],
                    });
                }
                return data;
            }),
        },
        {
            label: "Syncing notes",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "notes", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Notes"
                );
                if (Array.isArray(data?.notes)) {
                    await storeNotesList(data.notes);
                }
                return data;
            }),
        },
        {
            label: "Syncing rent revisions",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "rentrevisions", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Rent revisions"
                );
                if (Array.isArray(data?.revisions)) {
                    await storeRentRevisions(data.revisions);
                }
                return data;
            }),
        },
        {
            label: "Syncing rent revision reminders",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "rentrevisionreminders", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Rent revision reminders"
                );
                if (Array.isArray(data?.reminders)) {
                    await storeRentRevisionReminders(data.reminders);
                }
                return data;
            }),
        },
        {
            label: "Syncing generated bills",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "generatedbills", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Generated bills"
                );
                if (data && Array.isArray(data.bills)) {
                    await setLocalData(LOCAL_KEYS.generatedBills, data);
                }
                return data;
            }),
        },
        {
            label: "Syncing payments",
            run: skipIfExported(async () => {
                const data = await runWithTimeout(
                    callAppScript({ url, action: "payments", cache: { useLocal: false } }),
                    SYNC_TIMEOUT_MS,
                    "Payments"
                );
                if (Array.isArray(data?.payments)) {
                    await setLocalData(LOCAL_KEYS.payments, data.payments);
                }
                return data;
            }),
        },
    ];
}

export async function initSyncManager() {
    const pending = await queueCount();
    await updateSyncIndicatorWithCount(pending > 0 ? "pending" : "synced", { count: pending });
    emitSyncBusy();
}

export async function startInitialSync() {
    if (initialSyncRunning) return { ok: false, reason: "already-running" };
    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to sync data", "warning"),
    });
    if (!url) return { ok: false, reason: "missing-url" };

    initialSyncRunning = true;
    emitSyncBusy();
    try {
        const pending = await queueCount();
        if (pending > 0) {
            const confirmed = window.confirm(
                `You have ${pending} unsynced change${pending === 1 ? "" : "s"}. Resyncing will discard them. Continue?`
            );
            if (!confirmed) {
                updateSyncIndicator("pending", "Sync paused");
                setProgressVisible(false);
                return { ok: false, reason: "cancelled" };
            }
        }

        updateSyncIndicator("syncing");
        setProgressVisible(true);
        updateSyncProgress(0, "Preparing sync...");
        await cacheDeletePrefix("");
        await queueClear();
        clearLocalStorageCaches();

        const tasks = buildSyncTasks(url);
        const total = tasks.length;
        const errors = [];

        for (let i = 0; i < tasks.length; i += 1) {
            const task = tasks[i];
            const percent = ((i) / total) * 100;
            updateSyncProgress(percent, task.label);
            try {
                await task.run();
            } catch (err) {
                console.warn("Initial sync task failed", task.label, err);
                errors.push({ label: task.label, error: String(err) });
            }
        }

        updateSyncProgress(100, errors.length ? "Sync finished with errors" : "Sync complete");
        setTimeout(() => setProgressVisible(false), 600);
        if (errors.length) {
            updateSyncIndicator("pending", "Sync incomplete");
        } else {
            updateSyncIndicator("synced");
        }

        await flushSyncQueue();
        const ok = errors.length === 0;
        if (ok) {
            try {
                localStorage.setItem(STORAGE_KEYS.LAST_FULL_SYNC_AT, String(Date.now()));
            } catch (err) {
                console.warn("Failed to persist sync timestamp", err);
            }
        }
        dispatchUpdateEvent("sync:completed", { ok, errors });
        return { ok, errors };
    } finally {
        initialSyncRunning = false;
        emitSyncBusy();
    }
}

export async function enqueueSyncJob({ action, payload, method = "POST", params = {} } = {}) {
    if (!action) return null;
    const id = await queueAdd({ action, payload, method, params });
    const pending = await queueCount();
    void updateSyncIndicatorWithCount(queueFlushRunning ? "syncing" : "pending", { count: pending });
    if (navigator.onLine) {
        if (queueFlushRunning) {
            queueFlushRequested = true;
        } else {
            flushSyncQueue();
        }
    }
    return id;
}

export async function flushSyncQueue() {
    if (queueFlushRunning) {
        queueFlushRequested = true;
        return;
    }
    const url = ensureAppScriptUrl();
    if (!url) {
        await updateSyncIndicatorWithCount("pending", { message: "Sync pending" });
        return;
    }

    queueFlushRunning = true;
    queueFlushRequested = false;
    emitSyncBusy();
    try {
        let pendingCount = await queueCount();
        if (!pendingCount) {
            await updateSyncIndicatorWithCount("synced", { count: 0 });
            return;
        }

        await updateSyncIndicatorWithCount("syncing", { count: pendingCount });
        let jobs = await queueList();
        while (jobs.length) {
            for (const job of jobs) {
                try {
                    const result = await callAppScript({
                        url,
                        action: job.action,
                        method: job.method || "POST",
                        params: job.params || {},
                        payload: job.payload || {},
                        cache: { useLocal: false, revalidate: false },
                    });
                    if (job.action === "deleteTenantDocument" && result?.ok === false) {
                        throw new Error("Delete document failed");
                    }
                    await invalidateCachesForWriteAction(url, job.action);
                    if (job.action === "deleteTenantDocument") {
                        const resolvedId =
                            result?.docId || job.payload?.docId || job.payload?.doc_id || "";
                        await removeDeletedDocFromLocal(resolvedId);
                        await clearPendingDocDelete(resolvedId);
                    }
                    await queueDelete(job.id);
                } catch (err) {
                    console.warn("Sync job failed", job.action, err);
                    updateSyncIndicator("pending", "Sync paused");
                    return;
                }

                pendingCount = await queueCount();
                if (pendingCount > 0) {
                    await updateSyncIndicatorWithCount("syncing", { count: pendingCount });
                }
            }

            jobs = await queueList();
        }

        await updateSyncIndicatorWithCount("synced", { count: 0 });
        await refreshAllSheetsSnapshot(url);
    } finally {
        queueFlushRunning = false;
        emitSyncBusy();
        if (queueFlushRequested && navigator.onLine) {
            queueFlushRequested = false;
            setTimeout(() => flushSyncQueue(), 0);
        }
    }
}

export async function invalidateCacheForAction(url, action) {
    if (!url || !action) return;
    const prefix = `${url}|${action}|`;
    await cacheDeletePrefix(prefix);
}

export async function invalidateCachesForWriteAction(url, writeAction) {
    const actions = WRITE_INVALIDATIONS[writeAction] || [];
    if (!actions.length) return;
    await Promise.all(actions.map((action) => invalidateCacheForAction(url, action)));
    clearLocalStorageCaches();
}
