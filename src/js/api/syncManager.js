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
import { LOCAL_KEYS, setLocalData } from "./localStore.js";
import { showToast, updateSyncIndicator } from "../utils/ui.js";

let initialSyncRunning = false;
let queueFlushRunning = false;
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
    saveClauses: ["clauses"],
    savePayment: ["payments", "billsminimal", "generatedbills"],
    deleteAttachment: ["payments"],
    saveBillingRecord: ["generatedbills", "billsminimal"],
    saveRentRevision: ["tenants"],
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
    if (data.generatedBills && Array.isArray(data.generatedBills.bills)) {
        await setLocalData(LOCAL_KEYS.generatedBills, data.generatedBills);
    }
}

function setProgressVisible(show) {
    const wrap = document.getElementById("syncProgressWrap");
    if (!wrap) return;
    wrap.classList.toggle("hidden", !show);
}

function updateSyncProgress(percent, label) {
    const bar = document.getElementById("syncProgressBar");
    const percentLabel = document.getElementById("syncProgressPercent");
    const textLabel = document.getElementById("syncProgressLabel");
    if (bar) bar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
    if (percentLabel) percentLabel.textContent = `${Math.round(percent)}%`;
    if (textLabel && label) textLabel.textContent = label;
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
    updateSyncIndicator(pending > 0 ? "pending" : "synced");
}

export async function startInitialSync() {
    if (initialSyncRunning) return { ok: false, reason: "already-running" };
    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to sync data", "warning"),
    });
    if (!url) return { ok: false, reason: "missing-url" };

    initialSyncRunning = true;
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
    initialSyncRunning = false;

    if (errors.length) {
        updateSyncIndicator("pending", "Sync incomplete");
    } else {
        updateSyncIndicator("synced");
    }

    await flushSyncQueue();
    return { ok: errors.length === 0, errors };
}

export async function enqueueSyncJob({ action, payload, method = "POST", params = {} } = {}) {
    if (!action) return null;
    const id = await queueAdd({ action, payload, method, params });
    updateSyncIndicator("pending");
    if (navigator.onLine) {
        flushSyncQueue();
    }
    return id;
}

export async function flushSyncQueue() {
    if (queueFlushRunning) return;
    const url = ensureAppScriptUrl();
    if (!url) {
        updateSyncIndicator("pending", "Sync pending");
        return;
    }

    queueFlushRunning = true;
    let jobs = await queueList();
    if (!jobs.length) {
        queueFlushRunning = false;
        updateSyncIndicator("synced");
        return;
    }

    updateSyncIndicator("syncing");
    for (const job of jobs) {
        try {
            await callAppScript({
                url,
                action: job.action,
                method: job.method || "POST",
                params: job.params || {},
                payload: job.payload || {},
                cache: { useLocal: false, revalidate: false },
            });
            await invalidateCachesForWriteAction(url, job.action);
            await queueDelete(job.id);
        } catch (err) {
            console.warn("Sync job failed", job.action, err);
            updateSyncIndicator("pending", "Sync paused");
            queueFlushRunning = false;
            return;
        }
    }

    jobs = await queueList();
    updateSyncIndicator(jobs.length ? "pending" : "synced");
    queueFlushRunning = false;
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
