import { cacheDelete, cacheGet, cacheSet } from "./localDb.js";

export const LOCAL_KEYS = {
    wings: "data:wings",
    units: "data:units",
    landlords: "data:landlords",
    tenants: "data:tenants",
    tenancies: "data:tenancies",
    familyMembers: "data:familyMembers",
    clauses: "data:clauses",
    payments: "data:payments",
    attachments: "data:attachments",
    billLines: "data:billLines",
    wingMonthlyConfig: "data:wingMonthlyConfig",
    tenantMonthlyReadings: "data:tenantMonthlyReadings",
    generatedBills: "data:generatedbills",
    rentRevisionsAll: "data:rentRevisions:all",
    rentRevisions: (tenancyId) => `data:rentRevisions:${tenancyId || "unknown"}`,
    allSheets: "data:allSheets",
};

function coerceArray(value) {
    return Array.isArray(value) ? value : [];
}

function normalizeId(value) {
    return value === undefined || value === null ? "" : String(value);
}

export async function getLocalEntry(key) {
    if (!key) return null;
    const entry = await cacheGet(key);
    if (!entry || typeof entry !== "object") return null;
    return entry;
}

export async function getLocalData(key, fallback = null) {
    const entry = await getLocalEntry(key);
    if (!entry) return fallback;
    return entry.value ?? fallback;
}

export async function getLocalList(key, fallback = []) {
    const data = await getLocalData(key, fallback);
    return coerceArray(data);
}

export async function setLocalData(key, value) {
    if (!key) return null;
    await cacheSet(key, value ?? null);
    return value;
}

export async function clearLocalData(key) {
    if (!key) return false;
    await cacheDelete(key);
    return true;
}

export async function updateLocalData(key, updater, fallback = null) {
    const current = await getLocalData(key, fallback);
    const next = typeof updater === "function" ? updater(current) : current;
    if (typeof next !== "undefined") {
        await setLocalData(key, next);
        return next;
    }
    return current;
}

export async function upsertLocalListItem(key, idField, item) {
    if (!key || !idField || !item) return coerceArray(await getLocalData(key, []));
    return updateLocalData(
        key,
        (list) => {
            const next = coerceArray(list).slice();
            const idValue = normalizeId(item[idField]);
            if (!idValue) {
                next.push({ ...item });
                return next;
            }
            const idx = next.findIndex((row) => normalizeId(row?.[idField]) === idValue);
            if (idx >= 0) {
                next[idx] = { ...next[idx], ...item };
            } else {
                next.push({ ...item });
            }
            return next;
        },
        []
    );
}

export async function removeLocalListItem(key, idField, idValue) {
    if (!key || !idField) return coerceArray(await getLocalData(key, []));
    const matchId = normalizeId(idValue);
    return updateLocalData(
        key,
        (list) => coerceArray(list).filter((row) => normalizeId(row?.[idField]) !== matchId),
        []
    );
}

export async function replaceLocalList(key, list) {
    return setLocalData(key, coerceArray(list));
}

export async function appendLocalListItems(key, items) {
    return updateLocalData(
        key,
        (list) => coerceArray(list).concat(coerceArray(items)),
        []
    );
}
