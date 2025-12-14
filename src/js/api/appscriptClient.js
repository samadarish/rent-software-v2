/**
 * Shared Apps Script client helpers.
 * Centralizes URL resolution and JSON fetch/post handling so that
 * feature-specific API wrappers can stay concise.
 */
import { getAppScriptUrl } from "./config.js";
import { showModal } from "../utils/ui.js";

const inflightGets = new Map();
const responseCache = new Map();
const CACHE_TTL_MS = 2 * 60 * 1000;
const STORAGE_PREFIX = "appscriptCache:";

function isCacheEntryFresh(entry, ttlMs = CACHE_TTL_MS) {
    if (!entry || typeof entry.timestamp !== "number") return false;
    return Date.now() - entry.timestamp < ttlMs;
}

function getPersistedCache(cacheKey) {
    if (typeof localStorage === "undefined") return null;
    try {
        const raw = localStorage.getItem(`${STORAGE_PREFIX}${cacheKey}`);
        if (!raw) return null;
        return JSON.parse(raw);
    } catch (e) {
        console.warn("Unable to read cached response", e);
        return null;
    }
}

function setPersistedCache(cacheKey, entry) {
    if (typeof localStorage === "undefined") return;
    try {
        localStorage.setItem(`${STORAGE_PREFIX}${cacheKey}`, JSON.stringify(entry));
    } catch (e) {
        // Storage quota failures shouldn't break execution; just warn.
        console.warn("Unable to persist cached response", e);
    }
}

function clearPersistedCache(cacheKey) {
    if (typeof localStorage === "undefined") return;
    try {
        localStorage.removeItem(`${STORAGE_PREFIX}${cacheKey}`);
    } catch (e) {
        console.warn("Unable to clear cached response", e);
    }
}

function buildCacheKey({ url, action, params }) {
    const sortedEntries = Object.entries(params || {})
        .filter(([, value]) => value !== undefined && value !== null)
        .sort(([a], [b]) => a.localeCompare(b));
    const paramString = sortedEntries
        .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
        .join("&");
    return `${url}|${action}|${paramString}`;
}

function getCachedResponse(cacheKey, cacheTtlMs = CACHE_TTL_MS) {
    const inMemory = responseCache.get(cacheKey);
    if (isCacheEntryFresh(inMemory, cacheTtlMs)) return inMemory.data;

    if (inMemory) {
        responseCache.delete(cacheKey);
    }

    const persisted = getPersistedCache(cacheKey);
    if (isCacheEntryFresh(persisted, cacheTtlMs)) {
        responseCache.set(cacheKey, persisted);
        return persisted.data;
    }

    if (persisted) clearPersistedCache(cacheKey);
    return null;
}

function setCachedResponse(cacheKey, data) {
    const entry = { data, timestamp: Date.now() };
    responseCache.set(cacheKey, entry);
    setPersistedCache(cacheKey, entry);
}

export function getCachedGetResponse({ url, action, params = {}, cacheTtlMs } = {}) {
    if (!url || !action) return null;
    const cacheKey = buildCacheKey({ url, action, params });
    return getCachedResponse(cacheKey, cacheTtlMs);
}

export function invalidateCachedGets({ action } = {}) {
    const matchingKeys = [];
    const shouldCheckAction = Boolean(action);

    const maybeCollectKey = (cacheKey) => {
        if (!shouldCheckAction) {
            matchingKeys.push(cacheKey);
            return;
        }
        const [, cachedAction] = cacheKey.split("|");
        if (cachedAction === action) matchingKeys.push(cacheKey);
    };

    Array.from(responseCache.keys()).forEach(maybeCollectKey);

    if (typeof localStorage !== "undefined") {
        Object.keys(localStorage)
            .filter((key) => key.startsWith(STORAGE_PREFIX))
            .forEach((key) => {
                const cacheKey = key.replace(STORAGE_PREFIX, "");
                maybeCollectKey(cacheKey);
            });
    }

    matchingKeys.forEach((cacheKey) => {
        responseCache.delete(cacheKey);
        clearPersistedCache(cacheKey);
    });
}

function showConfigModal() {
    const modal = document.getElementById("appscriptModal");
    if (modal) showModal(modal);
}

export function ensureAppScriptUrl({ onMissing, promptForConfig = false } = {}) {
    const url = getAppScriptUrl();
    if (!url) {
        if (typeof onMissing === "function") onMissing();
        if (promptForConfig) showConfigModal();
    }
    return url;
}

export async function callAppScript({
    url,
    action,
    method = "GET",
    params = {},
    payload,
    cacheTtlMs = CACHE_TTL_MS,
    bypassCache = false,
}) {
    if (!url) return null;

    const search = new URLSearchParams({ action, ...params });
    const target = method === "GET" ? `${url}?${search.toString()}` : url;
    const options = {
        method,
        headers: {
            Accept: "application/json",
        },
        cache: "no-store",
    };

    if (method !== "GET") {
        // Use a simple text payload to avoid CORS preflight failures on Apps Script.
        const body = JSON.stringify({ action, payload });
        options.headers["Content-Type"] = "text/plain;charset=UTF-8";
        options.body = body;

        // Avoid keepalive on large payloads (e.g., payment images) because browsers
        // will reject bodies >64kb when keepalive is enabled, causing fetch to fail.
        const bodySize = body?.length || 0;
        if (bodySize <= 60_000) {
            options.keepalive = true;
        }
    }

    const isCacheableGet = method === "GET" && !bypassCache;
    const cacheKey = isCacheableGet ? buildCacheKey({ url, action, params }) : null;
    if (isCacheableGet) {
        const cached = getCachedResponse(cacheKey, cacheTtlMs);
        if (cached) return cached;
    }

    if (isCacheableGet && inflightGets.has(cacheKey)) {
        return inflightGets.get(cacheKey);
    }

    const fetchPromise = fetch(target, options).then(async (res) => {
        if (!res.ok) throw new Error(`Non-200: ${res.status}`);
        return res.json();
    });

    if (isCacheableGet) inflightGets.set(cacheKey, fetchPromise);

    try {
        const data = await fetchPromise;
        if (isCacheableGet) setCachedResponse(cacheKey, data);
        return data;
    } finally {
        if (isCacheableGet) inflightGets.delete(cacheKey);
    }
}
