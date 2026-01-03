/**
 * Shared Apps Script client helpers.
 * Centralizes URL resolution and JSON fetch/post handling so that
 * feature-specific API wrappers can stay concise.
 */
import { getAppScriptUrl, openAppScriptModal } from "./config.js";
import { cacheGet, cacheSet } from "./localDb.js";

const inflightGets = new Map();
const timingBuffer = globalThis.__appscriptTimings || [];
globalThis.__appscriptTimings = timingBuffer;
const TIMING_LIMIT = 50;
const LOCAL_CACHE_TTL_MS = 5 * 60 * 1000;

function recordTiming({ action, method, url, ok, durationMs }) {
    try {
        timingBuffer.unshift({
            action,
            method,
            url,
            ok,
            durationMs,
            at: new Date().toISOString(),
        });
        if (timingBuffer.length > TIMING_LIMIT) {
            timingBuffer.length = TIMING_LIMIT;
        }
        const label = ok ? "ok" : "fail";
        console.debug(`[AppsScript] ${method} ${action} ${label} in ${durationMs.toFixed(1)}ms`);
    } catch (err) {
        console.warn("Unable to record Apps Script timing", err);
    }
}

/**
 * Builds a deterministic cache key for GET calls so repeated requests can be deduped.
 * @param {{ url: string, action: string, params?: Record<string,string> }} config
 * @returns {string} cache identifier combining URL, action, and sorted params
 */
function buildCacheKey({ url, action, params }) {
    const sortedEntries = Object.entries(params || {})
        .filter(([, value]) => value !== undefined && value !== null)
        .sort(([a], [b]) => a.localeCompare(b));
    const paramString = sortedEntries
        .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
        .join("&");
    return `${url}|${action}|${paramString}`;
}

async function readLocalCache(cacheKey, ttlMs, allowStale) {
    if (!cacheKey) return null;
    const entry = await cacheGet(cacheKey);
    if (!entry || typeof entry !== "object") return null;
    const updatedAt = Number(entry.updated_at || entry.updatedAt || 0);
    if (ttlMs && updatedAt && Date.now() - updatedAt > ttlMs && !allowStale) return null;
    return entry.value ?? null;
}

/**
 * Opens the Apps Script URL configuration modal when the user needs to set it up.
 */
function showConfigModal() {
    openAppScriptModal({ mode: "input" });
}

/**
 * Returns the configured Apps Script URL, optionally triggering callbacks when missing.
 * @param {{ onMissing?: Function, promptForConfig?: boolean }} options
 * @returns {string} Stored Apps Script URL or empty string.
 */
export function ensureAppScriptUrl({ onMissing, promptForConfig = false } = {}) {
    const url = getAppScriptUrl();
    if (!url) {
        if (typeof onMissing === "function") onMissing();
        if (promptForConfig) showConfigModal();
    }
    return url;
}

/**
 * Executes a JSON-based call to the Apps Script backend.
 * Dedupes GET requests, handles payload encoding, and returns parsed JSON.
 * @param {{ url: string, action: string, method?: string, params?: object, payload?: any }} options
 * @returns {Promise<any>} Parsed JSON response.
 */
export async function callAppScript({
    url,
    action,
    method = "GET",
    params = {},
    payload,
    cache = {},
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

    const isCacheableGet = method === "GET";
    const cacheKey = isCacheableGet ? buildCacheKey({ url, action, params }) : null;
    const cacheOptions = cache && typeof cache === "object" ? cache : {};
    const allowLocalCache = cacheOptions.useLocal !== false;
    const writeLocalCache = cacheOptions.write !== false;
    const ttlMs =
        typeof cacheOptions.ttlMs === "number" ? cacheOptions.ttlMs : LOCAL_CACHE_TTL_MS;
    const revalidate = cacheOptions.revalidate !== false;
    const allowStale = cacheOptions.allowStale === true || !navigator.onLine;

    const runFetch = () => {
        const start = (performance && performance.now && performance.now()) || Date.now();
        return fetch(target, options)
            .then(async (res) => {
                if (!res.ok) throw new Error(`Non-200: ${res.status}`);
                return res.json();
            })
            .then((data) => {
                const durationMs =
                    ((performance && performance.now && performance.now()) || Date.now()) - start;
                recordTiming({ action, method, url: target, ok: true, durationMs });
                if (isCacheableGet && cacheKey && writeLocalCache) {
                    cacheSet(cacheKey, data);
                }
                return data;
            })
            .catch((err) => {
                const durationMs =
                    ((performance && performance.now && performance.now()) || Date.now()) - start;
                recordTiming({ action, method, url: target, ok: false, durationMs });
                throw err;
            });
    };

    if (isCacheableGet && allowLocalCache && cacheKey) {
        const cached = await readLocalCache(cacheKey, ttlMs, allowStale);
        if (cached) {
            if (revalidate && !inflightGets.has(cacheKey)) {
                const refreshPromise = runFetch().finally(() => inflightGets.delete(cacheKey));
                inflightGets.set(cacheKey, refreshPromise);
                refreshPromise.catch(() => null);
            }
            return cached;
        }
    }

    if (isCacheableGet && inflightGets.has(cacheKey)) {
        return inflightGets.get(cacheKey);
    }

    const fetchPromise = runFetch();
    if (isCacheableGet && cacheKey) inflightGets.set(cacheKey, fetchPromise);

    try {
        return await fetchPromise;
    } finally {
        if (isCacheableGet && cacheKey) inflightGets.delete(cacheKey);
    }
}
