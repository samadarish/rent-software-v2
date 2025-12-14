/**
 * Shared Apps Script client helpers.
 * Centralizes URL resolution and JSON fetch/post handling so that
 * feature-specific API wrappers can stay concise.
 */
import { getAppScriptUrl } from "./config.js";
import { showModal } from "../utils/ui.js";

const inflightGets = new Map();

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

/**
 * Opens the Apps Script URL configuration modal when the user needs to set it up.
 */
function showConfigModal() {
    const modal = document.getElementById("appscriptModal");
    if (modal) showModal(modal);
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
export async function callAppScript({ url, action, method = "GET", params = {}, payload }) {
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
    if (isCacheableGet && inflightGets.has(cacheKey)) {
        return inflightGets.get(cacheKey);
    }

    const fetchPromise = fetch(target, options).then(async (res) => {
        if (!res.ok) throw new Error(`Non-200: ${res.status}`);
        return res.json();
    });

    if (isCacheableGet) inflightGets.set(cacheKey, fetchPromise);

    try {
        return await fetchPromise;
    } finally {
        if (isCacheableGet) inflightGets.delete(cacheKey);
    }
}
