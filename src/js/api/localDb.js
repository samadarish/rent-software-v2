const tauriInvoke = window.__TAURI__?.core?.invoke;
const hasTauri = typeof tauriInvoke === "function";

async function safeInvoke(command, args) {
    if (!hasTauri) return null;
    try {
        return await tauriInvoke(command, args);
    } catch (err) {
        console.warn(`Local DB command failed: ${command}`, err);
        return null;
    }
}

export function isLocalDbAvailable() {
    return hasTauri;
}

export async function cacheGet(key) {
    if (!key) return null;
    return await safeInvoke("cache_get", { key });
}

export async function cacheSet(key, value) {
    if (!key) return false;
    const res = await safeInvoke("cache_set", { key, value });
    return res !== null;
}

export async function cacheDelete(key) {
    if (!key) return false;
    const res = await safeInvoke("cache_delete", { key });
    return res !== null;
}

export async function cacheDeletePrefix(prefix) {
    if (prefix === undefined || prefix === null) return false;
    const res = await safeInvoke("cache_delete_prefix", { prefix });
    return res !== null;
}

export async function queueAdd({ action, payload, method = "POST", params = {} } = {}) {
    if (!action) return null;
    return await safeInvoke("queue_add", { action, payload, method, params });
}

export async function queueList(limit) {
    return (await safeInvoke("queue_list", { limit })) || [];
}

export async function queueDelete(id) {
    if (!id && id !== 0) return false;
    const res = await safeInvoke("queue_delete", { id });
    return res !== null;
}

export async function queueClear() {
    const res = await safeInvoke("queue_clear", {});
    return res !== null;
}

export async function queueCount() {
    const res = await safeInvoke("queue_count", {});
    return typeof res === "number" ? res : 0;
}
