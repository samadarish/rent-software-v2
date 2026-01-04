/**
 * Configuration Management
 * 
 * Handles App Script URL configuration and storage.
 */

import { STORAGE_KEYS } from "../constants.js";
import { hideModal, showModal, showToast, updateConnectionIndicator } from "../utils/ui.js";
import { startInitialSync } from "./syncManager.js";

const APPS_SCRIPT_PATH_REGEX = /^\/macros\/s\/[A-Za-z0-9_-]+\/(exec|dev)(\/)?$/;
const FULL_SYNC_STALE_MS = 2 * 60 * 60 * 1000;
let modalProgressBound = false;

/**
 * Retrieves the Google Apps Script URL from local storage
 * @returns {string} The stored URL or empty string if not configured
 */
export function getAppScriptUrl() {
    return (localStorage.getItem(STORAGE_KEYS.APP_SCRIPT_URL) || "").trim();
}

function getLastFullSyncAt() {
    const raw = localStorage.getItem(STORAGE_KEYS.LAST_FULL_SYNC_AT);
    const parsed = Number(raw || 0);
    return Number.isNaN(parsed) ? 0 : parsed;
}

function isFullSyncStale() {
    const last = getLastFullSyncAt();
    if (!last) return true;
    return Date.now() - last > FULL_SYNC_STALE_MS;
}

function setAppScriptModalMode(mode = "input") {
    const modal = document.getElementById("appscriptModal");
    if (!modal) return;
    const form = document.getElementById("appscriptFormSection");
    const sync = document.getElementById("appscriptSyncSection");
    const closeBtn = document.getElementById("appscriptModalClose");
    const isSync = mode === "syncing";
    modal.dataset.syncState = isSync ? "syncing" : "idle";
    if (form) form.classList.toggle("hidden", isSync);
    if (sync) sync.classList.toggle("hidden", !isSync);
    if (closeBtn) closeBtn.classList.toggle("hidden", isSync);
}

function updateAppScriptSyncProgress(percent, label) {
    const bar = document.getElementById("appscriptSyncBar");
    const percentLabel = document.getElementById("appscriptSyncPercent");
    const textLabel = document.getElementById("appscriptSyncLabel");
    if (bar) bar.style.width = `${Math.max(0, Math.min(100, percent || 0))}%`;
    if (percentLabel) percentLabel.textContent = `${Math.round(percent || 0)}%`;
    if (textLabel && label) textLabel.textContent = label;
}

function bindModalProgress() {
    if (modalProgressBound || typeof document === "undefined") return;
    modalProgressBound = true;
    document.addEventListener("sync:progress", (event) => {
        const modal = document.getElementById("appscriptModal");
        if (!modal || modal.dataset.syncState !== "syncing") return;
        const detail = event?.detail || {};
        updateAppScriptSyncProgress(detail.percent || 0, detail.label || "");
    });
}

export function openAppScriptModal({ mode = "input" } = {}) {
    const modal = document.getElementById("appscriptModal");
    if (!modal) return;
    if (mode !== "syncing") {
        const input = document.getElementById("appscript_url");
        if (input) input.value = getAppScriptUrl();
    }
    setAppScriptModalMode(mode === "syncing" ? "syncing" : "input");
    showModal(modal);
}

async function runFullSyncWithModal({ reason = "manual" } = {}) {
    const modalVariant = reason === "stale" ? "simple" : "config";
    const modal =
        modalVariant === "simple"
            ? document.getElementById("autoSyncModal")
            : document.getElementById("appscriptModal");
    if (!modal) return { ok: false, reason: "missing-modal" };
    bindModalProgress();
    if (modalVariant === "config") {
        setAppScriptModalMode("syncing");
        showModal(modal);
        updateAppScriptSyncProgress(0, "Preparing sync...");
    } else {
        modal.dataset.syncState = "syncing";
        showModal(modal);
    }
    const result = await startInitialSync();
    if (modalVariant === "config") {
        hideModal(modal);
        setAppScriptModalMode("input");
    } else {
        modal.dataset.syncState = "idle";
        hideModal(modal);
    }
    if (result?.ok === false) {
        showToast("Sync finished with errors. Check your connection.", "warning");
    }
    return result;
}

/**
 * Validates and normalizes an Apps Script web app URL.
 * @param {string} rawUrl - Raw URL text from the settings field.
 * @returns {URL | null} Parsed URL object when valid; otherwise null.
 */
function parseAppScriptUrl(rawUrl) {
    try {
        const parsed = new URL(rawUrl);
        if (parsed.protocol !== "https:") return null;
        if (parsed.hostname !== "script.google.com") return null;
        if (!APPS_SCRIPT_PATH_REGEX.test(parsed.pathname)) return null;
        return parsed;
    } catch (error) {
        console.warn("Invalid Apps Script URL", error);
        return null;
    }
}

/**
 * Saves the Google Apps Script URL to local storage and refreshes data
 * Shows modal if URL is not provided
 */
export function saveAppScriptUrl() {
    const input = document.getElementById("appscript_url");
    if (!input) return;

    const raw = input.value.trim();
    if (!raw) {
        showToast("Please enter an Apps Script Web App URL.", "error");
        return;
    }

    const parsed = parseAppScriptUrl(raw);
    if (!parsed) {
        showToast(
            "Enter a https://script.google.com/macros/s/<deployment>/(exec|dev) URL.",
            "error"
        );
        return;
    }

    const normalizedUrl = `${parsed.origin}${parsed.pathname}`;

    localStorage.setItem(STORAGE_KEYS.APP_SCRIPT_URL, normalizedUrl);

    updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
    document.dispatchEvent(
        new CustomEvent("appscript:url-updated", { detail: { url: normalizedUrl } })
    );

    runFullSyncWithModal({ reason: "manual" });
}

/**
 * Ensures the App Script URL is configured
 * Opens the configuration modal if URL is not set
 */
export async function ensureAppScriptConfigured({ autoSync = false } = {}) {
    bindModalProgress();
    const url = getAppScriptUrl();
    const input = document.getElementById("appscript_url");
    if (input && url) input.value = url;

    if (!url) {
        openAppScriptModal({ mode: "input" });
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        return { ok: false, reason: "missing-url" };
    }

    if (autoSync && navigator.onLine && isFullSyncStale()) {
        await runFullSyncWithModal({ reason: "stale" });
    }

    return { ok: true };
}

/**
 * Reads landlord defaults from local storage and returns an object shape.
 * @returns {{ name?: string, aadhaar?: string, address?: string }} persisted defaults
 */
function getLandlordDefaults() {
    try {
        const raw = localStorage.getItem(STORAGE_KEYS.LANDLORD_DEFAULTS);
        return raw ? JSON.parse(raw) : {};
    } catch (e) {
        console.warn("Unable to parse landlord defaults", e);
        return {};
    }
}

/**
 * Prefills landlord fields in the agreement form based on saved defaults.
 * @param {boolean} force - When true, overwrite existing values.
 */
export function applyLandlordDefaultsToForm(force = false) {
    const defaults = getLandlordDefaults();
    if (!defaults) return;

    const setValue = (id, value) => {
        const el = document.getElementById(id);
        if (!el || !value) return;
        if (force || !el.value) {
            el.value = value;
        }
    };

    setValue("Landlord_name", defaults.name || "");
    setValue("landlord_aadhar", defaults.aadhaar || "");
    setValue("landlord_address", defaults.address || "");
}

/**
 * Opens the landlord configuration modal and hydrates inputs with stored defaults.
 */
export function openLandlordConfigModal() {
    const modal = document.getElementById("landlordConfigModal");
    if (!modal) return;

    const defaults = getLandlordDefaults();
    const name = document.getElementById("landlordDefaultName");
    const aadhaar = document.getElementById("landlordDefaultAadhaar");
    const address = document.getElementById("landlordDefaultAddress");
    const wing = document.getElementById("landlordDefaultWing");

    if (name) name.value = defaults.name || "";
    if (aadhaar) aadhaar.value = defaults.aadhaar || "";
    if (address) address.value = defaults.address || "";
    if (wing) wing.value = "";

    showModal(modal);
}

/**
 * Persists landlord defaults to local storage and reapplies them to the form.
 */
export function saveLandlordDefaults(options = {}) {
    const name = document.getElementById("landlordDefaultName")?.value.trim() || "";
    const aadhaar = document.getElementById("landlordDefaultAadhaar")?.value.trim() || "";
    const address = document.getElementById("landlordDefaultAddress")?.value.trim() || "";

    const { closeModal = true, showMessage = true } = options || {};

    localStorage.setItem(
        STORAGE_KEYS.LANDLORD_DEFAULTS,
        JSON.stringify({ name, aadhaar, address })
    );

    applyLandlordDefaultsToForm(true);

    if (closeModal) {
        const modal = document.getElementById("landlordConfigModal");
        if (modal) hideModal(modal);
    }

    if (showMessage) {
        showToast("Landlord defaults saved", "success");
    }
}

/**
 * Adds a wing from within the landlord config modal and refreshes the wing list.
 */
export async function saveWingFromLandlordConfig() {
    const wingInput = document.getElementById("landlordDefaultWing");
    if (!wingInput) return;

    const wing = (wingInput.value || "").trim();
    if (!wing) {
        showToast("Please enter a wing name", "warning");
        return;
    }

    const { addWingToSheet } = await import("./sheets.js");
    const result = await addWingToSheet(wing);
    if (result && result.ok !== false) {
        wingInput.value = "";
        showToast("Wing saved", "success");
    }
}
