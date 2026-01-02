/**
 * Configuration Management
 * 
 * Handles App Script URL configuration and storage.
 */

import { STORAGE_KEYS } from "../constants.js";
import { hideModal, showModal, showToast, updateConnectionIndicator } from "../utils/ui.js";

const APPS_SCRIPT_PATH_REGEX = /^\/macros\/s\/[A-Za-z0-9_-]+\/(exec|dev)(\/)?$/;

/**
 * Retrieves the Google Apps Script URL from local storage
 * @returns {string} The stored URL or empty string if not configured
 */
export function getAppScriptUrl() {
    return (localStorage.getItem(STORAGE_KEYS.APP_SCRIPT_URL) || "").trim();
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
    const modal = document.getElementById("appscriptModal");
    if (modal) hideModal(modal);

    updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
    document.dispatchEvent(
        new CustomEvent("appscript:url-updated", { detail: { url: normalizedUrl } })
    );

    // Refresh data after URL is saved
    import("./sheets.js").then(({ fetchWingsFromSheet, loadClausesFromSheet }) => {
        fetchWingsFromSheet(true);
        loadClausesFromSheet(true, true);
    });
}

/**
 * Ensures the App Script URL is configured
 * Opens the configuration modal if URL is not set
 */
export function ensureAppScriptConfigured() {
    const url = getAppScriptUrl();
    const input = document.getElementById("appscript_url");
    if (input && url) input.value = url;

    if (!url) {
        const modal = document.getElementById("appscriptModal");
        if (modal) showModal(modal);
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
    }
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
    const select = document.getElementById("landlordExistingSelect");
    if (select) select.value = "";

    showModal(modal);
}

/**
 * Persists landlord defaults to local storage and reapplies them to the form.
 */
export function saveLandlordDefaults() {
    const name = document.getElementById("landlordDefaultName")?.value.trim() || "";
    const aadhaar = document.getElementById("landlordDefaultAadhaar")?.value.trim() || "";
    const address = document.getElementById("landlordDefaultAddress")?.value.trim() || "";

    localStorage.setItem(
        STORAGE_KEYS.LANDLORD_DEFAULTS,
        JSON.stringify({ name, aadhaar, address })
    );

    applyLandlordDefaultsToForm(true);

    const modal = document.getElementById("landlordConfigModal");
    if (modal) hideModal(modal);

    showToast("Landlord defaults saved", "success");
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

    const { addWingToSheet, fetchWingsFromSheet } = await import("./sheets.js");
    const result = await addWingToSheet(wing);
    if (result && result.ok !== false) {
        wingInput.value = "";
        fetchWingsFromSheet(true);
    }
}
