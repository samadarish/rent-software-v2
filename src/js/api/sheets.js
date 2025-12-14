/**
 * Google Sheets API Communication
 * 
 * Handles all communication with the Google Apps Script backend.
 */

import { clauseSections } from "../constants.js";
import { currentFlow } from "../state.js";
import { showToast, updateConnectionIndicator } from "../utils/ui.js";
import { debouncePromise } from "../utils/timing.js";
import {
    callAppScript,
    ensureAppScriptUrl,
    getCachedGetResponse,
    invalidateCachedGets,
} from "./appscriptClient.js";

// Detect accidental double-loading so we can surface clearer diagnostics
if (globalThis.__sheetsApiLoaded) {
    console.warn(
        "sheets.js loaded more than once; check duplicate script imports to avoid redeclaration errors"
    );
} else {
    globalThis.__sheetsApiLoaded = true;
}

function applyWingsToDropdown(wings = []) {
    const sel = document.getElementById("wing");
    if (sel) {
        sel.innerHTML = '<option value="">Select wing</option>';
        wings.forEach((w) => {
            const opt = document.createElement("option");
            opt.value = w;
            opt.textContent = w;
            sel.appendChild(opt);
        });

        const billingSel = document.getElementById("billingWingSelect");
        if (billingSel) {
            billingSel.innerHTML = sel.innerHTML;
        }
        document.dispatchEvent(new CustomEvent("wings:updated"));
    }
}

function hydrateWingsFromCache(url) {
    const cached = getCachedGetResponse({ url, action: "wings" });
    if (cached?.wings?.length) {
        applyWingsToDropdown(cached.wings);
        return true;
    }
    return false;
}

/**
 * Fetches available wings from Google Sheets and populates the dropdown
 */
export async function fetchWingsFromSheet() {
    const url = ensureAppScriptUrl({
        onMissing: () =>
            updateConnectionIndicator(
                navigator.onLine ? "online" : "offline",
                "Set Apps Script URL"
            ),
    });
    if (!url) return;

    hydrateWingsFromCache(url);

    try {
        const data = await callAppScript({ url, action: "wings" });
        if (!Array.isArray(data.wings)) return;
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");

        applyWingsToDropdown(data.wings);
    } catch (e) {
        console.warn("Could not fetch wings", e);
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Wing sync failed" : "Offline");
    }
}

export async function addWingToSheet(wing) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return { ok: false };

    const cleaned = (wing || "").trim();
    if (!cleaned) {
        showToast("Please enter a wing name", "warning");
        return { ok: false };
    }

    try {
        const data = await callAppScript({
            url,
            action: "addWing",
            method: "POST",
            payload: { wing: cleaned },
        });
        const msg = (data && data.message) || "Wing saved";
        showToast(msg, "success");
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
        invalidateCachedGets({ action: "wings" });
        return data;
    } catch (e) {
        console.error("addWingToSheet error", e);
        showToast("Failed to save wing", "error");
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Wing save failed" : "Offline");
        return { ok: false };
    }
}

export async function removeWingFromSheet(wing) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return { ok: false };

    const cleaned = (wing || "").trim();
    if (!cleaned) {
        showToast("Please enter a wing to remove", "warning");
        return { ok: false };
    }

    try {
        const data = await callAppScript({
            url,
            action: "removeWing",
            method: "POST",
            payload: { wing: cleaned },
        });
        showToast("Wing removed", "success");
        invalidateCachedGets({ action: "wings" });
        return data;
    } catch (e) {
        console.error("removeWingFromSheet error", e);
        showToast("Failed to remove wing", "error");
        return { ok: false };
    }
}

export async function fetchLandlordsFromSheet() {
    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view landlords", "warning"),
    });
    if (!url) return { landlords: [] };

    try {
        const data = await callAppScript({ url, action: "landlords" });
        if (Array.isArray(data.landlords)) {
            document.dispatchEvent(new CustomEvent("landlords:updated", { detail: data.landlords }));
        }
        return data;
    } catch (e) {
        console.error("fetchLandlordsFromSheet error", e);
        showToast("Could not fetch landlords", "error");
        return { landlords: [] };
    }
}

export async function saveLandlordConfig(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    if (!url) return { ok: false };

    try {
        const data = await callAppScript({
            url,
            action: "saveLandlord",
            method: "POST",
            payload,
        });
        if (data?.ok) {
            showToast("Landlord saved", "success");
            invalidateCachedGets({ action: "landlords" });
        }
        return data;
    } catch (e) {
        console.error("saveLandlordConfig error", e);
        showToast("Failed to save landlord", "error");
        return { ok: false };
    }
}

export async function deleteLandlordConfig(landlordId) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    if (!url) return { ok: false };

    try {
        const data = await callAppScript({
            url,
            action: "deleteLandlord",
            method: "POST",
            payload: { landlordId },
        });
        if (data?.ok) {
            showToast("Landlord removed", "success");
            invalidateCachedGets({ action: "landlords" });
        }
        return data;
    } catch (e) {
        console.error("deleteLandlordConfig error", e);
        showToast("Failed to delete landlord", "error");
        return { ok: false };
    }
}

export async function fetchGeneratedBills() {
    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view generated bills", "warning"),
    });
    if (!url) return { bills: [] };

    try {
        return await callAppScript({ url, action: "generatedbills" });
    } catch (e) {
        console.error("fetchGeneratedBills error", e);
        showToast("Could not fetch generated bills", "error");
        return { bills: [] };
    }
}

export async function fetchBillingRecord(monthKey, wing) {
    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view saved billing", "warning"),
    });
    if (!url) return {};
    if (!monthKey || !wing) return {};

    try {
        return await callAppScript({ url, action: "getbillingrecord", params: { month: monthKey, wing } });
    } catch (e) {
        console.error("fetchBillingRecord error", e);
        showToast("Could not load saved billing", "error");
        return {};
    }
}

export async function fetchUnitsFromSheet() {
    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view units", "warning"),
    });
    if (!url) return { units: [] };

    try {
        return await callAppScript({ url, action: "units" });
    } catch (e) {
        console.error("fetchUnitsFromSheet error", e);
        showToast("Could not fetch units", "error");
        return { units: [] };
    }
}

export async function saveBillingRecord(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    if (!url) return {};

    try {
        const data = await callAppScript({
            url,
            action: "saveBillingRecord",
            method: "POST",
            payload,
        });
        if (data?.ok) {
            showToast("Billing saved", "success");
        }
        return data;
    } catch (e) {
        console.error("saveBillingRecord error", e);
        showToast("Failed to save billing to Google Sheets", "error");
        return {};
    }
}

export async function saveUnitConfig(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    if (!url) return {};

    try {
        const data = await callAppScript({
            url,
            action: "saveUnit",
            method: "POST",
            payload,
        });
        if (data?.ok) {
            showToast("Unit saved", "success");
            invalidateCachedGets({ action: "units" });
        }
        return data;
    } catch (e) {
        console.error("saveUnitConfig error", e);
        showToast("Failed to save unit", "error");
        return {};
    }
}

export async function deleteUnitConfig(unitId) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    if (!url) return {};

    try {
        const data = await callAppScript({
            url,
            action: "deleteUnit",
            method: "POST",
            payload: { unitId },
        });
        if (data?.ok) {
            showToast("Unit removed", "success");
            invalidateCachedGets({ action: "units" });
        }
        return data;
    } catch (e) {
        console.error("deleteUnitConfig error", e);
        showToast("Failed to delete unit", "error");
        return {};
    }
}

export async function fetchPayments() {
    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view payments", "warning"),
    });
    if (!url) return { payments: [] };

    try {
        return await callAppScript({ url, action: "payments" });
    } catch (e) {
        console.error("fetchPayments error", e);
        showToast("Could not fetch payments", "error");
        return { payments: [] };
    }
}

export async function fetchAttachmentPreview(attachmentUrl) {
    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view attachments", "warning"),
    });

    if (!url || !attachmentUrl) return {};

    try {
        return await callAppScript({
            url,
            action: "attachmentpreview",
            params: { attachmentUrl },
        });
    } catch (e) {
        console.error("fetchAttachmentPreview error", e);
        showToast("Could not load attachment preview", "error");
        return {};
    }
}

export async function savePaymentRecord(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    if (!url) return {};

    try {
        const data = await callAppScript({
            url,
            action: "savePayment",
            method: "POST",
            payload,
        });
        if (data?.ok) {
            showToast("Payment saved", "success");
            invalidateCachedGets({ action: "payments" });
        }
        return data;
    } catch (e) {
        console.error("savePaymentRecord error", e);
        showToast("Failed to save payment", "error");
        return {};
    }
}

function applyClausesPayload(data) {
    clauseSections.tenant.items = Array.isArray(data?.tenant) ? data.tenant : [];
    clauseSections.landlord.items = Array.isArray(data?.landlord) ? data.landlord : [];
    clauseSections.penalties.items = Array.isArray(data?.penalties) ? data.penalties : [];
    clauseSections.misc.items = Array.isArray(data?.misc) ? data.misc : [];

    normalizeClauseSections();
    renderClausesUI();
    setClausesDirty(false);
}

function hydrateClausesFromCache(url) {
    const cached = getCachedGetResponse({ url, action: "clauses" });
    if (!cached) return false;
    applyClausesPayload(cached);
    updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Cached clauses");
    return true;
}

const saveClausesDebounced = debouncePromise(async ({ url, payload }) => {
    const data = await callAppScript({
        url,
        action: "saveClauses",
        method: "POST",
        payload,
        bypassCache: true,
    });
    invalidateCachedGets({ action: "clauses" });
    return data;
}, 400);

/**
 * Loads clauses from Google Sheets and updates the UI
 * @param {boolean} showNotification - Whether to show a toast notification
 */
export async function loadClausesFromSheet(showNotification = false) {
    const url = ensureAppScriptUrl({
        onMissing: () =>
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL"),
    });

    // Import the clauses module functions
    const { normalizeClauseSections, renderClausesUI, setClausesDirty } =
        await import("../features/agreements/clauses.js");

    if (!url) {
        normalizeClauseSections();
        renderClausesUI();
        setClausesDirty(false);
        return;
    }

    hydrateClausesFromCache(url);

    try {
        const data = await callAppScript({ url, action: "clauses" });
        applyClausesPayload(data);
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");

        if (showNotification) {
            showToast("Latest clauses loaded from Google Sheets", "info");
        }
    } catch (e) {
        console.warn("Could not fetch clauses; leaving current UI", e);
        normalizeClauseSections();
        renderClausesUI();
        setClausesDirty(false);
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Clause sync failed" : "Offline");
        if (showNotification) {
            showToast("Failed to load clauses from Google Sheets", "error");
        }
    }
}

/**
 * Saves clauses to Google Sheets
 */
export async function saveClausesToSheet() {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    if (!url) return;

    // Import the clauses module functions
    const { normalizeClauseSections, setClausesDirty } =
        await import("../features/agreements/clauses.js");

    normalizeClauseSections();

    const payload = {
        tenant: clauseSections.tenant.items,
        landlord: clauseSections.landlord.items,
        penalties: clauseSections.penalties.items,
        misc: clauseSections.misc.items,
    };

    try {
        const data = await saveClausesDebounced({ url, payload });
        showToast(
            (data && data.message) || "Clauses saved to Google Sheets",
            "success"
        );
        setClausesDirty(false);
    } catch (e) {
        console.error("saveClausesToSheet error", e);
        showToast("Failed to save clauses to Google Sheets", "error");
    }
}

/**
 * Fetches tenant directory (tenants + family) from Google Sheets
 * @returns {Promise<{ tenants: Array }>} API response with tenants array
 */
export async function fetchTenantDirectory() {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            showToast("Please configure the Apps Script URL to view tenants", "warning");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return { tenants: [] };

    try {
        const data = await callAppScript({ url, action: "tenants" });
        if (!data || !Array.isArray(data.tenants)) {
            throw new Error("Invalid tenant response");
        }
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
        return data;
    } catch (e) {
        console.error("fetchTenantDirectory error", e);
        showToast("Failed to load tenants from Google Sheets", "error");
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Tenant fetch failed" : "Offline");
        return { tenants: [] };
    }
}

/**
 * Updates an existing tenant + family entries in Google Sheets
 * @param {object} payload - update payload (tenantId, tenancyId, grn, updates, familyMembers)
 */
const updateTenantRecordDebounced = debouncePromise(async ({ url, payload }) => {
    const data = await callAppScript({
        url,
        action: "updateTenant",
        method: "POST",
        payload,
        bypassCache: true,
    });
    invalidateCachedGets({ action: "tenants" });
    invalidateCachedGets({ action: "units" });
    return data;
}, 400);

export async function updateTenantRecord(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return;

    try {
        const data = await updateTenantRecordDebounced({ url, payload });
        showToast((data && data.message) || "Tenant updated", "success");
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
        return data;
    } catch (e) {
        console.error("updateTenantRecord error", e);
        showToast("Failed to update tenant", "error");
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Update failed" : "Offline");
        throw e;
    }
}

/**
 * Saves tenant data to Google Sheets database
 */
export async function saveTenantToDb() {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return;

    // Import the form module to collect data
    const { collectFullPayloadForDb } = await import("../features/tenants/form.js");
    const payload = collectFullPayloadForDb();
    const action = "saveTenant";

    try {
        const data = await callAppScript({
            url,
            action,
            method: "POST",
            payload,
        });

        let msg = "Tenant data sent to Google Sheets.";
        if (data && data.message) msg = data.message;
        alert(msg);
        const { refreshUnitOptions } = await import("../features/tenants/form.js");
        refreshUnitOptions(true);
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
        invalidateCachedGets({ action: "tenants" });
        invalidateCachedGets({ action: "units" });
    } catch (e) {
        console.error("saveTenantToDb error", e);
        alert("Failed to call Apps Script. Check URL / deployment.");
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Save failed" : "Offline");
    }
}
