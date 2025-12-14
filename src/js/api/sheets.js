/**
 * Google Sheets API Communication
 * 
 * Handles all communication with the Google Apps Script backend.
 */

import { clauseSections } from "../constants.js";
import { currentFlow } from "../state.js";
import { showToast, updateConnectionIndicator } from "../utils/ui.js";
import { callAppScript, ensureAppScriptUrl } from "./appscriptClient.js";

// Detect accidental double-loading so we can surface clearer diagnostics
if (globalThis.__sheetsApiLoaded) {
    console.warn(
        "sheets.js loaded more than once; check duplicate script imports to avoid redeclaration errors"
    );
} else {
    globalThis.__sheetsApiLoaded = true;
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

    try {
        const data = await callAppScript({ url, action: "wings" });
        if (!Array.isArray(data.wings)) return;
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");

        const sel = document.getElementById("wing");
        if (sel) {
            sel.innerHTML = '<option value="">Select wing</option>';
            data.wings.forEach((w) => {
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
    } catch (e) {
        console.warn("Could not fetch wings", e);
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Wing sync failed" : "Offline");
    }
}

/**
 * Persists a new wing value to Google Sheets and refreshes UI indicators.
 * @param {string} wing - Wing name entered by the user.
 * @returns {Promise<object>} API response shape from the Apps Script endpoint.
 */
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
        return data;
    } catch (e) {
        console.error("addWingToSheet error", e);
        showToast("Failed to save wing", "error");
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Wing save failed" : "Offline");
        return { ok: false };
    }
}

/**
 * Removes an existing wing from Google Sheets.
 * @param {string} wing - Wing identifier to delete.
 * @returns {Promise<object>} Result payload from the Apps Script endpoint.
 */
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
        return data;
    } catch (e) {
        console.error("removeWingFromSheet error", e);
        showToast("Failed to remove wing", "error");
        return { ok: false };
    }
}

/**
 * Retrieves saved landlord profiles from Google Sheets.
 * @returns {Promise<{landlords: Array}>} Fetched landlord collection (empty on failure).
 */
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

/**
 * Saves or updates a landlord configuration record.
 * @param {object} payload - Landlord details (id, name, aadhaar, address, defaults).
 * @returns {Promise<object>} Apps Script response with ok/message flags.
 */
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
        if (data?.ok) showToast("Landlord saved", "success");
        return data;
    } catch (e) {
        console.error("saveLandlordConfig error", e);
        showToast("Failed to save landlord", "error");
        return { ok: false };
    }
}

/**
 * Deletes a landlord configuration from Sheets storage.
 * @param {string} landlordId - Identifier of the landlord to remove.
 * @returns {Promise<object>} API response containing ok/message metadata.
 */
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
        if (data?.ok) showToast("Landlord removed", "success");
        return data;
    } catch (e) {
        console.error("deleteLandlordConfig error", e);
        showToast("Failed to delete landlord", "error");
        return { ok: false };
    }
}

/**
 * Loads the list of previously generated bills for quick access in the UI.
 * @returns {Promise<{bills: Array}>} Generated bill summaries or empty list on error.
 */
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

/**
 * Retrieves a saved billing record for a given month/wing combination.
 * @param {string} monthKey - Month key in YYYY-MM format.
 * @param {string} wing - Wing identifier to scope the record lookup.
 * @returns {Promise<object>} Billing payload (charges, meta, tenants) or empty object.
 */
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

/**
 * Fetches all configured units from Google Sheets for selector population.
 * @returns {Promise<{units: Array}>} Unit list payload (empty on failure).
 */
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

/**
 * Persists billing calculations to Google Sheets.
 * @param {object} payload - Billing metadata, tenant charges, and notes.
 * @returns {Promise<object>} Apps Script response reflecting save status.
 */
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

/**
 * Saves or updates an individual unit configuration.
 * @param {object} payload - Unit fields including wing, number, direction, and floor.
 * @returns {Promise<object>} Apps Script response describing persistence outcome.
 */
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
        if (data?.ok) showToast("Unit saved", "success");
        return data;
    } catch (e) {
        console.error("saveUnitConfig error", e);
        showToast("Failed to save unit", "error");
        return {};
    }
}

/**
 * Deletes an existing unit record by id.
 * @param {string} unitId - Unique identifier for the unit.
 * @returns {Promise<object>} Apps Script response noting deletion status.
 */
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
        if (data?.ok) showToast("Unit removed", "success");
        return data;
    } catch (e) {
        console.error("deleteUnitConfig error", e);
        showToast("Failed to delete unit", "error");
        return {};
    }
}

/**
 * Loads payment records from Google Sheets for display in the Payments tab.
 * @returns {Promise<{payments: Array}>} Collection of payment rows or empty array.
 */
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

/**
 * Requests a pre-signed URL or preview blob for an attachment stored remotely.
 * @param {string} attachmentUrl - URL returned from Sheets storage.
 * @returns {Promise<object>} Preview payload or empty object on failure.
 */
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

/**
 * Saves a payment entry against a tenant/unit combination.
 * @param {object} payload - Payment details including amount, date, and attachment info.
 * @returns {Promise<object>} Result payload including ok/message flags.
 */
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
        }
        return data;
    } catch (e) {
        console.error("savePaymentRecord error", e);
        showToast("Failed to save payment", "error");
        return {};
    }
}

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

    try {
        const data = await callAppScript({ url, action: "clauses" });

        clauseSections.tenant.items = Array.isArray(data.tenant) ? data.tenant : [];
        clauseSections.landlord.items = Array.isArray(data.landlord) ? data.landlord : [];
        clauseSections.penalties.items = Array.isArray(data.penalties) ? data.penalties : [];
        clauseSections.misc.items = Array.isArray(data.misc) ? data.misc : [];

        normalizeClauseSections();
        renderClausesUI();
        setClausesDirty(false);
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
        const data = await callAppScript({
            url,
            action: "saveClauses",
            method: "POST",
            payload,
        });
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
        const data = await callAppScript({
            url,
            action: "updateTenant",
            method: "POST",
            payload,
        });
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

export async function getRentRevisions(tenancyId) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return { ok: false, revisions: [] };

    try {
        const data = await callAppScript({
            url,
            action: "getRentRevisions",
            method: "POST",
            payload: { tenancyId },
        });
        return data;
    } catch (e) {
        console.error("getRentRevisions error", e);
        showToast("Failed to load rent history", "error");
        return { ok: false, revisions: [] };
    }
}

export async function saveRentRevision(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return { ok: false };

    try {
        const data = await callAppScript({
            url,
            action: "saveRentRevision",
            method: "POST",
            payload,
        });
        if (data?.ok) showToast("Rent revision saved", "success");
        return data;
    } catch (e) {
        console.error("saveRentRevision error", e);
        showToast("Failed to save rent revision", "error");
        return { ok: false };
    }
}

export async function deleteRentRevision(revisionId) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return { ok: false };

    try {
        const data = await callAppScript({
            url,
            action: "deleteRentRevision",
            method: "POST",
            payload: { revisionId },
        });
        if (data?.ok) showToast("Rent revision removed", "success");
        return data;
    } catch (e) {
        console.error("deleteRentRevision error", e);
        showToast("Failed to delete rent revision", "error");
        return { ok: false };
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
    } catch (e) {
        console.error("saveTenantToDb error", e);
        alert("Failed to call Apps Script. Check URL / deployment.");
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", navigator.onLine ? "Save failed" : "Offline");
    }
}
