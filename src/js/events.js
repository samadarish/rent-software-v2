/**
 * Event Handlers Module
 * 
 * Attaches all event listeners for the application.
 * This keeps event handling centralized and organized.
 */

import { syncTenantToFamilyTable, createFamilyRow } from "./features/tenants/family.js";
import { switchFlow } from "./features/navigation/flow.js";
import {
    saveAppScriptUrl,
    getAppScriptUrl,
    openLandlordConfigModal,
    saveLandlordDefaults,
    saveWingFromLandlordConfig,
} from "./api/config.js";
import {
    saveClausesToSheet,
    loadClausesFromSheet,
    saveTenantToDb,
    saveUnitConfig,
    deleteUnitConfig,
    removeWingFromSheet,
    fetchWingsFromSheet,
    fetchLandlordsFromSheet,
    saveLandlordConfig,
    deleteLandlordConfig,
} from "./api/sheets.js";
import { exportDocxFromTemplate } from "./features/agreements/docx.js";
import { numberToIndianWords } from "./utils/formatters.js";
import { clearAllDrafts, promptAndSaveDraft } from "./features/shared/drafts.js";
import { hideModal, showModal } from "./utils/ui.js";

let unitConfigCache = [];
let landlordConfigCache = [];

/**
 * Copies the wing dropdown options into the unit configuration modal selectors.
 */
function syncUnitConfigWingOptions() {
    const wingSource = document.getElementById("wing");
    const unitWing = document.getElementById("unitConfigWing");
    if (!wingSource || !unitWing) return;
    unitWing.innerHTML = wingSource.innerHTML;
}

/**
 * Populates the unit configuration modal list and dropdown from the provided units array.
 * @param {Array} units - Units fetched from Google Sheets.
 */
function syncUnitConfigList(units) {
    unitConfigCache = Array.isArray(units) ? units : [];
    const select = document.getElementById("unitConfigExisting");
    if (!select) return;
    const previous = select.value;
    select.innerHTML = '<option value="">Select existing unit</option>';
    unitConfigCache.forEach((u) => {
        const opt = document.createElement("option");
        opt.value = u.unit_id;
        opt.textContent = [u.wing, u.unit_number].filter(Boolean).join(" - ") || u.unit_id;
        select.appendChild(opt);
    });
    if (previous && Array.from(select.options).some((o) => o.value === previous)) select.value = previous;

    const list = document.getElementById("unitConfigList");
    if (list) {
        list.innerHTML = "";
        unitConfigCache.forEach((u) => {
            const row = document.createElement("div");
            row.className = "flex items-center justify-between border border-slate-200 rounded px-2 py-1 bg-white";
            const label = document.createElement("div");
            label.textContent = `${[u.wing, u.unit_number].filter(Boolean).join(" - ") || u.unit_id} • ${u.floor || ""} ${
                u.direction || ""
            }`;
            const actions = document.createElement("div");
            actions.className = "flex gap-2";
            const loadBtn = document.createElement("button");
            loadBtn.textContent = "Load";
            loadBtn.className = "text-[10px] px-2 py-0.5 rounded border";
            loadBtn.addEventListener("click", () => {
                select.value = u.unit_id;
                select.dispatchEvent(new Event("change"));
            });
            const deleteBtn = document.createElement("button");
            deleteBtn.textContent = "Delete";
            deleteBtn.className = "text-[10px] px-2 py-0.5 rounded bg-rose-100 text-rose-700";
            deleteBtn.addEventListener("click", async () => {
                await deleteUnitConfig(u.unit_id);
                const { refreshUnitOptions } = await import("./features/tenants/form.js");
                refreshUnitOptions(true);
            });
            actions.append(loadBtn, deleteBtn);
            row.append(label, actions);
            list.appendChild(row);
        });
    }
}

/**
 * Updates the landlord config dropdown with the latest saved landlords.
 * @param {Array} landlords - Landlord records retrieved from Sheets.
 */
function syncLandlordConfigList(landlords) {
    landlordConfigCache = Array.isArray(landlords) ? landlords : [];
    const select = document.getElementById("landlordExistingSelect");
    if (select) {
        const previous = select.value;
        select.innerHTML = '<option value="">Select saved landlord</option>';
        landlordConfigCache.forEach((l) => {
            const opt = document.createElement("option");
            opt.value = l.landlord_id;
            opt.textContent = l.name || l.landlord_id;
            select.appendChild(opt);
        });
        if (previous && Array.from(select.options).some((o) => o.value === previous)) select.value = previous;
    }
}

/**
 * Renders the list of wings as removable pills next to the wing selector.
 */
function syncWingList() {
    const list = document.getElementById("wingList");
    const wingSelect = document.getElementById("wing");
    if (!list || !wingSelect) return;
    const options = Array.from(wingSelect.options).filter((opt) => opt.value);
    list.innerHTML = "";
    options.forEach((opt) => {
        const pill = document.createElement("button");
        pill.textContent = opt.value;
        pill.className = "px-2 py-1 rounded-full bg-white border text-[11px] flex items-center gap-1";
        pill.addEventListener("click", async () => {
            await removeWingFromSheet(opt.value);
            fetchWingsFromSheet();
        });
        list.appendChild(pill);
    });
}

/**
 * Generates a pseudo GRN value when the user indicates they have none.
 * @returns {string} A random NoGRN value.
 */
function generateNoGrnValue() {
    const rand = Math.floor(10000 + Math.random() * 90000);
    return `NoGRN${rand}`;
}

/**
 * Toggles the GRN input between manual and auto-generated modes.
 */
function toggleNoGrnMode() {
    const input = document.getElementById("grn_number");
    const checkbox = document.getElementById("grnNoCheckbox");
    if (!input || !checkbox) return;

    const noGrnActive = checkbox.checked;

    if (noGrnActive) {
        input.dataset.noGrn = "1";
        input.dataset.prevGrn = input.value;
        input.disabled = true;
        input.value = generateNoGrnValue();
        return;
    }

    input.dataset.noGrn = "0";
    input.disabled = false;
    input.value = "";
    delete input.dataset.prevGrn;
}

/**
 * Attaches all event handlers for the application
 * Called once on DOM load
 */
export function attachEventHandlers() {
    // Keep tenant row in family table synced with tenant form inputs
    ["Tenant_Full_Name", "tenant_Aadhar", "Tenant_Permanent_Address", "Tenant_occupation"].forEach(
        (id) => {
            const el = document.getElementById(id);
            if (el) {
                el.addEventListener("input", () => {
                    syncTenantToFamilyTable();
                });
            }
        }
    );

    const noGrnCheckbox = document.getElementById("grnNoCheckbox");
    if (noGrnCheckbox) {
        noGrnCheckbox.addEventListener("change", toggleNoGrnMode);
    }

    // Navigation: Create Agreement button
    const navCreateAgreementBtn = document.getElementById("navCreateAgreementBtn");
    if (navCreateAgreementBtn) {
        navCreateAgreementBtn.addEventListener("click", () => {
            switchFlow("agreement");
        });
    }

    // Navigation: View Tenants button
    const navViewTenantsBtn = document.getElementById("navViewTenantsBtn");
    if (navViewTenantsBtn) {
        navViewTenantsBtn.addEventListener("click", () => {
            switchFlow("viewTenants");
        });
    }

    // Navigation: Generate Bill button
    const navGenerateBillBtn = document.getElementById("navGenerateBillBtn");
    if (navGenerateBillBtn) {
        navGenerateBillBtn.addEventListener("click", () => {
            switchFlow("generateBill");
        });
    }

    // Navigation: Dashboard button
    const navDashboardBtn = document.getElementById("navDashboardBtn");
    if (navDashboardBtn) {
        navDashboardBtn.addEventListener("click", () => {
            switchFlow("dashboard");
        });
    }

    // Navigation: Payments button
    const navPaymentsBtn = document.getElementById("navPaymentsBtn");
    if (navPaymentsBtn) {
        navPaymentsBtn.addEventListener("click", () => {
            switchFlow("payments");
        });
    }

    // Navigation: Create Tenant button – opens mode chooser modal
    const navCreateTenantBtn = document.getElementById("navCreateTenantBtn");
    if (navCreateTenantBtn) {
        navCreateTenantBtn.addEventListener("click", () => {
            const modal = document.getElementById("tenantModeModal");
            if (modal) showModal(modal);
        });
    }

    // Tenant mode modal buttons
    const tenantModeCancelBtn = document.getElementById("tenantModeCancelBtn");
    if (tenantModeCancelBtn) {
        tenantModeCancelBtn.addEventListener("click", () => {
            const modal = document.getElementById("tenantModeModal");
            if (modal) hideModal(modal);
        });
    }

    const tenantModeNewBtn = document.getElementById("tenantModeNewBtn");
    if (tenantModeNewBtn) {
        tenantModeNewBtn.addEventListener("click", () => {
            const modal = document.getElementById("tenantModeModal");
            if (modal) hideModal(modal);
            switchFlow("createTenantNew");
        });
    }

    const tenantModePastBtn = document.getElementById("tenantModePastBtn");
    if (tenantModePastBtn) {
        tenantModePastBtn.addEventListener("click", () => {
            const modal = document.getElementById("tenantModeModal");
            if (modal) hideModal(modal);
            switchFlow("addPastTenant");
        });
    }

    // Amount to words converters

    // Rent amount → words
    const rentAmountEl = document.getElementById("rent_amount");
    if (rentAmountEl) {
        rentAmountEl.addEventListener("input", () => {
            const val = parseInt(rentAmountEl.value, 10);
            const out = document.getElementById("rent_amount_words");
            if (!out) return;
            out.value =
                isNaN(val) || val <= 0 ? "" : numberToIndianWords(val) + " only";
        });
    }

    // Security deposit amount → words
    const secuEl = document.getElementById("secu_depo");
    if (secuEl) {
        secuEl.addEventListener("input", () => {
            const val = parseInt(secuEl.value, 10);
            const out = document.getElementById("secu_amount_words");
            if (!out) return;
            out.value =
                isNaN(val) || val <= 0 ? "" : numberToIndianWords(val) + " only";
        });
    }

    // Rent increase amount → words
    const incEl = document.getElementById("rent_inc");
    if (incEl) {
        incEl.addEventListener("input", () => {
            const out = document.getElementById("increase_amount_word");
            if (!out) return;
            const val = parseInt(incEl.value, 10);
            out.value =
                isNaN(val) || val <= 0 ? "" : numberToIndianWords(val) + " only";
        });
    }

    // Add family row button
    const addFamilyRowBtn = document.getElementById("addFamilyRowBtn");
    if (addFamilyRowBtn) {
        addFamilyRowBtn.addEventListener("click", () => {
            const tbody = document.querySelector("#familyTable tbody");
            if (!tbody) return;
            const row = createFamilyRow();
            tbody.appendChild(row);
        });
    }

    // App Script URL configuration modal
    const navConfigBtn = document.getElementById("navConfigBtn");
    if (navConfigBtn) {
        navConfigBtn.addEventListener("click", () => {
            const url = getAppScriptUrl();
            const input = document.getElementById("appscript_url");
            if (input) input.value = url || "";
            const modal = document.getElementById("appscriptModal");
            if (modal) showModal(modal);
        });
    }

    const navLandlordConfigBtn = document.getElementById("navLandlordConfigBtn");
    if (navLandlordConfigBtn) {
        navLandlordConfigBtn.addEventListener("click", openLandlordConfigModal);
    }

    const appscriptCancelBtn = document.getElementById("appscriptCancelBtn");
    if (appscriptCancelBtn) {
        appscriptCancelBtn.addEventListener("click", () => {
            const modal = document.getElementById("appscriptModal");
            if (modal) hideModal(modal);
        });
    }

    const appscriptSaveBtn = document.getElementById("appscriptSaveBtn");
    if (appscriptSaveBtn) {
        appscriptSaveBtn.addEventListener("click", saveAppScriptUrl);
    }

    const landlordDefaultsSaveBtn = document.getElementById("landlordDefaultsSaveBtn");
    if (landlordDefaultsSaveBtn) {
        landlordDefaultsSaveBtn.addEventListener("click", async () => {
            const name = document.getElementById("landlordDefaultName")?.value.trim() || "";
            const aadhaar = document.getElementById("landlordDefaultAadhaar")?.value.trim() || "";
            const address = document.getElementById("landlordDefaultAddress")?.value.trim() || "";
            const landlordId = document.getElementById("landlordExistingSelect")?.value || "";
            await saveLandlordConfig({ landlordId, name, aadhaar, address });
            saveLandlordDefaults();
            fetchLandlordsFromSheet();
        });
    }

    const landlordLoadBtn = document.getElementById("landlordLoadBtn");
    if (landlordLoadBtn) {
        landlordLoadBtn.addEventListener("click", () => {
            const selected = document.getElementById("landlordExistingSelect")?.value;
            const match = landlordConfigCache.find((l) => l.landlord_id === selected);
            if (!match) return;
            document.getElementById("landlordDefaultName").value = match.name || "";
            document.getElementById("landlordDefaultAadhaar").value = match.aadhaar || "";
            document.getElementById("landlordDefaultAddress").value = match.address || "";
        });
    }

    const landlordDeleteBtn = document.getElementById("landlordDeleteBtn");
    if (landlordDeleteBtn) {
        landlordDeleteBtn.addEventListener("click", async () => {
            const selected = document.getElementById("landlordExistingSelect")?.value;
            if (!selected) return;
            await deleteLandlordConfig(selected);
            fetchLandlordsFromSheet();
        });
    }

    const landlordAddWingBtn = document.getElementById("landlordDefaultWingBtn");
    if (landlordAddWingBtn) {
        landlordAddWingBtn.addEventListener("click", saveWingFromLandlordConfig);
    }

    const landlordRemoveWingBtn = document.getElementById("landlordDefaultWingRemoveBtn");
    if (landlordRemoveWingBtn) {
        landlordRemoveWingBtn.addEventListener("click", async () => {
            const wing = document.getElementById("landlordDefaultWing")?.value || "";
            if (!wing) return;
            await removeWingFromSheet(wing);
            fetchWingsFromSheet();
        });
    }

    const unitConfigSaveBtn = document.getElementById("unitConfigSaveBtn");
    if (unitConfigSaveBtn) {
        unitConfigSaveBtn.addEventListener("click", async () => {
            const wing = document.getElementById("unitConfigWing")?.value || "";
            const unitNumber = document.getElementById("unitConfigNumber")?.value || "";
            if (!wing || !unitNumber) return;
            await saveUnitConfig({
                unitId: document.getElementById("unitConfigExisting")?.value || "",
                wing,
                unitNumber,
                floor: document.getElementById("unitConfigFloor")?.value || "",
                direction: document.getElementById("unitConfigDirection")?.value || "",
                meterNumber: document.getElementById("unitConfigMeter")?.value || "",
                notes: document.getElementById("unitConfigNotes")?.value || "",
                isOccupied: false,
            });

            const { refreshUnitOptions } = await import("./features/tenants/form.js");
            refreshUnitOptions(true);
        });
    }

    const unitConfigLoadBtn = document.getElementById("unitConfigLoadBtn");
    if (unitConfigLoadBtn) {
        unitConfigLoadBtn.addEventListener("click", () => {
            const selected = document.getElementById("unitConfigExisting")?.value;
            if (!selected) return;
            const unit = unitConfigCache.find((u) => u.unit_id === selected);
            if (!unit) return;
            document.getElementById("unitConfigWing").value = unit.wing || "";
            document.getElementById("unitConfigNumber").value = unit.unit_number || "";
            document.getElementById("unitConfigFloor").value = unit.floor || "";
            document.getElementById("unitConfigDirection").value = unit.direction || "";
            document.getElementById("unitConfigMeter").value = unit.meter_number || "";
            document.getElementById("unitConfigNotes").value = unit.notes || "";
        });
    }

    const unitConfigDeleteBtn = document.getElementById("unitConfigDeleteBtn");
    if (unitConfigDeleteBtn) {
        unitConfigDeleteBtn.addEventListener("click", async () => {
            const selected = document.getElementById("unitConfigExisting")?.value;
            if (!selected) return;
            await deleteUnitConfig(selected);
            const { refreshUnitOptions } = await import("./features/tenants/form.js");
            refreshUnitOptions(true);
        });
    }

    const landlordDefaultsCancelBtn = document.getElementById("landlordDefaultsCancelBtn");
    if (landlordDefaultsCancelBtn) {
        landlordDefaultsCancelBtn.addEventListener("click", () => {
            const modal = document.getElementById("landlordConfigModal");
            if (modal) hideModal(modal);
        });
    }

    document.addEventListener("landlords:updated", (e) => {
        syncLandlordConfigList(e.detail || []);
    });

    document.addEventListener("wings:updated", () => {
        syncUnitConfigWingOptions();
        syncWingList();
    });

    // Clause management buttons
    const saveClausesBtn = document.getElementById("saveClausesBtn");
    if (saveClausesBtn) {
        saveClausesBtn.addEventListener("click", saveClausesToSheet);
    }

    const reloadClausesBtn = document.getElementById("reloadClausesBtn");
    if (reloadClausesBtn) {
        reloadClausesBtn.addEventListener("click", () => {
            loadClausesFromSheet(true);
        });
    }

    syncUnitConfigWingOptions();
    syncUnitConfigList(unitConfigCache);
    syncWingList();

    // Action buttons (class-based selectors for multiple instances)

    // Save tenant (Agreement mode)
    document.querySelectorAll(".btn-save-agreement").forEach(btn => {
        btn.addEventListener("click", saveTenantToDb);
    });

    // Create new tenant (Active)
    document.querySelectorAll(".btn-create-new").forEach(btn => {
        btn.addEventListener("click", saveTenantToDb);
    });

    // Save past tenant
    document.querySelectorAll(".btn-save-past").forEach(btn => {
        btn.addEventListener("click", saveTenantToDb);
    });

    // Save draft locally
    document.querySelectorAll(".btn-save-draft").forEach(btn => {
        btn.addEventListener("click", promptAndSaveDraft);
    });

    const clearDraftsBtn = document.getElementById("clearDraftsBtn");
    if (clearDraftsBtn) {
        clearDraftsBtn.addEventListener("click", clearAllDrafts);
    }

    // Export DOCX
    document.querySelectorAll(".btn-export-docx").forEach(btn => {
        btn.addEventListener("click", exportDocxFromTemplate);
    });

    document.addEventListener("wings:updated", syncUnitConfigWingOptions);
    document.addEventListener("units:updated", (e) => syncUnitConfigList(e.detail));
}
