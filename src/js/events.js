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
    openAppScriptModal,
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
    saveLandlordConfig,
    deleteLandlordConfig,
} from "./api/sheets.js";
import { refreshLandlords } from "./store/masters.js";
import { exportDocxFromTemplate } from "./features/agreements/docx.js";
import { buildUnitLabel, numberToIndianWords } from "./utils/formatters.js";
import { clearAllDrafts, promptAndSaveDraft } from "./features/shared/drafts.js";
import { handleNoGrnToggle } from "./features/tenants/formState.js";
import { cloneSelectOptions, hideModal } from "./utils/ui.js";

let unitConfigCache = [];
let landlordConfigCache = [];

/**
 * Copies the wing dropdown options into the unit configuration modal selectors.
 */
function syncUnitConfigWingOptions() {
    cloneSelectOptions("wing", "unitConfigWing", { preserveSelection: false });
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
        opt.textContent = buildUnitLabel(u);
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
            const unitLabel = buildUnitLabel(u);
            label.textContent = `${unitLabel} - ${u.floor || ""} ${u.direction || ""}`.trim();
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
            fetchWingsFromSheet(true);
        });
        list.appendChild(pill);
    });
}


function wireAmountToWords(inputId, outputId) {
    const input = document.getElementById(inputId);
    if (!input) return;

    input.addEventListener("input", () => {
        const val = parseInt(input.value, 10);
        const out = document.getElementById(outputId);
        if (!out) return;
        out.value = isNaN(val) || val <= 0 ? "" : `${numberToIndianWords(val)} only`;
    });
}

function bindClick(id, handler) {
    const el = document.getElementById(id);
    if (el) {
        el.addEventListener("click", handler);
    }
    return el;
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
        noGrnCheckbox.addEventListener("change", handleNoGrnToggle);
    }

    const navTargets = {
        navDashboardBtn: "dashboard",
        navGenerateBillBtn: "generateBill",
        navPaymentsBtn: "payments",
        navViewTenantsBtn: "viewTenants",
        navCreateTenantBtn: "createTenantNew",
        navCreateAgreementBtn: "agreement",
        navExportDataBtn: "exportData",
    };

    Object.entries(navTargets).forEach(([id, mode]) => {
        bindClick(id, () => switchFlow(mode));
    });

    // Amount to words converters
    wireAmountToWords("rent_amount", "rent_amount_words");
    wireAmountToWords("secu_depo", "secu_amount_words");
    wireAmountToWords("rent_inc", "increase_amount_word");

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
    bindClick("navConfigBtn", () => {
        const url = getAppScriptUrl();
        const input = document.getElementById("appscript_url");
        if (input) input.value = url || "";
        openAppScriptModal({ mode: "input" });
    });

    bindClick("navLandlordConfigBtn", openLandlordConfigModal);

    bindClick("appscriptCancelBtn", () => {
        const modal = document.getElementById("appscriptModal");
        if (modal) hideModal(modal);
    });
    bindClick("appscriptModalClose", () => {
        const modal = document.getElementById("appscriptModal");
        if (modal) hideModal(modal);
    });

    bindClick("appscriptSaveBtn", saveAppScriptUrl);

    const landlordDefaultsSaveBtn = document.getElementById("landlordDefaultsSaveBtn");
    if (landlordDefaultsSaveBtn) {
        landlordDefaultsSaveBtn.addEventListener("click", async () => {
            const name = document.getElementById("landlordDefaultName")?.value.trim() || "";
            const aadhaar = document.getElementById("landlordDefaultAadhaar")?.value.trim() || "";
            const address = document.getElementById("landlordDefaultAddress")?.value.trim() || "";
            const landlordId = document.getElementById("landlordExistingSelect")?.value || "";
            await saveLandlordConfig({ landlordId, name, aadhaar, address });
            saveLandlordDefaults();
            refreshLandlords(true);
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
            refreshLandlords(true);
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
            fetchWingsFromSheet(true);
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
            loadClausesFromSheet(true, true);
        });
    }

    syncUnitConfigWingOptions();
    syncUnitConfigList(unitConfigCache);
    syncWingList();

    // Action buttons (class-based selectors for multiple instances)

    // Save tenant
    document.querySelectorAll(".btn-save-agreement, .btn-create-new").forEach((btn) => {
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

    document.addEventListener("units:updated", (e) => syncUnitConfigList(e.detail));
}
