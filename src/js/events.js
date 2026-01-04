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
    saveLandlordConfig,
    deleteLandlordConfig,
} from "./api/sheets.js";
import { exportDocxFromTemplate } from "./features/agreements/docx.js";
import { buildUnitLabel, numberToIndianWords } from "./utils/formatters.js";
import { clearAllDrafts, promptAndSaveDraft } from "./features/shared/drafts.js";
import { handleNoGrnToggle } from "./features/tenants/formState.js";
import { cloneSelectOptions, hideModal, showToast } from "./utils/ui.js";

let unitConfigCache = [];
let landlordConfigCache = [];
let unitConfigEditId = "";
let landlordConfigEditId = "";

const UNIT_SAVE_LABEL = "Save unit";
const UNIT_UPDATE_LABEL = "Update unit";
const UNIT_SAVE_CLASSES =
    "px-3 py-1.5 rounded bg-indigo-600 text-white text-[11px] font-semibold hover:bg-indigo-500";
const UNIT_UPDATE_CLASSES =
    "px-3 py-1.5 rounded bg-amber-600 text-white text-[11px] font-semibold hover:bg-amber-500";
const LANDLORD_SAVE_LABEL = "Save landlord";
const LANDLORD_UPDATE_LABEL = "Update landlord";
const LANDLORD_SAVE_CLASSES =
    "px-3 py-1.5 rounded bg-emerald-600 text-white text-[11px] font-semibold hover:bg-emerald-500";
const LANDLORD_UPDATE_CLASSES =
    "px-3 py-1.5 rounded bg-amber-600 text-white text-[11px] font-semibold hover:bg-amber-500";

/**
 * Copies the wing dropdown options into the unit configuration modal selectors.
 */
function syncUnitConfigWingOptions() {
    cloneSelectOptions("wing", "unitConfigWing", { preserveSelection: false });
}

/**
 * Populates the unit configuration modal list from the provided units array.
 * @param {Array} units - Units fetched from Google Sheets.
 */
function syncUnitConfigList(units) {
    unitConfigCache = Array.isArray(units) ? units : [];
    const list = document.getElementById("unitConfigList");
    if (list) {
        list.innerHTML = "";
        unitConfigCache.forEach((u) => {
            const row = document.createElement("div");
            row.className =
                "flex items-center justify-between border border-slate-200 rounded-lg px-2 py-1.5 bg-white/90 hover:border-slate-300";
            const label = document.createElement("div");
            label.className = "flex flex-col";
            const unitLabel = buildUnitLabel(u) || u.unit_id || "Unit";
            const title = document.createElement("div");
            title.className = "text-slate-800 font-semibold";
            title.textContent = unitLabel;
            const meta = document.createElement("div");
            meta.className = "text-[10px] text-slate-500";
            const metaParts = [u.floor, u.direction, u.meter_number || u.meterNumber]
                .map((val) => (val || "").toString().trim())
                .filter(Boolean);
            meta.textContent = metaParts.join(" | ");
            label.append(title, meta);
            const actions = document.createElement("div");
            actions.className = "flex gap-2";
            const editBtn = document.createElement("button");
            editBtn.textContent = "Edit";
            editBtn.className =
                "text-[10px] px-2 py-0.5 rounded border border-indigo-200 bg-indigo-50 text-indigo-700 hover:bg-indigo-100";
            editBtn.addEventListener("click", () => {
                setUnitEditMode(u);
            });
            const deleteBtn = document.createElement("button");
            deleteBtn.textContent = "Delete";
            deleteBtn.className =
                "text-[10px] px-2 py-0.5 rounded border border-rose-200 bg-rose-50 text-rose-700 hover:bg-rose-100";
            deleteBtn.addEventListener("click", async () => {
                await deleteUnitConfig(u.unit_id);
                if (unitConfigEditId && unitConfigEditId === u.unit_id) {
                    clearUnitEditMode();
                }
            });
            actions.append(editBtn, deleteBtn);
            row.append(label, actions);
            list.appendChild(row);
        });
    }

    if (unitConfigEditId && !unitConfigCache.some((u) => u.unit_id === unitConfigEditId)) {
        clearUnitEditMode();
    }
}

/**
 * Updates the landlord config list with the latest saved landlords.
 * @param {Array} landlords - Landlord records retrieved from Sheets.
 */
function syncLandlordConfigList(landlords) {
    landlordConfigCache = Array.isArray(landlords) ? landlords : [];
    const list = document.getElementById("landlordConfigList");
    if (list) {
        list.innerHTML = "";
        landlordConfigCache.forEach((l) => {
            const row = document.createElement("div");
            row.className =
                "flex items-center justify-between border border-slate-200 rounded-lg px-2 py-1.5 bg-white/90 hover:border-slate-300";
            const label = document.createElement("div");
            label.className = "flex flex-col";
            const title = document.createElement("div");
            title.className = "text-slate-800 font-semibold";
            title.textContent = l.name || l.landlord_id || "Landlord";
            const meta = document.createElement("div");
            meta.className = "text-[10px] text-slate-500";
            const metaParts = [l.aadhaar, l.address].map((val) => (val || "").toString().trim()).filter(Boolean);
            meta.textContent = metaParts.join(" | ");
            label.append(title, meta);
            const actions = document.createElement("div");
            actions.className = "flex gap-2";
            const editBtn = document.createElement("button");
            editBtn.textContent = "Edit";
            editBtn.className =
                "text-[10px] px-2 py-0.5 rounded border border-indigo-200 bg-indigo-50 text-indigo-700 hover:bg-indigo-100";
            editBtn.addEventListener("click", () => {
                setLandlordEditMode(l);
            });
            const deleteBtn = document.createElement("button");
            deleteBtn.textContent = "Delete";
            deleteBtn.className =
                "text-[10px] px-2 py-0.5 rounded border border-rose-200 bg-rose-50 text-rose-700 hover:bg-rose-100";
            deleteBtn.addEventListener("click", async () => {
                await deleteLandlordConfig(l.landlord_id);
                if (landlordConfigEditId && landlordConfigEditId === l.landlord_id) {
                    clearLandlordEditMode();
                }
            });
            actions.append(editBtn, deleteBtn);
            row.append(label, actions);
            list.appendChild(row);
        });
    }

    if (landlordConfigEditId && !landlordConfigCache.some((l) => l.landlord_id === landlordConfigEditId)) {
        clearLandlordEditMode();
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
        const pill = document.createElement("div");
        pill.className =
            "px-2 py-1 rounded-full bg-white border border-slate-200 text-[11px] flex items-center gap-1";
        const label = document.createElement("span");
        label.textContent = opt.value;
        const close = document.createElement("button");
        close.type = "button";
        close.innerHTML = "&times;";
        close.className =
            "text-[14px] leading-none text-slate-400 w-5 h-5 flex items-center justify-center rounded hover:rounded-full hover:text-rose-600 hover:bg-rose-100";
        close.addEventListener("click", async (event) => {
            event.preventDefault();
            event.stopPropagation();
            const res = await removeWingFromSheet(opt.value);
            if (res?.ok !== false) {
                showToast("Wing removed", "success");
            }
        });
        pill.append(label, close);
        list.appendChild(pill);
    });
}

function setUnitEditMode(unit) {
    if (!unit) return;
    unitConfigEditId = unit.unit_id || "";
    if (!unitConfigEditId) return;
    document.getElementById("unitConfigWing").value = unit.wing || "";
    document.getElementById("unitConfigNumber").value = unit.unit_number || "";
    document.getElementById("unitConfigFloor").value = unit.floor || "";
    document.getElementById("unitConfigDirection").value = unit.direction || "";
    document.getElementById("unitConfigMeter").value = unit.meter_number || "";
    document.getElementById("unitConfigNotes").value = unit.notes || "";
    applyUnitEditState(true);
}

function clearUnitEditMode() {
    unitConfigEditId = "";
    resetUnitConfigFields();
    applyUnitEditState(false);
}

function applyUnitEditState(isEditing) {
    const saveBtn = document.getElementById("unitConfigSaveBtn");
    const notice = document.getElementById("unitConfigEditNotice");
    const cancelBtn = document.getElementById("unitConfigCancelEditBtn");
    if (notice) notice.classList.toggle("hidden", !isEditing);
    if (cancelBtn) cancelBtn.classList.toggle("hidden", !isEditing);
    if (saveBtn) {
        saveBtn.textContent = isEditing ? UNIT_UPDATE_LABEL : UNIT_SAVE_LABEL;
        saveBtn.className = isEditing ? UNIT_UPDATE_CLASSES : UNIT_SAVE_CLASSES;
    }
}

function resetUnitConfigFields() {
    const ids = [
        "unitConfigWing",
        "unitConfigNumber",
        "unitConfigFloor",
        "unitConfigDirection",
        "unitConfigMeter",
        "unitConfigNotes",
    ];
    ids.forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.value = "";
    });
}

function setLandlordEditMode(landlord) {
    if (!landlord) return;
    landlordConfigEditId = landlord.landlord_id || "";
    if (!landlordConfigEditId) return;
    document.getElementById("landlordDefaultName").value = landlord.name || "";
    document.getElementById("landlordDefaultAadhaar").value = landlord.aadhaar || "";
    document.getElementById("landlordDefaultAddress").value = landlord.address || "";
    applyLandlordEditState(true);
}

function clearLandlordEditMode() {
    landlordConfigEditId = "";
    resetLandlordConfigFields();
    applyLandlordEditState(false);
}

function applyLandlordEditState(isEditing) {
    const saveBtn = document.getElementById("landlordDefaultsSaveBtn");
    const notice = document.getElementById("landlordEditNotice");
    const cancelBtn = document.getElementById("landlordDefaultsCancelEditBtn");
    if (notice) notice.classList.toggle("hidden", !isEditing);
    if (cancelBtn) cancelBtn.classList.toggle("hidden", !isEditing);
    if (saveBtn) {
        saveBtn.textContent = isEditing ? LANDLORD_UPDATE_LABEL : LANDLORD_SAVE_LABEL;
        saveBtn.className = isEditing ? LANDLORD_UPDATE_CLASSES : LANDLORD_SAVE_CLASSES;
    }
}

function resetLandlordConfigFields() {
    const ids = ["landlordDefaultName", "landlordDefaultAadhaar", "landlordDefaultAddress"];
    ids.forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.value = "";
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

    bindClick("navLandlordConfigBtn", () => {
        clearUnitEditMode();
        clearLandlordEditMode();
        openLandlordConfigModal();
    });

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
            const wasEditing = !!landlordConfigEditId;
            const landlordId = landlordConfigEditId || "";
            const res = await saveLandlordConfig({ landlordId, name, aadhaar, address });
            if (res?.ok !== false) {
                saveLandlordDefaults({ closeModal: false, showMessage: false });
                showToast(wasEditing ? "Landlord updated" : "Landlord saved", "success");
                clearLandlordEditMode();
            }
        });
    }

    const landlordAddWingBtn = document.getElementById("landlordDefaultWingBtn");
    if (landlordAddWingBtn) {
        landlordAddWingBtn.addEventListener("click", saveWingFromLandlordConfig);
    }

    const landlordCancelEditBtn = document.getElementById("landlordDefaultsCancelEditBtn");
    if (landlordCancelEditBtn) {
        landlordCancelEditBtn.addEventListener("click", () => {
            clearLandlordEditMode();
        });
    }

    const unitConfigSaveBtn = document.getElementById("unitConfigSaveBtn");
    if (unitConfigSaveBtn) {
        unitConfigSaveBtn.addEventListener("click", async () => {
            const wing = document.getElementById("unitConfigWing")?.value || "";
            const unitNumber = document.getElementById("unitConfigNumber")?.value || "";
            if (!wing || !unitNumber) return;
            const wasEditing = !!unitConfigEditId;
            const unitId = unitConfigEditId || "";
            const res = await saveUnitConfig({
                unitId,
                wing,
                unitNumber,
                floor: document.getElementById("unitConfigFloor")?.value || "",
                direction: document.getElementById("unitConfigDirection")?.value || "",
                meterNumber: document.getElementById("unitConfigMeter")?.value || "",
                notes: document.getElementById("unitConfigNotes")?.value || "",
                isOccupied: false,
            });
            if (res?.ok !== false) {
                showToast(wasEditing ? "Unit updated" : "Unit saved", "success");
                clearUnitEditMode();
            }
        });
    }

    const unitCancelEditBtn = document.getElementById("unitConfigCancelEditBtn");
    if (unitCancelEditBtn) {
        unitCancelEditBtn.addEventListener("click", () => {
            clearUnitEditMode();
        });
    }

    const landlordDefaultsCancelBtn = document.getElementById("landlordDefaultsCancelBtn");
    if (landlordDefaultsCancelBtn) {
        landlordDefaultsCancelBtn.addEventListener("click", () => {
            const modal = document.getElementById("landlordConfigModal");
            if (modal) hideModal(modal);
            clearUnitEditMode();
            clearLandlordEditMode();
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
