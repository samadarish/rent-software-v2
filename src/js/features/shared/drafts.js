/**
 * Draft Management Feature Module
 *
 * Saves and restores form drafts per flow using browser localStorage.
 */

import { currentFlow } from "../../state.js";
import { getFamilyMembersFromTable, setFamilyMembersInTable, syncTenantToFamilyTable } from "../tenants/family.js";
import { pruneDraftValues } from "../tenants/draftHydration.js";
import { collectDraftFormValues, hydrateTenantFormFromDraft, resetTenantFormValues } from "../tenants/formState.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

const DRAFT_KEY_PREFIX = "rent_sw_draft_";
const GUARDED_FLOWS = ["agreement", "createTenantNew"];

const draftDirtyState = {
    agreement: false,
    createTenantNew: false,
};

const LANDLORD_FIELD_IDS = new Set(["Landlord_name", "landlord_aadhar", "landlord_address"]);

let suppressDirtyEvents = false;

let pendingDraftNameHandler = null;
let pendingNavigationHandler = null;
let pendingNavigationAfterSave = null;

function getDraftKey(mode) {
    return `${DRAFT_KEY_PREFIX}${mode}`;
}

function updateSaveDraftButtons(mode) {
    if (!GUARDED_FLOWS.includes(mode)) return;
    const dirty = Boolean(draftDirtyState[mode]);
    document.querySelectorAll(".btn-save-draft").forEach((btn) => {
        btn.classList.toggle("hidden", !dirty);
    });
}

function refreshDraftPickers(mode, selectedName = "") {
    const drafts = getDraftList(mode);
    document.querySelectorAll(".draft-picker").forEach((picker) => {
        picker.innerHTML = "";
        const placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = drafts.length ? "Load draft" : "No drafts saved";
        picker.appendChild(placeholder);

        drafts.forEach((draft) => {
            const opt = document.createElement("option");
            opt.value = draft.name;
            opt.textContent = draft.name;
            if (draft.name === selectedName) opt.selected = true;
            picker.appendChild(opt);
        });

        picker.disabled = drafts.length === 0;
    });
}

export function markDraftDirty(mode) {
    if (!GUARDED_FLOWS.includes(mode)) return;
    if (suppressDirtyEvents) return;
    draftDirtyState[mode] = true;
    updateSaveDraftButtons(mode);
}

export function setDraftClean(mode) {
    if (!GUARDED_FLOWS.includes(mode)) return;
    draftDirtyState[mode] = false;
    updateSaveDraftButtons(mode);
}

export function isDraftDirty(mode) {
    return Boolean(draftDirtyState[mode]);
}

function getDraftNameSuggestion() {
    const tenantNameInput = document.getElementById("Tenant_Full_Name");
    if (tenantNameInput?.value) return tenantNameInput.value.trim();
    return "";
}

function getDraftList(mode) {
    const key = getDraftKey(mode);
    const raw = localStorage.getItem(key);
    if (!raw) return [];
    try {
        const parsed = JSON.parse(raw);
        const drafts = Array.isArray(parsed) ? parsed : (parsed && typeof parsed === "object" ? [parsed] : []);
        if (!drafts.length) return [];
        let didPrune = false;
        const normalized = drafts.map((draft) => {
            const values = pruneDraftValues(draft?.values || {});
            if (
                draft?.values &&
                ("rent_amount_words" in draft.values || "secu_amount_words" in draft.values)
            ) {
                didPrune = true;
            }
            return { ...draft, values };
        });
        if (didPrune) persistDraftList(mode, normalized);
        return normalized;
    } catch (err) {
        console.error("Error parsing draft list", err);
        return [];
    }
}

function persistDraftList(mode, drafts) {
    const key = getDraftKey(mode);
    localStorage.setItem(key, JSON.stringify(drafts));
}

function getModeLabel(mode) {
    if (mode === "createTenantNew") return "Create tenant";
    return "Agreement";
}

export function saveDraftForCurrentFlow(name) {
    const draftName = (name || getDraftNameSuggestion()).trim();
    if (!draftName) {
        showToast("Draft name is required", "error");
        return;
    }

    const drafts = getDraftList(currentFlow);
    const payload = {
        name: draftName,
        savedAt: new Date().toISOString(),
        values: collectDraftFormValues(),
        family: getFamilyMembersFromTable(),
    };
    try {
        const existingIndex = drafts.findIndex((draft) => draft.name === draftName);
        if (existingIndex >= 0) {
            drafts[existingIndex] = payload;
        } else {
            drafts.push(payload);
        }
        persistDraftList(currentFlow, drafts);
        showToast(`${getModeLabel(currentFlow)} draft "${draftName}" saved locally`, "success");
        setDraftClean(currentFlow);
        refreshDraftPickers(currentFlow, draftName);
        if (pendingNavigationAfterSave) {
            const navHandler = pendingNavigationAfterSave;
            pendingNavigationAfterSave = null;
            pendingNavigationHandler = null;
            navHandler();
        }
    } catch (err) {
        console.error("Error saving draft", err);
        showToast("Could not save draft", "error");
    }
}

export function loadDraftForFlow(mode, draftName = null) {
    const drafts = getDraftList(mode);
    const draftToLoad = draftName
        ? drafts.find((draft) => draft.name === draftName)
        : drafts[drafts.length - 1];

    if (!draftToLoad) {
        suppressDirtyEvents = true;
        try {
            resetTenantFormValues();
            setFamilyMembersInTable([]);
            syncTenantToFamilyTable();
        } finally {
            suppressDirtyEvents = false;
        }
        setDraftClean(mode);
        refreshDraftPickers(mode);
        return;
    }

    try {
        suppressDirtyEvents = true;
        try {
            resetTenantFormValues();
            hydrateTenantFormFromDraft(draftToLoad.values || {});
            setFamilyMembersInTable(draftToLoad.family || []);
            syncTenantToFamilyTable();
        } finally {
            suppressDirtyEvents = false;
        }
        showToast(`${getModeLabel(mode)} draft "${draftToLoad.name}" loaded`, "info");
        setDraftClean(mode);
        refreshDraftPickers(mode, draftToLoad.name);
    } catch (err) {
        console.error("Error loading draft", err);
        suppressDirtyEvents = true;
        try {
            resetTenantFormValues();
            setFamilyMembersInTable([]);
            syncTenantToFamilyTable();
        } finally {
            suppressDirtyEvents = false;
        }
        markDraftDirty(mode);
        refreshDraftPickers(mode);
        showToast("Draft could not be loaded", "error");
    }
}

export function clearAllDrafts() {
    try {
        Object.keys(localStorage)
            .filter((key) => key.startsWith(DRAFT_KEY_PREFIX))
            .forEach((key) => localStorage.removeItem(key));
        showToast("All drafts cleared", "success");
        refreshDraftPickers(currentFlow);
        markDraftDirty(currentFlow);
    } catch (err) {
        console.error("Error clearing drafts", err);
        showToast("Unable to clear drafts", "error");
    }
}

export function promptAndSaveDraft() {
    const suggested = getDraftNameSuggestion();
    openDraftNameModal(suggested);
}

export function initDraftUi() {
    const formSection = document.getElementById("formSection");
    if (formSection) {
        ["input", "change"].forEach((evt) => {
            formSection.addEventListener(evt, (event) => {
                const target = event.target;
                if (target instanceof HTMLElement && target.id && LANDLORD_FIELD_IDS.has(target.id)) return;
                if (suppressDirtyEvents) return;
                markDraftDirty(currentFlow);
            });
        });
    }

    document.querySelectorAll(".draft-picker").forEach((picker) => {
        picker.addEventListener("change", (e) => {
            const target = e.target;
            if (!(target instanceof HTMLSelectElement)) return;
            if (!target.value) return;
            loadDraftForFlow(currentFlow, target.value);
        });
    });

    refreshDraftPickers(currentFlow);
    updateSaveDraftButtons(currentFlow);

    const draftNameModal = document.getElementById("draftNameModal");
    const draftNameSaveBtn = document.getElementById("draftNameSaveBtn");
    const draftNameCancelBtn = document.getElementById("draftNameCancelBtn");
    const draftNameInput = document.getElementById("draftNameInput");

    if (draftNameSaveBtn && draftNameInput && draftNameModal) {
        draftNameSaveBtn.addEventListener("click", () => {
            const name = (draftNameInput.value || "").trim();
            if (!name) {
                draftNameInput.focus();
                return;
            }
            draftNameModal.classList.add("hidden");
            if (pendingDraftNameHandler) pendingDraftNameHandler(name);
            pendingDraftNameHandler = null;
        });
        draftNameInput.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                draftNameSaveBtn.click();
            }
        });
    }

    if (draftNameCancelBtn && draftNameModal) {
        draftNameCancelBtn.addEventListener("click", () => {
            draftNameModal.classList.add("hidden");
            pendingDraftNameHandler = null;
            pendingNavigationAfterSave = null;
        });
    }

    const navModal = document.getElementById("unsavedDraftModal");
    const navStayBtn = document.getElementById("unsavedDraftStayBtn");
    const navLeaveBtn = document.getElementById("unsavedDraftLeaveBtn");
    const navSaveBtn = document.getElementById("unsavedDraftSaveBtn");

    if (navStayBtn && navModal) {
        navStayBtn.addEventListener("click", () => {
            hideModal(navModal);
            pendingNavigationHandler = null;
            pendingNavigationAfterSave = null;
        });
    }

    if (navLeaveBtn && navModal) {
        navLeaveBtn.addEventListener("click", () => {
            hideModal(navModal);
            if (pendingNavigationHandler) pendingNavigationHandler();
            pendingNavigationHandler = null;
            pendingNavigationAfterSave = null;
        });
    }

    if (navSaveBtn && navModal) {
        navSaveBtn.addEventListener("click", () => {
            hideModal(navModal);
            pendingNavigationAfterSave = pendingNavigationHandler;
            promptAndSaveDraft();
        });
    }
}

export function syncDraftUiForFlow(mode) {
    refreshDraftPickers(mode);
    updateSaveDraftButtons(mode);
}

export function openDraftNameModal(value = "") {
    const modal = document.getElementById("draftNameModal");
    const input = document.getElementById("draftNameInput");
    if (!modal || !input) return;
    input.value = value;
    input.placeholder = value || "Enter a name for this draft";
    // Pre-select the suggested tenant name so users can accept or edit it quickly without native prompts
    input.focus();
    input.select();
    pendingDraftNameHandler = (name) => saveDraftForCurrentFlow(name);
    showModal(modal);
}

export function openUnsavedDraftModal(onConfirm) {
    const modal = document.getElementById("unsavedDraftModal");
    if (!modal) {
        onConfirm();
        return;
    }
    pendingNavigationHandler = onConfirm;
    pendingNavigationAfterSave = null;
    showModal(modal);
}
