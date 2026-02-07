import { getLocalList, LOCAL_KEYS } from "../../api/localStore.js";
import { currentFlow } from "../../state.js";
import { showToast } from "../../utils/ui.js";
import { isValidationEnabled } from "../shared/validationMode.js";

const REQUIRED_FIELD_IDS = [
    "agreement_date",
    "grn_number",
    "tenancy_comm",
    "landlord_selector",
    "Landlord_name",
    "landlord_address",
    "landlord_aadhar",
    "Tenant_Full_Name",
    "Tenant_occupation",
    "tenant_Aadhar",
    "tenant_mobile",
    "Tenant_Permanent_Address",
    "unit_selector",
    "wing",
    "unit_number_display",
    "floor_of_building",
    "direction_build",
    "meter_number",
    "pet_text_area",
    "rent_amount",
    "payable_date",
    "secu_depo",
    "rent_inc",
    "rent_rev_number",
    "rent_rev_unit",
    "notice_num_t",
    "notice_num_l",
    "late_rent",
    "late_days",
];

const INVALID_INPUT_CLASSES = ["border-rose-500", "bg-rose-50", "text-rose-700", "ring-1", "ring-rose-200"];

let tenantFormValidationBound = false;

function normalizeToken(value) {
    return (value ?? "").toString().trim().toLowerCase();
}

function normalizeGrn(value) {
    return (value ?? "").toString().trim().toUpperCase();
}

function getFieldById(id) {
    return document.getElementById(id);
}

function getFieldLabel(field, fallbackId = "") {
    if (!field) return fallbackId || "field";
    if (field.labels && field.labels.length) {
        const text = (field.labels[0].textContent || "").trim();
        if (text) return text.replace(/\s+/g, " ");
    }
    const labelledByFor = document.querySelector(`label[for="${field.id}"]`);
    if (labelledByFor) {
        const text = (labelledByFor.textContent || "").trim();
        if (text) return text.replace(/\s+/g, " ");
    }
    const wrapper = field.closest("div");
    const inlineLabel = wrapper ? wrapper.querySelector("label") : null;
    if (inlineLabel) {
        const text = (inlineLabel.textContent || "").trim();
        if (text) return text.replace(/\s+/g, " ");
    }
    return fallbackId || field.id || "field";
}

function isEmptyValue(field) {
    if (!field) return true;
    const raw = (field.value ?? "").toString();
    return raw.trim() === "";
}

function clearFieldValidationState(field) {
    if (!field) return;
    INVALID_INPUT_CLASSES.forEach((cls) => field.classList.remove(cls));
    delete field.dataset.validationInvalid;
    delete field.dataset.validationMessage;
}

function setFieldValidationError(field, message = "") {
    if (!field) return;
    INVALID_INPUT_CLASSES.forEach((cls) => field.classList.add(cls));
    field.dataset.validationInvalid = "1";
    field.dataset.validationMessage = message;
}

function clearFormValidationState() {
    REQUIRED_FIELD_IDS.forEach((fieldId) => {
        clearFieldValidationState(getFieldById(fieldId));
    });
}

function collectMissingRequiredFields() {
    const errors = [];
    REQUIRED_FIELD_IDS.forEach((fieldId) => {
        const field = getFieldById(fieldId);
        if (!field) return;
        if (!isEmptyValue(field)) return;
        errors.push({
            fieldId,
            field,
            message: `${getFieldLabel(field, fieldId)} is required.`,
        });
    });
    return errors;
}

async function findDuplicateGrn(grnValue, { tenantId = "", tenancyId = "" } = {}) {
    const normalized = normalizeGrn(grnValue);
    if (!normalized) return null;

    const [tenancies, tenants] = await Promise.all([
        getLocalList(LOCAL_KEYS.tenancies, []),
        getLocalList(LOCAL_KEYS.tenants, []),
    ]);

    const duplicateInTenancies = (Array.isArray(tenancies) ? tenancies : []).find((entry) => {
        const currentGrn = normalizeGrn(entry?.grn_number || entry?.grnNumber || "");
        if (!currentGrn || currentGrn !== normalized) return false;
        const entryTenancyId = normalizeToken(entry?.tenancy_id || entry?.tenancyId || "");
        if (tenancyId && entryTenancyId && entryTenancyId === normalizeToken(tenancyId)) return false;
        return true;
    });
    if (duplicateInTenancies) return duplicateInTenancies;

    return (Array.isArray(tenants) ? tenants : []).find((entry) => {
        const currentGrn = normalizeGrn(
            entry?.grnNumber || entry?.grn_number || entry?.templateData?.["GRN number"] || ""
        );
        if (!currentGrn || currentGrn !== normalized) return false;
        const entryTenantId = normalizeToken(entry?.tenantId || entry?.tenant_id || "");
        if (tenantId && entryTenantId && entryTenantId === normalizeToken(tenantId)) return false;
        return true;
    });
}

export async function validateTenantFormBeforeSave(payload = {}) {
    if (!isValidationEnabled()) {
        clearFormValidationState();
        return { ok: true, skipped: true };
    }

    clearFormValidationState();
    const errors = collectMissingRequiredFields();

    const grnField = getFieldById("grn_number");
    const grnValue = grnField?.value || payload?.templateData?.["GRN number"] || payload?.grn || "";
    if ((grnValue || "").toString().trim()) {
        const duplicate = await findDuplicateGrn(grnValue, {
            tenantId: payload?.tenantId || "",
            tenancyId: payload?.tenancyId || "",
        });
        if (duplicate) {
            errors.push({
                fieldId: "grn_number",
                field: grnField,
                message: "GRN Number already exists in local data.",
            });
        }
    }

    if (!errors.length) {
        return { ok: true };
    }

    errors.forEach((entry) => setFieldValidationError(entry.field, entry.message));
    const first = errors[0]?.field || null;
    if (first && typeof first.scrollIntoView === "function") {
        first.scrollIntoView({ behavior: "smooth", block: "center" });
        if (!first.disabled && typeof first.focus === "function") first.focus();
    }

    const flowLabel = currentFlow === "createTenantNew" ? "Create New Tenant" : "Create Tenant on DB";
    const firstMessage = errors[0]?.message || "Please fill all required fields.";
    const suffix = errors.length > 1 ? ` (${errors.length} issues found)` : "";
    showToast(`${flowLabel}: ${firstMessage}${suffix}`, "error");

    return { ok: false, errors };
}

export function initTenantFormValidation() {
    if (tenantFormValidationBound) return;
    tenantFormValidationBound = true;

    REQUIRED_FIELD_IDS.forEach((fieldId) => {
        const field = getFieldById(fieldId);
        if (!field) return;
        const handler = () => clearFieldValidationState(field);
        field.addEventListener("input", handler);
        field.addEventListener("change", handler);
    });

    document.addEventListener("flow:changed", clearFormValidationState);
    document.addEventListener("app:validation-mode-changed", clearFormValidationState);
}

