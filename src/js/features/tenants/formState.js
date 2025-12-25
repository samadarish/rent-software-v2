import { generateNoGrnValue } from "../../utils/grn.js";
import { pruneDraftValues, resolveDraftHydration } from "./draftHydration.js";
import { applyLandlordToForm, applyUnitToPremises, getLandlordCache, getUnitCache } from "./form.js";

const FORM_FIELDS_SELECTOR = "#formSection input, #formSection textarea, #formSection select";
const DERIVED_FIELD_IDS = new Set([
    "Landlord_name",
    "landlord_aadhar",
    "landlord_address",
    "wing",
    "unit_number_display",
    "floor_of_building",
    "direction_build",
    "meter_number",
]);
const PRESERVE_SELECT_IDS = new Set(["landlord_selector", "unit_selector"]);

function applyFormValues(values = {}, options = {}) {
    const skip = options.skip instanceof Set ? options.skip : new Set();
    document.querySelectorAll(FORM_FIELDS_SELECTOR).forEach((el) => {
        if (!el.id || !(el.id in values)) return;
        if (skip.has(el.id)) return;
        if (el.type === "checkbox" || el.type === "radio") {
            el.checked = Boolean(values[el.id]);
        } else if (el.tagName === "SELECT") {
            const nextValue = values[el.id];
            if (PRESERVE_SELECT_IDS.has(el.id) && nextValue) {
                const hasOption = Array.from(el.options).some((opt) => opt.value === nextValue);
                if (!hasOption) {
                    const opt = document.createElement("option");
                    opt.value = nextValue;
                    opt.textContent = "Saved selection";
                    opt.dataset.placeholder = "true";
                    el.appendChild(opt);
                }
            }
            el.value = nextValue;
        } else {
            el.value = values[el.id];
        }
    });
}

function refreshAmountWords() {
    const rent = document.getElementById("rent_amount");
    if (rent) rent.dispatchEvent(new Event("input", { bubbles: true }));
    const deposit = document.getElementById("secu_depo");
    if (deposit) deposit.dispatchEvent(new Event("input", { bubbles: true }));
    const increase = document.getElementById("rent_inc");
    if (increase) increase.dispatchEvent(new Event("input", { bubbles: true }));
}

export function collectDraftFormValues() {
    const values = {};
    document.querySelectorAll(FORM_FIELDS_SELECTOR).forEach((el) => {
        if (!el.id) return;
        if (el.type === "checkbox" || el.type === "radio") {
            values[el.id] = el.checked;
        } else {
            values[el.id] = el.value;
        }
    });
    return pruneDraftValues(values);
}

export function setNoGrnState({ checked, grnValue = "", source = "user" } = {}) {
    const input = document.getElementById("grn_number");
    const checkbox = document.getElementById("grnNoCheckbox");
    if (!input || !checkbox) return;

    const isChecked = Boolean(checked);
    checkbox.checked = isChecked;

    if (isChecked) {
        const nextValue = grnValue || generateNoGrnValue();
        input.value = nextValue;
        input.disabled = true;
        input.readOnly = true;
        input.dataset.noGrn = "1";
        return;
    }

    input.disabled = false;
    input.readOnly = false;
    input.dataset.noGrn = "0";
    input.value = source === "user" ? "" : grnValue || "";
}

export function handleNoGrnToggle() {
    const checkbox = document.getElementById("grnNoCheckbox");
    if (!checkbox) return;
    setNoGrnState({ checked: checkbox.checked, source: "user" });
}

export function resetTenantFormValues() {
    document.querySelectorAll(FORM_FIELDS_SELECTOR).forEach((el) => {
        if (!el.id) return;
        if (el.type === "checkbox" || el.type === "radio") {
            el.checked = false;
        } else {
            el.value = "";
        }
    });
    applyLandlordToForm(null);
    applyUnitToPremises(null);
    setNoGrnState({ checked: false, source: "reset" });
}

export function hydrateTenantFormFromDraft(rawValues = {}) {
    const values = pruneDraftValues(rawValues);
    const skip = new Set(DERIVED_FIELD_IDS);
    if (values.grnNoCheckbox) skip.add("grn_number");

    applyFormValues(values, { skip });

    const hydration = resolveDraftHydration(values, {
        landlords: getLandlordCache(),
        units: getUnitCache(),
        generateNoGrnValue,
    });

    if (hydration.landlordId) {
        const landlord = {
            name: hydration.landlordFields.name,
            aadhaar: hydration.landlordFields.aadhaar,
            address: hydration.landlordFields.address,
        };
        applyLandlordToForm(landlord);
    } else {
        applyLandlordToForm(null);
    }

    if (hydration.unitId) {
        const unit = {
            wing: hydration.unitFields.wing,
            unit_number: hydration.unitFields.unitNumber,
            floor: hydration.unitFields.floor,
            direction: hydration.unitFields.direction,
            meter_number: hydration.unitFields.meter,
        };
        applyUnitToPremises(unit);
    } else {
        applyUnitToPremises(null);
    }

    setNoGrnState({
        checked: hydration.noGrnChecked,
        grnValue: hydration.grnValue,
        source: "draft",
    });

    refreshAmountWords();
}
