/**
 * Form Feature Module
 * 
 * Handles all form-related logic:
 * - Initializing form dropdowns and options
 * - Collecting form data for template generation
 * - Collecting full payload for database storage
 */

import { floorOptions } from "../../constants.js";
import { toOrdinal, formatDateForDoc } from "../../utils/formatters.js";
import { htmlToMarkedText } from "../../utils/htmlUtils.js";
import { getSelectedClauses } from "../agreements/clauses.js";
import { getFamilyMembersFromTable } from "./family.js";
import { currentFlow } from "../../state.js";
import { fetchUnitsFromSheet, fetchLandlordsFromSheet } from "../../api/sheets.js";

let unitCache = [];
let unitsLoaded = false;
let landlordCache = [];
let landlordsLoaded = false;

function normalizeOccupiedFlag(raw) {
    if (raw === true || raw === false) return raw;
    if (typeof raw === "string") return raw.toLowerCase() === "true";
    return !!raw;
}

function getLandlordById(landlordId) {
    return landlordCache.find((l) => l.landlord_id === landlordId);
}

function applyLandlordToForm(landlord) {
    const name = document.getElementById("Landlord_name");
    const aadhaar = document.getElementById("landlord_aadhar");
    const address = document.getElementById("landlord_address");
    if (name) name.value = landlord?.name || "";
    if (aadhaar) aadhaar.value = landlord?.aadhaar || "";
    if (address) address.value = landlord?.address || "";
}

function getUnitById(unitId) {
    return unitCache.find((u) => u.unit_id === unitId);
}

function applyUnitToPremises(unit) {
    const wing = document.getElementById("wing");
    const unitNumber = document.getElementById("unit_number_display");
    const floor = document.getElementById("floor_of_building");
    const direction = document.getElementById("direction_build");
    const meter = document.getElementById("meter_number");

    if (wing) wing.value = unit?.wing || "";
    if (unitNumber) unitNumber.value = unit?.unit_number || "";
    if (floor) floor.value = unit?.floor || "";
    if (direction) direction.value = unit?.direction || "";
    if (meter) meter.value = unit?.meter_number || "";
}

function populateLandlordSelect() {
    const select = document.getElementById("landlord_selector");
    if (!select) return;
    const previous = select.value;
    select.innerHTML = '<option value="">Choose a saved landlord</option>';
    landlordCache.forEach((l) => {
        const opt = document.createElement("option");
        opt.value = l.landlord_id;
        opt.textContent = l.name || l.landlord_id;
        opt.dataset.aadhaar = l.aadhaar || "";
        opt.dataset.address = l.address || "";
        select.appendChild(opt);
    });

    if (previous && Array.from(select.options).some((o) => o.value === previous)) {
        select.value = previous;
        applyLandlordToForm(getLandlordById(previous));
    } else if (!previous && landlordCache.length === 1) {
        select.value = landlordCache[0].landlord_id;
        applyLandlordToForm(landlordCache[0]);
    } else {
        applyLandlordToForm(null);
    }
}

function populateUnitSelectForFlow() {
    const select = document.getElementById("unit_selector");
    if (!select) return;

    const allowOccupied = currentFlow !== "createTenantNew";
    const available = allowOccupied
        ? unitCache
        : unitCache.filter((u) => !normalizeOccupiedFlag(u.is_occupied));

    const previous = select.value;
    select.innerHTML = '<option value="">Select a unit to auto-fill details</option>';
    available.forEach((u) => {
        const opt = document.createElement("option");
        opt.value = u.unit_id;
        const labelParts = [u.wing, u.unit_number].filter(Boolean).join(" - ");
        opt.textContent = labelParts || u.unit_id;
        opt.dataset.wing = u.wing || "";
        opt.dataset.floor = u.floor || "";
        opt.dataset.direction = u.direction || "";
        opt.dataset.meter = u.meter_number || "";
        select.appendChild(opt);
    });

    if (previous && Array.from(select.options).some((o) => o.value === previous)) {
        select.value = previous;
        const unit = getUnitById(previous);
        applyUnitToPremises(unit);
    } else {
        applyUnitToPremises(null);
    }
}

function syncUnitsFromEvent(event) {
    if (event?.detail && Array.isArray(event.detail)) {
        unitCache = event.detail.map((u) => ({
            ...u,
            is_occupied: normalizeOccupiedFlag(u.is_occupied),
        }));
        unitsLoaded = true;
    }
    populateUnitSelectForFlow();
}

function getSelectedUnitForForm() {
    const select = document.getElementById("unit_selector");
    if (!select || !select.value) return null;
    return getUnitById(select.value);
}

export async function refreshUnitOptions(force = false) {
    if (unitsLoaded && !force) {
        populateUnitSelectForFlow();
        return unitCache;
    }

    const data = await fetchUnitsFromSheet(force);
    unitCache = Array.isArray(data.units)
        ? data.units.map((u) => ({
              ...u,
              is_occupied: normalizeOccupiedFlag(u.is_occupied),
          }))
        : [];
    unitsLoaded = true;
    populateUnitSelectForFlow();
    document.dispatchEvent(new CustomEvent("units:updated", { detail: unitCache }));
    return unitCache;
}

export async function refreshLandlordOptions(force = false) {
    if (landlordsLoaded && !force) {
        populateLandlordSelect();
        return landlordCache;
    }

    const data = await fetchLandlordsFromSheet(force);
    landlordCache = Array.isArray(data.landlords)
        ? data.landlords.map((l) => ({ ...l }))
        : [];
    landlordsLoaded = true;
    populateLandlordSelect();
    document.dispatchEvent(new CustomEvent("landlords:updated", { detail: landlordCache }));
    return landlordCache;
}

/**
 * Initializes all form dropdowns and select options
 * Populates floor options, payable date, notice periods, and rent review periods
 */
export function initFormOptions() {
    // Floor dropdown
    const floorSelect = document.getElementById("floor_of_building");
    if (floorSelect && floorSelect.tagName === "SELECT") {
        floorOptions.forEach((label) => {
            const opt = document.createElement("option");
            opt.value = label;
            opt.textContent = label;
            floorSelect.appendChild(opt);
        });
    }

    // Payable date dropdown (1-31)
    const paySel = document.getElementById("payable_date");
    if (paySel) {
        for (let i = 1; i <= 31; i++) {
            const opt = document.createElement("option");
            opt.value = i;
            opt.textContent = toOrdinal(i);
            paySel.appendChild(opt);
        }
    }

    // Notice period dropdowns (1-12 months)
    const noticeTenant = document.getElementById("notice_num_t");
    const noticeLandlord = document.getElementById("notice_num_l");

    for (let i = 0; i <= 12; i++) {
        const label = i.toString();
        if (i > 0) {
            if (noticeTenant) {
                const opt2 = new Option(label, label);
                noticeTenant.appendChild(opt2);
            }
            if (noticeLandlord) {
                const opt3 = new Option(label, label);
                noticeLandlord.appendChild(opt3);
            }
        }
    }

    const unitSelect = document.getElementById("unit_selector");
    if (unitSelect) {
        unitSelect.addEventListener("change", (e) => {
            const unit = getUnitById(e.target.value);
            applyUnitToPremises(unit);
        });
    }

    const landlordSelect = document.getElementById("landlord_selector");
    if (landlordSelect) {
        landlordSelect.addEventListener("change", (e) => {
            const landlord = getLandlordById(e.target.value);
            applyLandlordToForm(landlord);
        });
    }

    document.addEventListener("units:updated", syncUnitsFromEvent);
    document.addEventListener("flow:changed", populateUnitSelectForFlow);
    document.addEventListener("landlords:updated", () => populateLandlordSelect());
}

/**
 * Collects all form data and formats it for the DOCX template
 * @returns {Object} Formatted data object for template generation
 */
export function collectFormDataForTemplate() {
    const agreementDate = document.getElementById("agreement_date").value;
    const tenancyComm = document.getElementById("tenancy_comm").value;
    const tenancyEndEl = document.getElementById("tenancy_end");
    const tenancyEnd = tenancyEndEl ? tenancyEndEl.value : "";
    const selectedUnit = getSelectedUnitForForm();
    const landlordSelect = document.getElementById("landlord_selector");
    const selectedLandlord = landlordSelect && landlordSelect.value ? getLandlordById(landlordSelect.value) : null;

    const clausesBySection = getSelectedClauses();

    // Convert clauses to array format with markdown-style bold
    const tenantClausesArray = (clausesBySection.tenant || []).map((c) => ({
        text: htmlToMarkedText(c.html || c.text),
    }));
    const landlordClausesArray = (clausesBySection.landlord || []).map((c) => ({
        text: htmlToMarkedText(c.html || c.text),
    }));
    const penaltyClausesArray = (clausesBySection.penalties || []).map((c) => ({
        text: htmlToMarkedText(c.html || c.text),
    }));
    const miscClausesArray = (clausesBySection.misc || []).map((c) => ({
        text: htmlToMarkedText(c.html || c.text),
    }));

    // Get family members
    const familyArr = getFamilyMembersFromTable().filter((f) => f.name);

    // Format family members as a text block
    const familyBlock = familyArr
        .map(
            (f, i) =>
                `${i + 1}. ${f.name} (${f.relationship || "-"}, ${f.occupation || "-"}) ` +
                `– Aadhaar: ${f.aadhaar || "-"}, Address: ${f.address || "-"}`
        )
        .join("\n");

    // Calculate rent increase amount in words
    const increaseAmountWord = "";

    const landlordName = selectedLandlord?.name ?? document.getElementById("Landlord_name").value;
    const landlordAddress = selectedLandlord?.address ?? document.getElementById("landlord_address").value;
    const landlordAadhaar = selectedLandlord?.aadhaar ?? document.getElementById("landlord_aadhar").value;

    const data = {
        "GRN number": document.getElementById("grn_number").value.trim(),
        agreement_date: formatDateForDoc(agreementDate),
        agreement_date_raw: agreementDate || "",

        landlord_id: selectedLandlord?.landlord_id || "",
        Landlord_name: String(landlordName || "").trim(),
        landlord_address: String(landlordAddress || "").trim(),
        landlord_aadhar: String(landlordAadhaar || "").trim(),

        tenancy_comm: formatDateForDoc(tenancyComm),
        tenancy_comm_raw: tenancyComm || "",
        tenancy_end: formatDateForDoc(tenancyEnd),
        tenancy_end_raw: tenancyEnd || "",
        notice_num_t: document.getElementById("notice_num_t").value.trim(),
        notice_num_l: document.getElementById("notice_num_l").value.trim(),

        Tenant_Full_Name: document.getElementById("Tenant_Full_Name").value.trim(),
        Tenant_Permanent_Address: document
            .getElementById("Tenant_Permanent_Address")
            .value.trim(),
        tenant_Aadhar: document.getElementById("tenant_Aadhar").value.trim(),
        tenant_mobile: document.getElementById("tenant_mobile").value.trim(),
        unit_id: selectedUnit?.unit_id || "",
        unit_number: selectedUnit?.unit_number || "",

        "floor_of_building": document.getElementById("floor_of_building").value,
        direction_build: document.getElementById("direction_build").value,
        meter_number: document.getElementById("meter_number").value.trim(),

        rent_amount: document.getElementById("rent_amount").value.trim(),
        rent_amount_words: document
            .getElementById("rent_amount_words")
            .value.trim(),

        payable_date_raw: document.getElementById("payable_date").value.trim(),
        payable_date: document.getElementById("payable_date").value
            ? toOrdinal(parseInt(document.getElementById("payable_date").value, 10))
            : "",

        secu_depo: document.getElementById("secu_depo").value.trim(),
        secu_amount_words: document
            .getElementById("secu_amount_words")
            .value.trim(),
        rent_inc: "",
        rent_rev_number: "",
        "rent_rev year_mon": "",

        pet_text_area: document.getElementById("pet_text_area").value.trim(),
        late_rent: document.getElementById("late_rent").value.trim(),
        late_days: document.getElementById("late_days").value.trim(),

        tenant_clauses: tenantClausesArray,
        landlord_clauses: landlordClausesArray,
        penalty_clauses: penaltyClausesArray,
        misc_clauses: miscClausesArray,

        family: familyArr,
        family_block: familyBlock,

        increase_amount_word: increaseAmountWord,

        checkbox: "",
        add_add_more_clauses: "",
    };

    return data;
}

/**
 * Collects full payload for saving to the database
 * Includes template data plus additional fields for the Tenants sheet
 * @returns {Object} Complete payload for database storage
 */
export function collectFullPayloadForDb() {
    const templateData = collectFormDataForTemplate();
    const selectedUnit = getSelectedUnitForForm();
    return {
        templateData,
        Tenant_occupation: document.getElementById("Tenant_occupation").value.trim(),
        wing: document.getElementById("wing").value.trim(),
        floor_of_building: document.getElementById("floor_of_building").value,
        familyMembers: getFamilyMembersFromTable(),
        unitId: selectedUnit?.unit_id || "",
        unit_number: selectedUnit?.unit_number || templateData.unit_number || "",
        landlordId: templateData.landlord_id || "",
        activeTenant: currentFlow === "addPastTenant" ? "No" : "Yes",
    };
}
