/**
 * Form Feature Module
 * 
 * Handles all form-related logic:
 * - Initializing form dropdowns and options
 * - Collecting form data for template generation
 * - Collecting full payload for database storage
 */

import { floorOptions } from "../../constants.js";
import { buildUnitLabel, formatDateForDoc, numberToIndianWords, toOrdinal } from "../../utils/formatters.js";
import { htmlToMarkedText } from "../../utils/htmlUtils.js";
import { getSelectedClauses } from "../agreements/clauses.js";
import { getFamilyMembersFromTable } from "./family.js";
import { currentFlow } from "../../state.js";
import {
    refreshUnits,
    refreshLandlords,
    getUnitCache as getUnitCacheStore,
    getLandlordCache as getLandlordCacheStore,
} from "../../store/masters.js";

const getValue = (id) => document.getElementById(id)?.value ?? "";
const getTrimmedValue = (id) => getValue(id).trim();

function getLandlordById(landlordId) {
    const landlordCache = getLandlordCacheStore();
    return landlordCache.find((l) => l.landlord_id === landlordId);
}

export function applyLandlordToForm(landlord) {
    const name = document.getElementById("Landlord_name");
    const aadhaar = document.getElementById("landlord_aadhar");
    const address = document.getElementById("landlord_address");
    if (name) name.value = landlord?.name || "";
    if (aadhaar) aadhaar.value = landlord?.aadhaar || "";
    if (address) address.value = landlord?.address || "";
}

function getUnitById(unitId) {
    const unitCache = getUnitCacheStore();
    return unitCache.find((u) => u.unit_id === unitId);
}

export function applyUnitToPremises(unit) {
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
    const landlordCache = getLandlordCacheStore();
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
    const unitCache = getUnitCacheStore();
    const select = document.getElementById("unit_selector");
    if (!select) return;

    const allowOccupied = currentFlow !== "createTenantNew";
    const available = allowOccupied
        ? unitCache
        : unitCache.filter((u) => !u.is_occupied);

    const previous = select.value;
    select.innerHTML = '<option value="">Select a unit to auto-fill details</option>';
    available.forEach((u) => {
        const opt = document.createElement("option");
        opt.value = u.unit_id;
        opt.textContent = buildUnitLabel(u);
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

function getSelectedUnitForForm() {
    const unitId = getValue("unit_selector");
    if (!unitId) return null;
    return getUnitById(unitId);
}

export function getUnitCache() {
    return getUnitCacheStore();
}

export function getLandlordCache() {
    return getLandlordCacheStore();
}

export async function refreshUnitOptions(force = false) {
    const units = await refreshUnits(force);
    populateUnitSelectForFlow();
    return units;
}

export async function refreshLandlordOptions(force = false) {
    const landlords = await refreshLandlords(force);
    populateLandlordSelect();
    return landlords;
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

    const rentRevisionNumber = document.getElementById("rent_rev_number");
    if (rentRevisionNumber) {
        rentRevisionNumber.innerHTML = "";
        for (let i = 0; i <= 31; i++) {
            const opt = document.createElement("option");
            opt.value = String(i);
            opt.textContent = String(i);
            rentRevisionNumber.appendChild(opt);
        }
    }

    // Notice period dropdowns (1-12 months)
    const noticeTenant = document.getElementById("notice_num_t");
    const noticeLandlord = document.getElementById("notice_num_l");

    for (let i = 1; i <= 12; i++) {
        const label = i.toString();
        if (noticeTenant) {
            const opt2 = new Option(label, label);
            noticeTenant.appendChild(opt2);
        }
        if (noticeLandlord) {
            const opt3 = new Option(label, label);
            noticeLandlord.appendChild(opt3);
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

    document.addEventListener("units:updated", populateUnitSelectForFlow);
    document.addEventListener("flow:changed", populateUnitSelectForFlow);
    document.addEventListener("landlords:updated", () => populateLandlordSelect());
}

/**
 * Collects all form data and formats it for the DOCX template
 * @returns {Object} Formatted data object for template generation
 */
export function collectFormDataForTemplate() {
    const agreementDate = getValue("agreement_date");
    const tenancyComm = getValue("tenancy_comm");
    const tenancyEnd = getValue("tenancy_end");
    const selectedUnit = getSelectedUnitForForm();
    const landlordId = getValue("landlord_selector");
    const selectedLandlord = landlordId ? getLandlordById(landlordId) : null;

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
                `- Aadhaar: ${f.aadhaar || "-"}, Address: ${f.address || "-"}`
        )
        .join("\n");

    // Calculate rent increase amount in words
    const rentIncreaseRaw = getTrimmedValue("rent_inc");
    const rentIncreaseVal = parseInt(rentIncreaseRaw, 10);
    const increaseAmountWord =
        !isNaN(rentIncreaseVal) && rentIncreaseVal > 0
            ? `${numberToIndianWords(rentIncreaseVal)} only`
            : getTrimmedValue("increase_amount_word");
    const rentRevisionUnitRaw = getTrimmedValue("rent_rev_unit");
    const rentRevisionUnitForDoc =
        rentRevisionUnitRaw === "YEAR"
            ? "year(s)"
            : rentRevisionUnitRaw === "MONTH"
                ? "month(s)"
                : rentRevisionUnitRaw;

    const landlordName = selectedLandlord?.name ?? getValue("Landlord_name");
    const landlordAddress = selectedLandlord?.address ?? getValue("landlord_address");
    const landlordAadhaar = selectedLandlord?.aadhaar ?? getValue("landlord_aadhar");
    const payableDateRaw = getTrimmedValue("payable_date");

    const data = {
        "GRN number": getTrimmedValue("grn_number"),
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
        notice_num_t: getTrimmedValue("notice_num_t"),
        notice_num_l: getTrimmedValue("notice_num_l"),

        Tenant_Full_Name: getTrimmedValue("Tenant_Full_Name"),
        Tenant_Permanent_Address: getTrimmedValue("Tenant_Permanent_Address"),
        tenant_Aadhar: getTrimmedValue("tenant_Aadhar"),
        tenant_mobile: getTrimmedValue("tenant_mobile"),
        unit_id: selectedUnit?.unit_id || "",
        unit_number: selectedUnit?.unit_number || "",

        "floor_of_building": getValue("floor_of_building"),
        direction_build: getValue("direction_build"),
        meter_number: getTrimmedValue("meter_number"),

        rent_amount: getTrimmedValue("rent_amount"),
        rent_amount_words: getTrimmedValue("rent_amount_words"),

        payable_date_raw: payableDateRaw,
        payable_date: payableDateRaw ? toOrdinal(parseInt(payableDateRaw, 10)) : "",

        secu_depo: getTrimmedValue("secu_depo"),
        secu_amount_words: getTrimmedValue("secu_amount_words"),
        rent_inc: getTrimmedValue("rent_inc"),
        rent_rev_number: getTrimmedValue("rent_rev_number"),
        "rent_rev year_mon": rentRevisionUnitForDoc,

        pet_text_area: getTrimmedValue("pet_text_area"),
        late_rent: getTrimmedValue("late_rent"),
        late_days: getTrimmedValue("late_days"),

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
    const sanitizedTemplate = { ...templateData };
    sanitizedTemplate["rent_rev year_mon"] = getTrimmedValue("rent_rev_unit");
    delete sanitizedTemplate.rent_amount_words;
    delete sanitizedTemplate.secu_amount_words;
    const selectedUnit = getSelectedUnitForForm();
    return {
        templateData: sanitizedTemplate,
        Tenant_occupation: getTrimmedValue("Tenant_occupation"),
        wing: getTrimmedValue("wing"),
        floor_of_building: getValue("floor_of_building"),
        familyMembers: getFamilyMembersFromTable(),
        unitId: selectedUnit?.unit_id || "",
        unit_number: selectedUnit?.unit_number || templateData.unit_number || "",
        landlordId: templateData.landlord_id || "",
        activeTenant: "Yes",
    };
}
