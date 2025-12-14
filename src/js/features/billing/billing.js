/**
 * Billing Feature Module
 *
 * Calculates monthly charges, renders bill previews, and synchronizes
 * billing metadata with Google Sheets. This module keeps the billing
 * state, parses tenant/unit data, and wires modal interactions.
 */

import { numberToIndianWords } from "../../utils/formatters.js";
import { hideModal, showModal, showToast, smoothToggle } from "../../utils/ui.js";
import { ensureTenantDirectoryLoaded, getActiveTenantsForWing } from "../tenants/tenants.js";
import { fetchBillingRecord, fetchGeneratedBills, saveBillingRecord } from "../../api/sheets.js";

const billingState = {
    selectedMonthKey: null,
    selectedMonthLabel: "",
    selectedWing: "",
    selectedWingNormalized: "",
    selectedWingLabel: "",
    lastGeneratedSummaries: [],
    sendStatus: new Map(),
    calendarCoverage: new Map(),
    availableWings: [],
    coverageLoaded: false,
    motorSnapshot: null,
    meta: {
        electricityRate: "",
        sweepingPerFlat: "",
        motorPrev: "",
        motorNew: "",
    },
    tenants: [],
};

const billingRecordCache = new Map();

function getBillingCacheKey(monthKey, wing) {
    const normalizedMonth = normalizeMonthKey(monthKey);
    const normalizedWing = normalizeWing(wing);
    return `${normalizedMonth || ""}__${normalizedWing || ""}`;
}

function getWingVariants(rawWing) {
    const canonical = (rawWing || "").toString().trim();
    const normalized = normalizeWing(canonical);
    if (!canonical) return [];
    if (canonical.toLowerCase() === normalized) return [canonical];
    return [canonical, normalized];
}

function normalizeMonthKey(value) {
    if (!value) return "";

    if (value instanceof Date && !Number.isNaN(value)) {
        const month = `${value.getUTCMonth() + 1}`.padStart(2, "0");
        return `${value.getUTCFullYear()}-${month}`;
    }

    const str = value.toString().trim();
    if (!str) return "";

    // Excel / Sheets serial numbers (roughly covers 1990-2150 ranges)
    const numericValue = Number(str);
    if (!Number.isNaN(numericValue) && /^\d+(?:\.0+)?$/.test(str) && numericValue > 30000) {
        const excelEpoch = Date.UTC(1899, 11, 30);
        const utcDate = new Date(excelEpoch + numericValue * 24 * 60 * 60 * 1000);
        if (!Number.isNaN(utcDate)) {
            const month = `${utcDate.getUTCMonth() + 1}`.padStart(2, "0");
            return `${utcDate.getUTCFullYear()}-${month}`;
        }
    }

    const compact = str.match(/^(\d{4})(\d{2})$/);
    if (compact) {
        return `${compact[1]}-${compact[2]}`;
    }

    const ymd = str.match(/^(\d{4})[-/.](\d{1,2})/);
    if (ymd) {
        const month = `${ymd[2]}`.padStart(2, "0");
        return `${ymd[1]}-${month}`;
    }

    const my = str.match(/^(\d{1,2})[-/.](\d{4})$/);
    if (my) {
        const month = `${my[1]}`.padStart(2, "0");
        return `${my[2]}-${month}`;
    }

    const mdy = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
    if (mdy) {
        const rawYear = mdy[3].length === 2 ? `20${mdy[3]}` : mdy[3];
        const first = parseInt(mdy[1], 10);
        const second = parseInt(mdy[2], 10);
        const monthNum = first > 12 && second <= 12 ? second : first;
        const month = `${monthNum}`.padStart(2, "0");
        return `${rawYear}-${month}`;
    }

    const monthName = str.match(/^([A-Za-z]{3,})\s+(\d{2,4})$/);
    if (monthName) {
        const monthIdx = new Date(`${monthName[1]} 1, 2000`).getMonth();
        if (!Number.isNaN(monthIdx)) {
            const month = `${monthIdx + 1}`.padStart(2, "0");
            const year = monthName[2].length === 2 ? `20${monthName[2]}` : monthName[2];
            return `${year}-${month}`;
        }
    }

    const parsedDate = new Date(str);
    if (!Number.isNaN(parsedDate)) {
        const month = `${parsedDate.getUTCMonth() + 1}`.padStart(2, "0");
        return `${parsedDate.getUTCFullYear()}-${month}`;
    }

    return str;
}

function normalizeWing(value) {
    return (value || "").toString().trim().toLowerCase();
}

function getSelectedWingNormalized() {
    return billingState.selectedWingNormalized || normalizeWing(billingState.selectedWing);
}

function getCanonicalWingValue(rawWing) {
    const wingSelect = document.getElementById("wing");
    const targetWing = (rawWing || "").toString().trim();
    if (!wingSelect || !targetWing) return targetWing;
    const match = Array.from(wingSelect.options).find(
        (opt) => normalizeWing(opt.value) === normalizeWing(targetWing)
    );
    return match ? match.value : targetWing;
}

function normalizeMetaPayload(meta = {}) {
    return {
        electricityRate: meta.electricityRate ?? meta.electricity_rate ?? "",
        sweepingPerFlat: meta.sweepingPerFlat ?? meta.sweeping_per_flat ?? "",
        motorPrev: meta.motorPrev ?? meta.motor_prev ?? "",
        motorNew: meta.motorNew ?? meta.motor_new ?? "",
    };
}

function hasAnyMetaValue(meta = {}) {
    return [meta.electricityRate, meta.sweepingPerFlat, meta.motorPrev, meta.motorNew].some((v) =>
        v !== undefined && v !== null && v !== ""
    );
}

function isTenantIncluded(tenant) {
    const flag = tenant?.included;
    if (flag === false) return false;
    if (typeof flag === "string" && flag.toLowerCase() === "false") return false;
    return true;
}

function getIncludedTenants() {
    return billingState.tenants.filter((t) => isTenantIncluded(t));
}

function getLastTwelveMonths() {
    const months = [];
    const now = new Date();
    for (let i = 12; i >= 1; i -= 1) {
        const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
        const monthNumber = `${d.getMonth() + 1}`.padStart(2, "0");
        months.push({
            key: `${d.getFullYear()}-${monthNumber}`,
            label: d.toLocaleString("default", { month: "long", year: "numeric" }),
        });
    }
    return months;
}

function getAvailableWings() {
    const select = document.getElementById("wing");
    if (!select) return [];

    return Array.from(select.options)
        .map((opt) => opt.value)
        .filter(Boolean)
        .map((w) => normalizeWing(w));
}

function getPreviousMonthKey(currentKey) {
    if (!currentKey) return null;
    const [year, month] = currentKey.split("-").map((v) => parseInt(v, 10));
    if (!year || !month) return null;
    const prevDate = new Date(year, month - 2, 1);
    const prevMonth = `${prevDate.getMonth() + 1}`.padStart(2, "0");
    return `${prevDate.getFullYear()}-${prevMonth}`;
}

function getNextMonthLabel(currentKey) {
    if (!currentKey) return "";
    const [year, month] = currentKey.split("-").map((v) => parseInt(v, 10));
    if (!year || !month) return "";
    const nextDate = new Date(year, month, 1);
    return nextDate.toLocaleString("default", { month: "long" });
}

function calculateChargesForTenant(tenant, motorPerTenant) {
    const rate = parseNumber(billingState.meta.electricityRate, true);
    const sweep = parseNumber(billingState.meta.sweepingPerFlat, true);
    const rent = roundToTwo(parseNumber(tenant.rentAmount, true));
    const prevVal = parseNumber(tenant.prevReading, false);
    const newVal = parseNumber(tenant.newReading, false);
    const units = Math.max(newVal - prevVal, 0);
    const isIncluded = isTenantIncluded(tenant);
    const electricity = isIncluded ? roundToTwo(units * rate) : 0;
    const motorShare = isIncluded ? roundToTwo(motorPerTenant ?? computeMotorShare().perTenant) : 0;
    const sweepAmount = isIncluded ? roundToTwo(sweep) : 0;
    const totalBeforeRound = isIncluded ? rent + electricity + motorShare + sweepAmount : 0;
    const total = roundToNearest(totalBeforeRound);

    return {
        units,
        electricity,
        motorShare,
        sweepAmount,
        total,
    };
}

function parseNumber(val, allowDecimal = false) {
    if (val === "" || val === null || val === undefined) return 0;
    const parsed = allowDecimal ? parseFloat(val) : parseInt(val, 10);
    return Number.isNaN(parsed) ? 0 : parsed;
}

function roundToTwo(val) {
    const num = parseFloat(val ?? 0);
    if (Number.isNaN(num)) return 0;
    return Math.round(num * 100) / 100;
}

function roundToNearest(val) {
    const num = parseFloat(val ?? 0);
    if (Number.isNaN(num)) return 0;
    return roundToTwo(Math.round(num));
}

function formatCurrency(amount) {
    const numeric = roundToTwo(amount || 0);
    if (Number.isNaN(numeric)) return "₹0";
    return `₹${numeric.toFixed(2)}`;
}

function updateSelectionChips() {
    const monthChip = document.getElementById("billingSelectedMonth");
    const wingChip = document.getElementById("billingSelectedWing");
    const displayWing = billingState.selectedWing ? billingState.selectedWing.toUpperCase() : "Pick wing";
    if (monthChip) monthChip.textContent = billingState.selectedMonthLabel || "Select month";
    if (wingChip) wingChip.textContent = displayWing;
}

function setStep(step) {
    const selectWingStep = document.getElementById("billingStepSelectWing");
    const detailsStep = document.getElementById("billingStepDetails");
    if (selectWingStep) selectWingStep.classList.toggle("hidden", step !== "wing");
    if (detailsStep) detailsStep.classList.toggle("hidden", step !== "details");
}

function resetBillingForm() {
    billingState.meta = {
        electricityRate: "",
        sweepingPerFlat: "",
        motorPrev: "",
        motorNew: "",
    };
    billingState.tenants = [];
    const tenantBody = document.getElementById("billingTenantTableBody");
    if (tenantBody) tenantBody.innerHTML = "";
    const emptyState = document.getElementById("billingTenantEmpty");
    if (emptyState) emptyState.classList.add("hidden");
}

function openBillingModal(month) {
    const modal = document.getElementById("billingWingModal");
    const monthText = document.getElementById("billingModalMonthLabel");
    const wingSelect = document.getElementById("billingWingSelect");
    billingState.selectedMonthKey = normalizeMonthKey(month.key);
    billingState.selectedMonthLabel = month.label;
    billingState.selectedWing = "";
    billingState.selectedWingNormalized = "";
    resetBillingForm();
    setStep("wing");

    if (monthText) monthText.textContent = month.label;
    if (wingSelect) wingSelect.value = "";
    updateSelectionChips();
    if (modal) showModal(modal);
}

function buildMonthCoverage(bills = []) {
    const coverage = new Map();
    bills.forEach((bill) => {
        const monthKey = normalizeMonthKey(
            bill.monthKey || bill.month_key || bill.month || bill.monthLabel || bill.month_label
        );
        const wing = normalizeWing(bill.wing);
        if (!monthKey || !wing) return;
        if (!coverage.has(monthKey)) {
            coverage.set(monthKey, new Set());
        }
        coverage.get(monthKey).add(wing);
    });
    return coverage;
}

function getMonthGenerationStatus(monthKey) {
    const totalWings = billingState.availableWings.length;
    const normalizedMonth = normalizeMonthKey(monthKey);
    const coveredWings = billingState.calendarCoverage.get(normalizedMonth)?.size || 0;

    if (!billingState.coverageLoaded || !totalWings) {
        return { key: "unknown", label: "Checking status...", textClass: "text-slate-500" };
    }

    if (coveredWings === 0) {
        return { key: "none", label: "No bills generated", textClass: "text-rose-700" };
    }

    if (coveredWings >= totalWings) {
        return { key: "full", label: "All wings billed", textClass: "text-emerald-700" };
    }

    return {
        key: "partial",
        label: `${coveredWings}/${totalWings} wings billed`,
        textClass: "text-amber-700",
    };
}

function markCoverageForSelection() {
    const monthKey = normalizeMonthKey(billingState.selectedMonthKey);
    const wing = getSelectedWingNormalized();

    if (!monthKey || !wing) return;

    if (!billingState.availableWings.length) {
        billingState.availableWings = getAvailableWings();
    }

    const coverage = billingState.calendarCoverage.get(monthKey) || new Set();
    coverage.add(wing);
    billingState.calendarCoverage.set(monthKey, coverage);
    billingState.coverageLoaded = true;
    renderBillingCalendar();
}

function renderBillingCalendar() {
    const grid = document.getElementById("billingMonthsGrid");
    if (!grid) return;

    if (!billingState.availableWings.length) {
        billingState.availableWings = getAvailableWings();
    }

    grid.innerHTML = "";
    const months = getLastTwelveMonths();

    months.forEach((month) => {
        const card = document.createElement("button");
        card.type = "button";
        const status = getMonthGenerationStatus(month.key);

        const baseClasses =
            "w-full aspect-[4/3] max-h-28 rounded-xl border bg-gradient-to-br shadow-sm hover:shadow-lg transition transform hover:-translate-y-1 flex items-center justify-center text-center p-1.5";
        const statusClasses = {
            full: "border-emerald-200 from-emerald-50 to-white",
            partial: "border-amber-200 from-amber-50 to-white",
            none: "border-rose-200 from-rose-50 to-white",
            unknown: "border-slate-200 from-slate-50 to-white",
        };

        card.className = `${baseClasses} ${statusClasses[status.key] || statusClasses.unknown}`;

        card.innerHTML = `
            <div class="space-y-1">
                <p class="text-sm md:text-base font-semibold text-slate-800">${month.label}</p>
                <p class="text-[11px] font-semibold ${status.textClass}">${status.label}</p>
            </div>
        `;

        card.addEventListener("click", () => openBillingModal(month));
        grid.appendChild(card);
    });
}

async function refreshBillingCalendarCoverage() {
    const { bills, coverage } = await fetchGeneratedBills();
    billingState.availableWings = getAvailableWings();
    const coverageSource = Array.isArray(coverage) && coverage.length ? coverage : bills;
    billingState.calendarCoverage = buildMonthCoverage(Array.isArray(coverageSource) ? coverageSource : []);
    billingState.coverageLoaded = true;
    renderBillingCalendar();
}

function computeMotorShare() {
    const rate = parseNumber(billingState.meta.electricityRate, true);
    const motorPrev = parseNumber(billingState.meta.motorPrev, false);
    const motorNew = parseNumber(billingState.meta.motorNew, false);
    const units = Math.max(motorNew - motorPrev, 0);
    const cost = roundToTwo(units * rate);
    const count = getIncludedTenants().length;
    return { units, cost, perTenant: count ? roundToTwo(cost / count) : 0 };
}

function renderMotorSummary() {
    const summary = document.getElementById("motorUnitsSummary");
    const share = document.getElementById("motorShareSummary");
    const { units, cost, perTenant } = computeMotorShare();
    if (summary) summary.textContent = `${units.toFixed(0)} units (${formatCurrency(cost)})`;
    const includedCount = getIncludedTenants().length;
    if (share) share.textContent = `${formatCurrency(perTenant)} per selected tenant (${includedCount || 0})`;
}

function renderTenantTable() {
    const tbody = document.getElementById("billingTenantTableBody");
    const emptyState = document.getElementById("billingTenantEmpty");
    if (!tbody || !emptyState) return;

    tbody.innerHTML = "";
    if (!billingState.tenants.length) {
        emptyState.classList.remove("hidden");
        return;
    }
    emptyState.classList.add("hidden");

    const motor = computeMotorShare();
    const motorPerTenant = motor.perTenant;

    const fragment = document.createDocumentFragment();
    billingState.tenants.forEach((tenant, idx) => {
        const charges = calculateChargesForTenant(tenant, motorPerTenant);
        const prevVal = parseNumber(tenant.prevReading, false);
        const newVal = parseNumber(tenant.newReading, false);
        const invalidReading = newVal < prevVal;
        const isIncluded = isTenantIncluded(tenant);
        const isConfirmed = tenant.hasBill;
        const rowHighlight = invalidReading ? "bg-rose-50" : isConfirmed ? "bg-emerald-50" : "";

        const tr = document.createElement("tr");
        tr.className = `border-b last:border-0 hover:bg-slate-50 ${rowHighlight}`;
        tr.dataset.grn = tenant.grn || "";
        tr.innerHTML = `
            <td class="px-2 py-2 text-center text-xs align-middle w-8">
                ${
                    isConfirmed
                        ? ""
                        : `<input type="checkbox" class="tenant-include h-4 w-4" data-index="${idx}" ${
                              isIncluded ? "checked" : ""
                          } />`
                }
            </td>
            <td class="px-2 py-2 text-xs">
                <div class="font-semibold text-[12px] leading-tight">${tenant.name}</div>
            </td>
            <td class="px-2 py-2 text-xs font-semibold">${formatCurrency(tenant.rentAmount)}</td>
            <td class="px-2 py-2 text-xs"><input type="number" inputmode="numeric" step="1" min="0" pattern="[0-9]*" class="w-full border rounded px-2 py-1 text-xs tenant-prev input-no-spinner ${invalidReading ? "border-rose-500 bg-rose-50 text-rose-700" : "border-slate-200"}" value="${tenant.prevReading ?? ""}" /></td>
            <td class="px-2 py-2 text-xs"><input type="number" inputmode="numeric" step="1" min="0" pattern="[0-9]*" class="w-full border rounded px-2 py-1 text-xs tenant-new input-no-spinner ${invalidReading ? "border-rose-500 bg-rose-50 text-rose-700" : "border-slate-200"}" value="${tenant.newReading ?? ""}" /></td>
            <td class="px-2 py-2 text-xs font-semibold text-indigo-700">${formatCurrency(charges.electricity)}</td>
            <td class="px-2 py-2 text-xs">${formatCurrency(charges.motorShare)}</td>
            <td class="px-2 py-2 text-xs">${formatCurrency(charges.sweepAmount)}</td>
            <td class="px-2 py-2 text-xs font-bold text-slate-800">${formatCurrency(charges.total)}</td>
            <td class="px-2 py-2 text-xs text-right">
                ${
                    isConfirmed
                        ? '<span class="inline-flex items-center gap-1 text-emerald-700 font-semibold">✓</span>'
                        : ""
                }
            </td>
        `;

        const prevInput = tr.querySelector(".tenant-prev");
        const newInput = tr.querySelector(".tenant-new");
        if (prevInput) {
            prevInput.addEventListener("change", (e) => {
                tenant.prevReading = e.target.value;
                renderTenantTable();
                renderMotorSummary();
            });
        }
        if (newInput) {
            newInput.addEventListener("change", (e) => {
                tenant.newReading = e.target.value;
                renderTenantTable();
                renderMotorSummary();
            });
        }

        const includeCheckbox = tr.querySelector(".tenant-include");
        if (includeCheckbox) {
            includeCheckbox.addEventListener("change", (e) => {
                tenant.included = e.target.checked;
                tenant.hasBill = false;
                renderTenantTable();
                renderMotorSummary();
            });
        }

        fragment.appendChild(tr);
    });

    tbody.appendChild(fragment);
}

function populateInputsFromState() {
    const rateInput = document.getElementById("billingElectricityRate");
    const sweepInput = document.getElementById("billingSweepRate");
    const motorPrevInput = document.getElementById("billingMotorPrev");
    const motorNewInput = document.getElementById("billingMotorNew");

    if (rateInput) rateInput.value = billingState.meta.electricityRate;
    if (sweepInput) sweepInput.value = billingState.meta.sweepingPerFlat;
    if (motorPrevInput) motorPrevInput.value = billingState.meta.motorPrev;
    if (motorNewInput) motorNewInput.value = billingState.meta.motorNew;

    applyMotorValidation();
}

function applyMotorValidation() {
    const motorPrevInput = document.getElementById("billingMotorPrev");
    const motorNewInput = document.getElementById("billingMotorNew");
    if (!motorPrevInput || !motorNewInput) return;

    const motorPrevVal = parseNumber(billingState.meta.motorPrev, false);
    const motorNewVal = parseNumber(billingState.meta.motorNew, false);
    const invalid = motorNewVal < motorPrevVal;

    const base = "w-full border rounded px-3 py-2 text-[12px]";
    const invalidClass = "border-rose-500 bg-rose-50 text-rose-700";

    motorPrevInput.className = `${base} ${invalid ? invalidClass : "border-slate-200"}`;
    motorNewInput.className = `${base} ${invalid ? invalidClass : "border-slate-200"}`;
}

function applyMetaListeners() {
    const rateInput = document.getElementById("billingElectricityRate");
    const sweepInput = document.getElementById("billingSweepRate");
    const motorPrevInput = document.getElementById("billingMotorPrev");
    const motorNewInput = document.getElementById("billingMotorNew");

    [
        [rateInput, "electricityRate", true],
        [sweepInput, "sweepingPerFlat", true],
        [motorPrevInput, "motorPrev", false],
        [motorNewInput, "motorNew", false],
    ].forEach(([input, key]) => {
        if (!input) return;
        input.addEventListener("input", (e) => {
            billingState.meta[key] = e.target.value;
            renderTenantTable();
            renderMotorSummary();
            applyMotorValidation();
        });
    });
}

function normalizeTenantKey(raw) {
    return (raw || "").toString().trim().toLowerCase();
}

function getTenantKeyCandidates(raw) {
    const keys = new Set();
    const identity = getTenantIdentityKey(raw);
    const wing = normalizeWing(raw?.wing || "");
    const unit = normalizeTenantKey(raw?.unitId || raw?.unit_id);
    const name = normalizeTenantKey(raw?.tenantFullName || raw?.name || raw?.tenant_name);
    const grn = normalizeTenantKey(raw?.grnNumber || raw?.grn_number || raw?.grn);

    if (identity) keys.add(identity);
    if (grn) keys.add(grn);
    if (unit) keys.add(unit);
    if (name) keys.add(name);
    if (wing && unit) keys.add(`${wing}|${unit}`);
    if (wing && name) keys.add(`${wing}|${name}`);
    if (wing && unit && name) keys.add(`${wing}|${unit}|${name}`);

    return Array.from(keys).filter(Boolean);
}

function normalizeIncludedFlag(flag) {
    if (flag === false) return false;
    if (typeof flag === "string" && flag.toLowerCase() === "false") return false;
    return true;
}

function getTenantIdentityKey(raw) {
    const candidates = [
        raw?.tenantKey,
        raw?.tenancyId,
        raw?.tenancy_id,
        raw?.grnNumber,
        raw?.grn_number,
        raw?.grn,
        raw?.name,
        raw?.tenantFullName,
        raw?.tenant_name,
    ];

    for (const candidate of candidates) {
        const key = normalizeTenantKey(candidate);
        if (key) return key;
    }

    return "";
}

function mergeTenantData(activeTenants, savedEntries = [], previousEntries = []) {
    const buildLookup = (entries) => {
        const lookup = new Map();
        const canonical = new Map();

        entries.forEach((item, idx) => {
            const candidateKeys = getTenantKeyCandidates(item);
            if (!candidateKeys.length) return;
            const identityKey = getTenantIdentityKey(item) || normalizeTenantKey(item.grn || item.name);
            const enriched = { ...item, tenantKey: identityKey };
            canonical.set(idx, enriched);
            candidateKeys.forEach((key) => {
                if (!lookup.has(key)) lookup.set(key, idx);
            });
        });

        return { lookup, canonical };
    };

    const { lookup: savedLookup, canonical: savedCanonical } = buildLookup(savedEntries);
    const { lookup: previousLookup, canonical: previousCanonical } = buildLookup(previousEntries);

    const activeMap = new Map();
    activeTenants.forEach((tenant) => {
        const keys = getTenantKeyCandidates(tenant);
        const identityKey = keys[0] || normalizeTenantKey(`${tenant.wing || ""}-${tenant.unitId || ""}-${tenant.tenantFullName || tenant.name || ""}`);
        if (identityKey && !activeMap.has(identityKey)) {
            activeMap.set(identityKey, { ...tenant, tenantKey: identityKey, __keys: keys });
        }
    });

    const merged = [];

    activeMap.forEach((tenant, mapKey) => {
        const keys = tenant.__keys || [mapKey];
        let saved = null;
        for (const key of keys) {
            const savedIdx = savedLookup.get(key);
            if (savedIdx !== undefined) {
                saved = savedCanonical.get(savedIdx);
                break;
            }
        }

        let previous = null;
        if (!saved) {
            for (const key of keys) {
                const prevIdx = previousLookup.get(key);
                if (prevIdx !== undefined) {
                    previous = previousCanonical.get(prevIdx);
                    break;
                }
            }
        }
        const savedPrev = saved?.prevReading;
        const savedNew = saved?.newReading;
        const previousPrev = previous?.newReading ?? previous?.prevReading ?? "";
        const included = saved && saved.included !== undefined ? normalizeIncludedFlag(saved.included) : true;

        merged.push({
            tenantKey: tenant.tenantKey || mapKey,
            grn: tenant.grnNumber || tenant.grn || "",
            name: tenant.tenantFullName || tenant.name || "Unnamed",
            tenancyId: tenant.tenancyId || "",
            unitId: tenant.unitId || "",
            rentAmount: tenant.rentAmount || saved?.rentAmount || saved?.rent_amount || 0,
            mobile: tenant.tenantMobile || tenant.mobile || tenant.phone || "",
            prevReading:
                savedPrev !== undefined && savedPrev !== null && savedPrev !== ""
                    ? savedPrev
                    : previousPrev,
            newReading: savedNew ?? "",
            included,
            hasBill: !!saved,
            payableDate: saved?.payableDate || tenant.payableDate || "",
        });
    });

    return merged;
}

function extractBillingRecordPayload(response) {
    if (!response) return null;
    if (response.record) return response.record;
    if (response.data) return response.data;
    if (response.meta || response.tenants) return response;
    return null;
}

async function getBillingRecordCached(monthKey, wing, { force = false } = {}) {
    const normalizedMonth = normalizeMonthKey(monthKey);
    const wingVariants = getWingVariants(wing);
    if (!normalizedMonth || !wingVariants.length) return null;

    let lastResponse = null;
    for (const variant of wingVariants) {
        const normalizedWing = normalizeWing(variant);
        const key = getBillingCacheKey(normalizedMonth, normalizedWing);
        if (!force && billingRecordCache.has(key)) return billingRecordCache.get(key);

        const response = await fetchBillingRecord(normalizedMonth, variant);
        lastResponse = response;
        const record = extractBillingRecordPayload(response);
        if (record) {
            const normalizedMeta = normalizeMetaPayload(record.meta || {});
            const hasConfig = record.hasConfig ?? hasAnyMetaValue(normalizedMeta);
            const hasReadings = record.hasReadings ?? (Array.isArray(record.tenants) && record.tenants.length > 0);
            const wrapped = {
                ...response,
                record: {
                    ...record,
                    meta: normalizedMeta,
                    hasConfig,
                    hasReadings,
                    monthKey: normalizeMonthKey(record.monthKey || record.month_key || normalizedMonth),
                    wing: normalizeWing(record.wing || wing),
                },
            };
            wingVariants.forEach((v) => {
                const variantKey = getBillingCacheKey(normalizedMonth, normalizeWing(v));
                billingRecordCache.set(variantKey, wrapped);
            });
            return wrapped;
        }
    }

    return lastResponse;
}

async function loadBillingData({ force = false } = {}) {
    const loader = document.getElementById("billingDetailsLoader");
    const form = document.getElementById("billingDetailsForm");
    if (loader) loader.classList.remove("hidden");
    if (form) form.classList.add("opacity-50");

    const loadTenantsPromise = ensureTenantDirectoryLoaded();
    const normalizedMonth = normalizeMonthKey(billingState.selectedMonthKey);
    const selectedWing = billingState.selectedWing || billingState.selectedWingNormalized;
    const prevKey = getPreviousMonthKey(normalizedMonth);
    const [_, current, previous] = await Promise.all([
        loadTenantsPromise,
        getBillingRecordCached(normalizedMonth, selectedWing, { force }),
        prevKey ? getBillingRecordCached(prevKey, selectedWing, { force }) : Promise.resolve(null),
    ]);

    const normalizedWing = getSelectedWingNormalized();
    const activeTenants = getActiveTenantsForWing(normalizedWing);
    const previousMeta = normalizeMetaPayload(previous?.record?.meta || {});
    const previousEntriesRaw = Array.isArray(previous?.record?.tenants) ? previous.record.tenants : [];
    const previousHasReadings = previous?.record?.hasReadings ?? previousEntriesRaw.length > 0;
    const previousEntries = previousHasReadings ? previousEntriesRaw : [];

    const record = current?.record;
    const currentMeta = normalizeMetaPayload(record?.meta || {});
    const hasCurrentConfig = record?.hasConfig ?? hasAnyMetaValue(currentMeta);
    const hasCurrentReadings = record?.hasReadings ?? (Array.isArray(record?.tenants) && record.tenants.length > 0);
    const savedEntries = hasCurrentReadings && Array.isArray(record?.tenants) ? record.tenants : [];

    billingState.meta = {
        electricityRate: currentMeta.electricityRate || previousMeta.electricityRate || "",
        sweepingPerFlat: currentMeta.sweepingPerFlat || previousMeta.sweepingPerFlat || "",
        motorPrev: currentMeta.motorPrev || previousMeta.motorNew || previousMeta.motorPrev || "",
        motorNew: currentMeta.motorNew || "",
    };

    if (!hasCurrentConfig && !billingState.meta.motorPrev && previousMeta.motorNew) {
        billingState.meta.motorPrev = previousMeta.motorNew;
    }

    billingState.tenants = mergeTenantData(activeTenants, savedEntries, previousEntries);
    populateInputsFromState();
    renderMotorSummary();
    renderTenantTable();

    if (loader) loader.classList.add("hidden");
    if (form) form.classList.remove("opacity-50");
}

async function handleNextStep() {
    const wingSelect = document.getElementById("billingWingSelect");
    if (!wingSelect || !wingSelect.value) {
        showToast("Please select a wing to continue", "error");
        return;
    }
    const canonicalWing = getCanonicalWingValue(wingSelect.value);
    billingState.selectedWing = canonicalWing;
    billingState.selectedWingNormalized = normalizeWing(canonicalWing);
    billingState.selectedWingLabel = canonicalWing;
    updateSelectionChips();
    setStep("details");
    await loadBillingData();
}

function closeBillingModal() {
    const modal = document.getElementById("billingWingModal");
    if (modal) hideModal(modal);
}

async function handleGenerateAndPrompt() {
    const saved = await handleSaveBills();
    if (saved) {
        closeBillingModal();
        toggleSendPrompt(true);
    }
}

async function handleSaveBills() {
    if (!billingState.selectedMonthKey || !billingState.selectedWing) {
        showToast("Pick a month and wing first", "error");
        return false;
    }

    const selectedTenants = getIncludedTenants();
    if (!selectedTenants.length) {
        showToast("Select at least one tenant to generate bills", "error");
        return false;
    }

    const motor = computeMotorShare();
    const motorPerTenant = motor.perTenant;

    const normalizedMonthKey = normalizeMonthKey(billingState.selectedMonthKey);
    const normalizedWing = getSelectedWingNormalized();
    const payload = {
        monthKey: normalizedMonthKey,
        monthLabel: billingState.selectedMonthLabel,
        wing: normalizedWing,
        wingLabel: billingState.selectedWingLabel || billingState.selectedWing,
        meta: { ...billingState.meta },
        tenants: billingState.tenants.map((t) => ({
            tenantKey: getTenantIdentityKey(t),
            grn: t.grn,
            name: t.name,
            tenancyId: t.tenancyId,
            rentAmount: roundToTwo(t.rentAmount),
            prevReading: roundToTwo(t.prevReading),
            newReading: roundToTwo(t.newReading),
            payableDate: t.payableDate,
            included: isTenantIncluded(t),
            ...(() => {
                const isIncluded = isTenantIncluded(t);
                const charges = isIncluded
                    ? calculateChargesForTenant(t, motorPerTenant)
                    : {
                          electricity: 0,
                          motorShare: 0,
                          sweepAmount: 0,
                          total: 0,
                      };
                return {
                    electricityAmount: charges.electricity,
                    motorShare: charges.motorShare,
                    sweepAmount: charges.sweepAmount,
                    totalAmount: charges.total,
                };
            })(),
        })),
    };

    const res = await saveBillingRecord(payload);
    if (res?.ok) {
        showToast("Bills saved to Google Sheets", "success");
        billingState.tenants = billingState.tenants.map((t) => ({
            ...t,
            hasBill: isTenantIncluded(t),
        }));
        billingState.motorSnapshot = {
            units: motor.units,
            rate: parseNumber(billingState.meta.electricityRate, true),
            prev: parseNumber(billingState.meta.motorPrev, false),
            next: parseNumber(billingState.meta.motorNew, false),
            includedCount: selectedTenants.length,
        };
        const normalizedWingKey = normalizeWing(payload.wing || billingState.selectedWing);
        const cacheKey = getBillingCacheKey(normalizedMonthKey, normalizedWingKey);
        const cachedRecord = {
            record: {
                ...payload,
                monthKey: normalizedMonthKey,
                wing: normalizedWingKey,
                hasConfig: true,
                hasReadings: true,
            },
        };
        billingRecordCache.set(cacheKey, cachedRecord);
        billingState.lastGeneratedSummaries = selectedTenants.map((t) => {
            const charges = calculateChargesForTenant(t, motorPerTenant);
            return {
                id: t.grn || t.name,
                name: t.name,
                mobile: getPrimaryMobile(t.mobile),
                rent: roundToTwo(parseNumber(t.rentAmount, true)),
                prevReading: roundToTwo(parseNumber(t.prevReading, false)),
                newReading: roundToTwo(parseNumber(t.newReading, false)),
                electricity: charges.electricity,
                motorShare: charges.motorShare,
                motorUnits: motor.units,
                sweep: charges.sweepAmount,
                units: charges.units,
                total: charges.total,
                payableDay: t.payableDate || "",
            };
        });
        // Refresh coverage directly from the latest saved bills to keep calendar accurate
        const coverageSource = Array.isArray(res.coverage) && res.coverage.length ? res.coverage : res.bills;
        if (Array.isArray(coverageSource)) {
            billingState.availableWings = getAvailableWings();
            billingState.calendarCoverage = buildMonthCoverage(coverageSource);
            billingState.coverageLoaded = true;
            renderBillingCalendar();
        }
        billingState.sendStatus = new Map();
        markCoverageForSelection();
        renderTenantTable();
        return true;
    }

    showToast("Unable to save bills. Please try again.", "error");
    return false;
}

function toggleSendPrompt(show) {
    const prompt = document.getElementById("billingSendPrompt");
    if (!prompt) return;
    smoothToggle(prompt, show, { baseClass: "fade-overlay" });
}

function toggleSendList(show) {
    const modal = document.getElementById("billingSendListModal");
    if (!modal) return;
    if (show) {
        showModal(modal);
    } else {
        hideModal(modal);
    }
    if (show) {
        renderSendList();
    }
}

function formatWhatsappMessage(summary) {
    const month = billingState.selectedMonthLabel || "this month";
    const motorCount = billingState.motorSnapshot?.includedCount || 1;
    const electricityRate =
        billingState.motorSnapshot?.rate || parseNumber(billingState.meta.electricityRate, true);
    const motorUnits = billingState.motorSnapshot?.units ?? 0;
    const motorPrev = billingState.motorSnapshot?.prev ?? parseNumber(billingState.meta.motorPrev, false);
    const motorNew = billingState.motorSnapshot?.next ?? parseNumber(billingState.meta.motorNew, false);
    const formatAmount = (val) => roundToTwo(val).toFixed(2);
    const totalWords = numberToIndianWords(Math.round(summary.total));
    const payableMonth = getNextMonthLabel(billingState.selectedMonthKey);
    const payableSuffix = summary.payableDay
        ? `Pay on or before *${summary.payableDay}${payableMonth ? ` ${payableMonth}` : ""}*. Thank you!`
        : "Thank you!";

    const lines = [
        `Hi *${summary.name}*, your rent bill for *${month}* has been generated.`,
        "",
        `Electricity rate for the *${month}* = Rs. *${electricityRate}*/ Unit`,
        `Number of residents = *${motorCount}*`,
        "",
        `Rent = Rs. *${formatAmount(summary.rent)}*`,
        "",
        "Electricity :",
        `Previous reading = *${summary.prevReading}*`,
        `Current reading = *${summary.newReading}*`,
        `Previous - Current =  *${summary.units}* Units`,
        `(*${summary.units}* Units) x *${formatAmount(electricityRate)}* =  Rs. *${formatAmount(summary.electricity)}*`,
        "",
        "Motor :",
        `Previous reading = *${motorPrev}*`,
        `Current reading = *${motorNew}*`,
        `Previous - Current =  *${motorUnits}* Units`,
        `(*${motorUnits}* Units x *${formatAmount(electricityRate)}* ) / *${motorCount}* =  Rs. *${formatAmount(summary.motorShare)}*`,
        "",
        `Sweeping = *${formatAmount(summary.sweep)}*`,
        "",
        `Total = *${formatAmount(summary.rent)}* + *${formatAmount(summary.electricity)}* + *${formatAmount(summary.motorShare)}* + *${formatAmount(summary.sweep)}*`,
        "",
        `= Rs. *${formatAmount(summary.total)}* (*${totalWords}*) only.`,
        "",
        payableSuffix,
    ];

    return lines.join("\n");
}

function getPrimaryMobile(value) {
    if (Array.isArray(value)) {
        return String(value.find(Boolean) ?? "");
    }
    if (value === null || value === undefined) {
        return "";
    }
    return String(value);
}

async function openWhatsappExternally(url) {
    try {
        const shell = window.__TAURI__?.shell;
        if (shell?.open) {
            await shell.open(url);
            return true;
        }
    } catch (err) {
        console.error("Unable to open WhatsApp in default browser", err);
    }

    // Use a synthetic anchor click instead of window.open to avoid popup blockers.
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.target = "_blank";
    anchor.rel = "noopener noreferrer";
    anchor.style.display = "none";
    document.body.appendChild(anchor);
    anchor.click();
    requestAnimationFrame(() => anchor.remove());

    return true;
}

function renderSendList() {
    const tbody = document.getElementById("billingSendTableBody");
    const header = document.getElementById("billingSendMonthLabel");
    const empty = document.getElementById("billingSendEmpty");
    if (!tbody || !header || !empty) return;

    header.textContent = billingState.selectedMonthLabel || "Selected month";
    tbody.innerHTML = "";

    if (!billingState.lastGeneratedSummaries.length) {
        empty.classList.remove("hidden");
        return;
    }

    empty.classList.add("hidden");

    billingState.lastGeneratedSummaries.forEach((summary) => {
        const tr = document.createElement("tr");
        tr.className = "border-b last:border-0";
        const sent = billingState.sendStatus.get(summary.id);
        tr.innerHTML = `
            <td class="px-3 py-2 text-[12px] font-semibold text-slate-800">${summary.name}</td>
            <td class="px-3 py-2 text-[12px] font-semibold text-slate-700">₹${summary.total.toFixed(2)}</td>
            <td class="px-3 py-2 text-right text-[12px]">
                ${
                    sent
                        ? '<span class="inline-flex items-center gap-1 text-emerald-700 font-semibold">✓ Sent</span>'
                        : `<button class="send-bill-btn px-3 py-1.5 rounded bg-emerald-600 text-white text-[11px] font-semibold hover:bg-emerald-500" data-id="${
                              summary.id
                          }">Send Bill (WhatsApp)</button>`
                }
            </td>
        `;

        const sendBtn = tr.querySelector(".send-bill-btn");
        if (sendBtn) {
            sendBtn.addEventListener("click", () => {
                const number = getPrimaryMobile(summary.mobile).replace(/\D/g, "");
                if (!number) {
                    showToast(`Missing mobile number for ${summary.name}`, "error");
                    return;
                }
                const message = formatWhatsappMessage(summary);
                const url = `https://web.whatsapp.com/send?phone=${number}&text=${encodeURIComponent(
                    message
                )}&app_absent=0`;
                openWhatsappExternally(url).then((opened) => {
                    if (!opened) return;
                    billingState.sendStatus.set(summary.id, true);
                    renderSendList();
                });
            });
        }

        tbody.appendChild(tr);
    });
}

function setupModalEvents() {
    const modal = document.getElementById("billingWingModal");
    const closeButtons = document.querySelectorAll(".billing-modal-close");
    const nextButton = document.getElementById("billingNextBtn");
    const generateBtn = document.getElementById("billingGenerateBtn");
    const wingSelect = document.getElementById("billingWingSelect");
    const promptNo = document.getElementById("billingSendPromptNo");
    const promptYes = document.getElementById("billingSendPromptYes");
    const promptClose = document.getElementById("billingSendPromptClose");
    const sendCloseButtons = document.querySelectorAll(".billingSendListClose");

    closeButtons.forEach((btn) =>
        btn.addEventListener("click", () => {
            if (modal) hideModal(modal);
        })
    );

    if (wingSelect) {
        wingSelect.addEventListener("change", () => {
            const canonicalWing = getCanonicalWingValue(wingSelect.value);
            billingState.selectedWing = canonicalWing;
            billingState.selectedWingNormalized = normalizeWing(canonicalWing);
            updateSelectionChips();
        });
    }

    if (nextButton) {
        nextButton.addEventListener("click", () => {
            handleNextStep();
        });
    }

    if (generateBtn) {
        generateBtn.addEventListener("click", () => {
            handleGenerateAndPrompt();
        });
    }

    if (promptNo) {
        promptNo.addEventListener("click", () => {
            toggleSendPrompt(false);
        });
    }

    if (promptClose) {
        promptClose.addEventListener("click", () => {
            toggleSendPrompt(false);
        });
    }

    if (promptYes) {
        promptYes.addEventListener("click", () => {
            toggleSendPrompt(false);
            toggleSendList(true);
        });
    }

    sendCloseButtons.forEach((btn) =>
        btn.addEventListener("click", () => {
            toggleSendList(false);
        })
    );

    applyMetaListeners();
}

function cloneWingOptions(targetId) {
    const source = document.getElementById("wing");
    const target = document.getElementById(targetId);
    if (!source || !target) return;

    const previousValue = target.value;
    target.innerHTML = "";
    Array.from(source.options).forEach((opt) => {
        const clone = opt.cloneNode(true);
        target.appendChild(clone);
    });

    if (previousValue && Array.from(target.options).some((o) => o.value === previousValue)) {
        target.value = previousValue;
    }
}

/**
 * Initializes the Billing tab by loading cached data, wiring events,
 * and rendering the default view state.
 */
export function initBillingFeature() {
    renderBillingCalendar();
    cloneWingOptions("billingWingSelect");
    setupModalEvents();
    refreshBillingCalendarCoverage();

    document.addEventListener("wings:updated", () => {
        cloneWingOptions("billingWingSelect");
        billingState.availableWings = getAvailableWings();
        renderBillingCalendar();
    });
}
