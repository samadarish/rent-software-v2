/**
 * Billing Feature Module
 *
 * Calculates monthly charges, renders bill previews, and synchronizes
 * billing metadata with Google Sheets. This module keeps the billing
 * state, parses tenant/unit data, and wires modal interactions.
 */

import {
    formatCurrency as formatCurrencyBase,
    normalizeMonthKey,
    numberToIndianWords,
    numberToIndianWordsHindi,
} from "../../utils/formatters.js";
import { normalizeWing } from "../../utils/normalizers.js";
import { escapeHtml } from "../../utils/htmlUtils.js";
import { cloneSelectOptions, hideModal, showModal, showToast, smoothToggle } from "../../utils/ui.js";
import { ensureTenantDirectoryLoaded, getActiveTenantsForWing } from "../tenants/tenants.js";
import { fetchBillingRecord, fetchGeneratedBills, saveBillingRecord } from "../../api/sheets.js";

const CALENDAR_PAGE_SIZE = 12;

const billingState = {
    selectedMonthKey: null,
    selectedMonthLabel: "",
    selectedWing: "",
    selectedWingNormalized: "",
    selectedWingLabel: "",
    lastGeneratedSummaries: [],
    sendStatus: new Map(),
    selections: new Map(), // key: id -> { lang: 'en'|'hi', selected: boolean }
    calendarCoverage: new Map(),
    calendarCoverageCache: new Map(),
    availableWings: [],
    coverageLoaded: false,
    calendarPage: 0,
    calendarRangeKey: "",
    motorSnapshot: null,
    savedSnapshot: null,
    meta: {
        electricityRate: "",
        sweepingPerFlat: "",
        motorPrev: "",
        motorNew: "",
    },
    tenants: [],
};

const billingRecordCache = new Map();
let billingRenderHandle = null;

function scheduleBillingRender() {
    if (billingRenderHandle !== null) return;
    billingRenderHandle = requestAnimationFrame(() => {
        billingRenderHandle = null;
        renderTenantTable();
        renderMotorSummary();
        updateGenerateButtonState();
    });
}

function bindTenantTableEvents() {
    const tbody = document.getElementById("billingTenantTableBody");
    if (!tbody || tbody.dataset.bound === "true") return;
    tbody.dataset.bound = "true";
    tbody.addEventListener("change", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        const row = target.closest("tr");
        if (!row) return;
        const idx = Number(row.dataset.index);
        if (Number.isNaN(idx)) return;
        const tenant = billingState.tenants[idx];
        if (!tenant) return;

        if (target.classList.contains("tenant-prev")) {
            tenant.prevReading = target.value;
            scheduleBillingRender();
            return;
        }
        if (target.classList.contains("tenant-new")) {
            tenant.newReading = target.value;
            scheduleBillingRender();
            return;
        }
        if (target.classList.contains("tenant-include")) {
            tenant.included = target.checked;
            tenant.hasBill = false;
            scheduleBillingRender();
        }
    });
}
const formatCurrency = (amount) =>
    formatCurrencyBase(amount, {
        useGrouping: false,
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
        coerceEmptyToZero: true,
        roundTo: 2,
        invalidValue: "\u20B90.00",
    });

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

function normalizeNumberValue(value) {
    if (value === "" || value === null || value === undefined) return "";
    const num = Number(value);
    return Number.isNaN(num) ? String(value).trim() : num;
}

function getTenantSnapshotKey(tenant) {
    const tenancyId = tenant?.tenancyId || tenant?.tenancy_id || "";
    if (tenancyId) return tenancyId;
    const identity = getTenantIdentityKey(tenant);
    if (identity) return identity;
    return normalizeTenantKey(
        tenant?.tenantKey || tenant?.tenantName || tenant?.tenant_name || tenant?.name || ""
    );
}

function buildSavedSnapshot(record) {
    const entries = Array.isArray(record?.tenants) ? record.tenants : [];
    const tenantMap = new Map();
    entries.forEach((entry) => {
        const key = getTenantSnapshotKey(entry);
        if (!key) return;
        tenantMap.set(key, {
            included: normalizeIncludedFlag(entry?.included),
            prevReading: entry?.prevReading ?? entry?.prev_reading ?? "",
            newReading: entry?.newReading ?? entry?.new_reading ?? "",
            rentAmount: entry?.rentAmount ?? entry?.rent_amount ?? "",
        });
    });
    return {
        meta: normalizeMetaPayload(record?.meta || {}),
        tenantMap,
        hasBills: record?.hasReadings ?? entries.length > 0,
    };
}

function hasBillingChanges() {
    const snapshot = billingState.savedSnapshot;
    if (!snapshot || !snapshot.hasBills) return false;

    const currentMeta = normalizeMetaPayload(billingState.meta || {});
    if (
        normalizeNumberValue(currentMeta.electricityRate) !==
            normalizeNumberValue(snapshot.meta.electricityRate) ||
        normalizeNumberValue(currentMeta.sweepingPerFlat) !==
            normalizeNumberValue(snapshot.meta.sweepingPerFlat) ||
        normalizeNumberValue(currentMeta.motorPrev) !== normalizeNumberValue(snapshot.meta.motorPrev) ||
        normalizeNumberValue(currentMeta.motorNew) !== normalizeNumberValue(snapshot.meta.motorNew)
    ) {
        return true;
    }

    for (const tenant of billingState.tenants) {
        const key = getTenantSnapshotKey(tenant);
        if (!key) continue;
        const saved = snapshot.tenantMap.get(key);
        if (!saved) return true;
        if (normalizeIncludedFlag(tenant?.included) !== normalizeIncludedFlag(saved.included)) return true;
        if (normalizeNumberValue(tenant?.prevReading) !== normalizeNumberValue(saved.prevReading)) return true;
        if (normalizeNumberValue(tenant?.newReading) !== normalizeNumberValue(saved.newReading)) return true;
        if (normalizeNumberValue(tenant?.rentAmount) !== normalizeNumberValue(saved.rentAmount)) return true;
    }

    return false;
}

function updateGenerateButtonState() {
    const btn = document.getElementById("billingGenerateBtn");
    if (!btn) return;
    const shouldUpdate = hasBillingChanges();
    if (shouldUpdate) {
        btn.textContent = "Update bills";
        btn.classList.remove("bg-emerald-600", "hover:bg-emerald-500");
        btn.classList.add("bg-amber-500", "hover:bg-amber-400");
        return;
    }
    btn.textContent = "Generate bills";
    btn.classList.remove("bg-amber-500", "hover:bg-amber-400");
    btn.classList.add("bg-emerald-600", "hover:bg-emerald-500");
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

function buildMonthEntry(date) {
    const monthNumber = `${date.getMonth() + 1}`.padStart(2, "0");
    return {
        key: `${date.getFullYear()}-${monthNumber}`,
        label: date.toLocaleString("default", { month: "long", year: "numeric" }),
    };
}

function getBillingCalendarPageInfo(pageIndex = 0) {
    const page = Math.max(0, Number(pageIndex) || 0);
    const now = new Date();
    const endOffset = -1 - page * CALENDAR_PAGE_SIZE;
    const startOffset = endOffset - (CALENDAR_PAGE_SIZE - 1);
    const months = [];

    for (let offset = startOffset; offset <= endOffset; offset += 1) {
        const d = new Date(now.getFullYear(), now.getMonth() + offset, 1);
        months.push(buildMonthEntry(d));
    }

    const fromMonth = months[0]?.key || "";
    const toMonth = months[months.length - 1]?.key || "";
    const rangeLabel =
        months.length > 0 ? `${months[0].label} - ${months[months.length - 1].label}` : "Last 12 months";

    return {
        page,
        months,
        fromMonth,
        toMonth,
        rangeLabel,
    };
}

function getCalendarRangeKey(rangeInfo) {
    return `${rangeInfo.fromMonth || ""}__${rangeInfo.toMonth || ""}`;
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

function getBillIdInitials(name) {
    const trimmed = (name || "").toString().trim();
    if (!trimmed) return "XX";
    const parts = trimmed.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) {
        const first = parts[0][0] || "X";
        const last = parts[parts.length - 1][0] || "X";
        return `${first}${last}`.toUpperCase();
    }
    const word = parts[0] || "";
    const first = word[0] || "X";
    const last = word[word.length - 1] || first || "X";
    return `${first}${last}`.toUpperCase();
}

function getBillIdMonthToken(monthKey) {
    return (monthKey || "").toString().replace(/[^0-9]/g, "");
}

function getBillIdTimeToken(date) {
    const hours = `${date.getHours()}`.padStart(2, "0");
    const minutes = `${date.getMinutes()}`.padStart(2, "0");
    return `${hours}${minutes}`;
}

function generateBillId(name, monthKey, timeToken) {
    const initials = getBillIdInitials(name);
    const monthToken = getBillIdMonthToken(monthKey);
    const randomToken = Math.random().toString(36).replace(/[^a-z0-9]/g, "").slice(2, 5).padEnd(3, "0");
    return `${initials}-${monthToken}-${randomToken}-${timeToken}`;
}

function getSendStatusKey(id, lang) {
    return `${id || ""}__${lang || "en"}`;
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
    billingState.savedSnapshot = null;
    const tenantBody = document.getElementById("billingTenantTableBody");
    if (tenantBody) tenantBody.innerHTML = "";
    const emptyState = document.getElementById("billingTenantEmpty");
    if (emptyState) emptyState.classList.add("hidden");
    updateGenerateButtonState();
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

function getBillingCalendarControlNodes() {
    return {
        labels: Array.from(document.querySelectorAll("[data-billing-months-label]")),
        prevButtons: Array.from(document.querySelectorAll("[data-billing-months-prev]")),
        nextButtons: Array.from(document.querySelectorAll("[data-billing-months-next]")),
    };
}

function updateBillingCalendarControls(rangeInfo) {
    const isNewest = rangeInfo.page === 0;
    const { labels, prevButtons, nextButtons } = getBillingCalendarControlNodes();

    labels.forEach((label) => {
        label.textContent = rangeInfo.rangeLabel || "Last 12 months";
    });

    nextButtons.forEach((nextBtn) => {
        nextBtn.disabled = isNewest;
        nextBtn.classList.toggle("opacity-50", isNewest);
        nextBtn.classList.toggle("cursor-not-allowed", isNewest);
    });

    prevButtons.forEach((prevBtn) => {
        prevBtn.disabled = false;
        prevBtn.classList.remove("opacity-50", "cursor-not-allowed");
    });
}

function setBillingCalendarPage(pageIndex) {
    const nextPage = Math.max(0, Number(pageIndex) || 0);
    if (nextPage === billingState.calendarPage) return;
    billingState.calendarPage = nextPage;

    const pageInfo = getBillingCalendarPageInfo(nextPage);
    const rangeKey = getCalendarRangeKey(pageInfo);
    const cached = billingState.calendarCoverageCache.get(rangeKey);

    if (cached) {
        billingState.calendarCoverage = cached;
        billingState.coverageLoaded = true;
    } else {
        billingState.coverageLoaded = false;
    }

    renderBillingCalendar();
    if (!cached) {
        refreshBillingCalendarCoverage();
    }
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
    if (billingState.calendarRangeKey) {
        billingState.calendarCoverageCache.set(billingState.calendarRangeKey, billingState.calendarCoverage);
    }
    renderBillingCalendar();
}

function renderBillingCalendar() {
    const grid = document.getElementById("billingMonthsGrid");
    if (!grid) return;

    if (!billingState.availableWings.length) {
        billingState.availableWings = getAvailableWings();
    }

    grid.innerHTML = "";
    const pageInfo = getBillingCalendarPageInfo(billingState.calendarPage);
    billingState.calendarRangeKey = getCalendarRangeKey(pageInfo);
    updateBillingCalendarControls(pageInfo);
    const months = pageInfo.months;

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

async function refreshBillingCalendarCoverage(pageIndex = billingState.calendarPage) {
    const pageInfo = getBillingCalendarPageInfo(pageIndex);
    const rangeKey = getCalendarRangeKey(pageInfo);

    billingState.calendarRangeKey = rangeKey;
    billingState.availableWings = getAvailableWings();

    const cached = billingState.calendarCoverageCache.get(rangeKey);
    if (cached) {
        billingState.calendarCoverage = cached;
        billingState.coverageLoaded = true;
        renderBillingCalendar();
        return;
    }

    const { bills, coverage } = await fetchGeneratedBills({
        fromMonth: pageInfo.fromMonth,
        toMonth: pageInfo.toMonth,
    });
    const coverageSource = Array.isArray(coverage) && coverage.length ? coverage : bills;
    const coverageMap = buildMonthCoverage(Array.isArray(coverageSource) ? coverageSource : []);

    billingState.calendarCoverage = coverageMap;
    billingState.calendarCoverageCache.set(rangeKey, coverageMap);
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

    bindTenantTableEvents();
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
        const safeName = escapeHtml(tenant.name || "");
        const safeUnit = escapeHtml(tenant.unitNumber || "-");
        const safePrev = escapeHtml(tenant.prevReading ?? "");
        const safeNew = escapeHtml(tenant.newReading ?? "");
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
        tr.dataset.index = String(idx);
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
                <div class="font-semibold text-[12px] leading-tight">${safeName}</div>
            </td>
            <td class="px-2 py-2 text-xs">${safeUnit}</td>
            <td class="px-2 py-2 text-xs font-semibold">${formatCurrency(tenant.rentAmount)}</td>
            <td class="px-2 py-2 text-xs"><input type="number" inputmode="numeric" step="1" min="0" pattern="[0-9]*" class="w-full border rounded px-2 py-1 text-xs tenant-prev input-no-spinner ${invalidReading ? "border-rose-500 bg-rose-50 text-rose-700" : "border-slate-200"}" value="${safePrev}" /></td>
            <td class="px-2 py-2 text-xs"><input type="number" inputmode="numeric" step="1" min="0" pattern="[0-9]*" class="w-full border rounded px-2 py-1 text-xs tenant-new input-no-spinner ${invalidReading ? "border-rose-500 bg-rose-50 text-rose-700" : "border-slate-200"}" value="${safeNew}" /></td>
            <td class="px-2 py-2 text-xs font-semibold text-indigo-700">${formatCurrency(charges.electricity)}</td>
            <td class="px-2 py-2 text-xs">${formatCurrency(charges.motorShare)}</td>
            <td class="px-2 py-2 text-xs">${formatCurrency(charges.sweepAmount)}</td>
            <td class="px-2 py-2 text-xs font-bold text-slate-800">${formatCurrency(charges.total)}</td>
            <td class="px-2 py-2 text-xs text-right">
                ${
                    isConfirmed
                        ? '<span class="inline-flex items-center gap-1 text-emerald-700 font-semibold">OK</span>'
                        : ""
                }
            </td>
        `;

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
            scheduleBillingRender();
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
    const merged = [];

    activeTenants.forEach((tenant, idx) => {
        const candidateKeys = getTenantKeyCandidates(tenant);
        const identityKey =
            candidateKeys[0] ||
            normalizeTenantKey(`${tenant.wing || ""}-${tenant.unitId || ""}-${tenant.tenantFullName || tenant.name || ""}`) ||
            `tenant-${idx}`;
        const keys = candidateKeys.length ? candidateKeys : [identityKey];

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
        const hasBill = saved ? normalizeIncludedFlag(saved.included) : false;
        const savedBillId = saved?.billId || saved?.bill_id || "";

        merged.push({
            tenantKey: tenant.tenantKey || identityKey,
            grn: tenant.grnNumber || tenant.grn || "",
            name: tenant.tenantFullName || tenant.name || "Unnamed",
            tenancyId: tenant.tenancyId || "",
            unitId: tenant.unitId || "",
            unitNumber:
                tenant.unitNumber ||
                tenant.unit_number ||
                saved?.unitNumber ||
                saved?.unit_number ||
                previous?.unitNumber ||
                previous?.unit_number ||
                "",
            rentAmount: tenant.rentAmount || saved?.rentAmount || saved?.rent_amount || 0,
            mobile: tenant.tenantMobile || tenant.mobile || tenant.phone || "",
            prevReading:
                savedPrev !== undefined && savedPrev !== null && savedPrev !== ""
                    ? savedPrev
                    : previousPrev,
            newReading: savedNew ?? "",
            included,
            hasBill,
            billId: savedBillId,
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
    billingState.savedSnapshot = buildSavedSnapshot(record);
    populateInputsFromState();
    renderMotorSummary();
    renderTenantTable();
    updateGenerateButtonState();

    if (loader) loader.classList.add("hidden");
    if (form) form.classList.remove("opacity-50");
}

export async function refreshBillingData(force = false) {
    refreshBillingCalendarCoverage();
    if (!billingState.selectedMonthKey || !getSelectedWingNormalized()) {
        return;
    }
    await loadBillingData({ force });
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

    const invalidTenant = selectedTenants.find((t) => {
        const hasMissing =
            t.prevReading === "" ||
            t.prevReading === undefined ||
            t.newReading === "" ||
            t.newReading === undefined;
        const prevVal = Number(t.prevReading);
        const newVal = Number(t.newReading);
        const notNumbers = Number.isNaN(prevVal) || Number.isNaN(newVal);
        const hasRollback = Number.isFinite(prevVal) && Number.isFinite(newVal) && newVal < prevVal;
        return hasMissing || notNumbers || hasRollback;
    });
    if (invalidTenant) {
        showToast(
            `Enter valid meter readings for ${invalidTenant.name || "tenant"} (new reading must be >= previous).`,
            "error"
        );
        return false;
    }

    const normalizedMonthKey = normalizeMonthKey(billingState.selectedMonthKey);
    let billTimeToken = "";
    selectedTenants.forEach((tenant) => {
        const existingBillId = (tenant.billId || "").toString().trim();
        if (existingBillId) {
            tenant.billId = existingBillId;
            return;
        }
        if (!billTimeToken) {
            billTimeToken = getBillIdTimeToken(new Date());
        }
        tenant.billId = generateBillId(tenant.name || tenant.tenantName || "", normalizedMonthKey, billTimeToken);
    });

    const motor = computeMotorShare();
    const motorPerTenant = motor.perTenant;

    const normalizedWing = getSelectedWingNormalized();
    const calendarRange = getBillingCalendarPageInfo(billingState.calendarPage);
    const payload = {
        monthKey: normalizedMonthKey,
        monthLabel: billingState.selectedMonthLabel,
        wing: normalizedWing,
        wingLabel: billingState.selectedWingLabel || billingState.selectedWing,
        coverageFrom: calendarRange.fromMonth,
        coverageTo: calendarRange.toMonth,
        meta: { ...billingState.meta },
        tenants: billingState.tenants.map((t) => ({
            tenantKey: getTenantIdentityKey(t),
            grn: t.grn,
            name: t.name,
            tenancyId: t.tenancyId,
            billId: t.billId || "",
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
                billId: t.billId || "",
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
            const coverageMap = buildMonthCoverage(coverageSource);
            const rangeKey = getCalendarRangeKey(calendarRange);
            billingState.availableWings = getAvailableWings();
            billingState.calendarCoverage = coverageMap;
            billingState.coverageLoaded = true;
            billingState.calendarCoverageCache.set(rangeKey, coverageMap);
            renderBillingCalendar();
        }
        billingState.sendStatus = new Map();
        billingState.savedSnapshot = buildSavedSnapshot({
            meta: payload.meta,
            tenants: payload.tenants,
            hasReadings: true,
        });
        updateGenerateButtonState();
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
        renderSendList();
        
        // START POLLING for Login State Changes (Every 500ms)
        if (window._billingStatePoller) clearInterval(window._billingStatePoller);
        window._billingStatePoller = setInterval(() => {
             // Pass null to force read from storage
             // This ensures if event was missed, we still catch up
             updateAutoSendButton(null);
        }, 500);

    } else {
        hideModal(modal);
        
        // STOP POLLING
        if (window._billingStatePoller) {
            clearInterval(window._billingStatePoller);
            window._billingStatePoller = null;
        }
    }
}

function formatWhatsappMessage(summary, language = "en") {
    const month = billingState.selectedMonthLabel || "this month";
    const billId = summary.billId || summary.bill_id || "";
    const motorCount = billingState.motorSnapshot?.includedCount || 1;
    const electricityRate =
        billingState.motorSnapshot?.rate || parseNumber(billingState.meta.electricityRate, true);
    const motorUnits = billingState.motorSnapshot?.units ?? 0;
    const motorPrev = billingState.motorSnapshot?.prev ?? parseNumber(billingState.meta.motorPrev, false);
    const motorNew = billingState.motorSnapshot?.next ?? parseNumber(billingState.meta.motorNew, false);
    const formatAmount = (val) => roundToTwo(val).toFixed(2);
    const totalWords = numberToIndianWords(Math.round(summary.total));
    const totalWordsHi = numberToIndianWordsHindi(Math.round(summary.total));
    const payableMonth = getNextMonthLabel(billingState.selectedMonthKey);
    const payableSuffixEn = summary.payableDay
        ? `Pay on or before *${summary.payableDay}${payableMonth ? ` ${payableMonth}` : ""}*. Thank you!`
        : "Thank you!";
    const payableSuffixHi = summary.payableDay
        ? `कृपया *${summary.payableDay}${payableMonth ? ` ${payableMonth}` : ""}* तक भुगतान करें। धन्यवाद!`
        : "धन्यवाद!";

    const englishLines = [
        `Hi *${summary.name}*, your rent bill for *${month}* has been generated. Your Bill ID is *${billId}*.`,
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
        payableSuffixEn,
    ];

    const hindiLines = [
        `*${summary.name}*, *${month}* के लिए आपका किराया बिल तैयार हो गया है। आपका बिल आईडी *${billId}* है।`,
        "",
        `*${month}* के लिए बिजली रेट = Rs. *${electricityRate}*/ यूनिट`,
        `निवासियों की संख्या = *${motorCount}*`,
        "",
        `किराया = Rs. *${formatAmount(summary.rent)}*`,
        "",
        "बिजली :",
        `पिछला रीडिंग = *${summary.prevReading}*`,
        `वर्तमान रीडिंग = *${summary.newReading}*`,
        `पिछला - वर्तमान =  *${summary.units}* यूनिट`,
        `(*${summary.units}* यूनिट) x *${formatAmount(electricityRate)}* =  Rs. *${formatAmount(summary.electricity)}*`,
        "",
        "मोटर :",
        `पिछला रीडिंग = *${motorPrev}*`,
        `वर्तमान रीडिंग = *${motorNew}*`,
        `पिछला - वर्तमान =  *${motorUnits}* यूनिट`,
        `(*${motorUnits}* यूनिट x *${formatAmount(electricityRate)}* ) / *${motorCount}* =  Rs. *${formatAmount(summary.motorShare)}*`,
        "",
        `सफाई = *${formatAmount(summary.sweep)}*`,
        "",
        `कुल = *${formatAmount(summary.rent)}* + *${formatAmount(summary.electricity)}* + *${formatAmount(summary.motorShare)}* + *${formatAmount(summary.sweep)}*`,
        "",
        `= Rs. *${formatAmount(summary.total)}* (*${totalWordsHi}*) मात्र।`,
        "",
        payableSuffixHi,
    ];

    return (language === "hi" ? hindiLines : englishLines).join("\n");
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

function bindSendListEvents() {
    const tbody = document.getElementById("billingSendTableBody");
    if (!tbody || tbody.dataset.bound === "true") return;
    tbody.dataset.bound = "true";
    tbody.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        
        // Handle Copy Button
        const copyBtn = target.closest(".send-copy-btn");
        if (copyBtn) {
             const id = copyBtn.dataset.id || "";
             const row = copyBtn.closest("tr");
             if (!row) return;

             // Determine language from radio
             const langRadio = row.querySelector(`input[name="lang_${id}"]:checked`);
             const lang = langRadio ? langRadio.value : "en";
             
             const summary = billingState.lastGeneratedSummaries.find((item) => String(item.id) === String(id));
             if (!summary) return;
             
             const message = formatWhatsappMessage(summary, lang);
             copyToClipboard(message).then(() => {
                 showToast("Message copied to clipboard", "success");
             }).catch(err => {
                 console.error("Copy failed", err);
                 showToast("Failed to copy message", "error");
             });
             return;
        }

        // Handle Send Button
        const btn = target.closest(".send-bill-btn");
        if (!btn) return;
        const id = btn.dataset.id || "";
        const row = btn.closest("tr");
        if (!row) return;

        // Determine language from radio
        const langRadio = row.querySelector(`input[name="lang_${id}"]:checked`);
        const lang = langRadio ? langRadio.value : "en";
        
        const summary = billingState.lastGeneratedSummaries.find((item) => String(item.id) === String(id));
        if (!summary) return;
        const number = getPrimaryMobile(summary.mobile).replace(/\D/g, "");
        if (!number) {
            showToast(`Missing mobile number for ${summary.name}`, "error");
            return;
        }
        const message = formatWhatsappMessage(summary, lang);
        
        let url = ""; 
        
        // If technical user / native request:
        // Use custom scheme or backend command if desired for manual send too?
        // User asked to "use injecttion hooks" for "Auto-send".
        // For manual "Send Manual", maybe standard web link is fine, but since we have "whatsapp-login"
        // in backend, maybe we should use that session?
        // Let's stick to standard URL for manual send unless specified otherwise, 
        // OR better, try to use the same mechanism if running in Tauri to be consistent.
        
        // Actually, for manual send, opening the browser/window is expected.
        url = `https://web.whatsapp.com/send?phone=${number}&text=${encodeURIComponent(message)}&app_absent=0`;
        
        openWhatsappExternally(url).then((opened) => {
            if (!opened) return;
            billingState.sendStatus.set(getSendStatusKey(summary.id, lang), true);
            // We can't easily detect "sent" status for manual external open without more complex hooks
            // But we can mark it as opened/attempted.
            // renderSendList() replaces the DOM, so finding the element again is tricky if we re-render immediately.
            // But we should re-render to show the checkmark.
            renderSendList();
        });
    });
}

function renderSendList() {
    const tbody = document.getElementById("billingSendTableBody");
    const header = document.getElementById("billingSendMonthLabel");
    const emptyDiv = document.getElementById("billingSendEmpty");
    if (!tbody || !header || !emptyDiv) return;

    bindSendListEvents();
    header.textContent = billingState.selectedMonthLabel || "Selected month";
    tbody.innerHTML = "";

    if (!billingState.lastGeneratedSummaries.length) {
        tbody.closest("table").classList.add("hidden");
        emptyDiv.classList.remove("hidden");
        return;
    }
    tbody.closest("table").classList.remove("hidden");
    emptyDiv.classList.add("hidden");

    billingState.lastGeneratedSummaries.forEach((summary) => {
        const safeId = escapeHtml(summary.id || "");
        
        // Logged In Check
        const isLoggedIn = localStorage.getItem("wa_logged_in") === "true";

        // Load persisted state or default
        // If logged out: force unselected but don't overwrite persistent state yet? 
        // User said "uncheck everybody".
        // If logged in: default to selected if not previously set? 
        
        let persisted = billingState.selections.get(safeId);
        
        // Default Logic if no persistence
        if (!persisted) {
             persisted = { lang: 'en', selected: isLoggedIn }; // Select all if logged in by default
        }
        
        // Override if logged out -> always unchecked valid for rendering
        const isSelected = isLoggedIn ? persisted.selected : false;
        const currentLang = persisted.lang;

        // Check if already sent (any lang? or specific?)
        // If sent, maybe we show green row already?
        const sentHindi = billingState.sendStatus.get(getSendStatusKey(summary.id, "hi"));
        const sentEnglish = billingState.sendStatus.get(getSendStatusKey(summary.id, "en"));
        // If "Auto-sent" logic applies, maybe we want to know if *either* was just auto-sent.
        // But for now, just render normally.

        const tr = document.createElement("tr");
        const isSentAny = sentHindi || sentEnglish; 
        
        // If we want to show "green" row if fully processed? 
        // User asked "once auto send has been done remove the checkbox and make the row for that tenant green"
        // So if we detect it's been sent *in this session* or status is true, maybe we do it?
        // Let's rely on standard logic: if unchecked (removed) and sent, show green.
        // Actually `handleAutoSend` unchecks it.
        
        if (isSentAny && !isSelected) { 
             // Maybe user wants to send again manually? 
             // Ideally we shouldn't block manual sending even if auto-sent.
             // But let's apply the requested style.
             tr.className = "border-b last:border-0 hover:bg-emerald-50 bg-emerald-50 transition-colors";
        } else {
             tr.className = "border-b last:border-0 hover:bg-slate-50 transition-colors";
        }

        const safeName = escapeHtml(summary.name || "");
        const totalValue = Number(summary.total) || 0;
        
        // Checkbox HTML
        const checkboxHtml = (isSentAny && !isSelected) ? 
             `<span class="text-emerald-600 font-bold text-[10px]">✓</span>` :
             `<input type="checkbox" class="send-check h-3.5 w-3.5 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer" data-id="${safeId}" ${isSelected ? "checked" : ""} ${!isLoggedIn ? "disabled" : ""}>`;

        tr.innerHTML = `
            <td class="px-3 py-2 w-8 text-center align-middle">
                 ${checkboxHtml}
            </td>
            <td class="px-3 py-2 text-[12px] font-semibold text-slate-800 align-middle">${safeName}</td>
            <td class="px-3 py-2 text-[12px] font-semibold text-slate-700 align-middle">Rs. ${totalValue.toFixed(2)}</td>
            
            <td class="px-3 py-2 text-center align-middle">
                <div class="flex items-center justify-center gap-1">
                   <input type="radio" name="lang_${safeId}" value="hi" class="lang-radio h-3 w-3 border-gray-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer" ${currentLang === 'hi' ? 'checked' : ''} data-id="${safeId}">
                </div>
            </td>
            <td class="px-3 py-2 text-center align-middle">
                <div class="flex items-center justify-center gap-1">
                    <input type="radio" name="lang_${safeId}" value="en" class="lang-radio h-3 w-3 border-gray-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer" ${currentLang === 'en' ? 'checked' : ''} data-id="${safeId}">
                </div>
            </td>
            
            <td class="px-3 py-2 text-right text-[12px] align-middle">
                <div class="inline-flex items-center justify-end gap-2">
                    <button class="send-copy-btn p-1.5 rounded text-slate-400 hover:text-slate-600 hover:bg-slate-100" title="Copy message" data-id="${safeId}">
                        <svg class="w-3.5 h-3.5 pointer-events-none" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
                    </button>
                    <button class="send-bill-btn inline-flex items-center gap-1 px-3 py-1.5 rounded bg-[#25D366] text-white border border-[#20ba5a] text-[12px] font-semibold hover:bg-[#20ba5a] shadow-sm" data-id="${safeId}">
                        <span>Send Manual</span>
                    </button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    // Bind checkbox changes for state persistence
    tbody.querySelectorAll(".send-check").forEach(cb => {
        cb.addEventListener("change", (e) => {
             const id = e.target.dataset.id;
             const current = billingState.selections.get(id) || { lang: 'en', selected: false };
             current.selected = e.target.checked;
             billingState.selections.set(id, current);
             updateAutoSendButton();
        });
    });

    // Bind radio changes for state persistence
    tbody.querySelectorAll(".lang-radio").forEach(radio => {
        radio.addEventListener("change", (e) => {
             if(e.target.checked) {
                 const id = e.target.dataset.id;
                 const val = e.target.value;
                 const current = billingState.selections.get(id) || { lang: 'en', selected: false };
                 current.lang = val;
                 billingState.selections.set(id, current);
             }
        });
    });
    
    updateAutoSendButton();
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

    sendCloseButtons.forEach((btn) =>
        btn.addEventListener("click", () => {
            toggleSendList(false);
        })
    );

    const toggleAllCheckbox = document.getElementById("billingSendToggleAll");
    if (toggleAllCheckbox) {
        toggleAllCheckbox.addEventListener("change", (e) => {
            const checked = e.target.checked;
            document.querySelectorAll(".send-check").forEach((cb) => {
                const id = cb.dataset.id;
                cb.checked = checked;
                
                // Persist state
                const current = billingState.selections.get(id) || { lang: 'en', selected: false };
                current.selected = checked;
                billingState.selections.set(id, current);
            });
            updateAutoSendButton();
        });
    }

    const autoSendBtn = document.getElementById("billingAutoSendBtn");
    if (autoSendBtn) {
        autoSendBtn.addEventListener("click", handleAutoSend);
    }

    applyMetaListeners();
    ensureBillingListeners();
    updateAutoSendButton();
}

let _billingListenersBound = false;
function ensureBillingListeners() {
    if (_billingListenersBound) return;
    _billingListenersBound = true;

    if (window.__TAURI__) {
        window.__TAURI__.event.listen("whatsapp-login-success", () => {
             console.log("Billing: Login Event Received");
             localStorage.setItem("wa_logged_in", "true"); // Force persist immediately
             updateAutoSendButton(true);
        });
        window.__TAURI__.event.listen("whatsapp-logout", () => {
             console.log("Billing: Logout Event Received");
             localStorage.setItem("wa_logged_in", "false"); // Force persist immediately
             updateAutoSendButton(false);
        });
    }

    // Fallback: Check on window focus (e.g. after login window closes)
    // Add delay to allow main.js to update localStorage from the event first
    window.addEventListener("focus", () => {
        console.log("Billing: Window Focussed - Checking State (Delayed)");
        setTimeout(() => {
             updateAutoSendButton(null);
        }, 500);
    });

    // LINK WITH MAIN APP STATUS (User Request)
    // This listens to the exact event that updates the navigation/toast
    document.addEventListener("app:whatsapp-status-change", (e) => {
         console.log("Billing: App Status Change Detected", e.detail);
         if (e.detail && typeof e.detail.loggedIn === "boolean") {
             updateAutoSendButton(e.detail.loggedIn);
         }
    });
}

function updateAutoSendButton(forceState = null) {
    const btn = document.getElementById("billingAutoSendBtn");
    const toggleAll = document.getElementById("billingSendToggleAll");
    // Note: billingAutoSendCount may not exist if button is currently showing "Sign in" state
    // It gets recreated when we set innerHTML for logged-in state
    if (!btn) return;
    
    // Use forced state if provided, otherwise check storage
    const isLoggedIn = forceState !== null ? forceState : (localStorage.getItem("wa_logged_in") === "true");
    
    // Logic for State Transitions
    // Check if we are physically transitioning by checking if inputs are disabled vs new login state
    const inputs = document.querySelectorAll(".send-check");
    const areDisabled = inputs.length > 0 && inputs[0].disabled;
    
    if (!isLoggedIn) {
        // ==> LOGGED OUT STATE
        // Disable inputs if not already
        if (!areDisabled) {
             inputs.forEach(ci => { ci.disabled = true; ci.checked = false; });
             if(toggleAll) { toggleAll.disabled = true; toggleAll.checked = false; }
             // Update persist logic? Maybe just visual for now to keep state safe?
             // User said "uncheck everybody".
             inputs.forEach(cb => {
                 const id = cb.dataset.id;
                 const current = billingState.selections.get(id) || { lang: 'en', selected: false };
                 current.selected = false;
                 billingState.selections.set(id, current);
             });
        }
        
        // Show Sign In State
        btn.innerHTML = `<span>To send Sign in to whatsapp</span>`;
        // Use bg-red-400 for "less red" (standard red is bg-red-500/600, rose is pinkish/red)
        btn.className = "px-6 py-2.5 rounded bg-red-400 text-white text-[12px] font-semibold hover:bg-red-500 shadow-sm flex items-center gap-2 transform active:scale-95 transition-all";
        btn.disabled = false;
        btn.dataset.action = "login";
        return;
    }

    // ==> LOGGED IN STATE
    // Logic: If inputs are disabled (just logged in) OR count is 0 (first load logged in?), 
    // force enable and CHECK ALL to ensure "Auto-send" is ready to go.
    
    // Calculate initial count of selected checkboxes
    let count = document.querySelectorAll(".send-check:checked").length;
    
    // Check if we need to "Activate" the form
    const needsActivation = areDisabled || count === 0;

    if (needsActivation) {
         console.log("Billing: Activating Checkboxes (Enable & Select All)");
         inputs.forEach(ci => { 
             ci.disabled = false; 
             ci.checked = true; 
             
             // Update persistent state so it sticks
             const id = ci.dataset.id;
             const current = billingState.selections.get(id) || { lang: 'en', selected: false };
             current.selected = true;
             billingState.selections.set(id, current);
         });
         
         if(toggleAll) { 
             toggleAll.disabled = false; 
             toggleAll.checked = true; 
         }
         
         // Re-query checked inputs to update count correctly below
         const freshCheckboxes = Array.from(document.querySelectorAll(".send-check:checked"));
         // Update local variable
         count = freshCheckboxes.length; 
    } else {
        // Just ensure they are enabled if they were disabled
        if (areDisabled) {
             console.log("Billing: Enabling Checkboxes only");
             inputs.forEach(ci => { ci.disabled = false; });
             if(toggleAll) toggleAll.disabled = false;
        }
    }

    // Restore Auto Send State
    btn.dataset.action = "send";
    
    // Explicitly reset classes
    btn.className = "px-6 py-2.5 rounded bg-emerald-600 text-white text-[12px] font-semibold hover:bg-emerald-500 shadow-sm flex items-center gap-2 transform active:scale-95 transition-all";
    
    // Re-calculate count
    const finalCount = document.querySelectorAll(".send-check:checked").length;
    
    btn.innerHTML = `<span>Auto-send Selected</span> <span id="billingAutoSendCount" class="bg-white/20 px-1.5 rounded text-[10px] ${finalCount === 0 ? 'hidden' : ''}">${finalCount}</span>`;
    
    if (finalCount === 0) {
        btn.classList.add("opacity-50", "cursor-not-allowed");
        btn.disabled = true;
    } else {
        btn.classList.remove("opacity-50", "cursor-not-allowed");
        btn.disabled = false;
    }
}

async function handleAutoSend() {
    const btn = document.getElementById("billingAutoSendBtn");

    // Handle Login Action if button state is 'login'
    if (btn && btn.dataset.action === "login") {
        try {
            if (window.__TAURI__) {
                await window.__TAURI__.core.invoke("open_whatsapp");
            } else {
                alert("Cannot open WhatsApp in browser mode");
            }
        } catch (err) {
            console.error("Failed to open WhatsApp login", err);
            showToast("Failed to open WhatsApp window", "error");
        }
        return;
    }

    const checkboxes = Array.from(document.querySelectorAll(".send-check:checked"));
    if (!checkboxes.length) return;

    if (btn) {
        btn.textContent = "Sending...";
        btn.disabled = true;
    }

    let sentCount = 0;
    const total = checkboxes.length;

    for (const [index, checkbox] of checkboxes.entries()) {
        const id = checkbox.dataset.id;
        const row = checkbox.closest("tr");
        if(!row) continue;

        // Determine language
        // Since we persist, we could read from map, but DOM is fine
        const langRadio = row.querySelector(`input[name="lang_${id}"]:checked`);
        const lang = langRadio ? langRadio.value : "en";
        
        const summary = billingState.lastGeneratedSummaries.find((item) => String(item.id) === String(id));
        if (!summary) continue;
        
        const number = getPrimaryMobile(summary.mobile).replace(/\D/g, "");
        if (!number) {
            showToast(`Skipped ${summary.name}: No mobile number`, "error");
            continue;
        }

        const message = formatWhatsappMessage(summary, lang);
        
        try {
             if (window.__TAURI__) {
                 // Pass progress label (1-based index)
                 const progressLabel = `${index + 1}/${total} Sending`;
                 
                 // This is now awaitable and blocks until window closes!
                 await window.__TAURI__.core.invoke("send_whatsapp_message", { 
                     phone: number, 
                     message: message,
                     progressLabel: progressLabel
                 });
                 
                 billingState.sendStatus.set(getSendStatusKey(summary.id, lang), true);
                 
                 // UI Updates on Success
                 // 1. Remove Checkbox
                 checkbox.checked = false;
                 checkbox.disabled = true; 
                 // Update persist state
                 const current = billingState.selections.get(id) || { lang, selected: false };
                 current.selected = false;
                 billingState.selections.set(id, current);

                 // 2. Make row green logic
                 row.classList.add("bg-emerald-50");
                 // Maybe add a "Sent" badge in the checkbox column?
                 const checkCell = row.querySelector("td:first-child");
                 if(checkCell) checkCell.innerHTML = `<span class="text-emerald-600 font-bold text-[10px]">✓</span>`;

                 sentCount++;
             } else {
                 console.warn("Auto-send only works in Tauri environment");
                 showToast("Auto-send unavailable in browser mode", "error");
                 break;
             }
        } catch (err) {
            console.error("Failed to auto-send for " + summary.name, err);
            showToast(`Failed to send to ${summary.name}`, "error");
            // If timeout or error, we continue to next? Yes.
        }
    }
    
    if (btn) {
        btn.innerHTML = `<span>Auto-send Selected</span> <span id="billingAutoSendCount" class="bg-white/20 px-1.5 rounded text-[10px] hidden">0</span>`;
        btn.disabled = false;
        updateAutoSendButton(); // Reset state
    }
    
    showToast(`Auto-send completed. ${sentCount}/${total} sent.`, "success");
}

function copyToClipboard(text) {
    if (!navigator.clipboard) {
         // Fallback
         const ta = document.createElement('textarea');
         ta.value = text;
         document.body.appendChild(ta);
         ta.select();
         document.execCommand('copy');
         document.body.removeChild(ta);
         return Promise.resolve();
    }
    return navigator.clipboard.writeText(text);
}

/**
 * Initializes the Billing tab by loading cached data, wiring events,
 * and rendering the default view state.
 */
export function initBillingFeature() {
    renderBillingCalendar();
    cloneSelectOptions("wing", "billingWingSelect");
    setupModalEvents();
    refreshBillingCalendarCoverage();

    const { prevButtons, nextButtons } = getBillingCalendarControlNodes();
    prevButtons.forEach((prevBtn) => {
        prevBtn.addEventListener("click", () => {
            setBillingCalendarPage(billingState.calendarPage + 1);
        });
    });
    nextButtons.forEach((nextBtn) => {
        nextBtn.addEventListener("click", () => {
            setBillingCalendarPage(billingState.calendarPage - 1);
        });
    });

    document.addEventListener("wings:updated", () => {
        cloneSelectOptions("wing", "billingWingSelect");
        billingState.availableWings = getAvailableWings();
        renderBillingCalendar();
    });

    document.addEventListener("rentRevisions:updated", () => {
        billingRecordCache.clear();
        const modal = document.getElementById("billingWingModal");
        const isOpen = modal && !modal.classList.contains("hidden");
        if (!isOpen) return;
        if (!billingState.selectedMonthKey || !getSelectedWingNormalized()) return;
        refreshBillingData(true);
    });
}
