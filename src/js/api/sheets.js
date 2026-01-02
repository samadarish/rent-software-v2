/**
 * Google Sheets API Communication
 * 
 * Handles all communication with the Google Apps Script backend.
 */

import { clauseSections } from "../constants.js";
import { currentFlow } from "../state.js";
import { buildUnitLabel } from "../utils/formatters.js";
import { normalizeMonthKey, normalizeWing } from "../utils/normalizers.js";
import { showToast, updateConnectionIndicator } from "../utils/ui.js";
import { callAppScript, ensureAppScriptUrl } from "./appscriptClient.js";
import {
    LOCAL_KEYS,
    getLocalData,
    getLocalEntry,
    setLocalData,
    upsertLocalListItem,
    removeLocalListItem,
} from "./localStore.js";
import { enqueueSyncJob } from "./syncManager.js";

const LOCAL_REVALIDATE_MS = 5 * 60 * 1000;

function shouldRevalidate(entry, force) {
    if (force) return true;
    if (!entry) return true;
    const updatedAt = Number(entry.updated_at || entry.updatedAt || 0);
    if (!updatedAt) return true;
    return Date.now() - updatedAt > LOCAL_REVALIDATE_MS;
}

function normalizeBooleanValue(value) {
    if (value === true || value === false) return value;
    if (typeof value === "string") {
        return value.toLowerCase() !== "false" && value !== "";
    }
    return !!value;
}

function normalizeMonthRange(fromMonth, toMonth) {
    let from = normalizeMonthKey(fromMonth || "");
    let to = normalizeMonthKey(toMonth || "");
    if (from && to && from > to) {
        [from, to] = [to, from];
    }
    return { from, to };
}

function getBillMonthKey(bill) {
    if (!bill || typeof bill !== "object") return "";
    return normalizeMonthKey(
        bill.monthKey ||
            bill.month_key ||
            bill.month ||
            bill.monthLabel ||
            bill.month_label ||
            ""
    );
}

function getMonthIndex(monthKey) {
    const normalized = normalizeMonthKey(monthKey || "");
    const match = normalized.match(/^(\d{4})-(\d{2})$/);
    if (!match) return null;
    return Number(match[1]) * 12 + (Number(match[2]) - 1);
}

function isWithinRecentMonths(monthKey, monthsBack) {
    const limit = Number(monthsBack) || 0;
    if (!limit || limit <= 0) return true;
    const idx = getMonthIndex(monthKey);
    if (idx === null) return true;
    const now = new Date();
    const currentIdx = now.getFullYear() * 12 + now.getMonth();
    return idx >= currentIdx - (limit - 1);
}

function formatMonthLabelForDisplay(monthKey) {
    const normalized = normalizeMonthKey(monthKey || "");
    const match = normalized.match(/^(\d{4})-(\d{2})$/);
    if (!match) return normalized;
    const monthIndex = Number(match[2]) - 1;
    if (monthIndex < 0 || monthIndex > 11) return normalized;
    const monthNames = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ];
    return `${monthNames[monthIndex]} ${match[1]}`;
}

function getLatestRentFromRevisions(revisions = []) {
    let latestMonth = "";
    let latestCreated = 0;
    let latestAmount = null;
    revisions.forEach((rev) => {
        const monthKey = normalizeMonthKey(rev?.effective_month || rev?.effectiveMonth || "");
        const createdRaw = rev?.created_at || rev?.createdAt || "";
        const createdAt = createdRaw ? new Date(createdRaw).getTime() : 0;
        const monthCompare = (monthKey || "").localeCompare(latestMonth || "");
        if (latestAmount === null || monthCompare > 0 || (monthCompare === 0 && createdAt > latestCreated)) {
            latestMonth = monthKey || "";
            latestCreated = createdAt;
            const amount = Number(rev?.rent_amount ?? rev?.rentAmount);
            latestAmount = Number.isNaN(amount) ? 0 : amount;
        }
    });
    return latestAmount;
}

function resolveBillPaidFlag(bill = {}) {
    const remainingRaw = bill?.remainingAmount ?? bill?.remaining_amount;
    const amountPaidRaw = bill?.amountPaid ?? bill?.amount_paid;
    const isPaidRaw = bill?.isPaid ?? bill?.is_paid;
    const hasPaidRaw = bill?.hasPaid ?? bill?.has_paid;

    if (isPaidRaw !== null && isPaidRaw !== undefined && isPaidRaw !== "") {
        return normalizeBooleanValue(isPaidRaw);
    }
    if (hasPaidRaw !== null && hasPaidRaw !== undefined && hasPaidRaw !== "") {
        return normalizeBooleanValue(hasPaidRaw);
    }
    if (remainingRaw !== null && remainingRaw !== undefined && remainingRaw !== "") {
        return (Number(remainingRaw) || 0) <= 0;
    }
    if (amountPaidRaw !== null && amountPaidRaw !== undefined && amountPaidRaw !== "") {
        const totalAmount = Number(bill?.totalAmount ?? bill?.total_amount) || 0;
        const remaining = Math.max(0, totalAmount - (Number(amountPaidRaw) || 0));
        if (totalAmount > 0) {
            return remaining <= 0;
        }
    }
    return null;
}

function deriveBillPaymentStateLocal(bill = {}) {
    const totalAmount = Number(bill?.total_amount ?? bill?.totalAmount) || 0;
    const amountPaidRaw = bill?.amount_paid ?? bill?.amountPaid;
    const isPaidRaw = bill?.is_paid ?? bill?.isPaid;
    const hasAmountPaid =
        amountPaidRaw !== "" && amountPaidRaw !== null && amountPaidRaw !== undefined;
    const hasIsPaid = isPaidRaw !== "" && isPaidRaw !== null && isPaidRaw !== undefined;
    const amountPaid = hasAmountPaid ? Number(amountPaidRaw) || 0 : null;
    const isPaid = hasIsPaid ? normalizeBooleanValue(isPaidRaw) : null;
    const paidByTotal = totalAmount <= 0;
    const paidByAmount = amountPaid !== null && amountPaid + 0.005 >= totalAmount;
    let derivedIsPaid = isPaid;
    if (derivedIsPaid === null) {
        if (amountPaid !== null) {
            derivedIsPaid = paidByTotal || paidByAmount;
        } else if (paidByTotal) {
            derivedIsPaid = true;
        }
    }
    const remainingAmount =
        amountPaid !== null
            ? Math.max(0, totalAmount - amountPaid)
            : derivedIsPaid === true
                ? 0
                : null;
    return { totalAmount, amountPaid, isPaid: derivedIsPaid, remainingAmount };
}

function filterBillsByOptions(bills, { status, fromMonth, toMonth, monthsBack } = {}) {
    if (!Array.isArray(bills) || !bills.length) return [];
    const normalizedStatus = (status || "").toString().trim().toLowerCase();
    const { from, to } = normalizeMonthRange(fromMonth, toMonth);
    const recentLimit = Number(monthsBack) || 0;

    return bills.filter((bill) => {
        const monthKey = getBillMonthKey(bill);
        if (from && (!monthKey || monthKey < from)) return false;
        if (to && (!monthKey || monthKey > to)) return false;
        if (recentLimit > 0 && (!monthKey || !isWithinRecentMonths(monthKey, recentLimit))) {
            return false;
        }
        if (normalizedStatus === "paid") {
            return resolveBillPaidFlag(bill) === true;
        }
        if (normalizedStatus === "pending") {
            return resolveBillPaidFlag(bill) !== true;
        }
        return true;
    });
}

function filterCoverageByRange(coverage, { fromMonth, toMonth } = {}) {
    if (!Array.isArray(coverage) || !coverage.length) return [];
    const { from, to } = normalizeMonthRange(fromMonth, toMonth);
    if (!from && !to) return coverage;
    return coverage.filter((entry) => {
        const monthKey = normalizeMonthKey(
            entry?.monthKey ||
                entry?.month_key ||
                entry?.month ||
                entry?.monthLabel ||
                entry?.month_label ||
                ""
        );
        if (!monthKey) return false;
        if (from && monthKey < from) return false;
        if (to && monthKey > to) return false;
        return true;
    });
}

async function buildGeneratedBillsFromLocalData() {
    const [
        billLinesEntry,
        tenanciesEntry,
        unitsEntry,
        tenantsEntry,
        readingsEntry,
        configEntry,
    ] = await Promise.all([
        getLocalEntry(LOCAL_KEYS.billLines),
        getLocalEntry(LOCAL_KEYS.tenancies),
        getLocalEntry(LOCAL_KEYS.units),
        getLocalEntry(LOCAL_KEYS.tenants),
        getLocalEntry(LOCAL_KEYS.tenantMonthlyReadings),
        getLocalEntry(LOCAL_KEYS.wingMonthlyConfig),
    ]);

    const hasLocalData = [
        billLinesEntry,
        tenanciesEntry,
        unitsEntry,
        tenantsEntry,
        readingsEntry,
        configEntry,
    ].some((entry) => entry && entry.value !== undefined);
    if (!hasLocalData) return null;

    const billLines = Array.isArray(billLinesEntry?.value) ? billLinesEntry.value : [];
    const tenancies = Array.isArray(tenanciesEntry?.value) ? tenanciesEntry.value : [];
    const units = Array.isArray(unitsEntry?.value) ? unitsEntry.value : [];
    const tenantDirectory = Array.isArray(tenantsEntry?.value) ? tenantsEntry.value : [];
    const readings = Array.isArray(readingsEntry?.value) ? readingsEntry.value : [];
    const configs = Array.isArray(configEntry?.value) ? configEntry.value : [];

    const tenancyMap = new Map();
    tenancies.forEach((tenancy) => {
        const id = tenancy?.tenancy_id || tenancy?.tenancyId;
        if (id) tenancyMap.set(id, tenancy);
    });

    const unitMap = new Map();
    units.forEach((unit) => {
        const id = unit?.unit_id || unit?.unitId;
        if (id) unitMap.set(id, unit);
    });

    const tenantByTenancy = new Map();
    const tenantById = new Map();
    tenantDirectory.forEach((entry) => {
        if (!entry || typeof entry !== "object") return;
        if (entry.tenancyId) tenantByTenancy.set(entry.tenancyId, entry);
        if (entry.tenantId) tenantById.set(entry.tenantId, entry);
    });

    const readingMap = new Map();
    readings.forEach((reading) => {
        if (!reading || typeof reading !== "object") return;
        const tenancyId = reading.tenancy_id || reading.tenancyId || "";
        const monthKey = normalizeMonthKey(reading.month_key || reading.monthKey || "");
        if (!tenancyId || !monthKey) return;
        readingMap.set(`${monthKey}__${tenancyId}`, reading);
    });

    const configMap = new Map();
    configs.forEach((cfg) => {
        if (!cfg || typeof cfg !== "object") return;
        const monthKey = normalizeMonthKey(cfg.month_key || cfg.monthKey || "");
        const wing = normalizeWing(cfg.wing || "");
        if (!monthKey || !wing) return;
        configMap.set(`${monthKey}__${wing}`, cfg);
    });

    const coverage = configs
        .map((cfg) => ({
            monthKey: normalizeMonthKey(cfg?.month_key || cfg?.monthKey || ""),
            wing: (cfg?.wing || "").toString().trim(),
        }))
        .filter((entry) => entry.monthKey && entry.wing);

    const bills = billLines.map((bill) => {
        const tenancyId = bill?.tenancy_id || bill?.tenancyId || "";
        const rawMonth = bill?.month_key || bill?.monthKey || "";
        const normalizedMonth = normalizeMonthKey(rawMonth);
        const monthKey = normalizedMonth || rawMonth;
        const tenancy = tenancyMap.get(tenancyId) || {};
        const unit = unitMap.get(tenancy.unit_id || tenancy.unitId) || {};
        const tenantEntry =
            tenantByTenancy.get(tenancyId) ||
            tenantById.get(tenancy.tenant_id || tenancy.tenantId) ||
            {};
        const wingValue = (unit.wing || tenantEntry.wing || bill?.wing || "").toString().trim();
        const unitNumber =
            unit.unit_number ||
            tenantEntry.unitNumber ||
            bill?.unitNumber ||
            bill?.unit_number ||
            "";
        const tenantName = tenantEntry.tenantFullName || tenantEntry.tenantName || "";
        const tenantKey =
            tenancy.grn_number ||
            tenantEntry.grnNumber ||
            tenantEntry.tenantKey ||
            tenantName ||
            "";
        const readingKey = `${normalizeMonthKey(monthKey)}__${tenancyId}`;
        const reading = readingMap.get(readingKey) || {};
        const cfg =
            configMap.get(`${normalizeMonthKey(monthKey)}__${normalizeWing(wingValue)}`) || {};
        const paymentState = deriveBillPaymentStateLocal(bill);

        return {
            monthKey,
            monthLabel: formatMonthLabelForDisplay(monthKey),
            wing: wingValue,
            unitNumber,
            tenantKey,
            tenantName,
            rentAmount: Number(bill?.rent_amount ?? bill?.rentAmount) || 0,
            electricityAmount: Number(bill?.electricity_amount ?? bill?.electricityAmount) || 0,
            motorShare: Number(bill?.motor_share_amount ?? bill?.motorShare) || 0,
            sweepAmount: Number(bill?.sweep_amount ?? bill?.sweepAmount) || 0,
            totalAmount: paymentState.totalAmount,
            amountPaid: paymentState.amountPaid,
            remainingAmount: paymentState.remainingAmount,
            isPaid: paymentState.isPaid,
            included: normalizeBooleanValue(reading.included),
            payableDate:
                bill?.payable_date ||
                bill?.payableDate ||
                tenancy.rent_payable_day ||
                tenantEntry.payableDate ||
                "",
            prevReading: reading.prev_reading ?? reading.prevReading ?? "",
            newReading: reading.new_reading ?? reading.newReading ?? "",
            electricityRate: cfg.electricity_rate || "",
            sweepingPerFlat: cfg.sweeping_per_flat || "",
            motorPrev: cfg.motor_prev || "",
            motorNew: cfg.motor_new || "",
            billLineId: bill?.bill_line_id || bill?.billLineId || "",
            tenancyId,
        };
    });

    const payload = { ok: true, bills, coverage };
    await setLocalData(LOCAL_KEYS.generatedBills, payload);
    return payload;
}

async function loadGeneratedBillsBase(url) {
    const entry = await getLocalEntry(LOCAL_KEYS.generatedBills);
    const cached = entry?.value && typeof entry.value === "object" ? entry.value : null;
    if (cached && Array.isArray(cached.bills)) {
        if (shouldRevalidate(entry, false) && navigator.onLine && url) {
            callAppScript({
                url,
                action: "generatedbills",
                cache: { useLocal: false, revalidate: false },
            })
                .then(async (data) => {
                    if (data && Array.isArray(data.bills)) {
                        await setLocalData(LOCAL_KEYS.generatedBills, data);
                    }
                })
                .catch(() => null);
        }
        return cached;
    }

    const localBuilt = await buildGeneratedBillsFromLocalData();
    if (localBuilt && Array.isArray(localBuilt.bills)) {
        return localBuilt;
    }

    if (!url) return null;
    try {
        const data = await callAppScript({
            url,
            action: "generatedbills",
            cache: { useLocal: false, revalidate: false },
        });
        if (data && Array.isArray(data.bills)) {
            await setLocalData(LOCAL_KEYS.generatedBills, data);
        }
        return data;
    } catch (err) {
        return cached;
    }
}

function hasAnyMetaValue(meta = {}) {
    return [meta.electricityRate, meta.sweepingPerFlat, meta.motorPrev, meta.motorNew].some(
        (value) => value !== undefined && value !== null && value !== ""
    );
}

function buildBillingRecordFromBills(bills, monthKey, wing) {
    if (!Array.isArray(bills) || !bills.length) return null;
    const normalizedMonth = normalizeMonthKey(monthKey || "");
    const normalizedWing = normalizeWing(wing || "");
    if (!normalizedMonth || !normalizedWing) return null;

    const matches = bills.filter((bill) => {
        const billMonth = getBillMonthKey(bill);
        const billWing = normalizeWing(bill?.wing || "");
        return billMonth === normalizedMonth && billWing === normalizedWing;
    });

    if (!matches.length) return null;
    const first = matches[0] || {};
    const meta = {
        electricityRate: first.electricityRate ?? first.electricity_rate ?? "",
        sweepingPerFlat: first.sweepingPerFlat ?? first.sweeping_per_flat ?? "",
        motorPrev: first.motorPrev ?? first.motor_prev ?? "",
        motorNew: first.motorNew ?? first.motor_new ?? "",
    };
    const monthLabel = first.monthLabel || first.month_label || "";

    const tenants = matches.map((bill) => ({
        tenancyId: bill.tenancyId ?? bill.tenancy_id ?? "",
        tenantKey: bill.tenantKey ?? bill.tenant_key ?? "",
        tenantName: bill.tenantName ?? bill.tenant_name ?? "",
        wing: bill.wing ?? "",
        unitNumber: bill.unitNumber ?? bill.unit_number ?? "",
        prevReading: bill.prevReading ?? bill.prev_reading ?? "",
        newReading: bill.newReading ?? bill.new_reading ?? "",
        included: bill.included,
        rentAmount: bill.rentAmount ?? bill.rent_amount ?? "",
        payableDate: bill.payableDate ?? bill.payable_date ?? "",
    }));

    return {
        ok: true,
        monthKey: normalizedMonth,
        monthLabel,
        wing: wing || "",
        hasConfig: hasAnyMetaValue(meta),
        hasReadings: true,
        meta,
        tenants,
    };
}

async function buildBillingRecordFromLocalData(monthKey, wing) {
    const normalizedMonth = normalizeMonthKey(monthKey || "");
    const normalizedWing = normalizeWing(wing || "");
    if (!normalizedMonth || !normalizedWing) return null;

    const [
        configEntry,
        readingsEntry,
        tenanciesEntry,
        unitsEntry,
        tenantsEntry,
        rentEntry,
    ] = await Promise.all([
        getLocalEntry(LOCAL_KEYS.wingMonthlyConfig),
        getLocalEntry(LOCAL_KEYS.tenantMonthlyReadings),
        getLocalEntry(LOCAL_KEYS.tenancies),
        getLocalEntry(LOCAL_KEYS.units),
        getLocalEntry(LOCAL_KEYS.tenants),
        getLocalEntry(LOCAL_KEYS.rentRevisionsAll),
    ]);

    const hasLocalData = [
        configEntry,
        readingsEntry,
        tenanciesEntry,
        unitsEntry,
        tenantsEntry,
    ].some((entry) => entry && entry.value !== undefined);
    if (!hasLocalData) return null;

    const configs = Array.isArray(configEntry?.value) ? configEntry.value : [];
    const readings = Array.isArray(readingsEntry?.value) ? readingsEntry.value : [];
    const tenancies = Array.isArray(tenanciesEntry?.value) ? tenanciesEntry.value : [];
    const units = Array.isArray(unitsEntry?.value) ? unitsEntry.value : [];
    const tenantDirectory = Array.isArray(tenantsEntry?.value) ? tenantsEntry.value : [];
    const rentRevisions = Array.isArray(rentEntry?.value) ? rentEntry.value : [];

    const configRow = configs.find(
        (cfg) =>
            normalizeMonthKey(cfg.month_key || cfg.monthKey || "") === normalizedMonth &&
            normalizeWing(cfg.wing || "") === normalizedWing
    );
    const config = configRow || {
        month_key: normalizedMonth,
        wing,
        electricity_rate: "",
        sweeping_per_flat: "",
        motor_prev: "",
        motor_new: "",
        motor_units: "",
    };

    const unitMap = new Map();
    units.forEach((unit) => {
        const id = unit?.unit_id || unit?.unitId;
        if (!id) return;
        if (normalizeWing(unit.wing || "") === normalizedWing) {
            unitMap.set(id, unit);
        }
    });

    const tenancyMap = new Map();
    tenancies.forEach((tenancy) => {
        const id = tenancy?.tenancy_id || tenancy?.tenancyId;
        const unitId = tenancy?.unit_id || tenancy?.unitId;
        if (!id) return;
        if (unitId && unitMap.has(unitId)) {
            tenancyMap.set(id, tenancy);
        }
    });

    const tenantByTenancy = new Map();
    const tenantById = new Map();
    tenantDirectory.forEach((entry) => {
        if (!entry || typeof entry !== "object") return;
        if (entry.tenancyId) tenantByTenancy.set(entry.tenancyId, entry);
        if (entry.tenantId) tenantById.set(entry.tenantId, entry);
    });

    const matchedReadings = readings.filter((reading) => {
        const tenancyId = reading?.tenancy_id || reading?.tenancyId || "";
        const readingMonth = normalizeMonthKey(reading?.month_key || reading?.monthKey || "");
        if (!tenancyId || readingMonth !== normalizedMonth) return false;
        if (tenancyMap.has(tenancyId)) return true;
        const entry = tenantByTenancy.get(tenancyId);
        return entry ? normalizeWing(entry.wing || "") === normalizedWing : false;
    });

    const revisionsByTenancy = new Map();
    rentRevisions.forEach((rev) => {
        const tenancyId = rev?.tenancy_id || rev?.tenancyId;
        if (!tenancyId) return;
        const list = revisionsByTenancy.get(tenancyId) || [];
        list.push(rev);
        revisionsByTenancy.set(tenancyId, list);
    });

    const tenants = matchedReadings.map((reading) => {
        const tenancyId = reading?.tenancy_id || reading?.tenancyId || "";
        const tenancy = tenancyMap.get(tenancyId) || {};
        const unit = unitMap.get(tenancy.unit_id || tenancy.unitId) || {};
        const tenantEntry =
            tenantByTenancy.get(tenancyId) ||
            tenantById.get(tenancy.tenant_id || tenancy.tenantId) ||
            {};
        const revisions = revisionsByTenancy.get(tenancyId) || [];
        const latestRent = revisions.length ? getLatestRentFromRevisions(revisions) : null;
        const fallbackRent =
            latestRent !== null
                ? latestRent
                : tenantEntry.rentAmount ?? tenantEntry.currentRent ?? "";
        const overrideRent = reading.override_rent ?? reading.overrideRent ?? "";
        const rentAmount = overrideRent || fallbackRent || "";
        const tenantName = tenantEntry.tenantFullName || tenantEntry.tenantName || "";
        const tenantKey =
            tenancy.grn_number ||
            tenantEntry.grnNumber ||
            tenantEntry.tenantKey ||
            tenantName ||
            "";

        return {
            tenancyId,
            tenantKey,
            tenantName,
            wing: unit.wing || tenantEntry.wing || wing || "",
            unitNumber: unit.unit_number || tenantEntry.unitNumber || "",
            prevReading: reading.prev_reading ?? reading.prevReading ?? "",
            newReading: reading.new_reading ?? reading.newReading ?? "",
            included: normalizeBooleanValue(reading.included),
            override_rent: overrideRent,
            rentAmount,
            payableDate: tenancy.rent_payable_day || tenantEntry.payableDate || "",
            direction: unit.direction || tenantEntry.direction || "",
            floor: unit.floor || tenantEntry.floor || "",
            meterNumber: unit.meter_number || tenantEntry.meterNumber || "",
        };
    });

    const computedMotorUnits =
        (Number(config.motor_new || 0) || 0) - (Number(config.motor_prev || 0) || 0);
    const motorUnitsRaw = config.motor_units;
    const motorUnits =
        motorUnitsRaw !== "" && motorUnitsRaw !== null && motorUnitsRaw !== undefined
            ? motorUnitsRaw
            : computedMotorUnits;

    return {
        ok: true,
        monthKey: normalizedMonth,
        monthLabel: formatMonthLabelForDisplay(normalizedMonth),
        wing: config.wing || wing || "",
        hasConfig: !!configRow,
        hasReadings: tenants.length > 0,
        meta: {
            month_key: config.month_key || normalizedMonth,
            wing: config.wing || wing || "",
            electricityRate: config.electricity_rate || "",
            sweepingPerFlat: config.sweeping_per_flat || "",
            motorPrev: config.motor_prev || "",
            motorNew: config.motor_new || "",
            motor_units: motorUnits,
        },
        tenants,
    };
}

function createLocalId(prefix = "local") {
    try {
        if (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function") {
            return `${prefix}-${globalThis.crypto.randomUUID()}`;
        }
    } catch (err) {
        // Ignore UUID generation failures.
    }
    return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeWingValue(value) {
    return normalizeWing(value || "");
}

async function updateLocalWingsList(list) {
    const normalized = Array.isArray(list)
        ? list
              .map((w) => (w || "").toString().trim())
              .filter(Boolean)
        : [];
    const seen = new Set();
    const deduped = normalized.filter((wing) => {
        const key = normalizeWingValue(wing);
        if (!key || seen.has(key)) return false;
        seen.add(key);
        return true;
    });
    await setLocalData(LOCAL_KEYS.wings, deduped);
    applyWingOptions(deduped);
    return deduped;
}

async function updateLocalUnitsList(list) {
    const next = Array.isArray(list) ? list : [];
    await setLocalData(LOCAL_KEYS.units, next);
    document.dispatchEvent(new CustomEvent("units:updated", { detail: next }));
    return next;
}

async function updateLocalLandlordsList(list) {
    const next = Array.isArray(list) ? list : [];
    await setLocalData(LOCAL_KEYS.landlords, next);
    document.dispatchEvent(new CustomEvent("landlords:updated", { detail: next }));
    return next;
}

async function updateLocalTenantsList(list) {
    const next = Array.isArray(list) ? list : [];
    await setLocalData(LOCAL_KEYS.tenants, next);
    document.dispatchEvent(new CustomEvent("tenants:updated", { detail: next }));
    return next;
}

async function updateLocalPaymentsList(list) {
    const next = Array.isArray(list) ? list : [];
    await setLocalData(LOCAL_KEYS.payments, next);
    return next;
}

async function updateLocalGeneratedBills(data) {
    const payload = data && typeof data === "object" ? data : { bills: [], coverage: [] };
    await setLocalData(LOCAL_KEYS.generatedBills, payload);
    return payload;
}

function buildBillKey(monthKey, tenancyId) {
    return `${normalizeMonthKey(monthKey || "")}__${tenancyId || ""}`;
}

function resolveBillAmounts({ totalAmount, amountPaid }) {
    const total = Number(totalAmount) || 0;
    const paid = Number(amountPaid) || 0;
    const isPaid = total <= 0 || paid + 0.005 >= total;
    const remainingAmount = isPaid ? 0 : Math.max(0, total - paid);
    return { totalAmount: total, amountPaid: paid, isPaid, remainingAmount };
}

async function buildLocalBillsFromPayload(payload = {}) {
    const monthKey = normalizeMonthKey(payload.monthKey || "");
    const wing = payload.wing || "";
    const monthLabel = payload.monthLabel || "";
    const meta = payload.meta || {};
    const tenantsPayload = Array.isArray(payload.tenants) ? payload.tenants : [];
    const existingData = await getLocalData(LOCAL_KEYS.generatedBills, { bills: [], coverage: [] });
    const existingBills = Array.isArray(existingData?.bills) ? existingData.bills : [];
    const existingByKey = new Map();
    existingBills.forEach((bill) => {
        const key = buildBillKey(bill.monthKey || bill.month_key, bill.tenancyId || bill.tenancy_id);
        if (key) existingByKey.set(key, bill);
    });

    const tenantDirectory = await getLocalData(LOCAL_KEYS.tenants, []);
    const tenancyMap = new Map();
    (Array.isArray(tenantDirectory) ? tenantDirectory : []).forEach((t) => {
        if (t && t.tenancyId) tenancyMap.set(t.tenancyId, t);
    });

    const newBills = tenantsPayload.map((tenant) => {
        const tenancyId = tenant.tenancyId || tenant.tenancy_id || "";
        const billKey = buildBillKey(monthKey, tenancyId);
        const existing = existingByKey.get(billKey) || {};
        const tenantEntry = tenancyMap.get(tenancyId) || {};
        const tenantName = tenant.name || tenant.tenantName || tenantEntry.tenantFullName || "";
        const tenantKey =
            tenant.tenantKey || tenant.grn || tenantEntry.grnNumber || tenantEntry.tenantKey || tenantName || "";
        const unitNumber = tenant.unitNumber || tenant.unit_number || tenantEntry.unitNumber || "";
        const wingValue = wing || tenantEntry.wing || tenantEntry.wing || "";
        const totalAmount = tenant.totalAmount ?? tenant.total_amount ?? 0;
        const amountPaid = existing.amountPaid ?? existing.amount_paid ?? 0;
        const state = resolveBillAmounts({ totalAmount, amountPaid });

        return {
            monthKey,
            monthLabel,
            wing: wingValue,
            unitNumber,
            tenantKey,
            tenantName,
            rentAmount: Number(tenant.rentAmount ?? tenant.rent_amount) || 0,
            electricityAmount: Number(tenant.electricityAmount ?? tenant.electricity_amount) || 0,
            motorShare: Number(tenant.motorShare ?? tenant.motor_share_amount) || 0,
            sweepAmount: Number(tenant.sweepAmount ?? tenant.sweep_amount) || 0,
            totalAmount: state.totalAmount,
            amountPaid: state.amountPaid,
            remainingAmount: state.remainingAmount,
            isPaid: state.isPaid,
            included: tenant.included,
            payableDate: tenant.payableDate || tenant.payable_date || tenantEntry.payableDate || "",
            prevReading: tenant.prevReading ?? tenant.prev_reading ?? "",
            newReading: tenant.newReading ?? tenant.new_reading ?? "",
            electricityRate: meta.electricityRate ?? meta.electricity_rate ?? "",
            sweepingPerFlat: meta.sweepingPerFlat ?? meta.sweeping_per_flat ?? "",
            motorPrev: meta.motorPrev ?? meta.motor_prev ?? "",
            motorNew: meta.motorNew ?? meta.motor_new ?? "",
            billLineId: existing.billLineId || existing.bill_line_id || createLocalId("bill"),
            tenancyId,
        };
    });

    const filteredExisting = existingBills.filter((bill) => {
        const key = buildBillKey(bill.monthKey || bill.month_key, bill.tenancyId || bill.tenancy_id);
        return !newBills.some((nb) => buildBillKey(nb.monthKey, nb.tenancyId) === key);
    });

    const combinedBills = filteredExisting.concat(newBills);
    const coverageExisting = Array.isArray(existingData?.coverage) ? existingData.coverage : [];
    const coverageKey = `${monthKey}__${normalizeWingValue(wing)}`;
    const filteredCoverage = coverageExisting.filter((c) => {
        const key = `${normalizeMonthKey(c.monthKey || c.month_key || "")}__${normalizeWingValue(c.wing || "")}`;
        return key !== coverageKey;
    });
    const coverage = monthKey && wing ? filteredCoverage.concat([{ monthKey, wing }]) : filteredCoverage;

    return { bills: combinedBills, coverage };
}

function buildLocalPaymentRecord(payload = {}) {
    const id = payload.id || createLocalId("payment");
    const date = payload.date || new Date().toISOString().slice(0, 10);
    const createdAt = payload.createdAt || new Date().toISOString();
    return {
        id,
        date,
        amount: Number(payload.amount) || 0,
        mode: payload.mode || "",
        reference: payload.reference || "",
        notes: payload.notes || "",
        tenantKey: payload.tenantKey || payload.tenantName || "",
        tenantName: payload.tenantName || "",
        wing: payload.wing || "",
        attachmentName: payload.attachmentName || "",
        attachmentUrl: payload.attachmentUrl || "",
        attachmentId: payload.attachmentId || "",
        monthKey: payload.monthKey || "",
        monthLabel: payload.monthLabel || "",
        billTotal: Number(payload.billTotal) || 0,
        rentAmount: Number(payload.rentAmount) || 0,
        electricityAmount: Number(payload.electricityAmount) || 0,
        motorShare: Number(payload.motorShare ?? payload.motorAmount) || 0,
        sweepAmount: Number(payload.sweepAmount) || 0,
        prevReading: payload.prevReading || "",
        newReading: payload.newReading || "",
        payableDate: payload.payableDate || "",
        billLineId: payload.billLineId || payload.bill_line_id || "",
        tenancyId: payload.tenancyId || payload.tenancy_id || "",
        createdAt,
    };
}

function updateBillsWithPayment(bills = [], payments = [], billLineId) {
    if (!billLineId) return bills;
    const targetId = String(billLineId);
    const paidTotal = payments
        .filter((p) => String(p.billLineId || p.bill_line_id || "") === targetId)
        .reduce((sum, p) => sum + (Number(p.amount) || 0), 0);

    return bills.map((bill) => {
        const id = bill.billLineId || bill.bill_line_id || "";
        if (String(id) !== targetId) return bill;
        const state = resolveBillAmounts({
            totalAmount: bill.totalAmount ?? bill.total_amount ?? 0,
            amountPaid: paidTotal,
        });
        return {
            ...bill,
            amountPaid: state.amountPaid,
            amount_paid: state.amountPaid,
            remainingAmount: state.remainingAmount,
            remaining_amount: state.remainingAmount,
            isPaid: state.isPaid,
            is_paid: state.isPaid,
        };
    });
}

function coalesceValue(...values) {
    for (const value of values) {
        if (value !== undefined && value !== null && value !== "") {
            return value;
        }
    }
    return "";
}

function buildTenantEntryFromPayload(payload, { existing = {}, unit = {}, landlord = {}, tenantId, tenancyId, unitId, landlordId } = {}) {
    const template = payload?.templateData || {};
    const updates = payload?.updates || {};

    const wing = coalesceValue(payload.wing, updates.wing, template.wing, unit.wing, existing.wing);
    const unitNumber = coalesceValue(
        updates.unitNumber,
        payload.unit_number,
        template.unit_number,
        unit.unit_number,
        existing.unitNumber
    );
    const floor = coalesceValue(
        updates.floor,
        payload.floor_of_building,
        template.floor_of_building,
        unit.floor,
        existing.floor
    );
    const direction = coalesceValue(
        updates.direction,
        template.direction_build,
        unit.direction,
        existing.direction
    );
    const meterNumber = coalesceValue(
        updates.meterNumber,
        template.meter_number,
        unit.meter_number,
        existing.meterNumber
    );

    const tenantFullName = coalesceValue(
        updates.tenantFullName,
        template.Tenant_Full_Name,
        existing.tenantFullName
    );
    const tenantOccupation = coalesceValue(
        payload.Tenant_occupation,
        updates.tenantOccupation,
        template.tenant_occupation,
        existing.tenantOccupation
    );
    const tenantPermanentAddress = coalesceValue(
        updates.tenantPermanentAddress,
        template.Tenant_Permanent_Address,
        existing.tenantPermanentAddress
    );
    const tenantMobile = coalesceValue(updates.tenantMobile, template.tenant_mobile, existing.tenantMobile);
    const tenantAadhaar = coalesceValue(updates.tenantAadhaar, template.tenant_Aadhar, existing.tenantAadhaar);
    const grnNumber = coalesceValue(updates.grnNumber, template["GRN number"], payload.grn, existing.grnNumber);

    const rentAmount = Number(coalesceValue(updates.rentAmount, template.rent_amount, existing.rentAmount)) || 0;
    const payableDate = coalesceValue(updates.payableDate, template.payable_date_raw, existing.payableDate);

    const agreementDate = coalesceValue(updates.agreementDateRaw, template.agreement_date_raw, template.agreement_date, existing.agreementDate);
    const tenancyCommencement = coalesceValue(
        updates.tenancyCommencementRaw,
        template.tenancy_comm_raw,
        template.tenancy_comm,
        existing.tenancyCommencement
    );
    const tenancyEndDate = coalesceValue(updates.tenancyEndRaw, template.tenancy_end_raw, existing.tenancyEndDate);

    const activeTenant =
        typeof updates.activeTenant !== "undefined" ? updates.activeTenant : existing.activeTenant ?? true;

    const landlordName = coalesceValue(landlord.name, template.Landlord_name, existing.landlordName);
    const landlordAadhaar = coalesceValue(landlord.aadhaar, template.landlord_aadhar, existing.landlordAadhaar);
    const landlordAddress = coalesceValue(landlord.address, template.landlord_address, existing.landlordAddress);

    const historyEntry = {
        tenancyId,
        unitLabel: buildUnitLabel(unitId ? unit : { wing, unit_number: unitNumber }),
        startDate: tenancyCommencement || agreementDate || "",
        endDate: tenancyEndDate || "",
        status: activeTenant ? "ACTIVE" : "ENDED",
        grnNumber,
    };

    const existingHistory = Array.isArray(existing.tenancyHistory) ? existing.tenancyHistory : [];
    const mergedHistory = [historyEntry, ...existingHistory.filter((h) => h.tenancyId !== tenancyId)];

    const templateData = {
        ...(existing.templateData || {}),
        ...(template || {}),
        tenant_id: tenantId,
        tenancy_id: tenancyId,
        unit_id: unitId,
        landlord_id: landlordId,
        wing,
        unit_number: unitNumber,
        floor_of_building: floor,
        direction_build: direction,
        meter_number: meterNumber,
        "GRN number": grnNumber,
        Tenant_Full_Name: tenantFullName,
        Tenant_Permanent_Address: tenantPermanentAddress,
        tenant_Aadhar: tenantAadhaar,
        tenant_mobile: tenantMobile,
        tenant_occupation: tenantOccupation,
        Landlord_name: landlordName,
        landlord_address: landlordAddress,
        landlord_aadhar: landlordAadhaar,
        rent_amount: rentAmount,
        payable_date_raw: payableDate,
        agreement_date_raw: agreementDate,
        tenancy_comm_raw: tenancyCommencement,
        tenancy_end_raw: tenancyEndDate,
    };

    return {
        ...existing,
        tenantId,
        tenancyId,
        grnNumber,
        tenantFullName,
        tenantOccupation,
        tenantPermanentAddress,
        tenantAadhaar,
        tenantMobile,
        wing,
        unitId,
        unitNumber,
        floor,
        direction,
        meterNumber,
        landlordId,
        landlordName,
        landlordAadhaar,
        landlordAddress,
        unitOccupied: true,
        rentAmount,
        currentRent: rentAmount,
        payableDate,
        securityDeposit: coalesceValue(updates.securityDeposit, template.secu_depo, existing.securityDeposit),
        rentIncrease: coalesceValue(updates.rentIncreaseAmount, template.rent_inc, existing.rentIncrease),
        rentRevisionNumber: coalesceValue(updates.rentRevisionNumber, template.rent_rev_number, existing.rentRevisionNumber),
        rentRevisionUnit: coalesceValue(updates.rentRevisionUnit, template["rent_rev year_mon"], existing.rentRevisionUnit),
        tenantNoticeMonths: coalesceValue(updates.tenantNoticeMonths, template.notice_num_t, existing.tenantNoticeMonths),
        landlordNoticeMonths: coalesceValue(updates.landlordNoticeMonths, template.notice_num_l, existing.landlordNoticeMonths),
        lateRentPerDay: coalesceValue(updates.lateRentPerDay, template.late_rent, existing.lateRentPerDay),
        lateGracePeriodDays: coalesceValue(updates.lateGracePeriodDays, template.late_days, existing.lateGracePeriodDays),
        agreementDate,
        tenancyCommencement,
        tenancyEndDate,
        activeTenant,
        vacateReason: coalesceValue(updates.vacateReason, existing.vacateReason),
        family: Array.isArray(payload.familyMembers) ? payload.familyMembers : existing.family || [],
        tenancyHistory: mergedHistory,
        templateData,
    };
}

function buildTenancyRecordFromEntry(entry = {}, { tenantId, tenancyId, unitId, landlordId } = {}) {
    const createdAt = entry.createdAt || entry.created_at || new Date().toISOString();
    return {
        tenancy_id: tenancyId,
        tenant_id: tenantId,
        grn_number: entry.grnNumber || "",
        unit_id: unitId,
        landlord_id: landlordId || "",
        agreement_date: entry.agreementDate || "",
        commencement_date: entry.tenancyCommencement || "",
        end_date: entry.tenancyEndDate || "",
        status: entry.activeTenant ? "ACTIVE" : "ENDED",
        vacate_reason: entry.vacateReason || "",
        security_deposit: entry.securityDeposit || "",
        rent_payable_day: entry.payableDate || "",
        tenant_notice_months: entry.tenantNoticeMonths || "",
        landlord_notice_months: entry.landlordNoticeMonths || "",
        pet_policy: entry.petPolicy || "",
        late_rent_per_day: entry.lateRentPerDay || "",
        late_grace_days: entry.lateGracePeriodDays || "",
        rent_revision_unit: entry.rentRevisionUnit || "",
        rent_revision_number: entry.rentRevisionNumber || "",
        created_at: createdAt,
        rent_increase_amount: entry.rentIncrease || "",
    };
}

function buildFamilyRecordsFromEntry(entry = {}, tenantId) {
    const now = new Date().toISOString();
    const members = Array.isArray(entry.family) ? entry.family : [];
    return members.map((member) => ({
        member_id: member.member_id || member.memberId || createLocalId("family"),
        tenant_id: tenantId,
        name: member.name || "",
        relationship: member.relationship || "",
        occupation: member.occupation || "",
        aadhaar: member.aadhaar || "",
        address: member.address || "",
        created_at: member.created_at || now,
    }));
}


async function runWriteAction({
    url,
    action,
    payload,
    params,
    method = "POST",
    queuedMessage = "Saved locally. Sync pending.",
    fallback = { ok: true, queued: true },
    localUpdate,
}) {
    let localResult = null;
    if (typeof localUpdate === "function") {
        try {
            localResult = await localUpdate();
        } catch (err) {
            console.warn(`Local update failed for ${action}`, err);
        }
    }

    await enqueueSyncJob({ action, payload, params, method });
    if (queuedMessage) {
        showToast(queuedMessage, "warning");
    }

    if (localResult && typeof localResult === "object" && localResult.queued === undefined) {
        localResult.queued = true;
    }

    return localResult || fallback;
}

// Detect accidental double-loading so we can surface clearer diagnostics
if (globalThis.__sheetsApiLoaded) {
    console.warn(
        "sheets.js loaded more than once; check duplicate script imports to avoid redeclaration errors"
    );
} else {
    globalThis.__sheetsApiLoaded = true;
}

/**
 * Fetches available wings from Google Sheets and populates the dropdown
 */
function applyWingOptions(wings = []) {
    const sel = document.getElementById("wing");
    if (sel) {
        sel.innerHTML = '<option value="">Select wing</option>';
        wings.forEach((w) => {
            const opt = document.createElement("option");
            opt.value = w;
            opt.textContent = w;
            sel.appendChild(opt);
        });

        const billingSel = document.getElementById("billingWingSelect");
        if (billingSel) {
            billingSel.innerHTML = sel.innerHTML;
        }
        document.dispatchEvent(new CustomEvent("wings:updated", { detail: wings }));
    }
}

export async function fetchWingsFromSheet(force = false) {
    const entry = await getLocalEntry(LOCAL_KEYS.wings);
    const cached = Array.isArray(entry?.value) ? entry.value : null;
    if (cached) {
        applyWingOptions(cached);
        if (!shouldRevalidate(entry, force)) {
            return { wings: cached };
        }
    }

    const url = ensureAppScriptUrl({
        onMissing: () =>
            updateConnectionIndicator(
                navigator.onLine ? "online" : "offline",
                "Set Apps Script URL"
            ),
    });
    if (!url) return { wings: cached || [] };

    const runFetch = async () => {
        const data = await callAppScript({
            url,
            action: "wings",
            cache: { useLocal: false, revalidate: false },
        });
        if (Array.isArray(data?.wings)) {
            await setLocalData(LOCAL_KEYS.wings, data.wings);
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
            applyWingOptions(data.wings);
            return { wings: data.wings };
        }
        return { wings: cached || [] };
    };

    if (cached) {
        if (navigator.onLine) {
            runFetch().catch((e) => {
                console.warn("Could not fetch wings", e);
                updateConnectionIndicator(
                    navigator.onLine ? "online" : "offline",
                    navigator.onLine ? "Wing sync failed" : "Offline"
                );
            });
        }
        return { wings: cached };
    }

    try {
        return await runFetch();
    } catch (e) {
        console.warn("Could not fetch wings", e);
        updateConnectionIndicator(
            navigator.onLine ? "online" : "offline",
            navigator.onLine ? "Wing sync failed" : "Offline"
        );
        return { wings: cached || [] };
    }
}

/**
 * Persists a new wing value to Google Sheets and refreshes UI indicators.
 * @param {string} wing - Wing name entered by the user.
 * @returns {Promise<object>} API response shape from the Apps Script endpoint.
 */
export async function addWingToSheet(wing) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });

    const cleaned = (wing || "").trim();
    if (!cleaned) {
        showToast("Please enter a wing name", "warning");
        return { ok: false };
    }

    try {
        const localUpdate = async () => {
            const existing = await getLocalData(LOCAL_KEYS.wings, []);
            const next = Array.isArray(existing) ? existing.slice() : [];
            const normalized = normalizeWingValue(cleaned);
            const hasWing = next.some((w) => normalizeWingValue(w) === normalized);
            if (!hasWing) next.push(cleaned);
            await updateLocalWingsList(next);
            return { ok: true, wing: cleaned, wings: next };
        };
        const data = await runWriteAction({
            url,
            action: "addWing",
            payload: { wing: cleaned },
            queuedMessage: "Wing saved locally. Sync pending.",
            fallback: { ok: true, queued: true, message: "Wing queued" },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("addWingToSheet error", e);
        showToast("Failed to save wing", "error");
        updateConnectionIndicator(
            navigator.onLine ? "online" : "offline",
            navigator.onLine ? "Wing save failed" : "Offline"
        );
        return { ok: false };
    }
}

/**
 * Removes an existing wing from Google Sheets.
 * @param {string} wing - Wing identifier to delete.
 * @returns {Promise<object>} Result payload from the Apps Script endpoint.
 */
export async function removeWingFromSheet(wing) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });

    const cleaned = (wing || "").trim();
    if (!cleaned) {
        showToast("Please enter a wing to remove", "warning");
        return { ok: false };
    }

    try {
        const localUpdate = async () => {
            const existing = await getLocalData(LOCAL_KEYS.wings, []);
            const normalized = normalizeWingValue(cleaned);
            const next = Array.isArray(existing)
                ? existing.filter((w) => normalizeWingValue(w) !== normalized)
                : [];
            await updateLocalWingsList(next);
            return { ok: true, wing: cleaned, wings: next };
        };
        const data = await runWriteAction({
            url,
            action: "removeWing",
            payload: { wing: cleaned },
            queuedMessage: "Wing removal queued. Sync pending.",
            fallback: { ok: true, queued: true, wing: cleaned },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("removeWingFromSheet error", e);
        showToast("Failed to remove wing", "error");
        return { ok: false };
    }
}

/**
 * Retrieves saved landlord profiles from Google Sheets.
 * @returns {Promise<{landlords: Array}>} Fetched landlord collection (empty on failure).
 */
export async function fetchLandlordsFromSheet(force = false) {
    const entry = await getLocalEntry(LOCAL_KEYS.landlords);
    const cached = Array.isArray(entry?.value) ? entry.value : null;
    if (cached) {
        document.dispatchEvent(new CustomEvent("landlords:updated", { detail: cached }));
        if (!shouldRevalidate(entry, force)) {
            return { landlords: cached };
        }
    }

    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view landlords", "warning"),
    });
    if (!url) return { landlords: cached || [] };

    const runFetch = async () => {
        const data = await callAppScript({
            url,
            action: "landlords",
            cache: { useLocal: false, revalidate: false },
        });
        if (Array.isArray(data?.landlords)) {
            await setLocalData(LOCAL_KEYS.landlords, data.landlords);
            document.dispatchEvent(new CustomEvent("landlords:updated", { detail: data.landlords }));
            return { landlords: data.landlords };
        }
        return { landlords: cached || [] };
    };

    if (cached) {
        if (navigator.onLine) {
            runFetch().catch((e) => {
                console.error("fetchLandlordsFromSheet error", e);
            });
        }
        return { landlords: cached };
    }

    try {
        return await runFetch();
    } catch (e) {
        console.error("fetchLandlordsFromSheet error", e);
        showToast("Could not fetch landlords", "error");
        return { landlords: cached || [] };
    }
}

/**
 * Saves or updates a landlord configuration record.
 * @param {object} payload - Landlord details (id, name, aadhaar, address, defaults).
 * @returns {Promise<object>} Apps Script response with ok/message flags.
 */
export async function saveLandlordConfig(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    const nextPayload = payload && typeof payload === "object" ? { ...payload } : {};
    if (!nextPayload.landlordId) {
        nextPayload.landlordId = createLocalId("landlord");
    }

    try {
        const localUpdate = async () => {
            const landlord = {
                landlord_id: nextPayload.landlordId,
                name: nextPayload.name || "",
                aadhaar: nextPayload.aadhaar || "",
                address: nextPayload.address || "",
                created_at: nextPayload.created_at || new Date().toISOString(),
            };
            const next = await upsertLocalListItem(LOCAL_KEYS.landlords, "landlord_id", landlord);
            await updateLocalLandlordsList(next);
            return { ok: true, landlord };
        };
        const data = await runWriteAction({
            url,
            action: "saveLandlord",
            payload: nextPayload,
            queuedMessage: "Landlord saved locally. Sync pending.",
            fallback: { ok: true, queued: true, landlord: nextPayload },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("saveLandlordConfig error", e);
        showToast("Failed to save landlord", "error");
        return { ok: false };
    }
}

/**
 * Deletes a landlord configuration from Sheets storage.
 * @param {string} landlordId - Identifier of the landlord to remove.
 * @returns {Promise<object>} API response containing ok/message metadata.
 */
export async function deleteLandlordConfig(landlordId) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });

    try {
        const localUpdate = async () => {
            const next = await removeLocalListItem(LOCAL_KEYS.landlords, "landlord_id", landlordId);
            await updateLocalLandlordsList(next);
            return { ok: true, landlordId };
        };
        const data = await runWriteAction({
            url,
            action: "deleteLandlord",
            payload: { landlordId },
            queuedMessage: "Landlord removal queued. Sync pending.",
            fallback: { ok: true, queued: true, landlordId },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("deleteLandlordConfig error", e);
        showToast("Failed to delete landlord", "error");
        return { ok: false };
    }
}

/**
 * Loads the list of previously generated bills for quick access in the UI.
 * @returns {Promise<{bills: Array}>} Generated bill summaries or empty list on error.
 */
export async function fetchBillsMinimal(status = "pending", options = {}) {
    const normalized = (status || "pending").toString().trim().toLowerCase();
    const monthsBack =
        options && typeof options === "object" ? Number(options.monthsBack) || 0 : 0;

    const base = await loadGeneratedBillsBase();
    if (base && Array.isArray(base.bills)) {
        const filtered = filterBillsByOptions(base.bills, {
            status: normalized,
            monthsBack,
        });
        return { bills: filtered };
    }

    const url = ensureAppScriptUrl();
    if (!url) {
        showToast("Configure Apps Script URL to view generated bills", "warning");
        return { bills: [] };
    }

    const params = {};
    if (normalized) params.status = normalized;
    if (monthsBack > 0) params.monthsBack = monthsBack;
    const finalParams = Object.keys(params).length ? params : undefined;

    try {
        return await callAppScript({
            url,
            action: "billsminimal",
            params: finalParams,
            cache: { useLocal: false, revalidate: false },
        });
    } catch (e) {
        console.error("fetchBillsMinimal error", e);
        showToast("Could not fetch bills", "error");
        return { bills: [] };
    }
}

export async function fetchGeneratedBills(options = {}) {
    const status =
        typeof options === "string"
            ? options
            : (options && typeof options === "object" ? options.status : "");
    const fromMonth =
        options && typeof options === "object" ? options.fromMonth : "";
    const toMonth =
        options && typeof options === "object" ? options.toMonth : "";
    const params = {};
    if (status) params.status = status;
    if (fromMonth) params.fromMonth = fromMonth;
    if (toMonth) params.toMonth = toMonth;
    const finalParams = Object.keys(params).length ? params : undefined;

    const base = await loadGeneratedBillsBase();
    if (base && Array.isArray(base.bills)) {
        const filteredBills = filterBillsByOptions(base.bills, { status, fromMonth, toMonth });
        const coverageSource =
            Array.isArray(base.coverage) && base.coverage.length ? base.coverage : base.bills;
        const filteredCoverage = filterCoverageByRange(coverageSource, { fromMonth, toMonth });
        return {
            ...base,
            bills: filteredBills,
            coverage: filteredCoverage,
        };
    }

    const url = ensureAppScriptUrl();
    if (!url) {
        showToast("Configure Apps Script URL to view generated bills", "warning");
        return { bills: [] };
    }

    try {
        const data = await callAppScript({
            url,
            action: "generatedbills",
            params: finalParams,
            cache: { useLocal: false, revalidate: false },
        });
        if (data && Array.isArray(data.bills)) {
            await setLocalData(LOCAL_KEYS.generatedBills, data);
        }
        return data;
    } catch (e) {
        console.error("fetchGeneratedBills error", e);
        showToast("Could not fetch generated bills", "error");
        return { bills: [] };
    }
}

/**
 * Fetches full bill details for a single bill line.
 * @param {string} billLineId - Bill line identifier.
 * @returns {Promise<{ok: boolean, bill?: object}>} Full bill details payload.
 */
export async function fetchBillDetails(billLineId) {
    const cleaned = (billLineId || "").toString().trim();
    if (!cleaned) return { ok: false };

    const base = await loadGeneratedBillsBase();
    if (base && Array.isArray(base.bills)) {
        const match = base.bills.find(
            (bill) => (bill.billLineId || bill.bill_line_id || "").toString().trim() === cleaned
        );
        if (match) {
            return { ok: true, bill: match };
        }
    }

    const url = ensureAppScriptUrl();
    if (!url) {
        showToast("Configure Apps Script URL to view bill details", "warning");
        return { ok: false };
    }

    try {
        return await callAppScript({
            url,
            action: "billdetails",
            params: { billLineId: cleaned },
            cache: { useLocal: false, revalidate: false },
        });
    } catch (e) {
        console.error("fetchBillDetails error", e);
        showToast("Could not fetch bill details", "error");
        return { ok: false };
    }
}

/**
 * Retrieves a saved billing record for a given month/wing combination.
 * @param {string} monthKey - Month key in YYYY-MM format.
 * @param {string} wing - Wing identifier to scope the record lookup.
 * @returns {Promise<object>} Billing payload (charges, meta, tenants) or empty object.
 */
export async function fetchBillingRecord(monthKey, wing) {
    if (!monthKey || !wing) return {};

    const localRecord = await buildBillingRecordFromLocalData(monthKey, wing);
    if (localRecord && (localRecord.hasConfig || localRecord.hasReadings)) {
        return localRecord;
    }

    const base = await loadGeneratedBillsBase();
    if (base && Array.isArray(base.bills)) {
        const record = buildBillingRecordFromBills(base.bills, monthKey, wing);
        if (record) {
            return record;
        }
    }

    if (localRecord) {
        return localRecord;
    }

    const url = ensureAppScriptUrl();
    if (!url) {
        showToast("Configure Apps Script URL to view saved billing", "warning");
        return {};
    }

    try {
        return await callAppScript({
            url,
            action: "getbillingrecord",
            params: { month: monthKey, wing },
            cache: { useLocal: false, revalidate: false },
        });
    } catch (e) {
        console.error("fetchBillingRecord error", e);
        showToast("Could not load saved billing", "error");
        return {};
    }
}

/**
 * Fetches all configured units from Google Sheets for selector population.
 * @returns {Promise<{units: Array}>} Unit list payload (empty on failure).
 */
export async function fetchUnitsFromSheet(force = false) {
    const entry = await getLocalEntry(LOCAL_KEYS.units);
    const cached = Array.isArray(entry?.value) ? entry.value : null;
    if (cached) {
        if (!shouldRevalidate(entry, force)) {
            return { units: cached };
        }
    }

    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view units", "warning"),
    });
    if (!url) return { units: cached || [] };

    const runFetch = async () => {
        const data = await callAppScript({
            url,
            action: "units",
            cache: { useLocal: false, revalidate: false },
        });
        if (Array.isArray(data?.units)) {
            await setLocalData(LOCAL_KEYS.units, data.units);
            return { units: data.units };
        }
        return { units: cached || [] };
    };

    if (cached) {
        if (navigator.onLine) {
            runFetch().catch((e) => {
                console.error("fetchUnitsFromSheet error", e);
            });
        }
        return { units: cached };
    }

    try {
        return await runFetch();
    } catch (e) {
        console.error("fetchUnitsFromSheet error", e);
        showToast("Could not fetch units", "error");
        return { units: cached || [] };
    }
}

/**
 * Persists billing calculations to Google Sheets.
 * @param {object} payload - Billing metadata, tenant charges, and notes.
 * @returns {Promise<object>} Apps Script response reflecting save status.
 */
export async function saveBillingRecord(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    const nextPayload = payload && typeof payload === "object" ? { ...payload } : {};

    try {
        const localUpdate = async () => {
            const generated = await buildLocalBillsFromPayload(nextPayload);
            const stored = await updateLocalGeneratedBills(generated);
            const monthKey = normalizeMonthKey(nextPayload.monthKey || "");
            const wing = nextPayload.wing || "";
            const meta = nextPayload.meta || {};
            if (monthKey && wing) {
                const wingConfigs = Array.isArray(await getLocalData(LOCAL_KEYS.wingMonthlyConfig, []))
                    ? await getLocalData(LOCAL_KEYS.wingMonthlyConfig, [])
                    : [];
                const existingConfig = wingConfigs.find(
                    (cfg) =>
                        normalizeMonthKey(cfg.month_key) === monthKey &&
                        normalizeWing(cfg.wing) === normalizeWing(wing)
                );
                const wingConfig = {
                    wing_month_id: existingConfig?.wing_month_id || createLocalId("wingmonth"),
                    month_key: monthKey,
                    wing,
                    electricity_rate: meta.electricityRate ?? meta.electricity_rate ?? "",
                    sweeping_per_flat: meta.sweepingPerFlat ?? meta.sweeping_per_flat ?? "",
                    motor_prev: meta.motorPrev ?? meta.motor_prev ?? "",
                    motor_new: meta.motorNew ?? meta.motor_new ?? "",
                    motor_units:
                        (Number(meta.motorNew ?? meta.motor_new ?? 0) || 0) -
                        (Number(meta.motorPrev ?? meta.motor_prev ?? 0) || 0),
                    created_at: existingConfig?.created_at || new Date().toISOString(),
                };
                const nextConfigs = wingConfigs.filter(
                    (cfg) =>
                        !(
                            normalizeMonthKey(cfg.month_key) === monthKey &&
                            normalizeWing(cfg.wing) === normalizeWing(wing)
                        )
                );
                nextConfigs.push(wingConfig);
                await setLocalData(LOCAL_KEYS.wingMonthlyConfig, nextConfigs);
            }

            const tenantsPayload = Array.isArray(nextPayload.tenants) ? nextPayload.tenants : [];
            if (monthKey && tenantsPayload.length) {
                const readings = Array.isArray(await getLocalData(LOCAL_KEYS.tenantMonthlyReadings, []))
                    ? await getLocalData(LOCAL_KEYS.tenantMonthlyReadings, [])
                    : [];
                const readingKeys = new Set();
                const nextReadings = tenantsPayload.map((tenant) => {
                    const tenancyId = tenant.tenancyId || tenant.tenancy_id || "";
                    const key = `${monthKey}__${tenancyId}`;
                    readingKeys.add(key);
                    const existing = readings.find(
                        (r) => `${normalizeMonthKey(r.month_key)}__${r.tenancy_id}` === key
                    );
                    return {
                        reading_id: existing?.reading_id || createLocalId("reading"),
                        month_key: monthKey,
                        tenancy_id: tenancyId,
                        prev_reading: tenant.prevReading ?? tenant.prev_reading ?? "",
                        new_reading: tenant.newReading ?? tenant.new_reading ?? "",
                        included: tenant.included,
                        override_rent: tenant.override_rent || tenant.rentAmount || "",
                        notes: tenant.notes || "",
                        created_at: existing?.created_at || new Date().toISOString(),
                    };
                });
                const retainedReadings = readings.filter((r) => {
                    const key = `${normalizeMonthKey(r.month_key)}__${r.tenancy_id}`;
                    return !readingKeys.has(key);
                });
                await setLocalData(LOCAL_KEYS.tenantMonthlyReadings, retainedReadings.concat(nextReadings));
            }

            if (monthKey && Array.isArray(stored.bills)) {
                const billLines = Array.isArray(await getLocalData(LOCAL_KEYS.billLines, []))
                    ? await getLocalData(LOCAL_KEYS.billLines, [])
                    : [];
                const tenantMap = new Map();
                tenantsPayload.forEach((tenant) => {
                    const tenancyId = tenant.tenancyId || tenant.tenancy_id || "";
                    if (!tenancyId) return;
                    tenantMap.set(tenancyId, tenant);
                });
                const billKeys = new Set();
                const nextBills = stored.bills
                    .filter((bill) => normalizeMonthKey(bill.monthKey) === monthKey)
                    .map((bill) => {
                        const tenancyId = bill.tenancyId || bill.tenancy_id || "";
                        const key = `${monthKey}__${tenancyId}`;
                        billKeys.add(key);
                        const existing = billLines.find(
                            (line) => `${normalizeMonthKey(line.month_key)}__${line.tenancy_id}` === key
                        );
                        const tenant = tenantMap.get(tenancyId) || {};
                        const prevReading = Number(tenant.prevReading ?? tenant.prev_reading ?? 0) || 0;
                        const newReading = Number(tenant.newReading ?? tenant.new_reading ?? 0) || 0;
                        const units = Math.max(newReading - prevReading, 0);
                        const totalAmount = Number(bill.totalAmount ?? bill.total_amount) || 0;
                        const amountPaid = Number(bill.amountPaid ?? bill.amount_paid) || 0;
                        const isPaid =
                            bill.isPaid ??
                            bill.is_paid ??
                            (totalAmount <= 0 || amountPaid + 0.005 >= totalAmount);
                        return {
                            bill_line_id: bill.billLineId || bill.bill_line_id || createLocalId("bill"),
                            month_key: monthKey,
                            tenancy_id: tenancyId,
                            rent_amount: Number(bill.rentAmount ?? bill.rent_amount) || 0,
                            electricity_units: units,
                            electricity_amount: Number(bill.electricityAmount ?? bill.electricity_amount) || 0,
                            motor_share_amount: Number(bill.motorShare ?? bill.motor_share_amount) || 0,
                            sweep_amount: Number(bill.sweepAmount ?? bill.sweep_amount) || 0,
                            total_amount: totalAmount,
                            payable_date: bill.payableDate || bill.payable_date || "",
                            generated_at: existing?.generated_at || new Date().toISOString(),
                            amount_paid: amountPaid,
                            is_paid: isPaid,
                        };
                    });
                const retainedBills = billLines.filter((line) => {
                    const key = `${normalizeMonthKey(line.month_key)}__${line.tenancy_id}`;
                    return !billKeys.has(key);
                });
                await setLocalData(LOCAL_KEYS.billLines, retainedBills.concat(nextBills));
            }
            return { ok: true, bills: stored.bills || [], coverage: stored.coverage || [] };
        };
        const data = await runWriteAction({
            url,
            action: "saveBillingRecord",
            payload: nextPayload,
            queuedMessage: "Billing saved locally. Sync pending.",
            fallback: { ok: true, queued: true },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("saveBillingRecord error", e);
        showToast("Failed to save billing to Google Sheets", "error");
        return {};
    }
}

/**
 * Saves or updates an individual unit configuration.
 * @param {object} payload - Unit fields including wing, number, direction, and floor.
 * @returns {Promise<object>} Apps Script response describing persistence outcome.
 */
export async function saveUnitConfig(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    const nextPayload = payload && typeof payload === "object" ? { ...payload } : {};
    if (!nextPayload.unitId) {
        nextPayload.unitId = createLocalId("unit");
    }

    try {
        const localUpdate = async () => {
            const unit = {
                unit_id: nextPayload.unitId,
                wing: nextPayload.wing || "",
                unit_number: nextPayload.unitNumber || nextPayload.unit_number || "",
                floor: nextPayload.floor || "",
                direction: nextPayload.direction || "",
                meter_number: nextPayload.meterNumber || nextPayload.meter_number || "",
                landlord_id: nextPayload.landlordId || nextPayload.landlord_id || "",
                notes: nextPayload.notes || "",
                is_occupied: !!nextPayload.isOccupied,
                current_tenancy_id: nextPayload.currentTenancyId || nextPayload.current_tenancy_id || "",
                created_at: nextPayload.created_at || new Date().toISOString(),
            };
            const next = await upsertLocalListItem(LOCAL_KEYS.units, "unit_id", unit);
            await updateLocalUnitsList(next);
            if (unit.wing) {
                const existingWings = await getLocalData(LOCAL_KEYS.wings, []);
                const normalized = normalizeWingValue(unit.wing);
                const hasWing = Array.isArray(existingWings)
                    ? existingWings.some((w) => normalizeWingValue(w) === normalized)
                    : false;
                if (!hasWing) {
                    await updateLocalWingsList([...(existingWings || []), unit.wing]);
                }
            }
            return { ok: true, unit };
        };
        const data = await runWriteAction({
            url,
            action: "saveUnit",
            payload: nextPayload,
            queuedMessage: "Unit saved locally. Sync pending.",
            fallback: { ok: true, queued: true, unit: nextPayload },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("saveUnitConfig error", e);
        showToast("Failed to save unit", "error");
        return {};
    }
}

/**
 * Deletes an existing unit record by id.
 * @param {string} unitId - Unique identifier for the unit.
 * @returns {Promise<object>} Apps Script response noting deletion status.
 */
export async function deleteUnitConfig(unitId) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });

    try {
        const localUpdate = async () => {
            const next = await removeLocalListItem(LOCAL_KEYS.units, "unit_id", unitId);
            await updateLocalUnitsList(next);
            return { ok: true, unitId };
        };
        const data = await runWriteAction({
            url,
            action: "deleteUnit",
            payload: { unitId },
            queuedMessage: "Unit removal queued. Sync pending.",
            fallback: { ok: true, queued: true, unitId },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("deleteUnitConfig error", e);
        showToast("Failed to delete unit", "error");
        return {};
    }
}

/**
 * Loads payment records from Google Sheets for display in the Payments tab.
 * @returns {Promise<{payments: Array}>} Collection of payment rows or empty array.
 */
export async function fetchPayments() {
    const entry = await getLocalEntry(LOCAL_KEYS.payments);
    const cached = Array.isArray(entry?.value) ? entry.value : null;
    if (cached) {
        if (!shouldRevalidate(entry, false)) {
            return { payments: cached };
        }
    }

    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view payments", "warning"),
    });
    if (!url) return { payments: cached || [] };

    const runFetch = async () => {
        const data = await callAppScript({
            url,
            action: "payments",
            cache: { useLocal: false, revalidate: false },
        });
        if (Array.isArray(data?.payments)) {
            await setLocalData(LOCAL_KEYS.payments, data.payments);
            return { payments: data.payments };
        }
        return { payments: cached || [] };
    };

    if (cached) {
        if (navigator.onLine) {
            runFetch().catch((e) => {
                console.error("fetchPayments error", e);
            });
        }
        return { payments: cached };
    }

    try {
        return await runFetch();
    } catch (e) {
        console.error("fetchPayments error", e);
        showToast("Could not fetch payments", "error");
        return { payments: cached || [] };
    }
}

/**
 * Requests a pre-signed URL or preview blob for an attachment stored remotely.
 * @param {string} attachmentUrl - URL returned from Sheets storage.
 * @returns {Promise<object>} Preview payload or empty object on failure.
 */
export async function fetchAttachmentPreview(attachmentUrl) {
    const cleaned = (attachmentUrl || "").toString().trim();
    if (cleaned.startsWith("data:")) {
        return { ok: true, previewUrl: cleaned, attachmentUrl: cleaned };
    }

    const url = ensureAppScriptUrl({
        onMissing: () => showToast("Configure Apps Script URL to view attachments", "warning"),
    });

    if (!url || !cleaned) return {};

    try {
        return await callAppScript({
            url,
            action: "attachmentpreview",
            params: { attachmentUrl: cleaned },
            cache: { useLocal: false, write: false },
        });
    } catch (e) {
        console.error("fetchAttachmentPreview error", e);
        showToast("Could not load attachment preview", "error");
        return {};
    }
}

/**
 * Saves a payment entry against a tenant/unit combination.
 * @param {object} payload - Payment details including amount, date, and attachment info.
 * @returns {Promise<object>} Result payload including ok/message flags.
 */
export async function savePaymentRecord(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });
    const nextPayload = payload && typeof payload === "object" ? { ...payload } : {};
    if (!nextPayload.id) {
        nextPayload.id = createLocalId("payment");
    }
    const localBillLineId = nextPayload.billLineId || "";
    if (typeof localBillLineId === "string" && localBillLineId.startsWith("bill-")) {
        nextPayload.billLineId = "";
    }

    try {
        const localUpdate = async () => {
            const payment = buildLocalPaymentRecord({ ...nextPayload, billLineId: localBillLineId });
            const existingPayments = await getLocalData(LOCAL_KEYS.payments, []);
            const nextPayments = Array.isArray(existingPayments)
                ? existingPayments.filter((p) => String(p.id) !== String(payment.id)).concat(payment)
                : [payment];
            await updateLocalPaymentsList(nextPayments);

            if (payment.attachmentId || payment.attachment_id) {
                const attachmentId = payment.attachmentId || payment.attachment_id;
                const attachments = Array.isArray(await getLocalData(LOCAL_KEYS.attachments, []))
                    ? await getLocalData(LOCAL_KEYS.attachments, [])
                    : [];
                const record = {
                    attachment_id: attachmentId,
                    file_name: payment.attachmentName || "",
                    file_url: payment.attachmentUrl || "",
                    file_drive_id: "",
                    uploaded_at: new Date().toISOString(),
                };
                const nextAttachments = attachments
                    .filter((a) => a.attachment_id !== attachmentId)
                    .concat(record);
                await setLocalData(LOCAL_KEYS.attachments, nextAttachments);
            }

            if (payment.billLineId) {
                const generated = await getLocalData(LOCAL_KEYS.generatedBills, { bills: [], coverage: [] });
                const bills = Array.isArray(generated?.bills) ? generated.bills : [];
                const updatedBills = updateBillsWithPayment(bills, nextPayments, payment.billLineId);
                await updateLocalGeneratedBills({ ...generated, bills: updatedBills });

                const billLines = Array.isArray(await getLocalData(LOCAL_KEYS.billLines, []))
                    ? await getLocalData(LOCAL_KEYS.billLines, [])
                    : [];
                const paidTotal = nextPayments
                    .filter((p) => String(p.billLineId || p.bill_line_id || "") === String(payment.billLineId))
                    .reduce((sum, p) => sum + (Number(p.amount) || 0), 0);
                const updatedBillLines = billLines.map((line) => {
                    const id = line.bill_line_id || line.billLineId || "";
                    if (String(id) !== String(payment.billLineId)) return line;
                    const total = Number(line.total_amount) || 0;
                    const isPaid = total <= 0 || paidTotal + 0.005 >= total;
                    return {
                        ...line,
                        amount_paid: paidTotal,
                        is_paid: isPaid,
                    };
                });
                await setLocalData(LOCAL_KEYS.billLines, updatedBillLines);
            }

            return { ok: true, payment };
        };
        const data = await runWriteAction({
            url,
            action: "savePayment",
            payload: nextPayload,
            queuedMessage: "Payment saved locally. Sync pending.",
            fallback: { ok: true, queued: true, payment: nextPayload },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("savePaymentRecord error", e);
        showToast("Failed to save payment", "error");
        return {};
    }
}

/**
 * Loads clauses from Google Sheets and updates the UI
 * @param {boolean} showNotification - Whether to show a toast notification
 */
export async function loadClausesFromSheet(showNotification = false, force = false) {
    const { normalizeClauseSections, renderClausesUI, setClausesDirty } =
        await import("../features/agreements/clauses.js");

    const entry = await getLocalEntry(LOCAL_KEYS.clauses);
    const cached = entry?.value && typeof entry.value === "object" ? entry.value : null;
    if (cached) {
        clauseSections.tenant.items = Array.isArray(cached.tenant) ? cached.tenant : [];
        clauseSections.landlord.items = Array.isArray(cached.landlord) ? cached.landlord : [];
        clauseSections.penalties.items = Array.isArray(cached.penalties) ? cached.penalties : [];
        clauseSections.misc.items = Array.isArray(cached.misc) ? cached.misc : [];
        normalizeClauseSections();
        renderClausesUI();
        setClausesDirty(false);
        if (!shouldRevalidate(entry, force)) {
            return;
        }
    }

    const url = ensureAppScriptUrl({
        onMissing: () =>
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL"),
    });

    if (!url) {
        normalizeClauseSections();
        renderClausesUI();
        setClausesDirty(false);
        return;
    }

    const runFetch = async () => {
        const data = await callAppScript({
            url,
            action: "clauses",
            cache: { useLocal: false, revalidate: false },
        });

        clauseSections.tenant.items = Array.isArray(data.tenant) ? data.tenant : [];
        clauseSections.landlord.items = Array.isArray(data.landlord) ? data.landlord : [];
        clauseSections.penalties.items = Array.isArray(data.penalties) ? data.penalties : [];
        clauseSections.misc.items = Array.isArray(data.misc) ? data.misc : [];

        normalizeClauseSections();
        renderClausesUI();
        setClausesDirty(false);
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
        await setLocalData(LOCAL_KEYS.clauses, {
            tenant: clauseSections.tenant.items,
            landlord: clauseSections.landlord.items,
            penalties: clauseSections.penalties.items,
            misc: clauseSections.misc.items,
        });

        if (showNotification) {
            showToast("Latest clauses loaded from Google Sheets", "info");
        }
    };

    if (cached && navigator.onLine) {
        runFetch().catch((e) => {
            console.warn("Could not fetch clauses; leaving current UI", e);
            updateConnectionIndicator(
                navigator.onLine ? "online" : "offline",
                navigator.onLine ? "Clause sync failed" : "Offline"
            );
            if (showNotification) {
                showToast("Failed to load clauses from Google Sheets", "error");
            }
        });
        return;
    }

    try {
        await runFetch();
    } catch (e) {
        console.warn("Could not fetch clauses; leaving current UI", e);
        normalizeClauseSections();
        renderClausesUI();
        setClausesDirty(false);
        updateConnectionIndicator(
            navigator.onLine ? "online" : "offline",
            navigator.onLine ? "Clause sync failed" : "Offline"
        );
        if (showNotification) {
            showToast("Failed to load clauses from Google Sheets", "error");
        }
    }
}

export async function uploadPaymentAttachment(payload, options = {}) {
    const localOnly = options.localOnly !== false;
    const localDataUrl = payload?.dataUrl || payload?.attachmentDataUrl || "";
    const localName =
        payload?.attachmentName ||
        payload?.fileName ||
        payload?.name ||
        "receipt";
    if (localOnly) {
        return {
            ok: true,
            attachment: {
                attachment_id: createLocalId("attachment"),
                attachmentUrl: localDataUrl,
                attachmentName: localName,
            },
        };
    }

    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => showToast("Configure Apps Script URL to upload receipts", "warning"),
    });
    if (!url) return { ok: false };
    const onProgress = options && typeof options.onProgress === "function" ? options.onProgress : null;
    const onAbort = options && typeof options.onAbort === "function" ? options.onAbort : null;
    const uploadId = options && options.uploadId ? String(options.uploadId) : `upload-${Date.now()}`;
    const tauriInvoke = window.__TAURI__?.core?.invoke;
    const tauriEvents = window.__TAURI__?.event;

    if (typeof tauriInvoke === "function") {
        let unlisten = null;
        if (onProgress && typeof tauriEvents?.listen === "function") {
            unlisten = await tauriEvents.listen("upload-progress", (event) => {
                const payloadData = event?.payload || {};
                const id = payloadData.uploadId || payloadData.upload_id || "";
                if (!id || id !== uploadId) return;
                const loaded = Number(payloadData.loaded) || 0;
                const total = Number(payloadData.total) || 0;
                const percent = total ? (loaded / total) * 100 : null;
                onProgress({ loaded, total, percent });
                if (payloadData.done && typeof unlisten === "function") {
                    unlisten();
                    unlisten = null;
                }
            });
        }
        if (onAbort) {
            onAbort(() => tauriInvoke("cancel_upload", { uploadId }));
        }
        try {
            const result = await tauriInvoke("upload_payment_attachment", {
                url,
                payload,
                uploadId,
            });
            if (typeof unlisten === "function") unlisten();
            return result || { ok: false };
        } catch (err) {
            if (typeof unlisten === "function") unlisten();
            console.error("uploadPaymentAttachment error", err);
            showToast("Failed to upload receipt", "error");
            return { ok: false, error: String(err) };
        }
    }

    if (!onProgress) {
        try {
            const data = await callAppScript({
                url,
                action: "uploadPaymentAttachment",
                method: "POST",
                payload,
            });
            return data || { ok: false };
        } catch (err) {
            console.error("uploadPaymentAttachment error", err);
            showToast("Failed to upload receipt", "error");
            return { ok: false, error: String(err) };
        }
    }

    try {
        const body = JSON.stringify({ action: "uploadPaymentAttachment", payload });
        const result = await new Promise((resolve) => {
            const xhr = new XMLHttpRequest();
            xhr.open("POST", url, true);
            xhr.setRequestHeader("Content-Type", "text/plain");

            xhr.upload.onprogress = (event) => {
                if (!onProgress) return;
                const total = event.lengthComputable ? event.total : null;
                const loaded = event.loaded || 0;
                const percent = total ? (loaded / total) * 100 : null;
                onProgress({ loaded, total, percent });
            };

            xhr.onload = () => {
                if (xhr.status < 200 || xhr.status >= 300) {
                    resolve({ ok: false, error: `Upload failed (${xhr.status})` });
                    return;
                }
                try {
                    resolve(JSON.parse(xhr.responseText));
                } catch (err) {
                    resolve({ ok: false, error: "Upload response parse failed" });
                }
            };

            xhr.onerror = () => resolve({ ok: false, error: "Upload failed" });
            xhr.onabort = () => resolve({ ok: false, error: "Upload cancelled" });

            if (onAbort) {
                onAbort(() => xhr.abort());
            }
            xhr.send(body);
        });

        if (!result?.ok) {
            showToast("Failed to upload receipt", "error");
        }
        return result || { ok: false };
    } catch (err) {
        console.error("uploadPaymentAttachment error", err);
        showToast("Failed to upload receipt", "error");
        return { ok: false, error: String(err) };
    }
}

export async function deleteAttachment(attachmentId) {
    const url = ensureAppScriptUrl({
        promptForConfig: false,
    });
    if (!attachmentId) return { ok: false };
    try {
        const localUpdate = async () => {
            const payments = Array.isArray(await getLocalData(LOCAL_KEYS.payments, []))
                ? await getLocalData(LOCAL_KEYS.payments, [])
                : [];
            const next = payments.map((p) => {
                if ((p.attachmentId || p.attachment_id) !== attachmentId) return p;
                return {
                    ...p,
                    attachmentId: "",
                    attachment_id: "",
                    attachmentName: "",
                    attachmentUrl: "",
                };
            });
            await updateLocalPaymentsList(next);
            const attachments = Array.isArray(await getLocalData(LOCAL_KEYS.attachments, []))
                ? await getLocalData(LOCAL_KEYS.attachments, [])
                : [];
            const nextAttachments = attachments.filter((a) => a.attachment_id !== attachmentId);
            await setLocalData(LOCAL_KEYS.attachments, nextAttachments);
            return { ok: true, attachmentId };
        };
        const data = await runWriteAction({
            url,
            action: "deleteAttachment",
            payload: { attachmentId },
            queuedMessage: "Attachment delete queued. Sync pending.",
            fallback: { ok: true, queued: true, attachmentId },
            localUpdate,
        });
        return data || { ok: false };
    } catch (err) {
        console.error("deleteAttachment error", err);
        return { ok: false };
    }
}

/**
 * Saves clauses to Google Sheets
 */
export async function saveClausesToSheet() {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => alert("Please configure the Apps Script URL first."),
    });

    // Import the clauses module functions
    const { normalizeClauseSections, setClausesDirty } =
        await import("../features/agreements/clauses.js");

    normalizeClauseSections();

    const payload = {
        tenant: clauseSections.tenant.items,
        landlord: clauseSections.landlord.items,
        penalties: clauseSections.penalties.items,
        misc: clauseSections.misc.items,
    };

    try {
        const data = await runWriteAction({
            url,
            action: "saveClauses",
            payload,
            queuedMessage: "Clauses saved locally. Sync pending.",
            fallback: { ok: true, queued: true, message: "Clauses queued" },
            localUpdate: async () => {
                await setLocalData(LOCAL_KEYS.clauses, payload);
                return { ok: true, message: "Clauses saved locally" };
            },
        });
        if (data?.ok && data?.message) {
            showToast(data.message, "success");
        }
        setClausesDirty(false);
    } catch (e) {
        console.error("saveClausesToSheet error", e);
        showToast("Failed to save clauses to Google Sheets", "error");
    }
}

/**
 * Fetches tenant directory (tenants + family) from Google Sheets
 * @returns {Promise<{ tenants: Array }>} API response with tenants array
 */
export async function fetchTenantDirectory() {
    const entry = await getLocalEntry(LOCAL_KEYS.tenants);
    const cached = Array.isArray(entry?.value) ? entry.value : null;
    if (cached) {
        if (!shouldRevalidate(entry, false)) {
            return { tenants: cached };
        }
    }

    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            showToast("Please configure the Apps Script URL to view tenants", "warning");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return { tenants: cached || [] };

    const runFetch = async () => {
        const data = await callAppScript({
            url,
            action: "tenants",
            cache: { useLocal: false, revalidate: false },
        });
        if (!data || !Array.isArray(data.tenants)) {
            throw new Error("Invalid tenant response");
        }
        await setLocalData(LOCAL_KEYS.tenants, data.tenants);
        updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Internet online");
        return { tenants: data.tenants };
    };

    if (cached) {
        if (navigator.onLine) {
            runFetch().catch((e) => {
                console.error("fetchTenantDirectory error", e);
            });
        }
        return { tenants: cached };
    }

    try {
        return await runFetch();
    } catch (e) {
        console.error("fetchTenantDirectory error", e);
        showToast("Failed to load tenants from Google Sheets", "error");
        updateConnectionIndicator(
            navigator.onLine ? "online" : "offline",
            navigator.onLine ? "Tenant fetch failed" : "Offline"
        );
        return { tenants: cached || [] };
    }
}

/**
 * Updates an existing tenant + family entries in Google Sheets
 * @param {object} payload - update payload (tenantId, tenancyId, grn, updates, familyMembers)
 */
export async function updateTenantRecord(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    const nextPayload = payload && typeof payload === "object" ? { ...payload } : {};
    if (!nextPayload.tenantId) {
        nextPayload.tenantId = createLocalId("tenant");
    }
    if (!nextPayload.tenancyId && !nextPayload.forceNewTenancyId) {
        nextPayload.tenancyId = createLocalId("tenancy");
    }

    try {
        const localUpdate = async () => {
            const tenants = Array.isArray(await getLocalData(LOCAL_KEYS.tenants, []))
                ? (await getLocalData(LOCAL_KEYS.tenants, []))
                : [];
            const units = Array.isArray(await getLocalData(LOCAL_KEYS.units, []))
                ? (await getLocalData(LOCAL_KEYS.units, []))
                : [];
            const landlords = Array.isArray(await getLocalData(LOCAL_KEYS.landlords, []))
                ? (await getLocalData(LOCAL_KEYS.landlords, []))
                : [];

            const tenantId = nextPayload.tenantId;
            const tenancyId = nextPayload.forceNewTenancyId || nextPayload.tenancyId;
            const unitId = nextPayload.unitId || nextPayload.updates?.unitId || nextPayload.templateData?.unit_id || "";
            let landlordId = nextPayload.landlordId || nextPayload.updates?.landlordId || nextPayload.templateData?.landlord_id || "";
            nextPayload.tenancyId = tenancyId;
            nextPayload.unitId = unitId;

            let unit = units.find((u) => u.unit_id === unitId);
            if (!unit && unitId) {
                unit = {
                    unit_id: unitId,
                    wing: nextPayload.wing || nextPayload.templateData?.wing || "",
                    unit_number: nextPayload.templateData?.unit_number || "",
                    floor: nextPayload.templateData?.floor_of_building || "",
                    direction: nextPayload.templateData?.direction_build || "",
                    meter_number: nextPayload.templateData?.meter_number || "",
                    landlord_id: landlordId || "",
                    notes: "",
                    is_occupied: true,
                    current_tenancy_id: tenancyId,
                    created_at: new Date().toISOString(),
                };
                units.push(unit);
            }
            if (unit) {
                unit.is_occupied = true;
                unit.current_tenancy_id = tenancyId;
            }

            let landlord = landlords.find((l) => l.landlord_id === landlordId);
            const templateLandlordName = nextPayload.templateData?.Landlord_name;
            if (!landlord && templateLandlordName) {
                if (!landlordId) {
                    landlordId = createLocalId("landlord");
                }
                landlord = {
                    landlord_id: landlordId,
                    name: templateLandlordName || "",
                    aadhaar: nextPayload.templateData?.landlord_aadhar || "",
                    address: nextPayload.templateData?.landlord_address || "",
                    created_at: new Date().toISOString(),
                };
                landlords.push(landlord);
            }
            nextPayload.landlordId = landlordId;

            const list = tenants.slice();
            let idx = list.findIndex((t) => t.tenancyId === tenancyId);
            if (idx < 0) {
                idx = list.findIndex((t) => t.tenantId === tenantId);
            }
            const existing = idx >= 0 ? list[idx] : {};

            const updated = buildTenantEntryFromPayload(nextPayload, {
                existing,
                unit: unit || {},
                landlord: landlord || {},
                tenantId,
                tenancyId,
                unitId,
                landlordId,
            });

            if (idx >= 0) {
                list[idx] = updated;
            } else {
                list.push(updated);
            }

            if (nextPayload.createNewTenancy && nextPayload.previousTenancyId && !nextPayload.keepPreviousActive) {
                const prevIdx = list.findIndex((t) => t.tenancyId === nextPayload.previousTenancyId);
                if (prevIdx >= 0) {
                    const prev = list[prevIdx];
                    list[prevIdx] = {
                        ...prev,
                        activeTenant: false,
                        tenancyEndDate: prev.tenancyEndDate || new Date().toISOString().slice(0, 10),
                    };
                }
            }

            await updateLocalUnitsList(units);
            await updateLocalLandlordsList(landlords);
            await updateLocalTenantsList(list);
            const tenancies = Array.isArray(await getLocalData(LOCAL_KEYS.tenancies, []))
                ? await getLocalData(LOCAL_KEYS.tenancies, [])
                : [];
            let nextTenancies = tenancies
                .filter((t) => t.tenancy_id !== tenancyId)
                .concat(
                    buildTenancyRecordFromEntry(updated, {
                        tenantId,
                        tenancyId,
                        unitId,
                        landlordId,
                    })
                );
            if (nextPayload.createNewTenancy && nextPayload.previousTenancyId && !nextPayload.keepPreviousActive) {
                nextTenancies = nextTenancies.map((t) => {
                    if (t.tenancy_id !== nextPayload.previousTenancyId) return t;
                    return {
                        ...t,
                        status: "ENDED",
                        end_date: t.end_date || new Date().toISOString().slice(0, 10),
                    };
                });
            }
            await setLocalData(LOCAL_KEYS.tenancies, nextTenancies);

            const familyMembers = Array.isArray(await getLocalData(LOCAL_KEYS.familyMembers, []))
                ? await getLocalData(LOCAL_KEYS.familyMembers, [])
                : [];
            const updatedFamily = buildFamilyRecordsFromEntry(updated, tenantId);
            const nextFamily = familyMembers
                .filter((m) => m.tenant_id !== tenantId)
                .concat(updatedFamily);
            await setLocalData(LOCAL_KEYS.familyMembers, nextFamily);
            return { ok: true, tenantId, tenancyId };
        };
        const data = await runWriteAction({
            url,
            action: "updateTenant",
            payload: nextPayload,
            queuedMessage: "Tenant update saved locally. Sync pending.",
            fallback: { ok: true, queued: true, message: "Tenant update queued" },
            localUpdate,
        });
        return data;
    } catch (e) {
        console.error("updateTenantRecord error", e);
        showToast("Failed to update tenant", "error");
        updateConnectionIndicator(
            navigator.onLine ? "online" : "offline",
            navigator.onLine ? "Update failed" : "Offline"
        );
        throw e;
    }
}

export async function getRentRevisions(tenancyId) {
    const key = LOCAL_KEYS.rentRevisions(tenancyId);
    const entry = await getLocalEntry(key);
    const cached = Array.isArray(entry?.value) ? entry.value : null;
    if (cached) {
        if (!shouldRevalidate(entry, false)) {
            return { ok: true, revisions: cached };
        }
    }

    const allEntry = await getLocalEntry(LOCAL_KEYS.rentRevisionsAll);
    const allRevisions = Array.isArray(allEntry?.value) ? allEntry.value : null;
    const filteredAll = allRevisions
        ? allRevisions.filter((rev) => (rev?.tenancy_id || rev?.tenancyId) === tenancyId)
        : null;
    if (filteredAll && !shouldRevalidate(allEntry, false)) {
        await setLocalData(key, filteredAll);
        return { ok: true, revisions: filteredAll };
    }

    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });
    if (!url) return { ok: !!(cached || filteredAll), revisions: cached || filteredAll || [] };

    const runFetch = async () => {
        const data = await callAppScript({
            url,
            action: "getRentRevisions",
            method: "POST",
            payload: { tenancyId },
            cache: { useLocal: false, revalidate: false },
        });
        if (Array.isArray(data?.revisions)) {
            await setLocalData(key, data.revisions);
        }
        return data;
    };

    const runFetchAll = async () => {
        const data = await callAppScript({
            url,
            action: "rentrevisions",
            cache: { useLocal: false, revalidate: false },
        });
        if (Array.isArray(data?.revisions)) {
            await setLocalData(LOCAL_KEYS.rentRevisionsAll, data.revisions);
            const next = data.revisions.filter(
                (rev) => (rev?.tenancy_id || rev?.tenancyId) === tenancyId
            );
            await setLocalData(key, next);
            return { ok: true, revisions: next };
        }
        return data;
    };

    if (cached) {
        if (navigator.onLine) {
            runFetch().catch((e) => console.error("getRentRevisions error", e));
        }
        return { ok: true, revisions: cached };
    }

    if (filteredAll) {
        if (navigator.onLine) {
            runFetchAll().catch((e) => console.error("getRentRevisions error", e));
        }
        await setLocalData(key, filteredAll);
        return { ok: true, revisions: filteredAll };
    }

    try {
        return await runFetch();
    } catch (e) {
        console.error("getRentRevisions error", e);
        showToast("Failed to load rent history", "error");
        return { ok: false, revisions: cached || filteredAll || [] };
    }
}

export async function saveRentRevision(payload) {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });

    const nextPayload = payload && typeof payload === "object" ? { ...payload } : {};
    if (!nextPayload.tenancyId && nextPayload.tenancy_id) {
        nextPayload.tenancyId = nextPayload.tenancy_id;
    }

    try {
        const localUpdate = async () => {
            const tenancyId = nextPayload.tenancyId || "";
            const key = LOCAL_KEYS.rentRevisions(tenancyId);
            const revisions = Array.isArray(await getLocalData(key, []))
                ? await getLocalData(key, [])
                : [];
            const effectiveMonth = nextPayload.effectiveMonth || nextPayload.effective_month || "";
            const record = {
                revision_id: nextPayload.revision_id || createLocalId("revision"),
                tenancy_id: tenancyId,
                effective_month: effectiveMonth,
                rent_amount: Number(nextPayload.rentAmount ?? nextPayload.rent_amount) || 0,
                note: nextPayload.note || "",
                created_at: nextPayload.created_at || new Date().toISOString(),
            };
            const next = revisions.filter(
                (r) => normalizeMonthKey(r.effective_month) !== normalizeMonthKey(record.effective_month)
            );
            next.push(record);
            await setLocalData(key, next);
            const allRevisions = Array.isArray(await getLocalData(LOCAL_KEYS.rentRevisionsAll, []))
                ? await getLocalData(LOCAL_KEYS.rentRevisionsAll, [])
                : [];
            const nextAll = allRevisions.filter(
                (r) =>
                    (r?.tenancy_id || r?.tenancyId) !== tenancyId ||
                    normalizeMonthKey(r.effective_month || r.effectiveMonth) !== normalizeMonthKey(record.effective_month)
            );
            nextAll.push(record);
            await setLocalData(LOCAL_KEYS.rentRevisionsAll, nextAll);

            const tenants = Array.isArray(await getLocalData(LOCAL_KEYS.tenants, []))
                ? await getLocalData(LOCAL_KEYS.tenants, [])
                : [];
            const rentAmount = getLatestRentFromRevisions(
                nextAll.filter((rev) => (rev?.tenancy_id || rev?.tenancyId) === tenancyId)
            );
            if (rentAmount !== null) {
                const idx = tenants.findIndex((t) => t.tenancyId === tenancyId);
                if (idx >= 0) {
                    tenants[idx] = {
                        ...tenants[idx],
                        rentAmount,
                        currentRent: rentAmount,
                    };
                    await updateLocalTenantsList(tenants);
                }
            }
            return { ok: true, revision: record, revisions: next };
        };
        const data = await runWriteAction({
            url,
            action: "saveRentRevision",
            payload: nextPayload,
            queuedMessage: "Rent revision saved locally. Sync pending.",
            fallback: { ok: true, queued: true },
            localUpdate,
        });
        if (data?.ok && data?.revision) showToast("Rent revision saved locally", "success");
        return data;
    } catch (e) {
        console.error("saveRentRevision error", e);
        showToast("Failed to save rent revision", "error");
        return { ok: false };
    }
}

/**
 * Saves tenant data to Google Sheets database
 */
export async function saveTenantToDb() {
    const url = ensureAppScriptUrl({
        promptForConfig: true,
        onMissing: () => {
            alert("Please configure the Apps Script URL first.");
            updateConnectionIndicator(navigator.onLine ? "online" : "offline", "Set Apps Script URL");
        },
    });

    // Import the form module to collect data
    const { collectFullPayloadForDb } = await import("../features/tenants/form.js");
    const payload = collectFullPayloadForDb();
    const action = "saveTenant";

    try {
        const nextPayload = payload && typeof payload === "object" ? { ...payload } : {};
        if (!nextPayload.tenantId) {
            nextPayload.tenantId = createLocalId("tenant");
        }
        if (!nextPayload.tenancyId) {
            nextPayload.tenancyId = createLocalId("tenancy");
        }
        if (!nextPayload.unitId) {
            nextPayload.unitId = createLocalId("unit");
        }
        if (!nextPayload.landlordId && nextPayload.templateData?.Landlord_name) {
            nextPayload.landlordId = createLocalId("landlord");
        }

        const localUpdate = async () => {
            const tenants = Array.isArray(await getLocalData(LOCAL_KEYS.tenants, []))
                ? await getLocalData(LOCAL_KEYS.tenants, [])
                : [];
            const units = Array.isArray(await getLocalData(LOCAL_KEYS.units, []))
                ? await getLocalData(LOCAL_KEYS.units, [])
                : [];
            const landlords = Array.isArray(await getLocalData(LOCAL_KEYS.landlords, []))
                ? await getLocalData(LOCAL_KEYS.landlords, [])
                : [];

            const tenantId = nextPayload.tenantId;
            const tenancyId = nextPayload.tenancyId;
            const unitId = nextPayload.unitId;
            let landlordId = nextPayload.landlordId || "";

            let unit = units.find((u) => u.unit_id === unitId);
            if (!unit) {
                unit = {
                    unit_id: unitId,
                    wing: nextPayload.wing || nextPayload.templateData?.wing || "",
                    unit_number: nextPayload.unit_number || nextPayload.templateData?.unit_number || "",
                    floor: nextPayload.floor_of_building || nextPayload.templateData?.floor_of_building || "",
                    direction: nextPayload.direction_build || nextPayload.templateData?.direction_build || "",
                    meter_number: nextPayload.meter_number || nextPayload.templateData?.meter_number || "",
                    landlord_id: landlordId || "",
                    notes: "",
                    is_occupied: true,
                    current_tenancy_id: tenancyId,
                    created_at: new Date().toISOString(),
                };
                units.push(unit);
            }
            unit.is_occupied = true;
            unit.current_tenancy_id = tenancyId;

            let landlord = landlords.find((l) => l.landlord_id === landlordId);
            if (!landlord && nextPayload.templateData?.Landlord_name) {
                if (!landlordId) {
                    landlordId = createLocalId("landlord");
                    nextPayload.landlordId = landlordId;
                }
                landlord = {
                    landlord_id: landlordId,
                    name: nextPayload.templateData?.Landlord_name || "",
                    aadhaar: nextPayload.templateData?.landlord_aadhar || "",
                    address: nextPayload.templateData?.landlord_address || "",
                    created_at: new Date().toISOString(),
                };
                landlords.push(landlord);
            }

            const entry = buildTenantEntryFromPayload(nextPayload, {
                existing: {},
                unit: unit || {},
                landlord: landlord || {},
                tenantId,
                tenancyId,
                unitId,
                landlordId,
            });

            const nextTenants = tenants.filter((t) => t.tenancyId !== tenancyId).concat(entry);
            await updateLocalUnitsList(units);
            await updateLocalLandlordsList(landlords);
            await updateLocalTenantsList(nextTenants);
            const tenancies = Array.isArray(await getLocalData(LOCAL_KEYS.tenancies, []))
                ? await getLocalData(LOCAL_KEYS.tenancies, [])
                : [];
            const tenancyRecord = buildTenancyRecordFromEntry(entry, {
                tenantId,
                tenancyId,
                unitId,
                landlordId,
            });
            const nextTenancies = tenancies
                .filter((t) => t.tenancy_id !== tenancyId)
                .concat(tenancyRecord);
            await setLocalData(LOCAL_KEYS.tenancies, nextTenancies);

            const familyMembers = Array.isArray(await getLocalData(LOCAL_KEYS.familyMembers, []))
                ? await getLocalData(LOCAL_KEYS.familyMembers, [])
                : [];
            const updatedFamily = buildFamilyRecordsFromEntry(entry, tenantId);
            const nextFamily = familyMembers
                .filter((m) => m.tenant_id !== tenantId)
                .concat(updatedFamily);
            await setLocalData(LOCAL_KEYS.familyMembers, nextFamily);
            if (entry.wing) {
                const existingWings = await getLocalData(LOCAL_KEYS.wings, []);
                const normalized = normalizeWingValue(entry.wing);
                const hasWing = Array.isArray(existingWings)
                    ? existingWings.some((w) => normalizeWingValue(w) === normalized)
                    : false;
                if (!hasWing) {
                    await updateLocalWingsList([...(existingWings || []), entry.wing]);
                }
            }
            return { ok: true, tenantId, tenancyId };
        };

        const data = await runWriteAction({
            url,
            action,
            payload: nextPayload,
            queuedMessage: "Tenant saved locally. Sync pending.",
            fallback: { ok: true, queued: true, message: "Tenant queued" },
            localUpdate,
        });

        if (data?.ok) {
            alert("Tenant saved locally. Sync pending.");
            const { refreshUnitOptions } = await import("../features/tenants/form.js");
            refreshUnitOptions(true);
            return;
        }
        alert("Failed to save tenant locally.");
    } catch (e) {
        console.error("saveTenantToDb error", e);
        alert("Failed to call Apps Script. Check URL / deployment.");
        updateConnectionIndicator(
            navigator.onLine ? "online" : "offline",
            navigator.onLine ? "Save failed" : "Offline"
        );
    }
}
