import { LOCAL_KEYS, getLocalEntry } from "../../api/localStore.js";
import { normalizeMonthKey } from "../../utils/formatters.js";

const billingStatusState = {
    initialized: false,
    loading: false,
    error: "",
    currentMonthKey: "",
    generatedCurrentMonth: 0,
    pendingAll: 0,
    totalBills: 0,
};

function getElements() {
    return {
        widget: document.getElementById("billingStatusWidget"),
        currentMonthLabel: document.getElementById("billingStatusCurrentMonthLabel"),
        currentMonthCount: document.getElementById("billingStatusCurrentMonthCount"),
        pendingCount: document.getElementById("billingStatusPendingCount"),
        meta: document.getElementById("billingStatusMeta"),
    };
}

function normalizeBooleanValue(value) {
    if (value === true || value === false) return value;
    if (typeof value === "string") {
        return value.toLowerCase() !== "false" && value !== "";
    }
    return !!value;
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

function getBillMonthKey(bill = {}) {
    return normalizeMonthKey(
        bill?.monthKey ||
            bill?.month_key ||
            bill?.month ||
            bill?.monthLabel ||
            bill?.month_label ||
            ""
    );
}

function formatMonthLabel(monthKey = "") {
    if (!/^\d{4}-\d{2}$/.test(monthKey)) return "current month";
    const [year, month] = monthKey.split("-").map((value) => Number(value));
    const date = new Date(year, month - 1, 1);
    if (Number.isNaN(date.getTime())) return monthKey;
    return date.toLocaleDateString("en-GB", { month: "short", year: "numeric" });
}

function getPreviousMonthKeyFromNow() {
    const now = new Date();
    const previousMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    return normalizeMonthKey(previousMonthDate);
}

async function readGeneratedBillsFromLocalDb() {
    const entry = await getLocalEntry(LOCAL_KEYS.generatedBills);
    const raw = entry?.value;
    if (Array.isArray(raw?.bills)) return raw.bills;
    if (Array.isArray(raw)) return raw;
    return [];
}

function computeCounts(bills = []) {
    const currentMonthKey = getPreviousMonthKeyFromNow();
    let generatedCurrentMonth = 0;
    let pendingAll = 0;

    bills.forEach((bill) => {
        const monthKey = getBillMonthKey(bill);
        if (monthKey === currentMonthKey) generatedCurrentMonth += 1;
        if (resolveBillPaidFlag(bill) !== true) pendingAll += 1;
    });

    return {
        currentMonthKey,
        generatedCurrentMonth,
        pendingAll,
        totalBills: bills.length,
    };
}

function render() {
    const { widget, currentMonthLabel, currentMonthCount, pendingCount, meta } = getElements();
    if (!widget) return;

    if (currentMonthLabel) {
        currentMonthLabel.textContent = `Generated bills (${formatMonthLabel(billingStatusState.currentMonthKey)})`;
    }
    if (currentMonthCount) {
        currentMonthCount.textContent = String(billingStatusState.generatedCurrentMonth);
    }
    if (pendingCount) {
        pendingCount.textContent = String(billingStatusState.pendingAll);
    }

    if (!meta) return;
    if (billingStatusState.loading) {
        meta.textContent = "Refreshing billing status...";
        return;
    }
    if (billingStatusState.error) {
        meta.textContent = billingStatusState.error;
        return;
    }
    meta.textContent = `Loaded ${billingStatusState.totalBills} generated bills from cache.`;
}

async function refreshBillingStatus() {
    if (billingStatusState.loading) return;
    billingStatusState.loading = true;
    billingStatusState.error = "";
    render();

    try {
        const bills = await readGeneratedBillsFromLocalDb();
        const counts = computeCounts(Array.isArray(bills) ? bills : []);
        billingStatusState.currentMonthKey = counts.currentMonthKey;
        billingStatusState.generatedCurrentMonth = counts.generatedCurrentMonth;
        billingStatusState.pendingAll = counts.pendingAll;
        billingStatusState.totalBills = counts.totalBills;
    } catch (err) {
        console.error("Failed to load billing status widget counts", err);
        billingStatusState.currentMonthKey = getPreviousMonthKeyFromNow();
        billingStatusState.generatedCurrentMonth = 0;
        billingStatusState.pendingAll = 0;
        billingStatusState.totalBills = 0;
        billingStatusState.error = "Failed to read billing status.";
    } finally {
        billingStatusState.loading = false;
        render();
    }
}

function bindRefreshEvents() {
    document.addEventListener("sync:completed", refreshBillingStatus);
    document.addEventListener("payment:saved", refreshBillingStatus);
    document.addEventListener("paid-bills:updated", refreshBillingStatus);
    document.addEventListener("flow:changed", (event) => {
        const mode = event?.detail?.mode;
        if (mode === "dashboard") {
            refreshBillingStatus();
        }
    });
}

export function initBillingStatus() {
    if (billingStatusState.initialized) return;
    billingStatusState.initialized = true;

    const { widget } = getElements();
    if (!widget) return;

    bindRefreshEvents();
    refreshBillingStatus();
}
