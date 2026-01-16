import { formatCurrency } from "../../utils/formatters.js";
import { escapeHtml } from "../../utils/htmlUtils.js";
import { normalizeMonthKey } from "../../utils/normalizers.js";
import { fetchPaidBillsSummary, openPaidBillDetails } from "../billing/payments.js";

const MAX_RECENT_PAYMENTS = 5;

const recentPaymentsState = {
    bills: [],
    widgets: [],
    loading: false,
    loadPromise: null,
};

let recentPaymentsInitialized = false;

function getBillLineId(bill) {
    return (bill?.billLineId || bill?.bill_line_id || "").toString().trim();
}

function buildUnitLabel(bill) {
    const wing = (bill?.wing || "").toString().trim();
    const unitNumber = (bill?.unitNumber || bill?.unit_number || "").toString().trim();
    if (wing && unitNumber) return `${wing} - ${unitNumber}`;
    return wing || unitNumber || "-";
}

function formatMonthLabel(monthKey, monthLabel) {
    const label = (monthLabel || "").toString().trim();
    const normalized = (monthKey || "").toString().trim();
    const isIsoLabel = label.includes("T") || label.includes("Z");
    const candidate = label && !isIsoLabel ? label : normalized;
    const monthKeyValue = normalizeMonthKey(candidate);
    if (/^\d{4}-\d{2}$/.test(monthKeyValue)) {
        const [year, month] = monthKeyValue.split("-");
        const date = new Date(Number(year), Number(month) - 1, 1);
        if (!Number.isNaN(date.getTime())) {
            return date.toLocaleDateString("en-GB", { month: "short", year: "numeric" });
        }
    }
    return candidate || "-";
}

function formatAmountLabel(value) {
    return formatCurrency(value, { emptyValue: "-", invalidValue: "-" });
}

function renderWidget(widget, bills) {
    if (!widget?.list) return;
    const { list, empty, variant } = widget;
    list.innerHTML = "";

    if (!bills.length) {
        if (empty) {
            empty.textContent = "No paid bills yet.";
            empty.classList.remove("hidden");
        }
        return;
    }

    if (empty) empty.classList.add("hidden");

    bills.forEach((bill) => {
        const tenantName = escapeHtml(bill?.tenantName || bill?.tenant_name || "Unknown tenant");
        const unitLabel = escapeHtml(buildUnitLabel(bill));
        const monthLabel = escapeHtml(formatMonthLabel(bill?.monthKey, bill?.monthLabel));
        const amountLabel = escapeHtml(
            formatAmountLabel(bill?.totalAmount ?? bill?.total_amount ?? bill?.amountPaid ?? bill?.amount_paid ?? 0)
        );
        const billLineId = escapeHtml(getBillLineId(bill));

        const item = document.createElement("div");
        const isDashboard = variant === "dashboard";
        item.className = isDashboard
            ? "flex items-center justify-between gap-2 rounded-lg border border-slate-200/60 bg-white px-2 py-2 hover:bg-slate-50"
            : "flex items-center justify-between gap-2 rounded-md px-1.5 py-1 hover:bg-slate-50";
        item.innerHTML = `
            <div class="min-w-0">
                <div class="${isDashboard ? "text-[12px]" : "text-[10px]"} font-semibold text-slate-800 truncate">${tenantName}</div>
                <div class="${isDashboard ? "text-[10px]" : "text-[9px]"} text-slate-500 truncate">${unitLabel} | ${monthLabel}</div>
            </div>
            <div class="text-right">
                <div class="${isDashboard ? "text-[12px]" : "text-[10px]"} font-semibold text-slate-700">${amountLabel}</div>
                <button type="button"
                    class="${isDashboard ? "text-[10px]" : "text-[8px]"} font-semibold text-indigo-600 hover:underline"
                    data-bill-line-id="${billLineId}">
                    View
                </button>
            </div>
        `;
        list.appendChild(item);
    });
}

function setLoadingState() {
    recentPaymentsState.widgets.forEach(({ list, empty }) => {
        if (list) list.innerHTML = "";
        if (empty) {
            empty.textContent = "Loading recent payments...";
            empty.classList.remove("hidden");
        }
    });
}

async function loadRecentPayments(options = {}) {
    if (recentPaymentsState.loadPromise) return recentPaymentsState.loadPromise;
    setLoadingState();
    recentPaymentsState.loading = true;

    const run = (async () => {
        try {
            const bills = await fetchPaidBillsSummary({
                limit: MAX_RECENT_PAYMENTS,
                force: options.force === true,
            });
            recentPaymentsState.bills = Array.isArray(bills) ? bills : [];
        } catch (err) {
            console.error("Failed to load recent payments", err);
            recentPaymentsState.bills = [];
        }

        recentPaymentsState.widgets.forEach((widget) => {
            renderWidget(widget, recentPaymentsState.bills);
        });
    })();

    recentPaymentsState.loadPromise = run;
    try {
        await run;
    } finally {
        recentPaymentsState.loading = false;
        recentPaymentsState.loadPromise = null;
    }
}

function bindViewHandler(list) {
    if (!list || list.dataset.bound === "true") return;
    list.dataset.bound = "true";
    list.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        const button = target.closest("button[data-bill-line-id]");
        if (!button) return;
        const billLineId = button.dataset.billLineId || "";
        if (!billLineId) return;
        const match = recentPaymentsState.bills.find(
            (bill) => getBillLineId(bill) === billLineId
        );
        if (match) {
            openPaidBillDetails(match);
        }
    });
}

export function initRecentPayments() {
    if (recentPaymentsInitialized) return;
    recentPaymentsInitialized = true;

    const dashboardList = document.getElementById("dashboardRecentPaymentsList");
    const dashboardEmpty = document.getElementById("dashboardRecentPaymentsEmpty");
    const sidebarList = document.getElementById("sidebarRecentPaymentsList");
    const sidebarEmpty = document.getElementById("sidebarRecentPaymentsEmpty");

    const widgets = [];
    if (dashboardList) {
        widgets.push({ list: dashboardList, empty: dashboardEmpty, variant: "dashboard" });
        bindViewHandler(dashboardList);
    }
    if (sidebarList) {
        widgets.push({ list: sidebarList, empty: sidebarEmpty, variant: "sidebar" });
        bindViewHandler(sidebarList);
    }

    if (!widgets.length) return;
    recentPaymentsState.widgets = widgets;

    document.addEventListener("payment:saved", () => {
        loadRecentPayments({ force: true });
    });

    document.addEventListener("sync:completed", () => {
        loadRecentPayments({ force: true });
    });

    document.addEventListener("paid-bills:updated", () => {
        loadRecentPayments();
    });

    loadRecentPayments();
}
