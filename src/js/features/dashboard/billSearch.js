import { fetchGeneratedBills } from "../../api/sheets.js";
import { formatCurrency } from "../../utils/formatters.js";
import { escapeHtml } from "../../utils/htmlUtils.js";

const SEARCH_DEBOUNCE_MS = 140;
const MAX_RESULTS = 8;

const billSearchState = {
    loaded: false,
    loading: false,
    loadPromise: null,
    bills: [],
    results: [],
    requestId: 0,
};

function getBillIdValue(bill) {
    return (bill?.billId || bill?.bill_id || "").toString().trim();
}

function getBillLineIdValue(bill) {
    return (bill?.billLineId || bill?.bill_line_id || "").toString().trim();
}

function normalizeSearchValue(value) {
    return (value || "").toString().trim().toLowerCase();
}

function formatMonthLabel(monthKey, monthLabel) {
    const label = (monthLabel || "").toString().trim();
    const normalized = (monthKey || "").toString().trim();
    const isIsoLabel = label.includes("T") || label.includes("Z");
    const candidate = label && !isIsoLabel ? label : normalized;
    if (/^\d{4}-\d{2}$/.test(candidate)) {
        const [year, month] = candidate.split("-");
        const d = new Date(Number(year), Number(month) - 1, 1);
        if (!isNaN(d.getTime())) {
            return d.toLocaleDateString("en-GB", { month: "short", year: "numeric" });
        }
    }
    return candidate || "-";
}

function buildUnitLabel(bill) {
    const wing = (bill?.wing || "").toString().trim();
    const unitNumber = (bill?.unitNumber || bill?.unit_number || "").toString().trim();
    const parts = [wing, unitNumber].filter(Boolean);
    return parts.length ? parts.join(" - ") : "-";
}

async function loadBills({ force = false, statusEl } = {}) {
    if (billSearchState.loading && billSearchState.loadPromise) {
        return billSearchState.loadPromise;
    }
    if (billSearchState.loaded && !force) return billSearchState.bills;
    billSearchState.loading = true;
    if (statusEl) statusEl.textContent = "Loading bills...";
    const run = (async () => {
        try {
            const data = await fetchGeneratedBills();
            billSearchState.bills = Array.isArray(data?.bills) ? data.bills : [];
            billSearchState.loaded = true;
        } finally {
            billSearchState.loading = false;
            billSearchState.loadPromise = null;
            if (statusEl) {
                statusEl.textContent = billSearchState.bills.length
                    ? `Loaded ${billSearchState.bills.length}`
                    : "No bills yet";
            }
        }
        return billSearchState.bills;
    })();
    billSearchState.loadPromise = run;
    return run;
}

function filterBills(query, bills) {
    if (!query) return [];
    const needle = normalizeSearchValue(query);
    if (!needle) return [];
    return bills
        .filter((bill) => normalizeSearchValue(getBillIdValue(bill)).includes(needle))
        .slice(0, MAX_RESULTS);
}

function renderResults({ results, dropdown, list, empty }) {
    if (!dropdown || !list) return;
    list.innerHTML = "";
    if (!results.length) {
        if (empty) empty.classList.remove("hidden");
        dropdown.classList.remove("hidden");
        return;
    }
    if (empty) empty.classList.add("hidden");

    results.forEach((bill) => {
        const tenantName = escapeHtml(bill?.tenantName || bill?.tenant_name || "Unknown tenant");
        const unitLabel = escapeHtml(buildUnitLabel(bill));
        const monthLabel = escapeHtml(formatMonthLabel(bill?.monthKey, bill?.monthLabel));
        const totalLabel = escapeHtml(formatCurrency(bill?.totalAmount ?? bill?.total_amount ?? 0));
        const billIdLabel = escapeHtml(getBillIdValue(bill) || "-");
        const billLineId = getBillLineIdValue(bill);

        const row = document.createElement("button");
        row.type = "button";
        row.className = "w-full text-left px-3 py-2 hover:bg-slate-50";
        row.dataset.billLineId = billLineId;
        row.innerHTML = `
            <div class="flex items-start justify-between gap-2">
                <div>
                    <div class="text-[12px] font-semibold text-slate-900">${tenantName}</div>
                    <div class="text-[10px] text-slate-500">${unitLabel} | ${monthLabel}</div>
                    <div class="text-[10px] text-slate-400 mt-0.5">Bill ID: ${billIdLabel}</div>
                </div>
                <div class="text-[12px] font-semibold text-slate-800">${totalLabel}</div>
            </div>
        `;
        list.appendChild(row);
    });

    dropdown.classList.remove("hidden");
}

function hideDropdown(dropdown) {
    if (dropdown) dropdown.classList.add("hidden");
}

async function openBillFromResults(bill) {
    if (!bill) return;
    const mod = await import("../billing/payments.js");
    if (typeof mod.openBillFromDashboard === "function") {
        await mod.openBillFromDashboard(bill);
    }
}

export function initBillIdSearch() {
    const input = document.getElementById("billIdSearchInput");
    const dropdown = document.getElementById("billIdSearchDropdown");
    const list = document.getElementById("billIdSearchList");
    const empty = document.getElementById("billIdSearchEmpty");
    const statusEl = document.getElementById("billIdSearchStatus");
    if (!input || !dropdown || !list) return;

    let debounceTimer = null;
    const runSearch = async (query) => {
        const requestId = ++billSearchState.requestId;
        const bills = await loadBills({ statusEl });
        if (requestId !== billSearchState.requestId) return;
        const results = filterBills(query, bills);
        billSearchState.results = results;
        renderResults({ results, dropdown, list, empty });
    };

    const handleInput = () => {
        const query = input.value || "";
        if (!query.trim()) {
            billSearchState.results = [];
            if (empty) empty.classList.add("hidden");
            hideDropdown(dropdown);
            return;
        }
        if (debounceTimer) clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
            runSearch(query);
        }, SEARCH_DEBOUNCE_MS);
    };

    input.addEventListener("input", handleInput);
    input.addEventListener("focus", () => {
        if (input.value.trim()) handleInput();
    });
    input.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
            hideDropdown(dropdown);
        }
        if (event.key === "Enter" && billSearchState.results.length === 1) {
            event.preventDefault();
            openBillFromResults(billSearchState.results[0]);
            hideDropdown(dropdown);
        }
    });

    list.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        const button = target.closest("button[data-bill-line-id]");
        if (!button) return;
        const billLineId = button.dataset.billLineId || "";
        const match = billSearchState.results.find(
            (bill) => getBillLineIdValue(bill) === billLineId
        );
        if (!match) return;
        openBillFromResults(match);
        hideDropdown(dropdown);
    });

    document.addEventListener("click", (event) => {
        if (dropdown.classList.contains("hidden")) return;
        if (dropdown.contains(event.target) || input.contains(event.target)) return;
        hideDropdown(dropdown);
    });

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape") hideDropdown(dropdown);
    });

    document.addEventListener("sync:completed", () => {
        loadBills({ force: true, statusEl }).then(() => {
            if (input.value.trim()) handleInput();
        });
    });

}
