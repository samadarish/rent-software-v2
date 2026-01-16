import { fetchGeneratedBills } from "../../api/sheets.js";
import { formatCurrency } from "../../utils/formatters.js";
import { escapeHtml } from "../../utils/htmlUtils.js";

const SEARCH_DEBOUNCE_MS = 140;
const MAX_RESULTS = 8;
const MAX_RECENT_SEARCHES = 3;
const RECENT_SEARCHES_KEY = "billIdRecentSearches";

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

// --- Recent Searches Functions ---

function getRecentSearches() {
    try {
        const stored = localStorage.getItem(RECENT_SEARCHES_KEY);
        return stored ? JSON.parse(stored) : [];
    } catch {
        return [];
    }
}

function saveRecentSearch(bill) {
    if (!bill) return;
    const billId = getBillIdValue(bill);
    const billLineId = getBillLineIdValue(bill);
    if (!billId || !billLineId) return;

    const entry = {
        billId,
        billLineId,
        tenantName: bill?.tenantName || bill?.tenant_name || "Unknown",
        wing: bill?.wing || "",
        unitNumber: bill?.unitNumber || bill?.unit_number || "",
        monthKey: bill?.monthKey || "",
        monthLabel: bill?.monthLabel || "",
        totalAmount: bill?.totalAmount ?? bill?.total_amount ?? 0,
        savedAt: Date.now(),
    };

    let recent = getRecentSearches();
    // Remove existing entry with same billLineId
    recent = recent.filter((r) => r.billLineId !== billLineId);
    // Add to front
    recent.unshift(entry);
    // Keep only last MAX_RECENT_SEARCHES
    recent = recent.slice(0, MAX_RECENT_SEARCHES);

    try {
        localStorage.setItem(RECENT_SEARCHES_KEY, JSON.stringify(recent));
        // Dispatch event so all search widgets can refresh their recent searches
        document.dispatchEvent(new CustomEvent("recentSearches:updated"));
    } catch {
        // Ignore storage errors
    }
}

function renderRecentSearches(container, onOpenBill) {
    if (!container) return;
    const recent = getRecentSearches();
    container.innerHTML = "";

    if (!recent.length) {
        container.innerHTML = `<span class="text-[9px] text-slate-400 italic">No recent searches</span>`;
        return;
    }

    recent.forEach((entry) => {
        const tenantName = escapeHtml(entry.tenantName || "Unknown");
        const billIdDisplay = escapeHtml(entry.billId || "-");
        const unitLabel = escapeHtml(buildUnitLabel(entry));
        const monthLabel = escapeHtml(formatMonthLabel(entry.monthKey, entry.monthLabel));

        const item = document.createElement("button");
        item.type = "button";
        item.className = "w-full flex items-center justify-between gap-2 px-2 py-1 rounded hover:bg-slate-100 text-left group transition";
        item.dataset.recentBillLineId = entry.billLineId;
        item.innerHTML = `
            <div class="min-w-0 flex-1">
                <div class="text-[10px] font-medium text-slate-700 truncate">${tenantName}</div>
                <div class="text-[9px] text-slate-400 truncate">${unitLabel} • ${monthLabel}</div>
            </div>
            <div class="text-[8px] text-slate-400 font-mono shrink-0 max-w-[80px] truncate">${billIdDisplay}</div>
        `;
        item.title = `Open bill: ${billIdDisplay}`;

        item.addEventListener("click", () => {
            if (typeof onOpenBill === "function") {
                onOpenBill(entry);
            }
        });

        container.appendChild(item);
    });
}

/**
 * Creates a bill search widget for a given set of DOM element IDs.
 * Allows the same logic to power multiple search widgets (dashboard, sidebar, etc.)
 */
function createBillSearchWidget({ inputId, dropdownId, listId, emptyId, statusId, recentContainerId }) {
    const input = document.getElementById(inputId);
    const dropdown = document.getElementById(dropdownId);
    const list = document.getElementById(listId);
    const empty = document.getElementById(emptyId);
    const statusEl = document.getElementById(statusId);
    const recentContainer = document.getElementById(recentContainerId);
    if (!input || !dropdown || !list) return null;

    let debounceTimer = null;
    let localResults = [];

    const runSearch = async (query) => {
        const requestId = ++billSearchState.requestId;
        const bills = await loadBills({ statusEl });
        if (requestId !== billSearchState.requestId) return;
        const results = filterBills(query, bills);
        localResults = results;
        billSearchState.results = results;
        renderResults({ results, dropdown, list, empty });
    };

    const handleInput = () => {
        const query = input.value || "";
        if (!query.trim()) {
            localResults = [];
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

    // Handler for opening a bill (from search or recent)
    const handleOpenBill = async (bill) => {
        saveRecentSearch(bill);
        await openBillFromResults(bill);
        hideDropdown(dropdown);
        input.value = "";
        // Refresh recent searches display
        if (recentContainer) {
            renderRecentSearches(recentContainer, handleOpenBill);
        }
    };

    // Handler for opening from recent searches
    const handleOpenFromRecent = async (entry) => {
        // Try to find the full bill data in loaded bills
        const bills = await loadBills({ statusEl });
        const match = bills.find((b) => getBillLineIdValue(b) === entry.billLineId);
        if (match) {
            await handleOpenBill(match);
        } else {
            // Fallback: use the saved entry data
            await handleOpenBill(entry);
        }
    };

    input.addEventListener("input", handleInput);
    input.addEventListener("focus", () => {
        if (input.value.trim()) handleInput();
    });
    input.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
            hideDropdown(dropdown);
        }
        if (event.key === "Enter" && localResults.length === 1) {
            event.preventDefault();
            handleOpenBill(localResults[0]);
        }
    });

    list.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        const button = target.closest("button[data-bill-line-id]");
        if (!button) return;
        const billLineId = button.dataset.billLineId || "";
        const match = localResults.find(
            (bill) => getBillLineIdValue(bill) === billLineId
        );
        if (!match) return;
        handleOpenBill(match);
    });

    document.addEventListener("click", (event) => {
        if (dropdown.classList.contains("hidden")) return;
        if (dropdown.contains(event.target) || input.contains(event.target)) return;
        hideDropdown(dropdown);
    });

    document.addEventListener("sync:completed", () => {
        loadBills({ force: true, statusEl }).then(() => {
            if (input.value.trim()) handleInput();
        });
    });

    // Refresh bill cache when a payment is saved
    document.addEventListener("payment:saved", () => {
        loadBills({ force: true, statusEl });
    });

    // Listen for recent searches updates from other widgets
    document.addEventListener("recentSearches:updated", () => {
        if (recentContainer) {
            renderRecentSearches(recentContainer, handleOpenFromRecent);
        }
    });

    // Initial render of recent searches
    if (recentContainer) {
        renderRecentSearches(recentContainer, handleOpenFromRecent);
    }

    return { input, dropdown, list, empty, statusEl, handleInput };
}

export function initBillIdSearch() {
    createBillSearchWidget({
        inputId: "billIdSearchInput",
        dropdownId: "billIdSearchDropdown",
        listId: "billIdSearchList",
        emptyId: "billIdSearchEmpty",
        statusId: "billIdSearchStatus",
        recentContainerId: "billIdRecentSearches",
    });
}

export function initSidebarBillIdSearch() {
    createBillSearchWidget({
        inputId: "sidebarBillIdSearchInput",
        dropdownId: "sidebarBillIdSearchDropdown",
        listId: "sidebarBillIdSearchList",
        emptyId: "sidebarBillIdSearchEmpty",
        statusId: "sidebarBillIdSearchStatus",
        recentContainerId: "sidebarBillIdRecentSearches",
    });
}
