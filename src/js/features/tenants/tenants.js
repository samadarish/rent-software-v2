import {
    fetchLandlordsFromSheet,
    fetchTenantDirectory,
    fetchUnitsFromSheet,
    getRentRevisions,
    saveRentRevision,
    updateTenantRecord,
} from "../../api/sheets.js";
import { toOrdinal } from "../../utils/formatters.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

let tenantCache = [];
let tenantRowsCache = [];
let hasLoadedTenants = false;
let unitCache = [];
let landlordCache = [];
let currentStatusFilter = "active"; // all | active | inactive
let currentSearch = "";
let activeTenantForModal = null;
let activeRentRevisions = [];
let activeRentHistoryContext = null;
let selectedTenantForSidebar = null;
let pendingVacateTenant = null;
let pendingNewTenancyTenant = null;
let tenantModalEditable = false;
let tenantModalMode = "tenant"; // tenant | tenancy

const statusClassMap = {
    active: "bg-emerald-100 text-emerald-700 border-emerald-200",
    inactive: "bg-rose-100 text-rose-700 border-rose-200",
};

/**
 * Returns all active tenants for a given wing, ensuring duplicates are filtered out.
 * @param {string} wing - Wing identifier from the dropdown.
 * @returns {Array} Filtered tenant list.
 */
export function getActiveTenantsForWing(wing) {
    const normalizedWing = (wing || "").toString().trim().toLowerCase();
    if (!normalizedWing) return [];

    const source = tenantRowsCache.length ? tenantRowsCache : tenantCache;

    return source.filter((t) => {
        const matchesWing = (t.wing || "").toString().trim().toLowerCase() === normalizedWing;
        const isActive = !!t.activeTenant;
        return matchesWing && isActive;
    });
}

/**
 * Returns a shallow copy of the current tenant cache for read-only usage.
 */
export function getTenantDirectorySnapshot() {
    return [...tenantCache];
}

/**
 * Formats the rent amount for display in the tenant list.
 * @param {number|string} amount
 * @returns {string}
 */
function formatRent(amount) {
    if (!amount) return "-";
    const numeric = parseInt(amount, 10);
    if (isNaN(numeric)) return `₹${amount}`;
    return `₹${numeric.toLocaleString("en-IN")}`;
}

function pickRentValue(...candidates) {
    for (const val of candidates) {
        if (val !== null && typeof val !== "undefined" && val !== "") {
            return val;
        }
    }
    return "";
}

/**
 * Combines wing and floor data into a concise badge label for UI chips.
 * @param {object} tenant
 * @returns {string}
 */
function formatWingFloor(tenant) {
    const wing = tenant.wing || "-";
    const floor = tenant.floor || "-";
    return `${wing} / ${floor}`;
}

/**
 * Converts date inputs into the browser-friendly YYYY-MM-DD format.
 * @param {string|Date} raw
 * @returns {string}
 */
function formatDateForInput(raw) {
    if (!raw) return "";
    const d = new Date(raw);
    if (isNaN(d)) return "";
    const year = d.getFullYear();
    const month = `${d.getMonth() + 1}`.padStart(2, "0");
    const day = `${d.getDate()}`.padStart(2, "0");
    return `${year}-${month}-${day}`;
}

/**
 * Ensures a falsy-safe, formatted date string for any provided value.
 * @param {string|Date} raw
 * @returns {string}
 */
function normalizeDateInputValue(raw) {
    const formatted = formatDateForInput(raw);
    return formatted || "";
}

/**
 * Computes a unique identity key for a tenant entry to collapse duplicates.
 * @param {object} raw
 * @returns {string}
 */
function getTenantIdentityKey(raw) {
    const candidates = [
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
        const key = (candidate || "").toString().trim().toLowerCase();
        if (key) return key;
    }

    return "";
}

/**
 * Builds a stable key for grouping tenant records into a single row.
 * @param {object} raw
 * @returns {string}
 */
function getTenantGroupKey(raw) {
    const candidates = [raw?.tenantId, raw?.tenant_id, raw?.grnNumber, raw?.grn_number, raw?.tenantFullName, raw?.tenant_name];
    for (const candidate of candidates) {
        const key = (candidate || "").toString().trim().toLowerCase();
        if (key) return key;
    }
    return "";
}

/**
 * Returns a deterministic key for tenancy history rows to de-dupe merged tenants.
 * @param {object} history
 * @returns {string}
 */
function getTenancyHistoryKey(history = {}) {
    const candidates = [history.tenancyId, history.tenancy_id, history.tenancy_id, history.grnNumber];
    for (const candidate of candidates) {
        const key = (candidate || "").toString().trim().toLowerCase();
        if (key) return key;
    }
    const start = history.startDate || history.start_date || "";
    const unit = history.unitLabel || history.unit || "";
    const end = history.endDate || history.end_date || "";
    return `${unit}::${start}::${end}`.toLowerCase();
}

/**
 * Normalizes and merges tenancy history entries for a tenant to avoid duplicates.
 * @param {Array} existing
 * @param {Array} incoming
 * @returns {Array}
 */
function mergeTenancyHistory(existing = [], incoming = []) {
    const map = new Map();
    [...existing, ...incoming].forEach((entry = {}) => {
        const key = getTenancyHistoryKey(entry);
        if (!key) return;
        const current = map.get(key) || {};
        map.set(key, { ...current, ...entry });
    });
    return Array.from(map.values());
}

/**
 * Collapses multiple tenant rows (often one per tenancy) into a single tenant entry with merged history.
 * @param {Array} rawTenants
 * @returns {Array}
 */
function collapseTenantRows(rawTenants = []) {
    const grouped = new Map();
    rawTenants.forEach((t, idx) => {
        const key = getTenantGroupKey(t) || `row-${idx}`;
        const history = Array.isArray(t.tenancyHistory) ? t.tenancyHistory : [];
        if (grouped.has(key)) {
            const existing = grouped.get(key);
            const mergedHistory = mergeTenancyHistory(existing.tenancyHistory, history);
            grouped.set(key, {
                ...existing,
                ...t,
                tenancyHistory: mergedHistory,
                activeTenant: existing.activeTenant || t.activeTenant,
                family: existing.family && existing.family.length ? existing.family : t.family || [],
            });
        } else {
            grouped.set(key, { ...t, tenancyHistory: history });
        }
    });
    return Array.from(grouped.values());
}

/**
 * Formats tenancy end dates for table display.
 * @param {string} raw
 * @returns {string}
 */
function formatTenancyEndDate(raw) {
    if (!raw) return "-";
    const formatted = formatDateForInput(raw);
    return formatted || raw;
}

function renderRentHistory(baseRent) {
    const tbody = document.getElementById("rentHistoryTableBody");
    const empty = document.getElementById("rentHistoryEmpty");
    const currentRentEl = document.getElementById("rentHistoryCurrentValue");
    if (!tbody || !empty) return;

    tbody.innerHTML = "";
    const revisions = Array.isArray(activeRentRevisions) ? activeRentRevisions : [];
    if (!revisions.length) {
        empty.classList.remove("hidden");
    } else {
        empty.classList.add("hidden");
        revisions.forEach((rev) => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td class="px-2 py-1 text-[11px] text-slate-700">${formatMonthLabel(rev.effective_month)}</td>
                <td class="px-2 py-1 text-[11px] font-semibold">₹${(Number(rev.rent_amount) || 0).toLocaleString("en-IN")}</td>
                <td class="px-2 py-1 text-[11px] text-slate-600">${rev.note || ""}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    if (currentRentEl) {
        const today = new Date();
        const monthKey = `${today.getFullYear()}-${`${today.getMonth() + 1}`.padStart(2, "0")}`;
        const effective = getEffectiveRentFromRevisions(revisions, monthKey, baseRent);
        currentRentEl.textContent =
            effective === null || typeof effective === "undefined"
                ? "-"
                : `₹${(Number(effective) || 0).toLocaleString("en-IN")}`;
    }
}

function toggleTenantModalSections(mode) {
    tenantModalMode = mode;
    const tenantSections = document.querySelectorAll(".tenant-section");
    const tenancySections = document.querySelectorAll(".tenancy-section");

    tenantSections.forEach((el) => el.classList.toggle("hidden", mode !== "tenant"));
    tenancySections.forEach((el) => el.classList.toggle("hidden", mode !== "tenancy"));
}

async function loadRentHistoryForTenancy(tenancyId, baseRent) {
    const loader = document.getElementById("rentHistoryLoader");
    if (loader) loader.classList.remove("hidden");
    activeRentRevisions = [];
    renderRentHistory(baseRent);
    if (!tenancyId) {
        if (loader) loader.classList.add("hidden");
        return;
    }
    try {
        const res = await getRentRevisions(tenancyId);
        activeRentRevisions = Array.isArray(res?.revisions) ? res.revisions : [];
    } catch (e) {
        console.error("Failed to load rent revisions", e);
    }
    renderRentHistory(baseRent);
    if (loader) loader.classList.add("hidden");
}

async function handleRentRevisionSave() {
    if (!activeRentHistoryContext) return;
    const tenancyId = activeRentHistoryContext.tenancyId || activeRentHistoryContext.templateData?.tenancy_id;
    if (!tenancyId) {
        showToast("Save tenancy first to add revisions", "warning");
        return;
    }

    const monthInput = document.getElementById("rentRevisionMonth");
    const amountInput = document.getElementById("rentRevisionAmount");
    const noteInput = document.getElementById("rentRevisionNote");

    const effectiveMonth = normalizeMonthKey(monthInput?.value || "");
    const rentAmount = Number(amountInput?.value || 0);
    const note = noteInput?.value || "";

    if (!effectiveMonth || !/^\d{4}-\d{2}$/.test(effectiveMonth)) {
        showToast("Enter a valid effective month (YYYY-MM)", "warning");
        return;
    }
    if (isNaN(rentAmount) || rentAmount <= 0) {
        showToast("Enter a valid rent amount", "warning");
        return;
    }

    const saveBtn = document.getElementById("rentRevisionSaveBtn");
    if (saveBtn) saveBtn.disabled = true;
    try {
        const res = await saveRentRevision({ tenancyId, effectiveMonth, rentAmount, note });
        if (res?.ok) {
            activeRentRevisions = Array.isArray(res.revisions) ? res.revisions : activeRentRevisions;
            const baseRent =
                activeRentHistoryContext.rentAmount ||
                activeRentHistoryContext.templateData?.rent_amount ||
                activeRentHistoryContext.currentRent ||
                "";
            renderRentHistory(baseRent);
            if (monthInput) monthInput.value = "";
            if (amountInput) amountInput.value = "";
            if (noteInput) noteInput.value = "";
        }
    } catch (e) {
        console.error("Failed to save rent revision", e);
    } finally {
        if (saveBtn) saveBtn.disabled = false;
    }
}

function normalizeMonthKey(raw) {
    if (!raw) return "";
    const str = raw.toString().trim();
    const match = str.match(/^(\d{4})-(\d{1,2})/);
    if (match) return `${match[1]}-${match[2].padStart(2, "0")}`;
    const compact = str.match(/^(\d{4})(\d{2})$/);
    if (compact) return `${compact[1]}-${compact[2]}`;
    return str;
}

function formatMonthLabel(monthKey) {
    const normalized = normalizeMonthKey(monthKey);
    if (!normalized) return "";
    const match = normalized.match(/^(\d{4})-(\d{2})$/);
    if (!match) return normalized;
    const d = new Date(Number(match[1]), Number(match[2]) - 1, 1);
    return d.toLocaleDateString("en-US", { month: "short", year: "numeric" });
}

function getEffectiveRentFromRevisions(revisions, monthKey, baseRent) {
    const normalizedMonth = normalizeMonthKey(monthKey);
    if (!normalizedMonth) return baseRent ?? null;
    const ordered = [...(revisions || [])].sort(
        (a, b) => (normalizeMonthKey(b.effective_month) || "").localeCompare(normalizeMonthKey(a.effective_month) || "")
    );
    for (const rev of ordered) {
        const eff = normalizeMonthKey(rev.effective_month);
        if (eff && eff <= normalizedMonth) {
            const amt = Number(rev.rent_amount);
            return isNaN(amt) ? baseRent ?? null : amt;
        }
    }
    return baseRent ?? null;
}

/**
 * Clones option elements from one select to another while preserving selection.
 * @param {string} sourceId
 * @param {string} targetId
 */
function cloneSelectOptions(sourceId, targetId) {
    const source = document.getElementById(sourceId);
    const target = document.getElementById(targetId);
    if (!source || !target) return;
    if (!source.options || !target.options) return;

    const previous = target.value;
    target.innerHTML = "";
    Array.from(source.options).forEach((opt) => {
        const clone = opt.cloneNode(true);
        target.appendChild(clone);
    });

    if (previous && Array.from(target.options).some((o) => o.value === previous)) {
        target.value = previous;
    }
}

/**
 * Refreshes all tenant modal picklists so they match the primary form values.
 */
function syncTenantModalPicklists() {
    cloneSelectOptions("wing", "tenantModalWing");
    cloneSelectOptions("payable_date", "tenantModalPayable");
    cloneSelectOptions("notice_num_t", "tenantModalTenantNotice");
    cloneSelectOptions("notice_num_l", "tenantModalLandlordNotice");
    cloneSelectOptions("landlord_selector", "tenantModalLandlord");
}

/**
 * Builds the list of available units in the tenant modal while respecting occupancy.
 * @param {string} selectedUnitId
 * @param {string} tenancyId
 */
function populateUnitDropdown(selectedUnitId, tenancyId) {
    const select = document.getElementById("tenantModalUnit");
    if (!select) return;
    const previous = select.value;
    select.innerHTML = '<option value="">Select unit</option>';

    const available = unitCache.filter(
        (u) => !u.is_occupied || (tenancyId && u.current_tenancy_id === tenancyId)
    );

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

    if (selectedUnitId && Array.from(select.options).some((o) => o.value === selectedUnitId)) {
        select.value = selectedUnitId;
    } else if (previous && Array.from(select.options).some((o) => o.value === previous)) {
        select.value = previous;
    }
}

/**
 * Applies unit attributes from the cache to the modal fields.
 * @param {string} unitId
 */
function applyUnitSelectionToModal(unitId) {
    if (!unitId) return;
    const unit = unitCache.find((u) => u.unit_id === unitId);
    if (!unit) return;
    const wing = document.getElementById("tenantModalWing");
    const unitNumber = document.getElementById("tenantModalUnitNumber");
    const floor = document.getElementById("tenantModalFloor");
    const direction = document.getElementById("tenantModalDirection");
    const meter = document.getElementById("tenantModalMeter");
    if (wing) wing.value = unit.wing || "";
    if (unitNumber) unitNumber.value = unit.unit_number || "";
    if (floor) floor.value = unit.floor || "";
    if (direction) direction.value = unit.direction || "";
    if (meter) meter.value = unit.meter_number || "";
}

/**
 * Syncs landlord selection into the modal fields for Aadhaar and address.
 * @param {string} landlordId
 */
function applyLandlordSelectionToModal(landlordId) {
    const landlord = landlordCache.find((l) => l.landlord_id === landlordId);
    const select = document.getElementById("tenantModalLandlord");
    if (select && landlordId && select.value !== landlordId) {
        select.value = landlordId;
    }
    const aadhaar = document.getElementById("tenantModalLandlordAadhaar");
    const address = document.getElementById("tenantModalLandlordAddress");
    if (aadhaar) aadhaar.value = landlord?.aadhaar || "";
    if (address) address.value = landlord?.address || "";
}

/**
 * Builds a concise label for a unit for dropdown and table usage.
 * @param {object} unit
 * @returns {string}
 */
function buildUnitLabel(unit) {
    if (!unit) return "";
    const parts = [unit.wing, unit.unit_number].filter(Boolean);
    return parts.join(" - ") || unit.unit_id || "";
}

/**
 * Updates cached unit occupancy to keep dropdowns in sync without refetching.
 * @param {string} unitId
 * @param {boolean} occupied
 * @param {string} tenancyId
 */
function markUnitOccupancy(unitId, occupied, tenancyId) {
    if (!unitId) return;
    const target = unitCache.find((u) => u.unit_id === unitId);
    if (target) {
        target.is_occupied = occupied;
        target.current_tenancy_id = occupied ? tenancyId : "";
    }
}

/**
 * Normalizes day-of-month values by stripping non-digits for storage/display.
 * @param {string|number} val
 * @returns {string}
 */
function normalizeDayValue(val) {
    if (!val) return "";
    const match = String(val).match(/\d+/);
    return match ? match[0] : String(val);
}

/**
 * Updates the badge within the tenant modal to reflect active/inactive status.
 * @param {boolean} active
 */
function setTenantModalStatusPill(active) {
    const pill = document.getElementById("tenantModalStatusPill");
    if (!pill) return;
    pill.textContent = active ? "Active" : "Inactive";
    pill.className = `text-[10px] px-2 py-1 rounded-full border font-semibold ${
        active ? statusClassMap.active : statusClassMap.inactive
    }`;
}

/**
 * Shows or hides tenant list loading states while data is fetched.
 * @param {boolean} isLoading
 */
function setTenantListLoading(isLoading) {
    const table = document.getElementById("tenantTableBody");
    const emptyState = document.getElementById("tenantListEmpty");
    const loader = document.getElementById("tenantListLoader");

    if (loader) loader.classList.toggle("hidden", !isLoading);
    if (table) table.classList.toggle("opacity-50", isLoading);
    if (emptyState && isLoading) emptyState.classList.add("hidden");
}

/**
 * Converts a boolean active flag into user-facing status text.
 * @param {boolean} active
 * @returns {string}
 */
function getStatusLabel(active) {
    return active ? "Active" : "Inactive";
}

/**
 * Builds a status pill element for the tenant list based on active state.
 * @param {boolean} active
 * @returns {HTMLSpanElement}
 */
function renderStatusPill(active) {
    const span = document.createElement("span");
    span.className = `text-[10px] px-2 py-1 rounded-full border font-semibold ${active ? statusClassMap.active : statusClassMap.inactive}`;
    span.textContent = getStatusLabel(active);
    return span;
}

/**
 * Renders tenant rows into the directory table and binds actions.
 * @param {Array} rows
 */
function renderTenantRows(rows) {
    const tbody = document.getElementById("tenantTableBody");
    const emptyState = document.getElementById("tenantListEmpty");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (!rows || rows.length === 0) {
        if (emptyState) emptyState.classList.remove("hidden");
        return;
    }

    if (emptyState) emptyState.classList.add("hidden");

    rows.forEach((t) => {
        const tr = document.createElement("tr");
        tr.className = "border-b last:border-0 hover:bg-slate-50 cursor-pointer";
        tr.dataset.grn = t.grnNumber || "";

        const activeHistory = Array.isArray(t.tenancyHistory)
            ? t.tenancyHistory.filter((h) => (h.status || "").toLowerCase() === "active")
            : [];
        const activeCount = activeHistory.length;
        const unitLabels = activeHistory.map((h) => h.unitLabel).filter(Boolean).join(", ") || "—";
        const isActive = activeCount > 0 || !!t.activeTenant;
        const statusCell = renderStatusPill(isActive);

        tr.innerHTML = `
            <td class="px-2 py-1.5">
                <div class="font-semibold text-xs leading-tight">${t.tenantFullName || "Unnamed"}</div>
            </td>
            <td class="px-2 py-1.5 text-xs">${activeCount}</td>
            <td class="px-2 py-1.5 text-xs">${unitLabels}</td>
            <td class="px-2 py-1.5 text-xs">
                <div class="flex items-center gap-1 status-slot"></div>
            </td>
            <td class="px-2 py-1.5 text-xs">
                <button class="text-indigo-600 hover:text-indigo-800 font-semibold text-[11px] underline">Edit Tenant</button>
            </td>
        `;

        const statusSlot = tr.querySelector(".status-slot");
        if (statusSlot) statusSlot.appendChild(statusCell);

        const actionBtn = tr.querySelector("button");
        if (actionBtn) {
            actionBtn.addEventListener("click", (e) => {
                e.stopPropagation();
                setSidebarSelection(t);
                openTenantModal(t);
            });
        }

        tr.addEventListener("click", () => {
            setSidebarSelection(t);
        });

        tbody.appendChild(tr);
    });

    highlightSelectedRow();
}

/**
 * Applies search and status filters to the cached tenant list before rendering.
 */
function applyTenantFilters() {
    let filtered = [...tenantCache];

    if (currentStatusFilter !== "all") {
        const shouldBeActive = currentStatusFilter === "active";
        filtered = filtered.filter((t) => {
            const activeHistory = Array.isArray(t.tenancyHistory)
                ? t.tenancyHistory.some((h) => (h.status || "").toLowerCase() === "active")
                : false;
            return (activeHistory || !!t.activeTenant) === shouldBeActive;
        });
    }

    if (currentSearch) {
        const q = currentSearch.toLowerCase();
        filtered = filtered.filter((t) => {
            return (
                (t.tenantFullName || "").toLowerCase().includes(q) ||
                (t.wing || "").toLowerCase().includes(q) ||
                (t.floor || "").toLowerCase().includes(q) ||
                (t.grnNumber || "").toLowerCase().includes(q)
            );
        });
    }

    renderTenantRows(filtered);
    const totalLabel = document.getElementById("tenantListCount");
    if (totalLabel) totalLabel.textContent = `${filtered.length} shown / ${tenantCache.length} total`;
}

/**
 * Updates the active/inactive filter buttons to reflect the selected filter.
 */
function syncStatusButtons() {
    document.querySelectorAll("[data-tenant-status-filter]").forEach((btn) => {
        const val = btn.getAttribute("data-tenant-status-filter");
        const isActive = val === currentStatusFilter;
        btn.classList.toggle("bg-slate-900", isActive);
        btn.classList.toggle("text-white", isActive);
        btn.classList.toggle("shadow-sm", isActive);
        btn.classList.toggle("bg-white", !isActive);
        btn.classList.toggle("text-slate-700", !isActive);
    });
}

/**
 * Ensures tenants are loaded exactly once; subsequent calls are no-ops.
 */
export async function ensureTenantDirectoryLoaded() {
    if (hasLoadedTenants) return;
    await loadTenantDirectory();
}

/**
 * Fetches tenant, unit, and landlord data then renders the directory.
 * @param {boolean} forceReload - When true, bypasses the cache.
 */
export async function loadTenantDirectory(forceReload = false) {
    const previousSelectedGrn = selectedTenantForSidebar?.grnNumber;
    if (hasLoadedTenants && !forceReload) {
        applyTenantFilters();
        return;
    }

    setTenantListLoading(true);
    const [unitData, landlordData, data] = await Promise.all([
        fetchUnitsFromSheet(),
        fetchLandlordsFromSheet(),
        fetchTenantDirectory(),
    ]);
    unitCache = Array.isArray(unitData.units)
        ? unitData.units.map((u) => ({
              ...u,
              is_occupied:
                  u.is_occupied === true ||
                  (typeof u.is_occupied === "string" && u.is_occupied.toLowerCase() === "true"),
          }))
        : [];
    landlordCache = Array.isArray(landlordData.landlords) ? landlordData.landlords : [];
    const normalizedTenants = Array.isArray(data.tenants)
        ? data.tenants.map((t) => ({
              ...t,
              tenancyHistory: Array.isArray(t.tenancyHistory) ? t.tenancyHistory : [],
              family: Array.isArray(t.family) ? t.family : [],
          }))
        : [];
    tenantRowsCache = normalizedTenants;
    tenantCache = collapseTenantRows(normalizedTenants);
    document.dispatchEvent(new CustomEvent("landlords:updated", { detail: landlordCache }));
    hasLoadedTenants = true;
    setTenantListLoading(false);
    applyTenantFilters();

    if (tenantCache.length) {
        const match = tenantCache.find((t) => t.grnNumber === previousSelectedGrn);
        setSidebarSelection(match || tenantCache[0]);
    } else {
        selectedTenantForSidebar = null;
        updateSidebarSnapshot();
    }
}

function createFamilyRow(member = {}) {
    const tr = document.createElement("tr");
    tr.className = "border-b last:border-0";

    tr.innerHTML = `
        <td class="px-2 py-1.5 text-xs"><input type="text" class="w-full border rounded px-2 py-1 text-xs" value="${member.name || ""}" /></td>
        <td class="px-2 py-1.5 text-xs"><input type="text" class="w-full border rounded px-2 py-1 text-xs" value="${member.relationship || ""}" /></td>
        <td class="px-2 py-1.5 text-xs"><input type="text" class="w-full border rounded px-2 py-1 text-xs" value="${member.occupation || ""}" /></td>
        <td class="px-2 py-1.5 text-xs"><input type="text" class="w-full border rounded px-2 py-1 text-xs" value="${member.aadhaar || ""}" /></td>
        <td class="px-2 py-1.5 text-xs"><input type="text" class="w-full border rounded px-2 py-1 text-xs" value="${member.address || ""}" /></td>
        <td class="px-2 py-1.5 text-xs text-right">
            <button class="text-rose-600 text-[11px] underline tenant-family-remove">Remove</button>
        </td>
    `;

    const removeBtn = tr.querySelector("button");
    if (removeBtn) {
        removeBtn.addEventListener("click", () => {
            tr.remove();
        });
    }

    setTenantModalEditable(tenantModalEditable);

    return tr;
}

function highlightSelectedRow() {
    const tbody = document.getElementById("tenantTableBody");
    if (!tbody) return;
    const selectedGrn = selectedTenantForSidebar?.grnNumber || "";
    tbody.querySelectorAll("tr").forEach((row) => {
        row.classList.toggle("bg-indigo-50", selectedGrn && row.dataset.grn === selectedGrn);
    });
}

function setTenantModalEditable(enabled) {
    tenantModalEditable = enabled;
    const modal = document.getElementById("tenantDetailModal");
    if (!modal) return;

    const alwaysReadOnly = new Set([
        "tenantModalWing",
        "tenantModalUnitNumber",
        "tenantModalFloor",
        "tenantModalDirection",
        "tenantModalMeter",
        "tenantModalLandlordAadhaar",
        "tenantModalLandlordAddress",
    ]);

    modal.querySelectorAll("input, select, textarea").forEach((el) => {
        if (alwaysReadOnly.has(el.id)) {
            el.disabled = true;
            el.classList.add("bg-slate-50", "text-slate-700");
            return;
        }
        el.disabled = !enabled;
    });

    modal.querySelectorAll("#tenantFamilyTableBody button").forEach((btn) => {
        btn.disabled = !enabled;
        btn.classList.toggle("opacity-50", !enabled);
    });

    const addFamilyBtn = document.getElementById("tenantModalAddFamilyBtn");
    if (addFamilyBtn) {
        addFamilyBtn.disabled = !enabled;
        addFamilyBtn.classList.toggle("opacity-60", !enabled);
    }

    const saveBtn = document.getElementById("tenantModalSaveBtn");
    if (saveBtn) {
        saveBtn.disabled = !enabled;
        saveBtn.classList.toggle("opacity-60", !enabled);
    }

    document.querySelectorAll(".tenant-modal-edit-toggle").forEach((btn) => {
        btn.textContent = enabled ? "Editing enabled" : "Edit details";
    });
}

function updateSidebarSnapshot() {
    const emptyState = document.getElementById("sidebarSnapshotEmpty");
    const detailsPanel = document.getElementById("tenantDetailsPanel");
    const nameEl = document.getElementById("sidebarTenantName");
    const statusEl = document.getElementById("sidebarStatusPill");
    const mobileEl = document.getElementById("sidebarTenantMobile");
    const occupationEl = document.getElementById("sidebarTenantOccupation");
    const addressEl = document.getElementById("sidebarTenantAddress");
    const familyCount = document.getElementById("sidebarFamilyCount");
    const familyList = document.getElementById("sidebarFamilyList");
    const historyList = document.getElementById("sidebarUnitHistory");

    if (!selectedTenantForSidebar) {
        if (emptyState) emptyState.classList.remove("hidden");
        if (detailsPanel) detailsPanel.classList.add("hidden");
        if (nameEl) nameEl.textContent = "Select a tenant";
        if (statusEl) {
            statusEl.textContent = "Status";
            statusEl.className = "text-[10px] px-2 py-1 rounded-full border bg-slate-50 text-slate-600";
        }
        if (mobileEl) mobileEl.textContent = "-";
        if (occupationEl) occupationEl.textContent = "-";
        if (addressEl) addressEl.textContent = "-";
        if (familyCount) familyCount.textContent = "0 members";
        if (familyList) familyList.innerHTML = '<li class="text-[10px] text-slate-500">No tenant selected.</li>';
        if (historyList) historyList.innerHTML = '<div class="text-[10px] text-slate-500">No tenant selected.</div>';
        return;
    }

    const t = selectedTenantForSidebar;
    if (emptyState) emptyState.classList.add("hidden");
    if (detailsPanel) detailsPanel.classList.remove("hidden");
    if (nameEl) nameEl.textContent = t.tenantFullName || "Tenant";
    const isActive = (t.tenancyHistory || []).some((h) => (h.status || "").toLowerCase() === "active") || t.activeTenant;
    if (statusEl) {
        statusEl.textContent = getStatusLabel(isActive);
        statusEl.className = `text-[10px] px-2 py-1 rounded-full border font-semibold ${isActive ? statusClassMap.active : statusClassMap.inactive}`;
    }
    if (mobileEl) mobileEl.textContent = t.tenantMobile || "-";
    if (occupationEl) occupationEl.textContent = t.tenantOccupation || "-";
    if (addressEl) addressEl.textContent = t.tenantPermanentAddress || "-";

    if (familyList) {
        familyList.innerHTML = "";
        (t.family || []).forEach((member) => {
            const li = document.createElement("li");
            li.textContent = `${member.name || "Member"} – ${member.relationship || "Relation"}`;
            familyList.appendChild(li);
        });
        if (!(t.family || []).length) {
            familyList.innerHTML = '<li class="text-[10px] text-slate-500">No family members recorded.</li>';
        }
    }
    if (familyCount) {
        familyCount.textContent = `${t.family?.length || 0} members`;
    }

    if (historyList) {
        const history = Array.isArray(t.tenancyHistory) ? [...t.tenancyHistory] : [];
        history.sort((a, b) => {
            const aKey = `${(a.status || "").toLowerCase() === "active" ? 0 : 1}${a.startDate || ""}`;
            const bKey = `${(b.status || "").toLowerCase() === "active" ? 0 : 1}${b.startDate || ""}`;
            return aKey.localeCompare(bKey);
        });
        if (!history.length) {
            historyList.innerHTML = '<div class="text-[10px] text-slate-500">No tenancy history yet.</div>';
        } else {
            historyList.innerHTML = "";
            history.forEach((h) => {
                const card = document.createElement("div");
                card.className = "border rounded-lg p-2 bg-slate-50";
                card.dataset.tenancyId = h.tenancyId || h.tenancy_id || "";
                card.dataset.grn = t.grnNumber || "";
                card.dataset.unitLabel = h.unitLabel || "";
                card.dataset.startDate = h.startDate || "";
                card.dataset.endDate = h.endDate || "";
                card.dataset.status = h.status || "";
                card.dataset.currentRent = pickRentValue(h.currentRent, h.rentAmount, t.currentRent, t.rentAmount, "");
                const statusPill = document.createElement("span");
                statusPill.className = `text-[9px] px-2 py-0.5 rounded-full border ${
                    (h.status || "").toLowerCase() === "active" ? statusClassMap.active : statusClassMap.inactive
                }`;
                statusPill.textContent = h.status || "-";
                const dates = [formatDateForInput(h.startDate) || "", formatTenancyEndDate(h.endDate) || "Present"]
                    .filter(Boolean)
                    .join(" → ");
                card.innerHTML = `
                    <div class="flex items-center justify-between gap-2">
                        <div>
                            <p class="font-semibold text-[12px]">${h.unitLabel || "Unit"}</p>
                            <p class="text-[10px] text-slate-500">${dates}</p>
                            <p class="text-[10px] text-slate-500">Current rent: ${formatRent(
                                h.currentRent ?? t.currentRent ?? t.rentAmount
                            )}</p>
                        </div>
                        <div class="flex flex-col gap-1 items-end">
                            <span class="self-end">${statusPill.outerHTML}</span>
                            <div class="flex gap-1">
                                <button type="button" class="px-2 py-1 rounded text-[10px] bg-white border border-slate-200 font-semibold tenancy-edit-btn">Edit tenancy</button>
                                <button type="button" class="px-2 py-1 rounded text-[10px] bg-white border border-indigo-200 text-indigo-700 font-semibold rent-history-btn">Rent history</button>
                            </div>
                        </div>
                    </div>
                `;
                const pillHolder = card.querySelector("span.self-end");
                if (pillHolder) pillHolder.replaceWith(statusPill);
                const editBtn = card.querySelector(".tenancy-edit-btn");
                const rentBtn = card.querySelector(".rent-history-btn");
                if (editBtn) {
                    editBtn.addEventListener("click", () => openTenancyModal(h, t));
                }
                if (rentBtn) {
                    rentBtn.dataset.tenancyId =
                        card.dataset.tenancyId ||
                        t.tenancyId ||
                        t.tenancy_id ||
                        t.templateData?.tenancy_id ||
                        "";
                    rentBtn.dataset.grn = t.grnNumber || t.grn_number || t.templateData?.grn_number || "";
                    rentBtn.dataset.unitLabel = h.unitLabel || t.unitLabel || t.unitNumber || t.templateData?.unit_number || "";
                    rentBtn.dataset.startDate = h.startDate || t.tenancyCommencement || t.templateData?.tenancy_comm_raw || "";
                    rentBtn.dataset.endDate =
                        h.endDate ||
                        t.tenancyEndRaw ||
                        t.tenancyEndDate ||
                        t.templateData?.tenancy_end_raw ||
                        "";
                    rentBtn.dataset.status = h.status || t.status || "";
                    rentBtn.dataset.currentRent = pickRentValue(
                        h.currentRent,
                        h.rentAmount,
                        t.currentRent,
                        t.rentAmount,
                        t.templateData?.rent_amount,
                        ""
                    );
                    rentBtn.addEventListener("click", (event) => handleRentHistoryClick(event, t));
                }
                historyList.appendChild(card);
            });
        }
    }

    highlightSelectedRow();
}

function setSidebarSelection(tenant) {
    selectedTenantForSidebar = tenant;
    updateSidebarSnapshot();
}

function openTenancyModal(tenancy, tenant) {
    const base = tenant || selectedTenantForSidebar;
    if (!base) return;
    const merged = {
        ...base,
        tenancyId: tenancy?.tenancyId || base.tenancyId,
        unitId: tenancy?.unitId || base.unitId,
        unitNumber: tenancy?.unitLabel || base.unitNumber,
        tenancyCommencementRaw: tenancy?.startDate || base.tenancyCommencement,
        tenancyEndDate:
            tenancy?.endDate || tenancy?.tenancyEndDate || tenancy?.tenancyEndRaw || base.tenancyEndRaw || base.tenancyEndDate,
        rentAmount: pickRentValue(tenancy?.currentRent, base.rentAmount),
        activeTenant: (tenancy?.status || "").toLowerCase() === "active",
    };
    // UX note: tenancy edit is invoked from the unit history action; tenant identity/family stay within tenant mode.
    populateTenantModal(merged, "tenancy");
    setTenantModalEditable(true);
}

function resolveTenantFromElement(element, tenancyId = "") {
    const grn = element?.dataset.grn || element?.dataset.tenantGrn || "";
    if (grn) {
        const match = tenantCache.find((t) => (t.grnNumber || t.grn_number || "").toString() === grn);
        if (match) return match;
    }

    if (tenancyId) {
        const match = tenantCache.find((t) => {
            const target = tenancyId.toString();
            const directIdMatches = [t.tenancyId, t.tenancy_id].some((id) => id && id.toString() === target);
            const historyMatches =
                Array.isArray(t.tenancyHistory) &&
                t.tenancyHistory.some((h) => (h.tenancyId || h.tenancy_id)?.toString() === target);
            return directIdMatches || historyMatches;
        });
        if (match) return match;
    }

    return selectedTenantForSidebar || null;
}

function buildRentHistoryContext(btn, fallbackTenant) {
    const card = btn?.closest?.("[data-tenancy-id]");
    const tenant = fallbackTenant || resolveTenantFromElement(btn || card, btn?.dataset?.tenancyId || card?.dataset?.tenancyId) || {};
    const fallbackTenancyId =
        tenant.tenancyId || tenant.tenancy_id || tenant.templateData?.tenancy_id || tenant.templateData?.tenancyId || "";
    const tenancyId = btn?.dataset?.tenancyId || card?.dataset?.tenancyId || fallbackTenancyId || "";
    const historyEntry = tenancyId
        ? (tenant.tenancyHistory || []).find((h) => (h.tenancyId || h.tenancy_id)?.toString() === tenancyId.toString())
        : (tenant.tenancyHistory || []).find((h) => (h.status || "").toLowerCase() === "active") || null;

    const tenancy = {
        tenancyId: tenancyId || historyEntry?.tenancyId || historyEntry?.tenancy_id,
        unitLabel:
            btn?.dataset?.unitLabel ||
            card?.dataset?.unitLabel ||
            historyEntry?.unitLabel ||
            tenant.unitLabel ||
            tenant.unitNumber ||
            "",
        startDate: btn?.dataset?.startDate || card?.dataset?.startDate || historyEntry?.startDate || tenant.tenancyCommencement,
        endDate:
            btn?.dataset?.endDate ||
            card?.dataset?.endDate ||
            historyEntry?.endDate ||
            tenant.tenancyEndRaw ||
            tenant.tenancyEndDate,
        status: btn?.dataset?.status || card?.dataset?.status || historyEntry?.status || tenant.status,
        currentRent: pickRentValue(
            btn?.dataset?.currentRent,
            card?.dataset?.currentRent,
            historyEntry?.currentRent,
            historyEntry?.rentAmount,
            tenant.currentRent,
            tenant.rentAmount
        ),
        rentAmount: historyEntry?.rentAmount,
    };

    return { tenancy: tenancy.tenancyId ? tenancy : historyEntry || tenancy, tenant };
}

function resolveFallbackTenancyContext(tenant) {
    if (!tenant) return null;

    const history = Array.isArray(tenant.tenancyHistory) ? tenant.tenancyHistory : [];
    const activeHistory = history.find((h) => (h.status || "").toLowerCase() === "active");
    const fallbackHistory = activeHistory || history[0];

    if (fallbackHistory) {
        return {
            tenancyId: fallbackHistory.tenancyId || fallbackHistory.tenancy_id,
            unitLabel:
                fallbackHistory.unitLabel ||
                tenant.unitLabel ||
                tenant.unitNumber ||
                tenant.templateData?.unit_number,
            startDate: fallbackHistory.startDate,
            endDate: fallbackHistory.endDate,
            status: fallbackHistory.status,
            currentRent: pickRentValue(
                fallbackHistory.currentRent,
                fallbackHistory.rentAmount,
                tenant.currentRent,
                tenant.rentAmount
            ),
        };
    }

    if (tenant.tenancyId || tenant.tenancy_id) {
        return {
            tenancyId: tenant.tenancyId || tenant.tenancy_id,
            unitLabel: tenant.unitLabel || tenant.unitNumber || tenant.templateData?.unit_number,
            startDate: tenant.tenancyCommencement,
            endDate: tenant.tenancyEndRaw || tenant.tenancyEndDate,
            status: tenant.status,
            currentRent: pickRentValue(tenant.currentRent, tenant.rentAmount),
        };
    }

    return null;
}

function openRentHistoryModal(tenancy, tenant) {
    const modal = document.getElementById("rentHistoryModal");
    if (!modal) {
        showToast("Rent history modal is unavailable", "error");
        return;
    }

    const baseTenant = tenant || selectedTenantForSidebar || {};
    const resolvedTenancy = tenancy || resolveFallbackTenancyContext(baseTenant);
    if (!resolvedTenancy) {
        showToast("Select a tenancy first", "warning");
        return;
    }
    const title = document.getElementById("rentHistoryTitle");
    if (title) {
        const unitLabel =
            resolvedTenancy.unitLabel || baseTenant.unitNumber || baseTenant.templateData?.unit_number || "Unit";
        title.textContent = `Rent History — ${unitLabel}`;
    }

    activeRentHistoryContext = {
        ...resolvedTenancy,
        templateData: baseTenant.templateData || {},
        rentAmount: pickRentValue(
            resolvedTenancy.currentRent,
            resolvedTenancy.rentAmount,
            baseTenant.currentRent,
            baseTenant.rentAmount,
            baseTenant.templateData?.rent_amount
        ),
    };
    activeRentRevisions = [];
    const baseRent =
        activeRentHistoryContext.rentAmount ||
        activeRentHistoryContext.templateData?.rent_amount ||
        activeRentHistoryContext.currentRent ||
        "";
    renderRentHistory(baseRent);
    loadRentHistoryForTenancy(resolvedTenancy?.tenancyId, baseRent);
    showModal(modal);
}

function closeRentHistoryModal() {
    const modal = document.getElementById("rentHistoryModal");
    if (modal) hideModal(modal);
    activeRentHistoryContext = null;
    activeRentRevisions = [];
}

function handleRentHistoryClick(event, fallbackTenant) {
    const btn = event?.target?.closest?.(".rent-history-btn");
    if (!btn) return;

    event.preventDefault();
    event.stopPropagation();

    const { tenancy, tenant } = buildRentHistoryContext(btn, fallbackTenant);
    const contextTenant = tenant || selectedTenantForSidebar;
    const resolvedTenancy =
        (tenancy && tenancy.tenancyId && tenancy) || resolveFallbackTenancyContext(contextTenant) || tenancy || null;

    if (!resolvedTenancy || !resolvedTenancy.tenancyId) {
        showToast("Select a tenancy first", "warning");
        return;
    }

    openRentHistoryModal(resolvedTenancy, contextTenant);
}

function startNewTenancyFromSidebar() {
    if (!selectedTenantForSidebar) {
        showToast("Select a tenant first", "warning");
        return;
    }
    pendingNewTenancyTenant = selectedTenantForSidebar;
    const msg = document.getElementById("newTenancyConfirmMessage");
    if (msg) {
        msg.textContent =
            "Move tenant to a new unit? Click OK to move (previous tenancy ends). Click Cancel to keep previous tenancy active and add another unit.";
    }
    openNewTenancyConfirmModal();
}

function openNewTenancyConfirmModal() {
    const modal = document.getElementById("newTenancyConfirmModal");
    if (modal) showModal(modal);
}

function closeNewTenancyConfirmModal() {
    const modal = document.getElementById("newTenancyConfirmModal");
    if (modal) hideModal(modal);
    pendingNewTenancyTenant = null;
}

function handleNewTenancyChoice(moveChoice) {
    const base = pendingNewTenancyTenant || selectedTenantForSidebar;
    if (!base) {
        closeNewTenancyConfirmModal();
        return;
    }

    const keepPreviousActive = !moveChoice;
    const newTenancyId = typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : `tenancy-${Date.now()}`;
    const templateData = { ...(base.templateData || {}), tenancy_id: newTenancyId };

    const draftTenancy = {
        ...base,
        tenancyId: newTenancyId,
        previousTenancyId: moveChoice ? base.tenancyId : undefined,
        isNewTenancy: true,
        keepPreviousActive,
        unitId: "",
        wing: "",
        floor: "",
        direction: "",
        meterNumber: "",
        tenancyEndDate: "",
        activeTenant: true,
        templateData,
    };

    closeNewTenancyConfirmModal();
    populateTenantModal(draftTenancy, "tenancy");
    setTenantModalEditable(true);
}

function populateTenantModal(tenant, mode = "tenant") {
    activeTenantForModal = tenant;
    activeRentRevisions = [];
    const modal = document.getElementById("tenantDetailModal");
    if (!modal) return;

    toggleTenantModalSections(mode);

    const templateData = tenant.templateData || {};

    syncTenantModalPicklists();
    populateUnitDropdown(tenant.unitId || templateData.unit_id, tenant.tenancyId || templateData.tenancy_id);

    const title = document.getElementById("tenantModalTitle");
    if (title) {
        const statusLabel = tenant.activeTenant ? "Active" : "Inactive";
        const unitLabel = tenant.unitNumber || templateData.unit_number || templateData.unitNumber;
        if (mode === "tenancy") {
            if (unitLabel) {
                title.textContent = `Edit Tenancy — ${unitLabel} (${statusLabel})`;
            } else {
                title.textContent = `New Tenancy — ${tenant.tenantFullName || "Tenant"}`;
            }
        } else {
            title.textContent = tenant.tenantFullName || "Tenant";
        }
    }

    setTenantModalStatusPill(tenant.activeTenant);

    const fields = {
        tenantModalFullName: tenant.tenantFullName || templateData.Tenant_Full_Name || "",
        tenantModalOccupation: tenant.tenantOccupation || templateData.Tenant_occupation || "",
        tenantModalAddress: tenant.tenantPermanentAddress || templateData.Tenant_Permanent_Address || "",
        tenantModalAadhaar: tenant.tenantAadhaar || templateData.tenant_Aadhar || "",
        tenantModalGrn: tenant.grnNumber || templateData["GRN number"] || "",
        tenantModalUnit: tenant.unitId || templateData.unit_id || "",
        tenantModalWing: tenant.wing || templateData.wing || "",
        tenantModalUnitNumber: tenant.unitNumber || templateData.unit_number || templateData.unitNumber || "",
        tenantModalFloor: tenant.floor || templateData["floor_of_building "] || templateData.floor_of_building || "",
        tenantModalDirection: tenant.direction || templateData.direction_build || "",
        tenantModalMeter: tenant.meterNumber || templateData.meter_number || "",
        tenantModalPayable: normalizeDayValue(
            tenant.payableDate || templateData.payable_date_raw || templateData.payable_date || ""
        ),
        tenantModalDeposit: tenant.securityDeposit || templateData.secu_depo || "",
        tenantModalTenantNotice: tenant.tenantNoticeMonths || templateData.notice_num_t || "",
        tenantModalLandlordNotice: tenant.landlordNoticeMonths || templateData.notice_num_l || "",
        tenantModalLateRent: tenant.lateRentPerDay || templateData.late_rent || "",
        tenantModalGracePeriod: tenant.lateGracePeriodDays || templateData.late_days || "",
        tenantModalLandlord: tenant.landlordId || templateData.landlord_id || "",
        tenantModalLandlordAadhaar: tenant.landlordAadhaar || templateData.landlord_aadhar || "",
        tenantModalLandlordAddress: tenant.landlordAddress || templateData.landlord_address || "",
        tenantModalAgreementDate: formatDateForInput(
            tenant.agreementDateRaw || templateData.agreement_date_raw || ""
        ),
        tenantModalCommencement: formatDateForInput(
            tenant.tenancyCommencementRaw || templateData.tenancy_comm_raw || ""
        ),
        tenantModalEndDate: formatDateForInput(
            tenant.tenancyEndRaw || tenant.tenancyEndDate || templateData.tenancy_end_raw || ""
        ),
        tenantModalMobile: tenant.tenantMobile || templateData.tenant_mobile || "",
        tenantModalVacateReason: tenant.vacateReason || "",
        tenantModalInitialMeter: "",
        tenantModalRentRevisionUnit: tenant.rentRevisionUnit || templateData["rent_rev year_mon"] || "",
        tenantModalRentRevisionNumber: tenant.rentRevisionNumber || templateData.rent_rev_number || "",
        tenantModalPetPolicy: tenant.petPolicy || templateData.pet_text_area || "",
    };

    Object.entries(fields).forEach(([id, value]) => {
        const el = document.getElementById(id);
        if (el) el.value = value;
    });

    if (fields.tenantModalLandlord) {
        applyLandlordSelectionToModal(fields.tenantModalLandlord);
    }

    const familyTbody = document.getElementById("tenantFamilyTableBody");
    if (familyTbody) {
        familyTbody.innerHTML = "";
        (tenant.family || []).forEach((member) => {
            familyTbody.appendChild(createFamilyRow(member));
        });
    }

    applyUnitSelectionToModal(tenant.unitId || templateData.unit_id || "");

    if (mode === "rent") {
        renderRentHistory(tenant.rentAmount || templateData.rent_amount || "");
        loadRentHistoryForTenancy(tenant.tenancyId || templateData.tenancy_id, tenant.rentAmount || templateData.rent_amount || "");
    }

    setTenantModalEditable(false);
    showModal(modal);
}

function collectFamilyRows() {
    const tbody = document.getElementById("tenantFamilyTableBody");
    if (!tbody) return [];

    const members = [];
    tbody.querySelectorAll("tr").forEach((tr) => {
        const inputs = tr.querySelectorAll("input");
        if (inputs.length < 5) return;
        const [name, relationship, occupation, aadhaar, address] = inputs;
        if (!name.value && !relationship.value && !occupation.value && !aadhaar.value && !address.value) return;

        members.push({
            name: name.value || "",
            relationship: relationship.value || "",
            occupation: occupation.value || "",
            aadhaar: aadhaar.value || "",
            address: address.value || "",
        });
    });

    return members;
}

async function saveTenantModal() {
    if (!activeTenantForModal) return;

    if (!tenantModalEditable) {
        showToast("Click Edit details to enable updates", "warning");
        return;
    }

    const payableSelect = document.getElementById("tenantModalPayable");

    const updates = {
        tenantFullName: document.getElementById("tenantModalFullName")?.value || "",
        tenantOccupation: document.getElementById("tenantModalOccupation")?.value || "",
        tenantPermanentAddress: document.getElementById("tenantModalAddress")?.value || "",
        tenantAadhaar: document.getElementById("tenantModalAadhaar")?.value || "",
        grnNumber: document.getElementById("tenantModalGrn")?.value.trim() || "",
        unitId: document.getElementById("tenantModalUnit")?.value || "",
        wing: document.getElementById("tenantModalWing")?.value.trim() || "",
        floor: document.getElementById("tenantModalFloor")?.value.trim() || "",
        direction: document.getElementById("tenantModalDirection")?.value || "",
        meterNumber: document.getElementById("tenantModalMeter")?.value || "",
        payableDate:
            payableSelect && payableSelect.value
                ? toOrdinal(parseInt(payableSelect.value, 10))
                : payableSelect?.value || "",
        securityDeposit: document.getElementById("tenantModalDeposit")?.value || "",
        tenantNoticeMonths: document.getElementById("tenantModalTenantNotice")?.value || "",
        landlordNoticeMonths: document.getElementById("tenantModalLandlordNotice")?.value || "",
        lateRentPerDay: document.getElementById("tenantModalLateRent")?.value || "",
        lateGracePeriodDays: document.getElementById("tenantModalGracePeriod")?.value || "",
        landlordId: document.getElementById("tenantModalLandlord")?.value || "",
        agreementDateRaw: normalizeDateInputValue(
            document.getElementById("tenantModalAgreementDate")?.value || ""
        ),
        tenancyCommencementRaw: normalizeDateInputValue(
            document.getElementById("tenantModalCommencement")?.value || ""
        ),
        tenancyEndRaw: normalizeDateInputValue(document.getElementById("tenantModalEndDate")?.value || ""),
        tenantMobile: document.getElementById("tenantModalMobile")?.value || "",
        vacateReason: document.getElementById("tenantModalVacateReason")?.value || "",
        unitNumber: document.getElementById("tenantModalUnitNumber")?.value || "",
        rentRevisionUnit: document.getElementById("tenantModalRentRevisionUnit")?.value || "",
        rentRevisionNumber: document.getElementById("tenantModalRentRevisionNumber")?.value || "",
        petPolicy: document.getElementById("tenantModalPetPolicy")?.value || "",
    };
    const existingGrn = activeTenantForModal.grnNumber || activeTenantForModal.templateData?.["GRN number"] || "";
    const submittedGrn = updates.grnNumber || existingGrn;

    const familyMembers = collectFamilyRows();

    try {
        const res = await updateTenantRecord({
            tenantId: activeTenantForModal.tenantId,
            tenancyId: activeTenantForModal.tenancyId,
            previousTenancyId: activeTenantForModal.previousTenancyId,
            createNewTenancy: !!activeTenantForModal.isNewTenancy,
            forceNewTenancyId: activeTenantForModal.tenancyId,
            keepPreviousActive: !!activeTenantForModal.keepPreviousActive,
            unitId: updates.unitId || activeTenantForModal.unitId,
            grn: submittedGrn,
            templateData: activeTenantForModal.templateData,
            updates,
            familyMembers,
        });

        const unit = unitCache.find((u) => u.unit_id === (updates.unitId || activeTenantForModal.unitId));
        const landlord = landlordCache.find((l) => l.landlord_id === updates.landlordId) || {};
        const merged = {
            ...activeTenantForModal,
            ...updates,
            landlordId: updates.landlordId || activeTenantForModal.landlordId,
            landlordName: landlord.name || activeTenantForModal.landlordName,
            landlordAadhaar: landlord.aadhaar || activeTenantForModal.landlordAadhaar,
            landlordAddress: landlord.address || activeTenantForModal.landlordAddress,
            grnNumber: submittedGrn,
            activeTenant:
                activeTenantForModal.isNewTenancy || typeof updates.activeTenant !== "undefined"
                    ? updates.activeTenant ?? true
                    : activeTenantForModal.activeTenant,
            tenancyEndDate: updates.tenancyEndRaw || activeTenantForModal.tenancyEndDate,
            unitId: updates.unitId || activeTenantForModal.unitId,
            family: familyMembers,
            isNewTenancy: false,
            previousTenancyId: undefined,
        };

        const historyEntry = {
            tenancyId: merged.tenancyId,
            unitLabel: buildUnitLabel(unit),
            startDate: updates.tenancyCommencementRaw || merged.tenancyCommencement,
            endDate: merged.tenancyEndDate,
            status: merged.activeTenant ? "ACTIVE" : "ENDED",
            grnNumber: merged.grnNumber,
        };
        const existingHistory = Array.isArray(activeTenantForModal.tenancyHistory)
            ? activeTenantForModal.tenancyHistory
            : [];
        const adjustedHistory = existingHistory.map((h) => {
            if (
                activeTenantForModal.previousTenancyId &&
                !activeTenantForModal.keepPreviousActive &&
                h.tenancyId === activeTenantForModal.previousTenancyId
            ) {
                return { ...h, status: "ENDED", endDate: h.endDate || formatDateForInput(new Date()) };
            }
            return h;
        });
        merged.tenancyHistory = [historyEntry, ...adjustedHistory.filter((h) => h.tenancyId !== merged.tenancyId)];

        if (activeTenantForModal.isNewTenancy) {
            tenantCache.push(merged);
            selectedTenantForSidebar = merged;
        } else {
            Object.assign(activeTenantForModal, merged);
            const idx = tenantCache.findIndex((t) => t.tenancyId === merged.tenancyId);
            if (idx >= 0) tenantCache[idx] = merged;
        }

        activeTenantForModal = merged;

        applyTenantFilters();
        if (selectedTenantForSidebar && selectedTenantForSidebar.grnNumber === merged.grnNumber) {
            Object.assign(selectedTenantForSidebar, merged);
            updateSidebarSnapshot();
        }
        setTenantModalEditable(false);
        closeTenantModal();
    } catch (e) {
        console.error("Failed to update tenant", e);
    }
}

export function closeTenantModal() {
    const modal = document.getElementById("tenantDetailModal");
    if (modal) hideModal(modal);
    setTenantModalEditable(false);
    activeRentRevisions = [];
}

export function openTenantModal(tenant) {
    if (!tenant) {
        showToast("Tenant not found", "error");
        return;
    }
    populateTenantModal(tenant, "tenant");
}

function closeTenantInsightsModal() {
    const modal = document.getElementById("tenantInsightsModal");
    if (modal) hideModal(modal);
}

function openTenantInsightsModal(tenant) {
    const t = tenant || selectedTenantForSidebar;
    if (!t) {
        showToast("Select a tenant to view insights", "warning");
        return;
    }

    const rent = parseInt(t.rentAmount, 10) || 0;
    const familyCount = Array.isArray(t.family) ? t.family.length : 0;
    const activeMonths = t.activeTenant ? 12 : 6;
    const collected = rent * Math.max(1, Math.min(activeMonths, 12));
    const projectedNextQuarter = rent * 3;

    const setBarWidth = (id, percent) => {
        const el = document.getElementById(id);
        if (el) el.style.width = `${Math.min(100, Math.max(5, percent))}%`;
    };

    const occupancyLabel = t.activeTenant ? "Occupied" : "Inactive";
    const tenureNote = t.tenancyEndDate
        ? `Ends ${formatTenancyEndDate(t.tenancyEndDate)}`
        : "No end date set";

    const highlights = [
        `Rent set at ${formatRent(t.rentAmount)} for ${formatWingFloor(t)}.`,
        `Status: ${getStatusLabel(t.activeTenant)}${t.vacateReason ? ` (Reason: ${t.vacateReason})` : ""}.`,
        `Family linked: ${familyCount} member${familyCount === 1 ? "" : "s"}.`,
    ];
    if (t.grnNumber) highlights.push(`GRN ${t.grnNumber} recorded in directory.`);
    if (t.meterNumber) highlights.push(`Meter ${t.meterNumber} noted for utilities.`);

    const quarterPercent = rent ? Math.min(100, Math.max(10, Math.round((rent * 3) / (rent * 4) * 100))) : 25;
    const collectedPercent = rent ? Math.min(100, Math.max(15, Math.round((collected / (rent * 12)) * 100))) : 40;

    const modal = document.getElementById("tenantInsightsModal");
    if (!modal) return;

    const title = document.getElementById("tenantInsightsTitle");
    if (title) title.textContent = `Insights – ${t.tenantFullName || "Tenant"}`;
    const rentEl = document.getElementById("insightRent");
    if (rentEl) rentEl.textContent = formatRent(t.rentAmount);
    const statusEl = document.getElementById("insightStatus");
    if (statusEl) statusEl.textContent = `Status: ${getStatusLabel(t.activeTenant)}`;
    const collectedEl = document.getElementById("insightCollected");
    if (collectedEl) collectedEl.textContent = formatRent(collected);
    const occupancyEl = document.getElementById("insightOccupancy");
    if (occupancyEl) occupancyEl.textContent = occupancyLabel;
    const tenureEl = document.getElementById("insightTenure");
    if (tenureEl) tenureEl.textContent = tenureNote;
    const quarterEl = document.getElementById("insightQuarter");
    if (quarterEl) quarterEl.textContent = formatRent(rent * 3);
    const nextQuarterEl = document.getElementById("insightNextQuarter");
    if (nextQuarterEl) nextQuarterEl.textContent = formatRent(projectedNextQuarter);

    setBarWidth("insightCollectedBar", collectedPercent);
    setBarWidth("insightQuarterBar", quarterPercent);
    setBarWidth("insightNextQuarterBar", Math.min(100, quarterPercent + 10));

    const highlightsList = document.getElementById("insightHighlights");
    if (highlightsList) {
        highlightsList.innerHTML = "";
        highlights.slice(0, 5).forEach((line) => {
            const li = document.createElement("li");
            li.textContent = line;
            highlightsList.appendChild(li);
        });
    }

    showModal(modal);
}

function closeVacateModal() {
    const modal = document.getElementById("vacateModal");
    if (modal) hideModal(modal);
    pendingVacateTenant = null;
}

function openVacateModal(tenant) {
    pendingVacateTenant = tenant;
    setSidebarSelection(tenant);

    const modal = document.getElementById("vacateModal");
    const title = document.getElementById("vacateModalTitle");
    const reasonInput = document.getElementById("vacateReasonInput");
    const endDateInput = document.getElementById("vacateEndDateInput");

    if (title) title.textContent = `Vacate – ${tenant.tenantFullName || "Tenant"}`;
    if (reasonInput) reasonInput.value = tenant.vacateReason || "";
    if (endDateInput) {
        if (tenant.tenancyEndDate) {
            endDateInput.value = formatDateForInput(tenant.tenancyEndDate);
        } else {
            endDateInput.value = "";
        }
    }

    if (modal) showModal(modal);
}

async function saveVacateModal() {
    if (!pendingVacateTenant) return;

    const reasonInput = document.getElementById("vacateReasonInput");
    const endDateInput = document.getElementById("vacateEndDateInput");

    const vacateReason = reasonInput?.value?.trim() || "";
    const tenancyEndRaw = endDateInput?.value || "";

    if (!tenancyEndRaw) {
        showToast("Please add a tenancy end date", "warning");
        return;
    }

    const updates = {
        activeTenant: false,
        vacateReason,
        tenancyEndRaw,
    };

    try {
        await updateTenantRecord({
            tenantId: pendingVacateTenant.tenantId,
            tenancyId: pendingVacateTenant.tenancyId,
            grn: pendingVacateTenant.grnNumber,
            templateData: pendingVacateTenant.templateData,
            updates,
            familyMembers: pendingVacateTenant.family || [],
        });

        Object.assign(pendingVacateTenant, {
            ...updates,
            tenancyEndDate: tenancyEndRaw,
        });

        applyTenantFilters();
        if (selectedTenantForSidebar && selectedTenantForSidebar.grnNumber === pendingVacateTenant.grnNumber) {
            Object.assign(selectedTenantForSidebar, pendingVacateTenant);
            updateSidebarSnapshot();
        }
        closeVacateModal();
        showToast("Tenant marked inactive", "success");
    } catch (e) {
        console.error("Failed to mark tenant inactive", e);
    }
}

export function initTenantDirectory() {
    syncTenantModalPicklists();

    document.addEventListener("landlords:updated", (e) => {
        if (e?.detail && Array.isArray(e.detail)) {
            landlordCache = e.detail;
        }
        syncTenantModalPicklists();
    });

    const searchInput = document.getElementById("tenantSearchInput");
    if (searchInput) {
        searchInput.addEventListener("input", (e) => {
            currentSearch = (e.target.value || "").trim();
            applyTenantFilters();
        });
    }

    document.querySelectorAll("[data-tenant-status-filter]").forEach((btn) => {
        btn.addEventListener("click", () => {
            currentStatusFilter = btn.getAttribute("data-tenant-status-filter") || "all";
            syncStatusButtons();
            applyTenantFilters();
        });
    });

    const refreshBtn = document.getElementById("refreshTenantListBtn");
    if (refreshBtn) {
        refreshBtn.addEventListener("click", () => loadTenantDirectory(true));
    }

    const newTenancyBtn = document.getElementById("sidebarNewTenancyBtn");
    if (newTenancyBtn) {
        newTenancyBtn.addEventListener("click", startNewTenancyFromSidebar);
    }

    const newTenancyConfirmOkBtn = document.getElementById("newTenancyConfirmOk");
    if (newTenancyConfirmOkBtn) {
        newTenancyConfirmOkBtn.addEventListener("click", () => handleNewTenancyChoice(true));
    }

    const newTenancyConfirmCancelBtn = document.getElementById("newTenancyConfirmCancel");
    if (newTenancyConfirmCancelBtn) {
        newTenancyConfirmCancelBtn.addEventListener("click", () => handleNewTenancyChoice(false));
    }

    document.querySelectorAll(".new-tenancy-confirm-close").forEach((btn) =>
        btn.addEventListener("click", () => closeNewTenancyConfirmModal())
    );

    const editTenantBtn = document.getElementById("sidebarEditTenantBtn");
    if (editTenantBtn) {
        editTenantBtn.addEventListener("click", () => {
            if (!selectedTenantForSidebar) return;
            openTenantModal(selectedTenantForSidebar);
        });
    }

    const modalCloseButtons = document.querySelectorAll(".tenant-modal-close");
    modalCloseButtons.forEach((btn) => btn.addEventListener("click", closeTenantModal));

    document.querySelectorAll(".tenant-modal-edit-toggle").forEach((btn) =>
        btn.addEventListener("click", () => setTenantModalEditable(true))
    );

    const addFamilyBtn = document.getElementById("tenantModalAddFamilyBtn");
    if (addFamilyBtn) {
        addFamilyBtn.addEventListener("click", () => {
            const tbody = document.getElementById("tenantFamilyTableBody");
            if (tbody) tbody.appendChild(createFamilyRow());
        });
    }

    const unitSelect = document.getElementById("tenantModalUnit");
    if (unitSelect) {
        unitSelect.addEventListener("change", (e) => applyUnitSelectionToModal(e.target.value));
    }

    const landlordSelect = document.getElementById("tenantModalLandlord");
    if (landlordSelect) {
        landlordSelect.addEventListener("change", (e) => applyLandlordSelectionToModal(e.target.value));
    }

    const saveBtn = document.getElementById("tenantModalSaveBtn");
    if (saveBtn) {
        saveBtn.addEventListener("click", saveTenantModal);
    }

    const rentRevisionSaveBtn = document.getElementById("rentRevisionSaveBtn");
    if (rentRevisionSaveBtn) {
        rentRevisionSaveBtn.addEventListener("click", handleRentRevisionSave);
    }

    document.querySelectorAll(".rent-history-close").forEach((btn) => btn.addEventListener("click", closeRentHistoryModal));

    const insightsCloseBtns = document.querySelectorAll(".tenant-insights-close");
    insightsCloseBtns.forEach((btn) => btn.addEventListener("click", closeTenantInsightsModal));

    const vacateCloseBtns = document.querySelectorAll(".vacate-modal-close");
    vacateCloseBtns.forEach((btn) => btn.addEventListener("click", closeVacateModal));

    const vacateSaveBtn = document.getElementById("vacateSaveBtn");
    if (vacateSaveBtn) {
        vacateSaveBtn.addEventListener("click", saveVacateModal);
    }

    const sidebarHistory = document.getElementById("sidebarUnitHistory");
    if (sidebarHistory) {
        sidebarHistory.addEventListener("click", (event) => handleRentHistoryClick(event));
    }

    document.addEventListener("click", (event) => {
        const modal = document.getElementById("rentHistoryModal");
        if (modal && !modal.classList.contains("hidden")) return;
        handleRentHistoryClick(event);
    });

    syncStatusButtons();
}
