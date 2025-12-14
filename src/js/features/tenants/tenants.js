import {
    fetchLandlordsFromSheet,
    fetchTenantDirectory,
    fetchUnitsFromSheet,
    updateTenantRecord,
} from "../../api/sheets.js";
import { getCachedGetResponse } from "../../api/appscriptClient.js";
import { getAppScriptUrl } from "../../api/config.js";
import { toOrdinal } from "../../utils/formatters.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

let tenantCache = [];
let hasLoadedTenants = false;
let unitCache = [];
let landlordCache = [];
let currentStatusFilter = "active"; // all | active | inactive
let currentSearch = "";
let activeTenantForModal = null;
let selectedTenantForSidebar = null;
let pendingVacateTenant = null;
let tenantModalEditable = false;

const statusClassMap = {
    active: "bg-emerald-100 text-emerald-700 border-emerald-200",
    inactive: "bg-rose-100 text-rose-700 border-rose-200",
};

function hydrateTenantDirectoryFromCache(previousSelectedGrn) {
    const url = getAppScriptUrl();
    if (!url) return false;

    const cachedUnits = getCachedGetResponse({ url, action: "units" });
    const cachedLandlords = getCachedGetResponse({ url, action: "landlords" });
    const cachedTenants = getCachedGetResponse({ url, action: "tenants" });

    if (
        !Array.isArray(cachedUnits?.units) &&
        !Array.isArray(cachedLandlords?.landlords) &&
        !Array.isArray(cachedTenants?.tenants)
    ) {
        return false;
    }

    unitCache = Array.isArray(cachedUnits?.units)
        ? cachedUnits.units.map((u) => ({
              ...u,
              is_occupied:
                  u.is_occupied === true ||
                  (typeof u.is_occupied === "string" && u.is_occupied.toLowerCase() === "true"),
          }))
        : unitCache;

    landlordCache = Array.isArray(cachedLandlords?.landlords)
        ? cachedLandlords.landlords
        : landlordCache;

    tenantCache = Array.isArray(cachedTenants?.tenants)
        ? cachedTenants.tenants.map((t) => ({
              ...t,
              tenancyHistory: Array.isArray(t.tenancyHistory) ? t.tenancyHistory : [],
          }))
        : tenantCache;

    if (landlordCache.length) {
        document.dispatchEvent(new CustomEvent("landlords:updated", { detail: landlordCache }));
    }

    applyTenantFilters();

    if (tenantCache.length) {
        const match = tenantCache.find((t) => t.grnNumber === previousSelectedGrn);
        setSidebarSelection(match || tenantCache[0]);
    } else {
        selectedTenantForSidebar = null;
        updateSidebarSnapshot();
    }

    return true;
}

export function getActiveTenantsForWing(wing) {
    const normalizedWing = (wing || "").toString().trim().toLowerCase();
    if (!normalizedWing) return [];

    const seen = new Set();
    return tenantCache.filter((t) => {
        const matchesWing = (t.wing || "").toString().trim().toLowerCase() === normalizedWing;
        const isActive = !!t.activeTenant;
        if (!matchesWing || !isActive) return false;

        const key = getTenantIdentityKey(t);
        if (!key) return true;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

export function getTenantDirectorySnapshot() {
    return [...tenantCache];
}

function formatRent(amount) {
    if (!amount) return "-";
    const numeric = parseInt(amount, 10);
    if (isNaN(numeric)) return `₹${amount}`;
    return `₹${numeric.toLocaleString("en-IN")}`;
}

function formatWingFloor(tenant) {
    const wing = tenant.wing || "-";
    const floor = tenant.floor || "-";
    return `${wing} / ${floor}`;
}

function formatDateForInput(raw) {
    if (!raw) return "";
    const d = new Date(raw);
    if (isNaN(d)) return "";
    const year = d.getFullYear();
    const month = `${d.getMonth() + 1}`.padStart(2, "0");
    const day = `${d.getDate()}`.padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function normalizeDateInputValue(raw) {
    const formatted = formatDateForInput(raw);
    return formatted || "";
}

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

function formatTenancyEndDate(raw) {
    if (!raw) return "-";
    const formatted = formatDateForInput(raw);
    return formatted || raw;
}

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

function syncTenantModalPicklists() {
    cloneSelectOptions("wing", "tenantModalWing");
    cloneSelectOptions("payable_date", "tenantModalPayable");
    cloneSelectOptions("rent_rev_number", "tenantModalRentRevNumber");
    cloneSelectOptions("rent_rev_year_mon", "tenantModalRentRevUnit");
    cloneSelectOptions("notice_num_t", "tenantModalTenantNotice");
    cloneSelectOptions("notice_num_l", "tenantModalLandlordNotice");
    cloneSelectOptions("landlord_selector", "tenantModalLandlord");
}

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

function buildUnitLabel(unit) {
    if (!unit) return "";
    const parts = [unit.wing, unit.unit_number].filter(Boolean);
    return parts.join(" - ") || unit.unit_id || "";
}

function markUnitOccupancy(unitId, occupied, tenancyId) {
    if (!unitId) return;
    const target = unitCache.find((u) => u.unit_id === unitId);
    if (target) {
        target.is_occupied = occupied;
        target.current_tenancy_id = occupied ? tenancyId : "";
    }
}

function normalizeDayValue(val) {
    if (!val) return "";
    const match = String(val).match(/\d+/);
    return match ? match[0] : String(val);
}

function setTenantModalStatusPill(active) {
    const pill = document.getElementById("tenantModalStatusPill");
    if (!pill) return;
    pill.textContent = active ? "Active" : "Inactive";
    pill.className = `text-[10px] px-2 py-1 rounded-full border font-semibold ${
        active ? statusClassMap.active : statusClassMap.inactive
    }`;
}

function setTenantListLoading(isLoading) {
    const table = document.getElementById("tenantTableBody");
    const emptyState = document.getElementById("tenantListEmpty");
    const loader = document.getElementById("tenantListLoader");

    if (loader) loader.classList.toggle("hidden", !isLoading);
    if (table) table.classList.toggle("opacity-50", isLoading);
    if (emptyState && isLoading) emptyState.classList.add("hidden");
}

function getStatusLabel(active) {
    return active ? "Active" : "Inactive";
}

function renderStatusPill(active) {
    const span = document.createElement("span");
    span.className = `text-[10px] px-2 py-1 rounded-full border font-semibold ${active ? statusClassMap.active : statusClassMap.inactive}`;
    span.textContent = getStatusLabel(active);
    return span;
}

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

        const statusCell = renderStatusPill(t.activeTenant);

        const isActive = !!t.activeTenant;

        tr.innerHTML = `
            <td class="px-2 py-1.5">
                <div class="font-semibold text-xs leading-tight">${t.tenantFullName || "Unnamed"}</div>
            </td>
            <td class="px-2 py-1.5 text-xs">${formatWingFloor(t)}</td>
            <td class="px-2 py-1.5 text-xs">${formatRent(t.rentAmount)}</td>
            <td class="px-2 py-1.5 text-xs">
                <div class="flex items-center gap-1 status-slot"></div>
            </td>
            <td class="px-2 py-1.5 text-xs">
                <button class="text-indigo-600 hover:text-indigo-800 font-semibold text-[11px] underline">View</button>
            </td>
            <td class="px-2 py-1.5 text-xs">
                <button class="inline-flex items-center px-2.5 py-1 rounded font-semibold text-[11px] ${
                    isActive
                        ? "bg-rose-600 text-white border border-rose-700 hover:bg-rose-700"
                        : "bg-slate-200 text-slate-600 border border-slate-300 cursor-not-allowed"
                }">${isActive ? "Vacate" : "Vacated"}</button>
            </td>
        `;

        const statusSlot = tr.querySelector(".status-slot");
        if (statusSlot) statusSlot.appendChild(statusCell);

        const actionBtn = tr.querySelector("button");
        if (actionBtn) {
            actionBtn.addEventListener("click", () => {
                setSidebarSelection(t);
                openTenantModal(t);
            });
        }

        const vacateBtn = tr.querySelectorAll("button")[1];
        if (vacateBtn) {
            if (!isActive) {
                vacateBtn.disabled = true;
            }
            vacateBtn.addEventListener("click", (e) => {
                e.stopPropagation();
                if (!isActive) return;
                setSidebarSelection(t);
                openVacateModal(t);
            });
        }

        tr.addEventListener("click", (e) => {
            if (e.target instanceof HTMLButtonElement) return;
            setSidebarSelection(t);
        });

        tbody.appendChild(tr);
    });

    highlightSelectedRow();
}

function applyTenantFilters() {
    let filtered = [...tenantCache];

    if (currentStatusFilter !== "all") {
        const shouldBeActive = currentStatusFilter === "active";
        filtered = filtered.filter((t) => !!t.activeTenant === shouldBeActive);
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

export async function ensureTenantDirectoryLoaded() {
    if (hasLoadedTenants) return;
    await loadTenantDirectory();
}

export async function loadTenantDirectory(forceReload = false) {
    const previousSelectedGrn = selectedTenantForSidebar?.grnNumber;
    hydrateTenantDirectoryFromCache(previousSelectedGrn);

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
    tenantCache = Array.isArray(data.tenants)
        ? data.tenants.map((t) => ({
              ...t,
              tenancyHistory: Array.isArray(t.tenancyHistory) ? t.tenancyHistory : [],
          }))
        : [];
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
    const details = document.getElementById("sidebarSnapshotDetails");
    const nameEl = document.getElementById("sidebarTenantName");
    const grnEl = document.getElementById("sidebarTenantGRN");
    const statusEl = document.getElementById("sidebarStatusPill");
    const wingFloorEl = document.getElementById("sidebarTenantWingFloor");
    const rentEl = document.getElementById("sidebarTenantRent");
    const mobileEl = document.getElementById("sidebarTenantMobile");
    const meterEl = document.getElementById("sidebarTenantMeter");
    const endEl = document.getElementById("sidebarTenantEnd");
    const insightsBtn = document.getElementById("tenantInsightsBtn");
    const newTenancyBtn = document.getElementById("sidebarNewTenancyBtn");
    const familyCount = document.getElementById("sidebarFamilyCount");
    const familyList = document.getElementById("sidebarFamilyList");
    const moveList = document.getElementById("sidebarTenantMoves");

    if (!selectedTenantForSidebar) {
        if (emptyState) emptyState.classList.remove("hidden");
        if (details) details.classList.add("hidden");
        if (nameEl) nameEl.textContent = "Select a tenant";
        if (grnEl) grnEl.textContent = "GRN: -";
        if (statusEl) {
            statusEl.textContent = "Status";
            statusEl.className = "text-[10px] px-2 py-1 rounded-full border bg-slate-50 text-slate-600";
        }
        if (wingFloorEl) wingFloorEl.textContent = "-";
        if (rentEl) rentEl.textContent = "-";
        if (mobileEl) mobileEl.textContent = "-";
        if (meterEl) meterEl.textContent = "-";
        if (endEl) endEl.textContent = "-";
        if (insightsBtn) insightsBtn.disabled = true;
        if (newTenancyBtn) newTenancyBtn.disabled = true;
        if (familyCount) familyCount.textContent = "0 members";
        if (familyList) familyList.innerHTML = '<li class="text-[10px] text-slate-500">No tenant selected.</li>';
        if (moveList) moveList.innerHTML = '<li class="text-[10px] text-slate-500">No tenant selected.</li>';
        return;
    }

    const t = selectedTenantForSidebar;
    if (emptyState) emptyState.classList.add("hidden");
    if (details) details.classList.remove("hidden");
    if (nameEl) nameEl.textContent = t.tenantFullName || "Tenant";
    if (grnEl) grnEl.textContent = `GRN: ${t.grnNumber || "-"}`;
    if (statusEl) {
        statusEl.textContent = getStatusLabel(t.activeTenant);
        statusEl.className = `text-[10px] px-2 py-1 rounded-full border font-semibold ${t.activeTenant ? statusClassMap.active : statusClassMap.inactive}`;
    }
    if (wingFloorEl) wingFloorEl.textContent = formatWingFloor(t);
    if (rentEl) rentEl.textContent = formatRent(t.rentAmount);
    if (mobileEl) mobileEl.textContent = t.tenantMobile || "-";
    if (meterEl) meterEl.textContent = t.meterNumber || "-";
    if (endEl) endEl.textContent = formatTenancyEndDate(t.tenancyEndDate);

    if (insightsBtn) {
        insightsBtn.disabled = false;
        insightsBtn.onclick = () => openTenantInsightsModal(t);
    }

    if (newTenancyBtn) {
        newTenancyBtn.disabled = false;
    }

    if (familyList && Array.isArray(t.family)) {
        familyList.innerHTML = "";
        t.family.slice(0, 6).forEach((member) => {
            const li = document.createElement("li");
            li.textContent = `${member.name || "Member"} – ${member.relationship || "Relation"}`;
            familyList.appendChild(li);
        });
        if (familyCount) familyCount.textContent = `${t.family.length} member${t.family.length === 1 ? "" : "s"}`;
    } else if (familyList) {
        familyList.innerHTML = '<li class="text-[10px] text-slate-500">No family on record.</li>';
        if (familyCount) familyCount.textContent = "0 members";
    }

    if (moveList) {
        moveList.innerHTML = "";
        const history = Array.isArray(t.tenancyHistory) ? t.tenancyHistory.slice(0, 5) : [];
        if (!history.length) {
            moveList.innerHTML = '<li class="text-[10px] text-slate-500">No tenancy history yet.</li>';
        } else {
            history.forEach((m) => {
                const li = document.createElement("li");
                const start = formatDateForInput(m.startDate) || m.startDate || "";
                const end = formatDateForInput(m.endDate) || m.endDate || "";
                const status = (m.status || "").toLowerCase();
                const endLabel = end ? ` → ${end}` : "";
                li.textContent = `${m.unitLabel || "Unit"} (${start || ""}${endLabel}) ${status ? `- ${status}` : ""}`;
                moveList.appendChild(li);
            });
        }
    }

    highlightSelectedRow();
}

function setSidebarSelection(tenant) {
    selectedTenantForSidebar = tenant;
    updateSidebarSnapshot();
}

function startNewTenancyFromSidebar() {
    if (!selectedTenantForSidebar) {
        showToast("Select a tenant first", "warning");
        return;
    }
    const base = selectedTenantForSidebar;
    const moveChoice = window.confirm(
        "Move tenant to a new unit? Click OK to move (previous tenancy ends). Click Cancel to keep previous tenancy active and add another unit."
    );
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

    populateTenantModal(draftTenancy);
    setTenantModalEditable(true);
}

function populateTenantModal(tenant) {
    activeTenantForModal = tenant;
    const modal = document.getElementById("tenantDetailModal");
    if (!modal) return;

    const templateData = tenant.templateData || {};

    syncTenantModalPicklists();
    populateUnitDropdown(tenant.unitId || templateData.unit_id, tenant.tenancyId || templateData.tenancy_id);

    const title = document.getElementById("tenantModalTitle");
    if (title) title.textContent = tenant.tenantFullName || "Tenant";

    setTenantModalStatusPill(tenant.activeTenant);

    const fields = {
        tenantModalGRN: tenant.grnNumber || templateData["GRN number"] || "",
        tenantModalFullName: tenant.tenantFullName || templateData.Tenant_Full_Name || "",
        tenantModalOccupation: tenant.tenantOccupation || templateData.Tenant_occupation || "",
        tenantModalAddress: tenant.tenantPermanentAddress || templateData.Tenant_Permanent_Address || "",
        tenantModalAadhaar: tenant.tenantAadhaar || templateData.tenant_Aadhar || "",
        tenantModalUnit: tenant.unitId || templateData.unit_id || "",
        tenantModalWing: tenant.wing || templateData.wing || "",
        tenantModalUnitNumber: tenant.unitNumber || templateData.unit_number || templateData.unitNumber || "",
        tenantModalFloor: tenant.floor || templateData["floor_of_building "] || templateData.floor_of_building || "",
        tenantModalDirection: tenant.direction || templateData.direction_build || "",
        tenantModalMeter: tenant.meterNumber || templateData.meter_number || "",
        tenantModalRent: tenant.rentAmount || templateData.rent_amount || "",
        tenantModalPayable: normalizeDayValue(
            tenant.payableDate || templateData.payable_date_raw || templateData.payable_date || ""
        ),
        tenantModalDeposit: tenant.securityDeposit || templateData.secu_depo || "",
        tenantModalRentIncrease: tenant.rentIncrease || templateData.rent_inc || "",
        tenantModalRentRevNumber: tenant.rentRevisionNumber || templateData.rent_rev_number || "",
        tenantModalRentRevUnit: tenant.rentRevisionUnit || templateData["rent_rev year_mon"] || "",
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
        tenantModalMobile: tenant.tenantMobile || templateData.tenant_mobile || "",
        tenantModalVacateReason: tenant.vacateReason || "",
        tenantModalInitialMeter: "",
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
        grnNumber: document.getElementById("tenantModalGRN")?.value || "",
        tenantFullName: document.getElementById("tenantModalFullName")?.value || "",
        tenantOccupation: document.getElementById("tenantModalOccupation")?.value || "",
        tenantPermanentAddress: document.getElementById("tenantModalAddress")?.value || "",
        tenantAadhaar: document.getElementById("tenantModalAadhaar")?.value || "",
        unitId: document.getElementById("tenantModalUnit")?.value || "",
        wing: document.getElementById("tenantModalWing")?.value.trim() || "",
        floor: document.getElementById("tenantModalFloor")?.value.trim() || "",
        direction: document.getElementById("tenantModalDirection")?.value || "",
        meterNumber: document.getElementById("tenantModalMeter")?.value || "",
        rentAmount: document.getElementById("tenantModalRent")?.value || "",
        payableDate:
            payableSelect && payableSelect.value
                ? toOrdinal(parseInt(payableSelect.value, 10))
                : payableSelect?.value || "",
        securityDeposit: document.getElementById("tenantModalDeposit")?.value || "",
        rentIncrease: document.getElementById("tenantModalRentIncrease")?.value || "",
        rentRevisionNumber: document.getElementById("tenantModalRentRevNumber")?.value || "",
        rentRevisionUnit: document.getElementById("tenantModalRentRevUnit")?.value || "",
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
        tenantMobile: document.getElementById("tenantModalMobile")?.value || "",
        vacateReason: document.getElementById("tenantModalVacateReason")?.value || "",
        unitNumber: document.getElementById("tenantModalUnitNumber")?.value || "",
    };

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
            grn: updates.grnNumber || activeTenantForModal.grnNumber,
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
            grnNumber: updates.grnNumber || activeTenantForModal.grnNumber,
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
}

export function openTenantModal(tenant) {
    if (!tenant) {
        showToast("Tenant not found", "error");
        return;
    }
    populateTenantModal(tenant);
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

    const insightsCloseBtns = document.querySelectorAll(".tenant-insights-close");
    insightsCloseBtns.forEach((btn) => btn.addEventListener("click", closeTenantInsightsModal));

    const vacateCloseBtns = document.querySelectorAll(".vacate-modal-close");
    vacateCloseBtns.forEach((btn) => btn.addEventListener("click", closeVacateModal));

    const vacateSaveBtn = document.getElementById("vacateSaveBtn");
    if (vacateSaveBtn) {
        vacateSaveBtn.addEventListener("click", saveVacateModal);
    }

    syncStatusButtons();
}
