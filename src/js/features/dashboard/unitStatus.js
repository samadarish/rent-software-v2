/**
 * Unit Status Dashboard Widget
 * 
 * Displays all units with real-time occupancy status indicators.
 * Data is loaded from local SQLite and updates automatically on sync.
 * Hover over occupied units to see tenant names.
 */

import { getLocalList, LOCAL_KEYS } from "../../api/localStore.js";
import { escapeHtml } from "../../utils/htmlUtils.js";

const unitStatusState = {
    units: [],
    tenancies: [],
    tenants: [],
    loading: false,
    initialized: false,
    tooltipEl: null,
};

function normalizeOccupiedFlag(raw) {
    if (raw === true || raw === false) return raw;
    if (typeof raw === "string") return raw.toLowerCase() === "true";
    return !!raw;
}

function normalizeId(value) {
    return (value ?? "").toString().trim().toLowerCase();
}

function getUnitId(unit) {
    return (unit?.unit_id || unit?.unitId || "").toString().trim();
}

function getUnitNumber(unit) {
    return (unit?.unit_number || unit?.unitNumber || "").toString().trim();
}

function getWing(unit) {
    return (unit?.wing || "").toString().trim();
}

function isOccupied(unit) {
    return normalizeOccupiedFlag(unit?.is_occupied ?? unit?.isOccupied);
}

function getCurrentTenancyId(unit) {
    return (unit?.current_tenancy_id || unit?.currentTenancyId || "").toString().trim();
}

function getTenancyId(tenancy) {
    return (tenancy?.tenancy_id || tenancy?.tenancyId || "").toString().trim();
}

function getTenantIdFromTenancy(tenancy) {
    return (tenancy?.tenant_id || tenancy?.tenantId || "").toString().trim();
}

function getTenantId(tenant) {
    return (tenant?.tenant_id || tenant?.tenantId || "").toString().trim();
}

function getTenantName(tenant) {
    // Match the pattern used elsewhere in the codebase (exportData.js, tenants.js)
    const name = (
        tenant?.tenantFullName ||
        tenant?.full_name ||
        tenant?.tenantName ||
        tenant?.tenant_name ||
        tenant?.name ||
        tenant?.templateData?.Tenant_Full_Name ||
        tenant?.templateData?.tenantFullName ||
        ""
    ).toString().trim();
    return name || null;
}

/**
 * Build lookup maps for tenancy -> tenant name resolution
 */
function buildLookupMaps(tenancies, tenants) {
    // Map tenancy_id -> tenancy object
    const tenancyMap = new Map();
    tenancies.forEach((t) => {
        const id = normalizeId(getTenancyId(t));
        if (id) tenancyMap.set(id, t);
    });

    // Map tenant_id -> tenant object
    const tenantMap = new Map();
    tenants.forEach((t) => {
        const id = normalizeId(getTenantId(t));
        if (id) tenantMap.set(id, t);
    });

    return { tenancyMap, tenantMap };
}

/**
 * Get tenant name for a unit based on current_tenancy_id
 */
function getTenantNameForUnit(unit, tenancyMap, tenantMap) {
    const tenancyId = normalizeId(getCurrentTenancyId(unit));
    if (!tenancyId) return null;

    const tenancy = tenancyMap.get(tenancyId);
    if (!tenancy) return null;

    const tenantId = normalizeId(getTenantIdFromTenancy(tenancy));
    if (!tenantId) return null;

    const tenant = tenantMap.get(tenantId);
    if (!tenant) return null;

    return getTenantName(tenant);
}

function groupUnitsByWing(units) {
    const groups = new Map();
    units.forEach((unit) => {
        const wing = getWing(unit) || "Other";
        if (!groups.has(wing)) {
            groups.set(wing, []);
        }
        groups.get(wing).push(unit);
    });
    // Sort units within each wing by unit number
    groups.forEach((wingUnits) => {
        wingUnits.sort((a, b) => {
            const numA = getUnitNumber(a);
            const numB = getUnitNumber(b);
            return numA.localeCompare(numB, undefined, { numeric: true });
        });
    });
    return groups;
}

function getElements() {
    return {
        container: document.getElementById("unitStatusContainer"),
        list: document.getElementById("unitStatusList"),
        empty: document.getElementById("unitStatusEmpty"),
        occupied: document.getElementById("unitStatusOccupied"),
        vacant: document.getElementById("unitStatusVacant"),
    };
}

/**
 * Create or get the tooltip element (appended to body to avoid overflow clipping)
 */
function getTooltipEl() {
    if (unitStatusState.tooltipEl) return unitStatusState.tooltipEl;
    
    let el = document.getElementById("unit-tooltip");
    if (!el) {
        el = document.createElement("div");
        el.id = "unit-tooltip";
        document.body.appendChild(el);
    }
    unitStatusState.tooltipEl = el;
    return el;
}

function showTooltip(targetEl, text) {
    const tooltip = getTooltipEl();
    tooltip.textContent = text;
    
    const rect = targetEl.getBoundingClientRect();
    const tooltipWidth = tooltip.offsetWidth || 100;
    
    // Position above the element, centered
    tooltip.style.left = `${rect.left + rect.width / 2 - tooltipWidth / 2}px`;
    tooltip.style.top = `${rect.top - 30}px`;
    
    tooltip.classList.add("visible");
}

function hideTooltip() {
    const tooltip = getTooltipEl();
    tooltip.classList.remove("visible");
}

function renderWidget(units, tenancyMap, tenantMap) {
    const { list, empty, occupied, vacant } = getElements();
    if (!list) return;

    const occupiedCount = units.filter(isOccupied).length;
    const vacantCount = units.length - occupiedCount;

    // Update summary counts
    if (occupied) occupied.textContent = occupiedCount;
    if (vacant) vacant.textContent = vacantCount;

    // Clear list
    list.innerHTML = "";

    if (!units.length) {
        if (empty) {
            empty.textContent = "No units configured yet.";
            empty.classList.remove("hidden");
        }
        return;
    }

    if (empty) empty.classList.add("hidden");

    // Group by wing
    const grouped = groupUnitsByWing(units);

    // Render each wing group
    grouped.forEach((wingUnits, wing) => {
        // Wing header
        const wingHeader = document.createElement("div");
        wingHeader.className = "text-[9px] font-semibold text-slate-500 uppercase tracking-wide mt-2 first:mt-0 mb-1";
        wingHeader.textContent = wing;
        list.appendChild(wingHeader);

        // Unit grid for this wing - flex-wrap ensures no horizontal scroll
        const unitGrid = document.createElement("div");
        unitGrid.className = "flex flex-wrap gap-1.5 max-w-full";

        wingUnits.forEach((unit) => {
            const unitNumber = escapeHtml(getUnitNumber(unit) || getUnitId(unit) || "?");
            const occupied = isOccupied(unit);

            // Build tooltip only for occupied units
            let tooltip = null;
            if (occupied) {
                const tenantName = getTenantNameForUnit(unit, tenancyMap, tenantMap);
                tooltip = tenantName || "Occupied";
            }

            const item = document.createElement("div");
            const baseClasses = "flex items-center justify-center gap-1.5 min-w-[40px] px-2 py-1 rounded-md border text-[10px] font-medium cursor-default";
            const colorClasses = occupied 
                ? "bg-emerald-50 border-emerald-200 text-emerald-700" 
                : "bg-rose-50 border-rose-200 text-rose-700";
            item.className = `${baseClasses} ${colorClasses}`.trim().replace(/\s+/g, " ");
            
            // Add tooltip events for occupied units
            if (tooltip) {
                item.addEventListener("mouseenter", () => showTooltip(item, tooltip));
                item.addEventListener("mouseleave", hideTooltip);
            }

            const dot = document.createElement("span");
            dot.className = `w-2 h-2 rounded-full ${occupied ? "bg-emerald-500" : "bg-rose-500"}`;

            const label = document.createElement("span");
            label.textContent = unitNumber;

            item.appendChild(dot);
            item.appendChild(label);
            unitGrid.appendChild(item);
        });

        list.appendChild(unitGrid);
    });
}

function setLoadingState() {
    const { list, empty } = getElements();
    if (list) list.innerHTML = "";
    if (empty) {
        empty.textContent = "Loading units...";
        empty.classList.remove("hidden");
    }
}

async function loadUnits() {
    if (unitStatusState.loading) return;
    unitStatusState.loading = true;
    setLoadingState();

    try {
        // Load units, tenancies, and tenants in parallel
        const [units, tenancies, tenants] = await Promise.all([
            getLocalList(LOCAL_KEYS.units, []),
            getLocalList(LOCAL_KEYS.tenancies, []),
            getLocalList(LOCAL_KEYS.tenants, []),
        ]);
        
        unitStatusState.units = Array.isArray(units) ? units : [];
        unitStatusState.tenancies = Array.isArray(tenancies) ? tenancies : [];
        unitStatusState.tenants = Array.isArray(tenants) ? tenants : [];
    } catch (err) {
        console.error("Failed to load units for status widget", err);
        unitStatusState.units = [];
        unitStatusState.tenancies = [];
        unitStatusState.tenants = [];
    }

    const { tenancyMap, tenantMap } = buildLookupMaps(
        unitStatusState.tenancies,
        unitStatusState.tenants
    );
    renderWidget(unitStatusState.units, tenancyMap, tenantMap);
    unitStatusState.loading = false;
}

export function refreshUnitStatus() {
    loadUnits();
}

export function initUnitStatus() {
    if (unitStatusState.initialized) return;
    unitStatusState.initialized = true;

    const { container } = getElements();
    if (!container) return;

    // When units change, reload all data from SQLite to get fresh tenancies/tenants
    // This ensures tenant names are available after saving a new tenant
    document.addEventListener("units:updated", () => {
        loadUnits();
    });

    // When tenants change, reload all data to update tooltips
    document.addEventListener("tenants:updated", () => {
        loadUnits();
    });

    // When tenancies change, reload all data to update tooltips
    document.addEventListener("tenancies:updated", () => {
        loadUnits();
    });

    // Refresh on sync completion
    document.addEventListener("sync:completed", () => {
        loadUnits();
    });

    // Initial load
    loadUnits();
}


