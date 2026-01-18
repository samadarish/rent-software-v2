import { getLocalList, LOCAL_KEYS } from "../../api/localStore.js";
import { buildUnitLabel, formatDateForDoc, normalizeMonthKey, toOrdinal } from "../../utils/formatters.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

let initialized = false;
let elements = null;
const MAX_TENANT_RESULTS = 8;
let tenantSearchItems = [];
let tenantSearchResults = [];
let tenantSearchMap = new Map();
let selectedTenantEntry = null;

const exportState = {
    tenants: [],
    tenancies: [],
    units: [],
    directory: [],
    tenanciesByTenant: new Map(),
    unitMap: new Map(),
};

const selection = {
    tenantId: "",
    tenancyId: "",
};

let lastPdfDownloadUrl = "";
let lastPdfFileName = "";
let lastPdfFilePath = "";
let lastPdfBlob = null;
let pdfModalWired = false;

function getExportElements() {
    return {
        tenantInput: document.getElementById("exportTenantSearch"),
        tenantDropdown: document.getElementById("exportTenantDropdown"),
        tenantList: document.getElementById("exportTenantList"),
        tenantEmpty: document.getElementById("exportTenantEmpty"),
        tenantClear: document.getElementById("exportTenantClear"),
        tenancySelect: document.getElementById("exportTenancySelect"),
        exportBtn: document.getElementById("exportPdfBtn"),
        spinner: document.getElementById("exportPdfSpinner"),
        hint: document.getElementById("exportTenantHint"),
        empty: document.getElementById("exportDataEmpty"),
    };
}

function normalizeId(value) {
    if (value === undefined || value === null) return "";
    return String(value).trim();
}

function normalizeKey(value) {
    return normalizeId(value).toLowerCase();
}

function normalizeSearchValue(value) {
    return normalizeId(value).toLowerCase();
}

function getTemplateValue(record, key) {
    if (!record || typeof record !== "object") return undefined;
    const template = record.templateData;
    if (!template || typeof template !== "object") return undefined;
    return template[key];
}

function resolveTenantId(tenant) {
    if (!tenant || typeof tenant !== "object") return "";
    return (
        tenant.tenantId ||
        tenant.tenant_id ||
        tenant.templateData?.tenant_id ||
        ""
    );
}

function resolveTenancyId(tenancy) {
    if (!tenancy || typeof tenancy !== "object") return "";
    return (
        tenancy.tenancyId ||
        tenancy.tenancy_id ||
        tenancy.templateData?.tenancy_id ||
        ""
    );
}

function isTenancyActive(tenancy) {
    const status = (tenancy?.status || "").toString().trim().toLowerCase();
    if (status) return status === "active";
    return !normalizeId(tenancy?.end_date || tenancy?.endDate);
}

function resolveTenantStatus(tenant, tenancies) {
    if (Array.isArray(tenancies) && tenancies.some((t) => isTenancyActive(t))) {
        return "active";
    }
    if (tenant?.activeTenant === true) return "active";
    return "inactive";
}

function getTenantName(tenant) {
    return (
        tenant?.tenantFullName ||
        tenant?.tenantName ||
        tenant?.tenant_name ||
        tenant?.templateData?.Tenant_Full_Name ||
        "Tenant"
    );
}

function buildTenantLabel(tenant) {
    return getTenantName(tenant);
}

function getTenantGrn(tenant) {
    return (
        tenant?.grnNumber ||
        tenant?.grn_number ||
        getTemplateValue(tenant, "GRN number") ||
        ""
    );
}

function buildTenantSearchEntry(tenant) {
    const tenantId = normalizeId(tenant?.tenantId || tenant?.tenant_id);
    const name = getTenantName(tenant);
    const grn = normalizeId(getTenantGrn(tenant));
    const unitLabel = normalizeId(
        tenant?.unitLabel || tenant?.unitNumber || tenant?.unit_number || ""
    );
    const statusKey = normalizeId(tenant?.status).toLowerCase();
    const statusLabel = statusKey ? (statusKey === "active" ? "Active" : "Inactive") : "";
    const searchValue = normalizeSearchValue(
        [name, grn, unitLabel, statusLabel].filter(Boolean).join(" ")
    );
    return {
        tenantId,
        tenant,
        name,
        searchValue,
    };
}

function buildTenantDirectory() {
    const tenantMap = new Map();
    exportState.tenants.forEach((tenant) => {
        const tenantId = resolveTenantId(tenant);
        if (!tenantId) return;
        const existing = tenantMap.get(tenantId) || {};
        tenantMap.set(tenantId, { ...existing, ...tenant, tenantId });
    });

    const tenanciesByTenant = new Map();
    exportState.tenancies.forEach((tenancy) => {
        const tenantId = normalizeId(tenancy?.tenant_id || tenancy?.tenantId);
        if (!tenantId) return;
        const list = tenanciesByTenant.get(tenantId) || [];
        list.push(tenancy);
        tenanciesByTenant.set(tenantId, list);
    });

    const unitMap = new Map();
    exportState.units.forEach((unit) => {
        const unitId = normalizeId(unit?.unit_id || unit?.unitId);
        if (unitId) unitMap.set(unitId, unit);
    });

    const directory = Array.from(tenantMap.values()).map((tenant) => {
        const tenantId = normalizeId(tenant.tenantId);
        const tenantTenancies = tenanciesByTenant.get(tenantId) || [];
        const status = resolveTenantStatus(tenant, tenantTenancies);
        const activeTenancy = tenantTenancies.find((t) => isTenancyActive(t)) || null;
        const unit =
            unitMap.get(normalizeId(activeTenancy?.unit_id || activeTenancy?.unitId)) || null;
        const unitLabel = buildUnitLabel(unit) || tenant.unitNumber || tenant.unit_number || "";
        return {
            ...tenant,
            tenantId,
            status,
            unitLabel,
            tenancies: tenantTenancies,
        };
    });

    exportState.directory = directory;
    exportState.tenanciesByTenant = tenanciesByTenant;
    exportState.unitMap = unitMap;
}

function setExportHint(message) {
    if (elements?.hint) {
        elements.hint.textContent = message;
    }
}

function setExportButtonState(enabled) {
    if (!elements?.exportBtn) return;
    elements.exportBtn.disabled = !enabled;
    if (!enabled) {
        elements.exportBtn.classList.add("opacity-50", "cursor-not-allowed");
    } else {
        elements.exportBtn.classList.remove("opacity-50", "cursor-not-allowed");
    }
}

function clearTenantSelection({ keepInput = false } = {}) {
    selection.tenantId = "";
    selection.tenancyId = "";
    selectedTenantEntry = null;
    if (elements?.tenantInput && !keepInput) elements.tenantInput.value = "";
    if (elements?.tenancySelect) {
        elements.tenancySelect.innerHTML = '<option value="">Select a tenant first</option>';
        elements.tenancySelect.disabled = true;
    }
    setExportButtonState(false);
    setExportHint("Select a tenant and tenancy to enable export.");
    syncTenantClearButton();
}

function updateTenantOptions() {
    tenantSearchItems = [];
    tenantSearchMap = new Map();

    const filtered = exportState.directory
        .slice()
        .sort((a, b) => {
            const nameA = (a.tenantFullName || a.tenantName || "").toString();
            const nameB = (b.tenantFullName || b.tenantName || "").toString();
            return nameA.localeCompare(nameB);
        });

    filtered.forEach((tenant) => {
        const entry = buildTenantSearchEntry(tenant);
        if (!entry.tenantId) return;
        tenantSearchItems.push(entry);
        tenantSearchMap.set(entry.tenantId, entry);
    });

    const current = selection.tenantId;
    selectedTenantEntry = current ? tenantSearchMap.get(normalizeId(current)) || null : null;
    if (current && !selectedTenantEntry) {
        clearTenantSelection();
    }

    if (elements.empty) {
        const hasAnyTenants = exportState.directory.length > 0;
        elements.empty.classList.toggle("hidden", hasAnyTenants);
    }

    if (elements?.tenantInput) {
        const query = normalizeId(elements.tenantInput.value);
        const shouldRefresh =
            document.activeElement === elements.tenantInput ||
            (elements.tenantDropdown && !elements.tenantDropdown.classList.contains("hidden"));
        if (query && shouldRefresh) {
            renderTenantResults(filterTenantResults(query));
        } else {
            hideTenantDropdown();
        }
    }

    syncTenantClearButton();
}

function syncTenantClearButton() {
    if (!elements?.tenantClear || !elements?.tenantInput) return;
    const hasValue = Boolean(normalizeId(elements.tenantInput.value));
    elements.tenantClear.classList.toggle("hidden", !hasValue);
}

function hideTenantDropdown() {
    if (elements?.tenantDropdown) {
        elements.tenantDropdown.classList.add("hidden");
    }
}

function filterTenantResults(query) {
    const needle = normalizeSearchValue(query);
    if (!needle) return [];
    return tenantSearchItems
        .filter((entry) => entry.searchValue.includes(needle))
        .slice(0, MAX_TENANT_RESULTS);
}

function renderTenantResults(results) {
    if (!elements?.tenantDropdown || !elements?.tenantList) return;
    tenantSearchResults = results;
    elements.tenantList.innerHTML = "";

    if (!results.length) {
        if (elements.tenantEmpty) elements.tenantEmpty.classList.remove("hidden");
        elements.tenantDropdown.classList.remove("hidden");
        return;
    }

    if (elements.tenantEmpty) elements.tenantEmpty.classList.add("hidden");

    results.forEach((entry) => {
        const row = document.createElement("button");
        row.type = "button";
        row.className = "w-full text-left px-3 py-2 hover:bg-slate-50";
        row.dataset.tenantId = entry.tenantId;
        const nameLine = document.createElement("div");
        nameLine.className = "text-[11px] font-semibold text-slate-800";
        nameLine.textContent = entry.name || "Tenant";
        row.appendChild(nameLine);
        elements.tenantList.appendChild(row);
    });

    elements.tenantDropdown.classList.remove("hidden");
}

function findExactTenantMatch(value) {
    const needle = normalizeSearchValue(value);
    if (!needle) return null;
    const matches = tenantSearchItems.filter(
        (entry) => normalizeSearchValue(entry.name) === needle
    );
    return matches.length === 1 ? matches[0] : null;
}

function selectTenant(entry) {
    if (!entry) return;
    selection.tenantId = normalizeId(entry.tenantId);
    selection.tenancyId = "";
    selectedTenantEntry = entry;
    if (elements?.tenantInput) {
        elements.tenantInput.value = entry.name;
    }
    syncTenantClearButton();
    populateTenancyOptions(entry.tenant);
    hideTenantDropdown();
}

function formatDateShort(value) {
    const raw = normalizeId(value);
    if (!raw) return "";
    const match = raw.match(/^(\d{4}-\d{2}-\d{2})/);
    if (match) return match[1];
    const date = new Date(raw);
    if (!Number.isNaN(date.getTime())) {
        return date.toISOString().slice(0, 10);
    }
    return raw;
}

function coerceIsoDate(value) {
    if (!value) return "";
    // If already a Date object, use UTC methods to avoid timezone shift
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        const year = value.getUTCFullYear();
        const month = String(value.getUTCMonth() + 1).padStart(2, "0");
        const day = String(value.getUTCDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }
    const raw = normalizeId(value);
    if (!raw) return "";
    // Match YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD format - return directly without parsing
    const matchYmd = raw.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (matchYmd) {
        const [, year, month, day] = matchYmd;
        return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
    }
    // Match DD-MM-YYYY or DD/MM/YYYY format
    const matchDmy = raw.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
    if (matchDmy) {
        const [, day, month, year] = matchDmy;
        return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
    }
    // For other date formats, parse with T00:00:00 to ensure local time interpretation
    const hasTime = raw.includes("T") || raw.includes(" ");
    const parsed = new Date(hasTime ? raw : raw + "T00:00:00");
    if (!Number.isNaN(parsed.getTime())) {
        const year = parsed.getFullYear();
        const month = String(parsed.getMonth() + 1).padStart(2, "0");
        const day = String(parsed.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }
    return "";
}


function formatDateForExport(value) {
    const iso = coerceIsoDate(value);
    return iso ? formatDateForDoc(iso) : "";
}

function formatPayableDate(value) {
    const raw = normalizeId(value);
    if (!raw) return "-";
    const numeric = parseInt(raw, 10);
    if (!Number.isNaN(numeric) && /^\d+$/.test(raw)) {
        return `${toOrdinal(numeric)} of every month`;
    }
    return raw;
}

function formatRentRevision(numberValue, unitValue) {
    const num = Number(numberValue);
    if (Number.isNaN(num) || num <= 0) return "-";
    const unitRaw = normalizeId(unitValue).toLowerCase();
    let unit = unitRaw || "month";
    if (unit.startsWith("year")) unit = "year";
    if (unit.startsWith("mon")) unit = "month";
    const label = num === 1 ? unit : `${unit}s`;
    return `Every ${num} ${label}`;
}

function formatNumber(value, { decimals = 2 } = {}) {
    if (value === null || value === undefined || value === "") return "";
    const numeric = Number(value);
    if (Number.isNaN(numeric)) return "";
    return numeric.toLocaleString("en-IN", {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
    });
}

function formatUnits(value) {
    if (value === null || value === undefined || value === "") return "";
    const numeric = Number(value);
    if (Number.isNaN(numeric)) return "";
    return numeric.toLocaleString("en-IN", { maximumFractionDigits: 0 });
}

function parseBoolean(value) {
    if (value === true || value === false) return value;
    if (value === 1 || value === 0) return Boolean(value);
    if (typeof value === "string") {
        const normalized = value.trim().toLowerCase();
        if (!normalized) return null;
        return normalized !== "false" && normalized !== "no" && normalized !== "0";
    }
    return value != null ? Boolean(value) : null;
}

function getBillMonthKey(bill) {
    return normalizeMonthKey(bill?.month_key || bill?.monthKey || bill?.month || "");
}

function getPaymentSortTimestamp(payment = {}) {
    const raw =
        payment.paymentDateTime ||
        payment.payment_date ||
        payment.paymentDate ||
        payment.date ||
        payment.createdAt ||
        payment.created_at ||
        "";
    const date = new Date(raw);
    if (!Number.isNaN(date.getTime())) return date.getTime();
    const match = normalizeId(raw).match(/^(\d{4}-\d{2}-\d{2})/);
    if (match) {
        const parsed = new Date(`${match[1]}T00:00:00`);
        return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
    }
    return 0;
}

function buildTenancyLabel(tenancy, tenant, unitMap) {
    const unit =
        unitMap.get(normalizeId(tenancy?.unit_id || tenancy?.unitId)) ||
        unitMap.get(normalizeId(tenant?.unitId)) ||
        null;
    const fallbackWing = normalizeId(tenancy?.wing || tenant?.wing || "");
    const fallbackUnitNumber = normalizeId(
        tenancy?.unitNumber ||
            tenancy?.unit_number ||
            tenant?.unitNumber ||
            tenant?.unit_number ||
            ""
    );
    const fallbackLabel = [fallbackWing, fallbackUnitNumber].filter(Boolean).join(" - ");
    const unitLabel =
        buildUnitLabel(unit) ||
        fallbackLabel ||
        tenancy?.unitLabel ||
        tenant?.unitLabel ||
        tenant?.unitNumber ||
        tenant?.unit_number ||
        "Unit";
    // Use commencement_date directly from tenancy table; only use tenant data if tenancy is empty
    const startDate = formatDateForExport(
        tenancy?.commencement_date ||
            tenancy?.commencementDate ||
            tenancy?.agreement_date ||
            ""
    ) || formatDateForExport(
        tenant?.tenancyCommencement ||
            tenant?.templateData?.tenancy_comm_raw ||
            ""
    );
    const endDate = formatDateForExport(tenancy?.end_date || tenancy?.endDate || "");
    const statusLabel = isTenancyActive(tenancy) ? "Active" : "Inactive";
    const dateRange = startDate || endDate ? `${startDate || "-"} to ${endDate || "Ongoing"}` : "";
    return `${unitLabel} - ${statusLabel}${dateRange ? ` (${dateRange})` : ""}`;
}

function populateTenancyOptions(tenant) {
    if (!elements?.tenancySelect) return;
    const tenantId = normalizeId(tenant?.tenantId || tenant?.tenant_id);
    const tenancies = exportState.tenanciesByTenant.get(tenantId) || [];
    const historyFallback = Array.isArray(tenant?.tenancyHistory) ? tenant.tenancyHistory : [];

    const tenancyList = tenancies.length
        ? tenancies.map((tenancy) => ({
              tenancyId: resolveTenancyId(tenancy),
              raw: tenancy,
          }))
        : historyFallback.map((history) => ({
              tenancyId: resolveTenancyId(history),
              raw: history,
          }));

    elements.tenancySelect.innerHTML = '<option value="">Select tenancy</option>';
    elements.tenancySelect.disabled = tenancyList.length === 0;

    tenancyList
        .filter((item) => item.tenancyId)
        .sort((a, b) => {
            const activeDiff = Number(isTenancyActive(b.raw)) - Number(isTenancyActive(a.raw));
            if (activeDiff) return activeDiff;
            const aDate = formatDateShort(a.raw?.commencement_date || a.raw?.commencementDate || "");
            const bDate = formatDateShort(b.raw?.commencement_date || b.raw?.commencementDate || "");
            return aDate.localeCompare(bDate);
        })
        .forEach((item) => {
            const label = buildTenancyLabel(item.raw, tenant, exportState.unitMap);
            const option = document.createElement("option");
            option.value = item.tenancyId;
            option.textContent = label;
            elements.tenancySelect.appendChild(option);
        });

    selection.tenancyId = "";
    setExportButtonState(false);
    setExportHint(
        tenancyList.length
            ? "Select a tenancy to enable export."
            : "No tenancy history found for this tenant."
    );
}

function handleTenantInput({ force = false } = {}) {
    if (!elements?.tenantInput) return;
    const value = normalizeId(elements.tenantInput.value);
    syncTenantClearButton();
    if (!value) {
        clearTenantSelection();
        hideTenantDropdown();
        return;
    }

    if (selectedTenantEntry && normalizeSearchValue(selectedTenantEntry.name) !== normalizeSearchValue(value)) {
        clearTenantSelection({ keepInput: true });
    }

    const results = filterTenantResults(value);
    renderTenantResults(results);

    if (force) {
        const exact = findExactTenantMatch(value);
        if (exact) {
            selectTenant(exact);
            return;
        }
        if (results.length === 1) {
            selectTenant(results[0]);
            return;
        }
        showToast("Select a tenant from the list.", "warning");
        clearTenantSelection({ keepInput: true });
        hideTenantDropdown();
    }
}

function setExportLoading(loading) {
    if (elements?.spinner) {
        elements.spinner.classList.toggle("hidden", !loading);
    }
    if (elements?.exportBtn) {
        elements.exportBtn.disabled = loading || !selection.tenancyId;
        elements.exportBtn.classList.toggle("cursor-not-allowed", loading);
        elements.exportBtn.classList.toggle("opacity-60", loading);
    }
}

function ensurePdfLibsReady() {
    if (!window.jspdf || !window.jspdf.jsPDF) {
        showToast("PDF export library is not loaded.", "error");
        return false;
    }
    return true;
}

function sanitizeFileName(value, fallback) {
    const safe = normalizeId(value)
        .replace(/\s+/g, "_")
        .replace(/[^\w.-]/g, "");
    return safe || fallback || "Tenant_Export";
}

function buildExportPayload(tenantId, tenancyId, data) {
    const tenant =
        data.tenants.find((t) => normalizeId(t.tenancyId || t.tenancy_id) === tenancyId) ||
        data.tenants.find((t) => normalizeId(resolveTenantId(t)) === tenantId) ||
        {};
    const tenancy =
        data.tenancies.find((t) => normalizeId(t.tenancy_id || t.tenancyId) === tenancyId) || {};
    const unit =
        data.units.find(
            (u) => normalizeId(u.unit_id || u.unitId) === normalizeId(tenancy?.unit_id || tenancy?.unitId)
        ) || {};

    const resolvedTenantId = normalizeId(resolveTenantId(tenant)) || normalizeId(tenancy?.tenant_id || tenancy?.tenantId);
    const family = data.familyMembers.filter(
        (member) => normalizeId(member?.tenant_id || member?.tenantId) === resolvedTenantId
    );

    const wing = normalizeId(unit?.wing || tenant?.wing || tenancy?.wing);
    const bills = data.billLines.filter(
        (bill) => normalizeId(bill?.tenancy_id || bill?.tenancyId) === tenancyId
    );
    const billLineIds = new Set(
        bills
            .map((bill) => normalizeId(bill?.bill_line_id || bill?.billLineId))
            .filter(Boolean)
    );

    const readingsByMonth = new Map();
    data.readings.forEach((reading) => {
        const tenancyKey = normalizeId(reading?.tenancy_id || reading?.tenancyId);
        if (tenancyKey !== tenancyId) return;
        const monthKey = normalizeMonthKey(reading?.month_key || reading?.monthKey || "");
        if (monthKey) readingsByMonth.set(monthKey, reading);
    });

    const configByMonth = new Map();
    data.configs.forEach((cfg) => {
        const monthKey = normalizeMonthKey(cfg?.month_key || cfg?.monthKey || "");
        const cfgWing = normalizeId(cfg?.wing || "");
        if (monthKey && cfgWing && cfgWing.toLowerCase() === wing.toLowerCase()) {
            configByMonth.set(monthKey, cfg);
        }
    });

    const payments = data.payments.filter((payment) => {
        const matchTenancy = normalizeId(payment?.tenancyId || payment?.tenancy_id);
        if (matchTenancy && matchTenancy === tenancyId) return true;
        const billLineId = normalizeId(payment?.billLineId || payment?.bill_line_id);
        return billLineId ? billLineIds.has(billLineId) : false;
    });

    const revisions = data.revisions.filter(
        (rev) => normalizeId(rev?.tenancy_id || rev?.tenancyId) === tenancyId
    );

    return {
        tenant,
        tenancy,
        unit,
        family,
        wing,
        bills,
        readingsByMonth,
        configByMonth,
        payments,
        revisions,
    };
}

function matchPaymentsForBill(bill, payments, fallbackPayments, tenantKey, wing) {
    const billLineId = normalizeId(bill?.bill_line_id || bill?.billLineId);
    if (billLineId) {
        const matches = payments.filter(
            (payment) => normalizeId(payment?.billLineId || payment?.bill_line_id) === billLineId
        );
        if (matches.length) return matches;
    }

    const monthKey = getBillMonthKey(bill);
    if (!monthKey) return [];
    const normalizedTenantKey = normalizeKey(tenantKey);
    const normalizedWing = normalizeKey(wing);

    return fallbackPayments.filter((payment) => {
        const paymentMonth = normalizeMonthKey(payment?.monthKey || payment?.month_key || "");
        const paymentTenant = normalizeKey(payment?.tenantKey || payment?.tenantName || "");
        const paymentWing = normalizeKey(payment?.wing || "");
        if (paymentMonth !== monthKey) return false;
        if (normalizedTenantKey && paymentTenant && paymentTenant !== normalizedTenantKey) return false;
        if (normalizedWing && paymentWing && paymentWing !== normalizedWing) return false;
        return true;
    });
}

function computeTotals(bills) {
    return bills.reduce(
        (acc, bill) => {
            const total = Number(bill?.total_amount ?? bill?.totalAmount) || 0;
            const amountPaid = Number(bill?.amount_paid ?? bill?.amountPaid) || 0;
            const paid = total > 0 ? Math.min(amountPaid, total) : 0;
            const ratio = total > 0 ? paid / total : 0;
            acc.rentPaid += (Number(bill?.rent_amount ?? bill?.rentAmount) || 0) * ratio;
            acc.electricityPaid += (Number(bill?.electricity_amount ?? bill?.electricityAmount) || 0) * ratio;
            acc.motorPaid += (Number(bill?.motor_share_amount ?? bill?.motorShare) || 0) * ratio;
            acc.sweepPaid += (Number(bill?.sweep_amount ?? bill?.sweepAmount) || 0) * ratio;
            acc.electricityUnits += Number(bill?.electricity_units ?? bill?.electricityUnits) || 0;
            acc.totalPaid += paid;
            return acc;
        },
        {
            rentPaid: 0,
            electricityPaid: 0,
            motorPaid: 0,
            sweepPaid: 0,
            electricityUnits: 0,
            totalPaid: 0,
        }
    );
}

function buildRentUpdateRow(revision) {
    const rentLabel = formatNumber(revision.rentAmount, { decimals: 2 }) || "-";
    const note = normalizeId(revision.note);
    const monthLabel = normalizeId(revision.effectiveMonth);
    const lines = [`RENT UPDATE: ${rentLabel}${monthLabel ? ` (Effective ${monthLabel})` : ""}`];
    if (note) {
        lines.push(`NOTE: ${note}`);
    }
    const row = [
        {
            content: lines.join("\n"),
            colSpan: 20,
        },
    ];
    row.__type = "rent-update";
    return row;
}

function buildPaymentRow(payment) {
    const row = Array(20).fill("");
    row[13] = formatNumber(payment.amount, { decimals: 2 }) || "";
    row[14] = normalizeId(payment.mode || "");
    row[15] = formatDateShort(
        payment.paymentDateTime || payment.payment_date || payment.paymentDate || payment.date || ""
    );
    const link = normalizeId(payment.attachmentUrl || payment.attachment_url || "");
    row[16] = link && link.startsWith("http") ? link : normalizeId(payment.attachmentName || "");
    row[17] = normalizeId(payment.notes || payment.reference || "");
    row.__type = "payment";
    return row;
}

function buildBillRentRevisions(bills) {
    const revisions = [];
    let lastRent = null;
    bills.forEach((bill) => {
        const monthKey = bill._monthKey || normalizeMonthKey(bill?.month_key || bill?.monthKey || "");
        if (!monthKey) return;
        const rentAmount = Number(bill?.rent_amount ?? bill?.rentAmount) || 0;
        if (!rentAmount) return;
        if (lastRent === null) {
            lastRent = rentAmount;
            revisions.push({
                effectiveMonth: monthKey,
                rentAmount,
                note: "",
                createdAt: bill?.generated_at || bill?.generatedAt || bill?.created_at || "",
            });
            return;
        }
        if (rentAmount !== lastRent) {
            revisions.push({
                effectiveMonth: monthKey,
                rentAmount,
                note: "",
                createdAt: bill?.generated_at || bill?.generatedAt || bill?.created_at || "",
            });
            lastRent = rentAmount;
        }
    });
    return revisions;
}

function buildBillRows(payload) {
    const tenantKey =
        payload.tenant?.grnNumber ||
        payload.tenant?.tenantKey ||
        payload.tenant?.tenantFullName ||
        payload.tenant?.tenantName ||
        "";
    const wing = payload.wing || "";
    const paymentsFallback = payload.payments;

    const bills = payload.bills
        .map((bill) => ({
            ...bill,
            _monthKey: getBillMonthKey(bill) || normalizeId(bill?.month_key || bill?.monthKey || ""),
        }))
        .sort((a, b) => {
            if (a._monthKey && b._monthKey) {
                return a._monthKey.localeCompare(b._monthKey);
            }
            const aDate = formatDateShort(a.generated_at || a.created_at || "");
            const bDate = formatDateShort(b.generated_at || b.created_at || "");
            return aDate.localeCompare(bDate);
        });

    const revisions = payload.revisions
        .map((rev) => ({
            effectiveMonth: normalizeMonthKey(rev?.effective_month || rev?.effectiveMonth || ""),
            rentAmount: Number(rev?.rent_amount ?? rev?.rentAmount) || 0,
            note: normalizeId(rev?.note || ""),
            createdAt: rev?.created_at || rev?.createdAt || "",
        }))
        .filter((rev) => rev.effectiveMonth)
        .sort((a, b) => {
            const monthDiff = a.effectiveMonth.localeCompare(b.effectiveMonth);
            if (monthDiff !== 0) return monthDiff;
            return formatDateShort(a.createdAt).localeCompare(formatDateShort(b.createdAt));
        });

    const billRevisions = buildBillRentRevisions(bills);
    const revisionMap = new Map();

    revisions.forEach((rev) => {
        const existing = revisionMap.get(rev.effectiveMonth);
        if (!existing) {
            revisionMap.set(rev.effectiveMonth, rev);
            return;
        }
        const existingDate = formatDateShort(existing.createdAt);
        const nextDate = formatDateShort(rev.createdAt);
        if (nextDate && (!existingDate || nextDate >= existingDate)) {
            revisionMap.set(rev.effectiveMonth, rev);
        }
    });

    billRevisions.forEach((rev) => {
        const existing = revisionMap.get(rev.effectiveMonth);
        if (!existing || existing.rentAmount !== rev.rentAmount) {
            revisionMap.set(rev.effectiveMonth, {
                ...existing,
                ...rev,
                note: existing?.note || rev.note,
            });
        }
    });

    const mergedRevisions = Array.from(revisionMap.values()).sort((a, b) => {
        const monthDiff = a.effectiveMonth.localeCompare(b.effectiveMonth);
        if (monthDiff !== 0) return monthDiff;
        return formatDateShort(a.createdAt).localeCompare(formatDateShort(b.createdAt));
    });

    const baseRent =
        Number(payload.tenant?.rentAmount ?? payload.tenant?.currentRent) ||
        Number(bills[0]?.rent_amount ?? bills[0]?.rentAmount) ||
        0;

    if (baseRent && (!mergedRevisions.length || (bills[0]?._monthKey && mergedRevisions[0].effectiveMonth > bills[0]._monthKey))) {
        mergedRevisions.unshift({
            effectiveMonth: bills[0]?._monthKey || "",
            rentAmount: baseRent,
            note: "Initial rent",
            createdAt: "",
        });
    }

    const rows = [];
    let revisionIndex = 0;

    // Insert rent revision rows before the first bill of each effective month.
    bills.forEach((bill) => {
        const billMonth = bill._monthKey;
        while (revisionIndex < mergedRevisions.length && mergedRevisions[revisionIndex].effectiveMonth <= billMonth) {
            rows.push(buildRentUpdateRow(mergedRevisions[revisionIndex]));
            revisionIndex += 1;
        }

        const reading = payload.readingsByMonth.get(billMonth) || {};
        const config = payload.configByMonth.get(billMonth) || {};

        const electricityUnits =
            bill.electricity_units ??
            bill.electricityUnits ??
            Math.max(
                (Number(reading?.new_reading ?? reading?.newReading) || 0) -
                    (Number(reading?.prev_reading ?? reading?.prevReading) || 0),
                0
            );

        const billPayments = matchPaymentsForBill(bill, payload.payments, paymentsFallback, tenantKey, wing).sort(
            (a, b) => getPaymentSortTimestamp(a) - getPaymentSortTimestamp(b)
        );
        const paymentTotal = billPayments.reduce((sum, payment) => sum + (Number(payment?.amount) || 0), 0);
        const totalAmount = Number(bill?.total_amount ?? bill?.totalAmount) || 0;
        const rawPaid = Number(bill?.amount_paid ?? bill?.amountPaid) || 0;
        const amountPaid = rawPaid || paymentTotal;
        const paidFlag = parseBoolean(bill?.is_paid ?? bill?.isPaid);
        const status =
            paidFlag === true
                ? "Paid"
                : amountPaid > 0
                ? "Partial"
                : totalAmount <= 0
                ? "Paid"
                : "Unpaid";

        const primaryPayment = billPayments[0] || null;
        const includedValue = parseBoolean(reading?.included);
        const includedLabel =
            includedValue === true ? "Yes" : includedValue === false ? "No" : "-";
        const row = [
            billMonth || normalizeId(bill?.month_key || bill?.monthKey || ""),
            normalizeId(bill?.bill_id || bill?.billId || ""),
            formatDateShort(bill?.generated_at || bill?.generatedAt || ""),
            includedLabel,
            normalizeId(config?.electricity_rate || ""),
            normalizeId(reading?.prev_reading ?? reading?.prevReading ?? ""),
            normalizeId(reading?.new_reading ?? reading?.newReading ?? ""),
            formatUnits(electricityUnits),
            formatNumber(bill?.electricity_amount ?? bill?.electricityAmount, { decimals: 2 }),
            formatNumber(bill?.motor_share_amount ?? bill?.motorShare, { decimals: 2 }),
            formatNumber(bill?.sweep_amount ?? bill?.sweepAmount, { decimals: 2 }),
            formatNumber(bill?.rent_amount ?? bill?.rentAmount, { decimals: 2 }),
            formatNumber(totalAmount, { decimals: 2 }),
            primaryPayment ? formatNumber(primaryPayment.amount, { decimals: 2 }) : "",
            primaryPayment ? normalizeId(primaryPayment.mode || "") : "",
            primaryPayment
                ? formatDateShort(
                      primaryPayment.paymentDateTime ||
                          primaryPayment.payment_date ||
                          primaryPayment.paymentDate ||
                          primaryPayment.date ||
                          ""
                  )
                : "",
            primaryPayment
                ? (normalizeId(primaryPayment.attachmentUrl || primaryPayment.attachment_url || "").startsWith("http")
                      ? normalizeId(primaryPayment.attachmentUrl || primaryPayment.attachment_url || "")
                      : normalizeId(primaryPayment.attachmentName || ""))
                : "",
            primaryPayment ? normalizeId(primaryPayment.notes || primaryPayment.reference || "") : "",
            formatNumber(amountPaid, { decimals: 2 }),
            status,
        ];
        row.__type = "bill";
        rows.push(row);

        billPayments.slice(1).forEach((payment) => {
            rows.push(buildPaymentRow(payment));
        });
    });

    if (!rows.length) {
        const row = ["No bills recorded for this tenancy."];
        while (row.length < 20) row.push("");
        row.__type = "note";
        rows.push(row);
    }

    return rows;
}

function addSectionTitle(doc, title, y, layout) {
    doc.setFillColor(...layout.sectionFill);
    doc.rect(layout.padding, y, layout.contentWidth, 18, "F");
    doc.setFontSize(10);
    doc.setTextColor(...layout.sectionText);
    doc.setFont("helvetica", "bold");
    doc.text(title, layout.padding + 8, y + 12);
    doc.setFont("helvetica", "normal");
    return y + 24;
}

function applyPdfFooter(doc, layout, meta) {
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i += 1) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(...layout.mutedText);
        const footerY = layout.pageHeight - 14;
        doc.text(`Generated: ${meta.generatedAt}`, layout.padding, footerY);
        doc.text(`Page ${i} of ${pageCount}`, layout.pageWidth - layout.padding, footerY, {
            align: "right",
        });
    }
}

function buildPdfDocument(payload) {
    if (!ensurePdfLibsReady()) return null;
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    if (typeof doc.autoTable !== "function") {
        showToast("PDF table helper is not loaded.", "error");
        return null;
    }

    const layout = {
        pageWidth: doc.internal.pageSize.getWidth(),
        pageHeight: doc.internal.pageSize.getHeight(),
        padding: 12,
        headerHeight: 56,
        contentWidth: 0,
        headerFill: [15, 23, 42],
        sectionFill: [241, 245, 249],
        sectionText: [15, 23, 42],
        mutedText: [71, 85, 105],
    };
    layout.contentWidth = layout.pageWidth - layout.padding * 2;
    const tableMarginTop = layout.padding + 12;

    const tenantName =
        payload.tenant?.tenantFullName ||
        payload.tenant?.tenantName ||
        payload.tenant?.tenant_name ||
        payload.tenant?.templateData?.Tenant_Full_Name ||
        "Tenant";
    const unitLabel =
        buildUnitLabel(payload.unit) ||
        payload.unit?.unit_number ||
        payload.tenant?.unitNumber ||
        payload.tenant?.unit_number ||
        "";
    const statusLabel = isTenancyActive(payload.tenancy) ? "Active" : "Inactive";
    const tenantGrn =
        payload.tenant?.grnNumber ||
        payload.tenant?.grn_number ||
        getTemplateValue(payload.tenant, "GRN number") ||
        "-";
    const tenantAadhaar =
        payload.tenant?.tenantAadhaar ||
        payload.tenant?.tenantAadhar ||
        payload.tenant?.tenant_Aadhar ||
        payload.tenant?.templateData?.tenant_Aadhar ||
        "-";
    const tenantPhone =
        payload.tenant?.tenantMobile ||
        payload.tenant?.tenant_mobile ||
        payload.tenant?.phone ||
        payload.tenant?.templateData?.tenant_mobile ||
        "-";
    const tenantAddress =
        payload.tenant?.tenantPermanentAddress ||
        payload.tenant?.tenantAddress ||
        payload.tenant?.tenant_address ||
        payload.tenant?.templateData?.Tenant_Permanent_Address ||
        "-";
    const tenantOccupation =
        payload.tenant?.tenantOccupation ||
        payload.tenant?.tenant_occupation ||
        payload.tenant?.templateData?.tenant_occupation ||
        "-";

    doc.setFillColor(...layout.headerFill);
    doc.rect(0, 0, layout.pageWidth, layout.headerHeight, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("Tenant Data Export", layout.pageWidth / 2, 32, { align: "center" });
    doc.setFontSize(9);
    doc.setFont("helvetica", "normal");
    const subtitleParts = [tenantName, unitLabel, statusLabel].filter(Boolean);
    doc.text(subtitleParts.join(" | "), layout.pageWidth / 2, 48, { align: "center" });

    let cursorY = layout.headerHeight + 16;

    cursorY = addSectionTitle(doc, "Tenant Profile", cursorY, layout);
    doc.autoTable({
        startY: cursorY,
        theme: "plain",
        margin: {
            left: layout.padding,
            right: layout.padding,
            top: tableMarginTop,
            bottom: 32,
        },
        tableWidth: layout.contentWidth,
        styles: {
            font: "helvetica",
            fontSize: 9,
            cellPadding: 3,
            lineWidth: 0,
            textColor: layout.sectionText,
        },
        columnStyles: {
            0: { cellWidth: 80, fontStyle: "bold" },
            1: { cellWidth: 200 },
            2: { cellWidth: 80, fontStyle: "bold" },
            3: { cellWidth: 200 },
        },
        body: [
            ["Name", tenantName || "-", "Aadhaar", tenantAadhaar],
            ["Phone", tenantPhone, "Address", tenantAddress],
            ["Occupation", tenantOccupation, "", ""],
        ],
    });
    cursorY = doc.lastAutoTable.finalY + 12;

    cursorY = addSectionTitle(doc, "Tenancy Detail", cursorY, layout);
    // Use commencement_date directly from tenancy table; only use tenant data if tenancy is empty
    const tenancyStart = formatDateForExport(
        payload.tenancy?.commencement_date ||
            payload.tenancy?.commencementDate ||
            payload.tenancy?.agreement_date ||
            ""
    ) || formatDateForExport(
        payload.tenant?.tenancyCommencement ||
            payload.tenant?.templateData?.tenancy_comm_raw ||
            ""
    );
    const tenancyEnd = formatDateForExport(
        payload.tenancy?.end_date ||
            payload.tenancy?.endDate ||
            payload.tenant?.tenancyEndDate ||
            payload.tenant?.templateData?.tenancy_end_raw ||
            ""
    );
    doc.autoTable({
        startY: cursorY,
        theme: "plain",
        margin: {
            left: layout.padding,
            right: layout.padding,
            top: tableMarginTop,
            bottom: 32,
        },
        tableWidth: layout.contentWidth,
        styles: {
            font: "helvetica",
            fontSize: 9,
            cellPadding: 3,
            lineWidth: 0,
            textColor: layout.sectionText,
        },
        columnStyles: {
            0: { cellWidth: 100, fontStyle: "bold" },
            1: { cellWidth: 180 },
            2: { cellWidth: 120, fontStyle: "bold" },
            3: { cellWidth: 180 },
        },
        body: [
            ["GRN No", tenantGrn || "-", "Rent Revision", formatRentRevision(payload.tenancy?.rent_revision_number || payload.tenant?.rentRevisionNumber || payload.tenant?.templateData?.rent_rev_number, payload.tenancy?.rent_revision_unit || payload.tenant?.rentRevisionUnit || getTemplateValue(payload.tenant, "rent_rev year_mon"))],
            ["Wing", payload.wing || "-", "Unit", unitLabel || "-"],
            ["Deposit", formatNumber(payload.tenancy?.security_deposit || payload.tenant?.securityDeposit || payload.tenant?.templateData?.secu_depo, { decimals: 2 }) || "-", "Rent Payable Date", formatPayableDate(payload.tenancy?.rent_payable_day || payload.tenant?.payableDate || payload.tenant?.templateData?.payable_date_raw)],
            ["Floor", payload.unit?.floor || payload.tenant?.floor || payload.tenant?.templateData?.floor_of_building || "-", "Start Date", tenancyStart || "-"],
            ["End Date", tenancyEnd || "Ongoing", "Status", statusLabel],
        ],
    });
    cursorY = doc.lastAutoTable.finalY + 12;

    cursorY = addSectionTitle(doc, "Family Members", cursorY, layout);
    const familyRows = payload.family.length
        ? payload.family.map((member) => [
              member.name || "-",
              member.relationship || "-",
              member.occupation || "-",
              member.aadhaar || "-",
              member.address || "-",
          ])
        : [["No family members recorded.", "", "", "", ""]];
    doc.autoTable({
        startY: cursorY,
        theme: "grid",
        margin: {
            left: layout.padding,
            right: layout.padding,
            top: tableMarginTop,
            bottom: 32,
        },
        tableWidth: layout.contentWidth,
        head: [["Name", "Relationship", "Occupation", "Aadhaar", "Address"]],
        styles: {
            font: "helvetica",
            fontSize: 9,
            cellPadding: 3,
            textColor: layout.sectionText,
        },
        headStyles: {
            fillColor: layout.headerFill,
            textColor: 255,
            fontStyle: "bold",
        },
        body: familyRows,
    });
    cursorY = doc.lastAutoTable.finalY + 12;

    cursorY = addSectionTitle(doc, "Totals", cursorY, layout);
    const totals = computeTotals(payload.bills);
    doc.autoTable({
        startY: cursorY,
        theme: "grid",
        margin: {
            left: layout.padding,
            right: layout.padding,
            top: tableMarginTop,
            bottom: 32,
        },
        tableWidth: layout.contentWidth,
        styles: {
            font: "helvetica",
            fontSize: 9,
            cellPadding: 3,
            textColor: layout.sectionText,
        },
        columnStyles: {
            0: { cellWidth: 220, fontStyle: "bold" },
            1: { cellWidth: 120 },
        },
        body: [
            ["Total Rent Paid", formatNumber(totals.rentPaid, { decimals: 2 }) || "0.00"],
            ["Total Electricity Paid", formatNumber(totals.electricityPaid, { decimals: 2 }) || "0.00"],
            ["Total Electricity Units Consumed", formatUnits(totals.electricityUnits) || "0"],
            ["Total Motor Share Paid", formatNumber(totals.motorPaid, { decimals: 2 }) || "0.00"],
            ["Total Paid (All Charges)", formatNumber(totals.totalPaid, { decimals: 2 }) || "0.00"],
        ],
    });
    cursorY = doc.lastAutoTable.finalY + 12;

    cursorY = addSectionTitle(doc, "Bills", cursorY, layout);
    const billRows = buildBillRows(payload);
    const billHeaders = [
        "Month",
        "Bill ID",
        "Generated",
        "Included",
        "Elec Rate",
        "Prev Read",
        "New Read",
        "Units",
        "Elec Amt",
        "Motor",
        "Sweep",
        "Rent",
        "Total",
        "Payment",
        "Mode",
        "Paid On",
        "Attachment",
        "Notes",
        "Total Paid",
        "Status",
    ];

    doc.autoTable({
        startY: cursorY,
        theme: "grid",
        margin: {
            left: layout.padding,
            right: layout.padding,
            top: tableMarginTop,
            bottom: 32,
        },
        tableWidth: layout.contentWidth,
        head: [billHeaders],
        body: billRows,
        styles: {
            font: "helvetica",
            fontSize: 7,
            cellPadding: 2,
            overflow: "linebreak",
            textColor: layout.sectionText,
        },
        headStyles: {
            fillColor: layout.headerFill,
            textColor: 255,
            fontStyle: "bold",
        },
        columnStyles: {
            0: { cellWidth: 36 },
            1: { cellWidth: 60 },
            2: { cellWidth: 48 },
            3: { cellWidth: 28 },
            4: { cellWidth: 34 },
            5: { cellWidth: 34 },
            6: { cellWidth: 34 },
            7: { cellWidth: 28 },
            8: { cellWidth: 38 },
            9: { cellWidth: 34 },
            10: { cellWidth: 34 },
            11: { cellWidth: 38 },
            12: { cellWidth: 38 },
            13: { cellWidth: 38 },
            14: { cellWidth: 30 },
            15: { cellWidth: 48 },
            16: { cellWidth: 52 },
            17: { cellWidth: 94 },
            18: { cellWidth: 42 },
            19: { cellWidth: 30 },
        },
        didParseCell: (data) => {
            const rowType = data.row.raw?.__type;
            if (rowType === "rent-update") {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
                data.cell.styles.fontStyle = "bold";
                data.cell.styles.halign = "left";
            }
            if (rowType === "payment") {
                data.cell.styles.textColor = layout.mutedText;
            }
            if (rowType === "note") {
                data.cell.styles.textColor = layout.mutedText;
                data.cell.styles.fontStyle = "italic";
                if (data.column.index > 0) {
                    data.cell.text = [""];
                }
            }
        },
    });

    applyPdfFooter(doc, layout, { generatedAt: new Date().toLocaleString() });
    return doc;
}

function revokeLastPdfUrl() {
    if (lastPdfDownloadUrl) {
        URL.revokeObjectURL(lastPdfDownloadUrl);
        lastPdfDownloadUrl = "";
    }
}

function syncPdfExportModal(fileName, objectUrlOrPath) {
    const modal = document.getElementById("pdfExportModal");
    const fileNameLabel = document.getElementById("pdfExportFileName");
    const locationLabel = document.getElementById("pdfExportLocation");
    const openLink = document.getElementById("pdfExportOpen");

    const safeFileName = fileName || "Tenant_Export.pdf";
    lastPdfFileName = safeFileName;

    if (fileNameLabel) fileNameLabel.textContent = safeFileName;
    if (locationLabel) {
        locationLabel.textContent = `Saved to: Downloads/${safeFileName}`;
    }

    if (openLink) {
        openLink.removeAttribute("download");
        openLink.removeAttribute("href");
        openLink.target = "";

        const isTauri = !!window.__TAURI__;
        if (objectUrlOrPath) {
            if (isTauri) {
                openLink.href = "#";
                openLink.dataset.filepath = objectUrlOrPath;
            } else {
                openLink.href = objectUrlOrPath;
                openLink.download = safeFileName;
                openLink.target = "_blank";
                openLink.rel = "noopener";
            }
            openLink.dataset.openReady = "true";
        }

        openLink.classList.toggle("pointer-events-none", !objectUrlOrPath);
        openLink.classList.toggle("opacity-50", !objectUrlOrPath);
    }

    if (modal) showModal(modal);
}

function wirePdfExportModal() {
    if (pdfModalWired) return;
    pdfModalWired = true;

    const closeBtn = document.getElementById("pdfExportClose");
    const modal = document.getElementById("pdfExportModal");
    const openLink = document.getElementById("pdfExportOpen");

    if (closeBtn) {
        closeBtn.addEventListener("click", () => {
            if (modal) hideModal(modal);
        });
    }

    if (modal) {
        modal.addEventListener("click", (e) => {
            if (e.target === modal) {
                hideModal(modal);
            }
        });
    }

    if (openLink) {
        openLink.addEventListener("click", async (e) => {
            if (!lastPdfDownloadUrl && !lastPdfFilePath) {
                e.preventDefault();
                showToast("Export a PDF first, then try opening it again.", "warning");
                return;
            }

            const closeModal = () => {
                if (modal) hideModal(modal);
            };

            if (window.__TAURI__ && lastPdfFilePath) {
                e.preventDefault();
                try {
                    await window.__TAURI__.opener.openPath(lastPdfFilePath);
                } catch (err) {
                    console.error("Failed to open file:", err);
                    showToast(`Failed to open file: ${err}`, "error");
                }
                closeModal();
                return;
            }

            if (typeof navigator !== "undefined" && navigator.msSaveOrOpenBlob && lastPdfBlob) {
                e.preventDefault();
                navigator.msSaveOrOpenBlob(lastPdfBlob, lastPdfFileName || "Tenant_Export.pdf");
                closeModal();
                return;
            }

            closeModal();
        });
    }
}

async function handleExportClick() {
    if (!elements) return;
    const tenancyId = normalizeId(selection.tenancyId || elements.tenancySelect?.value);
    const tenantId = normalizeId(selection.tenantId);
    if (!tenancyId) {
        showToast("Select a tenancy first.", "warning");
        return;
    }
    if (!ensurePdfLibsReady()) return;

    setExportLoading(true);
    try {
        const [
            tenants,
            tenancies,
            units,
            familyMembers,
            billLines,
            readings,
            configs,
            payments,
            revisions,
        ] = await Promise.all([
            getLocalList(LOCAL_KEYS.tenants),
            getLocalList(LOCAL_KEYS.tenancies),
            getLocalList(LOCAL_KEYS.units),
            getLocalList(LOCAL_KEYS.familyMembers),
            getLocalList(LOCAL_KEYS.billLines),
            getLocalList(LOCAL_KEYS.tenantMonthlyReadings),
            getLocalList(LOCAL_KEYS.wingMonthlyConfig),
            getLocalList(LOCAL_KEYS.payments),
            getLocalList(LOCAL_KEYS.rentRevisionsAll),
        ]);

        const payload = buildExportPayload(tenantId, tenancyId, {
            tenants,
            tenancies,
            units,
            familyMembers,
            billLines,
            readings,
            configs,
            payments,
            revisions,
        });

        const doc = buildPdfDocument(payload);
        if (!doc) return;

        const grn =
            payload.tenant?.grnNumber ||
            payload.tenant?.grn_number ||
            getTemplateValue(payload.tenant, "GRN number") ||
            "GRN";
        const safeTenant = sanitizeFileName(payload.tenant?.tenantFullName || payload.tenant?.tenantName, "Tenant");
        const safeGrn = sanitizeFileName(grn, "GRN");
        const safeTenancy = sanitizeFileName(tenancyId, "Tenancy");
        const fileName = `${safeTenant}_${safeGrn}_${safeTenancy}_Export.pdf`;

        if (window.__TAURI__) {
            const out = doc.output("arraybuffer");
            const uint8 = new Uint8Array(out);
            const { path, fs } = window.__TAURI__;
            const downloadDir = await path.downloadDir();
            const filePath = await path.join(downloadDir, fileName);
            await fs.writeFile(filePath, uint8);
            lastPdfFilePath = filePath;
            revokeLastPdfUrl();
            syncPdfExportModal(fileName, filePath);
            showToast(`Saved to Downloads: ${fileName}`, "success");
            return;
        }

        const blob = doc.output("blob");
        revokeLastPdfUrl();
        lastPdfBlob = blob;
        lastPdfDownloadUrl = URL.createObjectURL(blob);
        saveAs(blob, fileName);
        syncPdfExportModal(fileName, lastPdfDownloadUrl);
        showToast(`PDF downloaded as "${fileName}"`, "success");
    } catch (err) {
        console.error("PDF export failed", err);
        showToast("PDF export failed. Check the console for details.", "error");
    } finally {
        setExportLoading(false);
    }
}

export function initExportDataFeature() {
    if (initialized) return;
    initialized = true;
    elements = getExportElements();
    if (!elements?.tenantInput || !elements?.tenancySelect || !elements?.exportBtn || !elements?.tenantDropdown || !elements?.tenantList) {
        return;
    }

    wirePdfExportModal();

    elements.tenantInput.addEventListener("input", () => {
        handleTenantInput();
    });

    elements.tenantInput.addEventListener("focus", () => {
        const value = normalizeId(elements.tenantInput.value);
        if (value) {
            renderTenantResults(filterTenantResults(value));
        }
    });

    elements.tenantInput.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
            hideTenantDropdown();
            return;
        }
        if (event.key === "Enter") {
            const value = normalizeId(elements.tenantInput.value);
            if (!value) return;
            const exact = findExactTenantMatch(value);
            if (exact) {
                event.preventDefault();
                selectTenant(exact);
                return;
            }
            if (tenantSearchResults.length === 1) {
                event.preventDefault();
                selectTenant(tenantSearchResults[0]);
            }
        }
    });

    elements.tenantInput.addEventListener("change", () => {
        handleTenantInput({ force: true });
    });

    if (elements.tenantClear) {
        elements.tenantClear.addEventListener("click", () => {
            clearTenantSelection();
            hideTenantDropdown();
            elements.tenantInput.focus();
        });
    }

    elements.tenantList.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        const button = target.closest("button[data-tenant-id]");
        if (!button) return;
        const tenantId = normalizeId(button.dataset.tenantId);
        const entry = tenantSearchMap.get(tenantId);
        if (!entry) return;
        selectTenant(entry);
    });

    document.addEventListener("click", (event) => {
        if (elements.tenantDropdown.classList.contains("hidden")) return;
        if (elements.tenantDropdown.contains(event.target) || elements.tenantInput.contains(event.target)) return;
        hideTenantDropdown();
    });

    elements.tenancySelect.addEventListener("change", () => {
        selection.tenancyId = normalizeId(elements.tenancySelect.value);
        if (selection.tenancyId) {
            setExportButtonState(true);
            setExportHint("Ready to export.");
        } else {
            setExportButtonState(false);
            setExportHint("Select a tenancy to enable export.");
        }
    });

    elements.exportBtn.addEventListener("click", handleExportClick);

    document.addEventListener("sync:completed", () => {
        refreshExportData();
    });

    syncTenantClearButton();
    refreshExportData();
}

export async function refreshExportData() {
    const [tenants, tenancies, units] = await Promise.all([
        getLocalList(LOCAL_KEYS.tenants),
        getLocalList(LOCAL_KEYS.tenancies),
        getLocalList(LOCAL_KEYS.units),
    ]);

    exportState.tenants = Array.isArray(tenants) ? tenants : [];
    exportState.tenancies = Array.isArray(tenancies) ? tenancies : [];
    exportState.units = Array.isArray(units) ? units : [];
    buildTenantDirectory();
    updateTenantOptions();
}
