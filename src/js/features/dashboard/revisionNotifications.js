import { getLocalList, LOCAL_KEYS } from "../../api/localStore.js";
import { buildUnitLabel, normalizeMonthKey } from "../../utils/formatters.js";
import { escapeHtml } from "../../utils/htmlUtils.js";
import { hideModal, showModal } from "../../utils/ui.js";

const revisionNotificationState = {
    items: [],
    initialized: false,
    loading: false,
};

function getElements() {
    return {
        widgetList: document.getElementById("revisionNotificationList"),
        widgetEmpty: document.getElementById("revisionNotificationEmpty"),
        widgetCount: document.getElementById("revisionNotificationCount"),
        expandBtn: document.getElementById("revisionNotificationExpand"),
        modal: document.getElementById("revisionNotificationModal"),
        modalClose: document.getElementById("revisionNotificationModalClose"),
        modalList: document.getElementById("revisionNotificationModalList"),
        modalEmpty: document.getElementById("revisionNotificationModalEmpty"),
        modalCount: document.getElementById("revisionNotificationModalCount"),
    };
}

function parseNumber(value, fallback = 0) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
}

function normalizeKey(value) {
    return (value ?? "").toString().trim().toLowerCase();
}

function parseDate(raw) {
    if (!raw) return null;
    if (raw instanceof Date && !Number.isNaN(raw.getTime())) return raw;
    const direct = new Date(raw);
    if (!Number.isNaN(direct.getTime())) return direct;

    if (typeof raw === "string") {
        const parts = raw.split(/[\/\-\.]/).map((part) => part.trim());
        if (parts.length === 3) {
            let [first, second, third] = parts;
            if (third.length === 2) third = `20${third}`;
            let day = parseInt(first, 10);
            let month = parseInt(second, 10);
            const year = parseInt(third, 10);
            if (Number.isNaN(day) || Number.isNaN(month) || Number.isNaN(year)) return null;
            if (month > 12 && day <= 12) {
                [day, month] = [month, day];
            }
            const alt = new Date(year, month - 1, day);
            if (!Number.isNaN(alt.getTime())) return alt;
        }
    }
    return null;
}

function addMonths(date, months) {
    const next = new Date(date.getTime());
    const day = next.getDate();
    next.setDate(1);
    next.setMonth(next.getMonth() + months);
    const maxDay = new Date(next.getFullYear(), next.getMonth() + 1, 0).getDate();
    next.setDate(Math.min(day, maxDay));
    return next;
}

function monthKeyFromDate(date) {
    if (!date) return "";
    return normalizeMonthKey(date) || "";
}

function formatMonthLabel(key) {
    const normalized = normalizeMonthKey(key);
    if (!normalized) return "-";
    if (/^\d{4}-\d{2}$/.test(normalized)) {
        const [year, month] = normalized.split("-");
        const date = new Date(Number(year), Number(month) - 1, 1);
        if (!Number.isNaN(date.getTime())) {
            return date.toLocaleDateString("en-GB", { month: "short", year: "numeric" });
        }
    }
    return normalized;
}

function formatIntervalLabel(interval) {
    if (!interval) return "";
    const unit = interval.unitLabel || "month";
    const plural = interval.number === 1 ? "" : "s";
    return `${interval.number} ${unit}${plural}`;
}

function resolveInterval(tenancy, tenant) {
    const number = parseNumber(
        tenancy?.rent_revision_number ?? tenant?.rentRevisionNumber ?? tenant?.templateData?.rent_rev_number
    );
    if (!number || number <= 0) return null;
    const unitRaw = (
        tenancy?.rent_revision_unit ??
        tenant?.rentRevisionUnit ??
        tenant?.templateData?.["rent_rev year_mon"] ??
        ""
    )
        .toString()
        .trim()
        .toUpperCase();
    if (!unitRaw) return null;
    const isYear = unitRaw.includes("YEAR");
    const isMonth = unitRaw.includes("MONTH");
    if (!isYear && !isMonth) return null;
    const intervalMonths = isYear ? number * 12 : number;
    return {
        number,
        unitRaw,
        intervalMonths,
        unitLabel: isYear ? "year" : "month",
    };
}

function resolveCommencementDate(tenancy, tenant) {
    return (
        parseDate(tenancy?.commencement_date) ||
        parseDate(tenancy?.commencementDate) ||
        parseDate(tenant?.tenancyCommencement) ||
        parseDate(tenant?.templateData?.tenancy_comm_raw) ||
        parseDate(tenant?.templateData?.tenancy_comm) ||
        null
    );
}

function getNextRevisionDate(commencementDate, intervalMonths, today) {
    if (!commencementDate || !intervalMonths) return null;
    let next = addMonths(commencementDate, intervalMonths);
    let guard = 0;
    while (next < today && guard < 600) {
        next = addMonths(next, intervalMonths);
        guard += 1;
    }
    return next;
}

function isTenancyActive(tenancy, tenant) {
    const status = normalizeKey(tenancy?.status || tenancy?.tenancyStatus || "");
    if (status) return status === "active";
    if (typeof tenant?.activeTenant === "boolean") return tenant.activeTenant;
    return true;
}

function buildTenantMaps(tenants) {
    const byId = new Map();
    const byTenancy = new Map();
    tenants.forEach((tenant) => {
        const tenantId = normalizeKey(tenant?.tenant_id || tenant?.tenantId);
        const tenancyId = normalizeKey(tenant?.tenancyId || tenant?.tenancy_id);
        if (tenantId) byId.set(tenantId, tenant);
        if (tenancyId) byTenancy.set(tenancyId, tenant);
    });
    return { byId, byTenancy };
}

function buildUnitMap(units) {
    const map = new Map();
    units.forEach((unit) => {
        const id = normalizeKey(unit?.unit_id || unit?.unitId);
        if (id) map.set(id, unit);
    });
    return map;
}

function resolveUnitLabel(tenancy, tenant, unitMap) {
    const unitId = normalizeKey(tenancy?.unit_id || tenancy?.unitId || tenant?.unitId || tenant?.unit_id);
    const unit = unitId ? unitMap.get(unitId) : null;
    if (unit) return buildUnitLabel(unit);
    return buildUnitLabel({
        wing: tenant?.wing || tenancy?.wing || "",
        unit_number: tenant?.unitNumber || tenant?.unit_number || tenancy?.unit_number || "",
        unit_id: tenancy?.unit_id || tenancy?.unitId || "",
    });
}

function buildItems(tenancies, tenants, units) {
    const { byId, byTenancy } = buildTenantMaps(tenants);
    const unitMap = buildUnitMap(units);
    const today = new Date();
    const currentMonthKey = monthKeyFromDate(today);
    const items = [];

    tenancies.forEach((tenancy) => {
        if (!tenancy || typeof tenancy !== "object") return;
        const tenancyId = normalizeKey(tenancy?.tenancy_id || tenancy?.tenancyId);
        const tenantId = normalizeKey(tenancy?.tenant_id || tenancy?.tenantId);
        const tenant = byId.get(tenantId) || byTenancy.get(tenancyId) || null;
        if (!isTenancyActive(tenancy, tenant)) return;

        const commencementDate = resolveCommencementDate(tenancy, tenant);
        const interval = resolveInterval(tenancy, tenant);
        if (!commencementDate || !interval) return;

        const noticeMonths = parseNumber(
            tenancy?.tenant_notice_months ?? tenant?.tenantNoticeMonths ?? tenant?.templateData?.notice_num_t,
            0
        );
        const nextRevisionDate = getNextRevisionDate(commencementDate, interval.intervalMonths, today);
        if (!nextRevisionDate) return;

        const nextRevisionMonth = monthKeyFromDate(nextRevisionDate);
        const visibleFromDate = addMonths(nextRevisionDate, -(noticeMonths + 2));
        const visibleFromMonth = monthKeyFromDate(visibleFromDate);

        if (!currentMonthKey || !visibleFromMonth) return;
        if (currentMonthKey < visibleFromMonth) return;

        items.push({
            tenancyId,
            tenantName:
                tenant?.tenantFullName ||
                tenant?.tenantName ||
                tenant?.tenant_name ||
                tenant?.full_name ||
                tenant?.templateData?.Tenant_Full_Name ||
                "Unknown tenant",
            unitLabel: resolveUnitLabel(tenancy, tenant, unitMap) || "-",
            interval,
            noticeMonths,
            nextRevisionDate,
            nextRevisionMonth,
            visibleFromMonth,
        });
    });

    items.sort((a, b) => a.nextRevisionDate - b.nextRevisionDate);
    return items;
}

function setCountLabel(el, count) {
    if (!el) return;
    const label = `${count} upcoming`;
    el.textContent = label;
}

function renderWidget(items) {
    const { widgetList, widgetEmpty, widgetCount } = getElements();
    if (!widgetList) return;

    widgetList.innerHTML = "";
    setCountLabel(widgetCount, items.length);

    if (!items.length) {
        if (widgetEmpty) {
            widgetEmpty.textContent = "No upcoming revisions.";
            widgetEmpty.classList.remove("hidden");
        }
        return;
    }

    if (widgetEmpty) widgetEmpty.classList.add("hidden");

    const preview = items.slice(0, 3);
    preview.forEach((item) => {
        const tenantName = escapeHtml(item.tenantName);
        const unitLabel = escapeHtml(item.unitLabel);
        const nextLabel = escapeHtml(formatMonthLabel(item.nextRevisionMonth));

        const row = document.createElement("div");
        row.className =
            "flex items-center justify-between gap-2 rounded-md border border-rose-100 bg-white/70 px-2 py-1.5";
        row.innerHTML = `
            <div class="min-w-0">
                <div class="text-[11px] font-semibold text-rose-900 truncate">${tenantName}</div>
                <div class="text-[9px] text-rose-700 truncate">${unitLabel}</div>
            </div>
            <div class="text-[10px] font-semibold text-rose-800">${nextLabel}</div>
        `;
        widgetList.appendChild(row);
    });
}

function renderModal(items) {
    const { modalList, modalEmpty, modalCount } = getElements();
    if (!modalList) return;

    modalList.innerHTML = "";
    setCountLabel(modalCount, items.length);

    if (!items.length) {
        if (modalEmpty) modalEmpty.classList.remove("hidden");
        return;
    }

    if (modalEmpty) modalEmpty.classList.add("hidden");

    items.forEach((item) => {
        const tenantName = escapeHtml(item.tenantName);
        const unitLabel = escapeHtml(item.unitLabel);
        const nextLabel = escapeHtml(formatMonthLabel(item.nextRevisionMonth));
        const visibleLabel = escapeHtml(formatMonthLabel(item.visibleFromMonth));
        const intervalLabel = escapeHtml(formatIntervalLabel(item.interval));
        const noticeLabel = escapeHtml(`${item.noticeMonths || 0} mo`);

        const row = document.createElement("div");
        row.className = "flex items-start justify-between gap-3 py-2";
        row.innerHTML = `
            <div class="min-w-0">
                <div class="text-[12px] font-semibold text-slate-800 truncate">${tenantName}</div>
                <div class="text-[10px] text-slate-500 truncate">${unitLabel} • Every ${intervalLabel} • Notice ${noticeLabel}</div>
            </div>
            <div class="text-right">
                <div class="text-[11px] font-semibold text-rose-700">${nextLabel}</div>
                <div class="text-[9px] text-slate-500">Visible from ${visibleLabel}</div>
            </div>
        `;
        modalList.appendChild(row);
    });
}

async function loadRevisionNotifications() {
    if (revisionNotificationState.loading) return;
    revisionNotificationState.loading = true;
    const { widgetEmpty } = getElements();
    if (widgetEmpty) {
        widgetEmpty.textContent = "Loading revisions...";
        widgetEmpty.classList.remove("hidden");
    }

    try {
        const [tenancies, tenants, units] = await Promise.all([
            getLocalList(LOCAL_KEYS.tenancies, []),
            getLocalList(LOCAL_KEYS.tenants, []),
            getLocalList(LOCAL_KEYS.units, []),
        ]);
        revisionNotificationState.items = buildItems(
            Array.isArray(tenancies) ? tenancies : [],
            Array.isArray(tenants) ? tenants : [],
            Array.isArray(units) ? units : []
        );
    } catch (err) {
        console.error("Failed to load revision notifications", err);
        revisionNotificationState.items = [];
    }

    renderWidget(revisionNotificationState.items);
    renderModal(revisionNotificationState.items);
    revisionNotificationState.loading = false;
}

function bindModalEvents() {
    const { modal, modalClose, expandBtn } = getElements();
    if (expandBtn && !expandBtn.dataset.bound) {
        expandBtn.dataset.bound = "1";
        expandBtn.addEventListener("click", () => {
            if (modal) showModal(modal);
        });
    }

    if (modal && !modal.dataset.bound) {
        modal.dataset.bound = "1";
        modal.addEventListener("click", (event) => {
            if (event.target === modal) hideModal(modal);
        });
    }

    if (modalClose && !modalClose.dataset.bound) {
        modalClose.dataset.bound = "1";
        modalClose.addEventListener("click", () => {
            if (modal) hideModal(modal);
        });
    }
}

export function initRevisionNotifications() {
    if (revisionNotificationState.initialized) return;
    const { widgetList } = getElements();
    if (!widgetList) return;
    revisionNotificationState.initialized = true;
    bindModalEvents();

    document.addEventListener("tenancies:updated", loadRevisionNotifications);
    document.addEventListener("tenants:updated", loadRevisionNotifications);
    document.addEventListener("units:updated", loadRevisionNotifications);
    document.addEventListener("sync:completed", loadRevisionNotifications);
    document.addEventListener("rentRevisions:updated", loadRevisionNotifications);

    loadRevisionNotifications();
}
