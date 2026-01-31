import { getLocalList, LOCAL_KEYS } from "../../api/localStore.js";
import { saveRentRevision, saveRentRevisionReminder } from "../../api/sheets.js";
import { buildUnitLabel, formatCurrency, normalizeMonthKey } from "../../utils/formatters.js";
import { escapeHtml } from "../../utils/htmlUtils.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

const revisionNotificationState = {
    items: [],
    selections: new Map(),
    sendStatus: new Map(),
    decisionContext: null,
    suspendReload: false,
    pendingReload: false,
    initialized: false,
    loading: false,
};

const DEBUG_MONTH_KEY = "tenantApp.debugMonthKey";

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
        toggleAll: document.getElementById("revisionNotificationSendToggleAll"),
        autoSendBtn: document.getElementById("revisionNotificationAutoSendBtn"),
        autoSendCount: document.getElementById("revisionNotificationAutoSendCount"),
        decisionModal: document.getElementById("revisionDecisionModal"),
        decisionTitle: document.getElementById("revisionDecisionTitle"),
        decisionAmount: document.getElementById("revisionDecisionAmount"),
        decisionClose: document.getElementById("revisionDecisionClose"),
        decisionCancel: document.getElementById("revisionDecisionCancel"),
        decisionConfirm: document.getElementById("revisionDecisionConfirm"),
    };
}

function parseNumber(value, fallback = null) {
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

function formatMoney(value) {
    if (value === null || typeof value === "undefined") return "-";
    return formatCurrency(value, {
        currencySymbol: "Rs. ",
        parseMode: "int",
        minimumFractionDigits: 0,
        maximumFractionDigits: 0,
        emptyValue: "-",
        invalidValue: "-",
    });
}

function toSafeToken(value) {
    return (value || "row")
        .toString()
        .replace(/[^a-zA-Z0-9_-]/g, "_")
        .slice(0, 40);
}

function getItemId(item) {
    return `${item.tenancyId || "unknown"}__${item.nextRevisionMonth || "unknown"}`;
}

function resolveInterval(tenancy, tenant) {
    const number = parseNumber(
        tenancy?.rent_revision_number ?? tenant?.rentRevisionNumber ?? tenant?.templateData?.rent_rev_number,
        0
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

function resolveIncreaseAmount(tenancy, tenant) {
    return parseNumber(
        tenancy?.rent_increase_amount ??
            tenancy?.rentIncreaseAmount ??
            tenant?.rentIncrease ??
            tenant?.rentIncreaseAmount ??
            tenant?.templateData?.rent_inc ??
            null,
        null
    );
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
    const currentMonthKey = monthKeyFromDate(today);
    while (
        currentMonthKey &&
        monthKeyFromDate(addMonths(next, 1)) < currentMonthKey &&
        guard < 600
    ) {
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

function normalizeRevisionMonth(revision) {
    return normalizeMonthKey(
        revision?.effective_month ||
            revision?.effectiveMonth ||
            revision?.month_key ||
            revision?.monthKey ||
            revision?.month ||
            ""
    );
}

function normalizeRevisionAmount(revision) {
    return parseNumber(
        revision?.rent_amount ??
            revision?.rentAmount ??
            revision?.rent ??
            revision?.amount ??
            null,
        null
    );
}

function buildRevisionMap(revisions) {
    const map = new Map();
    revisions.forEach((rev) => {
        if (!rev || typeof rev !== "object") return;
        const tenancyId = normalizeKey(rev?.tenancy_id || rev?.tenancyId);
        if (!tenancyId) return;
        const monthKey = normalizeRevisionMonth(rev);
        if (!monthKey) return;
        const list = map.get(tenancyId) || [];
        list.push({
            monthKey,
            rentAmount: normalizeRevisionAmount(rev),
            note: (rev?.note || rev?.notes || "").toString().trim(),
        });
        map.set(tenancyId, list);
    });

    map.forEach((list) => {
        list.sort((a, b) => a.monthKey.localeCompare(b.monthKey));
    });

    return map;
}

function buildReminderMap(reminders) {
    const map = new Map();
    reminders.forEach((raw) => {
        if (!raw || typeof raw !== "object") return;
        const tenancyId = normalizeKey(raw?.tenancy_id || raw?.tenancyId || "");
        const revisionMonth = normalizeMonthKey(raw?.revision_month || raw?.revisionMonth || "");
        if (!tenancyId || !revisionMonth) return;
        map.set(`${tenancyId}__${revisionMonth}`, {
            reminder_id: raw.reminder_id || raw.reminderId || "",
            tenancy_id: tenancyId,
            revision_month: revisionMonth,
            last_reminder_month: normalizeMonthKey(raw.last_reminder_month || raw.lastReminderMonth || ""),
            send_count: Number(raw.send_count ?? raw.sendCount ?? 0) || 0,
            last_message_type: raw.last_message_type || raw.lastMessageType || "",
            last_message_lang: raw.last_message_lang || raw.lastMessageLang || "",
            last_message_text: raw.last_message_text || raw.lastMessageText || "",
            last_sent_at: raw.last_sent_at || raw.lastSentAt || "",
            last_sent_via: raw.last_sent_via || raw.lastSentVia || "",
            status: (raw.status || "pending").toString().toLowerCase(),
            decision_date: raw.decision_date || raw.decisionDate || "",
            decision_amount: raw.decision_amount ?? raw.decisionAmount ?? "",
            audit_log: raw.audit_log || raw.auditLog || "",
            created_at: raw.created_at || raw.createdAt || "",
            updated_at: raw.updated_at || raw.updatedAt || "",
        });
    });
    return map;
}

const MAX_AUDIT_LOG_ENTRIES = 25;

function parseAuditLog(raw) {
    if (!raw) return [];
    if (Array.isArray(raw)) return raw.filter(Boolean);
    if (typeof raw === "string") {
        try {
            const parsed = JSON.parse(raw);
            return Array.isArray(parsed) ? parsed.filter(Boolean) : [];
        } catch (err) {
            return [];
        }
    }
    return [];
}

function buildAuditEntry({ messageType, messageLang, messageText, sentVia, sentBy }) {
    return {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        sentAt: new Date().toISOString(),
        messageType: messageType || "",
        messageLang: messageLang || "",
        messageText: messageText || "",
        sentVia: sentVia || "",
        sentBy: sentBy || "",
    };
}

function buildReminderPayload(item, overrides = {}) {
    const reminder = item?.reminder || {};
    return {
        reminder_id: reminder.reminder_id || reminder.reminderId || "",
        tenancyId: item?.tenancyId || "",
        revisionMonth: item?.nextRevisionMonth || "",
        lastReminderMonth:
            overrides.lastReminderMonth ??
            reminder.last_reminder_month ??
            reminder.lastReminderMonth ??
            "",
        sendCount: overrides.sendCount ?? reminder.send_count ?? reminder.sendCount ?? 0,
        lastMessageType:
            overrides.lastMessageType ??
            reminder.last_message_type ??
            reminder.lastMessageType ??
            "",
        lastMessageLang:
            overrides.lastMessageLang ??
            reminder.last_message_lang ??
            reminder.lastMessageLang ??
            "",
        lastMessageText:
            overrides.lastMessageText ??
            reminder.last_message_text ??
            reminder.lastMessageText ??
            "",
        lastSentAt: overrides.lastSentAt ?? reminder.last_sent_at ?? reminder.lastSentAt ?? "",
        lastSentVia: overrides.lastSentVia ?? reminder.last_sent_via ?? reminder.lastSentVia ?? "",
        status: (overrides.status ?? reminder.status ?? "pending").toString().toLowerCase(),
        decisionDate: overrides.decisionDate ?? reminder.decision_date ?? reminder.decisionDate ?? "",
        decisionAmount:
            overrides.decisionAmount ?? reminder.decision_amount ?? reminder.decisionAmount ?? "",
        auditLog: overrides.auditLog ?? reminder.audit_log ?? reminder.auditLog ?? "",
        createdAt: reminder.created_at ?? reminder.createdAt ?? "",
        updatedAt: overrides.updatedAt ?? reminder.updated_at ?? reminder.updatedAt ?? "",
    };
}

async function recordReminderSend(item, { lang, message, sentVia, sentBy } = {}) {
    if (!item) return;
    const reminder = item.reminder || {};
    const auditLog = parseAuditLog(reminder.audit_log || reminder.auditLog);
    const entry = buildAuditEntry({
        messageType: item.messageType,
        messageLang: lang,
        messageText: message,
        sentVia,
        sentBy,
    });
    const nextLog = [...auditLog, entry].slice(-MAX_AUDIT_LOG_ENTRIES);
    const nowIso = new Date().toISOString();
    const payload = buildReminderPayload(item, {
        lastReminderMonth: getCurrentMonthKey(),
        sendCount: (Number(reminder.send_count ?? reminder.sendCount ?? 0) || 0) + 1,
        lastMessageType: item.messageType,
        lastMessageLang: lang,
        lastMessageText: message,
        lastSentAt: nowIso,
        lastSentVia: sentVia || "whatsapp",
        status: "pending",
        auditLog: JSON.stringify(nextLog),
        updatedAt: nowIso,
    });
    await saveRentRevisionReminder(payload);
}

function getLatestRevisionEntry(list, currentMonthKey) {
    if (!Array.isArray(list) || !list.length) return null;
    if (!currentMonthKey) return list[list.length - 1] || null;
    const filtered = list.filter((entry) => entry.monthKey <= currentMonthKey);
    return filtered.length ? filtered[filtered.length - 1] : null;
}

function getRevisionEntryForMonth(list, monthKey) {
    if (!Array.isArray(list) || !list.length || !monthKey) return null;
    return list.find((entry) => entry.monthKey === monthKey) || null;
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

function resolveGrnNumber(tenancy, tenant) {
    return (
        tenancy?.grn_number ||
        tenancy?.grnNumber ||
        tenant?.grnNumber ||
        tenant?.grn_number ||
        tenant?.templateData?.["GRN number"] ||
        tenant?.templateData?.grn_number ||
        ""
    )
        .toString()
        .trim();
}

function formatDateLabel(date) {
    if (!date || !(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    return date.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
}

function getDebugMonthKey() {
    const raw = localStorage.getItem(DEBUG_MONTH_KEY) || "";
    const normalized = normalizeMonthKey(raw);
    return /^\d{4}-\d{2}$/.test(normalized) ? normalized : "";
}

function getNotificationToday() {
    const debugKey = getDebugMonthKey();
    if (debugKey) {
        const [year, month] = debugKey.split("-").map((part) => parseInt(part, 10));
        const date = new Date(year, month - 1, 1);
        if (!Number.isNaN(date.getTime())) return date;
    }
    return new Date();
}

function getCurrentMonthKey() {
    return monthKeyFromDate(getNotificationToday());
}

function applyReminderSentState(item, { lang } = {}) {
    const sentMonth = getCurrentMonthKey();
    revisionNotificationState.items = revisionNotificationState.items.map((entry) => {
        if (entry.id !== item.id) return entry;
        return {
            ...entry,
            lastReminderMonth: sentMonth,
            alreadySentThisMonth: true,
            canSend: false,
            lastMessageType: entry.messageType,
            lastMessageLang: lang || entry.lastMessageLang || "en",
        };
    });
}

function resolveTenantMobile(tenant) {
    return (
        tenant?.tenantMobile ||
        tenant?.tenant_mobile ||
        tenant?.mobile ||
        tenant?.templateData?.tenant_mobile ||
        tenant?.templateData?.tenantMobile ||
        ""
    );
}

function getPrimaryMobile(value) {
    if (Array.isArray(value)) {
        return String(value.find(Boolean) ?? "");
    }
    if (value === null || value === undefined) {
        return "";
    }
    return String(value);
}

function buildItems(tenancies, tenants, units, revisions, reminders) {
    const { byId, byTenancy } = buildTenantMaps(tenants);
    const unitMap = buildUnitMap(units);
    const revisionMap = buildRevisionMap(revisions);
    const reminderMap = buildReminderMap(reminders);
    const today = getNotificationToday();
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

        const noticeMonths =
            parseNumber(
                tenancy?.tenant_notice_months ?? tenant?.tenantNoticeMonths ?? tenant?.templateData?.notice_num_t,
                0
            ) || 0;
        const scheduledRevisionDate = getNextRevisionDate(commencementDate, interval.intervalMonths, today);
        const nextRevisionMonth = monthKeyFromDate(scheduledRevisionDate);
        if (!nextRevisionMonth) return;

        const revisionsForTenancy = revisionMap.get(tenancyId) || [];
        const nextRevisionEntry = getRevisionEntryForMonth(revisionsForTenancy, nextRevisionMonth);

        const nextRevisionDate = scheduledRevisionDate;
        if (!nextRevisionDate) return;
        const visibleFromDate = addMonths(nextRevisionDate, -(noticeMonths + 2));
        const visibleFromMonth = monthKeyFromDate(visibleFromDate);
        const sendFromDate = addMonths(nextRevisionDate, -(noticeMonths + 1));
        const sendFromMonth = monthKeyFromDate(sendFromDate);
        const sendUntilDate = nextRevisionDate;
        const sendUntilMonth = monthKeyFromDate(sendUntilDate);

        if (!currentMonthKey || !visibleFromMonth) return;
        if (currentMonthKey < visibleFromMonth) return;

        const reminderKey = `${tenancyId}__${nextRevisionMonth}`;
        const reminder = reminderMap.get(reminderKey) || null;
        const reminderStatus = (reminder?.status || "pending").toString().toLowerCase();
        if (reminderStatus === "approved" || reminderStatus === "rejected") return;

        const approvalMonth = monthKeyFromDate(addMonths(nextRevisionDate, 1));
        const isApprovalMonth = currentMonthKey === approvalMonth;
        const isRevisionMonth = currentMonthKey === nextRevisionMonth;
        const lastReminderMonth = normalizeMonthKey(
            reminder?.last_reminder_month || reminder?.lastReminderMonth || ""
        );
        const alreadySentThisMonth =
            !!lastReminderMonth && !!currentMonthKey && lastReminderMonth === currentMonthKey;
        const canSend =
            !isApprovalMonth &&
            !alreadySentThisMonth &&
            currentMonthKey >= sendFromMonth &&
            currentMonthKey <= sendUntilMonth;
        const messageType = isRevisionMonth ? "effective-month" : "regular";

        const latestRevisionEntry = getLatestRevisionEntry(revisionsForTenancy, currentMonthKey);
        const currentRent =
            latestRevisionEntry?.rentAmount ??
            parseNumber(
                tenant?.currentRent ??
                    tenant?.rentAmount ??
                    tenant?.templateData?.rent_amount ??
                    tenancy?.rent_amount ??
                    null,
                null
            );
        const configuredIncrease = resolveIncreaseAmount(tenancy, tenant);
        let revisedRent = nextRevisionEntry?.rentAmount ?? null;
        if (revisedRent === null && currentRent !== null && configuredIncrease !== null) {
            revisedRent = currentRent + configuredIncrease;
        }
        const increaseAmount =
            currentRent !== null && revisedRent !== null ? revisedRent - currentRent : configuredIncrease;

        items.push({
            id: tenancyId ? `${tenancyId}__${nextRevisionMonth}` : getItemId({ tenancyId, nextRevisionMonth }),
            tenancyId,
            tenantName:
                tenant?.tenantFullName ||
                tenant?.tenantName ||
                tenant?.tenant_name ||
                tenant?.full_name ||
                tenant?.templateData?.Tenant_Full_Name ||
                "Unknown tenant",
            mobile: resolveTenantMobile(tenant),
            unitLabel: resolveUnitLabel(tenancy, tenant, unitMap) || "-",
            grnNumber: resolveGrnNumber(tenancy, tenant),
            commencementDate,
            commencementMonth: monthKeyFromDate(commencementDate),
            interval,
            noticeMonths,
            nextRevisionDate,
            nextRevisionMonth,
            revisionNote: nextRevisionEntry?.note || "",
            visibleFromMonth,
            sendFromMonth,
            sendUntilMonth,
            canSend,
            isApprovalMonth,
            isRevisionMonth,
            messageType,
            reminder,
            reminderStatus,
            lastReminderMonth,
            alreadySentThisMonth,
            sendCount: reminder?.send_count || 0,
            lastMessageType: reminder?.last_message_type || "",
            currentRent,
            revisedRent,
            increaseAmount,
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

    items.forEach((item) => {
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

function getStatusMeta(item) {
    if ((item.reminderStatus || "") === "approved") {
        return {
            label: "Approved",
            className: "border-emerald-200 bg-emerald-50 text-emerald-700",
        };
    }
    if ((item.reminderStatus || "") === "rejected") {
        return {
            label: "Rejected",
            className: "border-rose-200 bg-rose-50 text-rose-700",
        };
    }
    if (item.isApprovalMonth) {
        return {
            label: "Pending approval",
            className: "border-amber-200 bg-amber-50 text-amber-700",
        };
    }
    if (item.alreadySentThisMonth || item.lastReminderMonth) {
        return {
            label: "Reminder sent",
            className: "border-emerald-200 bg-emerald-50 text-emerald-700",
        };
    }
    if (!item.canSend) {
        return {
            label: "Not sendable yet",
            className: "border-slate-200 bg-slate-50 text-slate-600",
        };
    }
    return {
        label: "Pending reminder",
        className: "border-indigo-200 bg-indigo-50 text-indigo-700",
    };
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
        const rawId = item.id || "";
        const safeId = escapeHtml(rawId);
        const token = toSafeToken(rawId);
        const tenantName = escapeHtml(item.tenantName);
        const unitLabel = escapeHtml(item.unitLabel);
        const commencementLabel = escapeHtml(formatMonthLabel(item.commencementMonth));
        const nextLabel = escapeHtml(formatMonthLabel(item.nextRevisionMonth));
        const visibleLabel = escapeHtml(formatMonthLabel(item.visibleFromMonth));
        const sendFromLabel = escapeHtml(formatMonthLabel(item.sendFromMonth));
        const sendUntilLabel = escapeHtml(formatMonthLabel(item.sendUntilMonth));
        const intervalLabel = escapeHtml(formatIntervalLabel(item.interval));
        const noticeLabel = escapeHtml(`${item.noticeMonths || 0} mo`);
        const rentBefore = escapeHtml(formatMoney(item.currentRent));
        const revisedRent = escapeHtml(formatMoney(item.revisedRent));
        const noteLabel = escapeHtml(item.revisionNote || "-");
        const statusMeta = getStatusMeta(item);
        const manualLabel = item.alreadySentThisMonth
            ? "Sent"
            : item.lastReminderMonth
                ? "Resend"
                : "Send Manual";

        const isLoggedIn = localStorage.getItem("wa_logged_in") === "true";
        const canSend = item.canSend;
        const isApprovalMonth = item.isApprovalMonth;

        let persisted = revisionNotificationState.selections.get(item.id);
        if (!persisted) {
            persisted = { selected: isLoggedIn && canSend, lang: "en" };
            revisionNotificationState.selections.set(item.id, persisted);
        }
        const isSelected = persisted.selected && isLoggedIn && canSend;
        const disabled = !canSend || isApprovalMonth;
        const lang = persisted.lang || "en";

        const row = document.createElement("tr");
        const shouldDimRow = item.alreadySentThisMonth || (!canSend && !isApprovalMonth);
        row.className = `border-b last:border-0 transition-colors ${
            shouldDimRow ? "bg-slate-50 opacity-70" : "hover:bg-slate-50"
        }`;
        row.dataset.langToken = token;
        row.innerHTML = `
            <td class="px-3 py-2 w-8 text-center align-middle">
                ${
                    canSend
                        ? `<input type="checkbox" class="revision-send-check h-3.5 w-3.5 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                    data-id="${safeId}" data-can-send="${canSend ? "1" : "0"}" ${isSelected ? "checked" : ""} ${disabled ? "disabled" : ""}>`
                        : ""
                }
            </td>
            <td class="px-3 py-2 align-middle">
                <div class="text-[12px] font-semibold text-slate-800">${tenantName}</div>
                <div class="text-[9px] text-slate-500">Every ${intervalLabel} | Notice ${noticeLabel}</div>
                <div class="mt-1">
                    <span class="inline-flex items-center px-2 py-0.5 rounded-full border text-[9px] font-semibold ${statusMeta.className}">
                        ${statusMeta.label}
                    </span>
                </div>
            </td>
            <td class="px-3 py-2 text-[11px] text-slate-700 align-middle">${unitLabel}</td>
            <td class="px-3 py-2 text-[10px] text-slate-600 align-middle">${commencementLabel}</td>
            <td class="px-3 py-2 align-middle">
                <div class="text-[11px] font-semibold text-rose-700">${nextLabel}</div>
                <div class="text-[9px] text-slate-500">${rentBefore}</div>
            </td>
            <td class="px-3 py-2 text-[11px] text-slate-700 align-middle">${revisedRent}</td>
            <td class="px-3 py-2 text-[10px] text-slate-500 align-middle">${noteLabel}</td>
            <td class="px-3 py-2 text-[10px] text-slate-600 align-middle">${visibleLabel}</td>
            <td class="px-3 py-2 text-[10px] text-slate-600 align-middle">${sendFromLabel} - ${sendUntilLabel}</td>
            <td class="px-3 py-2 text-center align-middle">
                <div class="flex items-center justify-center gap-1">
                    <input type="radio" name="rev_lang_${token}" value="hi"
                        class="revision-lang-radio h-3 w-3 border-gray-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                        data-id="${safeId}" ${lang === "hi" ? "checked" : ""} ${disabled ? "disabled" : ""}>
                </div>
            </td>
            <td class="px-3 py-2 text-center align-middle">
                <div class="flex items-center justify-center gap-1">
                    <input type="radio" name="rev_lang_${token}" value="en"
                        class="revision-lang-radio h-3 w-3 border-gray-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                        data-id="${safeId}" ${lang !== "hi" ? "checked" : ""} ${disabled ? "disabled" : ""}>
                </div>
            </td>
            <td class="px-3 py-2 text-right align-middle">
                <div class="inline-flex items-center justify-end gap-2">
                    ${
                        isApprovalMonth
                            ? `
                    <button class="rev-approve-btn inline-flex items-center gap-1 px-3 py-1.5 rounded bg-emerald-600 text-white border border-emerald-600 text-[11px] font-semibold hover:bg-emerald-500 shadow-sm"
                        data-id="${safeId}">
                        <span>Approve</span>
                    </button>
                    <button class="rev-reject-btn inline-flex items-center gap-1 px-3 py-1.5 rounded bg-rose-600 text-white border border-rose-600 text-[11px] font-semibold hover:bg-rose-500 shadow-sm"
                        data-id="${safeId}">
                        <span>Reject</span>
                    </button>
                            `
                            : `
                    <button class="rev-send-copy p-1.5 rounded text-slate-400 hover:text-slate-600 hover:bg-slate-100 ${canSend ? "" : "opacity-40 cursor-not-allowed"}"
                        title="Copy message" data-id="${safeId}" ${canSend ? "" : "disabled"}>
                        <svg class="w-3.5 h-3.5 pointer-events-none" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
                    </button>
                    <button class="rev-send-btn inline-flex items-center gap-1 px-3 py-1.5 rounded bg-[#25D366] text-white border border-[#20ba5a] text-[12px] font-semibold hover:bg-[#20ba5a] shadow-sm ${canSend ? "" : "opacity-40 cursor-not-allowed"}"
                        data-id="${safeId}" ${canSend ? "" : "disabled"}>
                        <span>${manualLabel}</span>
                    </button>
                            `
                    }
                </div>
            </td>
        `;
        modalList.appendChild(row);
    });

    bindSendListEvents();
    bindCheckboxEvents();
    bindLangEvents();
    updateAutoSendButton();
}

function getSendableCheckboxes() {
    return Array.from(document.querySelectorAll(".revision-send-check")).filter((cb) => !cb.disabled);
}

function bindCheckboxEvents() {
    document.querySelectorAll(".revision-send-check").forEach((cb) => {
        if (cb.dataset.bound === "1") return;
        cb.dataset.bound = "1";
        cb.addEventListener("change", (event) => {
            const target = event.target;
            if (!(target instanceof HTMLInputElement)) return;
            const id = target.dataset.id || "";
            if (!id) return;
            const current = revisionNotificationState.selections.get(id) || { selected: false, lang: "en" };
            current.selected = target.checked;
            if (!current.lang) current.lang = "en";
            revisionNotificationState.selections.set(id, current);
            updateAutoSendButton();
        });
    });
}

function bindLangEvents() {
    document.querySelectorAll(".revision-lang-radio").forEach((radio) => {
        if (radio.dataset.bound === "1") return;
        radio.dataset.bound = "1";
        radio.addEventListener("change", (event) => {
            const target = event.target;
            if (!(target instanceof HTMLInputElement)) return;
            if (!target.checked) return;
            const id = target.dataset.id || "";
            if (!id) return;
            const current = revisionNotificationState.selections.get(id) || { selected: false, lang: "en" };
            current.lang = target.value || "en";
            revisionNotificationState.selections.set(id, current);
        });
    });
}

function formatWhatsappMessage(item, lang = "en") {
    const monthLabel = formatMonthLabel(item.nextRevisionMonth);
    let billMonthLabel = "-";
    if (/^\d{4}-\d{2}$/.test(item.nextRevisionMonth || "")) {
        const [year, month] = item.nextRevisionMonth.split("-").map((part) => parseInt(part, 10));
        const date = new Date(year, month - 1, 1);
        const nextDate = addMonths(date, 1);
        billMonthLabel = formatMonthLabel(monthKeyFromDate(nextDate));
    }
    const currentRentValue = item.currentRent;
    const revisedRentValue = item.revisedRent;
    const currentRent = formatMoney(currentRentValue);
    const revisedRent = formatMoney(revisedRentValue);
    const hasAmounts = currentRentValue !== null && revisedRentValue !== null;
    const name = item.tenantName || "Tenant";
    const unit = item.unitLabel || "Unit";
    const grn = item.grnNumber || "";
    const commencementLabel = formatDateLabel(item.commencementDate) || formatMonthLabel(item.commencementMonth);
    if (lang === "hi") {
        const grnText = grn ? `समझौता संख्या ${grn}` : "समझौते के अनुसार";
        const commencementText = commencementLabel ? ` और आरंभ तिथि ${commencementLabel}` : "";
        const lines = [`नमस्ते ${name},`];
        if (hasAmounts) {
            if (item.messageType === "effective-month") {
                lines.push(
                    `${grnText}${commencementText} के अनुसार, ${unit} के लिए आपका किराया इस महीने ${currentRent} से ${revisedRent} हो जाएगा। यह अगले महीने (${billMonthLabel}) के बिल में दिखेगा।`
                );
            } else {
                lines.push(
                    `${grnText}${commencementText} के अनुसार, ${unit} के लिए आपका किराया ${currentRent} से ${revisedRent} होकर ${monthLabel} महीने से लागू होगा।`
                );
            }
        } else {
            if (item.messageType === "effective-month") {
                lines.push(
                    `${grnText}${commencementText} के अनुसार, ${unit} के लिए आपका किराया इस महीने से संशोधित होगा। यह अगले महीने (${billMonthLabel}) के बिल में दिखेगा।`
                );
            } else {
                lines.push(
                    `${grnText}${commencementText} के अनुसार, ${unit} के लिए आपका किराया ${monthLabel} महीने से संशोधित होगा।`
                );
            }
        }
        if (item.revisionNote) lines.push(`नोट: ${item.revisionNote}`);
        lines.push("किसी भी प्रश्न के लिए संपर्क करें।");
        return lines.join("\n");
    }

    const grnText = grn ? `agreement No ${grn}` : "the agreement";
    const commencementText = commencementLabel ? ` and commencement date ${commencementLabel}` : "";
    const lines = [`Hi ${name},`];
    if (hasAmounts) {
        if (item.messageType === "effective-month") {
            lines.push(
                `As per ${grnText}${commencementText}, your rent for ${unit} will be increased this month from ${currentRent} to ${revisedRent}. This will be reflected in next month's bill (${billMonthLabel}).`
            );
        } else {
            lines.push(
                `As per ${grnText}${commencementText}, your rent will be increased from ${currentRent} to ${revisedRent} for ${unit} in the month of ${monthLabel}.`
            );
        }
    } else if (item.messageType === "effective-month") {
        lines.push(
            `As per ${grnText}${commencementText}, your rent for ${unit} will be revised from this month. This will be reflected in next month's bill (${billMonthLabel}).`
        );
    } else {
        lines.push(`As per ${grnText}${commencementText}, your rent for ${unit} will be revised in ${monthLabel}.`);
    }
    if (item.revisionNote) lines.push(`Note: ${item.revisionNote}`);
    lines.push("Please contact us for any questions.");
    return lines.join("\n");
}

async function openWhatsappExternally(url) {
    try {
        const shell = window.__TAURI__?.shell;
        if (shell?.open) {
            await shell.open(url);
            return true;
        }
    } catch (err) {
        console.error("Unable to open WhatsApp in default browser", err);
    }

    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.target = "_blank";
    anchor.rel = "noopener noreferrer";
    anchor.style.display = "none";
    document.body.appendChild(anchor);
    anchor.click();
    requestAnimationFrame(() => anchor.remove());
    return true;
}

function copyToClipboard(text) {
    if (!navigator.clipboard) {
        const ta = document.createElement("textarea");
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        return Promise.resolve();
    }
    return navigator.clipboard.writeText(text);
}

function findItemById(id) {
    return revisionNotificationState.items.find((item) => item.id === id);
}

function bindSendListEvents() {
    const { modalList } = getElements();
    if (!modalList || modalList.dataset.bound === "true") return;
    modalList.dataset.bound = "true";

    modalList.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;

        const approveBtn = target.closest(".rev-approve-btn");
        if (approveBtn) {
            const id = approveBtn.dataset.id || "";
            const item = findItemById(id);
            if (!item) return;
            handleApproveAction(item);
            return;
        }

        const rejectBtn = target.closest(".rev-reject-btn");
        if (rejectBtn) {
            const id = rejectBtn.dataset.id || "";
            const item = findItemById(id);
            if (!item) return;
            openDecisionModal("reject", item);
            return;
        }

        const copyBtn = target.closest(".rev-send-copy");
        if (copyBtn) {
            const id = copyBtn.dataset.id || "";
            const item = findItemById(id);
            if (!item) return;
            if (!item.canSend) {
                if (item.alreadySentThisMonth) {
                    showToast("Reminder already sent for this month.", "warning");
                } else {
                    showToast("Sending is available closer to the revision month.", "warning");
                }
                return;
            }
            const row = copyBtn.closest("tr");
            const token = row?.dataset?.langToken || "";
            const radio = token ? row.querySelector(`input[name="rev_lang_${token}"]:checked`) : null;
            const lang = radio ? radio.value : "en";
            const message = formatWhatsappMessage(item, lang);
            copyToClipboard(message)
                .then(() => showToast("Message copied to clipboard", "success"))
                .catch((err) => {
                    console.error("Copy failed", err);
                    showToast("Failed to copy message", "error");
                });
            return;
        }

        const sendBtn = target.closest(".rev-send-btn");
        if (!sendBtn) return;
        const id = sendBtn.dataset.id || "";
        const item = findItemById(id);
        if (!item) return;
        if (!item.canSend) {
            if (item.alreadySentThisMonth) {
                showToast("Reminder already sent for this month.", "warning");
            } else {
                showToast("Sending is available closer to the revision month.", "warning");
            }
            return;
        }
        const number = getPrimaryMobile(item.mobile).replace(/\D/g, "");
        if (!number) {
            showToast(`Missing mobile number for ${item.tenantName}`, "error");
            return;
        }
        const row = sendBtn.closest("tr");
        const token = row?.dataset?.langToken || "";
        const radio = token ? row.querySelector(`input[name="rev_lang_${token}"]:checked`) : null;
        const lang = radio ? radio.value : "en";
        const message = formatWhatsappMessage(item, lang);
        const url = `https://web.whatsapp.com/send?phone=${number}&text=${encodeURIComponent(message)}&app_absent=0`;
        openWhatsappExternally(url).then(async () => {
            try {
                await recordReminderSend(item, {
                    lang,
                    message,
                    sentVia: "whatsapp",
                    sentBy: "manual",
                });
                const current = revisionNotificationState.selections.get(id) || { selected: false, lang };
                current.selected = false;
                current.lang = lang || current.lang || "en";
                revisionNotificationState.selections.set(id, current);
                applyReminderSentState(item, { lang });
                renderModal(revisionNotificationState.items);
            } catch (err) {
                console.error("Failed to save reminder send", err);
                showToast("Failed to save reminder status", "warning");
            }
            revisionNotificationState.sendStatus.set(item.id, true);
        });
    });
}

function updateAutoSendButton(forceState = null) {
    const { autoSendBtn, autoSendCount, toggleAll } = getElements();
    if (!autoSendBtn) return;

    const isLoggedIn =
        forceState !== null ? forceState : localStorage.getItem("wa_logged_in") === "true";

    const checkboxes = Array.from(document.querySelectorAll(".revision-send-check"));

    if (!isLoggedIn) {
        checkboxes.forEach((cb) => {
            cb.disabled = true;
            cb.checked = false;
            const id = cb.dataset.id || "";
            if (id) {
                const current = revisionNotificationState.selections.get(id) || { selected: false, lang: "en" };
                current.selected = false;
                if (!current.lang) current.lang = "en";
                revisionNotificationState.selections.set(id, current);
            }
        });
        if (toggleAll) {
            toggleAll.disabled = true;
            toggleAll.checked = false;
        }
        autoSendBtn.innerHTML = "<span>To send Sign in to whatsapp</span>";
        autoSendBtn.className =
            "px-6 py-2.5 rounded bg-red-400 text-white text-[12px] font-semibold hover:bg-red-500 shadow-sm flex items-center gap-2 transform active:scale-95 transition-all";
        autoSendBtn.disabled = false;
        autoSendBtn.dataset.action = "login";
        if (autoSendCount) autoSendCount.classList.add("hidden");
        return;
    }

    const sendable = getSendableCheckboxes();

    checkboxes.forEach((cb) => {
        const canSend = cb.dataset.canSend === "1";
        cb.disabled = !canSend;
        const id = cb.dataset.id || "";
        if (!id) return;
        let entry = revisionNotificationState.selections.get(id);
        if (!entry) {
            entry = { selected: canSend, lang: "en" };
            revisionNotificationState.selections.set(id, entry);
        }
        if (!entry.lang) entry.lang = "en";
        if (!canSend) {
            entry.selected = false;
            cb.checked = false;
        } else {
            cb.checked = entry.selected;
        }
    });

    let selectedCount = sendable.filter((cb) => cb.checked).length;

    if (toggleAll) {
        toggleAll.disabled = false;
        toggleAll.checked = sendable.length > 0 && sendable.every((cb) => cb.checked);
    }
    autoSendBtn.dataset.action = "send";
    autoSendBtn.className =
        "px-6 py-2.5 rounded bg-emerald-600 text-white text-[12px] font-semibold hover:bg-emerald-500 shadow-sm flex items-center gap-2 transform active:scale-95 transition-all";
    autoSendBtn.innerHTML = `<span>Send with WhatsApp</span> <span id="revisionNotificationAutoSendCount" class="bg-white/20 px-1.5 rounded text-[10px] ${selectedCount === 0 ? "hidden" : ""}">${selectedCount}</span>`;
    if (selectedCount === 0) {
        autoSendBtn.classList.add("opacity-50", "cursor-not-allowed");
        autoSendBtn.disabled = true;
    } else {
        autoSendBtn.classList.remove("opacity-50", "cursor-not-allowed");
        autoSendBtn.disabled = false;
    }

    const countEl = document.getElementById("revisionNotificationAutoSendCount");
    if (countEl) {
        countEl.textContent = `${selectedCount}`;
        countEl.classList.toggle("hidden", selectedCount === 0);
    }
}

async function handleAutoSend() {
    const { autoSendBtn } = getElements();
    if (autoSendBtn && autoSendBtn.dataset.action === "login") {
        try {
            if (window.__TAURI__) {
                await window.__TAURI__.core.invoke("open_whatsapp");
            } else {
                alert("Cannot open WhatsApp in browser mode");
            }
        } catch (err) {
            console.error("Failed to open WhatsApp login", err);
            showToast("Failed to open WhatsApp window", "error");
        }
        return;
    }

    const checkboxes = Array.from(document.querySelectorAll(".revision-send-check:checked"));
    if (!checkboxes.length) return;

    if (autoSendBtn) {
        autoSendBtn.textContent = "Sending...";
        autoSendBtn.disabled = true;
    }

    revisionNotificationState.suspendReload = true;
    revisionNotificationState.pendingReload = false;

    let sentCount = 0;
    const total = checkboxes.length;

    for (const [index, checkbox] of checkboxes.entries()) {
        const id = checkbox.dataset.id || "";
        const item = findItemById(id);
        if (!item) continue;
        const number = getPrimaryMobile(item.mobile).replace(/\D/g, "");
        if (!number) {
            showToast(`Skipped ${item.tenantName}: No mobile number`, "error");
            continue;
        }
        const row = checkbox.closest("tr");
        const token = row?.dataset?.langToken || "";
        const radio = token ? row.querySelector(`input[name="rev_lang_${token}"]:checked`) : null;
        const lang = radio ? radio.value : "en";
        const message = formatWhatsappMessage(item, lang);
        try {
            if (window.__TAURI__) {
                const progressLabel = `${index + 1}/${total} Sending`;
                await window.__TAURI__.core.invoke("send_whatsapp_message", {
                    phone: number,
                    message,
                    progressLabel,
                });
                try {
                    await recordReminderSend(item, {
                        lang,
                        message,
                        sentVia: "whatsapp",
                        sentBy: "auto",
                    });
                    applyReminderSentState(item, { lang });
                } catch (err) {
                    console.error("Failed to save reminder send", err);
                    showToast("Failed to save reminder status", "warning");
                }
                revisionNotificationState.sendStatus.set(item.id, true);
                checkbox.checked = false;
                checkbox.disabled = true;
                const current = revisionNotificationState.selections.get(id) || { selected: false };
                current.selected = false;
                revisionNotificationState.selections.set(id, current);

                const row = checkbox.closest("tr");
                if (row) {
                    row.classList.add("bg-emerald-50");
                    const cell = row.querySelector("td:first-child");
                    if (cell) {
                        cell.innerHTML = `<span class="text-emerald-600 font-bold text-[10px]">OK</span>`;
                    }
                }
                sentCount += 1;
            } else {
                showToast("Auto-send unavailable in browser mode", "error");
                break;
            }
        } catch (err) {
            console.error(`Failed to auto-send for ${item.tenantName}`, err);
            showToast(`Failed to send to ${item.tenantName}`, "error");
        }
    }

    revisionNotificationState.suspendReload = false;
    if (revisionNotificationState.pendingReload) {
        revisionNotificationState.pendingReload = false;
        loadRevisionNotifications();
    }

    if (autoSendBtn) {
        autoSendBtn.disabled = false;
        updateAutoSendButton();
    }

    showToast(`Auto-send completed. ${sentCount}/${total} sent.`, "success");
}

function handleReminderRefresh() {
    if (revisionNotificationState.suspendReload) {
        revisionNotificationState.pendingReload = true;
        return;
    }
    loadRevisionNotifications();
}

async function loadRevisionNotifications() {
    if (revisionNotificationState.loading) {
        revisionNotificationState.pendingReload = true;
        return;
    }
    revisionNotificationState.loading = true;
    const { widgetEmpty } = getElements();
    if (widgetEmpty) {
        widgetEmpty.textContent = "Loading revisions...";
        widgetEmpty.classList.remove("hidden");
    }

    try {
        const [tenancies, tenants, units, revisions, reminders] = await Promise.all([
            getLocalList(LOCAL_KEYS.tenancies, []),
            getLocalList(LOCAL_KEYS.tenants, []),
            getLocalList(LOCAL_KEYS.units, []),
            getLocalList(LOCAL_KEYS.rentRevisionsAll, []),
            getLocalList(LOCAL_KEYS.rentRevisionReminders, []),
        ]);
        revisionNotificationState.items = buildItems(
            Array.isArray(tenancies) ? tenancies : [],
            Array.isArray(tenants) ? tenants : [],
            Array.isArray(units) ? units : [],
            Array.isArray(revisions) ? revisions : [],
            Array.isArray(reminders) ? reminders : []
        );
    } catch (err) {
        console.error("Failed to load revision notifications", err);
        revisionNotificationState.items = [];
    }

    renderWidget(revisionNotificationState.items);
    renderModal(revisionNotificationState.items);
    revisionNotificationState.loading = false;
    if (revisionNotificationState.pendingReload && !revisionNotificationState.suspendReload) {
        revisionNotificationState.pendingReload = false;
        loadRevisionNotifications();
    }
}

function openDecisionModal(action, item) {
    const { decisionModal, decisionTitle, decisionAmount, decisionConfirm } = getElements();
    if (!decisionModal || !decisionTitle || !decisionAmount || !decisionConfirm) return;
    const amount = parseNumber(item?.revisedRent ?? item?.currentRent ?? null, null);
    decisionTitle.textContent =
        action === "approve" ? "Approve rent revision" : "Reject rent revision";
    decisionConfirm.textContent = action === "approve" ? "Approve" : "Reject";
    decisionAmount.value = amount !== null ? String(amount) : "";
    decisionModal.dataset.action = action || "";
    decisionModal.dataset.itemId = item?.id || "";
    revisionNotificationState.decisionContext = { action, itemId: item?.id || "" };
    showModal(decisionModal);
    decisionAmount.focus();
}

function closeDecisionModal() {
    const { decisionModal, decisionAmount } = getElements();
    if (decisionAmount) decisionAmount.value = "";
    if (decisionModal) {
        decisionModal.dataset.action = "";
        decisionModal.dataset.itemId = "";
        hideModal(decisionModal);
    }
    revisionNotificationState.decisionContext = null;
}

async function finalizeApproval(item, amount) {
    if (!item) return;
    const effectiveMonth = item.nextRevisionMonth;
    if (!effectiveMonth) {
        showToast("Missing revision month", "error");
        return;
    }
    const note = item.revisionNote || "Approved via revision notification";
    const result = await saveRentRevision({
        tenancyId: item.tenancyId,
        effectiveMonth,
        rentAmount: amount,
        note,
    });
    if (result?.ok === false) {
        showToast("Failed to approve revision", "error");
        return;
    }
    const nowIso = new Date().toISOString();
    await saveRentRevisionReminder(
        buildReminderPayload(item, {
            status: "approved",
            decisionDate: nowIso,
            decisionAmount: amount,
            updatedAt: nowIso,
        })
    );
    showToast("Revision approved", "success");
}

async function finalizeRejection(item, amount) {
    if (!item) return;
    const nowIso = new Date().toISOString();
    await saveRentRevisionReminder(
        buildReminderPayload(item, {
            status: "rejected",
            decisionDate: nowIso,
            decisionAmount: amount,
            updatedAt: nowIso,
        })
    );
    showToast("Revision rejected", "success");
}

function handleApproveAction(item) {
    if (!item) return;
    if (!item.isApprovalMonth) {
        showToast("Approval is available in the revision month only.", "warning");
        return;
    }
    const amount = parseNumber(item.revisedRent, null);
    if (amount === null) {
        openDecisionModal("approve", item);
        return;
    }
    finalizeApproval(item, amount);
}

async function handleDecisionConfirm() {
    const { decisionAmount } = getElements();
    const context = revisionNotificationState.decisionContext;
    if (!context || !decisionAmount) return;
    const item = findItemById(context.itemId);
    if (!item) {
        closeDecisionModal();
        return;
    }
    const amount = parseNumber(decisionAmount.value, null);
    if (amount === null || amount <= 0) {
        showToast("Enter a valid rent amount", "warning");
        return;
    }
    if (context.action === "approve") {
        await finalizeApproval(item, amount);
    } else {
        await finalizeRejection(item, amount);
    }
    closeDecisionModal();
}

function bindModalEvents() {
    const {
        modal,
        modalClose,
        expandBtn,
        toggleAll,
        autoSendBtn,
        decisionModal,
        decisionClose,
        decisionCancel,
        decisionConfirm,
    } = getElements();
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

    if (decisionModal && !decisionModal.dataset.bound) {
        decisionModal.dataset.bound = "1";
        decisionModal.addEventListener("click", (event) => {
            if (event.target === decisionModal) closeDecisionModal();
        });
    }

    if (decisionClose && !decisionClose.dataset.bound) {
        decisionClose.dataset.bound = "1";
        decisionClose.addEventListener("click", closeDecisionModal);
    }

    if (decisionCancel && !decisionCancel.dataset.bound) {
        decisionCancel.dataset.bound = "1";
        decisionCancel.addEventListener("click", closeDecisionModal);
    }

    if (decisionConfirm && !decisionConfirm.dataset.bound) {
        decisionConfirm.dataset.bound = "1";
        decisionConfirm.addEventListener("click", handleDecisionConfirm);
    }

    if (toggleAll && !toggleAll.dataset.bound) {
        toggleAll.dataset.bound = "1";
        toggleAll.addEventListener("change", (event) => {
            const checked = event.target.checked;
            getSendableCheckboxes().forEach((cb) => {
                cb.checked = checked;
                const id = cb.dataset.id || "";
                if (!id) return;
                const current = revisionNotificationState.selections.get(id) || { selected: false, lang: "en" };
                current.selected = checked;
                if (!current.lang) current.lang = "en";
                revisionNotificationState.selections.set(id, current);
            });
            updateAutoSendButton();
        });
    }

    if (autoSendBtn && !autoSendBtn.dataset.bound) {
        autoSendBtn.dataset.bound = "1";
        autoSendBtn.addEventListener("click", handleAutoSend);
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
    document.addEventListener("rentRevisionReminders:updated", handleReminderRefresh);
    document.addEventListener("app:whatsapp-status-change", (e) => {
        if (e.detail && typeof e.detail.loggedIn === "boolean") {
            updateAutoSendButton(e.detail.loggedIn);
        }
    });

    window.addEventListener("focus", () => {
        setTimeout(() => updateAutoSendButton(null), 400);
    });

    loadRevisionNotifications();
}



