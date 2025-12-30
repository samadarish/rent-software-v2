/**
 * Payments Feature Module
 *
 * Displays recorded payments, allows adding new receipts, and links
 * payments to generated bills and tenant directory context.
 */

import {
    fetchAttachmentPreview,
    fetchBillDetails,
    fetchBillsMinimal,
    fetchPayments,
    savePaymentRecord,
    uploadPaymentAttachment,
    deleteAttachment,
} from "../../api/sheets.js";
import { ensureTenantDirectoryLoaded, getActiveTenantsForWing } from "../tenants/tenants.js";
import { formatCurrency as formatCurrencyBase, normalizeMonthKey as normalizeMonthKeyBase } from "../../utils/formatters.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

const paymentsState = {
    items: [],
    loaded: false,
    generatedBills: {
        pending: [],
        paid: [],
    },
    billsLoaded: {
        pending: false,
        paid: false,
    },
    paymentIndex: createPaymentIndex(),
    activeBillTab: "pending",
    billContext: {
        monthKey: "",
        monthLabel: "",
        billTotal: 0,
        tenantName: "",
        tenantKey: "",
        wing: "",
        amount: 0,
        remaining: 0,
        rentAmount: 0,
        electricityAmount: 0,
        motorShare: 0,
        sweepAmount: 0,
        prevReading: "",
        newReading: "",
        payableDate: "",
    },
    attachmentDataUrl: "",
    attachmentName: "",
    attachmentUrl: "",
    attachmentPreviewUrl: "",
    attachmentViewUrl: "",
    attachmentId: "",
    attachmentCommitted: false,
    editingId: "",
    viewOnly: false,
    uploadingAttachment: false,
    uploadProgress: 0,
    uploadController: null,
    pendingPagination: {
        page: 1,
        pageSize: 9,
    },
    paidFilters: {
        fromMonth: "",
        toMonth: "",
        limit: 9,
        hasSearched: false,
    },
};

const BILL_MONTHS_BACK = 24;
const PENDING_PAGE_SIZE = 9;
const PAID_DEFAULT_LIMIT = 9;
const RECEIPT_MAX_DIM = 1100;
const RECEIPT_JPEG_QUALITY = 0.68;
const RECEIPT_TARGET_BYTES = 350 * 1024;
const RECEIPT_OUTPUT_MIME = "image/jpeg";

const attachmentPreviewCache = new Map();
let paymentHistoryLoading = false;
let paymentModalResetTimer = null;
let paymentModalRequestId = 0;

function getBillsForStatus(status = "pending") {
    return status === "paid" ? paymentsState.generatedBills.paid : paymentsState.generatedBills.pending;
}

function setBillsForStatus(status = "pending", bills = []) {
    if (status === "paid") {
        paymentsState.generatedBills.paid = bills;
        return;
    }
    paymentsState.generatedBills.pending = bills;
}

function isBillsLoaded(status = "pending") {
    return !!paymentsState.billsLoaded?.[status];
}

function setBillsLoaded(status = "pending", loaded) {
    if (!paymentsState.billsLoaded) {
        paymentsState.billsLoaded = { pending: false, paid: false };
    }
    paymentsState.billsLoaded[status] = !!loaded;
}

function getAllGeneratedBills() {
    return [...getBillsForStatus("pending"), ...getBillsForStatus("paid")];
}

function createPaymentIndex() {
    return {
        byTenantMonth: new Map(),
        byTenantMonthWing: new Map(),
    };
}

function buildPaymentIndex(payments = []) {
    const index = createPaymentIndex();

    payments.forEach((payment) => {
        const monthKey = normalizeMonthKey(payment.monthKey);
        const tenantKey = normalizeKey(payment.tenantKey || payment.tenantName);
        if (!monthKey || !tenantKey) return;

        const baseKey = `${monthKey}|${tenantKey}`;
        const wingKey = `${baseKey}|${normalizeKey(payment.wing)}`;

        if (!index.byTenantMonth.has(baseKey)) {
            index.byTenantMonth.set(baseKey, []);
        }
        index.byTenantMonth.get(baseKey).push(payment);

        if (!index.byTenantMonthWing.has(wingKey)) {
            index.byTenantMonthWing.set(wingKey, []);
        }
        index.byTenantMonthWing.get(wingKey).push(payment);
    });

    return index;
}

function getPaymentsForBill(bill, paymentIndex = paymentsState.paymentIndex) {
    const normalizedMonth = normalizeMonthKey(bill.monthKey);
    const normalizedTenant = normalizeKey(bill.tenantKey || bill.tenantName);
    const normalizedWing = normalizeKey(bill.wing);

    const baseKey = `${normalizedMonth}|${normalizedTenant}`;
    const wingKey = `${baseKey}|${normalizedWing}`;

    const matches = paymentIndex.byTenantMonth.get(baseKey) || [];
    const wingMatches = paymentIndex.byTenantMonthWing.get(wingKey) || [];

    return { matches, wingMatches };
}

function formatCurrency(amount) {
    const invalidValue = typeof amount === "string" && amount ? amount : "-";
    return formatCurrencyBase(amount, {
        emptyValue: "-",
        invalidValue,
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
        useGrouping: true,
    });
}

function normalizeBooleanValue(value) {
    if (value === true || value === false) return value;
    if (typeof value === "string") {
        return value.toLowerCase() !== "false" && value !== "";
    }
    return !!value;
}

function normalizeMonthRange(fromMonth, toMonth) {
    let from = normalizeMonthKey(fromMonth || "");
    let to = normalizeMonthKey(toMonth || "");
    if (from && to && from > to) {
        [from, to] = [to, from];
    }
    return { from, to };
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

function trimBillForList(bill = {}) {
    if (!bill || typeof bill !== "object") return bill;
    return {
        tenantName: bill.tenantName,
        tenantKey: bill.tenantKey,
        wing: bill.wing,
        unitNumber: bill.unitNumber,
        unit_number: bill.unit_number,
        monthKey: bill.monthKey,
        monthLabel: bill.monthLabel,
        totalAmount: bill.totalAmount,
        total_amount: bill.total_amount,
        remainingAmount: bill.remainingAmount,
        remaining_amount: bill.remaining_amount,
        amountPaid: bill.amountPaid,
        amount_paid: bill.amount_paid,
        isPaid: bill.isPaid,
        is_paid: bill.is_paid,
        hasPaid: bill.hasPaid,
        has_paid: bill.has_paid,
        payableDate: bill.payableDate,
        billLineId: bill.billLineId,
        tenancyId: bill.tenancyId,
    };
}

function filterPendingBills(bills = []) {
    return bills.filter((bill) => {
        const paidFlag = resolveBillPaidFlag(bill);
        return paidFlag === false || paidFlag === null;
    });
}

function applyPaidBillFilters(bills = [], filters = {}) {
    const { from, to } = normalizeMonthRange(filters.fromMonth, filters.toMonth);
    const limit = Number(filters.limit) || 0;

    const filtered = bills.filter((bill) => {
        if (resolveBillPaidFlag(bill) !== true) return false;
        const billMonth = normalizeMonthKey(bill.monthKey || bill.monthLabel);
        if (!billMonth) return false;
        if (from && billMonth < from) return false;
        if (to && billMonth > to) return false;
        return true;
    });

    if (!limit) return filtered;
    return filtered.slice(0, limit);
}

function clampPendingPage(totalItems) {
    const pageSize = paymentsState.pendingPagination.pageSize || PENDING_PAGE_SIZE;
    const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
    const page = Math.min(Math.max(1, paymentsState.pendingPagination.page || 1), totalPages);
    paymentsState.pendingPagination.page = page;
    return { page, pageSize, totalPages };
}

function updatePendingPaginationUi(totalItems) {
    const wrapper = document.getElementById("pendingBillsPagination");
    const info = document.getElementById("pendingBillsPageInfo");
    const prevBtn = document.getElementById("pendingBillsPrev");
    const nextBtn = document.getElementById("pendingBillsNext");
    if (!wrapper) return;

    const { page, totalPages } = clampPendingPage(totalItems);
    if (info) {
        info.textContent = totalItems
            ? `Page ${page} of ${totalPages} - ${totalItems} total`
            : "No pending bills";
    }
    if (prevBtn) {
        prevBtn.disabled = page <= 1;
        prevBtn.classList.toggle("opacity-50", prevBtn.disabled);
    }
    if (nextBtn) {
        nextBtn.disabled = page >= totalPages;
        nextBtn.classList.toggle("opacity-50", nextBtn.disabled);
    }

    wrapper.classList.toggle("hidden", totalItems <= (paymentsState.pendingPagination.pageSize || PENDING_PAGE_SIZE));
}

function readPaidFiltersFromUi() {
    const fromInput = document.getElementById("paidBillsFrom");
    const toInput = document.getElementById("paidBillsTo");
    const limitSelect = document.getElementById("paidBillsLimit");
    return {
        fromMonth: normalizeMonthKey(fromInput?.value || ""),
        toMonth: normalizeMonthKey(toInput?.value || ""),
        limit: Number(limitSelect?.value) || PAID_DEFAULT_LIMIT,
    };
}

function computeMonthsBackForRange(fromMonth, toMonth) {
    const { from } = normalizeMonthRange(fromMonth, toMonth);
    if (!from) return BILL_MONTHS_BACK;
    const [yearStr, monthStr] = from.split("-");
    const year = Number(yearStr);
    const month = Number(monthStr);
    if (!year || !month) return BILL_MONTHS_BACK;
    const now = new Date();
    const diff =
        (now.getFullYear() - year) * 12 +
        (now.getMonth() + 1 - month);
    return Math.max(1, diff + 1);
}

function summarizeBillPayments(bill) {
    const totalAmount = Number(bill?.totalAmount ?? bill?.total_amount) || 0;
    const remainingRaw = bill?.remainingAmount ?? bill?.remaining_amount;
    const amountPaidRaw = bill?.amountPaid ?? bill?.amount_paid;
    const isPaidRaw = bill?.isPaid ?? bill?.is_paid;
    const hasPaidRaw = bill?.hasPaid ?? bill?.has_paid;
    const hasRemaining = remainingRaw !== null && remainingRaw !== undefined && remainingRaw !== "";
    const hasAmountPaid = amountPaidRaw !== null && amountPaidRaw !== undefined && amountPaidRaw !== "";
    const hasIsPaid = isPaidRaw !== null && isPaidRaw !== undefined && isPaidRaw !== "";
    const hasHasPaid = hasPaidRaw !== null && hasPaidRaw !== undefined && hasPaidRaw !== "";
    const paidFlag = hasIsPaid
        ? normalizeBooleanValue(isPaidRaw)
        : hasHasPaid
        ? normalizeBooleanValue(hasPaidRaw)
        : null;

    if (hasRemaining || hasAmountPaid || hasIsPaid || hasHasPaid) {
        const remaining = hasRemaining
            ? Math.max(0, Number(remainingRaw) || 0)
            : paidFlag === true
            ? 0
            : Math.max(0, totalAmount - (Number(amountPaidRaw) || 0));
        return {
            remaining,
        };
    }
    return {
        remaining: totalAmount,
    };
}

async function getBillDetailsForModal(bill) {
    if (!bill || !bill.billLineId) return bill;
    if (typeof bill.prevReading !== "undefined" || typeof bill.electricityRate !== "undefined") {
        return bill;
    }

    try {
        const res = await fetchBillDetails(bill.billLineId);
        if (res && res.ok && res.bill) {
            return { ...bill, ...res.bill };
        }
    } catch (err) {
        console.error("Failed to load bill details", err);
        showToast("Unable to load full bill details", "error");
    }

    return bill;
}

function findPaymentForBill(bill) {
    const { matches, wingMatches } = getPaymentsForBill(bill);

    const byWing = wingMatches.find((p) => normalizeKey(p.wing) === normalizeKey(bill.wing));
    return byWing || matches[0] || null;
}

function getDueInfo(bill) {
    const raw = bill?.payableDate || "";
    const match = String(raw).match(/(\d{1,2})/);
    if (!match) return { label: "", overdue: false };

    const day = Number(match[1]);
    const baseLabel = `Due ${raw}`;
    const parts = (bill?.monthKey || "").split("-").map((p) => parseInt(p, 10));
    if (parts.length !== 2 || !parts[0] || !parts[1]) {
        return { label: baseLabel, overdue: false };
    }

    const dueDate = new Date(parts[0], parts[1] - 1, day);
    const today = new Date();
    return { label: baseLabel, overdue: today > dueDate };
}

function friendlyMonthLabel(monthKey = "", monthLabel = "") {
    const label = (monthLabel || "").toString();
    const normalized = (monthKey || "").toString();
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

function normalizeKey(value) {
    return (value || "").toString().trim().toLowerCase();
}

function normalizeMonthKey(value) {
    return normalizeMonthKeyBase(value, { lowercase: true });
}

function resolveUnitNumberForBill(bill) {
    const direct = bill?.unitNumber || bill?.unit_number;
    if (direct) return direct;

    const wing = bill?.wing || "";
    if (!wing) return "";

    const candidates = getActiveTenantsForWing(wing);
    if (!candidates.length) return "";

    const tenantKey = normalizeKey(bill?.tenantKey || bill?.tenantName);
    let match = null;
    if (tenantKey) {
        match = candidates.find((t) =>
            normalizeKey(
                t.grnNumber || t.grn_number || t.grn || t.tenantKey || t.tenantName || t.tenantFullName || t.name
            ) === tenantKey
        );
    }

    if (!match && bill?.tenantName) {
        match = candidates.find((t) => normalizeKey(t.tenantFullName || t.name) === normalizeKey(bill.tenantName));
    }

    return (
        match?.unitNumber ||
        match?.unit_number ||
        match?.unitLabel ||
        match?.templateData?.unit_number ||
        ""
    );
}

async function enrichBillsWithUnits(bills = [], options = {}) {
    if (!Array.isArray(bills) || !bills.length) return bills;
    const needsUnit = bills.some((bill) => !(bill.unitNumber || bill.unit_number));
    if (!needsUnit) return bills;

    const allowFetch = options.allowFetch !== false;
    if (allowFetch) {
        try {
            await ensureTenantDirectoryLoaded();
        } catch (err) {
            return bills;
        }
    }

    return bills.map((bill) => {
        const unitNumber = bill.unitNumber || bill.unit_number || resolveUnitNumberForBill(bill);
        return unitNumber ? { ...bill, unitNumber } : bill;
    });
}

function togglePaymentModal(show) {
    const modal = document.getElementById("paymentModal");
    if (modal) {
        show ? showModal(modal) : hideModal(modal);
    }
}

function setPaymentModalLoading(isLoading, label = "Loading details...") {
    const loader = document.getElementById("paymentModalLoading");
    if (!loader) return;
    const text = document.getElementById("paymentModalLoadingText");
    if (text && label) text.textContent = label;
    loader.classList.toggle("hidden", !isLoading);
}

function updatePaymentModalTitle(modalTitle, payment, context) {
    if (!modalTitle) return;
    const isSettled = payment && (context?.remaining ?? 0) <= 0;
    if (isSettled) {
        modalTitle.textContent = "Payment details";
        return;
    }
    modalTitle.textContent = payment ? "Edit payment" : "Record payment";
}

function mergePaymentOverrides(context, payment) {
    if (!payment) return context;
    return {
        ...context,
        amount: payment.amount ?? context.amount,
        rentAmount: payment.rentAmount ?? context.rentAmount,
        electricityAmount: payment.electricityAmount ?? context.electricityAmount,
        motorShare: payment.motorShare ?? context.motorShare,
        sweepAmount: payment.sweepAmount ?? context.sweepAmount,
        prevReading: payment.prevReading ?? context.prevReading,
        newReading: payment.newReading ?? context.newReading,
        payableDate: payment.payableDate ?? context.payableDate,
        notes: payment.notes ?? context.notes,
        mode: payment.mode ?? context.mode,
        date: payment.date ?? context.date,
        billLineId: payment.billLineId ?? context.billLineId,
        tenancyId: payment.tenancyId ?? context.tenancyId,
    };
}

function resetPaymentForm() {
    cancelAttachmentUpload();
    discardUploadedAttachment();
    paymentsState.editingId = "";
    paymentsState.attachmentDataUrl = "";
    paymentsState.attachmentName = "";
    paymentsState.attachmentUrl = "";
    paymentsState.attachmentPreviewUrl = "";
    paymentsState.attachmentViewUrl = "";
    paymentsState.attachmentId = "";
    paymentsState.viewOnly = false;
    paymentsState.billContext = {
        monthKey: "",
        monthLabel: "",
        billTotal: 0,
        tenantKey: "",
        tenantName: "",
        wing: "",
        amount: 0,
        remaining: 0,
        rentAmount: 0,
        electricityAmount: 0,
        motorShare: 0,
        sweepAmount: 0,
        prevReading: "",
        newReading: "",
        payableDate: "",
    };

    const todayIso = new Date().toISOString().slice(0, 10);

    const dateInput = document.getElementById("paymentDateInput");
    const modeSelect = document.getElementById("paymentModeSelect");
    const notesInput = document.getElementById("paymentNotesInput");
    const idField = document.getElementById("paymentRecordId");
    const amountInput = document.getElementById("paymentAmountInput");
    const billTitle = document.getElementById("paymentBillTitle");
    const billAmount = document.getElementById("paymentBillAmount");
    const billStatus = document.getElementById("paymentBillStatus");
    const billMonth = document.getElementById("paymentBillMonth");
    const wingBadge = document.getElementById("paymentTenantWing");
    const dueWrap = document.getElementById("paymentAmountDueWrap");
    const breakdownTotal = document.getElementById("paymentBreakdownTotal");
    const breakdownRent = document.getElementById("paymentBreakdownRent");
    const breakdownElec = document.getElementById("paymentBreakdownElectricity");
    const breakdownMotor = document.getElementById("paymentBreakdownMotor");
    const breakdownSweep = document.getElementById("paymentBreakdownSweep");
    const attachmentPreview = document.getElementById("paymentAttachmentPreview");
    const attachmentName = document.getElementById("paymentAttachmentName");
    const attachmentLink = document.getElementById("paymentAttachmentLink");
    const attachmentInput = document.getElementById("paymentAttachmentInput");
    const attachmentClear = document.getElementById("paymentAttachmentClear");
    const saveBtn = document.getElementById("paymentSaveBtn");
    const formFields = document.getElementById("paymentFormFields");
    const notesSection = document.getElementById("paymentNotesSection");
    const attachmentWrapper = document.getElementById("paymentAttachmentWrapper");
    const progressWrap = document.getElementById("paymentAttachmentProgress");
    const progressFill = document.getElementById("paymentAttachmentProgressFill");
    const progressLabel = document.getElementById("paymentAttachmentProgressLabel");
    const zeroLabel = formatCurrency(0);

    if (dateInput) dateInput.value = todayIso;
    if (modeSelect) modeSelect.value = "UPI";
    if (notesInput) notesInput.value = "";
    if (idField) idField.value = "";
    if (amountInput) amountInput.value = "";
    if (amountInput) {
        amountInput.disabled = false;
        amountInput.max = "";
        amountInput.placeholder = "Enter amount received";
    }

    if (billTitle) billTitle.textContent = "Select a bill to record";
    if (billAmount) billAmount.textContent = zeroLabel;
    if (billStatus) billStatus.textContent = "Open this modal from a bill card to link the payment automatically.";
    if (billMonth) billMonth.textContent = "Month • -";
    if (wingBadge) wingBadge.textContent = "Wing • -";
    if (breakdownTotal) breakdownTotal.textContent = zeroLabel;
    if (breakdownRent) breakdownRent.textContent = zeroLabel;
    if (breakdownElec) breakdownElec.textContent = zeroLabel;
    if (breakdownMotor) breakdownMotor.textContent = zeroLabel;
    if (breakdownSweep) breakdownSweep.textContent = zeroLabel;

    if (attachmentPreview) {
        attachmentPreview.classList.add("hidden");
        const img = attachmentPreview.querySelector("img");
        if (img) {
            img.src = "";
            img.classList.add("hidden");
        }
    }
    if (attachmentName) attachmentName.textContent = "No file selected";
    if (attachmentLink) attachmentLink.classList.add("hidden");
    if (attachmentInput) attachmentInput.disabled = false;
    if (attachmentClear) attachmentClear.disabled = false;
    if (saveBtn) saveBtn.textContent = "Save payment";
    if (formFields) formFields.classList.remove("hidden");
    if (notesSection) notesSection.classList.remove("hidden");
    if (attachmentWrapper) attachmentWrapper.classList.remove("hidden");
    if (progressWrap) progressWrap.classList.add("hidden");
    if (progressFill) progressFill.style.width = "0%";
    if (progressLabel) progressLabel.textContent = "";
    if (dueWrap) dueWrap.classList.remove("hidden");
    setPaymentModalLoading(false);
    const historyContainer = document.getElementById("paymentHistoryContainer");
    const historyList = document.getElementById("paymentHistoryList");
    const historySummary = document.getElementById("paymentHistorySummary");
    const historyEmpty = document.getElementById("paymentHistoryEmpty");
    if (historyContainer) historyContainer.classList.add("hidden");
    if (historyList) historyList.innerHTML = "";
    if (historySummary) historySummary.textContent = "Bill receipts will appear here.";
    if (historyEmpty) historyEmpty.classList.add("hidden");
}

function setPaymentFormReadOnly(isReadOnly) {
    const dateInput = document.getElementById("paymentDateInput");
    const modeSelect = document.getElementById("paymentModeSelect");
    const notesInput = document.getElementById("paymentNotesInput");
    const amountInput = document.getElementById("paymentAmountInput");
    const attachmentInput = document.getElementById("paymentAttachmentInput");
    const attachmentClear = document.getElementById("paymentAttachmentClear");
    const saveBtn = document.getElementById("paymentSaveBtn");
    const formFields = document.getElementById("paymentFormFields");
    const notesSection = document.getElementById("paymentNotesSection");
    const attachmentWrapper = document.getElementById("paymentAttachmentWrapper");

    if (dateInput) dateInput.disabled = isReadOnly;
    if (modeSelect) modeSelect.disabled = isReadOnly;
    if (notesInput) notesInput.disabled = isReadOnly;
    if (amountInput) amountInput.disabled = isReadOnly;
    if (attachmentInput) attachmentInput.disabled = isReadOnly;
    if (attachmentClear) attachmentClear.disabled = isReadOnly;
    if (saveBtn) saveBtn.textContent = isReadOnly ? "Close" : "Save payment";
    if (formFields) formFields.classList.toggle("hidden", isReadOnly);
    if (notesSection) notesSection.classList.toggle("hidden", isReadOnly);
    if (attachmentWrapper) attachmentWrapper.classList.toggle("hidden", isReadOnly);
}

function renderPaymentHistory(context = {}) {
    const container = document.getElementById("paymentHistoryContainer");
    const list = document.getElementById("paymentHistoryList");
    const empty = document.getElementById("paymentHistoryEmpty");
    const summary = document.getElementById("paymentHistorySummary");
    if (!container || !list) return;

    list.innerHTML = "";
    if (empty) empty.classList.add("hidden");

    if (!context.monthKey || !context.tenantName) {
        container.classList.add("hidden");
        return;
    }

    if (!paymentsState.loaded) {
        container.classList.remove("hidden");
        if (summary) summary.textContent = "Loading payment history...";
        if (empty) empty.classList.add("hidden");
        if (!paymentHistoryLoading) {
            paymentHistoryLoading = true;
            loadPayments()
                .then(() => renderPaymentHistory(context))
                .catch((err) => {
                    console.error("Failed to load payments for history", err);
                    if (summary) summary.textContent = "Unable to load payment history.";
                    if (empty) empty.classList.remove("hidden");
                })
                .finally(() => {
                    paymentHistoryLoading = false;
                });
        }
        return;
    }

    const { matches } = getPaymentsForBill(context);
    if (!matches.length) {
        container.classList.remove("hidden");
        if (summary) summary.textContent = "No payments recorded yet for this bill.";
        if (empty) empty.classList.remove("hidden");
        return;
    }

    const sorted = [...matches].sort((a, b) => {
        const aTime = new Date(a.createdAt || a.date || 0).getTime();
        const bTime = new Date(b.createdAt || b.date || 0).getTime();
        return aTime - bTime;
    });

    const downloadIcon = `
        <svg viewBox="0 0 20 20" fill="currentColor" class="w-4 h-4" aria-hidden="true">
            <path d="M10 2a1 1 0 0 1 1 1v7.59l2.3-2.3a1 1 0 1 1 1.4 1.42l-4 4a1 1 0 0 1-1.4 0l-4-4a1 1 0 1 1 1.4-1.42L9 10.59V3a1 1 0 0 1 1-1z"/>
            <path d="M4 14a1 1 0 0 1 1 1v1h10v-1a1 1 0 1 1 2 0v2a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-2a1 1 0 0 1 1-1z"/>
        </svg>
    `;
    const copyIcon = `
        <svg viewBox="0 0 20 20" fill="currentColor" class="w-4 h-4" aria-hidden="true">
            <path d="M6 2a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h1a1 1 0 1 0 0-2H6V4h7v1a1 1 0 1 0 2 0V4a2 2 0 0 0-2-2H6z"/>
            <path d="M9 7a2 2 0 0 0-2 2v7a2 2 0 0 0 2 2h5a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2H9zm0 2h5v7H9V9z"/>
        </svg>
    `;

    const header = document.createElement("div");
    header.className = "grid grid-cols-[32px,1fr,1fr,1fr,auto] gap-3 px-3 text-[10px] uppercase text-slate-500";
    header.innerHTML = `
        <div>No.</div>
        <div>Amount</div>
        <div>Mode</div>
        <div>Date</div>
        <div class="text-right">Image</div>
    `;
    list.appendChild(header);

    sorted.forEach((p, idx) => {
        const card = document.createElement("div");
        card.className = "border border-slate-200 rounded-lg bg-white shadow-sm overflow-hidden";

        const amountLabel = formatCurrency(p.amount);
        const dateLabel = p.createdAt || p.date || "-";
        const modeLabel = p.mode || "-";
        const notesLabel = p.notes || "";
        const rawUrl = p.attachmentUrl || "";
        const thumbUrl = rawUrl ? normalizeAttachmentUrl(rawUrl) : "";
        const hasAttachment = !!thumbUrl;
        const noteText = notesLabel || "No notes";

        card.innerHTML = `
            <div class="grid grid-cols-[32px,1fr,1fr,1fr,auto] gap-3 items-center px-3 py-2">
                <div class="text-[10px] font-semibold text-slate-500">#${idx + 1}</div>
                <div class="text-[12px] font-semibold text-slate-900">${amountLabel}</div>
                <div class="text-[11px] text-slate-700">
                    ${modeLabel !== "-" ? `<span class="px-2 py-0.5 rounded-full bg-slate-100 text-slate-700">${modeLabel}</span>` : "-"}
                </div>
                <div class="text-[10px] text-slate-500 whitespace-nowrap">${dateLabel}</div>
                <div class="flex items-center justify-end gap-2">
                    <div class="receipt-shell w-14 h-14 rounded-lg border bg-slate-50 flex items-center justify-center overflow-hidden">
                        <img class="receipt-thumb ${hasAttachment ? "" : "hidden"} w-full h-full object-cover" alt="Receipt preview" />
                        <span class="receipt-placeholder ${hasAttachment ? "hidden" : ""} text-[9px] text-slate-400">No image</span>
                    </div>
                    <div class="flex flex-col gap-1">
                        <button type="button" class="receipt-action receipt-download ${hasAttachment ? "text-slate-700 hover:text-slate-900" : "text-slate-400 cursor-not-allowed"}" title="Download receipt">
                            ${downloadIcon}
                        </button>
                        <button type="button" class="receipt-action receipt-copy ${hasAttachment ? "text-slate-700 hover:text-slate-900" : "text-slate-400 cursor-not-allowed"}" title="Copy receipt">
                            ${copyIcon}
                        </button>
                    </div>
                </div>
            </div>
            <div class="border-t bg-slate-50 px-3 py-2 text-[10px] text-slate-600">
                <span class="uppercase text-[9px] font-semibold text-slate-500">Note</span>
                <span class="ml-1">${noteText}</span>
            </div>
        `;

        const thumbShell = card.querySelector(".receipt-shell");
        const img = card.querySelector(".receipt-thumb");
        const placeholder = card.querySelector(".receipt-placeholder");
        const downloadBtn = card.querySelector(".receipt-download");
        const copyBtn = card.querySelector(".receipt-copy");
        let targetUrl = thumbUrl;

        const getTargetUrl = () => targetUrl || rawUrl || thumbUrl;
        const openTarget = () => {
            const href = getTargetUrl();
            if (href) openAttachmentViewer(href, p.attachmentName || "Receipt");
        };
        const showPlaceholder = () => {
            if (img) img.classList.add("hidden");
            if (placeholder) placeholder.classList.remove("hidden");
        };

        if (hasAttachment) {
            if (thumbShell) {
                thumbShell.classList.add("cursor-pointer");
                thumbShell.addEventListener("click", openTarget);
            }
            if (downloadBtn) {
                downloadBtn.disabled = false;
                downloadBtn.addEventListener("click", () =>
                    downloadAttachment(getTargetUrl(), p.attachmentName || "receipt")
                );
            }
            if (copyBtn) {
                copyBtn.disabled = false;
                copyBtn.addEventListener("click", () => copyAttachmentToClipboard(getTargetUrl()));
            }
        } else {
            if (downloadBtn) downloadBtn.disabled = true;
            if (copyBtn) copyBtn.disabled = true;
        }

        if (img && hasAttachment) {
            img.referrerPolicy = "no-referrer";
            img.src = thumbUrl;
            img.classList.remove("hidden");
            if (placeholder) placeholder.classList.add("hidden");
            resolveAttachmentPreview(thumbUrl)
                .then(({ previewUrl, viewUrl }) => {
                    const src = previewUrl || viewUrl || thumbUrl;
                    if (src) img.src = src;
                    targetUrl = previewUrl || viewUrl || thumbUrl;
                })
                .catch(() => {
                    targetUrl = thumbUrl || rawUrl;
                })
                .finally(() => {
                    if (!targetUrl) showPlaceholder();
                });
        } else {
            showPlaceholder();
        }

        list.appendChild(card);
    });

    if (summary) {
        const totalPaid = sorted.reduce((sum, p) => sum + (Number(p.amount) || 0), 0);
        summary.textContent = `Showing ${sorted.length} partial payment${sorted.length === 1 ? "" : "s"} • Total recorded ${formatCurrency(totalPaid)}`;
    }

    container.classList.remove("hidden");
}

function applyBillContext(context = {}) {
    const billTitle = document.getElementById("paymentBillTitle");
    const billAmount = document.getElementById("paymentBillAmount");
    const billStatus = document.getElementById("paymentBillStatus");
    const billMonth = document.getElementById("paymentBillMonth");
    const wingBadge = document.getElementById("paymentTenantWing");
    const dueWrap = document.getElementById("paymentAmountDueWrap");
    const breakdownTotal = document.getElementById("paymentBreakdownTotal");
    const breakdownRent = document.getElementById("paymentBreakdownRent");
    const breakdownElec = document.getElementById("paymentBreakdownElectricity");
    const breakdownMotor = document.getElementById("paymentBreakdownMotor");
    const breakdownSweep = document.getElementById("paymentBreakdownSweep");
    const amountInput = document.getElementById("paymentAmountInput");

    const monthLabel = friendlyMonthLabel(context.monthKey, context.monthLabel);
    const rentAmount = Number(context.rentAmount || 0) || 0;
    const electricityAmount = Number(context.electricityAmount || 0) || 0;
    const motorAmount = Number(context.motorShare || 0) || 0;
    const sweepAmount = Number(context.sweepAmount || 0) || 0;
    const billTotal = Number(
        context.totalAmount ||
        context.billTotal ||
        context.amount ||
        rentAmount + electricityAmount + motorAmount + sweepAmount
    ) || 0;
    const remaining = typeof context.remaining === "number" ? context.remaining : Number(context.amount || billTotal);

    paymentsState.billContext = {
        monthKey: context.monthKey || "",
        monthLabel,
        billTotal,
        amount: typeof context.amount === "number" ? context.amount : remaining,
        remaining,
        tenantName: context.tenantName || "",
        tenantKey: context.tenantKey || context.grn || "",
        wing: context.wing || "",
        rentAmount,
        electricityAmount,
        motorAmount,
        motorShare: motorAmount,
        sweepAmount,
        prevReading: context.prevReading || "",
        newReading: context.newReading || "",
        payableDate: context.payableDate || "",
        tenancyId: context.tenancyId || "",
        billLineId: context.billLineId || "",
    };

    const fallbackAmountLabel = formatCurrency(0);

    if (billTitle) billTitle.textContent = context.tenantName ? `${context.tenantName}` : "Select a bill to record";
    if (billAmount) {
        const headlineAmount = typeof remaining === "number" ? remaining : billTotal;
        billAmount.textContent = billTotal ? formatCurrency(headlineAmount) : fallbackAmountLabel;
    }
    if (dueWrap) dueWrap.classList.toggle("hidden", typeof remaining === "number" && remaining <= 0);
    if (billStatus) {
        const dueLabel = context.payableDate ? `Due ${context.payableDate}` : "";
        billStatus.textContent = dueLabel || "Bill selected";
    }
    if (billMonth) billMonth.textContent = `Month • ${monthLabel || "-"}`;
    if (wingBadge) wingBadge.textContent = `Wing • ${context.wing || "-"}`;
    if (breakdownTotal) breakdownTotal.textContent = formatCurrency(billTotal || remaining || 0);
    if (breakdownRent) breakdownRent.textContent = formatCurrency(rentAmount);
    if (breakdownElec) breakdownElec.textContent = formatCurrency(electricityAmount);
    if (breakdownMotor) breakdownMotor.textContent = formatCurrency(motorAmount);
    if (breakdownSweep) breakdownSweep.textContent = formatCurrency(sweepAmount);
    if (amountInput) {
        const defaultAmount =
            typeof paymentsState.billContext.amount === "number"
                ? paymentsState.billContext.amount
                : remaining || billTotal || 0;
        amountInput.value = defaultAmount ? Number(defaultAmount).toFixed(2) : "";
        amountInput.max = billTotal || "";
        amountInput.placeholder = billTotal ? `Up to ${formatCurrency(remaining || billTotal)}` : "Enter amount received";
    }

    renderPaymentHistory(paymentsState.billContext);
}

function normalizeAttachmentUrl(url = "") {
    if (!url) return "";

    const driveMatch = url.match(/\/file\/d\/([^/]+)\//);
    if (driveMatch?.[1]) {
        return `https://drive.google.com/uc?export=download&id=${driveMatch[1]}`;
    }

    const idParam = url.match(/[?&]id=([^&#]+)/);
    if (idParam?.[1]) {
        return `https://drive.google.com/uc?export=download&id=${idParam[1]}`;
    }

    return url;
}

async function resolveAttachmentPreview(url = "") {
    const normalized = normalizeAttachmentUrl(url);
    if (!normalized) return { previewUrl: "", viewUrl: "" };

    if (normalized.startsWith("data:")) {
        return { previewUrl: normalized, viewUrl: normalized };
    }

    if (attachmentPreviewCache.has(normalized)) {
        return attachmentPreviewCache.get(normalized);
    }

    try {
        const result = await fetchAttachmentPreview(normalized);
        const viewUrl = normalizeAttachmentUrl(result?.attachmentUrl || normalized);
        const resolved = {
            previewUrl: result?.previewUrl || viewUrl,
            viewUrl,
        };
        attachmentPreviewCache.set(normalized, resolved);
        return resolved;
    } catch (err) {
        console.warn("resolveAttachmentPreview fallback", err);
        const fallback = {
            previewUrl: normalized,
            viewUrl: normalized,
        };
        attachmentPreviewCache.set(normalized, fallback);
        return fallback;
    }
}

function openAttachmentViewer(url, title = "Receipt") {
    const modal = document.getElementById("attachmentViewerModal");
    const iframe = document.getElementById("attachmentViewerFrame");
    const image = document.getElementById("attachmentViewerImage");
    const caption = document.getElementById("attachmentViewerTitle");
    const fallback = document.getElementById("attachmentViewerFallback");

    if (!modal || !iframe || !image || !caption || !fallback) return;

    const viewUrl = normalizeAttachmentUrl(url || "");
    caption.textContent = title || "Receipt";

    const nameHint = (paymentsState.attachmentName || "").toLowerCase();
    const nameIsImage = /\.(png|jpe?g|gif|webp|bmp|svg)(\?|$)/i.test(nameHint);
    const isImageLike =
        nameIsImage ||
        viewUrl.startsWith("data:image") ||
        /\.(png|jpe?g|gif|webp|bmp|svg)(\?|$)/i.test(viewUrl) ||
        viewUrl.includes("export=view");

    image.classList.add("hidden");
    iframe.classList.add("hidden");
    fallback.classList.add("hidden");

    if (isImageLike && viewUrl) {
        image.referrerPolicy = "no-referrer";
        image.src = viewUrl;
        image.classList.remove("hidden");
    } else if (viewUrl) {
        const viewerUrl = `https://docs.google.com/gview?embedded=1&url=${encodeURIComponent(viewUrl)}`;
        iframe.src = viewerUrl;
        iframe.classList.remove("hidden");
    } else {
        fallback.classList.remove("hidden");
    }

    showModal(modal);
}

function closeAttachmentViewer() {
    const modal = document.getElementById("attachmentViewerModal");
    const iframe = document.getElementById("attachmentViewerFrame");
    const image = document.getElementById("attachmentViewerImage");
    const fallback = document.getElementById("attachmentViewerFallback");

    if (modal) hideModal(modal);
    if (iframe) iframe.src = "";
    if (image) image.src = "";
    if (fallback) fallback.classList.add("hidden");
}

function openAttachmentViewerFromState() {
    const viewUrl = paymentsState.attachmentPreviewUrl || paymentsState.attachmentViewUrl;
    if (!viewUrl) {
        showToast("No receipt available to preview", "warning");
        return;
    }
    openAttachmentViewer(viewUrl, paymentsState.attachmentName || "Receipt");
}

async function showAttachmentPreview(name, url) {
    const attachmentPreview = document.getElementById("paymentAttachmentPreview");
    const attachmentName = document.getElementById("paymentAttachmentName");
    const attachmentLink = document.getElementById("paymentAttachmentLink");
    const { previewUrl, viewUrl } = await resolveAttachmentPreview(url);

    paymentsState.attachmentPreviewUrl = previewUrl || "";
    paymentsState.attachmentViewUrl = viewUrl || "";

    if (attachmentPreview) {
        const img = attachmentPreview.querySelector("img");
        if (previewUrl) {
            if (img) {
                img.referrerPolicy = "no-referrer";
                img.src = previewUrl;
                img.classList.remove("hidden");
            }
            attachmentPreview.classList.remove("hidden");
            attachmentPreview.classList.add("cursor-pointer");
        } else {
            if (img) {
                img.src = "";
                img.classList.add("hidden");
            }
            attachmentPreview.classList.add("hidden");
            attachmentPreview.classList.remove("cursor-pointer");
        }
    }

    if (attachmentName) attachmentName.textContent = name || "No file selected";

    if (attachmentLink) {
        const anchor = attachmentLink.querySelector("a");
        if (anchor) {
            anchor.href = viewUrl || previewUrl || "#";
            anchor.target = "_self";
        }
        attachmentLink.classList.toggle("hidden", !previewUrl && !viewUrl);
    }
}

async function openPaymentModal(payment = null, billContext = null) {
    if (paymentModalResetTimer) {
        clearTimeout(paymentModalResetTimer);
        paymentModalResetTimer = null;
    }
    const modalRequestId = ++paymentModalRequestId;
    resetPaymentForm();

    const modalTitle = document.getElementById("paymentModalTitle");
    const dateInput = document.getElementById("paymentDateInput");
    const modeSelect = document.getElementById("paymentModeSelect");
    const notesInput = document.getElementById("paymentNotesInput");
    const idField = document.getElementById("paymentRecordId");

    let context = billContext ? { ...billContext } : null;
    if (payment) {
        const baseContext = {
            monthKey: payment.monthKey,
            monthLabel: payment.monthLabel,
            billTotal: payment.billTotal || payment.amount,
            totalAmount: payment.billTotal || payment.amount,
            tenantName: payment.tenantName,
            tenantKey: payment.tenantKey,
            wing: payment.wing,
            amount: payment.amount,
            billLineId: payment.billLineId,
            tenancyId: payment.tenancyId,
        };
        context = context ? { ...baseContext, ...context } : baseContext;
        context = mergePaymentOverrides(context, payment);
    }

    if (context && typeof context.remaining !== "number") {
        const status = summarizeBillPayments(context);
        if (typeof status.remaining === "number") {
            context = { ...context, remaining: status.remaining };
        }
    }

    updatePaymentModalTitle(modalTitle, payment, context);

    if (context) applyBillContext(context);

    if (payment) {
        paymentsState.viewOnly = (context?.remaining ?? 0) <= 0;
        paymentsState.editingId = payment.id;
        paymentsState.attachmentName = payment.attachmentName || "";
        paymentsState.attachmentUrl = normalizeAttachmentUrl(payment.attachmentUrl || "");
        paymentsState.attachmentId = payment.attachmentId || "";
        paymentsState.attachmentCommitted = !!(payment.attachmentUrl || payment.attachmentId);
        paymentsState.attachmentDataUrl = "";

        if (modeSelect && payment.mode) modeSelect.value = payment.mode;
        if (dateInput) dateInput.value = payment.date || dateInput.value;
        if (notesInput) notesInput.value = payment.notes || "";
        if (idField) idField.value = payment.id || "";
    } else {
        paymentsState.viewOnly = false;
    }

    setPaymentFormReadOnly(paymentsState.viewOnly);
    togglePaymentModal(true);

    if (payment?.attachmentUrl) {
        showAttachmentPreview(payment.attachmentName, payment.attachmentUrl);
    }

    const isActiveRequest = () =>
        modalRequestId === paymentModalRequestId && isPaymentModalVisible();
    let detailSource = context && context.billLineId ? context : null;

    if (!detailSource && payment?.billLineId) {
        detailSource = { ...(context || {}), billLineId: payment.billLineId };
    }

    if (!detailSource && payment) {
        const matchingBill = getAllGeneratedBills().find(
            (bill) =>
                normalizeKey(bill.tenantKey || bill.tenantName) ===
                    normalizeKey(payment.tenantKey || payment.tenantName) &&
                normalizeMonthKey(bill.monthKey) === normalizeMonthKey(payment.monthKey) &&
                (bill.wing || "").toLowerCase() === (payment.wing || "").toLowerCase()
        );
        if (matchingBill) detailSource = matchingBill;
    }

    const needsDetailsFetch =
        !!detailSource?.billLineId &&
        typeof detailSource.prevReading === "undefined" &&
        typeof detailSource.electricityRate === "undefined";

    if (needsDetailsFetch) {
        setPaymentModalLoading(true);
        getBillDetailsForModal(detailSource)
            .then((detailedBill) => {
                if (!isActiveRequest() || !detailedBill) return;
                const status = summarizeBillPayments(detailedBill);
                let updatedContext = { ...detailedBill, remaining: status.remaining };
                updatedContext = mergePaymentOverrides(updatedContext, payment);
                applyBillContext(updatedContext);
                if (payment) {
                    paymentsState.viewOnly = (updatedContext?.remaining ?? 0) <= 0;
                    setPaymentFormReadOnly(paymentsState.viewOnly);
                    updatePaymentModalTitle(modalTitle, payment, updatedContext);
                }
            })
            .finally(() => {
                if (isActiveRequest()) setPaymentModalLoading(false);
            });
    }
}

async function openPaymentModalFromBill(bill) {
    const status = summarizeBillPayments(bill);
    const context = {
        ...bill,
        remaining: status.remaining,
    };
    openPaymentModal(null, context);
}

function closePaymentModal() {
    cancelAttachmentUpload();
    togglePaymentModal(false);
    if (paymentModalResetTimer) clearTimeout(paymentModalResetTimer);
    paymentModalResetTimer = setTimeout(() => {
        resetPaymentForm();
        paymentModalResetTimer = null;
    }, 240);
}

function setBillsTab(tab, options = {}) {
    paymentsState.activeBillTab = tab;
    const pendingTab = document.getElementById("billsPendingTab");
    const paidTab = document.getElementById("billsPaidTab");
    const pendingBody = document.getElementById("pendingBillsBody");
    const paidBody = document.getElementById("paidBillsBody");
    const paidFilters = document.getElementById("paidBillsFilters");
    const pendingPagination = document.getElementById("pendingBillsPagination");
    const loader = document.getElementById("generatedBillsLoader");
    const emptyState = document.getElementById("generatedBillsEmpty");
    if (pendingTab) {
        pendingTab.classList.toggle("bg-white", tab === "pending");
        pendingTab.classList.toggle("text-slate-800", tab === "pending");
    }
    if (paidTab) {
        paidTab.classList.toggle("bg-white", tab === "paid");
        paidTab.classList.toggle("text-slate-800", tab === "paid");
    }
    if (pendingBody && paidBody) {
        pendingBody.classList.toggle("hidden", tab !== "pending");
        paidBody.classList.toggle("hidden", tab !== "paid");
    }
    if (paidFilters) paidFilters.classList.toggle("hidden", tab !== "paid");
    if (pendingPagination) pendingPagination.classList.toggle("hidden", tab !== "pending");

    const deferPaidLoad = tab === "paid" && !paymentsState.paidFilters?.hasSearched;

    if (!options.skipLoad && !isBillsLoaded(tab) && !deferPaidLoad) {
        if (loader) loader.classList.remove("hidden");
        if (emptyState) emptyState.classList.add("hidden");
        if (tab === "pending" && pendingBody) pendingBody.innerHTML = "";
        if (tab === "paid" && paidBody) paidBody.innerHTML = "";
    }
    if (options.forceLoad) {
        loadGeneratedBills(tab, true, options);
        return;
    }
    if (options.skipLoad || deferPaidLoad) {
        renderGeneratedBills();
        return;
    }
    if (isBillsLoaded(tab)) {
        renderGeneratedBills();
    } else {
        loadGeneratedBills(tab);
    }
}

async function downloadAttachment(url, name) {
    if (!url) {
        showToast("No receipt image to download", "warning");
        return;
    }
    const { viewUrl, previewUrl } = await resolveAttachmentPreview(url);
    const href = viewUrl || previewUrl || url;
    if (!href) {
        showToast("No receipt image to download", "warning");
        return;
    }
    const anchor = document.createElement("a");
    anchor.href = href;
    anchor.download = name || "receipt";
    anchor.rel = "noopener noreferrer";
    anchor.target = "_blank";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    const fileLabel = name || "receipt";
    showToast(`Download started: ${fileLabel}`, "success");
}

async function copyAttachmentToClipboard(url) {
    if (!url) {
        showToast("No receipt image to copy", "warning");
        return;
    }
    try {
        const { previewUrl, viewUrl } = await resolveAttachmentPreview(url);
        const target = previewUrl || viewUrl || url;
        if (!target) {
            showToast("No receipt image to copy", "warning");
            return;
        }
        if (navigator.clipboard && window.ClipboardItem) {
            const response = await fetch(target);
            const blob = await response.blob();
            const item = new ClipboardItem({ [blob.type || "image/png"]: blob });
            await navigator.clipboard.write([item]);
            showToast("Receipt image copied", "success");
            return;
        }
        if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(target);
            showToast("Receipt link copied", "success");
            return;
        }
        showToast("Clipboard not available", "warning");
    } catch (err) {
        console.error("Failed to copy receipt image", err);
        showToast("Unable to copy receipt image", "error");
    }
}

function setRecordButtonLoading(button, isLoading) {
    if (!button) return;
    if (isLoading) {
        if (!button.dataset.label) {
            button.dataset.label = button.textContent || "Record payment";
        }
        button.disabled = true;
        button.classList.add("opacity-75", "cursor-wait");
        button.innerHTML = `
            <span class="inline-flex items-center gap-2">
                <span class="inline-block h-3 w-3 rounded-full border-2 border-white border-t-transparent animate-spin"></span>
                ${button.dataset.label}
            </span>
        `;
        return;
    }
    button.disabled = false;
    button.classList.remove("opacity-75", "cursor-wait");
    button.textContent = button.dataset.label || "Record payment";
}

function renderGeneratedBills() {
    const pendingBody = document.getElementById("pendingBillsBody");
    const paidBody = document.getElementById("paidBillsBody");
    const loader = document.getElementById("generatedBillsLoader");
    const emptyState = document.getElementById("generatedBillsEmpty");
    if (!pendingBody || !paidBody) return;

    if (loader) loader.classList.add("hidden");

    const makeRow = (targetBody, { bill, status }, isPaid) => {
        const tr = document.createElement("tr");
        tr.className = "border-b last:border-0 hover:bg-slate-50";

        const monthLabel = friendlyMonthLabel(bill.monthKey, bill.monthLabel);
        const remainingLabel = status.remaining > 0 ? `${formatCurrency(status.remaining)} due` : "Paid";
        const totalLabel = formatCurrency(bill.totalAmount ?? bill.total_amount);
        const dueInfo = getDueInfo(bill);
        const dueBadge = bill.payableDate
            ? `<div class="text-[10px] ${dueInfo.overdue && !isPaid ? "text-rose-700" : "text-slate-500"}">${dueInfo.label}</div>`
            : "";

        tr.innerHTML = `
            <td class="px-3 py-2 text-[11px] font-semibold">${bill.tenantName || "-"}</td>
            <td class="px-3 py-2 text-[11px]">${bill.unitNumber || bill.unit_number || "-"}</td>
            <td class="px-3 py-2 text-[11px]">${bill.wing || "-"}</td>
            <td class="px-3 py-2 text-[11px]">${monthLabel}${dueBadge}</td>
            <td class="px-3 py-2 text-[11px] text-right">${totalLabel}</td>
            <td class="px-3 py-2 text-[11px] text-right ${status.remaining > 0 ? "text-amber-700" : "text-emerald-700"}">${remainingLabel}</td>
            <td class="px-3 py-2 text-[11px] text-right"></td>
        `;

        if (dueInfo.overdue && status.remaining > 0) {
            tr.classList.add("bg-amber-50");
        }

        const actionCell = tr.querySelector("td:last-child");
        if (actionCell) {
            if (isPaid) {
                const viewBtn = document.createElement("button");
                viewBtn.className =
                    "px-3 py-1.5 rounded-lg bg-slate-200 text-slate-800 text-[11px] font-semibold hover:bg-slate-300";
                viewBtn.textContent = "View details";
                viewBtn.addEventListener("click", async () => {
                    if (!paymentsState.loaded) {
                        try {
                            await loadPayments();
                        } catch (err) {
                            console.error("Failed to load payments for details", err);
                        }
                    }
                    const payment = findPaymentForBill(bill);
                    openPaymentModal(payment || null, { ...bill, remaining: status.remaining });
                });
                actionCell.appendChild(viewBtn);
            } else {
                const btn = document.createElement("button");
                btn.className = "px-3 py-1.5 rounded-lg bg-indigo-600 text-white text-[11px] font-semibold hover:bg-indigo-500";
                btn.textContent = "Record payment";
                btn.addEventListener("click", async () => {
                    if (btn.disabled) return;
                    setRecordButtonLoading(btn, true);
                    try {
                        await openPaymentModalFromBill({ ...bill, remaining: status.remaining });
                    } finally {
                        setRecordButtonLoading(btn, false);
                    }
                });
                actionCell.appendChild(btn);
            }
        }

        targetBody.appendChild(tr);
    };

    const activeTab = paymentsState.activeBillTab || "pending";
    if (activeTab === "pending") {
        pendingBody.innerHTML = "";
        paidBody.classList.add("hidden");
        pendingBody.classList.remove("hidden");

        const pendingBills = filterPendingBills(getBillsForStatus("pending"));
        if (!pendingBills.length) {
            if (emptyState) {
                emptyState.textContent = "No pending bills right now.";
                emptyState.classList.remove("hidden");
            }
            updatePendingPaginationUi(0);
            return;
        }

        if (emptyState) emptyState.classList.add("hidden");

        const { page, pageSize } = clampPendingPage(pendingBills.length);
        const startIdx = (page - 1) * pageSize;
        const pageBills = pendingBills.slice(startIdx, startIdx + pageSize);
        pageBills.forEach((bill) => {
            const status = summarizeBillPayments(bill);
            makeRow(pendingBody, { bill, status }, false);
        });
        updatePendingPaginationUi(pendingBills.length);
        return;
    }

    paidBody.innerHTML = "";
    pendingBody.classList.add("hidden");
    paidBody.classList.remove("hidden");

    if (!paymentsState.paidFilters?.hasSearched) {
        if (emptyState) {
            emptyState.textContent = "Select a month range and click Search to view paid bills.";
            emptyState.classList.remove("hidden");
        }
        return;
    }

    const paidBills = applyPaidBillFilters(getBillsForStatus("paid"), paymentsState.paidFilters);
    if (!paidBills.length) {
        if (emptyState) {
            emptyState.textContent = "No paid bills match this search.";
            emptyState.classList.remove("hidden");
        }
        return;
    }

    if (emptyState) emptyState.classList.add("hidden");

    paidBills.forEach((bill) => {
        const status = summarizeBillPayments(bill);
        makeRow(paidBody, { bill, status }, true);
    });
}

async function loadPayments() {
    const { payments } = await fetchPayments();
    paymentsState.items = Array.isArray(payments) ? payments : [];
    paymentsState.paymentIndex = buildPaymentIndex(paymentsState.items);
    paymentsState.loaded = true;
}

async function loadGeneratedBills(
    status = paymentsState.activeBillTab || "pending",
    force = false,
    options = {}
) {
    const bucket = status === "paid" ? "paid" : "pending";
    if (!force && isBillsLoaded(bucket)) {
        renderGeneratedBills();
        return;
    }

    const loader = document.getElementById("generatedBillsLoader");
    if (loader) loader.classList.remove("hidden");

    try {
        const monthsBack =
            typeof options.monthsBack === "number"
                ? options.monthsBack
                : bucket === "paid"
                ? computeMonthsBackForRange(
                      paymentsState.paidFilters?.fromMonth,
                      paymentsState.paidFilters?.toMonth
                  )
                : 0;
        const payload = await fetchBillsMinimal(bucket, { monthsBack });
        const bills = payload && payload.bills;
        const normalizedBills = await enrichBillsWithUnits(Array.isArray(bills) ? bills : [], { allowFetch: false });
        const trimmedBills = normalizedBills.map(trimBillForList);
        setBillsForStatus(bucket, trimmedBills);
        setBillsLoaded(bucket, true);
        if (bucket === "pending") {
            paymentsState.pendingPagination.page = 1;
            paymentsState.pendingPagination.pageSize = PENDING_PAGE_SIZE;
        }
        renderGeneratedBills();
    } catch (err) {
        console.error("Failed to load generated bills", err);
        setBillsForStatus(bucket, []);
        setBillsLoaded(bucket, false);
        renderGeneratedBills();
        showToast("Unable to load generated bills right now.", "error");
    } finally {
        if (loader) loader.classList.add("hidden");
    }
}

function changePendingPage(delta) {
    const pendingBills = filterPendingBills(getBillsForStatus("pending"));
    if (!pendingBills.length) return;
    const { totalPages } = clampPendingPage(pendingBills.length);
    const next = Math.min(
        Math.max(1, (paymentsState.pendingPagination.page || 1) + delta),
        totalPages
    );
    if (next === paymentsState.pendingPagination.page) return;
    paymentsState.pendingPagination.page = next;
    renderGeneratedBills();
}

async function handlePaidSearch() {
    const searchBtn = document.getElementById("paidBillsSearchBtn");
    if (searchBtn) {
        searchBtn.disabled = true;
        searchBtn.classList.add("opacity-60");
    }
    const filters = readPaidFiltersFromUi();
    paymentsState.paidFilters = {
        ...paymentsState.paidFilters,
        ...filters,
        hasSearched: true,
    };
    const monthsBack = computeMonthsBackForRange(filters.fromMonth, filters.toMonth);
    try {
        await loadGeneratedBills("paid", true, { monthsBack });
    } finally {
        if (searchBtn) {
            searchBtn.disabled = false;
            searchBtn.classList.remove("opacity-60");
        }
    }
}

async function handleSavePayment() {
    const dateInput = document.getElementById("paymentDateInput");
    const modeSelect = document.getElementById("paymentModeSelect");
    const notesInput = document.getElementById("paymentNotesInput");
    const idField = document.getElementById("paymentRecordId");
    const amountInput = document.getElementById("paymentAmountInput");
    const context = paymentsState.billContext || {};

    if (paymentsState.viewOnly) {
        closePaymentModal();
        return;
    }

    if (!context.monthKey || !context.tenantName) {
        showToast("Open the modal from a generated bill to link the payment", "warning");
        return;
    }
    if (paymentsState.uploadingAttachment) {
        showToast("Receipt image is still uploading. Please wait for it to finish.", "warning");
        return;
    }

    const dateValue = dateInput?.value || new Date().toISOString().slice(0, 10);
    const rawAmount = amountInput?.value || "";
    const amountVal = Number(rawAmount);
    const pendingAmount =
        typeof context.remaining === "number"
            ? context.remaining
            : typeof context.billTotal === "number"
            ? context.billTotal
            : 0;

    if (!Number.isFinite(amountVal) || amountVal <= 0) {
        showToast("Enter a payment amount greater than 0", "error");
        return;
    }
    if (pendingAmount > 0 && amountVal - pendingAmount > 0.005) {
        showToast(`Amount cannot exceed ${formatCurrency(pendingAmount)} pending`, "error");
        return;
    }

    const payload = {
        id: idField?.value || paymentsState.editingId || "",
        tenantKey: context.tenantKey || context.tenantName,
        tenantName: context.tenantName,
        wing: context.wing || "",
        monthKey: context.monthKey,
        monthLabel: context.monthLabel,
        billTotal: context.billTotal,
        rentAmount: context.rentAmount,
        electricityAmount: context.electricityAmount,
        motorShare: context.motorShare ?? context.motorAmount ?? 0,
        sweepAmount: context.sweepAmount,
        prevReading: context.prevReading,
        newReading: context.newReading,
        payableDate: context.payableDate,
        amount: amountVal,
        mode: modeSelect?.value || "",
        notes: notesInput?.value || "",
        date: dateValue,
        attachmentDataUrl: paymentsState.attachmentDataUrl,
        attachmentName: paymentsState.attachmentName,
        attachmentUrl: paymentsState.attachmentUrl,
        attachmentId: paymentsState.attachmentId,
        tenancyId: context.tenancyId,
        billLineId: context.billLineId,
    };

    const { ok, payment } = await savePaymentRecord(payload);
    if (!ok) return;

    const savedPayment = {
        ...payload,
        ...(payment || {}),
    };
    savedPayment.id = savedPayment.id || paymentsState.editingId || `payment-${Date.now()}`;
    savedPayment.amount = Number(savedPayment.amount) || 0;
    savedPayment.monthKey = normalizeMonthKey(savedPayment.monthKey);
    savedPayment.tenantKey = normalizeKey(savedPayment.tenantKey || savedPayment.tenantName);

    const existingIdx = paymentsState.items.findIndex((p) => p.id === savedPayment.id);
    if (existingIdx >= 0) {
        paymentsState.items[existingIdx] = savedPayment;
    } else {
        paymentsState.items.unshift(savedPayment);
    }

    paymentsState.paymentIndex = buildPaymentIndex(paymentsState.items);

    paymentsState.editingId = savedPayment.id;
    paymentsState.attachmentDataUrl = "";
    paymentsState.attachmentName = savedPayment.attachmentName || "";
    paymentsState.attachmentUrl = normalizeAttachmentUrl(savedPayment.attachmentUrl || "");
    paymentsState.attachmentId = savedPayment.attachmentId || paymentsState.attachmentId || "";
    paymentsState.attachmentCommitted = true;

    const remainingAfter =
        Math.max(
            0,
            (Number(context.remaining) || Number(context.billTotal) || 0) - (Number(savedPayment.amount) || 0)
        ) || 0;

    const activeTab = paymentsState.activeBillTab || "pending";
    setBillsTab(activeTab, { skipLoad: true });
    closePaymentModal();
    await loadGeneratedBills(activeTab, true);
}

function isPaymentModalVisible() {
    const modal = document.getElementById("paymentModal");
    return !!modal && !modal.classList.contains("hidden");
}

function cancelAttachmentUpload() {
    if (paymentsState.uploadController) {
        paymentsState.uploadController.cancelled = true;
        if (typeof paymentsState.uploadController.abort === "function") {
            paymentsState.uploadController.abort();
        }
    }
    paymentsState.uploadController = null;
    paymentsState.uploadingAttachment = false;
    paymentsState.uploadProgress = 0;
    const progressWrap = document.getElementById("paymentAttachmentProgress");
    const progressFill = document.getElementById("paymentAttachmentProgressFill");
    const progressLabel = document.getElementById("paymentAttachmentProgressLabel");
    if (progressWrap) progressWrap.classList.add("hidden");
    if (progressFill) progressFill.style.width = "0%";
    if (progressLabel) progressLabel.textContent = "";
}

function clearAttachmentState() {
    paymentsState.attachmentId = "";
    paymentsState.attachmentDataUrl = "";
    paymentsState.attachmentName = "";
    paymentsState.attachmentUrl = "";
    paymentsState.attachmentPreviewUrl = "";
    paymentsState.attachmentViewUrl = "";
    paymentsState.attachmentCommitted = false;
}

async function discardUploadedAttachment() {
    const committed = paymentsState.attachmentCommitted;
    const id = paymentsState.attachmentId;
    clearAttachmentState();
    if (!id || committed) return;
    try {
        await deleteAttachment(id);
    } catch (err) {
        console.warn("Unable to delete attachment", err);
    }
}

function setAttachmentUploadProgress(percent, label) {
    const progressWrap = document.getElementById("paymentAttachmentProgress");
    const progressFill = document.getElementById("paymentAttachmentProgressFill");
    const progressLabel = document.getElementById("paymentAttachmentProgressLabel");
    if (!progressWrap || !progressFill || !progressLabel) return;
    if (percent === null || percent === undefined) {
        progressWrap.classList.add("hidden");
        progressFill.style.width = "0%";
        progressLabel.textContent = "";
        paymentsState.uploadProgress = 0;
        return;
    }
    paymentsState.uploadProgress = Math.max(0, Math.min(100, percent));
    progressWrap.classList.remove("hidden");
    progressFill.style.width = `${paymentsState.uploadProgress}%`;
    progressLabel.textContent = label || "Uploading...";
}

function approxBytesFromDataUrl(dataUrl) {
    if (!dataUrl) return 0;
    const comma = dataUrl.indexOf(",");
    if (comma === -1) return 0;
    const base64 = dataUrl.slice(comma + 1);
    return Math.floor((base64.length * 3) / 4);
}

function isImageMimeType(mimeType) {
    return (mimeType || "").toLowerCase().startsWith("image/");
}

function isImageFileName(name) {
    return /\.(png|jpe?g|gif|webp|bmp)$/i.test(name || "");
}

function isImageDataUrl(dataUrl) {
    return /^data:image\//i.test(dataUrl || "");
}

function getMimeFromDataUrl(dataUrl, fallback = "") {
    const match = /^data:([^;]+);/i.exec(dataUrl || "");
    return match?.[1] || fallback;
}

function normalizeImageDataUrl(dataUrl, mimeTypeHint) {
    if (!dataUrl) return dataUrl;
    if (isImageDataUrl(dataUrl)) return dataUrl;
    const mimeType = isImageMimeType(mimeTypeHint) ? mimeTypeHint : "";
    if (!mimeType) return dataUrl;
    const comma = dataUrl.indexOf(",");
    if (comma === -1) return dataUrl;
    const payload = dataUrl.slice(comma + 1);
    return `data:${mimeType};base64,${payload}`;
}

function readBlobAsDataUrl(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result || "");
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(blob);
    });
}

function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result || "");
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(file);
    });
}

function loadImageFromFile(file, controller) {
    return new Promise((resolve, reject) => {
        if (!file) {
            reject(new Error("missing file"));
            return;
        }
        if (controller?.cancelled) {
            reject(new Error("cancelled"));
            return;
        }
        const objectUrl = URL.createObjectURL(file);
        const img = new Image();
        const finalize = (handler) => (value) => {
            URL.revokeObjectURL(objectUrl);
            if (controller?.cancelled) {
                reject(new Error("cancelled"));
                return;
            }
            handler(value);
        };
        img.onload = finalize(() => resolve(img));
        img.onerror = finalize((err) => reject(err));
        img.src = objectUrl;
    });
}

async function canvasToDataUrl(canvas, mimeType, quality) {
    if (typeof canvas.toBlob !== "function") {
        const fallbackDataUrl = canvas.toDataURL(mimeType, quality);
        const resolvedMime = getMimeFromDataUrl(fallbackDataUrl, mimeType);
        return {
            dataUrl: fallbackDataUrl,
            bytes: approxBytesFromDataUrl(fallbackDataUrl),
            mimeType: resolvedMime,
        };
    }
    return new Promise((resolve) => {
        canvas.toBlob(
            async (blob) => {
                if (!blob) {
                    const fallbackDataUrl = canvas.toDataURL(mimeType, quality);
                    const resolvedMime = getMimeFromDataUrl(fallbackDataUrl, mimeType);
                    resolve({
                        dataUrl: fallbackDataUrl,
                        bytes: approxBytesFromDataUrl(fallbackDataUrl),
                        mimeType: resolvedMime,
                    });
                    return;
                }
                try {
                    const dataUrl = await readBlobAsDataUrl(blob);
                    resolve({
                        dataUrl,
                        bytes: blob.size,
                        mimeType: blob.type || mimeType,
                    });
                } catch (err) {
                    const fallbackDataUrl = canvas.toDataURL(mimeType, quality);
                    resolve({
                        dataUrl: fallbackDataUrl,
                        bytes: approxBytesFromDataUrl(fallbackDataUrl),
                        mimeType,
                    });
                }
            },
            mimeType,
            quality
        );
    });
}

function getExtensionForMime(mimeType) {
    const normalized = (mimeType || "").toLowerCase();
    if (normalized === "image/jpeg" || normalized === "image/jpg") return "jpg";
    if (normalized === "image/png") return "png";
    if (normalized === "image/webp") return "webp";
    return "";
}

function buildAttachmentName(originalName, mimeType) {
    const extension = getExtensionForMime(mimeType);
    const baseName = (originalName || "receipt").replace(/\.[^.]+$/, "");
    if (!extension) {
        return originalName || "receipt";
    }
    return `${baseName}.${extension}`;
}

function buildQualityCandidates(primary) {
    const candidates = [primary, 0.6, 0.52, 0.45];
    const unique = [];
    candidates.forEach((quality) => {
        const clamped = Math.max(0.35, Math.min(0.9, Number(quality) || 0));
        if (!unique.some((value) => Math.abs(value - clamped) < 0.001)) {
            unique.push(clamped);
        }
    });
    return unique;
}

function dataUrlToFile(dataUrl, baseName) {
    const comma = dataUrl.indexOf(",");
    if (comma === -1) {
        return null;
    }
    const header = dataUrl.slice(0, comma);
    const data = dataUrl.slice(comma + 1);
    const isBase64 = /;base64/i.test(header);
    const mimeType = getMimeFromDataUrl(dataUrl, "");
    let bytes;
    if (isBase64) {
        const binary = atob(data);
        const len = binary.length;
        bytes = new Uint8Array(len);
        for (let i = 0; i < len; i += 1) {
            bytes[i] = binary.charCodeAt(i);
        }
    } else {
        const decoded = decodeURIComponent(data);
        bytes = new TextEncoder().encode(decoded);
    }
    const blob = new Blob([bytes], { type: mimeType || "" });
    const extension = getExtensionForMime(mimeType) || "png";
    const name = `${baseName || "clipboard-image"}.${extension}`;
    if (typeof File === "function") {
        return new File([blob], name, { type: mimeType || "" });
    }
    return Object.assign(blob, { name });
}

function extractImageDataUrl(clipboard) {
    if (!clipboard?.getData) return "";
    const html = clipboard.getData("text/html") || "";
    const htmlMatch = html.match(/data:image\/[a-z0-9.+-]+;base64,[a-z0-9+/=]+/i);
    if (htmlMatch?.[0]) return htmlMatch[0];
    const text = clipboard.getData("text/plain") || "";
    const textMatch = text.match(/data:image\/[a-z0-9.+-]+;base64,[a-z0-9+/=]+/i);
    return textMatch?.[0] || "";
}

function readClipboardStringItem(item) {
    return new Promise((resolve) => {
        if (!item || typeof item.getAsString !== "function") {
            resolve("");
            return;
        }
        try {
            item.getAsString((value) => resolve(value || ""));
        } catch (err) {
            resolve("");
        }
    });
}

async function getClipboardImage(clipboard) {
    const items = Array.from(clipboard?.items || []);
    const itemWithImage = items.find((item) => item.kind === "file" && isImageMimeType(item.type));
    let file = itemWithImage?.getAsFile() || null;
    let mimeTypeHint = itemWithImage?.type || "";
    if (!file) {
        const fileFromFiles = Array.from(clipboard?.files || []).find((candidate) => {
            if (!candidate) return false;
            return isImageMimeType(candidate.type) || isImageFileName(candidate.name);
        });
        if (fileFromFiles) {
            file = fileFromFiles;
            mimeTypeHint = fileFromFiles.type || "";
        }
    }
    if (!file) {
        const stringImageItem = items.find((item) => item.kind === "string" && isImageMimeType(item.type));
        if (stringImageItem) {
            const value = await readClipboardStringItem(stringImageItem);
            const base = value && !/^data:/i.test(value) && isImageMimeType(stringImageItem.type)
                ? `data:${stringImageItem.type};base64,${value}`
                : value;
            const normalized = normalizeImageDataUrl(base, stringImageItem.type);
            file = dataUrlToFile(normalized, "clipboard-image");
            mimeTypeHint = stringImageItem.type || mimeTypeHint;
        }
    }
    if (!file) {
        const dataUrl = extractImageDataUrl(clipboard);
        if (dataUrl) {
            const normalized = normalizeImageDataUrl(dataUrl, mimeTypeHint);
            file = dataUrlToFile(normalized, "clipboard-image");
            mimeTypeHint = getMimeFromDataUrl(normalized, mimeTypeHint);
        }
    }
    if (!file && navigator.clipboard?.read) {
        try {
            const clipboardItems = await navigator.clipboard.read();
            const imageItem = clipboardItems.find((item) =>
                item.types?.some((type) => isImageMimeType(type))
            );
            if (imageItem) {
                const type = imageItem.types.find((t) => isImageMimeType(t)) || "";
                const blob = await imageItem.getType(type);
                const extension = getExtensionForMime(type) || "png";
                const name = `clipboard-image.${extension}`;
                if (typeof File === "function") {
                    file = new File([blob], name, { type });
                } else {
                    file = Object.assign(blob, { name });
                }
                mimeTypeHint = type;
            }
        } catch (err) {
            console.warn("Clipboard read failed", err);
        }
    }

    return { file, mimeTypeHint };
}

async function compressImageInBrowser(file, controller, mimeTypeHint = "") {
    const originalBytes = file?.size || 0;
    const ensureNotCancelled = () => {
        if (controller?.cancelled) {
            throw new Error("cancelled");
        }
    };
    const buildFallback = async () => {
        ensureNotCancelled();
        const fallbackDataUrl = await readFileAsDataUrl(file);
        const resolvedMime = getMimeFromDataUrl(fallbackDataUrl, file?.type || mimeTypeHint || "");
        ensureNotCancelled();
        return {
            dataUrl: normalizeImageDataUrl(fallbackDataUrl, resolvedMime || mimeTypeHint),
            bytes: originalBytes || approxBytesFromDataUrl(fallbackDataUrl),
            mimeType: resolvedMime,
        };
    };

    ensureNotCancelled();
    let image = null;
    try {
        image = await loadImageFromFile(file, controller);
    } catch (err) {
        if (controller?.cancelled) throw new Error("cancelled");
        console.warn("Image load failed, using original image", err);
        return buildFallback();
    }

    ensureNotCancelled();
    const width = image.width || 0;
    const height = image.height || 0;
    if (!width || !height) {
        return buildFallback();
    }
    const scale = Math.min(1, RECEIPT_MAX_DIM / Math.max(width, height));
    const targetWidth = Math.max(1, Math.round(width * scale));
    const targetHeight = Math.max(1, Math.round(height * scale));

    const canvas = document.createElement("canvas");
    canvas.width = targetWidth;
    canvas.height = targetHeight;
    const ctx = canvas.getContext("2d", { alpha: false });
    if (!ctx) {
        return buildFallback();
    }

    if (scale < 1) {
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = "high";
    }
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, targetWidth, targetHeight);
    ctx.drawImage(image, 0, 0, targetWidth, targetHeight);

    let best = null;
    const qualities = buildQualityCandidates(RECEIPT_JPEG_QUALITY);
    for (const quality of qualities) {
        ensureNotCancelled();
        const attempt = await canvasToDataUrl(canvas, RECEIPT_OUTPUT_MIME, quality);
        ensureNotCancelled();
        if (!attempt?.dataUrl) continue;
        const attemptBytes = Number(attempt.bytes) || approxBytesFromDataUrl(attempt.dataUrl);
        const resolved = {
            dataUrl: attempt.dataUrl,
            bytes: attemptBytes,
            mimeType: attempt.mimeType || RECEIPT_OUTPUT_MIME,
        };
        if (!best || attemptBytes < best.bytes) {
            best = resolved;
        }
        if (attemptBytes && attemptBytes <= RECEIPT_TARGET_BYTES) {
            break;
        }
    }

    if (!best?.dataUrl) {
        return buildFallback();
    }

    const optimizedBytes = best.bytes || approxBytesFromDataUrl(best.dataUrl);
    if (originalBytes && optimizedBytes && optimizedBytes >= originalBytes) {
        return buildFallback();
    }

    return {
        dataUrl: best.dataUrl,
        bytes: optimizedBytes,
        mimeType: best.mimeType || RECEIPT_OUTPUT_MIME,
    };
}

async function updateAttachmentFromFile(file, options = {}) {
    const mimeTypeHint = options.mimeTypeHint || "";
    const allowUnknown = !!options.allowUnknown;
    if (!file) return false;
    const hasName = !!file.name;
    const hasType = !!file.type;
    const isImageHint = isImageMimeType(file.type) || isImageMimeType(mimeTypeHint) || isImageFileName(file.name);
    if (!allowUnknown && !isImageHint && (hasName || hasType)) {
        showToast("Please use an image file for the receipt", "warning");
        return false;
    }

    const context = paymentsState.billContext || {};
    if (!context.monthKey || !context.tenantName) {
        showToast("Open the payment modal from a bill before adding a receipt", "warning");
        return false;
    }

    if (paymentsState.attachmentId && !paymentsState.attachmentCommitted) {
        await discardUploadedAttachment();
    }

    cancelAttachmentUpload();
    const controller = { cancelled: false };
    paymentsState.uploadController = controller;
    paymentsState.uploadingAttachment = true;
    setAttachmentUploadProgress(5, "Reading image...");
    const attachmentPreview = document.getElementById("paymentAttachmentPreview");
    if (attachmentPreview) attachmentPreview.classList.remove("hidden");

    let dataUrl = "";
    let compressedBytes = 0;
    let uploadName = file.name || "receipt";
    try {
        setAttachmentUploadProgress(35, "Optimizing image...");
        const optimized = await compressImageInBrowser(file, controller, mimeTypeHint);
        if (controller.cancelled) throw new Error("cancelled");
        dataUrl = optimized.dataUrl || "";
        if (!isImageDataUrl(dataUrl) && !isImageMimeType(optimized.mimeType || mimeTypeHint)) {
            throw new Error("invalid-image");
        }
        compressedBytes = Number(optimized.bytes) || approxBytesFromDataUrl(dataUrl);
        uploadName = buildAttachmentName(
            file.name || "receipt",
            optimized.mimeType || file.type || mimeTypeHint || ""
        );
        const approxKb = Math.max(1, Math.round(compressedBytes / 1024));
        setAttachmentUploadProgress(55, `Optimized ~${approxKb} KB`);
        if (dataUrl && isImageDataUrl(dataUrl)) {
            const linkWrap = document.getElementById("paymentAttachmentLink");
            if (linkWrap) linkWrap.classList.add("hidden");
            const preview = document.getElementById("paymentAttachmentPreview");
            const previewImg = preview?.querySelector("img");
            if (previewImg) {
                previewImg.referrerPolicy = "no-referrer";
                previewImg.src = dataUrl;
                previewImg.classList.remove("hidden");
            }
            if (preview) preview.classList.remove("hidden");
            paymentsState.attachmentPreviewUrl = dataUrl;
            paymentsState.attachmentViewUrl = dataUrl;
        }
    } catch (err) {
        if (controller.cancelled || err?.message === "cancelled") {
            cancelAttachmentUpload();
            return false;
        }
        if (err?.message === "invalid-image") {
            showToast("Clipboard data isn't an image. Please paste an image.", "warning");
            cancelAttachmentUpload();
            return false;
        }
        console.warn("Unable to read attachment", err);
        showToast("Could not read the image. Please try again.", "error");
        cancelAttachmentUpload();
        return false;
    }
    if (controller.cancelled) {
        cancelAttachmentUpload();
        return false;
    }

    const totalBytes = compressedBytes || approxBytesFromDataUrl(dataUrl);
    const totalKb = totalBytes ? Math.max(1, Math.round(totalBytes / 1024)) : 0;
    const uploadBase = 60;
    const uploadSpan = 35;
    const uploadId = `receipt-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    setAttachmentUploadProgress(uploadBase, totalKb ? `Uploading 0 / ${totalKb} KB` : "Uploading...");
    const uploadResult = await uploadPaymentAttachment(
        {
            dataUrl,
            attachmentName: uploadName || file.name || "",
            tenantName: context.tenantName,
            monthKey: context.monthKey,
            paymentId: paymentsState.editingId || `temp-${Date.now()}`,
        },
        {
            uploadId,
            onProgress: ({ loaded, total, percent }) => {
                if (controller.cancelled) return;
                const resolvedTotal = total || totalBytes || 0;
                const loadedKb = Math.round((loaded || 0) / 1024);
                const resolvedKb = resolvedTotal ? Math.max(1, Math.round(resolvedTotal / 1024)) : 0;
                let progress = uploadBase;
                if (typeof percent === "number") {
                    progress = uploadBase + (percent / 100) * uploadSpan;
                } else if (resolvedTotal) {
                    progress = uploadBase + Math.min(1, (loaded || 0) / resolvedTotal) * uploadSpan;
                }
                const label = resolvedKb
                    ? `Uploading ${loadedKb} / ${resolvedKb} KB`
                    : `Uploading ${loadedKb} KB`;
                setAttachmentUploadProgress(Math.min(99, progress), label);
            },
            onAbort: (abortFn) => {
                controller.abort = abortFn;
            },
        }
    );

    if (controller.cancelled) {
        if (uploadResult?.attachment?.attachment_id) {
            try {
                await deleteAttachment(uploadResult.attachment.attachment_id);
            } catch (err) {
                console.warn("Unable to delete attachment", err);
            }
        }
        cancelAttachmentUpload();
        return false;
    }

    if (!uploadResult?.ok || !uploadResult.attachment) {
        if (controller.cancelled) {
            cancelAttachmentUpload();
            return false;
        }
        showToast("Failed to upload receipt. Please try again.", "error");
        cancelAttachmentUpload();
        return false;
    }

    const { attachment_id, attachmentUrl, attachmentName } = uploadResult.attachment;
    clearAttachmentState();
    paymentsState.attachmentId = attachment_id || "";
    paymentsState.attachmentName = attachmentName || uploadName || file.name || "";
    paymentsState.attachmentUrl = attachmentUrl || "";
    paymentsState.attachmentCommitted = false;
    await showAttachmentPreview(paymentsState.attachmentName, paymentsState.attachmentUrl || dataUrl || "");
    setAttachmentUploadProgress(100, "Uploaded");
    setTimeout(() => setAttachmentUploadProgress(null), 600);
    paymentsState.uploadController = null;
    paymentsState.uploadingAttachment = false;
    return true;
}

function wireAttachmentHandlers() {
    const input = document.getElementById("paymentAttachmentInput");
    const clearBtn = document.getElementById("paymentAttachmentClear");
    const nameLabel = document.getElementById("paymentAttachmentName");
    const preview = document.getElementById("paymentAttachmentPreview");
    const viewLink = document.querySelector("#paymentAttachmentLink a");
    const modal = document.getElementById("paymentModal");
    const pasteArea = document.getElementById("paymentAttachmentArea");
    const pasteZone = document.getElementById("paymentAttachmentPasteZone");
    const pasteHint = document.getElementById("paymentAttachmentPasteHint");
    const viewerCloseBtn = document.getElementById("attachmentViewerClose");
    const viewerBackdrop = document.getElementById("attachmentViewerModal");

    if (input) {
        input.addEventListener("change", async (e) => {
            const file = e.target.files?.[0];
            const loaded = await updateAttachmentFromFile(file);
            if (!loaded) {
                input.value = "";
                return;
            }
            if (nameLabel) nameLabel.textContent = paymentsState.attachmentName || "";
        });
    }

    if (clearBtn) {
        clearBtn.addEventListener("click", async () => {
            cancelAttachmentUpload();
            await discardUploadedAttachment();
            if (input) input.value = "";
            clearAttachmentState();
            await showAttachmentPreview("No file selected", "");
            if (nameLabel) nameLabel.textContent = "No file selected";
            if (pasteHint)
                pasteHint.textContent = "Click inside this box and press Ctrl+V (or Cmd+V) to attach an image from your clipboard.";
        });
    }

    let pasteFallbackTimer = null;
    const handleClipboardImage = async (clipboard, event) => {
        if (!isPaymentModalVisible()) return false;
        if (event && event.__paymentPasteHandled) return false;
        if (event) event.__paymentPasteHandled = true;
        const { file, mimeTypeHint } = await getClipboardImage(clipboard);
        if (!file) return false;

        if (event?.preventDefault) event.preventDefault();
        const loaded = await updateAttachmentFromFile(file, { mimeTypeHint, allowUnknown: true });
        if (loaded) {
            if (nameLabel) nameLabel.textContent = paymentsState.attachmentName || "";
            if (input) input.value = "";
            if (pasteHint) pasteHint.textContent = "Image added from clipboard. Paste again to replace.";
            if (pasteZone) pasteZone.blur();
        }
        return loaded;
    };

    const handlePaste = (e) => {
        if (pasteFallbackTimer) {
            clearTimeout(pasteFallbackTimer);
            pasteFallbackTimer = null;
        }
        handleClipboardImage(e.clipboardData, e);
    };

    if (modal) modal.addEventListener("paste", handlePaste);
    if (pasteArea) pasteArea.addEventListener("paste", handlePaste);
    if (pasteZone) {
        pasteZone.addEventListener("paste", handlePaste);
        pasteZone.addEventListener("beforeinput", (e) => {
            if (e.inputType !== "insertFromPaste") {
                e.preventDefault();
            }
        });
        pasteZone.addEventListener("keydown", (e) => {
            if (!e.ctrlKey && !e.metaKey && e.key.length === 1) {
                e.preventDefault();
            }
            if (e.key.toLowerCase() === "v" && (e.ctrlKey || e.metaKey)) {
                if (pasteFallbackTimer) clearTimeout(pasteFallbackTimer);
                pasteFallbackTimer = setTimeout(() => {
                    handleClipboardImage(null, null);
                }, 80);
            }
        });
        pasteZone.addEventListener("drop", (e) => e.preventDefault());
        pasteZone.addEventListener("focus", () => {
            if (pasteHint)
                pasteHint.textContent =
                    "Press Ctrl+V (or Cmd+V) here to drop a screenshot. It will save automatically.";
        });
    }

    if (preview) {
        preview.addEventListener("click", () => openAttachmentViewerFromState());
    }

    if (viewLink) {
        viewLink.addEventListener("click", (e) => {
            e.preventDefault();
            openAttachmentViewerFromState();
        });
    }

    if (viewerCloseBtn) viewerCloseBtn.addEventListener("click", closeAttachmentViewer);
    if (viewerBackdrop) {
        viewerBackdrop.addEventListener("click", (e) => {
            if (e.target === viewerBackdrop) {
                closeAttachmentViewer();
            }
        });
    }
}

/**
 * Boots the Payments tab by loading data, connecting click handlers,
 * and syncing bill tabs with the payment form.
 */
export function initPaymentsFeature() {
    const refreshBtn = document.getElementById("paymentsRefreshBtn");
    if (refreshBtn) refreshBtn.addEventListener("click", () => {
        refreshPaymentsIfNeeded(true);
    });

    const paidSearchBtn = document.getElementById("paidBillsSearchBtn");
    if (paidSearchBtn) paidSearchBtn.addEventListener("click", handlePaidSearch);

    const pendingPrev = document.getElementById("pendingBillsPrev");
    const pendingNext = document.getElementById("pendingBillsNext");
    if (pendingPrev) pendingPrev.addEventListener("click", () => changePendingPage(-1));
    if (pendingNext) pendingNext.addEventListener("click", () => changePendingPage(1));

    const paidLimit = document.getElementById("paidBillsLimit");
    if (paidLimit) {
        paidLimit.value = String(paymentsState.paidFilters?.limit || PAID_DEFAULT_LIMIT);
    }
    paymentsState.pendingPagination.pageSize = PENDING_PAGE_SIZE;
    [document.getElementById("paidBillsFrom"), document.getElementById("paidBillsTo"), paidLimit].forEach((el) => {
        if (!el) return;
        el.addEventListener("change", () => {
            paymentsState.paidFilters.hasSearched = false;
            if (paymentsState.activeBillTab === "paid") {
                renderGeneratedBills();
            }
        });
    });

    const pendingTab = document.getElementById("billsPendingTab");
    const paidTab = document.getElementById("billsPaidTab");
    if (pendingTab) pendingTab.addEventListener("click", () => setBillsTab("pending"));
    if (paidTab) paidTab.addEventListener("click", () => setBillsTab("paid"));
    setBillsTab(paymentsState.activeBillTab || "pending");

    const closeBtn = document.getElementById("paymentModalClose");
    const cancelBtn = document.getElementById("paymentModalCancel");
    if (closeBtn) closeBtn.addEventListener("click", closePaymentModal);
    if (cancelBtn) cancelBtn.addEventListener("click", closePaymentModal);

    const saveBtn = document.getElementById("paymentSaveBtn");
    if (saveBtn) saveBtn.addEventListener("click", handleSavePayment);
    wireAttachmentHandlers();
}

export async function refreshPaymentsIfNeeded(force = false) {
    const activeTab = paymentsState.activeBillTab || "pending";
    const deferPaidLoad = activeTab === "paid" && !paymentsState.paidFilters?.hasSearched;
    const shouldLoadBills = force || (!isBillsLoaded(activeTab) && !deferPaidLoad);
    if (shouldLoadBills) {
        await loadGeneratedBills(activeTab, force);
    } else {
        renderGeneratedBills();
    }
}





