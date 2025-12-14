/**
 * Payments Feature Module
 *
 * Displays recorded payments, allows adding new receipts, and links
 * payments to generated bills and tenant directory context.
 */

import { fetchAttachmentPreview, fetchGeneratedBills, fetchPayments, savePaymentRecord } from "../../api/sheets.js";
import { ensureTenantDirectoryLoaded } from "../tenants/tenants.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

const paymentsState = {
    items: [],
    loaded: false,
    generatedBills: [],
    billsLoaded: false,
    paymentIndex: createPaymentIndex(),
    activeBillTab: "pending",
    billContext: {
        monthKey: "",
        monthLabel: "",
        billTotal: 0,
        billed: false,
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
    editingId: "",
    viewOnly: false,
};

const attachmentPreviewCache = new Map();

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
    const num = Number(amount);
    if (isNaN(num)) return amount || "-";
    return `₹${num.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function summarizeBillPayments(bill) {
    const { matches } = getPaymentsForBill(bill);

    const paidAmount = matches.reduce((sum, p) => sum + (Number(p.amount) || 0), 0);
    const remaining = Math.max(0, (Number(bill.totalAmount) || 0) - paidAmount);

    return {
        paidAmount,
        remaining,
        receiptCount: matches.length,
    };
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
    if (!value) return "";

    const str = value.toString().trim();
    if (!str) return "";

    const compactMatch = str.match(/^(\d{4})(\d{2})$/);
    if (compactMatch) {
        const month = compactMatch[2].padStart(2, "0");
        return `${compactMatch[1]}-${month}`.toLowerCase();
    }

    const match = str.match(/^(\d{4})[-/.](\d{1,2})/);
    if (match) {
        const month = match[2].padStart(2, "0");
        return `${match[1]}-${month}`.toLowerCase();
    }

    return str.toLowerCase();
}

function togglePaymentModal(show) {
    const modal = document.getElementById("paymentModal");
    if (modal) {
        show ? showModal(modal) : hideModal(modal);
    }
}

function resetPaymentForm() {
    paymentsState.editingId = "";
    paymentsState.attachmentDataUrl = "";
    paymentsState.attachmentName = "";
    paymentsState.attachmentUrl = "";
    paymentsState.attachmentPreviewUrl = "";
    paymentsState.attachmentViewUrl = "";
    paymentsState.viewOnly = false;
    paymentsState.billContext = {
        monthKey: "",
        monthLabel: "",
        billTotal: 0,
        billed: false,
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
    const billTitle = document.getElementById("paymentBillTitle");
    const billAmount = document.getElementById("paymentBillAmount");
    const billStatus = document.getElementById("paymentBillStatus");
    const billMonth = document.getElementById("paymentBillMonth");
    const wingBadge = document.getElementById("paymentTenantWing");
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

    if (dateInput) dateInput.value = todayIso;
    if (modeSelect) modeSelect.value = "UPI";
    if (notesInput) notesInput.value = "";
    if (idField) idField.value = "";

    if (billTitle) billTitle.textContent = "Select a bill to record";
    if (billAmount) billAmount.textContent = "₹0";
    if (billStatus) billStatus.textContent = "Open this modal from a bill card to link the payment automatically.";
    if (billMonth) billMonth.textContent = "Month • -";
    if (wingBadge) wingBadge.textContent = "Wing • -";
    if (breakdownTotal) breakdownTotal.textContent = "₹0";
    if (breakdownRent) breakdownRent.textContent = "₹0";
    if (breakdownElec) breakdownElec.textContent = "₹0";
    if (breakdownMotor) breakdownMotor.textContent = "₹0";
    if (breakdownSweep) breakdownSweep.textContent = "₹0";

    if (attachmentPreview) {
        attachmentPreview.classList.add("hidden");
        const img = attachmentPreview.querySelector("img");
        if (img) img.src = "";
    }
    if (attachmentName) attachmentName.textContent = "No file selected";
    if (attachmentLink) attachmentLink.classList.add("hidden");
    if (attachmentInput) attachmentInput.disabled = false;
    if (attachmentClear) attachmentClear.disabled = false;
    if (saveBtn) saveBtn.textContent = "Save payment";
}

function setPaymentFormReadOnly(isReadOnly) {
    const dateInput = document.getElementById("paymentDateInput");
    const modeSelect = document.getElementById("paymentModeSelect");
    const notesInput = document.getElementById("paymentNotesInput");
    const attachmentInput = document.getElementById("paymentAttachmentInput");
    const attachmentClear = document.getElementById("paymentAttachmentClear");
    const saveBtn = document.getElementById("paymentSaveBtn");

    if (dateInput) dateInput.disabled = isReadOnly;
    if (modeSelect) modeSelect.disabled = isReadOnly;
    if (notesInput) notesInput.disabled = isReadOnly;
    if (attachmentInput) attachmentInput.disabled = isReadOnly;
    if (attachmentClear) attachmentClear.disabled = isReadOnly;
    if (saveBtn) saveBtn.textContent = isReadOnly ? "Close" : "Save payment";
}

function applyBillContext(context = {}) {
    const billTitle = document.getElementById("paymentBillTitle");
    const billAmount = document.getElementById("paymentBillAmount");
    const billStatus = document.getElementById("paymentBillStatus");
    const billMonth = document.getElementById("paymentBillMonth");
    const wingBadge = document.getElementById("paymentTenantWing");
    const breakdownTotal = document.getElementById("paymentBreakdownTotal");
    const breakdownRent = document.getElementById("paymentBreakdownRent");
    const breakdownElec = document.getElementById("paymentBreakdownElectricity");
    const breakdownMotor = document.getElementById("paymentBreakdownMotor");
    const breakdownSweep = document.getElementById("paymentBreakdownSweep");

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
        amount: remaining,
        remaining,
        billed: !!context.monthKey,
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
    };

    if (billTitle) billTitle.textContent = context.tenantName ? `${context.tenantName}` : "Select a bill to record";
    if (billAmount) billAmount.textContent = billTotal ? formatCurrency(remaining || billTotal) : "₹0";
    if (billStatus) {
        const dueLabel = context.payableDate ? ` • Due ${context.payableDate}` : "";
        billStatus.textContent = context.monthKey
            ? `${monthLabel}${context.wing ? ` • ${context.wing}` : ""}${dueLabel}`
            : "Open this modal from a bill row to link automatically.";
    }
    if (billMonth) billMonth.textContent = `Month • ${monthLabel || "-"}`;
    if (wingBadge) wingBadge.textContent = `Wing • ${context.wing || "-"}`;
    if (breakdownTotal) breakdownTotal.textContent = formatCurrency(billTotal || remaining || 0);
    if (breakdownRent) breakdownRent.textContent = formatCurrency(rentAmount);
    if (breakdownElec) breakdownElec.textContent = formatCurrency(electricityAmount);
    if (breakdownMotor) breakdownMotor.textContent = formatCurrency(motorAmount);
    if (breakdownSweep) breakdownSweep.textContent = formatCurrency(sweepAmount);
}

function normalizeAttachmentUrl(url = "") {
    if (!url) return "";

    const driveMatch = url.match(/\/file\/d\/([^/]+)\//);
    if (driveMatch?.[1]) {
        return `https://drive.google.com/uc?export=view&id=${driveMatch[1]}`;
    }

    const idParam = url.match(/[?&]id=([^&#]+)/);
    if (idParam?.[1]) {
        return `https://drive.google.com/uc?export=view&id=${idParam[1]}`;
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

    const viewUrl = url || "";
    caption.textContent = title || "Receipt";

    const isImageLike =
        viewUrl.startsWith("data:image") || /\.(png|jpe?g|gif|webp|bmp|svg)(\?|$)/i.test(viewUrl) ||
        viewUrl.includes("export=view");

    image.classList.add("hidden");
    iframe.classList.add("hidden");
    fallback.classList.add("hidden");

    if (isImageLike) {
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
            if (img) img.src = previewUrl;
            attachmentPreview.classList.remove("hidden");
            attachmentPreview.classList.add("cursor-pointer");
        } else {
            if (img) img.src = "";
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
    await ensureTenantDirectoryLoaded();
    resetPaymentForm();

    const modalTitle = document.getElementById("paymentModalTitle");
    const dateInput = document.getElementById("paymentDateInput");
    const modeSelect = document.getElementById("paymentModeSelect");
    const notesInput = document.getElementById("paymentNotesInput");
    const idField = document.getElementById("paymentRecordId");

    if (modalTitle) modalTitle.textContent = payment ? "Edit payment" : "Record payment";

    let context = billContext || null;
    if (payment) {
        context = {
            monthKey: payment.monthKey,
            monthLabel: payment.monthLabel,
            billTotal: payment.billTotal || payment.amount,
            totalAmount: payment.billTotal || payment.amount,
            tenantName: payment.tenantName,
            tenantKey: payment.tenantKey,
            wing: payment.wing,
            amount: payment.amount,
        };

        const matchingBill = (paymentsState.generatedBills || []).find(
            (bill) => normalizeKey(bill.tenantKey || bill.tenantName) === normalizeKey(payment.tenantKey || payment.tenantName)
                && normalizeMonthKey(bill.monthKey) === normalizeMonthKey(payment.monthKey)
                && (bill.wing || '').toLowerCase() === (payment.wing || '').toLowerCase()
        );

        if (matchingBill) {
            const status = summarizeBillPayments(matchingBill);
            context = {
                ...matchingBill,
                remaining: status.remaining,
            };
        }

        context = {
            ...context,
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
        };
    }

    if (!context && paymentsState.generatedBills.length) {
        const firstBill = paymentsState.generatedBills[0];
        const summary = summarizeBillPayments(firstBill);
        context = { ...firstBill, remaining: summary.remaining };
    }

    if (modalTitle) {
        const isSettled = payment && (context?.remaining ?? 0) <= 0;
        if (isSettled) {
            modalTitle.textContent = "Payment details";
        } else if (payment) {
            modalTitle.textContent = "Edit payment";
        } else {
            modalTitle.textContent = "Record payment";
        }
    }

    if (context) applyBillContext(context);

    if (payment) {
        paymentsState.viewOnly = (context?.remaining ?? 0) <= 0;
        paymentsState.editingId = payment.id;
        paymentsState.attachmentName = payment.attachmentName || "";
        paymentsState.attachmentUrl = normalizeAttachmentUrl(payment.attachmentUrl || "");
        paymentsState.attachmentDataUrl = "";

        if (modeSelect && payment.mode) modeSelect.value = payment.mode;
        if (dateInput) dateInput.value = payment.date || dateInput.value;
        if (notesInput) notesInput.value = payment.notes || "";
        if (idField) idField.value = payment.id || "";

        if (payment.attachmentUrl) {
            await showAttachmentPreview(payment.attachmentName, payment.attachmentUrl);
        }
    }

    setPaymentFormReadOnly(paymentsState.viewOnly);

    togglePaymentModal(true);
}

function openPaymentModalFromBill(bill) {
    const status = summarizeBillPayments(bill);
    const context = {
        ...bill,
        remaining: status.remaining,
    };
    openPaymentModal(null, context);
}

function closePaymentModal() {
    togglePaymentModal(false);
    resetPaymentForm();
}

function setBillsTab(tab) {
    paymentsState.activeBillTab = tab;
    const pendingTab = document.getElementById("billsPendingTab");
    const paidTab = document.getElementById("billsPaidTab");
    if (pendingTab) {
        pendingTab.classList.toggle("bg-white", tab === "pending");
        pendingTab.classList.toggle("text-slate-800", tab === "pending");
    }
    if (paidTab) {
        paidTab.classList.toggle("bg-white", tab === "paid");
        paidTab.classList.toggle("text-slate-800", tab === "paid");
    }
    if (paymentsState.billsLoaded) {
        renderGeneratedBills();
    }
}

function renderGeneratedBills() {
    const pendingBody = document.getElementById("pendingBillsBody");
    const paidBody = document.getElementById("paidBillsBody");
    const loader = document.getElementById("generatedBillsLoader");
    const emptyState = document.getElementById("generatedBillsEmpty");
    if (!pendingBody || !paidBody) return;

    if (loader) loader.classList.add("hidden");
    pendingBody.innerHTML = "";
    paidBody.innerHTML = "";

    const bills = paymentsState.generatedBills || [];
    if (!bills.length) {
        if (emptyState) emptyState.classList.remove("hidden");
        return;
    }

    if (emptyState) emptyState.classList.add("hidden");

    const pending = [];
    const paid = [];

    bills.forEach((bill) => {
        const status = summarizeBillPayments(bill);
        const bucket = status.remaining > 0 ? pending : paid;
        bucket.push({ bill, status });
    });

    const makeRow = (targetBody, { bill, status }, isPaid) => {
        const tr = document.createElement("tr");
        tr.className = "border-b last:border-0 hover:bg-slate-50";

        const monthLabel = friendlyMonthLabel(bill.monthKey, bill.monthLabel);
        const remainingLabel = status.remaining > 0 ? `${formatCurrency(status.remaining)} due` : "Paid";
        const totalLabel = formatCurrency(bill.totalAmount);
        const dueInfo = getDueInfo(bill);
        const dueBadge = bill.payableDate
            ? `<div class="text-[10px] ${dueInfo.overdue && !isPaid ? "text-rose-700" : "text-slate-500"}">${dueInfo.label}</div>`
            : "";

        tr.innerHTML = `
            <td class="px-3 py-2 text-[11px] font-semibold">${bill.tenantName || "-"}</td>
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
                viewBtn.addEventListener("click", () => {
                    const payment = findPaymentForBill(bill);
                    openPaymentModal(payment || null, { ...bill, remaining: status.remaining });
                });
                actionCell.appendChild(viewBtn);
            } else {
                const btn = document.createElement("button");
                btn.className = "px-3 py-1.5 rounded-lg bg-indigo-600 text-white text-[11px] font-semibold hover:bg-indigo-500";
                btn.textContent = "Record payment";
                btn.addEventListener("click", () => openPaymentModalFromBill({ ...bill, remaining: status.remaining }));
                actionCell.appendChild(btn);
            }
        }

        targetBody.appendChild(tr);
    };

    const activeTab = paymentsState.activeBillTab || "pending";
    if (pending.length) {
        pending.forEach((entry) => makeRow(pendingBody, entry, false));
    }
    if (paid.length) {
        paid.forEach((entry) => makeRow(paidBody, entry, true));
    }

    pendingBody.classList.toggle("hidden", activeTab !== "pending");
    paidBody.classList.toggle("hidden", activeTab !== "paid");

    const showEmptyPending = activeTab === "pending" && !pending.length;
    const showEmptyPaid = activeTab === "paid" && !paid.length;
    if (emptyState) {
        emptyState.textContent = showEmptyPaid ? "No paid bills yet." : "No pending bills right now.";
        emptyState.classList.toggle("hidden", pending.length + paid.length > 0 && !showEmptyPending && !showEmptyPaid);
    }
}

async function loadPayments() {
    const { payments } = await fetchPayments();
    paymentsState.items = Array.isArray(payments) ? payments : [];
    paymentsState.paymentIndex = buildPaymentIndex(paymentsState.items);
    paymentsState.loaded = true;
    renderGeneratedBills();
}

async function loadGeneratedBills() {
    const loader = document.getElementById("generatedBillsLoader");
    if (loader) loader.classList.remove("hidden");

    const { bills } = await fetchGeneratedBills();
    paymentsState.generatedBills = Array.isArray(bills) ? bills : [];
    paymentsState.billsLoaded = true;
    setBillsTab(paymentsState.activeBillTab || "pending");
    renderGeneratedBills();
}

async function handleSavePayment() {
    const dateInput = document.getElementById("paymentDateInput");
    const modeSelect = document.getElementById("paymentModeSelect");
    const notesInput = document.getElementById("paymentNotesInput");
    const idField = document.getElementById("paymentRecordId");
    const context = paymentsState.billContext || {};

    if (paymentsState.viewOnly) {
        closePaymentModal();
        return;
    }

    if (!context.monthKey || !context.tenantName) {
        showToast("Open the modal from a generated bill to link the payment", "warning");
        return;
    }

    const dateValue = dateInput?.value || new Date().toISOString().slice(0, 10);
    const amountVal = Number(typeof context.remaining === "number" ? context.remaining : context.billTotal || 0);

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
    };

    const { ok, payment } = await savePaymentRecord(payload);
    if (!ok || !payment) return;

    const existingIdx = paymentsState.items.findIndex((p) => p.id === payment.id);
    if (existingIdx >= 0) {
        paymentsState.items[existingIdx] = payment;
    } else {
        paymentsState.items.unshift(payment);
    }

    paymentsState.paymentIndex = buildPaymentIndex(paymentsState.items);

    paymentsState.editingId = payment.id;
    paymentsState.attachmentDataUrl = "";
    paymentsState.attachmentName = payment.attachmentName || "";
    paymentsState.attachmentUrl = normalizeAttachmentUrl(payment.attachmentUrl || "");

    renderGeneratedBills();
    closePaymentModal();
}

function isPaymentModalVisible() {
    const modal = document.getElementById("paymentModal");
    return !!modal && !modal.classList.contains("hidden");
}

function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result || "");
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(file);
    });
}

async function updateAttachmentFromFile(file) {
    if (!file) return false;
    if (!file.type?.startsWith("image/")) {
        showToast("Please use an image file for the receipt", "warning");
        return false;
    }
    if (file.size > 2 * 1024 * 1024) {
        showToast("Please choose an image under 2 MB", "warning");
        return false;
    }

    let dataUrl = "";
    try {
        dataUrl = await readFileAsDataUrl(file);
    } catch (err) {
        console.warn("Unable to read attachment", err);
        showToast("Could not read the image. Please try again.", "error");
        return false;
    }
    const ext = (file.type?.split("/")[1] || "png").split(";")[0];
    const name = file.name || `pasted-receipt.${ext}`;

    paymentsState.attachmentDataUrl = dataUrl || "";
    paymentsState.attachmentName = name;
    paymentsState.attachmentUrl = dataUrl || "";
    await showAttachmentPreview(name, dataUrl || "");
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
            if (input) input.value = "";
            paymentsState.attachmentDataUrl = "";
            paymentsState.attachmentName = "";
            paymentsState.attachmentUrl = "";
            paymentsState.attachmentPreviewUrl = "";
            paymentsState.attachmentViewUrl = "";
            await showAttachmentPreview("No file selected", "");
            if (nameLabel) nameLabel.textContent = "No file selected";
            if (pasteHint)
                pasteHint.textContent = "Click inside this box and press Ctrl+V (or Cmd+V) to attach an image from your clipboard.";
        });
    }

    const handlePaste = async (e) => {
        if (!isPaymentModalVisible()) return;
        const items = Array.from(e.clipboardData?.items || []);
        const file = items
            .map((item) => (item.kind === "file" ? item.getAsFile() : null))
            .find((f) => f && f.type?.startsWith("image/"));
        if (!file) return;

        e.preventDefault();
        const loaded = await updateAttachmentFromFile(file);
        if (loaded) {
            if (nameLabel) nameLabel.textContent = paymentsState.attachmentName || "";
            if (input) input.value = "";
            if (pasteHint) pasteHint.textContent = "Image added from clipboard. Paste again to replace.";
            if (pasteZone) pasteZone.blur();
        }
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
    const shouldLoad = force || !paymentsState.loaded;
    const shouldLoadBills = force || !paymentsState.billsLoaded;

    if (!shouldLoad) {
        renderGeneratedBills();
    }

    const tasks = [];
    if (shouldLoad) tasks.push(loadPayments());
    if (shouldLoadBills) tasks.push(loadGeneratedBills());
    await Promise.all(tasks);
}
