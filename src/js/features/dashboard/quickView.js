import { LOCAL_KEYS, getLocalEntry, getLocalList } from "../../api/localStore.js";
import { formatCurrency, normalizeMonthKey } from "../../utils/formatters.js";
import { hideModal, showModal } from "../../utils/ui.js";

const MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const quickViewState = {
    initialized: false,
    loading: false,
    escapeBound: false,
    monthlyRangeFrom: "",
    monthlyRangeTo: "",
    yearlyRangeFrom: "",
    yearlyRangeTo: "",
    metrics: {
        currentYear: new Date().getFullYear(),
        totalGeneratedAllTime: 0,
        totalCollectedAllTime: 0,
        totalGeneratedCurrentYear: 0,
        totalCollectedCurrentYear: 0,
        outstandingAllTime: 0,
        outstandingCurrentYear: 0,
        collectionRateAllTime: 0,
        collectionRateCurrentYear: 0,
        monthlyByYear: {},
        lastFiveYears: [],
        availableYears: [],
        yearlyAll: [],
        paymentCount: 0,
    },
    error: "",
};

function getElements() {
    return {
        widget: document.getElementById("quickViewWidget"),
        widgetAllGenerated: document.getElementById("quickViewAllGenerated"),
        widgetAllCollected: document.getElementById("quickViewAllCollected"),
        widgetCurrentGenerated: document.getElementById("quickViewCurrentGenerated"),
        widgetCurrentCollected: document.getElementById("quickViewCurrentCollected"),
        widgetStatus: document.getElementById("quickViewStatus"),
        widgetRefreshBtn: document.getElementById("quickViewRefreshBtn"),
        widgetExpandBtn: document.getElementById("quickViewExpand"),
        modal: document.getElementById("quickViewModal"),
        modalCloseBtn: document.getElementById("quickViewModalClose"),
        modalRefreshBtn: document.getElementById("quickViewModalRefreshBtn"),
        modalAllGenerated: document.getElementById("quickViewModalAllGenerated"),
        modalAllCollected: document.getElementById("quickViewModalAllCollected"),
        modalCurrentGenerated: document.getElementById("quickViewModalCurrentGenerated"),
        modalCurrentCollected: document.getElementById("quickViewModalCurrentCollected"),
        modalOutstandingAll: document.getElementById("quickViewModalOutstandingAll"),
        modalOutstandingCurrent: document.getElementById("quickViewModalOutstandingCurrent"),
        modalCollectionAll: document.getElementById("quickViewModalCollectionAll"),
        modalCollectionCurrent: document.getElementById("quickViewModalCollectionCurrent"),
        monthlyYearFrom: document.getElementById("quickViewMonthlyYearFrom"),
        monthlyYearTo: document.getElementById("quickViewMonthlyYearTo"),
        yearlyYearFrom: document.getElementById("quickViewYearlyYearFrom"),
        yearlyYearTo: document.getElementById("quickViewYearlyYearTo"),
        monthlyTrendHint: document.getElementById("quickViewMonthlyTrendHint"),
        yearlyTrendHint: document.getElementById("quickViewYearlyTrendHint"),
        modalMonthlyChart: document.getElementById("quickViewMonthlyChart"),
        modalYearlyChart: document.getElementById("quickViewYearlyChart"),
        modalYearLabels: Array.from(document.querySelectorAll("[data-quick-view-year]")),
    };
}

function parseAmount(value) {
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : 0;
}

function createEmptyMonthlySeries() {
    return MONTH_LABELS.map((label, idx) => ({
        monthLabel: label,
        monthIndex: idx,
        generated: 0,
        collected: 0,
    }));
}

function cloneMonthlySeries(series = []) {
    const source = Array.isArray(series) ? series : [];
    return MONTH_LABELS.map((label, idx) => {
        const row = source[idx] || {};
        return {
            monthLabel: label,
            monthIndex: idx,
            generated: Number(row.generated) || 0,
            collected: Number(row.collected) || 0,
        };
    });
}

function addMonthlySeries(target = [], source = []) {
    target.forEach((row, idx) => {
        row.generated += Number(source[idx]?.generated) || 0;
        row.collected += Number(source[idx]?.collected) || 0;
    });
}

function formatMoney(value) {
    return formatCurrency(value, {
        currencySymbol: "Rs. ",
        minimumFractionDigits: 0,
        maximumFractionDigits: 0,
        emptyValue: "Rs. 0",
        invalidValue: "Rs. 0",
    });
}

function formatPercent(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return "0%";
    return `${numeric.toFixed(1)}%`;
}

function getBillMonthKey(bill = {}) {
    return normalizeMonthKey(
        bill?.monthKey ||
            bill?.month_key ||
            bill?.month ||
            bill?.monthLabel ||
            bill?.month_label ||
            ""
    );
}

function extractYear(monthKey) {
    if (!/^\d{4}-\d{2}$/.test(monthKey || "")) return null;
    return Number(monthKey.slice(0, 4));
}

function getMonthIndex(monthKey) {
    if (!/^\d{4}-\d{2}$/.test(monthKey || "")) return -1;
    return Number(monthKey.slice(5, 7)) - 1;
}

async function readGeneratedBillsFromLocalDb() {
    const entry = await getLocalEntry(LOCAL_KEYS.generatedBills);
    const raw = entry?.value;
    if (Array.isArray(raw?.bills)) return raw.bills;
    if (Array.isArray(raw)) return raw;
    return [];
}

function computeMetrics(bills = [], payments = []) {
    const currentYear = new Date().getFullYear();
    let totalGeneratedAllTime = 0;
    let totalCollectedAllTime = 0;
    let totalGeneratedCurrentYear = 0;
    let totalCollectedCurrentYear = 0;
    const monthlyByYear = {};

    const yearlyMap = new Map();

    bills.forEach((bill) => {
        const totalAmount = parseAmount(bill?.totalAmount ?? bill?.total_amount);
        const amountPaid = parseAmount(bill?.amountPaid ?? bill?.amount_paid);
        const monthKey = getBillMonthKey(bill);
        const year = extractYear(monthKey);
        const monthIndex = getMonthIndex(monthKey);

        totalGeneratedAllTime += totalAmount;
        totalCollectedAllTime += amountPaid;

        if (year === currentYear) {
            totalGeneratedCurrentYear += totalAmount;
            totalCollectedCurrentYear += amountPaid;
        }

        if (year !== null && monthIndex >= 0 && monthIndex < 12) {
            const key = String(year);
            if (!monthlyByYear[key]) {
                monthlyByYear[key] = createEmptyMonthlySeries();
            }
            monthlyByYear[key][monthIndex].generated += totalAmount;
            monthlyByYear[key][monthIndex].collected += amountPaid;
        }

        if (year !== null) {
            const existing = yearlyMap.get(year) || { year, generated: 0, collected: 0 };
            existing.generated += totalAmount;
            existing.collected += amountPaid;
            yearlyMap.set(year, existing);
        }
    });

    const yearlyAll = Array.from(yearlyMap.values()).sort((a, b) => a.year - b.year);
    const availableYears = yearlyAll.map((row) => Number(row.year)).filter((year) => Number.isFinite(year));

    if (!availableYears.length) {
        availableYears.push(currentYear);
        yearlyAll.push({ year: currentYear, generated: 0, collected: 0 });
    }

    const lastFiveYears = Array.from({ length: 5 }, (_, idx) => currentYear - 4 + idx);
    lastFiveYears.forEach((year) => {
        const key = String(year);
        if (!monthlyByYear[key]) {
            monthlyByYear[key] = createEmptyMonthlySeries();
        }
    });

    availableYears.forEach((year) => {
        const key = String(year);
        if (!monthlyByYear[key]) {
            monthlyByYear[key] = createEmptyMonthlySeries();
        }
    });

    const outstandingAllTime = Math.max(0, totalGeneratedAllTime - totalCollectedAllTime);
    const outstandingCurrentYear = Math.max(0, totalGeneratedCurrentYear - totalCollectedCurrentYear);
    const collectionRateAllTime =
        totalGeneratedAllTime > 0 ? (totalCollectedAllTime / totalGeneratedAllTime) * 100 : 0;
    const collectionRateCurrentYear =
        totalGeneratedCurrentYear > 0 ? (totalCollectedCurrentYear / totalGeneratedCurrentYear) * 100 : 0;

    return {
        currentYear,
        totalGeneratedAllTime,
        totalCollectedAllTime,
        totalGeneratedCurrentYear,
        totalCollectedCurrentYear,
        outstandingAllTime,
        outstandingCurrentYear,
        collectionRateAllTime,
        collectionRateCurrentYear,
        monthlyByYear,
        lastFiveYears,
        availableYears,
        yearlyAll,
        paymentCount: Array.isArray(payments) ? payments.length : 0,
    };
}

function getDefaultRangeForYears(years = []) {
    if (!Array.isArray(years) || !years.length) {
        const year = new Date().getFullYear();
        return { from: year, to: year };
    }
    const sorted = [...years].sort((a, b) => a - b);
    const to = sorted[sorted.length - 1];
    const from = sorted[Math.max(0, sorted.length - 5)];
    return { from, to };
}

function normalizeYearRange(years = [], fromRaw, toRaw) {
    const sorted = Array.isArray(years) ? [...years].sort((a, b) => a - b) : [];
    if (!sorted.length) {
        const nowYear = new Date().getFullYear();
        return { from: nowYear, to: nowYear };
    }

    const fallback = getDefaultRangeForYears(sorted);
    let from = Number(fromRaw);
    let to = Number(toRaw);
    if (!Number.isFinite(from) || !sorted.includes(from)) from = fallback.from;
    if (!Number.isFinite(to) || !sorted.includes(to)) to = fallback.to;
    if (from > to) [from, to] = [to, from];
    return { from, to };
}

function formatRangeLabel(prefix, range) {
    return `${prefix} (${range.from}-${range.to})`;
}

function renderYearRangeSelects(fromEl, toEl, years, range) {
    if (!fromEl || !toEl) return;
    const sortedDesc = [...years].sort((a, b) => b - a);

    const fill = (selectEl, value) => {
        selectEl.innerHTML = "";
        sortedDesc.forEach((year) => {
            const opt = document.createElement("option");
            opt.value = String(year);
            opt.textContent = String(year);
            if (year === value) opt.selected = true;
            selectEl.appendChild(opt);
        });
    };

    fill(fromEl, range.from);
    fill(toEl, range.to);
}

function aggregateMonthlyForRange(metrics, range) {
    const monthlyByYear = metrics?.monthlyByYear || {};
    const rows = createEmptyMonthlySeries();
    for (let year = range.from; year <= range.to; year += 1) {
        addMonthlySeries(rows, monthlyByYear[String(year)] || createEmptyMonthlySeries());
    }
    return rows;
}

function filterYearlyRowsForRange(metrics, range) {
    const rows = Array.isArray(metrics?.yearlyAll) ? metrics.yearlyAll : [];
    return rows.filter((row) => Number(row.year) >= range.from && Number(row.year) <= range.to);
}

function renderMonthlyChart(list = [], hint = "") {
    const { modalMonthlyChart, monthlyTrendHint } = getElements();
    if (!modalMonthlyChart) return;
    modalMonthlyChart.innerHTML = "";
    if (monthlyTrendHint) monthlyTrendHint.textContent = hint || "Generated vs collected";
    if (!list.length) {
        modalMonthlyChart.innerHTML = '<div class="text-[11px] text-slate-500">No monthly data available.</div>';
        return;
    }

    const maxValue = Math.max(
        1,
        ...list.flatMap((item) => [Number(item.generated) || 0, Number(item.collected) || 0])
    );

    list.forEach((item) => {
        const generated = Number(item.generated) || 0;
        const collected = Number(item.collected) || 0;
        const generatedPct = Math.max(0, Math.min(100, (generated / maxValue) * 100));
        const collectedPct = Math.max(0, Math.min(100, (collected / maxValue) * 100));

        const row = document.createElement("div");
        row.className = "rounded-lg border border-slate-200 bg-white px-2 py-2";
        row.innerHTML = `
            <div class="flex items-center justify-between text-[10px] mb-1">
                <span class="font-semibold text-slate-700">${item.monthLabel}</span>
                <span class="text-slate-500">G ${formatMoney(generated)} | C ${formatMoney(collected)}</span>
            </div>
            <div class="space-y-1.5">
                <div>
                    <div class="text-[9px] text-indigo-700 mb-0.5">Generated</div>
                    <div class="h-2 rounded bg-indigo-100 overflow-hidden">
                        <div class="h-full bg-indigo-500" style="width: ${generatedPct}%"></div>
                    </div>
                </div>
                <div>
                    <div class="text-[9px] text-emerald-700 mb-0.5">Collected</div>
                    <div class="h-2 rounded bg-emerald-100 overflow-hidden">
                        <div class="h-full bg-emerald-500" style="width: ${collectedPct}%"></div>
                    </div>
                </div>
            </div>
        `;
        modalMonthlyChart.appendChild(row);
    });
}

function renderYearlyChart(list = [], hint = "") {
    const { modalYearlyChart, yearlyTrendHint } = getElements();
    if (!modalYearlyChart) return;
    modalYearlyChart.innerHTML = "";
    if (yearlyTrendHint) yearlyTrendHint.textContent = hint || "Generated vs collected by year";
    if (!list.length) {
        modalYearlyChart.innerHTML = '<div class="text-[11px] text-slate-500">No yearly trend data available.</div>';
        return;
    }

    const rows = [...list].sort((a, b) => Number(a.year) - Number(b.year));
    const maxValue = Math.max(
        1,
        ...rows.flatMap((item) => [Number(item.generated) || 0, Number(item.collected) || 0])
    );

    const chartHeight = 260;
    const paddingTop = 16;
    const paddingRight = 20;
    const paddingBottom = 40;
    const paddingLeft = 52;
    const plotHeight = chartHeight - paddingTop - paddingBottom;
    const groupWidth = 38;
    const groupGap = 24;
    const barWidth = 14;
    const chartWidth =
        paddingLeft +
        paddingRight +
        rows.length * groupWidth +
        Math.max(0, rows.length - 1) * groupGap;
    const plotBottom = paddingTop + plotHeight;

    const yTicks = 4;
    const gridLines = [];
    for (let i = 0; i <= yTicks; i += 1) {
        const ratio = i / yTicks;
        const y = plotBottom - ratio * plotHeight;
        const value = Math.round(maxValue * ratio);
        gridLines.push({ y, value });
    }

    const bars = rows
        .map((item, index) => {
            const generated = Number(item.generated) || 0;
            const collected = Number(item.collected) || 0;
            const groupX = paddingLeft + index * (groupWidth + groupGap);
            const generatedHeight = Math.max(0, (generated / maxValue) * plotHeight);
            const collectedHeight = Math.max(0, (collected / maxValue) * plotHeight);
            const generatedY = plotBottom - generatedHeight;
            const collectedY = plotBottom - collectedHeight;
            const yearX = groupX + groupWidth / 2;
            const generatedX = groupX + 4;
            const collectedX = generatedX + barWidth + 6;
            return `
                <rect x="${generatedX}" y="${generatedY}" width="${barWidth}" height="${generatedHeight}" rx="2" fill="#6366f1">
                    <title>${item.year} Generated: ${formatMoney(generated)}</title>
                </rect>
                <rect x="${collectedX}" y="${collectedY}" width="${barWidth}" height="${collectedHeight}" rx="2" fill="#10b981">
                    <title>${item.year} Collected: ${formatMoney(collected)}</title>
                </rect>
                <text x="${yearX}" y="${plotBottom + 16}" text-anchor="middle" font-size="10" fill="#475569">${item.year}</text>
            `;
        })
        .join("");

    const grids = gridLines
        .map(
            (line) => `
            <line x1="${paddingLeft}" y1="${line.y}" x2="${chartWidth - paddingRight}" y2="${line.y}" stroke="#e2e8f0" stroke-width="1"></line>
            <text x="${paddingLeft - 6}" y="${line.y + 3}" text-anchor="end" font-size="9" fill="#64748b">${line.value.toLocaleString("en-IN")}</text>
        `
        )
        .join("");

    modalYearlyChart.innerHTML = `
        <div class="mb-2 flex items-center gap-3 text-[10px] text-slate-600">
            <span class="inline-flex items-center gap-1">
                <span class="inline-block h-2.5 w-2.5 rounded-sm bg-indigo-500"></span>
                Generated
            </span>
            <span class="inline-flex items-center gap-1">
                <span class="inline-block h-2.5 w-2.5 rounded-sm bg-emerald-500"></span>
                Collected
            </span>
        </div>
        <div class="overflow-x-auto rounded-md border border-slate-200 bg-white p-2">
            <svg width="${chartWidth}" height="${chartHeight}" viewBox="0 0 ${chartWidth} ${chartHeight}" role="img" aria-label="Yearly generated versus collected rent graph">
                ${grids}
                <line x1="${paddingLeft}" y1="${plotBottom}" x2="${chartWidth - paddingRight}" y2="${plotBottom}" stroke="#94a3b8" stroke-width="1"></line>
                ${bars}
            </svg>
        </div>
    `;
}

function renderMetrics() {
    const {
        widgetAllGenerated,
        widgetAllCollected,
        widgetCurrentGenerated,
        widgetCurrentCollected,
        widgetStatus,
        modalAllGenerated,
        modalAllCollected,
        modalCurrentGenerated,
        modalCurrentCollected,
        modalOutstandingAll,
        modalOutstandingCurrent,
        modalCollectionAll,
        modalCollectionCurrent,
        monthlyYearFrom,
        monthlyYearTo,
        yearlyYearFrom,
        yearlyYearTo,
        modalYearLabels,
    } = getElements();

    const metrics = quickViewState.metrics;
    const allGenerated = formatMoney(metrics.totalGeneratedAllTime);
    const allCollected = formatMoney(metrics.totalCollectedAllTime);
    const currentGenerated = formatMoney(metrics.totalGeneratedCurrentYear);
    const currentCollected = formatMoney(metrics.totalCollectedCurrentYear);

    if (widgetAllGenerated) widgetAllGenerated.textContent = allGenerated;
    if (widgetAllCollected) widgetAllCollected.textContent = allCollected;
    if (widgetCurrentGenerated) widgetCurrentGenerated.textContent = currentGenerated;
    if (widgetCurrentCollected) widgetCurrentCollected.textContent = currentCollected;

    if (modalAllGenerated) modalAllGenerated.textContent = allGenerated;
    if (modalAllCollected) modalAllCollected.textContent = allCollected;
    if (modalCurrentGenerated) modalCurrentGenerated.textContent = currentGenerated;
    if (modalCurrentCollected) modalCurrentCollected.textContent = currentCollected;
    if (modalOutstandingAll) modalOutstandingAll.textContent = formatMoney(metrics.outstandingAllTime);
    if (modalOutstandingCurrent) modalOutstandingCurrent.textContent = formatMoney(metrics.outstandingCurrentYear);
    if (modalCollectionAll) modalCollectionAll.textContent = formatPercent(metrics.collectionRateAllTime);
    if (modalCollectionCurrent) modalCollectionCurrent.textContent = formatPercent(metrics.collectionRateCurrentYear);
    if (Array.isArray(modalYearLabels)) {
        modalYearLabels.forEach((el) => {
            if (el) el.textContent = String(metrics.currentYear);
        });
    }

    if (widgetStatus) {
        if (quickViewState.loading) {
            widgetStatus.textContent = "Refreshing from local SQLite...";
        } else if (quickViewState.error) {
            widgetStatus.textContent = quickViewState.error;
        } else {
            widgetStatus.textContent = `Loaded ${metrics.paymentCount} payments from local SQLite cache.`;
        }
    }

    const availableYears = Array.isArray(metrics.availableYears) ? metrics.availableYears : [];

    const monthlyRange = normalizeYearRange(
        availableYears,
        quickViewState.monthlyRangeFrom,
        quickViewState.monthlyRangeTo
    );
    quickViewState.monthlyRangeFrom = String(monthlyRange.from);
    quickViewState.monthlyRangeTo = String(monthlyRange.to);
    renderYearRangeSelects(monthlyYearFrom, monthlyYearTo, availableYears, monthlyRange);
    const monthlyRows = aggregateMonthlyForRange(metrics, monthlyRange);
    renderMonthlyChart(monthlyRows, formatRangeLabel("Generated vs collected", monthlyRange));

    const yearlyRange = normalizeYearRange(
        availableYears,
        quickViewState.yearlyRangeFrom,
        quickViewState.yearlyRangeTo
    );
    quickViewState.yearlyRangeFrom = String(yearlyRange.from);
    quickViewState.yearlyRangeTo = String(yearlyRange.to);
    renderYearRangeSelects(yearlyYearFrom, yearlyYearTo, availableYears, yearlyRange);
    const yearlyRows = filterYearlyRowsForRange(metrics, yearlyRange);
    renderYearlyChart(yearlyRows, formatRangeLabel("Generated vs collected by year", yearlyRange));
}

async function refreshQuickViewData() {
    if (quickViewState.loading) return;
    quickViewState.loading = true;
    quickViewState.error = "";
    renderMetrics();

    try {
        const [bills, payments] = await Promise.all([
            readGeneratedBillsFromLocalDb(),
            getLocalList(LOCAL_KEYS.payments, []),
        ]);
        quickViewState.metrics = computeMetrics(Array.isArray(bills) ? bills : [], payments);
    } catch (err) {
        console.error("Failed to load quick view metrics", err);
        quickViewState.error = "Failed to read quick view metrics.";
        quickViewState.metrics = computeMetrics([], []);
    } finally {
        quickViewState.loading = false;
        renderMetrics();
    }
}

function bindModalEvents() {
    const { modal, modalCloseBtn, widgetExpandBtn } = getElements();
    if (!modal) return;

    if (!modal.dataset.bound) {
        modal.dataset.bound = "1";
        modal.addEventListener("click", (event) => {
            if (event.target === modal) hideModal(modal);
        });
    }

    if (modalCloseBtn && !modalCloseBtn.dataset.bound) {
        modalCloseBtn.dataset.bound = "1";
        modalCloseBtn.addEventListener("click", () => hideModal(modal));
    }

    if (widgetExpandBtn && !widgetExpandBtn.dataset.bound) {
        widgetExpandBtn.dataset.bound = "1";
        widgetExpandBtn.addEventListener("click", () => showModal(modal));
    }

    if (!quickViewState.escapeBound) {
        quickViewState.escapeBound = true;
        document.addEventListener("keydown", (event) => {
            if (event.key !== "Escape") return;
            const { modal: activeModal } = getElements();
            if (!activeModal || activeModal.classList.contains("hidden")) return;
            hideModal(activeModal);
        });
    }
}

function bindRefreshButtons() {
    const { widgetRefreshBtn, modalRefreshBtn } = getElements();
    if (widgetRefreshBtn && !widgetRefreshBtn.dataset.bound) {
        widgetRefreshBtn.dataset.bound = "1";
        widgetRefreshBtn.addEventListener("click", refreshQuickViewData);
    }
    if (modalRefreshBtn && !modalRefreshBtn.dataset.bound) {
        modalRefreshBtn.dataset.bound = "1";
        modalRefreshBtn.addEventListener("click", refreshQuickViewData);
    }
}

function bindRangeSelects() {
    const { monthlyYearFrom, monthlyYearTo, yearlyYearFrom, yearlyYearTo } = getElements();

    const bind = (el, handler) => {
        if (!el || el.dataset.bound) return;
        el.dataset.bound = "1";
        el.addEventListener("change", (event) => {
            const target = event.target;
            if (!(target instanceof HTMLSelectElement)) return;
            handler(target.value || "");
            renderMetrics();
        });
    };

    bind(monthlyYearFrom, (value) => {
        quickViewState.monthlyRangeFrom = value;
    });
    bind(monthlyYearTo, (value) => {
        quickViewState.monthlyRangeTo = value;
    });
    bind(yearlyYearFrom, (value) => {
        quickViewState.yearlyRangeFrom = value;
    });
    bind(yearlyYearTo, (value) => {
        quickViewState.yearlyRangeTo = value;
    });
}

function bindDataRefreshEvents() {
    document.addEventListener("sync:completed", refreshQuickViewData);
    document.addEventListener("payment:saved", refreshQuickViewData);
    document.addEventListener("paid-bills:updated", refreshQuickViewData);
}

export function initQuickView() {
    if (quickViewState.initialized) return;
    quickViewState.initialized = true;

    const { widget } = getElements();
    if (!widget) return;

    bindModalEvents();
    bindRefreshButtons();
    bindRangeSelects();
    bindDataRefreshEvents();
    refreshQuickViewData();
}
