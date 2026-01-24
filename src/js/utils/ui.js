/**
 * UI Utility Functions
 * 
 * Functions for UI interactions like toast notifications.
 */

/**
 * Displays a toast notification to the user
 * @param {string} message - The message to display
 * @param {"success" | "error" | "info" | "warning"} type - The type of toast (determines color)
 */
export function showToast(message, type = "success") {
    saveToastHistoryEntry(message, type);
    const container = document.getElementById("toastContainer");
    if (!container) {
        console.log(message);
        return;
    }

    const toast = document.createElement("div");
    toast.className =
        "pointer-events-auto w-full flex items-start gap-3 rounded-xl px-4 py-3 font-medium leading-snug text-white shadow-lg opacity-0 translate-y-1 transition";

    if (type === "success") {
        toast.classList.add("bg-emerald-600");
    } else if (type === "error") {
        toast.classList.add("bg-red-600");
    } else if (type === "warning") {
        toast.classList.add("bg-amber-500");
    } else {
        toast.classList.add("bg-slate-700");
    }

    toast.innerHTML = `<span class="flex-1 break-words break-all">${message}</span>`;
    container.appendChild(toast);

    // Animate in
    requestAnimationFrame(() => {
        toast.classList.remove("opacity-0", "translate-y-1");
    });

    // Animate out and remove after 2.5 seconds
    setTimeout(() => {
        toast.classList.add("opacity-0", "translate-y-1");
        setTimeout(() => toast.remove(), 300);
    }, 2500);
}

const TOAST_HISTORY_KEY = "toastHistory.v1";
const TOAST_HISTORY_LIMIT = 200;
let toastHistoryBound = false;

function loadToastHistory() {
    if (!window.localStorage) return [];
    try {
        const raw = localStorage.getItem(TOAST_HISTORY_KEY);
        if (!raw) return [];
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
        return [];
    }
}

function persistToastHistory(list) {
    if (!window.localStorage) return;
    try {
        localStorage.setItem(TOAST_HISTORY_KEY, JSON.stringify(list));
    } catch (err) {
        // Ignore storage failures.
    }
}

function getToastTypeStyles(type) {
    if (type === "error") return { dot: "bg-rose-500", text: "text-rose-600" };
    if (type === "warning") return { dot: "bg-amber-500", text: "text-amber-600" };
    if (type === "success") return { dot: "bg-emerald-500", text: "text-emerald-600" };
    return { dot: "bg-slate-500", text: "text-slate-600" };
}

function formatToastTimestamp(value) {
    if (!value) return "";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    return date.toLocaleString();
}

function renderToastHistory(list) {
    const container = document.getElementById("toastHistoryList");
    const empty = document.getElementById("toastHistoryEmpty");
    if (!container || !empty) return;

    container.innerHTML = "";
    if (!list.length) {
        empty.classList.remove("hidden");
        return;
    }
    empty.classList.add("hidden");

    list.forEach((entry) => {
        const row = document.createElement("div");
        row.className = "px-3 py-2 border-b border-slate-100 last:border-0";
        const styles = getToastTypeStyles(entry?.type);
        const message = (entry?.message || "").toString();
        const time = formatToastTimestamp(entry?.timestamp);
        row.innerHTML = `
            <div class="flex items-start gap-2">
                <span class="mt-1.5 h-2 w-2 rounded-full ${styles.dot}"></span>
                <div class="flex-1">
                    <div class="text-[11px] font-semibold ${styles.text}">${entry?.type || "info"}</div>
                    <div class="text-[12px] text-slate-800 break-words break-all whitespace-normal">${message}</div>
                    <div class="text-[10px] text-slate-500 mt-0.5">${time}</div>
                </div>
            </div>
        `;
        container.appendChild(row);
    });
}

function updateToastBadge(list) {
    const badge = document.getElementById("toastHistoryBadge");
    if (!badge) return;
    const errorCount = list.filter((entry) => entry?.type === "error").length;
    if (errorCount <= 0) {
        badge.textContent = "";
        badge.classList.add("hidden");
        return;
    }
    badge.textContent = String(errorCount);
    badge.classList.remove("hidden");
}

function updateToastHistoryUi(list) {
    updateToastBadge(list);
    renderToastHistory(list);
}

function saveToastHistoryEntry(message, type) {
    const entry = {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        message: (message || "").toString(),
        type,
        timestamp: new Date().toISOString(),
    };
    const existing = loadToastHistory();
    const updated = [entry, ...existing].slice(0, TOAST_HISTORY_LIMIT);
    persistToastHistory(updated);
    updateToastHistoryUi(updated);
    return entry;
}

function toggleToastHistoryDropdown(show) {
    const dropdown = document.getElementById("toastHistoryDropdown");
    if (!dropdown) return;
    dropdown.classList.toggle("hidden", !show);
}

export function initToastHistoryUi() {
    if (toastHistoryBound) return;
    toastHistoryBound = true;

    const btn = document.getElementById("toastHistoryBtn");
    const dropdown = document.getElementById("toastHistoryDropdown");
    const clearBtn = document.getElementById("toastHistoryClearBtn");
    if (!btn || !dropdown) return;

    const refresh = () => {
        const list = loadToastHistory();
        updateToastHistoryUi(list);
    };

    refresh();

    btn.addEventListener("click", (event) => {
        event.stopPropagation();
        const isHidden = dropdown.classList.contains("hidden");
        toggleToastHistoryDropdown(isHidden);
        if (isHidden) refresh();
    });

    if (clearBtn) {
        clearBtn.addEventListener("click", (event) => {
            event.stopPropagation();
            persistToastHistory([]);
            updateToastHistoryUi([]);
        });
    }

    document.addEventListener("click", (event) => {
        if (dropdown.classList.contains("hidden")) return;
        if (dropdown.contains(event.target) || btn.contains(event.target)) return;
        toggleToastHistoryDropdown(false);
    });

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
            toggleToastHistoryDropdown(false);
        }
    });
}

/**
 * Updates the online/offline badge in the header with the latest state and helper text.
 * @param {"checking" | "online" | "offline"} status - Connectivity state.
 * @param {string} message - Optional label to display instead of default status text.
 */
export function updateConnectionIndicator(status = "checking", message = "") {
    const indicator = document.getElementById("connectionIndicator");
    if (!indicator) return;

    const dot = indicator.querySelector(".status-dot");
    const label = indicator.querySelector(".status-label");

    const variants = {
        online: {
            wrap: "bg-emerald-50 text-emerald-800 border-emerald-200",
            dot: "bg-emerald-500",
            label: "Online",
        },
        offline: {
            wrap: "bg-rose-50 text-rose-800 border-rose-200",
            dot: "bg-rose-500",
            label: "Offline",
        },
        checking: {
            wrap: "bg-slate-700 text-white border-slate-600",
            dot: "bg-slate-400",
            label: "Checking...",
        },
    };

    const onlineState = status === "offline" ? false : navigator.onLine;
    const key = status === "checking" ? "checking" : onlineState ? "online" : "offline";
    const variant = variants[key];

    indicator.className = `flex items-center gap-2 text-[11px] px-3 py-1 rounded-full border ${variant.wrap}`;

    if (dot) dot.className = `status-dot w-2 h-2 rounded-full ${variant.dot}`;
    if (label) label.textContent = message || variant.label;
}

/**
 * Updates the sync status badge in the header.
 * @param {"synced" | "pending" | "syncing" | "error"} status - Sync state.
 * @param {string} message - Optional label override.
 */
export function updateSyncIndicator(status = "synced", message = "") {
    const indicator = document.getElementById("syncIndicator");
    if (!indicator) return;

    const dot = indicator.querySelector(".sync-dot");
    const label = indicator.querySelector(".sync-label");

    const tickSvg = `
        <svg viewBox="0 0 16 16" class="w-3 h-3" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M3.5 8.5l3 3 6-6" stroke-linecap="round" stroke-linejoin="round"></path>
        </svg>
    `;

    const variants = {
        synced: {
            wrap: "bg-emerald-50 text-emerald-800 border-emerald-200",
            dot: "inline-flex items-center justify-center w-4 h-4 rounded-full bg-emerald-100 text-emerald-600",
            dotHtml: tickSvg,
            label: "Synced",
        },
        pending: {
            wrap: "bg-amber-50 text-amber-800 border-amber-200",
            dot: "w-2 h-2 rounded-full bg-amber-500",
            dotHtml: "",
            label: "Not synced",
        },
        syncing: {
            wrap: "bg-amber-50 text-amber-800 border-amber-200",
            dot: "inline-block h-3 w-3 rounded-full border-2 border-amber-500 border-t-transparent animate-spin",
            dotHtml: "",
            label: "Syncing",
        },
        error: {
            wrap: "bg-rose-50 text-rose-800 border-rose-200",
            dot: "w-2 h-2 rounded-full bg-rose-500",
            dotHtml: "",
            label: "Sync error",
        },
    };

    const variant = variants[status] || variants.synced;
    indicator.className = `flex items-center gap-2 text-[11px] px-3 py-1 rounded-full border ${variant.wrap}`;
    if (dot) {
        dot.className = `sync-dot ${variant.dot}`;
        dot.innerHTML = variant.dotHtml || "";
    }
    if (label) label.textContent = message || variant.label;
}

/**
 * Clones option elements from one select to another while preserving selection.
 * @param {string} sourceId
 * @param {string} targetId
 * @param {{ preserveSelection?: boolean }} options
 */
export function cloneSelectOptions(sourceId, targetId, options = {}) {
    const { preserveSelection = true } = options;
    const source = document.getElementById(sourceId);
    const target = document.getElementById(targetId);
    if (!source || !target) return;
    if (!source.options || !target.options) return;

    const previous = preserveSelection ? target.value : "";
    target.innerHTML = "";
    Array.from(source.options).forEach((opt) => {
        const clone = opt.cloneNode(true);
        target.appendChild(clone);
    });

    if (preserveSelection && previous && Array.from(target.options).some((o) => o.value === previous)) {
        target.value = previous;
    }
}

const UNIT_DROPDOWN_CLASS = "unit-select-dropdown";
let unitDropdownDismissBound = false;

function closeUnitDropdowns(except) {
    document.querySelectorAll(`.${UNIT_DROPDOWN_CLASS}`).forEach((wrapper) => {
        if (except && wrapper === except) return;
        const menu = wrapper.querySelector("[data-unit-menu]");
        if (!menu) return;
        menu.classList.add("hidden");
        wrapper.dataset.open = "0";
    });
}

function bindUnitDropdownDismiss() {
    if (unitDropdownDismissBound) return;
    unitDropdownDismissBound = true;
    document.addEventListener("click", (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        const wrapper = target.closest(`.${UNIT_DROPDOWN_CLASS}`);
        closeUnitDropdowns(wrapper || null);
    });
    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape") closeUnitDropdowns(null);
    });
}

function getSelectTextSizeClass(select) {
    const classes = (select.className || "").split(/\s+/);
    return (
        classes.find((c) => c.startsWith("text-[") || c === "text-xs" || c === "text-sm" || c === "text-base") ||
        "text-xs"
    );
}

function buildUnitTag(label, className) {
    const tag = document.createElement("span");
    tag.textContent = label;
    tag.className = `inline-flex items-center px-2 py-0.5 rounded-full border font-semibold ${className}`;
    return tag;
}

/**
 * Builds or refreshes a custom unit dropdown UI for a select element.
 * @param {string|HTMLSelectElement} selectOrId
 * @param {{ includeCurrentTag?: boolean }} options
 */
export function syncUnitSelectDropdown(selectOrId, options = {}) {
    const { includeCurrentTag = false, currentValue = "" } = options;
    const select = typeof selectOrId === "string" ? document.getElementById(selectOrId) : selectOrId;
    if (!select) return;

    bindUnitDropdownDismiss();

    let wrapper = select.parentElement?.querySelector(`.${UNIT_DROPDOWN_CLASS}[data-for="${select.id}"]`);
    if (!wrapper) {
        wrapper = document.createElement("div");
        wrapper.className = `${UNIT_DROPDOWN_CLASS} relative w-full`;
        wrapper.dataset.for = select.id || "";
        wrapper.dataset.open = "0";

        const button = document.createElement("button");
        button.type = "button";
        button.dataset.unitToggle = "1";
        button.className = `${select.className} flex items-center justify-between gap-2 text-left cursor-pointer`;

        const label = document.createElement("span");
        label.dataset.unitLabel = "1";
        label.className = "truncate";

        const caret = document.createElement("span");
        caret.textContent = "v";
        caret.className = "text-slate-500";

        button.append(label, caret);

        const menu = document.createElement("div");
        menu.dataset.unitMenu = "1";
        menu.className =
            "absolute left-0 right-0 mt-1 max-h-64 overflow-auto bg-white border border-slate-200 rounded shadow-lg z-30 hidden";

        wrapper.append(button, menu);
        select.insertAdjacentElement("afterend", wrapper);
        select.classList.add("hidden");

        button.addEventListener("click", (event) => {
            event.preventDefault();
            if (button.disabled) return;
            const isOpen = wrapper.dataset.open === "1";
            closeUnitDropdowns(wrapper);
            menu.classList.toggle("hidden", isOpen);
            wrapper.dataset.open = isOpen ? "0" : "1";
        });
    }

    wrapper.dataset.includeCurrent = includeCurrentTag ? "1" : "0";
    wrapper.dataset.currentValue = currentValue || "";

    if (!select.dataset.unitDropdownBound) {
        select.dataset.unitDropdownBound = "1";
        select.addEventListener("change", () => {
            syncUnitSelectDropdown(select, {
                includeCurrentTag: wrapper.dataset.includeCurrent === "1",
                currentValue: wrapper.dataset.currentValue || "",
            });
        });
    }

    const button = wrapper.querySelector("[data-unit-toggle]");
    const label = wrapper.querySelector("[data-unit-label]");
    const menu = wrapper.querySelector("[data-unit-menu]");
    if (!button || !label || !menu) return;

    const selectedOption = select.selectedOptions?.[0] || select.options[0];
    label.textContent = selectedOption?.textContent || "Select unit";

    button.disabled = select.disabled;
    button.classList.toggle("bg-slate-50", select.disabled);
    button.classList.toggle("text-slate-500", select.disabled);
    button.classList.toggle("cursor-not-allowed", select.disabled);
    button.classList.toggle("cursor-pointer", !select.disabled);

    if (select.disabled) {
        menu.classList.add("hidden");
        wrapper.dataset.open = "0";
    }

    const sizeClass = getSelectTextSizeClass(select);
    menu.innerHTML = "";

    Array.from(select.options).forEach((opt) => {
        const item = document.createElement("button");
        item.type = "button";
        item.className = `w-full text-left flex items-center justify-between gap-2 px-2 py-1.5 ${sizeClass} ${
            opt.value === select.value ? "bg-slate-50" : ""
        }`;
        item.disabled = opt.disabled;
        if (opt.value === "") {
            item.classList.add("text-slate-400");
        } else {
            item.classList.add("hover:bg-slate-100");
        }

        const text = document.createElement("span");
        text.textContent = opt.textContent || "";
        text.className = "truncate";

        const tags = document.createElement("span");
        tags.className = "flex items-center gap-1";

        const occupied = opt.dataset?.occupied;
        if (occupied === "1" || occupied === "0") {
            tags.appendChild(
                buildUnitTag(
                    occupied === "1" ? "Occupied" : "Vacant",
                    occupied === "1"
                        ? "bg-emerald-100 text-emerald-700 border-emerald-200"
                        : "bg-rose-100 text-rose-700 border-rose-200"
                )
            );
        }

        if (includeCurrentTag && opt.value && opt.value === currentValue) {
            tags.appendChild(
                buildUnitTag("Current", "bg-slate-100 text-slate-600 border-slate-200")
            );
        }

        item.append(text, tags);

        item.addEventListener("click", (event) => {
            event.preventDefault();
            if (item.disabled) return;
            select.value = opt.value;
            select.dispatchEvent(new Event("change", { bubbles: true }));
            menu.classList.add("hidden");
            wrapper.dataset.open = "0";
        });

        menu.appendChild(item);
    });
}

const DEFAULT_ANIM_MS = 200;

/**
 * Applies a smooth fade/slide toggle for any element by handling hidden class timing
 * @param {HTMLElement | null} element
 * @param {boolean} show
 * @param {{ baseClass?: string, activeClass?: string, hidingClass?: string, hiddenClass?: string, duration?: number }} options
 */
/**
 * Smoothly toggles visibility for a given element using utility classes.
 * @param {HTMLElement | null} element - Element being shown/hidden.
 * @param {boolean} show - Whether the element should be visible.
 * @param {{ baseClass?: string, activeClass?: string, hidingClass?: string, hiddenClass?: string, duration?: number }} options
 */
export function smoothToggle(element, show, options = {}) {
    if (!element) return;

    const {
        baseClass = "fade-section",
        activeClass = "is-visible",
        hidingClass = "is-hiding",
        hiddenClass = "hidden",
        duration = DEFAULT_ANIM_MS,
    } = options;

    if (element.__hideTimer) {
        clearTimeout(element.__hideTimer);
        element.__hideTimer = null;
    }

    element.classList.add(baseClass);

    if (!show && duration <= 0) {
        element.classList.remove(activeClass, hidingClass);
        element.classList.add(hiddenClass);
        return;
    }

    if (show) {
        element.classList.remove(hiddenClass, hidingClass);
        requestAnimationFrame(() => element.classList.add(activeClass));
    } else {
        element.classList.remove(activeClass);
        element.classList.add(hidingClass);
        element.__hideTimer = setTimeout(() => {
            element.classList.add(hiddenClass);
        }, duration);
    }
}

const modalAnimationOptions = {
    baseClass: "fade-overlay",
    duration: 220,
};

/**
 * Displays a modal overlay with a fade animation.
 * @param {HTMLElement | null} modal
 */
export function showModal(modal) {
    smoothToggle(modal, true, modalAnimationOptions);
}

/**
 * Hides a modal overlay with a fade animation.
 * @param {HTMLElement | null} modal
 */
export function hideModal(modal) {
    smoothToggle(modal, false, modalAnimationOptions);
}
