const STORAGE_KEY = "dashboard.layout.v1";
const SLOT_COUNT = 9;
const NONE_VALUE = "__none__";

function loadLayout() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return [];
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
        return [];
    }
}

function saveLayout(layout) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(layout));
    } catch (err) {
        // Ignore storage failures.
    }
}

function normalizeLayout(layout, moduleKeySet, slotCount) {
    const normalized = Array(slotCount).fill("");
    const used = new Set();
    if (!Array.isArray(layout)) return normalized;

    layout.slice(0, slotCount).forEach((key, index) => {
        const value = typeof key === "string" ? key.trim() : "";
        if (!value || !moduleKeySet.has(value) || used.has(value)) return;
        normalized[index] = value;
        used.add(value);
    });

    return normalized;
}

function getSlotIndex(slot, fallback) {
    const raw = slot?.dataset?.dashboardSlot;
    const parsed = Number(raw);
    if (Number.isFinite(parsed)) return parsed;
    return fallback;
}

function ensureEditButton(slot, moduleEl, state) {
    if (!slot || !moduleEl) return null;

    moduleEl.classList.remove("pr-10");
    const existingInModule = moduleEl.querySelector("[data-dashboard-edit]");
    if (existingInModule) existingInModule.remove();

    let button = slot.querySelector("[data-dashboard-edit]");
    if (!button) {
        button = document.createElement("button");
        button.type = "button";
        button.dataset.dashboardEdit = "1";
        button.className =
            "absolute -top-2 right-2 z-20 inline-flex h-7 w-7 items-center justify-center rounded-full border border-slate-200 bg-white text-slate-500 shadow-sm hover:bg-white hover:text-slate-700";
        button.setAttribute("aria-label", "Edit widget");
        button.title = "Edit widget";
        button.innerHTML = `
            <svg viewBox="0 0 24 24" class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M4 20h4l10-10-4-4L4 16v4z" stroke-linecap="round" stroke-linejoin="round"></path>
                <path d="M14 6l4 4" stroke-linecap="round" stroke-linejoin="round"></path>
            </svg>
        `;
        slot.appendChild(button);
    }

    if (!button.dataset.bound) {
        button.dataset.bound = "1";
        button.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            const slotEl = button.closest("[data-dashboard-slot]");
            if (!slotEl) return;
            const index = getSlotIndex(slotEl, state.slots.indexOf(slotEl));
            if (index < 0 || index >= state.layout.length) return;
            const isOpen = slotEl.dataset.pickerOpen === "1";
            state.slots.forEach((other, idx) => {
                const otherIndex = getSlotIndex(other, idx);
                if (!state.layout[otherIndex]) return;
                other.dataset.pickerOpen = other === slotEl && !isOpen ? "1" : "0";
            });
            renderSlots(state);
            refreshOptions(state);
            if (slotEl.dataset.pickerOpen === "1") {
                slotEl.querySelector("[data-dashboard-select]")?.focus();
            }
        });
    }

    return button;
}

function renderSlots(state) {
    const assigned = new Set();

    state.slots.forEach((slot, idx) => {
        const index = getSlotIndex(slot, idx);
        const value = state.layout[index] || "";
        const picker = slot.querySelector("[data-dashboard-picker]");
        const content = slot.querySelector("[data-dashboard-content]");
        if (!picker || !content) return;

        if (value && state.modules.has(value)) {
            const moduleEl = state.modules.get(value).el;
            assigned.add(value);

            if (content.firstChild !== moduleEl) {
                while (content.firstChild) {
                    state.pool.appendChild(content.firstChild);
                }
                content.appendChild(moduleEl);
            }

            ensureEditButton(slot, moduleEl, state);
            slot.dataset.dashboardEmpty = "0";

            const isOpen = slot.dataset.pickerOpen === "1";
            picker.classList.toggle("hidden", !isOpen);
        } else {
            slot.dataset.dashboardEmpty = "1";
            slot.dataset.pickerOpen = "1";
            picker.classList.remove("hidden");

            const editButton = slot.querySelector("[data-dashboard-edit]");
            if (editButton) editButton.remove();
            while (content.firstChild) {
                state.pool.appendChild(content.firstChild);
            }
        }
    });

    state.modules.forEach((module) => {
        if (!assigned.has(module.key) && module.el.parentElement !== state.pool) {
            state.pool.appendChild(module.el);
        }
    });
}

function refreshOptions(state) {
    const used = new Set(state.layout.filter(Boolean));

    state.slots.forEach((slot, idx) => {
        const index = getSlotIndex(slot, idx);
        const select = slot.querySelector("[data-dashboard-select]");
        if (!select) return;
        const current = state.layout[index] || "";
        const available = state.moduleOrder.filter(
            (key) => !used.has(key) || key === current
        );

        select.innerHTML = "";

        const placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.disabled = true;
        placeholder.selected = !current;
        placeholder.textContent = "Select module";
        select.appendChild(placeholder);

        const noneOption = document.createElement("option");
        noneOption.value = NONE_VALUE;
        noneOption.textContent = "None";
        select.appendChild(noneOption);

        available.forEach((key) => {
            const opt = document.createElement("option");
            opt.value = key;
            opt.textContent = state.modules.get(key)?.label || key;
            if (key === current) opt.selected = true;
            select.appendChild(opt);
        });

        if (current) {
            select.value = current;
        } else {
            select.value = "";
        }
    });
}

export function initDashboardLayout() {
    const grid = document.getElementById("dashboardGrid");
    const pool = document.getElementById("dashboardModulePool");
    if (!grid || !pool) return;

    const slots = Array.from(grid.querySelectorAll("[data-dashboard-slot]"));
    if (!slots.length) return;

    const moduleElements = Array.from(document.querySelectorAll("[data-dashboard-module]"));
    if (!moduleElements.length) return;

    const modules = new Map();
    const moduleOrder = [];
    moduleElements.forEach((el) => {
        const key = (el.dataset.dashboardModule || "").trim();
        if (!key || modules.has(key)) return;
        const label = (el.dataset.dashboardLabel || "").trim() || key;
        modules.set(key, { key, label, el });
        moduleOrder.push(key);
    });

    const moduleKeySet = new Set(moduleOrder);
    const slotCount = Math.max(SLOT_COUNT, slots.length);
    const layout = normalizeLayout(loadLayout(), moduleKeySet, slotCount);

    const state = {
        slots,
        pool,
        modules,
        moduleOrder,
        layout,
    };

    slots.forEach((slot, idx) => {
        const select = slot.querySelector("[data-dashboard-select]");
        if (!select || select.dataset.bound) return;
        select.dataset.bound = "1";
        select.addEventListener("change", () => {
            const raw = select.value;
            if (!raw) return;
            const index = getSlotIndex(slot, idx);
            if (index < 0 || index >= state.layout.length) return;
            const nextValue = raw === NONE_VALUE ? "" : raw;
            if (nextValue && !state.modules.has(nextValue)) {
                select.value = state.layout[index] || "";
                return;
            }
            state.layout[index] = nextValue;
            saveLayout(state.layout);
            slot.dataset.pickerOpen = nextValue ? "0" : "1";
            renderSlots(state);
            refreshOptions(state);
        });
    });

    renderSlots(state);
    refreshOptions(state);
}
