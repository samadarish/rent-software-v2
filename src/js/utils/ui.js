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

    toast.innerHTML = `<span class="flex-1">${message}</span>`;
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
