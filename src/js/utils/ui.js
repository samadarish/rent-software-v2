/**
 * UI Utility Functions
 * 
 * Functions for UI interactions like toast notifications.
 */

/**
 * Displays a toast notification to the user
 * @param {string} message - The message to display
 * @param {"success" | "error" | "info"} type - The type of toast (determines color)
 */
export function showToast(message, type = "success") {
    const container = document.getElementById("toastContainer");
    if (!container) {
        console.log(message);
        return;
    }

    const toast = document.createElement("div");
    toast.className =
        "pointer-events-auto flex items-center gap-2 px-3 py-1.5 rounded-lg shadow text-[11px] text-white opacity-0 translate-y-1 transition";

    if (type === "success") {
        toast.classList.add("bg-emerald-600");
    } else if (type === "error") {
        toast.classList.add("bg-red-600");
    } else {
        toast.classList.add("bg-slate-700");
    }

    toast.innerHTML = `<span>${message}</span>`;
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

const DEFAULT_ANIM_MS = 200;

/**
 * Applies a smooth fade/slide toggle for any element by handling hidden class timing
 * @param {HTMLElement | null} element
 * @param {boolean} show
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

export function showModal(modal) {
    smoothToggle(modal, true, modalAnimationOptions);
}

export function hideModal(modal) {
    smoothToggle(modal, false, modalAnimationOptions);
}
