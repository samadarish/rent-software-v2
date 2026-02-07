import { hideModal, showModal, showToast } from "../../utils/ui.js";

const VALIDATION_MODE_KEY = "tenantApp.validation.enabled";
const LOGO_CLICK_TARGET = 10;
const LOGO_CLICK_RESET_MS = 8000;

let logoClickCount = 0;
let logoClickTimer = null;
let validationModeBound = false;

function readValidationMode() {
    const raw = localStorage.getItem(VALIDATION_MODE_KEY);
    if (raw === null || raw === undefined || raw === "") return true;
    return raw !== "false";
}

function resetLogoClickCounter() {
    logoClickCount = 0;
    if (logoClickTimer) {
        clearTimeout(logoClickTimer);
        logoClickTimer = null;
    }
}

function scheduleLogoCounterReset() {
    if (logoClickTimer) clearTimeout(logoClickTimer);
    logoClickTimer = setTimeout(() => {
        logoClickCount = 0;
        logoClickTimer = null;
    }, LOGO_CLICK_RESET_MS);
}

function getModalElements() {
    return {
        modal: document.getElementById("validationMasterModal"),
        message: document.getElementById("validationMasterMessage"),
        close: document.getElementById("validationMasterClose"),
        cancel: document.getElementById("validationMasterCancel"),
        confirm: document.getElementById("validationMasterConfirm"),
    };
}

function openValidationToggleModal() {
    const { modal, message, confirm } = getModalElements();
    if (!modal || !message || !confirm) return;

    const currentlyEnabled = readValidationMode();
    const nextEnabled = !currentlyEnabled;
    modal.dataset.nextEnabled = nextEnabled ? "1" : "0";

    message.textContent = currentlyEnabled
        ? "Disable all form validations across the app?"
        : "Enable all form validations across the app?";
    confirm.textContent = currentlyEnabled ? "Disable validations" : "Enable validations";
    confirm.classList.remove("bg-rose-600", "hover:bg-rose-500", "bg-emerald-600", "hover:bg-emerald-500");
    if (nextEnabled) {
        confirm.classList.add("bg-emerald-600", "hover:bg-emerald-500");
    } else {
        confirm.classList.add("bg-rose-600", "hover:bg-rose-500");
    }

    showModal(modal);
}

function closeValidationToggleModal() {
    const { modal } = getModalElements();
    if (!modal) return;
    modal.dataset.nextEnabled = "";
    hideModal(modal);
}

export function isValidationEnabled() {
    return readValidationMode();
}

export function setValidationEnabled(enabled, options = {}) {
    const notify = options.notify !== false;
    const normalized = Boolean(enabled);
    localStorage.setItem(VALIDATION_MODE_KEY, normalized ? "true" : "false");
    if (notify) {
        showToast(
            normalized ? "Validations enabled for all forms." : "Validations disabled for all forms.",
            normalized ? "success" : "warning"
        );
    }
    if (typeof document !== "undefined") {
        document.dispatchEvent(
            new CustomEvent("app:validation-mode-changed", {
                detail: { enabled: normalized },
            })
        );
    }
}

export function initValidationMasterToggle() {
    if (validationModeBound) return;
    validationModeBound = true;

    const logo = document.getElementById("appLogoValidationToggle");
    const { modal, close, cancel, confirm } = getModalElements();

    if (logo) {
        logo.addEventListener("click", () => {
            logoClickCount += 1;
            scheduleLogoCounterReset();
            if (logoClickCount < LOGO_CLICK_TARGET) return;
            resetLogoClickCounter();
            openValidationToggleModal();
        });
    }

    if (modal) {
        modal.addEventListener("click", (event) => {
            if (event.target === modal) closeValidationToggleModal();
        });
    }

    if (close) close.addEventListener("click", closeValidationToggleModal);
    if (cancel) cancel.addEventListener("click", closeValidationToggleModal);
    if (confirm) {
        confirm.addEventListener("click", () => {
            const shouldEnable = (modal?.dataset?.nextEnabled || "") === "1";
            setValidationEnabled(shouldEnable);
            closeValidationToggleModal();
        });
    }
}

