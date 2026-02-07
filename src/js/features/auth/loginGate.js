import { STORAGE_KEYS } from "../../constants.js";
import { callAppScript } from "../../api/appscriptClient.js";

const APPS_SCRIPT_PATH_REGEX = /^\/macros\/s\/[A-Za-z0-9_-]+\/(exec|dev)(\/)?$/;
const CURTAIN_MS = 700;
const SESSION_REUSE_MS = 2 * 60 * 60 * 1000;
const IDLE_LOCK_MS = 30 * 60 * 1000;
const APP_SCRIPT_URL_SESSION_KEY = `${STORAGE_KEYS.APP_SCRIPT_URL}.session`;
const AUTH_STORAGE_KEYS = {
    LAST_UNLOCK_AT: "tenantApp.auth.lastUnlockAt",
};
const ACTIVITY_EVENTS = ["pointerdown", "keydown", "mousemove", "touchstart", "wheel", "scroll"];

const loginGateState = {
    resolved: false,
    resolving: false,
    isLocked: true,
    appScriptUrl: "",
    promise: null,
    resolvePromise: null,
    idleTimerId: 0,
    activityBound: false,
    activityHandler: null,
    visibilityHandler: null,
    stepReason: "startup",
};

function getElements() {
    return {
        overlay: document.getElementById("loginGateOverlay"),
        urlStep: document.getElementById("loginGateUrlStep"),
        keyStep: document.getElementById("loginGateKeyStep"),
        urlInput: document.getElementById("loginGateUrlInput"),
        keyInput: document.getElementById("loginGateKeyInput"),
        error: document.getElementById("loginGateError"),
        urlContinueBtn: document.getElementById("loginGateUrlContinueBtn"),
        keyBackBtn: document.getElementById("loginGateBackBtn"),
        unlockBtn: document.getElementById("loginGateUnlockBtn"),
        title: document.getElementById("loginGateTitle"),
        subtitle: document.getElementById("loginGateSubtitle"),
    };
}

function parseAppScriptUrl(rawUrl) {
    try {
        const parsed = new URL((rawUrl || "").trim());
        if (parsed.protocol !== "https:") return null;
        if (parsed.hostname !== "script.google.com") return null;
        if (!APPS_SCRIPT_PATH_REGEX.test(parsed.pathname)) return null;
        return parsed;
    } catch (error) {
        return null;
    }
}

function normalizeAppScriptUrl(rawUrl) {
    const parsed = parseAppScriptUrl(rawUrl);
    if (!parsed) return "";
    return `${parsed.origin}${parsed.pathname}`;
}

function getLastUnlockAt() {
    const raw = sessionStorage.getItem(AUTH_STORAGE_KEYS.LAST_UNLOCK_AT);
    const parsed = Number(raw || 0);
    return Number.isFinite(parsed) ? parsed : 0;
}

function setLastUnlockAt(timestampMs) {
    sessionStorage.setItem(AUTH_STORAGE_KEYS.LAST_UNLOCK_AT, String(Number(timestampMs) || Date.now()));
}

function clearLastUnlockAt() {
    sessionStorage.removeItem(AUTH_STORAGE_KEYS.LAST_UNLOCK_AT);
    // Remove old persisted value from previous builds.
    localStorage.removeItem(AUTH_STORAGE_KEYS.LAST_UNLOCK_AT);
}

function hasReusableSession() {
    const lastUnlockAt = getLastUnlockAt();
    if (!lastUnlockAt) return false;
    return Date.now() - lastUnlockAt <= SESSION_REUSE_MS;
}

function setError(message = "") {
    const { error } = getElements();
    if (!error) return;
    if (!message) {
        error.textContent = "";
        error.classList.add("hidden");
        return;
    }
    error.textContent = message;
    error.classList.remove("hidden");
}

function setBusy(isBusy) {
    const { urlContinueBtn, keyBackBtn, unlockBtn, urlInput, keyInput } = getElements();
    [urlContinueBtn, keyBackBtn, unlockBtn, urlInput, keyInput].forEach((el) => {
        if (!el) return;
        el.disabled = !!isBusy;
    });
    if (unlockBtn) {
        unlockBtn.textContent = isBusy ? "Checking..." : "Unlock app";
    }
}

function getStepSubtitle(step, reason) {
    if (step === "url") {
        return "Paste your Apps Script Web App URL.";
    }
    if (reason === "manual") {
        return "App is locked. Enter password to continue.";
    }
    if (reason === "idle") {
        return "App auto-locked after inactivity. Enter password to continue.";
    }
    return "Password is verified against the Key sheet.";
}

function setStep(step, reason = loginGateState.stepReason) {
    const { urlStep, keyStep, title, subtitle, keyInput } = getElements();
    loginGateState.stepReason = reason || "startup";
    const isUrlStep = step === "url";
    if (urlStep) urlStep.classList.toggle("hidden", !isUrlStep);
    if (keyStep) keyStep.classList.toggle("hidden", isUrlStep);
    if (title) title.textContent = isUrlStep ? "Connect to App Script" : "Unlock app";
    if (subtitle) {
        subtitle.textContent = getStepSubtitle(step, loginGateState.stepReason);
    }
    if (!isUrlStep && keyInput) {
        keyInput.value = "";
        keyInput.focus();
    }
    setError("");
}

function clearIdleTimer() {
    if (!loginGateState.idleTimerId) return;
    window.clearTimeout(loginGateState.idleTimerId);
    loginGateState.idleTimerId = 0;
}

function startIdleTimer() {
    clearIdleTimer();
    if (loginGateState.isLocked) return;
    loginGateState.idleTimerId = window.setTimeout(() => {
        lockAppNow({ reason: "idle" });
    }, IDLE_LOCK_MS);
}

function bindActivityTracking() {
    if (loginGateState.activityBound) return;
    loginGateState.activityBound = true;

    loginGateState.activityHandler = () => {
        if (loginGateState.isLocked) return;
        startIdleTimer();
    };
    ACTIVITY_EVENTS.forEach((eventName) => {
        window.addEventListener(eventName, loginGateState.activityHandler, { passive: true });
    });

    loginGateState.visibilityHandler = () => {
        if (document.visibilityState !== "visible" || loginGateState.isLocked) return;
        startIdleTimer();
    };
    document.addEventListener("visibilitychange", loginGateState.visibilityHandler);
}

function markUnlocked() {
    loginGateState.isLocked = false;
    setLastUnlockAt(Date.now());
    bindActivityTracking();
    startIdleTimer();
}

function resolveStartupIfPending() {
    loginGateState.resolved = true;
    if (typeof loginGateState.resolvePromise === "function") {
        loginGateState.resolvePromise({ ok: true });
    }
    loginGateState.resolvePromise = null;
    loginGateState.promise = null;
}

function showGate(step, reason = "startup") {
    const { overlay, urlInput } = getElements();
    if (!overlay) return;
    overlay.classList.remove("hidden", "is-opening");
    setBusy(false);
    setStep(step, reason);
    if (step === "url" && urlInput) {
        urlInput.value = loginGateState.appScriptUrl || "";
        urlInput.focus();
    }
}

function completeUnlock() {
    if (loginGateState.resolving) return;
    loginGateState.resolving = true;
    markUnlocked();
    const { overlay } = getElements();
    if (!overlay) {
        loginGateState.resolving = false;
        resolveStartupIfPending();
        return;
    }

    overlay.classList.add("is-opening");
    window.setTimeout(() => {
        overlay.classList.add("hidden");
        overlay.classList.remove("is-opening");
        loginGateState.resolving = false;
        resolveStartupIfPending();
    }, CURTAIN_MS);
}

async function handleUrlContinue() {
    const { urlInput } = getElements();
    const raw = (urlInput?.value || "").trim();
    const normalized = normalizeAppScriptUrl(raw);
    if (!normalized) {
        setError("Enter a valid Apps Script URL (exec/dev).");
        return;
    }
    loginGateState.appScriptUrl = normalized;
    setStep("key", loginGateState.stepReason || "startup");
}

async function handleUnlock() {
    const { keyInput } = getElements();
    const enteredKey = (keyInput?.value || "").trim();
    if (!loginGateState.appScriptUrl) {
        setStep("url");
        setError("Apps Script URL is required.");
        return;
    }
    if (!enteredKey) {
        setError("Enter password.");
        return;
    }

    setError("");
    setBusy(true);
    try {
        const result = await callAppScript({
            url: loginGateState.appScriptUrl,
            action: "validateKey",
            method: "POST",
            payload: { key: enteredKey },
            cache: { useLocal: false, write: false, revalidate: false },
        });

        const isValid = result?.valid === true;
        if (!isValid) {
            setError("Wrong password.");
            return;
        }

        localStorage.setItem(STORAGE_KEYS.APP_SCRIPT_URL, loginGateState.appScriptUrl);
        sessionStorage.setItem(APP_SCRIPT_URL_SESSION_KEY, loginGateState.appScriptUrl);
        document.dispatchEvent(
            new CustomEvent("appscript:url-updated", { detail: { url: loginGateState.appScriptUrl } })
        );
        completeUnlock();
    } catch (err) {
        console.error("Failed to validate app key", err);
        setError("Unable to verify key. Check URL/internet and try again.");
    } finally {
        setBusy(false);
    }
}

function bindEvents() {
    const { urlContinueBtn, keyBackBtn, unlockBtn, urlInput, keyInput } = getElements();

    if (urlContinueBtn && !urlContinueBtn.dataset.bound) {
        urlContinueBtn.dataset.bound = "1";
        urlContinueBtn.addEventListener("click", () => {
            handleUrlContinue();
        });
    }

    if (keyBackBtn && !keyBackBtn.dataset.bound) {
        keyBackBtn.dataset.bound = "1";
        keyBackBtn.addEventListener("click", () => {
            setStep("url", loginGateState.stepReason || "startup");
        });
    }

    if (unlockBtn && !unlockBtn.dataset.bound) {
        unlockBtn.dataset.bound = "1";
        unlockBtn.addEventListener("click", () => {
            handleUnlock();
        });
    }

    if (urlInput && !urlInput.dataset.bound) {
        urlInput.dataset.bound = "1";
        urlInput.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                handleUrlContinue();
            }
        });
    }

    if (keyInput && !keyInput.dataset.bound) {
        keyInput.dataset.bound = "1";
        keyInput.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                handleUnlock();
            }
        });
    }
}

function syncSavedUrl() {
    const localValue = (localStorage.getItem(STORAGE_KEYS.APP_SCRIPT_URL) || "").trim();
    const sessionValue = (sessionStorage.getItem(APP_SCRIPT_URL_SESSION_KEY) || "").trim();
    const savedUrl = localValue || sessionValue;
    loginGateState.appScriptUrl = normalizeAppScriptUrl(savedUrl);
    if (loginGateState.appScriptUrl) {
        if (localValue !== loginGateState.appScriptUrl) {
            localStorage.setItem(STORAGE_KEYS.APP_SCRIPT_URL, loginGateState.appScriptUrl);
        }
        if (sessionValue !== loginGateState.appScriptUrl) {
            sessionStorage.setItem(APP_SCRIPT_URL_SESSION_KEY, loginGateState.appScriptUrl);
        }
    }
}

function unlockFromReusableSession() {
    const { overlay } = getElements();
    loginGateState.isLocked = false;
    bindActivityTracking();
    startIdleTimer();
    if (overlay) {
        overlay.classList.add("hidden");
        overlay.classList.remove("is-opening");
    }
    resolveStartupIfPending();
    return { ok: true };
}

export function lockAppNow({ reason = "manual" } = {}) {
    syncSavedUrl();
    clearIdleTimer();
    clearLastUnlockAt();
    loginGateState.isLocked = true;
    const step = loginGateState.appScriptUrl ? "key" : "url";
    showGate(step, reason);
}

export function ensureLoginGate() {
    if (loginGateState.resolved && !loginGateState.isLocked) {
        return Promise.resolve({ ok: true });
    }

    if (loginGateState.promise) {
        return loginGateState.promise;
    }

    const { overlay, urlInput } = getElements();
    if (!overlay) {
        loginGateState.resolved = true;
        return Promise.resolve({ ok: true });
    }

    bindEvents();

    syncSavedUrl();
    if (loginGateState.appScriptUrl && hasReusableSession()) {
        return Promise.resolve(unlockFromReusableSession());
    }

    loginGateState.promise = new Promise((resolve) => {
        loginGateState.resolvePromise = resolve;
    });

    const step = loginGateState.appScriptUrl ? "key" : "url";
    showGate(step, "startup");
    if (step === "url" && urlInput) {
        urlInput.value = loginGateState.appScriptUrl || "";
    }

    return loginGateState.promise;
}
