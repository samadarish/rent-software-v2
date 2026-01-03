import { isSyncBusy } from "../api/syncManager.js";
import { hideModal, showModal } from "./ui.js";

let closeGuardReady = false;
let forceClosing = false;

function resolveAppWindow() {
    if (typeof window === "undefined") return null;
    const tauriWindow = window.__TAURI__ && window.__TAURI__.window;
    if (!tauriWindow) return null;
    if (typeof tauriWindow.getCurrentWindow === "function") {
        return tauriWindow.getCurrentWindow();
    }
    return tauriWindow.appWindow || null;
}

function wireModal(modal) {
    if (!modal) return;
    const okBtn = document.getElementById("syncCloseModalOk");
    const dismissBtn = document.getElementById("syncCloseModalDismiss");
    const hide = () => hideModal(modal);

    if (okBtn) okBtn.addEventListener("click", hide);
    if (dismissBtn) dismissBtn.addEventListener("click", hide);
    modal.addEventListener("click", (event) => {
        if (event.target === modal) hide();
    });
}

function requestWindowClose(appWindow) {
    if (!appWindow || forceClosing) return;
    forceClosing = true;
    const closeAction =
        typeof appWindow.close === "function"
            ? appWindow.close()
            : typeof appWindow.destroy === "function"
                ? appWindow.destroy()
                : null;
    Promise.resolve(closeAction)
        .catch((err) => console.warn("Window close failed", err))
        .finally(() => {
            forceClosing = false;
        });
}

function attachCloseHandler(appWindow, modal) {
    if (!appWindow || typeof appWindow.onCloseRequested !== "function") return;
    appWindow.onCloseRequested((event) => {
        if (forceClosing) return;
        if (event && typeof event.preventDefault === "function") {
            event.preventDefault();
        }
        if (isSyncBusy()) {
            if (modal) showModal(modal);
            return;
        }
        requestWindowClose(appWindow);
    });
}

export function initCloseGuard() {
    if (closeGuardReady) return;
    closeGuardReady = true;

    const modal = document.getElementById("syncCloseModal");
    wireModal(modal);

    document.addEventListener("sync:busy", (event) => {
        if (!modal) return;
        if (!event?.detail?.busy) {
            hideModal(modal);
        }
    });

    const appWindow = resolveAppWindow();
    if (!appWindow) return;

    if (typeof appWindow.then === "function") {
        appWindow.then((resolved) => attachCloseHandler(resolved, modal));
    } else {
        attachCloseHandler(appWindow, modal);
    }
}
