/**
 * Window Controls
 * 
 * Custom minimize, maximize, and close buttons for the Tauri window.
 * Used when decorations: false is set in tauri.conf.json.
 */

/**
 * Initialize custom window control buttons
 */
export function initWindowControls() {
    const minimizeBtn = document.getElementById("windowMinimizeBtn");
    const maximizeBtn = document.getElementById("windowMaximizeBtn");
    const closeBtn = document.getElementById("windowCloseBtn");

    if (!window.__TAURI__) {
        // Hide controls in browser mode
        [minimizeBtn, maximizeBtn, closeBtn].forEach(btn => {
            if (btn) btn.style.display = "none";
        });
        return;
    }

    // Minimize button
    if (minimizeBtn) {
        minimizeBtn.addEventListener("click", async () => {
            try {
                const { getCurrentWindow } = window.__TAURI__.window;
                await getCurrentWindow().minimize();
            } catch (err) {
                console.error("Failed to minimize window:", err);
            }
        });
    }

    // Maximize/restore button
    if (maximizeBtn) {
        maximizeBtn.addEventListener("click", async () => {
            try {
                const { getCurrentWindow } = window.__TAURI__.window;
                await getCurrentWindow().toggleMaximize();
            } catch (err) {
                console.error("Failed to toggle maximize:", err);
            }
        });
    }

    // Close button
    if (closeBtn) {
        closeBtn.addEventListener("click", async () => {
            try {
                const { getCurrentWindow } = window.__TAURI__.window;
                await getCurrentWindow().close();
            } catch (err) {
                console.error("Failed to close window:", err);
            }
        });
    }

    // Enable dragging on the header
    const header = document.querySelector("header");
    if (header) {
        header.addEventListener("mousedown", async (e) => {
            // Don't start drag if clicking on buttons or interactive elements
            if (e.target.closest("button, a, input, select, [role='button']")) {
                return;
            }
            try {
                const { getCurrentWindow } = window.__TAURI__.window;
                await getCurrentWindow().startDragging();
            } catch (err) {
                // Ignore drag errors
            }
        });
    }
}
