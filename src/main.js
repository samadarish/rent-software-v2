/**
 * Main Entry Point
 * 
 * This is the application's main entry point.
 * It imports all necessary modules and initializes the application.
 * 
 * The application is organized into:
 * - constants.js: Application constants
 * - state.js: Global state management
 * - utils/: Utility functions (HTML, formatters, UI)
 * - api/: Backend API communication (config, sheets)
 * - features/tenants/: Form and family table experiences
 * - features/agreements/: Clauses and DOCX export helpers
 * - features/billing/: Billing and payments experiences
 * - features/navigation/: Flow switching and routing helpers
 * - features/shared/: Cross-feature utilities such as drafts
 * - events.js: Event handler registration
 */

import { ensureAppScriptConfigured, applyLandlordDefaultsToForm } from "./js/api/config.js";
import { fetchWingsFromSheet, loadClausesFromSheet } from "./js/api/sheets.js";
import { initFamilyTable } from "./js/features/tenants/family.js";
import { initFormOptions, refreshUnitOptions, refreshLandlordOptions } from "./js/features/tenants/form.js";
import { switchFlow } from "./js/features/navigation/flow.js";
import { attachEventHandlers } from "./js/events.js";
import { initToastHistoryUi, showToast, updateConnectionIndicator } from "./js/utils/ui.js";
import { initDraftUi } from "./js/features/shared/drafts.js";
import { initNotesFeature } from "./js/features/shared/notes.js";
import { flushSyncQueue, initSyncManager, startInitialSync } from "./js/api/syncManager.js";
import { initCloseGuard } from "./js/utils/closeGuard.js";
import { currentFlow } from "./js/state.js";

/**
 * Application initialization
 * Runs when the DOM is fully loaded
 */
document.addEventListener("DOMContentLoaded", async () => {
  // Set initial view immediately to prevent agreement flash on first load
  switchFlow("dashboard");

  // Initialize form dropdowns and options
  initFormOptions();

  // Attach all event listeners
  attachEventHandlers();
  initDraftUi();
  initToastHistoryUi();
  initNotesFeature();

  // --- FEATURE: Sync Now ---
  document.getElementById("manualSyncBtn")?.addEventListener("click", () => {
    startInitialSync();
  });

  // --- FEATURE: WhatsApp Login ---
  const waBtn = document.getElementById("whatsappLoginBtn");
  const waStatus = document.getElementById("whatsappStatus");
  const waIcon = document.getElementById("whatsappIconState");

  function setWhatsAppLoggedIn() {
    if (waStatus) waStatus.innerText = "Logged in";
    if (waIcon) waIcon.innerText = "✅";
    localStorage.setItem("wa_logged_in", "true");
  }

  if (localStorage.getItem("wa_logged_in") === "true") {
    setWhatsAppLoggedIn();
  }

  if (waBtn) {
    waBtn.addEventListener("click", async () => {
      try {
        if (window.__TAURI__) {
            await window.__TAURI__.core.invoke("open_whatsapp");
        } else {
            console.warn("Tauri API not available (browser mode)");
        }
      } catch (err) {
        console.error("Failed to open WhatsApp", err);
        alert("Error opening WhatsApp: " + err);
      }
    });
  }

  if (window.__TAURI__) {
      window.__TAURI__.event.listen("whatsapp-login-success", () => {
          setWhatsAppLoggedIn();
          showToast("WhatsApp logged in successfully", "success");
      });
  }

  document.addEventListener("sync:completed", () => {
    void (async () => {
      const tasks = [
        fetchWingsFromSheet(),
        loadClausesFromSheet(false, true),
        refreshUnitOptions(),
        refreshLandlordOptions(),
      ];

      if (currentFlow === "viewTenants") {
        const mod = await import("./js/features/tenants/tenants.js");
        tasks.push(mod.loadTenantDirectory(true));
      }

      if (currentFlow === "payments") {
        const mod = await import("./js/features/billing/payments.js");
        tasks.push(mod.refreshPaymentsIfNeeded(true));
      }

      if (currentFlow === "generateBill") {
        const mod = await import("./js/features/billing/billing.js");
        if (typeof mod.refreshBillingData === "function") {
          tasks.push(mod.refreshBillingData(true));
        }
      }

      await Promise.allSettled(tasks);
    })();
  });

  // Check and prompt for App Script URL if not configured
  await ensureAppScriptConfigured({ autoSync: true });

  applyLandlordDefaultsToForm();
  initSyncManager();
  initCloseGuard();
  if (navigator.onLine) {
    flushSyncQueue();
  }

  // Fire all initial data fetches in parallel
  const initialFetches = [
    fetchWingsFromSheet(),
    loadClausesFromSheet(false),
    refreshUnitOptions(),
    refreshLandlordOptions(),
  ];

  // Initialize the family table with tenant's row
  initFamilyTable();
  await Promise.allSettled(initialFetches);

  updateConnectionIndicator(navigator.onLine ? "online" : "offline");

  const handleOnline = () => {
    updateConnectionIndicator("online", "Internet connected");
    flushSyncQueue();
  };
  window.addEventListener("online", handleOnline);
  window.addEventListener("offline", () => updateConnectionIndicator("offline", "No internet"));
});
