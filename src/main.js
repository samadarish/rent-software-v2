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
import { updateConnectionIndicator } from "./js/utils/ui.js";
import { initDraftUi } from "./js/features/shared/drafts.js";
import { flushSyncQueue, initSyncManager, startInitialSync } from "./js/api/syncManager.js";
import { initCloseGuard } from "./js/utils/closeGuard.js";

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

  // Check and prompt for App Script URL if not configured
  ensureAppScriptConfigured();

  applyLandlordDefaultsToForm();
  initSyncManager();
  initCloseGuard();
  if (navigator.onLine) {
    flushSyncQueue();
  }
  document.addEventListener("appscript:url-updated", () => {
    startInitialSync();
  });

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
